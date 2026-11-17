"""Local Word conversion companion. Run in the supplied restricted container."""
import json
import os
from pathlib import Path
import shutil
import signal
import subprocess
import tempfile
import threading
from http.server import BaseHTTPRequestHandler, ThreadingHTTPServer
from urllib.parse import parse_qs, urlsplit
import zipfile

MAX_BYTES = 32 * 1024 * 1024
OLE = bytes.fromhex("d0cf11e0a1b11ae1")
SLOTS = threading.BoundedSemaphore(2)
FORMATS = {"doc": "doc:MS Word 97", "docx": "docx:Office Open XML Text"}
MIMES = {"doc": "application/msword", "docx": "application/vnd.openxmlformats-officedocument.wordprocessingml.document"}


def validate(data, kind):
    if kind == "doc":
        if not data.startswith(OLE):
            raise ValueError("Choose a binary Word DOC file.")
        return
    import io
    with zipfile.ZipFile(io.BytesIO(data)) as archive:
        entries = archive.infolist()
        if len(entries) > 4096 or sum(part.file_size for part in entries) > 128 * 1024 * 1024:
            raise ValueError("The document exceeds the expanded archive limit.")
        if any(part.file_size > 48 * 1024 * 1024 or part.flag_bits & 1 for part in entries):
            raise ValueError("Encrypted or oversized archive entries are unsupported.")
        if "[Content_Types].xml" not in archive.namelist() or "word/document.xml" not in archive.namelist():
            raise ValueError("Choose a standard DOCX document.")


def convert(data, source, target):
    validate(data, source)
    with tempfile.TemporaryDirectory(prefix="word-convert-") as directory:
        root = Path(directory)
        profile = root / "profile" / "user"
        profile.mkdir(parents=True)
        # Never run document macros. Each request has an independent office profile.
        (profile / "registrymodifications.xcu").write_text('''<?xml version="1.0"?>
<oor:items xmlns:oor="http://openoffice.org/2001/registry">
<item oor:path="/org.openoffice.Office.Common/Security/Scripting"><prop oor:name="MacroSecurityLevel" oor:op="fuse" oor:finalized="true"><value>3</value></prop><prop oor:name="BlockUntrustedRefererLinks" oor:op="fuse" oor:finalized="true"><value>true</value></prop></item>
<item oor:path="/org.openoffice.Office.Writer/Content/Update"><prop oor:name="Link" oor:op="fuse" oor:finalized="true"><value>0</value></prop></item>
</oor:items>''')
        incoming = root / ("document." + source)
        incoming.write_bytes(data)
        output = root / "output"
        output.mkdir()
        command = [os.environ.get("LIBREOFFICE", "soffice"),
                   "-env:UserInstallation=" + (root / "profile").as_uri(),
                   "--headless", "--nologo", "--nodefault", "--norestore",
                   "--convert-to", FORMATS[target], "--outdir", str(output), str(incoming)]
        process = subprocess.Popen(command, stdout=subprocess.PIPE, stderr=subprocess.PIPE, start_new_session=True,
                                   env={**os.environ, "SAL_USE_VCLPLUGIN": "svp"})
        try:
            process.communicate(timeout=45)
        except subprocess.TimeoutExpired:
            os.killpg(process.pid, signal.SIGKILL)
            process.communicate()
            raise
        converted = output / ("document." + target)
        if process.returncode or not converted.is_file():
            raise ValueError("Word conversion failed. Encrypted or damaged documents cannot be converted.")
        if converted.stat().st_size > MAX_BYTES:
            raise ValueError("The converted document exceeds 32 MiB.")
        data = converted.read_bytes()
        validate(data, target)
        return data


class Handler(BaseHTTPRequestHandler):
    server_version = "WordConverter"

    def reply(self, status, data, mime="application/json"):
        self.send_response(status)
        self.send_header("Content-Type", mime)
        self.send_header("Content-Length", str(len(data)))
        self.send_header("Cache-Control", "no-store")
        self.send_header("X-Content-Type-Options", "nosniff")
        self.end_headers()
        self.wfile.write(data)

    def error(self, status, message):
        self.reply(status, json.dumps({"message": message}).encode())

    def do_GET(self):
        if self.path != "/health":
            return self.error(404, "Not found.")
        ready = shutil.which(os.environ.get("LIBREOFFICE", "soffice")) is not None
        self.reply(200 if ready else 503, json.dumps({"ready": ready}).encode())

    def do_POST(self):
        url = urlsplit(self.path)
        query = parse_qs(url.query)
        source, target = query.get("input", [""])[0], query.get("format", [""])[0]
        if url.path != "/convert" or source not in FORMATS or target not in FORMATS or source == target:
            return self.error(400, "Choose DOC to DOCX or DOCX to DOC conversion.")
        # Browser clients must explicitly send binary content through the same-origin proxy.
        # No CORS headers are emitted and cross-site form requests are rejected.
        if self.headers.get("Sec-Fetch-Site") == "cross-site" or self.headers.get("Content-Type", "").split(";")[0] not in (MIMES[source], "application/octet-stream"):
            return self.error(403, "Use the configured same-origin Word converter.")
        try:
            length = int(self.headers.get("Content-Length", "0"))
        except ValueError:
            return self.error(400, "Invalid request size.")
        if length <= 0 or length > MAX_BYTES:
            return self.error(413, "Choose a document smaller than 32 MiB.")
        if not SLOTS.acquire(blocking=False):
            return self.error(429, "The converter is busy. Try again shortly.")
        try:
            self.connection.settimeout(50)
            data = self.rfile.read(length)
            if len(data) != length:
                return self.error(400, "The upload was incomplete.")
            self.reply(200, convert(data, source, target), MIMES[target])
        except subprocess.TimeoutExpired:
            self.error(504, "Word conversion timed out.")
        except (ValueError, zipfile.BadZipFile) as problem:
            self.error(422, str(problem))
        except Exception:
            self.error(500, "The Word converter could not complete this request.")
        finally:
            SLOTS.release()

    def log_message(self, format, *args):
        # Paths contain format names only; document contents and filenames are never logged.
        super().log_message(format, *args)


if __name__ == "__main__":
    ThreadingHTTPServer((os.environ.get("HOST", "127.0.0.1"), int(os.environ.get("PORT", "4175"))), Handler).serve_forever()
