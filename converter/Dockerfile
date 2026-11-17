FROM python:3.12-slim-bookworm
RUN apt-get update && apt-get install -y --no-install-recommends libreoffice-writer fonts-liberation fonts-dejavu-core \
    && rm -rf /var/lib/apt/lists/* \
    && useradd --create-home --uid 10001 converter
COPY server.py /app/server.py
USER converter
ENV HOST=0.0.0.0 PORT=4175 HOME=/tmp/office-home
EXPOSE 4175
CMD ["python", "/app/server.py"]
