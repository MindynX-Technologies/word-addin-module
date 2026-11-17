/** ArrayBuffer storage avoids Blob/File IndexedDB failures in some WebKit hosts. */
export type StoredFile = { name: string; type: string; lastModified: number; bytes: ArrayBuffer };

export async function storeFile(file: File): Promise<StoredFile> {
  return { name: file.name, type: file.type, lastModified: file.lastModified, bytes: await file.arrayBuffer() };
}

export function loadFile(file: File | StoredFile): File {
  return file instanceof File ? file : new File([file.bytes], file.name, { type: file.type, lastModified: file.lastModified });
}
