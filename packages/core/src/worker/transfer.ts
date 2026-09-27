/**
 * Return an exactly-sized transferable buffer. A full ArrayBuffer-backed view
 * can transfer without copying; sliced/shared views are detached from their
 * backing allocation first so unrelated bytes never cross the boundary.
 */
export function exactTransferableArrayBuffer(bytes: Uint8Array): ArrayBuffer {
  if (
    bytes.byteOffset === 0
    && bytes.byteLength === bytes.buffer.byteLength
    && bytes.buffer instanceof ArrayBuffer
  ) return bytes.buffer;
  return bytes.slice().buffer as ArrayBuffer;
}
