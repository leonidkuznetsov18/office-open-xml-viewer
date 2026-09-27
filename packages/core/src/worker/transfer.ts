/**
 * Return an exactly-sized transferable buffer. A full ArrayBuffer-backed view
 * can transfer without copying; sliced/shared views are detached from their
 * backing allocation first so unrelated bytes never cross the boundary.
 * Use `copy` for application-supplied model-source bytes: a full-span view
 * may still be borrowed, so its original backing must not be detached.
 */
export function exactTransferableArrayBuffer(bytes: Uint8Array, copy = false): ArrayBuffer {
  if (copy) {
    const owned = new Uint8Array(bytes.byteLength);
    owned.set(bytes);
    return owned.buffer;
  }
  if (
    bytes.byteOffset === 0
    && bytes.byteLength === bytes.buffer.byteLength
    && bytes.buffer instanceof ArrayBuffer
  ) return bytes.buffer;
  return bytes.slice().buffer as ArrayBuffer;
}
