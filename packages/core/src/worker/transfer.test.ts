import { describe, expect, it } from 'vitest';
import { exactTransferableArrayBuffer } from './transfer';

describe('exactTransferableArrayBuffer', () => {
  it('copies source-owned pull bytes even when the view spans its backing buffer', () => {
    const sourceBytes = new Uint8Array([1, 2, 3]);
    const payload = exactTransferableArrayBuffer(sourceBytes, true);
    expect(payload).not.toBe(sourceBytes.buffer);
    structuredClone(payload, { transfer: [payload] });
    expect(sourceBytes).toEqual(new Uint8Array([1, 2, 3]));
  });

  it('transfers only the selected range of a source-owned subarray', () => {
    const sourceBytes = new Uint8Array([9, 1, 2, 3, 9]);
    const payload = exactTransferableArrayBuffer(sourceBytes.subarray(1, 4), true);
    expect(new Uint8Array(payload)).toEqual(new Uint8Array([1, 2, 3]));
    structuredClone(payload, { transfer: [payload] });
    expect(sourceBytes).toEqual(new Uint8Array([9, 1, 2, 3, 9]));
  });

  it('copies a Buffer slice whose slice() method would keep borrowed storage', () => {
    const borrowed = Buffer.from([9, 1, 2, 9]).subarray(1, 3);
    const payload = exactTransferableArrayBuffer(borrowed, true);
    expect(new Uint8Array(payload)).toEqual(new Uint8Array([1, 2]));
    structuredClone(payload, { transfer: [payload] });
    expect(borrowed).toEqual(Buffer.from([1, 2]));
  });
});
