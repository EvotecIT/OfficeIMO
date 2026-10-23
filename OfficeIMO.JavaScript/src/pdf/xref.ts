/** @internal One type byte, an eight-byte offset/container number and a two-byte generation/index. */
export const pdfXrefEntryBytes = 11;

interface PdfXrefIndices {
  readonly length: number;
  at(id: number): number | undefined;
}

function unsigned(bytes: Uint8Array, offset: number, value: number, width: number): void {
  if (!Number.isSafeInteger(value) || value < 0 || value >= 2 ** (width * 8)) throw new RangeError("Invalid PDF cross-reference field.");
  // Bitwise shifts truncate to 32 bits. Arithmetic preserves every supported byte offset.
  for (let index = width - 1; index >= 0; index--) { bytes[offset + index] = value % 256; value = Math.floor(value / 256); }
}

/** @internal Stream a complete single-section xref, including its own ordinary-object entry.
 * A defined member index means that the location is an object-stream number rather than a byte offset. */
export function* pdfXrefChunks(locations: readonly (number | undefined)[], indices: PdfXrefIndices): Iterable<Uint8Array> {
  if (locations.length !== indices.length) throw new Error("PDF cross-reference locations must align.");
  const perChunk = Math.floor(32768 / pdfXrefEntryBytes);
  for (let first = 0; first < locations.length; first += perChunk) {
    const count = Math.min(perChunk, locations.length - first), bytes = new Uint8Array(count * pdfXrefEntryBytes);
    for (let index = 0; index < count; index++) {
      const id = first + index, position = index * pdfXrefEntryBytes, location = locations[id], member = indices.at(id);
      if (location === undefined) throw new Error("PDF contains unwritten objects.");
      bytes[position] = id === 0 ? 0 : member === undefined ? 1 : 2;
      unsigned(bytes, position + 1, location, 8);
      unsigned(bytes, position + 9, id === 0 ? 65535 : member ?? 0, 2);
    }
    yield bytes;
  }
}
