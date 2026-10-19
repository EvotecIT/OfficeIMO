const crcTable = new Uint32Array(256).map((_, index) => {
  let value = index;
  for (let bit = 0; bit < 8; bit++) value = value & 1 ? 0xedb88320 ^ (value >>> 1) : value >>> 1;
  return value >>> 0;
});

/** Incremental CRC-32/ISO-HDLC shared by ZIP and PNG chunk validation. */
export class Crc32 {
  private crc = 0xffffffff;
  update(bytes: Uint8Array): void {
    let crc = this.crc;
    for (let index = 0; index < bytes.length; index++) crc = crcTable[(crc ^ bytes[index]!) & 255]! ^ (crc >>> 8);
    this.crc = crc;
  }
  get value(): number { return (this.crc ^ 0xffffffff) >>> 0; }
}
