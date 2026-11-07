// Little-endian readers for BIFF8 records, including strings that continue across CONTINUE records.

export class Segments {
  private index = 0;
  private offset = 0;
  constructor(private readonly parts: Uint8Array[]) {}

  get exhausted(): boolean {
    return this.index >= this.parts.length || (this.index === this.parts.length - 1 && this.offset >= this.parts[this.index]!.length);
  }

  private current(): Uint8Array | undefined {
    while (this.index < this.parts.length && this.offset >= this.parts[this.index]!.length) {
      this.index++;
      this.offset = 0;
    }
    return this.parts[this.index];
  }

  u8(): number {
    const part = this.current();
    if (!part) throw new RangeError("record overrun");
    return part[this.offset++]!;
  }

  u16(): number {
    return this.u8() | (this.u8() << 8);
  }

  u32(): number {
    return (this.u16() | (this.u16() << 16)) >>> 0;
  }

  skip(count: number) {
    for (let i = 0; i < count; i++) this.u8();
  }

  /** Reads `count` characters; a CONTINUE boundary inside the string repeats the flags byte. */
  chars(count: number, highByte: boolean): string {
    let out = "";
    let wide = highByte;
    let remaining = count;
    while (remaining > 0) {
      const part = this.current();
      if (!part) break;
      const available = part.length - this.offset;
      const fit = wide ? Math.floor(available / 2) : available;
      if (fit === 0) {
        // Boundary: move to the next record, which starts with a new flags byte.
        this.index++;
        this.offset = 0;
        const next = this.parts[this.index];
        if (!next) break;
        wide = (next[this.offset++]! & 1) === 1;
        continue;
      }
      const take = Math.min(fit, remaining);
      out += wide ? decodeUtf16(part, this.offset, take) : decodeLatin1(part, this.offset, take);
      this.offset += wide ? take * 2 : take;
      remaining -= take;
      if (remaining > 0 && this.offset >= part.length) {
        this.index++;
        this.offset = 0;
        const next = this.parts[this.index];
        if (!next) break;
        wide = (next[this.offset++]! & 1) === 1;
      }
    }
    return out;
  }
}

export function decodeUtf16(bytes: Uint8Array, offset: number, count: number): string {
  let out = "";
  for (let i = 0; i < count; i++) out += String.fromCharCode(bytes[offset + i * 2]! | (bytes[offset + i * 2 + 1]! << 8));
  return out;
}

export function decodeLatin1(bytes: Uint8Array, offset: number, count: number): string {
  let out = "";
  for (let i = 0; i < count; i++) out += String.fromCharCode(bytes[offset + i]!);
  return out;
}

export const u16 = (b: Uint8Array, o: number) => b[o]! | (b[o + 1]! << 8);
export const i16 = (b: Uint8Array, o: number) => (u16(b, o) << 16) >> 16;
export const u32 = (b: Uint8Array, o: number) => (b[o]! | (b[o + 1]! << 8) | (b[o + 2]! << 16) | (b[o + 3]! << 24)) >>> 0;

export function f64(b: Uint8Array, o: number): number {
  return new DataView(b.buffer, b.byteOffset + o, 8).getFloat64(0, true);
}

/** RK-encoded number. */
export function rk(value: number): number {
  let number: number;
  if (value & 0x02) {
    number = value >> 2;
  } else {
    const view = new DataView(new ArrayBuffer(8));
    view.setUint32(4, value & 0xfffffffc, true);
    number = view.getFloat64(0, true);
  }
  return value & 0x01 ? number / 100 : number;
}

/** XLUnicodeString: u16 count, flags, characters. */
export function xlString(b: Uint8Array, o: number): { text: string; size: number } {
  const count = u16(b, o);
  const wide = (b[o + 2]! & 1) === 1;
  return { text: wide ? decodeUtf16(b, o + 3, count) : decodeLatin1(b, o + 3, count), size: 3 + count * (wide ? 2 : 1) };
}

/** ShortXLUnicodeString: u8 count, flags, characters. */
export function shortString(b: Uint8Array, o: number): { text: string; size: number } {
  const count = b[o]!;
  const wide = (b[o + 1]! & 1) === 1;
  return { text: wide ? decodeUtf16(b, o + 2, count) : decodeLatin1(b, o + 2, count), size: 2 + count * (wide ? 2 : 1) };
}

export interface BiffRecord {
  type: number;
  data: Uint8Array;
  /** Payloads of CONTINUE records that follow. */
  continues: Uint8Array[];
  offset: number;
}

export const CONTINUE = 0x003c;

export function readRecords(stream: Uint8Array, start: number): BiffRecord[] {
  const out: BiffRecord[] = [];
  let offset = start;
  while (offset + 4 <= stream.length) {
    const type = u16(stream, offset);
    const length = u16(stream, offset + 2);
    const data = stream.subarray(offset + 4, offset + 4 + length);
    if (type === CONTINUE && out.length > 0) out[out.length - 1]!.continues.push(data);
    else out.push({ type, data, continues: [], offset });
    offset += 4 + length;
    if (type === 0x000a) break; // EOF
  }
  return out;
}
