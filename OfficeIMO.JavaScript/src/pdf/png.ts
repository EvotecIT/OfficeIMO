import { NotSupportedError, OfficeIMOError } from "../core/errors.js";
import { checkAbort, pause, taskYieldDue, withAbort } from "../core/iteration.js";
import { Crc32 } from "../internal/crc32.js";

interface PngHeader {
  readonly width: number;
  readonly height: number;
  readonly components: 3 | 4;
  readonly firstData: number;
  readonly dataEnd: number;
}
export interface PngPixels {
  readonly width: number;
  readonly height: number;
  readonly components: 3;
  readonly data: Uint8Array;
  readonly alpha?: Uint8Array;
}
const signature = [137, 80, 78, 71, 13, 10, 26, 10];
const unsupportedChunks = new Set(["PLTE", "tRNS", "iCCP", "sRGB", "gAMA", "cHRM", "cICP", "mDCV", "cLLI", "sBIT", "eXIf", "acTL", "fcTL", "fdAT"]);

/** Decode the static unprofiled 8-bit RGB/RGBA profile without canvas or a second compression engine. */
export async function pngPixels(inputSource: Uint8Array, maxPixels: number, signal?: AbortSignal, maxWorkingBytes?: number): Promise<PngPixels> {
  checkAbort(signal);
  const shape = pngShape(inputSource, maxPixels);
  const imagePixels = shape.width * shape.height;
  const workingBytes = inputSource.length + imagePixels * shape.components + (shape.components === 4 ? shape.width * 8 : 0);
  if (maxWorkingBytes !== undefined && workingBytes > maxWorkingBytes)
    throw new OfficeIMOError("RESOURCE_LIMIT", "maxImageWorkingBytes exceeded by PNG source, image planes and scanlines.");
  const source = new Uint8Array(inputSource);
  const header = await pngHeader(source, shape, signal);
  checkAbort(signal);
  let stream: DecompressionStream;
  try {
    if (typeof DecompressionStream !== "function") throw new TypeError("Missing DecompressionStream");
    stream = new DecompressionStream("deflate");
  } catch (cause) {
    throw new OfficeIMOError("PLATFORM_UNAVAILABLE", "PNG intake requires native DecompressionStream(\"deflate\").", { cause });
  }
  const { width, height, components } = header, rowBytes = width * components;
  const expected = (rowBytes + 1) * height;
  // All dimensions and the exact filtered length are bounded before allocating decoded pixels.
  if (!Number.isSafeInteger(expected) || expected > 0x7fffffff)
    throw new OfficeIMOError("RESOURCE_LIMIT", "PNG filtered data exceeds the supported decoded byte length.");
  // RGBA is reconstructed one scanline at a time directly into PDF RGB/alpha planes.
  // It never retains an additional full interleaved RGBA image.
  const pixels = new Uint8Array(width * height * 3), alpha = components === 4 ? new Uint8Array(imagePixels) : undefined;
  let previous = components === 4 ? new Uint8Array(rowBytes) : undefined, current = components === 4 ? new Uint8Array(rowBytes) : undefined;
  const view = new DataView(source.buffer, source.byteOffset, source.byteLength);
  let chunk = header.firstData, position = chunk + 8, end = position + view.getUint32(chunk);
  const input = new ReadableStream<BufferSource>({ pull(controller) {
    checkAbort(signal);
    while (position === end && chunk < header.dataEnd) {
      chunk = end + 4;
      if (chunk === header.dataEnd) { controller.close(); return; }
      position = chunk + 8; end = position + view.getUint32(chunk);
    }
    // Limit the input given to the platform inflater between downstream length checks.
    const next = Math.min(end, position + 4096);
    controller.enqueue(source.subarray(position, next)); position = next;
  } }, { highWaterMark: 0 });
  const reader = input.pipeThrough(stream).getReader();
  let inflated = 0, output = 0, rgbOutput = 0, alphaOutput = 0, column = 0, filter = 0;
  try {
    while (true) {
      const item = await withAbort(reader.read(), signal).catch(cause => {
        checkAbort(signal);
        throw new TypeError("Invalid PNG zlib image data.", { cause });
      });
      checkAbort(signal);
      if (item.done) break;
      if (item.value.length > expected - inflated) throw new TypeError("PNG inflated bytes exceed the exact IHDR scanline length.");
      for (let p = 0; p < item.value.length; p++) {
        const value = item.value[p]!;
        if (column === 0) {
          filter = value;
          if (filter > 4) throw new TypeError("Invalid PNG scanline filter.");
        } else {
          const index = column - 1;
          let decoded = value;
          if (filter !== 0) {
            const left = column > components ? current ? current[index - components]! : pixels[output - components]! : 0;
            const up = previous ? previous[index]! : output >= rowBytes ? pixels[output - rowBytes]! : 0;
            const upperLeft = column > components ? previous ? previous[index - components]! : output >= rowBytes ? pixels[output - rowBytes - components]! : 0 : 0;
            decoded = (value + (filter === 1 ? left : filter === 2 ? up :
              filter === 3 ? Math.floor((left + up) / 2) : paeth(left, up, upperLeft))) & 255;
          }
          if (current) current[index] = decoded;
          else pixels[output] = decoded;
          output++;
        }
        if (column === rowBytes) {
          column = 0;
          if (current && alpha) {
            // Split each completed row once: no per-byte division or channel dispatch,
            // and filter reconstruction still uses only these same two scanlines.
            for (let c = 0; c < rowBytes; c += 4) {
              pixels[rgbOutput++] = current[c]!;
              pixels[rgbOutput++] = current[c + 1]!;
              pixels[rgbOutput++] = current[c + 2]!;
              alpha[alphaOutput++] = current[c + 3]!;
              if ((c & 16383) === 16380) { if (taskYieldDue()) await pause(); checkAbort(signal); }
            }
            const completed = current; current = previous; previous = completed;
          }
        }
        else column++;
        if ((p & 16383) === 16383) { if (taskYieldDue()) await pause(); checkAbort(signal); }
      }
      inflated += item.value.length;
      if (taskYieldDue()) await pause(); checkAbort(signal);
    }
    if (inflated !== expected) throw new TypeError("PNG inflated bytes must match the exact IHDR scanline length.");
    return { width, height, components: 3, data: pixels, ...(alpha ? { alpha } : {}) };
  } catch (error) {
    void reader.cancel(error).catch(() => {});
    throw error;
  } finally { reader.releaseLock(); }
}

/** Verify the complete source and critical chunk order before any decompressed image allocation. */
async function pngHeader(source: Uint8Array, shape: Pick<PngHeader, "width" | "height" | "components">, signal?: AbortSignal): Promise<PngHeader> {
  const view = new DataView(source.buffer, source.byteOffset, source.byteLength);
  let firstData = 0, dataEnd = 0, dataClosed = false;
  for (let offset = 8; offset < source.length;) {
    checkAbort(signal);
    if (offset + 12 > source.length) throw new TypeError("Truncated PNG chunk.");
    const length = view.getUint32(offset), end = offset + 12 + length;
    if (length > 0x7fffffff || end > source.length) throw new TypeError("Truncated or oversized PNG chunk.");
    const type = String.fromCharCode(...source.subarray(offset + 4, offset + 8));
    if (!/^[A-Za-z]{2}[A-Z][A-Za-z]$/.test(type)) throw new TypeError("Invalid PNG chunk type.");
    const crc = new Crc32();
    for (let p = offset + 4; p < end - 4; p += 65536) {
      crc.update(source.subarray(p, Math.min(end - 4, p + 65536)));
      if (taskYieldDue()) await pause(); checkAbort(signal);
    }
    if (crc.value !== view.getUint32(end - 4)) throw new TypeError("PNG chunk CRC mismatch: " + type + ".");
    if (offset === 8 && type !== "IHDR") throw new TypeError("PNG IHDR must be the first chunk.");
    if (type === "IHDR") {
      if (offset !== 8 || length !== 13) throw new TypeError("PNG requires one 13-byte IHDR first.");
    } else if (type === "IDAT") {
      if (dataClosed) throw new TypeError("PNG IDAT chunks must be consecutive.");
      firstData ||= offset; dataEnd = end;
    } else {
      if (firstData) dataClosed = true;
      if (type === "IEND") {
        if (length !== 0 || !firstData || end !== source.length) throw new TypeError("PNG requires a final empty IEND after IDAT.");
        return { ...shape, firstData, dataEnd };
      }
      if (unsupportedChunks.has(type)) throw new NotSupportedError("PNG chunk " + type + "; normalize color, orientation, transparency or animation first");
      if ((source[offset + 4]! & 32) === 0) throw new NotSupportedError("Required PNG chunk " + type);
    }
    offset = end;
  }
  throw new TypeError("PNG requires a complete final IEND chunk.");
}
/** Check source metadata synchronously, before capturing source bytes or allocating any image plane. */
function pngShape(source: Uint8Array, maxPixels: number): Pick<PngHeader, "width" | "height" | "components"> {
  if (source.length < 33 || signature.some((value, index) => source[index] !== value)) throw new TypeError("Invalid PNG signature or truncated IHDR.");
  const view = new DataView(source.buffer, source.byteOffset, source.byteLength);
  if (view.getUint32(8) !== 13 || String.fromCharCode(...source.subarray(12, 16)) !== "IHDR") throw new TypeError("PNG requires one 13-byte IHDR first.");
  const width = view.getUint32(16), height = view.getUint32(20);
  if (!width || !height || width > 0x7fffffff || height > 0x7fffffff) throw new TypeError("Invalid PNG dimensions.");
  if (!Number.isSafeInteger(width * height) || width * height > maxPixels)
    throw new OfficeIMOError("RESOURCE_LIMIT", "PNG dimensions exceed maxImagePixels.");
  if (source[24] !== 8 || ![2, 6].includes(source[25]!)) throw new NotSupportedError("PNG intake requires 8-bit RGB or RGBA pixels");
  if (source[26] !== 0 || source[27] !== 0) throw new NotSupportedError("PNG compression or filter method");
  if (source[28] !== 0) throw new NotSupportedError("Interlaced PNG intake");
  const components = source[25] === 6 ? 4 : 3;
  if (!Number.isSafeInteger((width * components + 1) * height) || (width * components + 1) * height > 0x7fffffff)
    throw new OfficeIMOError("RESOURCE_LIMIT", "PNG filtered data exceeds the supported decoded byte length.");
  return { width, height, components };
}
function paeth(left: number, up: number, upperLeft: number): number {
  const prediction = left + up - upperLeft;
  const a = Math.abs(prediction - left), b = Math.abs(prediction - up), c = Math.abs(prediction - upperLeft);
  return a <= b && a <= c ? left : b <= c ? up : upperLeft;
}
