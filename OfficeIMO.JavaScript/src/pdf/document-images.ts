import { NotSupportedError, OfficeIMOError } from "../core/errors.js";
import { checkAbort, pause, taskYieldDue } from "../core/iteration.js";
import type { PdfImageBlock } from "./document-types.js";
import type { PdfObjects } from "./objects.js";
import type { PdfSettings } from "./settings.js";
import { pngPixels } from "./png.js";

interface ImageResource { readonly name: string; readonly id: number; readonly width: number; readonly height: number; }
/** Shared document image resources; source objects are captured on first consumption. */
export class PdfDocumentImages {
  private readonly images = new Map<PdfImageBlock["image"], ImageResource>();
  private bytes = 0;
  constructor(private readonly objects: PdfObjects, private readonly settings: PdfSettings) {}
  async prepare(image: PdfImageBlock["image"]): Promise<ImageResource> {
    const existing = this.images.get(image); if (existing) return existing;
    const { limits, options } = this.settings;
    if (!image || typeof image !== "object" || !(image.data instanceof Uint8Array)) throw new TypeError("PDF images require a byte array.");
    if (this.bytes + image.data.length > limits.maxImageBytes) throw new OfficeIMOError("RESOURCE_LIMIT", "maxImageBytes exceeded.");
    let width: number, height: number, components: number, decoded: Uint8Array | undefined;
    if (image.format === "png") ({ width, height, components, data: decoded } = await pngPixels(image.data, limits.maxImagePixels, options.signal));
    else if (image.format === "jpeg") ({ width, height, components } = jpegSize(image.data));
    else if (image.format === "rgb" || image.format === "rgba") {
      ({ width, height } = image); components = image.format === "rgba" ? 4 : 3;
    } else throw new NotSupportedError("PDF images support unprofiled PNG/JPEG or decoded RGB/RGBA pixels.");
    if (!Number.isSafeInteger(width) || !Number.isSafeInteger(height) || width < 1 || height < 1 ||
        !Number.isSafeInteger(width * height) || width * height > limits.maxImagePixels)
      throw new OfficeIMOError("RESOURCE_LIMIT", "Image dimensions must be positive integers within maxImagePixels.");
    if ((image.format === "rgb" || image.format === "rgba") && image.data.length !== width * height * components) throw new RangeError("Raw image bytes must match width × height × channels exactly.");
    checkAbort(options.signal);
    const source = decoded ?? new Uint8Array(image.data), id = this.objects.reserve(), name = "Image" + (this.images.size + 1);
    this.bytes += image.data.length;
    let pixels = source, mask: number | undefined;
    if (components === 4) {
      pixels = new Uint8Array(width * height * 3); const alpha = new Uint8Array(width * height);
      for (let p = 0; p < alpha.length; p++) {
        pixels[p * 3] = source[p * 4]!; pixels[p * 3 + 1] = source[p * 4 + 1]!; pixels[p * 3 + 2] = source[p * 4 + 2]!; alpha[p] = source[p * 4 + 3]!;
        if ((p & 16383) === 16383) { if (taskYieldDue()) await pause(); checkAbort(options.signal); }
      }
      mask = this.objects.reserve();
      await this.objects.stream(mask, alpha, "/Type /XObject /Subtype /Image /Width " + width + " /Height " + height + " /ColorSpace /DeviceGray /BitsPerComponent 8", options.compression !== false);
    }
    const dictionary = "/Type /XObject /Subtype /Image /Width " + width + " /Height " + height + " /ColorSpace /" +
      (components === 1 ? "DeviceGray" : "DeviceRGB") + " /BitsPerComponent 8" + (mask ? " /SMask " + mask + " 0 R" : "") +
      (image.format === "jpeg" ? " /Filter /DCTDecode" : "");
    await this.objects.stream(id, pixels, dictionary, image.format === "jpeg" ? false : options.compression !== false);
    const resource = { name, id, width, height }; this.images.set(image, resource); return resource;
  }
}

/** Inspect bounded JPEG headers; JPEG entropy decoding remains the independent reader's job. */
function jpegSize(bytes: Uint8Array): { width: number; height: number; components: number } {
  if (bytes.length < 4 || bytes[0] !== 255 || bytes[1] !== 216 || bytes[bytes.length - 2] !== 255 || bytes[bytes.length - 1] !== 217)
    throw new TypeError("JPEG images need complete SOI/EOI markers.");
  let result: { width: number; height: number; components: number } | undefined;
  for (let p = 2; p < bytes.length;) {
    if (bytes[p++] !== 255) throw new TypeError("Invalid JPEG marker.");
    while (bytes[p] === 255) p++;
    const marker = bytes[p++]!;
    if (marker === 218) { if (!result) throw new TypeError("JPEG scan precedes its frame header."); return result; }
    if (marker === 217) break;
    if (p + 2 > bytes.length) throw new TypeError("Truncated JPEG segment.");
    const length = bytes[p]! * 256 + bytes[p + 1]!, end = p + length;
    if (length < 2 || end > bytes.length) throw new TypeError("Truncated JPEG segment.");
    if (marker === 226 && length >= 14 && String.fromCharCode(...bytes.subarray(p + 2, p + 14)) === "ICC_PROFILE\0")
      throw new NotSupportedError("Normalize ICC-profiled JPEG colors before PDF embedding.");
    if (marker === 225 && length >= 8 && String.fromCharCode(...bytes.subarray(p + 2, p + 8)) === "Exif\0\0")
      throw new NotSupportedError("Normalize JPEG EXIF orientation before PDF embedding.");
    if (marker >= 192 && marker <= 207 && ![196, 200, 204].includes(marker)) {
      if (![192, 193, 194].includes(marker) || length < 8 || bytes[p + 2] !== 8) throw new NotSupportedError("PDF JPEG images require 8-bit baseline, sequential or progressive frames.");
      const components = bytes[p + 7]!;
      if (![1, 3].includes(components) || length !== 8 + components * 3) throw new NotSupportedError("PDF JPEG images require grayscale or RGB components; CMYK is unsupported.");
      if (components === 3) {
        const ids = [bytes[p + 8], bytes[p + 11], bytes[p + 14]].join(",");
        if (ids !== "1,2,3" && ids !== "82,71,66") throw new NotSupportedError("Unsupported JPEG component color layout.");
      }
      if (result) throw new NotSupportedError("Multiple JPEG frame headers are unsupported.");
      result = { width: bytes[p + 5]! * 256 + bytes[p + 6]!, height: bytes[p + 3]! * 256 + bytes[p + 4]!, components };
    }
    p = end;
  }
  throw new TypeError("JPEG image has no complete frame and scan.");
}
