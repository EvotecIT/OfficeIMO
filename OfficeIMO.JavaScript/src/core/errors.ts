/** Stable machine-readable failures for document and package operations. */
export type ErrorCode = "NOT_SUPPORTED" | "INVALID_XML" | "INVALID_PART_URI" | "INVALID_STATE" | "ZIP64_REQUIRED" | "PLATFORM_UNAVAILABLE" | "RESOURCE_LIMIT";

const errorBrand = Symbol.for("@evotecit/officeimo/OfficeIMOError");
const unsupportedBrand = Symbol.for("@evotecit/officeimo/NotSupportedError");

export class OfficeIMOError extends Error {
  static override [Symbol.hasInstance](value: unknown): boolean {
    // Inherited checks on user subclasses must retain their own prototype identity.
    if (this !== OfficeIMOError) return Function.prototype[Symbol.hasInstance].call(this, value);
    return !!value && typeof value === "object" && (value as Record<symbol, unknown>)[errorBrand] === true;
  }
  constructor(readonly code: ErrorCode, message: string, options?: ErrorOptions) {
    super(message, options);
    Object.defineProperty(this, errorBrand, { value: true });
    this.name = "OfficeIMOError";
  }
}

export class NotSupportedError extends OfficeIMOError {
  static override [Symbol.hasInstance](value: unknown): boolean {
    if (this !== NotSupportedError) return Function.prototype[Symbol.hasInstance].call(this, value);
    return !!value && typeof value === "object" && (value as Record<symbol, unknown>)[unsupportedBrand] === true;
  }
  constructor(readonly feature: string) {
    super("NOT_SUPPORTED", feature + " is not supported yet.");
    Object.defineProperty(this, unsupportedBrand, { value: true });
    this.name = "NotSupportedError";
  }
}
