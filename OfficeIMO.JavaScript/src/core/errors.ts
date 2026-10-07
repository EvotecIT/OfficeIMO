/** Stable machine-readable failures for document and package operations. */
export type ErrorCode = "NOT_SUPPORTED" | "INVALID_XML" | "INVALID_PART_URI" | "INVALID_STATE" | "ZIP64_REQUIRED" | "PLATFORM_UNAVAILABLE" | "RESOURCE_LIMIT";

export class OfficeIMOError extends Error {
  constructor(readonly code: ErrorCode, message: string, options?: ErrorOptions) {
    super(message, options);
    this.name = "OfficeIMOError";
  }
}

export class NotSupportedError extends OfficeIMOError {
  constructor(readonly feature: string) {
    super("NOT_SUPPORTED", feature + " is not supported yet.");
    this.name = "NotSupportedError";
  }
}
