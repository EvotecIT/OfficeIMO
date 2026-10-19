/** @internal Recognize native Date values across iframe realms without trusting a string tag. */
export function isDate(value: unknown): value is Date {
  if (value === null || typeof value !== "object") return false;
  try {
    // Ordinary rows take the cheap negative path. A custom string tag can be an
    // accessor: use the native brand check directly without reading that tag.
    if (!(value instanceof Date) && !(Symbol.toStringTag in value) && Object.prototype.toString.call(value) !== "[object Date]") return false;
    Date.prototype.getTime.call(value); return true;
  }
  catch { return false; }
}
