/** @internal Recognize native Date values across iframe realms without trusting a string tag. */
export function isDate(value: unknown): value is Date {
  if (value === null || typeof value !== "object") return false;
  // Ordinary domain rows are objects too. Avoid throwing for each one; the native
  // check below still rejects objects that merely advertise a Date string tag.
  if (!(value instanceof Date) && Object.prototype.toString.call(value) !== "[object Date]") return false;
  try { Date.prototype.getTime.call(value); return true; }
  catch { return false; }
}
