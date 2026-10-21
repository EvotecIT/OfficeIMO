/** Capture a structural option record before validation or asynchronous work.
 * Own fields and supported prototype accessors are read once. Nested values remain
 * borrowed until their owning format captures or validates them. */
export function captureFields<T extends object>(input: T, fields: readonly (keyof T)[]): T {
  const captured = { ...input };
  for (const field of fields) {
    if (Object.hasOwn(captured, field)) continue;
    const value = input[field];
    if (value !== undefined) captured[field] = value;
  }
  return captured;
}
