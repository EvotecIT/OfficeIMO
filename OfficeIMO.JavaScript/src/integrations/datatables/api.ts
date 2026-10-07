/** @internal Runtime-checked calls across the optional third-party API boundary. */
export function call(owner: unknown, name: string, ...args: unknown[]): unknown {
  const fn = member(owner, name);
  if (typeof fn !== "function") throw new TypeError("DataTables API requires " + name + "().");
  return Reflect.apply(fn, owner, args);
}
/** @internal */
export function member(owner: unknown, name: string): unknown {
  return owner != null && (typeof owner === "object" || typeof owner === "function")
    ? (owner as Record<string, unknown>)[name] : undefined;
}
/** @internal */
export function array(value: unknown): unknown[] {
  const result = Array.isArray(value) ? value : call(value, "toArray");
  if (!Array.isArray(result)) throw new TypeError("DataTables API must return an array.");
  return result;
}
/** @internal */
export function indexes(value: unknown): number[] {
  return array(call(value, "indexes")).map(index => {
    if (!Number.isSafeInteger(index) || (index as number) < 0) throw new TypeError("Invalid DataTables index.");
    return index as number;
  });
}
