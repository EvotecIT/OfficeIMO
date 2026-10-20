/** Explicit display strings for boolean values. XLSX boolean cell types remain unchanged. */
export interface BooleanText {
  readonly true: string;
  readonly false: string;
}

/** Compatibility display policy; hosts can supply labels without changing typed source values. */
export const defaultBooleanText: Readonly<BooleanText> = Object.freeze({ true: "True", false: "False" });

/** @internal Capture display policy once, before source iteration or asynchronous output. */
export function captureBooleanText(input: BooleanText | undefined): Readonly<BooleanText> {
  if (input === undefined) return defaultBooleanText;
  if (!input || typeof input !== "object") throw new TypeError("booleanText must contain true and false strings.");
  const yes = input.true, no = input.false;
  if (typeof yes !== "string" || typeof no !== "string") throw new TypeError("booleanText must contain true and false strings.");
  return Object.freeze({ true: yes, false: no });
}
