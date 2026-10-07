export function colorArgb(value: string): string {
  if (typeof value !== "string" || !/^#?(?:[0-9a-f]{6}|[0-9a-f]{8})$/i.test(value)) throw new TypeError("Color must be RGB or ARGB hex.");
  const color = value.replace(/^#/, "").toUpperCase(); return color.length === 6 ? "FF" + color : color;
}
