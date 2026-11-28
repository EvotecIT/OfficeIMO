import type { PdfPathCommand } from "../../pdf/document-types.js";

/** An affine matrix [a b c d e f]: x' = a x + c y + e, y' = b x + d y + f. */
export type Matrix = readonly [number, number, number, number, number, number];
export const identity: Matrix = [1, 0, 0, 1, 0, 0];

export function multiply(m: Matrix, n: Matrix): Matrix {
  return [m[0] * n[0] + m[2] * n[1], m[1] * n[0] + m[3] * n[1], m[0] * n[2] + m[2] * n[3], m[1] * n[2] + m[3] * n[3],
    m[0] * n[4] + m[2] * n[5] + m[4], m[1] * n[4] + m[3] * n[5] + m[5]];
}
export function apply(m: Matrix, x: number, y: number): [number, number] { return [m[0] * x + m[2] * y + m[4], m[1] * x + m[3] * y + m[5]]; }
export function scaleOf(m: Matrix): number { return Math.sqrt(Math.abs(m[0] * m[3] - m[1] * m[2])); }

const numberPattern = /[+-]?(?:\d+\.?\d*|\.\d+)(?:[eE][+-]?\d+)?/g;
export function numbers(text: string): number[] { return (text.match(numberPattern) ?? []).map(Number); }

/** Parse an SVG transform list. Returns undefined for syntax this reader does not understand. */
export function parseTransform(text: string | undefined): Matrix | undefined {
  if (text === undefined || !text.trim() || text.trim() === "none") return identity;
  let result = identity, rest = text.trim();
  const item = /^([a-zA-Z]+)\s*\(([^)]*)\)\s*,?\s*/;
  while (rest) {
    const match = item.exec(rest);
    if (!match) return undefined;
    const args = numbers(match[2]!.replace(/(deg|px)\b/g, "")), name = match[1]!;
    let next: Matrix;
    if (name === "matrix" && args.length === 6) next = args as unknown as Matrix;
    else if (name === "translate" && (args.length === 1 || args.length === 2)) next = [1, 0, 0, 1, args[0]!, args[1] ?? 0];
    else if (name === "scale" && (args.length === 1 || args.length === 2)) next = [args[0]!, 0, 0, args[1] ?? args[0]!, 0, 0];
    else if (name === "rotate" && (args.length === 1 || args.length === 3)) {
      const angle = args[0]! * Math.PI / 180, cos = Math.cos(angle), sin = Math.sin(angle), cx = args[1] ?? 0, cy = args[2] ?? 0;
      next = multiply(multiply([1, 0, 0, 1, cx, cy], [cos, sin, -sin, cos, 0, 0]), [1, 0, 0, 1, -cx, -cy]);
    } else if (name === "skewX" && args.length === 1) next = [1, 0, Math.tan(args[0]! * Math.PI / 180), 1, 0, 0];
    else if (name === "skewY" && args.length === 1) next = [1, Math.tan(args[0]! * Math.PI / 180), 0, 1, 0, 0];
    else return undefined;
    result = multiply(result, next); rest = rest.slice(match[0].length);
  }
  return result.every(Number.isFinite) ? result : undefined;
}

/** Absolute user-space commands; quadratic curves and arcs are already cubic. */
export type UserCommand = PdfPathCommand;

function arc(out: UserCommand[], x1: number, y1: number, rx: number, ry: number, rotation: number, large: boolean, sweep: boolean, x2: number, y2: number): void {
  if (x1 === x2 && y1 === y2) return;
  rx = Math.abs(rx); ry = Math.abs(ry);
  if (!rx || !ry) { out.push({ op: "line", x: x2, y: y2 }); return; }
  const phi = rotation * Math.PI / 180, cos = Math.cos(phi), sin = Math.sin(phi);
  const dx = (x1 - x2) / 2, dy = (y1 - y2) / 2, xp = cos * dx + sin * dy, yp = -sin * dx + cos * dy;
  const lambda = xp * xp / (rx * rx) + yp * yp / (ry * ry);
  if (lambda > 1) { const s = Math.sqrt(lambda); rx *= s; ry *= s; }
  const sign = large === sweep ? -1 : 1, numerator = rx * rx * ry * ry - rx * rx * yp * yp - ry * ry * xp * xp;
  const factor = sign * Math.sqrt(Math.max(0, numerator / (rx * rx * yp * yp + ry * ry * xp * xp)));
  const cxp = factor * rx * yp / ry, cyp = -factor * ry * xp / rx;
  const cx = cos * cxp - sin * cyp + (x1 + x2) / 2, cy = sin * cxp + cos * cyp + (y1 + y2) / 2;
  const angle = (ux: number, uy: number, vx: number, vy: number): number => Math.atan2(ux * vy - uy * vx, ux * vx + uy * vy);
  const start = angle(1, 0, (xp - cxp) / rx, (yp - cyp) / ry);
  let delta = angle((xp - cxp) / rx, (yp - cyp) / ry, (-xp - cxp) / rx, (-yp - cyp) / ry);
  if (!sweep && delta > 0) delta -= 2 * Math.PI; else if (sweep && delta < 0) delta += 2 * Math.PI;
  const segments = Math.max(1, Math.ceil(Math.abs(delta) / (Math.PI / 2) - 1e-9)), step = delta / segments, k = 4 / 3 * Math.tan(step / 4);
  const point = (t: number): [number, number] => [cx + rx * Math.cos(t) * cos - ry * Math.sin(t) * sin, cy + rx * Math.cos(t) * sin + ry * Math.sin(t) * cos];
  const derivative = (t: number): [number, number] => [-rx * Math.sin(t) * cos - ry * Math.cos(t) * sin, -rx * Math.sin(t) * sin + ry * Math.cos(t) * cos];
  for (let i = 0; i < segments; i++) {
    const t1 = start + i * step, t2 = t1 + step, [px1, py1] = point(t1), [px2, py2] = i === segments - 1 ? [x2, y2] : point(t2);
    const [d1x, d1y] = derivative(t1), [d2x, d2y] = derivative(t2);
    out.push({ op: "cubic", x1: px1 + k * d1x, y1: py1 + k * d1y, x2: px2 - k * d2x, y2: py2 - k * d2y, x: px2, y: py2 });
  }
}

/** Parse SVG path data. Stops at the first malformed segment, as SVG renderers do, and reports whether it did. */
export function parsePathData(d: string): { readonly commands: UserCommand[]; readonly malformed: boolean } {
  const out: UserCommand[] = [], tokens = /([MmLlHhVvCcSsQqTtAaZz])|([+-]?(?:\d+\.?\d*|\.\d+)(?:[eE][+-]?\d+)?)|([\s,]+)|(.)/g;
  const list: string[] = [];
  let malformed = false;
  for (const match of d.matchAll(tokens)) {
    if (match[4] !== undefined) { malformed = true; break; }
    if (match[1] !== undefined || match[2] !== undefined) list.push(match[1] ?? match[2]!);
  }
  let i = 0, x = 0, y = 0, startX = 0, startY = 0, command = "", lastControl: [number, number] | undefined, lastQuad: [number, number] | undefined;
  let open = false;
  const isCommand = (value: string | undefined): boolean => value !== undefined && /^[A-Za-z]$/.test(value);
  const take = (): number | undefined => {
    const value = list[i];
    if (value === undefined || isCommand(value)) return undefined;
    i++;
    const number = Number(value);
    return Number.isFinite(number) ? number : undefined;
  };
  // Arc flags may be written without separators ("a1 1 0 011 1"): split compact digit runs.
  const flag = (): number | undefined => {
    const raw = list[i];
    if (raw === undefined || isCommand(raw)) return undefined;
    if (raw === "0" || raw === "1") { i++; return Number(raw); }
    if (/^[01][\d.]/.test(raw)) { list[i] = raw.slice(1); return Number(raw[0]); }
    return undefined;
  };
  while (i < list.length) {
    const token = list[i];
    if (isCommand(token)) { command = token!; i++; }
    else if (!command || command === "Z" || command === "z") return { commands: out, malformed: true };
    const relative = command === command.toLowerCase(), base = command.toUpperCase();
    if (!out.length && base !== "M") return { commands: out, malformed: true };
    const ox = relative ? x : 0, oy = relative ? y : 0;
    if (base === "Z") {
      if (open) out.push({ op: "close" });
      x = startX; y = startY; open = false; lastControl = lastQuad = undefined; continue;
    }
    const ensure = (): void => { if (!open) { out.push({ op: "move", x, y }); open = true; } };
    if (base === "M") {
      const a = take(), b = take(); if (a === undefined || b === undefined) return { commands: out, malformed: true };
      x = ox + a; y = oy + b; startX = x; startY = y; out.push({ op: "move", x, y }); open = true;
      command = relative ? "l" : "L"; lastControl = lastQuad = undefined; continue;
    }
    if (base === "L" || base === "H" || base === "V") {
      const a = take(); if (a === undefined) return { commands: out, malformed: true };
      const b = base === "L" ? take() : undefined;
      if (base === "L" && b === undefined) return { commands: out, malformed: true };
      ensure();
      if (base !== "V") x = ox + a;
      if (base !== "H") y = oy + (b ?? a);
      out.push({ op: "line", x, y }); lastControl = lastQuad = undefined; continue;
    }
    if (base === "C" || base === "S") {
      const values: number[] = [];
      for (let n = base === "C" ? 6 : 4; n > 0; n--) { const v = take(); if (v === undefined) return { commands: out, malformed: true }; values.push(v); }
      const [c1x, c1y] = base === "C" ? [ox + values[0]!, oy + values[1]!] : lastControl ? [2 * x - lastControl[0], 2 * y - lastControl[1]] : [x, y];
      const rest = base === "C" ? values.slice(2) : values;
      const c2x = ox + rest[0]!, c2y = oy + rest[1]!, ex = ox + rest[2]!, ey = oy + rest[3]!;
      ensure(); out.push({ op: "cubic", x1: c1x, y1: c1y, x2: c2x, y2: c2y, x: ex, y: ey });
      lastControl = [c2x, c2y]; lastQuad = undefined; x = ex; y = ey; continue;
    }
    if (base === "Q" || base === "T") {
      const values: number[] = [];
      for (let n = base === "Q" ? 4 : 2; n > 0; n--) { const v = take(); if (v === undefined) return { commands: out, malformed: true }; values.push(v); }
      const [qx, qy] = base === "Q" ? [ox + values[0]!, oy + values[1]!] : lastQuad ? [2 * x - lastQuad[0], 2 * y - lastQuad[1]] : [x, y];
      const ex = ox + values[values.length - 2]!, ey = oy + values[values.length - 1]!;
      ensure(); out.push({ op: "cubic", x1: x + 2 / 3 * (qx - x), y1: y + 2 / 3 * (qy - y), x2: ex + 2 / 3 * (qx - ex), y2: ey + 2 / 3 * (qy - ey), x: ex, y: ey });
      lastQuad = [qx, qy]; lastControl = undefined; x = ex; y = ey; continue;
    }
    if (base === "A") {
      const rx = take(), ry = take(), rotation = take(), large = flag(), sweep = flag(), a = take(), b = take();
      if (rx === undefined || ry === undefined || rotation === undefined || large === undefined || sweep === undefined || a === undefined || b === undefined)
        return { commands: out, malformed: true };
      ensure(); arc(out, x, y, rx, ry, rotation, large === 1, sweep === 1, ox + a, oy + b);
      x = ox + a; y = oy + b; lastControl = lastQuad = undefined; continue;
    }
    return { commands: out, malformed: true };
  }
  return { commands: out, malformed };
}

const kappa = 0.5522847498307936;
/** Closed ellipse as four cubic quarters, starting at the rightmost point. */
export function ellipse(cx: number, cy: number, rx: number, ry: number): UserCommand[] {
  const ox = rx * kappa, oy = ry * kappa;
  return [{ op: "move", x: cx + rx, y: cy },
    { op: "cubic", x1: cx + rx, y1: cy + oy, x2: cx + ox, y2: cy + ry, x: cx, y: cy + ry },
    { op: "cubic", x1: cx - ox, y1: cy + ry, x2: cx - rx, y2: cy + oy, x: cx - rx, y: cy },
    { op: "cubic", x1: cx - rx, y1: cy - oy, x2: cx - ox, y2: cy - ry, x: cx, y: cy - ry },
    { op: "cubic", x1: cx + ox, y1: cy - ry, x2: cx + rx, y2: cy - oy, x: cx + rx, y: cy }, { op: "close" }];
}
export function rectangle(x: number, y: number, width: number, height: number, rx: number, ry: number): UserCommand[] {
  rx = Math.min(Math.max(rx, 0), width / 2); ry = Math.min(Math.max(ry, 0), height / 2);
  if (!rx || !ry) return [{ op: "move", x, y }, { op: "line", x: x + width, y }, { op: "line", x: x + width, y: y + height }, { op: "line", x, y: y + height }, { op: "close" }];
  const ox = rx * (1 - kappa), oy = ry * (1 - kappa), r = x + width, b = y + height;
  return [{ op: "move", x: x + rx, y }, { op: "line", x: r - rx, y }, { op: "cubic", x1: r - ox, y1: y, x2: r, y2: y + oy, x: r, y: y + ry },
    { op: "line", x: r, y: b - ry }, { op: "cubic", x1: r, y1: b - oy, x2: r - ox, y2: b, x: r - rx, y: b },
    { op: "line", x: x + rx, y: b }, { op: "cubic", x1: x + ox, y1: b, x2: x, y2: b - oy, x, y: b - ry },
    { op: "line", x, y: y + ry }, { op: "cubic", x1: x, y1: y + oy, x2: x + ox, y2: y, x: x + rx, y }, { op: "close" }];
}
export function polyline(points: readonly number[], closed: boolean): UserCommand[] {
  const out: UserCommand[] = [];
  for (let i = 0; i + 1 < points.length; i += 2) out.push({ op: i ? "line" : "move", x: points[i]!, y: points[i + 1]! });
  if (closed && out.length) out.push({ op: "close" });
  return out;
}
export function transformCommands(commands: readonly UserCommand[], m: Matrix): PdfPathCommand[] {
  return commands.map(c => {
    if (c.op === "close") return c;
    const [x, y] = apply(m, c.x, c.y);
    if (c.op === "cubic") { const [x1, y1] = apply(m, c.x1, c.y1), [x2, y2] = apply(m, c.x2, c.y2); return { op: "cubic", x1, y1, x2, y2, x, y }; }
    return { op: c.op, x, y };
  });
}
