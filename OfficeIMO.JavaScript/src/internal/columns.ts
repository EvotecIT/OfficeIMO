import type { ColumnSettings, ColumnValueContext } from "../core/index.js";

type ValueKey<T, V> = { [K in Extract<keyof T, string>]-?: T[K] extends V ? K : never }[Extract<keyof T, string>];
type Selector<T, V> = [T] extends [never] ? { readonly key?: string; readonly value?: never }
  : T extends readonly unknown[] ? T[number] extends V ? { readonly key?: string; readonly value?: never } : never
  : { readonly key: ValueKey<T, V>; readonly value?: never };

/** One selector definition, with values admitted by the destination's public contract. */
export type ProjectedColumn<T, V> = ColumnSettings & { readonly key?: string } & (
  | Selector<T, V>
  | { readonly key?: string; readonly value: (row: T, context: ColumnValueContext) => V }
);

/** Erased metadata for owners that invoke the captured getter with the source row. */
export type ProjectionColumn = ColumnSettings & {
  readonly key?: string;
  readonly value?: (row: never, context: ColumnValueContext) => unknown;
};
