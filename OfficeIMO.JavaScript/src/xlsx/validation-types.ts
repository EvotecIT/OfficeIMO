/** An inline dropdown applied only to the selected exported data column, after its final row count is known. */
export interface ListValidation {
  readonly type: "list";
  /** One-based column number, or an unambiguous declared column key. */
  readonly column: string | number;
  /** Literal dropdown choices. Commas are not supported; the quoted list source must fit 255 UTF-16 units. */
  readonly items: readonly string[];
  /** Defaults to true, matching OfficeIMO.Excel.ValidationList. */
  readonly allowBlank?: boolean;
  /** Defaults to true. Excel applies the alert when the user edits a validated cell. */
  readonly showErrorMessage?: boolean;
  /** Defaults to stop. Warning and information alerts allow the user to accept a value outside the list. */
  readonly errorStyle?: "stop" | "warning" | "information";
  /** At most 32 XML-valid UTF-16 units. */
  readonly errorTitle?: string;
  /** At most 255 XML-valid UTF-16 units. */
  readonly errorMessage?: string;
}
