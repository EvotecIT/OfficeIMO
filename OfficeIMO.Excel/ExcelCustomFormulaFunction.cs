namespace OfficeIMO.Excel {
    /// <summary>
    /// Identifies the value kind passed to or returned by a custom formula function.
    /// </summary>
    public enum ExcelFormulaValueKind {
        /// <summary>The formula value is blank.</summary>
        Blank,

        /// <summary>The formula value is numeric.</summary>
        Number,

        /// <summary>The formula value is text.</summary>
        Text,

        /// <summary>The formula value is an Excel error such as <c>#N/A</c>.</summary>
        Error
    }

    /// <summary>
    /// Typed scalar value used by custom formula functions.
    /// </summary>
    public readonly struct ExcelFormulaValue {
        private ExcelFormulaValue(ExcelFormulaValueKind kind, double number, string? text) {
            Kind = kind;
            Number = number;
            Text = text;
        }

        /// <summary>Blank formula value.</summary>
        public static ExcelFormulaValue Blank => default;

        /// <summary>Gets the value kind.</summary>
        public ExcelFormulaValueKind Kind { get; }

        /// <summary>Gets the number when <see cref="Kind"/> is <see cref="ExcelFormulaValueKind.Number"/>.</summary>
        public double Number { get; }

        /// <summary>Gets the text or error code for text and error values.</summary>
        public string? Text { get; }

        /// <summary>Creates a finite numeric formula value.</summary>
        public static ExcelFormulaValue FromNumber(double value) {
            if (double.IsNaN(value) || double.IsInfinity(value)) {
                throw new ArgumentOutOfRangeException(nameof(value), "Custom formula numbers must be finite.");
            }

            return new ExcelFormulaValue(ExcelFormulaValueKind.Number, value, null);
        }

        /// <summary>Creates a text formula value.</summary>
        public static ExcelFormulaValue FromText(string? value) {
            return new ExcelFormulaValue(ExcelFormulaValueKind.Text, 0d, value ?? string.Empty);
        }

        /// <summary>Creates an Excel error formula value.</summary>
        public static ExcelFormulaValue FromError(string errorCode) {
            if (string.IsNullOrWhiteSpace(errorCode)) {
                throw new ArgumentNullException(nameof(errorCode));
            }

            string normalized = errorCode.Trim();
            if (normalized.Length > 255 || normalized[0] != '#' || normalized.Any(char.IsWhiteSpace)) {
                throw new ArgumentException("Custom formula error codes must be compact Excel error literals that start with '#'.", nameof(errorCode));
            }

            return new ExcelFormulaValue(ExcelFormulaValueKind.Error, 0d, normalized);
        }
    }

    /// <summary>
    /// Read-only workbook location supplied to a custom formula function.
    /// </summary>
    public sealed class ExcelCustomFormulaFunctionContext {
        internal ExcelCustomFormulaFunctionContext(
            ExcelDocument workbook,
            ExcelSheet worksheet,
            string functionName,
            string? cellReference) {
            Workbook = workbook;
            Worksheet = worksheet;
            FunctionName = functionName;
            CellReference = cellReference;
        }

        /// <summary>Gets the workbook being evaluated.</summary>
        public ExcelDocument Workbook { get; }

        /// <summary>Gets the worksheet containing the formula.</summary>
        public ExcelSheet Worksheet { get; }

        /// <summary>Gets the normalized custom function name.</summary>
        public string FunctionName { get; }

        /// <summary>Gets the formula cell reference when evaluation originates from a cell.</summary>
        public string? CellReference { get; }
    }

    /// <summary>
    /// Evaluates a registered custom formula function. Return <see langword="null"/> when the supplied argument shape is unsupported.
    /// </summary>
    /// <remarks>
    /// Range arguments are expanded into scalar values in row-major workbook order. Callbacks should be deterministic and free of workbook mutations.
    /// </remarks>
    public delegate ExcelFormulaValue? ExcelCustomFormulaFunction(
        ExcelCustomFormulaFunctionContext context,
        IReadOnlyList<ExcelFormulaValue> arguments);
}
