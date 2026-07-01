namespace OfficeIMO.Excel.LegacyXls.Model {
    /// <summary>
    /// Represents a parsed legacy XLS cell using one-based row and column indexes.
    /// </summary>
    public sealed class LegacyXlsCell {
        /// <summary>
        /// Creates a parsed legacy XLS cell.
        /// </summary>
        /// <param name="row">One-based row index.</param>
        /// <param name="column">One-based column index.</param>
        /// <param name="kind">Cell value kind.</param>
        /// <param name="value">Parsed cell value.</param>
        /// <param name="styleIndex">Legacy XF style index.</param>
        /// <param name="isFormula">Whether the cell was parsed from a formula record.</param>
        /// <param name="formulaText">Decoded formula text, when token decoding was supported.</param>
        /// <param name="textFormattingRuns">Rich-text formatting runs for text cells.</param>
        public LegacyXlsCell(
            int row,
            int column,
            LegacyXlsCellValueKind kind,
            object? value,
            ushort styleIndex,
            bool isFormula = false,
            string? formulaText = null,
            IReadOnlyList<LegacyXlsTextFormattingRun>? textFormattingRuns = null) {
            Row = row;
            Column = column;
            Kind = kind;
            Value = value;
            StyleIndex = styleIndex;
            IsFormula = isFormula;
            FormulaText = formulaText;
            TextFormattingRuns = textFormattingRuns ?? Array.Empty<LegacyXlsTextFormattingRun>();
        }

        /// <summary>
        /// Gets the one-based row index.
        /// </summary>
        public int Row { get; }

        /// <summary>
        /// Gets the one-based column index.
        /// </summary>
        public int Column { get; }

        /// <summary>
        /// Gets the value kind.
        /// </summary>
        public LegacyXlsCellValueKind Kind { get; }

        /// <summary>
        /// Gets the parsed value.
        /// </summary>
        public object? Value { get; }

        /// <summary>
        /// Gets the legacy XF style index associated with the cell.
        /// </summary>
        public ushort StyleIndex { get; }

        /// <summary>
        /// Gets whether this cell came from a BIFF formula record and contains its cached result.
        /// </summary>
        public bool IsFormula { get; }

        /// <summary>
        /// Gets decoded formula text when the BIFF token stream was supported.
        /// </summary>
        public string? FormulaText { get; }

        /// <summary>
        /// Gets rich-text formatting runs for text cells.
        /// </summary>
        public IReadOnlyList<LegacyXlsTextFormattingRun> TextFormattingRuns { get; }
    }
}
