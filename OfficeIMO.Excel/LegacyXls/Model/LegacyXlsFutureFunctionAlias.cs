namespace OfficeIMO.Excel.LegacyXls.Model {
    /// <summary>
    /// Describes a future Excel function alias carried by a BIFF defined-name record.
    /// </summary>
    public sealed class LegacyXlsFutureFunctionAlias {
        /// <summary>
        /// Creates metadata for an Excel future-function compatibility alias.
        /// </summary>
        public LegacyXlsFutureFunctionAlias(
            string name,
            string functionName,
            int recordOffset,
            ushort recordType,
            byte? formulaToken,
            string? formulaTokenName,
            int? formulaTokenOffset) {
            Name = name ?? throw new ArgumentNullException(nameof(name));
            FunctionName = functionName ?? throw new ArgumentNullException(nameof(functionName));
            RecordOffset = recordOffset;
            RecordType = recordType;
            FormulaToken = formulaToken;
            FormulaTokenName = formulaTokenName;
            FormulaTokenOffset = formulaTokenOffset;
        }

        /// <summary>
        /// Gets the BIFF defined-name text, including the _xlfn. prefix.
        /// </summary>
        public string Name { get; }

        /// <summary>
        /// Gets the Excel future function name without the _xlfn. compatibility prefix.
        /// </summary>
        public string FunctionName { get; }

        /// <summary>
        /// Gets the byte offset of the BIFF Lbl record that carried the alias.
        /// </summary>
        public int RecordOffset { get; }

        /// <summary>
        /// Gets the BIFF record type that carried the alias.
        /// </summary>
        public ushort RecordType { get; }

        /// <summary>
        /// Gets the parsed-formula token byte used by Excel for the alias body, when available.
        /// </summary>
        public byte? FormulaToken { get; }

        /// <summary>
        /// Gets the parsed-formula token name used by Excel for the alias body, when available.
        /// </summary>
        public string? FormulaTokenName { get; }

        /// <summary>
        /// Gets the zero-based parsed-expression offset of the alias token, when available.
        /// </summary>
        public int? FormulaTokenOffset { get; }
    }
}
