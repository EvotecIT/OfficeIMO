namespace OfficeIMO.Excel.LegacyXls.Model {
    /// <summary>
    /// Describes one BIFF parsed-formula token observed while importing a legacy XLS workbook.
    /// </summary>
    public sealed class LegacyXlsFormulaTokenRecord {
        /// <summary>
        /// Creates formula token metadata for corpus diagnostics and XLS import planning.
        /// </summary>
        public LegacyXlsFormulaTokenRecord(
            string context,
            string? sheetName,
            string? cellReference,
            int recordOffset,
            ushort recordType,
            byte token,
            string tokenName,
            int tokenOffset,
            ushort? functionId = null,
            string? functionName = null,
            byte? attribute = null,
            string? attributeName = null) {
            Context = context ?? throw new ArgumentNullException(nameof(context));
            SheetName = sheetName;
            CellReference = cellReference;
            RecordOffset = recordOffset;
            RecordType = recordType;
            Token = token;
            TokenName = tokenName ?? throw new ArgumentNullException(nameof(tokenName));
            TokenOffset = tokenOffset;
            FunctionId = functionId;
            FunctionName = functionName;
            Attribute = attribute;
            AttributeName = attributeName;
        }

        /// <summary>
        /// Gets the formula source context, such as CellFormula, SharedFormula, or DefinedName.
        /// </summary>
        public string Context { get; }

        /// <summary>
        /// Gets the worksheet name associated with the token, when known.
        /// </summary>
        public string? SheetName { get; }

        /// <summary>
        /// Gets the formula cell reference associated with the token, when known.
        /// </summary>
        public string? CellReference { get; }

        /// <summary>
        /// Gets the byte offset of the BIFF record that supplied the formula token stream.
        /// </summary>
        public int RecordOffset { get; }

        /// <summary>
        /// Gets the BIFF record type that supplied the formula token stream.
        /// </summary>
        public ushort RecordType { get; }

        /// <summary>
        /// Gets the raw BIFF parsed-formula token byte.
        /// </summary>
        public byte Token { get; }

        /// <summary>
        /// Gets the stable BIFF parsed-formula token name.
        /// </summary>
        public string TokenName { get; }

        /// <summary>
        /// Gets the zero-based token offset within the parsed-expression token stream.
        /// </summary>
        public int TokenOffset { get; }

        /// <summary>
        /// Gets the BIFF built-in function identifier for function tokens, when available.
        /// </summary>
        public ushort? FunctionId { get; }

        /// <summary>
        /// Gets the function name resolved from <see cref="FunctionId"/>, when known.
        /// </summary>
        public string? FunctionName { get; }

        /// <summary>
        /// Gets the PtgAttr attribute byte, when the token is an attribute token.
        /// </summary>
        public byte? Attribute { get; }

        /// <summary>
        /// Gets the stable attribute name for PtgAttr records, when available.
        /// </summary>
        public string? AttributeName { get; }
    }
}
