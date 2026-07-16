namespace OfficeIMO.Excel.Xlsb.Model {
    /// <summary>Represents one BIFF12 workbook defined name.</summary>
    internal sealed class XlsbDefinedName {
        internal XlsbDefinedName(
            uint flags,
            byte shortcutKey,
            uint localSheetIndex,
            string name,
            byte[] formulaTokens,
            byte[] formulaExtraBytes,
            string? comment,
            long recordOffset,
            int payloadLength) {
            Flags = flags;
            ShortcutKey = shortcutKey;
            LocalSheetIndex = localSheetIndex;
            Name = name ?? throw new ArgumentNullException(nameof(name));
            FormulaTokens = formulaTokens ?? throw new ArgumentNullException(nameof(formulaTokens));
            FormulaExtraBytes = formulaExtraBytes ?? throw new ArgumentNullException(nameof(formulaExtraBytes));
            Comment = comment;
            RecordOffset = recordOffset;
            PayloadLength = payloadLength;
        }

        internal uint Flags { get; }

        internal byte ShortcutKey { get; }

        internal uint LocalSheetIndex { get; }

        internal string Name { get; }

        internal byte[] FormulaTokens { get; }

        internal byte[] FormulaExtraBytes { get; }

        internal string? Comment { get; }

        internal long RecordOffset { get; }

        internal int PayloadLength { get; }

        internal bool Hidden => (Flags & 0x00000001U) != 0;

        internal bool BuiltIn => (Flags & 0x00000020U) != 0;

        internal bool IsSimpleName => (Flags & 0x0003FFDEU) == 0 && ShortcutKey == 0 && FormulaExtraBytes.Length == 0;

        internal string? FormulaText { get; set; }
    }
}
