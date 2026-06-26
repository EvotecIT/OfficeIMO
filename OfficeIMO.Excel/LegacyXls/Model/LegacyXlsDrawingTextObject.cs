namespace OfficeIMO.Excel.LegacyXls.Model {
    /// <summary>
    /// Describes the fixed header of a legacy BIFF TxO text-object record.
    /// </summary>
    public sealed class LegacyXlsDrawingTextObject {
        /// <summary>
        /// Creates TxO text-object metadata.
        /// </summary>
        public LegacyXlsDrawingTextObject(
            ushort rawOptions,
            ushort rotation,
            ushort textCharacterCount,
            ushort formattingRunByteCount,
            ushort emptyFontIndex,
            int formulaByteCount) {
            if (formulaByteCount < 0) {
                throw new ArgumentOutOfRangeException(nameof(formulaByteCount));
            }

            RawOptions = rawOptions;
            HorizontalAlignment = (ushort)((rawOptions >> 1) & 0x0007);
            HorizontalAlignmentName = GetHorizontalAlignmentName(HorizontalAlignment);
            VerticalAlignment = (ushort)((rawOptions >> 4) & 0x0007);
            VerticalAlignmentName = GetVerticalAlignmentName(VerticalAlignment);
            LockedText = (rawOptions & 0x0200) != 0;
            JustifyLastLine = (rawOptions & 0x4000) != 0;
            SecretEdit = (rawOptions & 0x8000) != 0;
            Rotation = rotation;
            RotationName = GetRotationName(rotation);
            TextCharacterCount = textCharacterCount;
            FormattingRunByteCount = formattingRunByteCount;
            EmptyFontIndex = emptyFontIndex;
            FormulaByteCount = formulaByteCount;
        }

        /// <summary>Gets the raw TxO option bitfield.</summary>
        public ushort RawOptions { get; }

        /// <summary>Gets the decoded horizontal alignment value.</summary>
        public ushort HorizontalAlignment { get; }

        /// <summary>Gets a stable display name for the horizontal alignment value.</summary>
        public string HorizontalAlignmentName { get; }

        /// <summary>Gets the decoded vertical alignment value.</summary>
        public ushort VerticalAlignment { get; }

        /// <summary>Gets a stable display name for the vertical alignment value.</summary>
        public string VerticalAlignmentName { get; }

        /// <summary>Gets whether the TxO text is locked.</summary>
        public bool LockedText { get; }

        /// <summary>Gets whether the last line is justified.</summary>
        public bool JustifyLastLine { get; }

        /// <summary>Gets whether secret edit mode is enabled.</summary>
        public bool SecretEdit { get; }

        /// <summary>Gets the raw TxO rotation value.</summary>
        public ushort Rotation { get; }

        /// <summary>Gets a stable display name for the rotation value.</summary>
        public string RotationName { get; }

        /// <summary>Gets the declared character count of text stored in following Continue records.</summary>
        public ushort TextCharacterCount { get; }

        /// <summary>Gets the declared byte count of formatting runs stored in following Continue records.</summary>
        public ushort FormattingRunByteCount { get; }

        /// <summary>Gets the font index used when the text is empty.</summary>
        public ushort EmptyFontIndex { get; }

        /// <summary>Gets the number of bytes available for the optional TxO formula field.</summary>
        public int FormulaByteCount { get; }

        /// <summary>Gets whether this TxO declares text payload in following Continue records.</summary>
        public bool HasTextInContinueRecords => TextCharacterCount > 0;

        /// <summary>Gets whether this TxO declares formatting-run payload in following Continue records.</summary>
        public bool HasFormattingRunsInContinueRecords => FormattingRunByteCount > 0;

        private static string GetHorizontalAlignmentName(ushort value) {
            return value switch {
                1 => "Left",
                2 => "Center",
                3 => "Right",
                4 => "Justify",
                7 => "Distributed",
                _ => $"Unknown:{value}"
            };
        }

        private static string GetVerticalAlignmentName(ushort value) {
            return value switch {
                1 => "Top",
                2 => "Middle",
                3 => "Bottom",
                4 => "Justify",
                7 => "Distributed",
                _ => $"Unknown:{value}"
            };
        }

        private static string GetRotationName(ushort value) {
            return value switch {
                0 => "None",
                1 => "StackedOrVertical",
                2 => "RotatedCounterClockwise90",
                3 => "RotatedClockwise90",
                _ => $"Unknown:{value}"
            };
        }
    }
}
