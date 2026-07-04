namespace OfficeIMO.Excel.LegacyXls.Model {
    /// <summary>
    /// Preserve-only metadata for a conditional-formatting extension BIFF record.
    /// </summary>
    public sealed class LegacyXlsConditionalFormattingExtensionRecord {
        /// <summary>
        /// Creates preserve-only conditional-formatting extension metadata.
        /// </summary>
        public LegacyXlsConditionalFormattingExtensionRecord(
            string sheetName,
            int recordOffset,
            ushort recordType,
            int payloadLength,
            bool isCf12,
            ushort? headerId,
            ushort? ruleIndex,
            int? priority,
            bool? stopIfTrue,
            bool hasUnprojectedFormatting,
            bool matchedRule)
            : this(
                sheetName,
                recordOffset,
                recordType,
                payloadLength,
                isCf12,
                headerId,
                ruleIndex,
                priority,
                stopIfTrue,
                hasUnprojectedFormatting,
                matchedRule,
                inlineFormattingByteCount: null) {
        }

        /// <summary>
        /// Creates preserve-only conditional-formatting extension metadata.
        /// </summary>
        public LegacyXlsConditionalFormattingExtensionRecord(
            string sheetName,
            int recordOffset,
            ushort recordType,
            int payloadLength,
            bool isCf12,
            ushort? headerId,
            ushort? ruleIndex,
            int? priority,
            bool? stopIfTrue,
            bool hasUnprojectedFormatting,
            bool matchedRule,
            int? inlineFormattingByteCount) {
            if (payloadLength < 0) {
                throw new ArgumentOutOfRangeException(nameof(payloadLength));
            }

            if (inlineFormattingByteCount < 0) {
                throw new ArgumentOutOfRangeException(nameof(inlineFormattingByteCount));
            }

            SheetName = sheetName ?? throw new ArgumentNullException(nameof(sheetName));
            RecordOffset = recordOffset;
            RecordType = recordType;
            PayloadLength = payloadLength;
            IsCf12 = isCf12;
            HeaderId = headerId;
            RuleIndex = ruleIndex;
            Priority = priority;
            StopIfTrue = stopIfTrue;
            HasUnprojectedFormatting = hasUnprojectedFormatting;
            MatchedRule = matchedRule;
            InlineFormattingByteCount = inlineFormattingByteCount;
        }

        /// <summary>Gets the worksheet name associated with the extension record.</summary>
        public string SheetName { get; }

        /// <summary>Gets the byte offset of the BIFF record.</summary>
        public int RecordOffset { get; }

        /// <summary>Gets the BIFF record type identifier.</summary>
        public ushort RecordType { get; }

        /// <summary>Gets the BIFF record payload length in bytes.</summary>
        public int PayloadLength { get; }

        /// <summary>Gets whether the extension declares a CF12 rule payload shape.</summary>
        public bool IsCf12 { get; }

        /// <summary>Gets the conditional-formatting header identifier, when decoded.</summary>
        public ushort? HeaderId { get; }

        /// <summary>Gets the zero-based conditional-formatting rule index, when decoded.</summary>
        public ushort? RuleIndex { get; }

        /// <summary>Gets the extension priority, when decoded.</summary>
        public int? Priority { get; }

        /// <summary>Gets whether the extension requests stop-if-true behavior, when decoded.</summary>
        public bool? StopIfTrue { get; }

        /// <summary>Gets whether the extension declares formatting bytes that are not fully projected yet.</summary>
        public bool HasUnprojectedFormatting { get; private set; }

        /// <summary>Gets whether extension formatting was projected onto the matched conditional-formatting rule.</summary>
        public bool HasProjectedFormatting { get; private set; }

        /// <summary>Gets the decoded inline formatting byte count declared by the extension, when present.</summary>
        public int? InlineFormattingByteCount { get; }

        /// <summary>Gets whether the extension was matched to a parsed conditional-formatting rule.</summary>
        public bool MatchedRule { get; private set; }

        internal void MarkMatchedRule() {
            MatchedRule = true;
        }

        internal void MarkProjectedFormatting() {
            HasProjectedFormatting = true;
            HasUnprojectedFormatting = false;
        }
    }
}
