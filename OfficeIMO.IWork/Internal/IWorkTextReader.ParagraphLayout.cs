namespace OfficeIMO.IWork.Internal;

internal static partial class IWorkTextReader {
    /// <summary>Recovers qualified relative spacing and assesses remaining selected layout messages.</summary>
    private static void AssessParagraphLayout(IWorkWireMessage message, ParagraphStyleData data, IWorkProjectionBudget budget, IWorkArchiveRecord record,
        IWorkSourceReferenceIssueCollector references, ref bool complete) {
        var evidence = new StylePropertyEvidence(record, "12/", references.Declarations);
        // TSWP.ParagraphStylePropertiesArchive: line spacing and custom tabs.
        AssessMessage(12, 13, ref complete);
        AssessMessage(24, 25, ref complete);

        void AssessMessage(int clearField, int valueField, ref bool isComplete) {
            bool? clear = ReadBoolean(message, clearField, evidence, ref isComplete);
            if (valueField == 13 && (clear == true || message.HasField(valueField)))
                data.LineSpacingMultiplier = null;
            if (valueField == 25 && (clear == true || message.HasField(valueField)))
                data.TabStops = Array.Empty<IWorkTabStop>();
            if (!message.HasField(valueField)) return;
            if (clear == true || message.FieldCount(valueField) != 1
                || message.HasUnexpectedWireKind(valueField, IWorkWireKind.Bytes)) {
                evidence.Record(message, valueField);
                isComplete = false;
                return;
            }
            IWorkWireMessage? value = IWorkObjectIndex.TryGetMessage(message, valueField, out bool malformed);
            if (!malformed && value != null && valueField == 13
                && (!message.HasField(clearField) || clear.HasValue)
                && TryRelativeLineSpacing(value, out double multiplier)) {
                data.LineSpacingMultiplier = multiplier;
                return;
            }
            if (!malformed && value != null && valueField == 25
                && (!message.HasField(clearField) || clear.HasValue)
                && TryTabStops(value, budget, out var tabs)) {
                data.TabStops = tabs;
                return;
            }
            references.Declarations.Record(record,
                "12/" + valueField.ToString(System.Globalization.CultureInfo.InvariantCulture),
                message.FieldCount(valueField), malformed || value == null
                    ? IWorkSourceDeclarationIssueKind.MalformedMessage
                    : IWorkSourceDeclarationIssueKind.UnsupportedField);
            // Empty messages remain unassessed: do not invent default semantics or
            // let descendant resets erase an unsupported ancestor's declaration.
            isComplete = false;
        }
    }
    private static bool TryRelativeLineSpacing(IWorkWireMessage value, out double multiplier) {
        multiplier = 0;
        // Absent mode is the protobuf relative-mode default. An absent amount is
        // not a qualified single-line default. Unknown extensions remain unassessed.
        if (value.FieldCount(1) > 1 || value.HasUnexpectedWireKind(1, IWorkWireKind.Varint)
            || value.GetUnsigned(1).GetValueOrDefault() != 0
            || value.FieldCount(2) != 1 || value.HasUnexpectedWireKind(2, IWorkWireKind.Fixed32)
            || value.TotalFieldCount != value.FieldCount(1) + 1) return false;
        float? amount = value.GetFloat(2);
        if (!amount.HasValue || !IsFinitePositive(amount.Value)) return false;
        multiplier = amount.Value;
        return true;
    }
}
