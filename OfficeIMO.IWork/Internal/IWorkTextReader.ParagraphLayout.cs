namespace OfficeIMO.IWork.Internal;

internal static partial class IWorkTextReader {
    /// <summary>Assesses selected layout messages that have no shared paragraph representation.</summary>
    private static void AssessUnmappedParagraphLayout(IWorkWireMessage message, IWorkArchiveRecord record,
        IWorkSourceReferenceIssueCollector references, ref bool complete) {
        var evidence = new StylePropertyEvidence(record, "12/", references.Declarations);
        // TSWP.ParagraphStylePropertiesArchive: line spacing and custom tabs.
        AssessMessage(12, 13, ref complete);
        AssessMessage(24, 25, ref complete);

        void AssessMessage(int clearField, int valueField, ref bool isComplete) {
            bool? clear = ReadBoolean(message, clearField, evidence, ref isComplete);
            if (!message.HasField(valueField)) return;
            if (clear == true || message.FieldCount(valueField) != 1
                || message.HasUnexpectedWireKind(valueField, IWorkWireKind.Bytes)) {
                evidence.Record(message, valueField);
                isComplete = false;
                return;
            }
            IWorkWireMessage? value = IWorkObjectIndex.TryGetMessage(message, valueField, out bool malformed);
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
}
