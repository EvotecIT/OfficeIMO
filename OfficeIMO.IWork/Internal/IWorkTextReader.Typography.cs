namespace OfficeIMO.IWork.Internal;

internal static partial class IWorkTextReader {
    /// <summary>Retains selected typography that the shared text model cannot reconstruct.</summary>
    private static void AssessUnmappedTypography(IWorkWireMessage message, IWorkArchiveRecord record,
        IWorkSourceDeclarationIssueCollector declarations, ref bool complete) {
        bool unmapped = false;
        // TSWP.CharacterStylePropertiesArchive: script, capitalization, baseline shift, kerning, tracking.
        AssessEnum(10, 2);
        AssessEnum(13, 3);
        AssessFloat(14);
        AssessFloat(15);
        AssessFloat(27);
        if (unmapped) complete = false;

        void AssessEnum(int field, ulong maximum) {
            if (!message.HasField(field)) return;
            ulong? value = message.GetUnsigned(field);
            bool invalid = message.FieldCount(field) != 1
                || message.HasUnexpectedWireKind(field, IWorkWireKind.Varint)
                || !value.HasValue || value.Value > maximum;
            if (invalid || value != 0) Record(field, invalid);
        }

        void AssessFloat(int field) {
            if (!message.HasField(field)) return;
            float? value = message.GetFloat(field);
            bool invalid = message.FieldCount(field) != 1
                || message.HasUnexpectedWireKind(field, IWorkWireKind.Fixed32)
                || !value.HasValue || !IsFinite(value.Value);
            if (invalid || value != 0) Record(field, invalid);
        }

        void Record(int field, bool invalid) {
            declarations.Record(record, "11/" + field.ToString(System.Globalization.CultureInfo.InvariantCulture),
                message.FieldCount(field), invalid ? IWorkSourceDeclarationIssueKind.InvalidValue
                    : IWorkSourceDeclarationIssueKind.UnsupportedField);
            unmapped = true;
        }
    }
}
