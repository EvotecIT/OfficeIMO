using OfficeIMO.IWork.Internal;

namespace OfficeIMO.IWork;

internal static partial class IWorkKeynoteReader {
    private static void AssessTransition(IWorkWireMessage slideMessage, IWorkArchiveRecord slide,
        IWorkSourceReferenceIssueCollector references, List<IWorkDiagnostic> diagnostics,
        ref bool complete) {
        if (!slideMessage.HasField(4)) return;
        string path = "4";
        int? count = slideMessage.FieldCount(4);
        IWorkSourceDeclarationIssueKind kind = IWorkSourceDeclarationIssueKind.UnsupportedField;
        try {
            if (IsManualNoTransition()) return;
        } catch (InvalidDataException exception) when (!IWorkProtobuf.IsLimitException(exception)) {
            kind = IWorkSourceDeclarationIssueKind.MalformedMessage;
        }
        complete = false;
        references.Declarations.Record(slide, path, count, kind);
        diagnostics.Add(new IWorkDiagnostic(IWorkDiagnosticSeverity.Warning,
            "IWORK_KEYNOTE_TRANSITION_UNSUPPORTED",
            "A selected Keynote slide has an effect, automatic advance, or unqualified transition declaration that is not reconstructed; editable conversion is incomplete.",
            slide.EntryPath, slide.Identifier));

        bool IsManualNoTransition() {
            IWorkWireMessage? transition = Child(slideMessage, 4, "4");
            if (transition == null) return false;
            if (transition.TotalFieldCount != transition.FieldCount(2)) return false;
            IWorkWireMessage? attributes = Child(transition, 2, "4/2");
            if (attributes == null) return false;
            if (attributes.TotalFieldCount != attributes.FieldCount(8)) return false;
            IWorkWireMessage? animation = Child(attributes, 8, "4/2/8");
            if (animation == null) return false;
            int knownCount = 0;
            foreach (int field in new[] { 1, 2, 3, 5, 6, 11, 16 }) {
                knownCount += animation.FieldCount(field);
                IWorkWireKind wire = field is 1 or 2 ? IWorkWireKind.Bytes
                    : field is 3 or 5 ? IWorkWireKind.Fixed64 : IWorkWireKind.Varint;
                if (animation.FieldCount(field) > 1 || animation.HasUnexpectedWireKind(field, wire)) {
                    path = "4/2/8/" + field.ToString(System.Globalization.CultureInfo.InvariantCulture);
                    count = animation.FieldCount(field);
                    kind = IWorkSourceDeclarationIssueKind.MalformedMessage;
                    return false;
                }
            }
            if (animation.TotalFieldCount != knownCount) return false;
            if (animation.GetUnsigned(6) is ulong automatic && automatic > 1) {
                path = "4/2/8/6";
                count = 1;
                kind = IWorkSourceDeclarationIssueKind.MalformedMessage;
                return false;
            }
            string? type = animation.GetString(1, out bool typeComplete);
            string? effect = animation.GetString(2, out bool effectComplete);
            if (!typeComplete || !effectComplete) {
                kind = IWorkSourceDeclarationIssueKind.MalformedMessage;
                return false;
            }
            // Duration, delay and seed are dormant when the effect is none and
            // automatic advance is off. Native fixtures retain a delay of 0.5.
            return string.Equals(type, "Transition", StringComparison.Ordinal)
                && string.Equals(effect, "none", StringComparison.Ordinal)
                && (!animation.HasField(6) || animation.GetUnsigned(6) == 0);
        }

        IWorkWireMessage? Child(IWorkWireMessage parent, int field, string fieldPath) {
            path = fieldPath;
            count = parent.FieldCount(field);
            if (count != 1 || parent.HasUnexpectedWireKind(field, IWorkWireKind.Bytes)) {
                kind = IWorkSourceDeclarationIssueKind.MalformedMessage;
                return null;
            }
            return parent.ParseNestedMessage(parent.GetBytes(field)!);
        }
    }
}
