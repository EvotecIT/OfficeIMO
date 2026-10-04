namespace OfficeIMO.IWork.Internal;

internal static partial class IWorkTextReader {
    /// <summary>Prefers the selected modern text fill over its legacy font-color mirror.</summary>
    private static void OverlayTextColor(IWorkWireMessage message, TextStyleData data,
        IWorkArchiveRecord record, IWorkSourceReferenceIssueCollector references,
        StylePropertyEvidence evidence, ref bool complete) {
        bool? fillContainer = ReadBoolean(message, 47, evidence, ref complete);
        if (fillContainer == true) {
            evidence.Record(message, 47, IWorkSourceDeclarationIssueKind.UnsupportedField);
            complete = false;
        }
        bool? clearFill = ReadBoolean(message, 45, evidence, ref complete);
        if (clearFill == true) {
            // Clearing a modern fill is distinct from supplying an empty (no-fill) FillArchive.
            // Neither declaration has a qualified visible-text representation in the current model.
            evidence.Record(message, message.HasField(46) ? 46 : 45, message.HasField(46)
                ? IWorkSourceDeclarationIssueKind.InvalidValue : IWorkSourceDeclarationIssueKind.UnsupportedField);
            complete = false;
            return;
        }
        if (message.HasField(45) && !clearFill.HasValue) return;
        if (message.HasField(46)) {
            IWorkCellFill? fill = null;
            IWorkCellFillReader.Read(record, message, references, ref fill, ref complete, field: 46);
            if (fill?.Color != null) data.Color = fill.Color;
            else if (fill != null) {
                evidence.Record(message, 46, IWorkSourceDeclarationIssueKind.UnsupportedField);
                complete = false;
            }
            return;
        }
        bool? clearColor = ReadBoolean(message, 6, evidence, ref complete);
        if (clearColor == true) {
            if (message.HasField(7)) { evidence.Record(message, 7); complete = false; }
            data.Color = null;
        }
        else if ((!message.HasField(6) || clearColor.HasValue)
            && TryColor(message, 7, out IWorkColor? color, evidence, ref complete)) data.Color = color;
    }
}
