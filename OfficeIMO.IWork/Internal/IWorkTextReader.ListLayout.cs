namespace OfficeIMO.IWork.Internal;

internal static partial class IWorkTextReader {
    private static void OverlayListLayout(IWorkWireMessage message, ListStyleData data,
        IWorkProjectionBudget budget, StylePropertyEvidence evidence, ref bool complete) {
        IReadOnlyList<float> textIndents = message.GetRepeatedFloat(12);
        if (message.HasUnexpectedWireKind(12, IWorkWireKind.Fixed32)
            || textIndents.Any(value => !IsFinite(value) || value < 0)) {
            evidence.Record(message, 12); complete = false;
        } else if (textIndents.Count > 0) {
            budget.AddTextItems(textIndents.Count);
            data.TextIndents = textIndents;
        }
        if (!message.HasField(14)) return;
        IReadOnlyList<IWorkWireMessage> geometries = IWorkObjectIndex.TryGetMessages(message, 14, out bool malformed);
        budget.AddTextItems(geometries.Count);
        var scales = new List<float>(geometries.Count);
        bool valid = !malformed;
        foreach (IWorkWireMessage geometry in geometries) {
            float? scale = geometry.GetFloat(1);
            float? offset = geometry.GetFloat(2);
            ulong? relative = geometry.GetUnsigned(3);
            if (geometry.FieldCount(1) != 1 || geometry.HasUnexpectedWireKind(1, IWorkWireKind.Fixed32)
                || !scale.HasValue || !IsFinite(scale.Value) || scale <= 0
                || geometry.FieldCount(2) != 1 || geometry.HasUnexpectedWireKind(2, IWorkWireKind.Fixed32)
                || offset != 0
                || geometry.FieldCount(3) != 1 || geometry.HasUnexpectedWireKind(3, IWorkWireKind.Varint)
                || relative != 1) valid = false;
            else scales.Add(scale.Value);
        }
        if (!valid) { evidence.Record(message, 14); complete = false; }
        else data.MarkerScales = scales;
    }

    private static IWorkListLayout? ResolveListLayout(ListStyleData data, int level, ref bool complete) {
        if (data.TextIndents.Count == 0 && data.MarkerScales.Count == 0) return null;
        if (level < 0 || level >= data.LeftIndents.Count || level >= data.TextIndents.Count
            || level >= data.MarkerScales.Count) {
            complete = false; return null;
        }
        return new IWorkListLayout(data.LeftIndents[level], data.TextIndents[level], data.MarkerScales[level]);
    }
}
