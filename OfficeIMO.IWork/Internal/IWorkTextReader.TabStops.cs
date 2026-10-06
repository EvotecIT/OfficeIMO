namespace OfficeIMO.IWork.Internal;

internal static partial class IWorkTextReader {
    private static bool TryTabStops(IWorkWireMessage message, IWorkProjectionBudget budget,
        out IReadOnlyList<IWorkTabStop>? tabs) {
        tabs = null;
        int count = message.FieldCount(1);
        if (count != message.TotalFieldCount
            || message.HasUnexpectedWireKind(1, IWorkWireKind.Bytes)) return false;
        if (count == 0) {
            // The native empty custom-tab declaration clears inherited custom stops.
            tabs = Array.Empty<IWorkTabStop>();
            return true;
        }
        budget.AddTextItems(count);
        var result = new List<IWorkTabStop>(count);
        var positions = new HashSet<double>();
        foreach (byte[] bytes in message.EnumerateRepeatedBytes(1)) {
            IWorkWireMessage value;
            try { value = message.ParseNestedMessage(bytes); }
            catch (InvalidDataException exception) when (!IWorkProtobuf.IsLimitException(exception)) { return false; }
            if (value.FieldCount(1) != 1 || value.HasUnexpectedWireKind(1, IWorkWireKind.Fixed32)
                || value.FieldCount(2) > 1 || value.HasUnexpectedWireKind(2, IWorkWireKind.Varint)
                || value.GetUnsigned(2).GetValueOrDefault() > 3
                || value.FieldCount(3) > 1 || value.HasUnexpectedWireKind(3, IWorkWireKind.Bytes)
                || value.HasField(3) && value.GetBytes(3)!.Length != 0
                || value.TotalFieldCount != value.FieldCount(1) + value.FieldCount(2) + value.FieldCount(3)) return false;
            float? position = value.GetFloat(1);
            if (!position.HasValue || !IsFinite(position.Value) || position.Value < 0
                || !positions.Add(position.Value)) return false;
            result.Add(new IWorkTabStop(position.Value, (IWorkTabAlignment)value.GetUnsigned(2).GetValueOrDefault()));
        }
        tabs = Array.AsReadOnly(result.OrderBy(tab => tab.PositionPoints).ToArray());
        return true;
    }
}
