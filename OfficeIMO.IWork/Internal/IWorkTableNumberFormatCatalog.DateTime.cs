namespace OfficeIMO.IWork.Internal;

internal sealed partial class IWorkTableNumberFormatCatalog {
    private static IWorkNumberFormat? ReadDateTimeFormat(IWorkWireMessage message) {
        if (message.FieldCount(1) != 1 || message.HasUnexpectedWireKind(1, IWorkWireKind.Varint)
            || message.GetUnsigned(1) != 261 || message.FieldCount(14) != 1
            || message.HasUnexpectedWireKind(14, IWorkWireKind.Bytes)
            || message.TotalFieldCount != 2 + message.FieldCount(12) + message.FieldCount(13)) return null;
        // Suppression is not qualified. Explicit false is equivalent to the
        // absent defaults; repeated or non-Boolean flags cannot be ignored.
        foreach (int field in new[] { 12, 13 }) {
            if (message.FieldCount(field) > 1 || message.HasUnexpectedWireKind(field, IWorkWireKind.Varint)
                || (message.GetUnsigned(field) ?? 0) != 0) return null;
        }
        // Qualified patterns are ASCII and at most thirteen bytes. Reject
        // oversized unqualified text before decoding or retaining a string.
        byte[]? bytes = message.GetBytes(14);
        if (bytes is not { Length: > 0 and <= 13 }) return null;
        string? pattern = message.GetString(14);
        IWorkDateTimeFormat? format = pattern == null ? null : IWorkDateTimeFormat.CreateQualified(pattern);
        return format == null ? null : new IWorkNumberFormat(IWorkNumberFormatKind.DateTime,
            null, false, IWorkNegativeNumberStyle.Minus, dateTimeFormat: format);
    }
}
