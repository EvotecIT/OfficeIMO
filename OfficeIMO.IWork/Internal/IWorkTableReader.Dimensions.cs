namespace OfficeIMO.IWork.Internal;

internal static partial class IWorkTableReader {
    private static bool HasUnsupportedTileMetadata(IWorkWireMessage message, int metadataFieldCount) {
        int recognized = 0;
        foreach (int field in new[] { 1, 2, 3, 4, 6, 7, 8 }) {
            int count = message.FieldCount(field);
            if (count > 1 || message.HasUnexpectedWireKind(field, IWorkWireKind.Varint)) return true;
            recognized += count;
        }
        return recognized != metadataFieldCount || message.GetUnsigned(7) > 1 || message.GetUnsigned(8) > 1;
    }

    private static bool HasUnsupportedTableScalarEncoding(IWorkWireMessage message) =>
        new[] { 6, 7, 9, 10, 11, 16, 17 }.Any(field => message.FieldCount(field) > 1)
        || message.HasUnexpectedWireKind(6, IWorkWireKind.Varint)
        || message.HasUnexpectedWireKind(7, IWorkWireKind.Varint)
        || message.HasUnexpectedWireKind(9, IWorkWireKind.Varint)
        || message.HasUnexpectedWireKind(10, IWorkWireKind.Varint)
        || message.HasUnexpectedWireKind(11, IWorkWireKind.Varint)
        || message.HasUnexpectedWireKind(16, IWorkWireKind.Fixed64)
        || message.HasUnexpectedWireKind(17, IWorkWireKind.Fixed64);

    private static int CheckedSubDimension(ulong? value, int maximum, string label, IWorkArchiveRecord record) {
        ulong resolved = value ?? 0;
        if (resolved > (ulong)maximum || resolved > int.MaxValue) {
            throw new InvalidDataException($"iWork table {label} count {resolved} in object {record.Identifier} exceeds the table dimensions.");
        }
        return (int)resolved;
    }

    private static double? ValidDimension(double? value) => value.HasValue && IsFinite(value.Value) && value.Value > 0
        ? value
        : null;

    private static bool HasInvalidDeclaredDimension(IWorkWireMessage message, int field,
        double? value) => message.FieldCount(field) > 1
        || message.HasField(field)
            && (message.HasUnexpectedWireKind(field, IWorkWireKind.Fixed64)
                || !value.HasValue || !IsFinite(value.Value) || value.Value <= 0);

    private static int CheckedDimension(ulong? value, int maximum, string label, IWorkArchiveRecord record) {
        ulong resolved = value ?? 0;
        if (resolved > (ulong)maximum || resolved > int.MaxValue) {
            throw new InvalidDataException($"iWork table {label} count {resolved} in object {record.Identifier} exceeds the configured limit of {maximum}.");
        }
        return (int)resolved;
    }

}
