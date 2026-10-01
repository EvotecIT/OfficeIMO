using System.Globalization;

namespace OfficeIMO.IWork.Internal;

internal static partial class IWorkFormulaReader {
    // Native merge formulas can explicitly identify their owning table. The table
    // model's textual UUID uses the reverse byte order of the CFUUID word encoding.
    // Only the independently qualified four-word representation is accepted here.
    private static bool ReferencesOwningTable(IWorkWireMessage node, IWorkWireMessage? table) {
        if (table == null || table.FieldCount(1) != 1
            || table.HasUnexpectedWireKind(1, IWorkWireKind.Bytes)
            || node.FieldCount(28) != 1 || node.HasUnexpectedWireKind(28, IWorkWireKind.Bytes)) return false;
        string? tableId = table.GetString(1, out bool complete);
        if (!complete || !Guid.TryParseExact(tableId, "D", out Guid identifier)) return false;
        IWorkWireMessage? reference = IWorkObjectIndex.TryGetMessage(node, 28, out bool malformed);
        if (malformed || reference == null || reference.TotalFieldCount != 1
            || reference.FieldCount(1) != 1 || reference.HasUnexpectedWireKind(1, IWorkWireKind.Bytes)) return false;
        IWorkWireMessage? uuid = IWorkObjectIndex.TryGetMessage(reference, 1, out malformed);
        if (malformed || uuid == null || uuid.TotalFieldCount != 4) return false;
        var words = new uint[4];
        for (int index = 0; index < words.Length; index++) {
            int field = index + 2;
            if (uuid.FieldCount(field) != 1 || uuid.HasUnexpectedWireKind(field, IWorkWireKind.Varint)
                || uuid.GetUnsigned(field) is not ulong value || value > uint.MaxValue) return false;
            words[index] = (uint)value;
        }
        string encoded = words[3].ToString("x8", CultureInfo.InvariantCulture)
            + words[2].ToString("x8", CultureInfo.InvariantCulture)
            + words[1].ToString("x8", CultureInfo.InvariantCulture)
            + words[0].ToString("x8", CultureInfo.InvariantCulture);
        var reversed = new char[32];
        for (int index = 0; index < 16; index++) {
            reversed[index * 2] = encoded[(15 - index) * 2];
            reversed[index * 2 + 1] = encoded[(15 - index) * 2 + 1];
        }
        return string.Equals(identifier.ToString("N"), new string(reversed), StringComparison.Ordinal);
    }
}
