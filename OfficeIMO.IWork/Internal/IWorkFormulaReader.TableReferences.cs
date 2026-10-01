using System.Globalization;

namespace OfficeIMO.IWork.Internal;

internal static partial class IWorkFormulaReader {
    internal static Guid? ReadTableIdentifier(IWorkWireMessage? table) {
        if (table == null || table.FieldCount(1) != 1 || table.HasUnexpectedWireKind(1, IWorkWireKind.Bytes)) return null;
        string? text = table.GetString(1, out bool complete);
        return complete && Guid.TryParseExact(text, "D", out Guid identifier) ? identifier : null;
    }

    // Native textual table UUIDs use the reverse byte order of CFUUID words.
    // Accept only the independently qualified four-word representation.
    internal static Guid? ReadReferencedTableIdentifier(IWorkWireMessage node) {
        if (node.FieldCount(28) != 1 || node.HasUnexpectedWireKind(28, IWorkWireKind.Bytes)) return null;
        IWorkWireMessage? reference = IWorkObjectIndex.TryGetMessage(node, 28, out bool malformed);
        if (malformed || reference == null || reference.TotalFieldCount != 1
            || reference.FieldCount(1) != 1 || reference.HasUnexpectedWireKind(1, IWorkWireKind.Bytes)) return null;
        IWorkWireMessage? uuid = IWorkObjectIndex.TryGetMessage(reference, 1, out malformed);
        if (malformed || uuid == null || uuid.TotalFieldCount != 4) return null;
        var words = new uint[4];
        for (int index = 0; index < words.Length; index++) {
            int field = index + 2;
            if (uuid.FieldCount(field) != 1 || uuid.HasUnexpectedWireKind(field, IWorkWireKind.Varint)
                || uuid.GetUnsigned(field) is not ulong value || value > uint.MaxValue) return null;
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
        return Guid.ParseExact(new string(reversed), "N");
    }

    private static bool ReferencesOwningTable(IWorkWireMessage node, IWorkWireMessage? table) {
        Guid? own = ReadTableIdentifier(table);
        return own.HasValue && own == ReadReferencedTableIdentifier(node);
    }

    internal static string? QuoteTableName(string name, int maximumCharacters) {
        if (name.Length > maximumCharacters - 2) return null;
        long length = 2L + name.Length + name.Count(character => character == '\'');
        return length <= maximumCharacters ? "'" + name.Replace("'", "''") + "'" : null;
    }
}
