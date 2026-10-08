using System.Text;

namespace OfficeIMO.Access;

internal sealed partial class AccessNativeWriter {
    private static byte[]? Properties(IReadOnlyDictionary<string, object?> properties) {
        if (properties.Count == 0) return null;
        string[] names = properties.Keys.ToArray();
        return PropertyMaps(names, new[] { (Name: "", Type: (ushort)0, Values: properties.Select(p => (Index: Array.IndexOf(names, p.Key), Type: (byte)10, Payload: AccessNativeBinary.Unicode.GetBytes((string)p.Value!))).ToArray()) });
    }
    private static byte[]? TextColumnProperties(Table table) {
        Column[] text = table.Columns.Where(c => c.Type == 10 || c.Type == 12).ToArray();
        if (text.Length == 0) return null;
        return PropertyMaps(new[] { "AllowZeroLength" }, text.Select(c => (Name: c.Name, Type: (ushort)1, Values: new[] { (Index: 0, Type: (byte)1, Payload: new byte[] { 1 }) })).ToArray());
    }
    private static byte[] PropertyMaps(string[] names, (string Name, ushort Type, (int Index, byte Type, byte[] Payload)[] Values)[] maps) {
        using var result = new MemoryStream(); using var writer = new BinaryWriter(result, AccessNativeBinary.Unicode, true);
        writer.Write(new byte[] { (byte)'M', (byte)'R', (byte)'2', 0 });
        using (var block = new MemoryStream()) {
            using var nameWriter = new BinaryWriter(block, AccessNativeBinary.Unicode, true);
            foreach (string name in names) { byte[] bytes = AccessNativeBinary.Unicode.GetBytes(name); nameWriter.Write((ushort)bytes.Length); nameWriter.Write(bytes); }
            WritePropertyChunk(writer, 128, block.ToArray());
        }
        foreach (var map in maps) {
            using var block = new MemoryStream(); using var valueWriter = new BinaryWriter(block, AccessNativeBinary.Unicode, true);
            byte[] name = AccessNativeBinary.Unicode.GetBytes(map.Name); valueWriter.Write(6 + name.Length); valueWriter.Write((ushort)name.Length); valueWriter.Write(name);
            foreach (var value in map.Values) {
                valueWriter.Write(checked((ushort)(8 + value.Payload.Length))); valueWriter.Write((byte)0); valueWriter.Write(value.Type);
                valueWriter.Write((ushort)value.Index); valueWriter.Write((ushort)value.Payload.Length); valueWriter.Write(value.Payload);
            }
            WritePropertyChunk(writer, map.Type, block.ToArray());
        }
        return result.ToArray();
    }
    private static void WritePropertyChunk(BinaryWriter writer, ushort type, byte[] block) { writer.Write(block.Length + 6); writer.Write(type); writer.Write(block); }
}
