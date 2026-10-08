using System.Collections.ObjectModel;
using OfficeIMO.Drawing;
using static OfficeIMO.Access.AccessNativeBinary;

namespace OfficeIMO.Access;

internal sealed partial class AccessNativeDatabase {
    private void LoadProperties(AccessTable table, byte[]? bytes, CancellationToken cancellation) {
        if (bytes == null || bytes.Length == 0) return;
        if (bytes.Length > MaxMetadataBytes) throw new InvalidDataException("Native Access properties exceed MaxMetadataBytes.");
        table.NativeProperties = new AccessOpaqueValue(0, bytes, "Persisted table/column property map; retained without expression evaluation.");
        var data = new OfficeByteView(bytes);
        if (data.Length < 4 || data[0] != 'M' || data[1] != 'R' || data[2] != '2' || data[3] != 0) {
            table.Diagnostics = Array.AsReadOnly(new[] { new AccessDiagnostic("access.properties.opaque", "Unqualified property-map signature is retained exactly.", table.Id) }); return;
        }
        var names = new List<string>(); int position = 4;
        while (position < data.Length) {
            cancellation.ThrowIfCancellationRequested(); int length = I32(data, position); int type = U16(data, position + 4);
            if (length < 6) throw new InvalidDataException("Native Access property chunk makes no progress.");
            var block = Slice(data, position + 6, length - 6); position = checked(position + length);
            if (type == 128) {
                names.Clear(); int item = 0;
                while (item < block.Length) { if (names.Count == MaxCatalogObjects) throw new InvalidDataException("Native Access property names exceed their limit."); names.Add(PropertyName(block, ref item)); }
                continue;
            }
            if (type != 0 && type != 1 && type != 2) {
                table.Diagnostics = table.Diagnostics.Concat(new[] { new AccessDiagnostic("access.properties.unknown-chunk", "Unknown property chunk is retained in NativeProperties.", table.Id) }).ToArray(); continue;
            }
            if (block.Length == 0) continue;
            int nameLength = I32(block, 0);
            if (nameLength < 4 || nameLength > block.Length) throw new InvalidDataException("Native Access property map name block is invalid.");
            int offset = 4; string mapName = nameLength > 6 ? PropertyName(Slice(block, 0, nameLength), ref offset) : string.Empty;
            offset = nameLength; var values = new Dictionary<string, object?>(StringComparer.OrdinalIgnoreCase);
            while (offset < block.Length) {
                cancellation.ThrowIfCancellationRequested(); int valueLength = U16(block, offset);
                if (valueLength < 8) throw new InvalidDataException("Native Access property record makes no progress.");
                var value = Slice(block, offset, valueLength); offset = checked(offset + valueLength);
                byte nativeType = value[3]; int nameIndex = U16(value, 4), size = U16(value, 6);
                if (nameIndex >= names.Count || size > value.Length - 8) throw new InvalidDataException("Native Access property name/value reference is invalid.");
                var payload = Slice(value, 8, size); string name = names[nameIndex];
                object? decoded = nativeType == 10 || nativeType == 12 ? Text(payload) : nativeType == 9 || nativeType == 11 ? payload.ToArray() : DecodeScalar(new AccessNativeColumn { Type = nativeType }, payload, cancellation);
                if (values.ContainsKey(name)) throw new InvalidDataException("Native Access property names are ambiguous.");
                values.Add(name, decoded);
            }
            var properties = new ReadOnlyDictionary<string, object?>(values);
            if (type == 0 && mapName.Length == 0) table.Properties = properties;
            else if (type == 1) {
                var column = table.Columns.Items.SingleOrDefault(x => StringComparer.OrdinalIgnoreCase.Equals(x.Name, mapName));
                if (column != null) { column.Properties = properties; column.IsRichText = values.TryGetValue("TextFormat", out var format) && Convert.ToInt32(format) == 1; }
            }
        }
    }
    private string PropertyName(OfficeByteView block, ref int position) {
        int length = U16(block, position); position = checked(position + 2);
        if (length == 0 || length > MaxMetadataBytes || (length & 1) != 0) throw new InvalidDataException("Native Access property name length is invalid.");
        string name = Text(Slice(block, position, length)); position = checked(position + length);
        if (name.IndexOf('\0') >= 0) throw new InvalidDataException("Native Access property name contains a null character.");
        return name;
    }
}
