using OfficeIMO.Drawing;
using System.Collections.ObjectModel;
using static OfficeIMO.Access.AccessNativeBinary;

namespace OfficeIMO.Access {
    internal sealed partial class AccessNativeDatabase {
        private void LoadProperties(AccessTable table, byte[]? bytes, CancellationToken cancellation, bool macroMap = false) {
            if (bytes == null || bytes.Length == 0) return;
            if (bytes.Length > MaxMetadataBytes) throw new InvalidDataException("Native Access properties exceed MaxMetadataBytes.");
            table.NativeProperties = new AccessOpaqueValue(0, bytes, "Persisted table/column property map; retained without expression evaluation.");
            OfficeByteView data = new OfficeByteView(bytes);
            if (data.Length < 4 || data[0] != 'M' || data[1] != 'R' || data[2] != '2' || data[3] != 0) {
                table.Diagnostics = Array.AsReadOnly(new[] { new AccessDiagnostic("access.properties.opaque", "Unqualified property-map signature is retained exactly.", table.Id) }); return;
            }
            List<string> names = new List<string>(); int position = 4;
            while (position < data.Length) {
                cancellation.ThrowIfCancellationRequested(); int length = I32(data, position); int type = U16(data, position + 4);
                if (length < 6) throw new InvalidDataException("Native Access property chunk makes no progress.");
                OfficeByteView block = Slice(data, position + 6, length - 6); position = checked(position + length);
                if (type == 128) {
                    names.Clear(); int item = 0;
                    while (item < block.Length) { if (names.Count == MaxCatalogObjects) throw new InvalidDataException("Native Access property names exceed their limit."); names.Add(PropertyName(block, ref item)); }
                    continue;
                }
                if (type != 0 && type != 1 && type != 2 && !(macroMap && type == 3)) {
                    table.Diagnostics = table.Diagnostics.Concat(new[] { new AccessDiagnostic("access.properties.unknown-chunk", "Unknown property chunk is retained in NativeProperties.", table.Id) }).ToArray(); continue;
                }
                if (block.Length == 0) continue;
                int nameLength = I32(block, 0);
                if (nameLength < 4 || nameLength > block.Length) throw new InvalidDataException("Native Access property map name block is invalid.");
                int offset = macroMap && type == 3 ? 8 : 4;
                if (macroMap && type == 3 && (nameLength < 10 || I32(block, 4) != 0)) throw new InvalidDataException("Native Access data-macro map header is unqualified.");
                string mapName = nameLength > offset + 2 ? PropertyName(Slice(block, 0, nameLength), ref offset) : string.Empty;
                offset = nameLength; Dictionary<string, object?> values = new Dictionary<string, object?>(StringComparer.OrdinalIgnoreCase);
                while (offset < block.Length) {
                    cancellation.ThrowIfCancellationRequested(); bool wide = macroMap && type == 3;
                    int headerLength = wide ? 12 : 8;
                    int valueLength = wide ? I32(block, offset) : U16(block, offset);
                    if (valueLength < headerLength) throw new InvalidDataException("Native Access property record makes no progress.");
                    OfficeByteView value = Slice(block, offset, valueLength); offset = checked(offset + valueLength);
                    byte nativeType = value[wide ? 5 : 3]; int nameIndex = U16(value, wide ? 6 : 4), size = wide ? I32(value, 8) : U16(value, 6);
                    if (nameIndex >= names.Count || size < 0 || size > value.Length - headerLength) throw new InvalidDataException("Native Access property name/value reference is invalid.");
                    OfficeByteView payload = Slice(value, headerLength, size); string name = names[nameIndex];
                    int scalarWidth = nativeType switch { 1 => 1, 2 => 1, 3 => 2, 4 => 4, 5 => 8, 6 => 4, 7 => 8, 8 => 8, 15 => 16, 19 => 8, 16 => 17, 18 => 4, 20 => 42, _ => -1 };
                    object? decoded;
                    if (scalarWidth >= 0 && payload.Length != scalarWidth) {
                        decoded = new AccessOpaqueValue(nativeType, payload.ToArray(), "This property uses an unqualified width for its nominal native type; the exact payload is retained.");
                        table.Diagnostics = table.Diagnostics.Concat(new[] { new AccessDiagnostic("access.properties.opaque-value", "An unqualified property value representation is retained without coercion.", table.Id) }).ToArray();
                    } else decoded = nativeType == 10 || nativeType == 12 ? Text(payload) : nativeType == 9 || nativeType == 11 ? payload.ToArray() : DecodeScalar(new AccessNativeColumn { Type = nativeType }, payload, cancellation);
                    if (values.ContainsKey(name)) throw new InvalidDataException("Native Access property names are ambiguous.");
                    values.Add(name, decoded);
                }
                ReadOnlyDictionary<string, object?> properties = new ReadOnlyDictionary<string, object?>(values);
                if ((type == 0 || macroMap && type == 3) && mapName.Length == 0) table.Properties = properties;
                else if (type == 1) {
                    AccessColumn? column = table.Columns.Items.SingleOrDefault(x => StringComparer.OrdinalIgnoreCase.Equals(x.Name, mapName));
                    if (column != null) { column.Properties = properties; column.IsRichText = values.TryGetValue("TextFormat", out object? format) && Convert.ToInt32(format) == 1; }
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
}
