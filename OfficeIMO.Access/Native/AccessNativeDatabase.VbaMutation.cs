using OfficeIMO.Core.Internal;
using System.Text;

namespace OfficeIMO.Access {
    internal sealed partial class AccessNativeDatabase {
        /// <summary>Adapts a shared project artifact to native Access storage and its module catalog.</summary>
        internal AccessNativeWriter BuildVbaMutation(OfficeVbaProject project, OfficeCompoundFile compound,
            long maximumBytes, CancellationToken cancellation, ISet<string> hosts, IReadOnlyDictionary<OfficeVbaModule, string>? appliedNames = null) {
            const string prefix = "VBA/VBAProject/";
            var replacements = compound.Streams.ToDictionary(x => prefix + x.Key, x => x.Value, StringComparer.OrdinalIgnoreCase);
            var removals = new HashSet<string>(_applicationStreams.Keys.Where(x => x.StartsWith(prefix, StringComparison.OrdinalIgnoreCase) && !replacements.ContainsKey(x)), StringComparer.OrdinalIgnoreCase);
            AccessNativeTable catalog = _tables["MSysObjects"];
            var original = _catalog.Where(x => x.Type == -32761).ToDictionary(x => x.Name, StringComparer.OrdinalIgnoreCase);
            OfficeVbaModule[] ordinary = project.Modules.Where(x => !hosts.Contains(x.Name) && (x.Kind == OfficeVbaModuleKind.Standard || x.Kind == OfficeVbaModuleKind.Class)).ToArray();
            // Core retains the detached model's load identity. A reused model must instead
            // address the native names accepted by its immediately preceding application.
            var sourceNames = ordinary.ToDictionary(x => x, x => appliedNames != null && appliedNames.TryGetValue(x, out string? applied)
                ? applied : x.IsNew ? null : x.OriginalName);
            var renamed = ordinary.Where(x => sourceNames[x] != null && original.ContainsKey(sourceNames[x]!) && x.Name != sourceNames[x])
                .ToDictionary(x => sourceNames[x]!, x => x.Name, StringComparer.OrdinalIgnoreCase);
            string[] deletedNames = original.Keys.Where(name => !renamed.ContainsKey(name) && !ordinary.Any(x => x.Name.Equals(name, StringComparison.OrdinalIgnoreCase))).ToArray();
            OfficeVbaModule[] added = ordinary.Where(x => (!original.ContainsKey(x.Name) || renamed.ContainsKey(x.Name))
                && (sourceNames[x] == null || !renamed.ContainsKey(sourceNames[x]!))).ToArray();
            var deletedIds = new HashSet<int>(deletedNames.Select(name => original[name].Id));
            var additions = new List<object?[]>(); var catalogRemovals = new HashSet<uint>(); var catalogRenames = new Dictionary<uint, string>();
            int namespaceId = 0; byte[]? owner = null;
            if (added.Length != 0 || deletedNames.Length != 0 || renamed.Count != 0) {
                NativeCatalogRecord moduleNamespace = _catalog.SingleOrDefault(x => x.Type == 3 && x.Name == "Modules")
                    ?? throw new NotSupportedException("The native module namespace is not qualified.");
                namespaceId = moduleNamespace.Id;
                if (original.Values.Any(x => x.ParentId != namespaceId || x.Flags != 0))
                    throw new NotSupportedException("The native module catalog has unqualified namespace or flag metadata.");
                bool firstProject = !_applicationStreams.Keys.Any(x => x.StartsWith(prefix, StringComparison.OrdinalIgnoreCase));
                if (!_applicationStreams.TryGetValue("Modules/\u0003DirData", out AccessStorageStream? directory) && (!firstProject || original.Count != 0))
                    throw new NotSupportedException("The native module directory is unavailable.");
                Dictionary<string, int> slots = directory == null ? new Dictionary<string, int>(StringComparer.OrdinalIgnoreCase) : ReadObjectDirectory(directory.Payload.GetBytes())
                    ?? throw new NotSupportedException("The native module directory is outside the qualified layout.");
                if (slots.Count != original.Count || original.Keys.Any(x => !slots.ContainsKey(x)))
                    throw new InvalidDataException("The native module catalog and storage directory disagree.");
                int nextSlot = slots.Values.DefaultIfEmpty(-1).Max();
                var renamedSlots = renamed.ToDictionary(x => x.Value, x => slots[x.Key], StringComparer.OrdinalIgnoreCase);
                foreach (string name in deletedNames) {
                    string storage = "Modules/" + slots[name].ToString(System.Globalization.CultureInfo.InvariantCulture) + "/";
                    foreach (string path in _applicationStreams.Keys.Where(x => x.StartsWith(storage, StringComparison.OrdinalIgnoreCase))) removals.Add(path);
                    slots.Remove(name);
                }
                foreach (string name in renamed.Keys) slots.Remove(name);
                foreach (var pair in renamedSlots) {
                    if (slots.ContainsKey(pair.Key)) throw new ArgumentException("A renamed module collides with its native catalog identity.", nameof(project));
                    slots.Add(pair.Key, pair.Value);
                }
                int nextId = _catalog.Where(x => x.Id < 0).Select(x => x.Id).DefaultIfEmpty(int.MinValue).Max();
                if (nextId >= -1 || (long)nextId + added.Length >= 0) throw new InvalidDataException("The native application object identity space is exhausted.");
                using (var rows = new AccessNativeRowCursor(catalog, cancellation, rowLimit: MaxCatalogObjects)) while (rows.Read(cancellation)) {
                    int id = Convert.ToInt32(RequiredField(catalog, rows, "Id", cancellation));
                    if (id == namespaceId) owner = Field(catalog, rows, "Owner", cancellation) as byte[];
                    if (deletedIds.Contains(id)) catalogRemovals.Add(rows.CurrentPointer);
                    NativeCatalogRecord? module = original.Values.FirstOrDefault(x => x.Id == id);
                    if (module != null && renamed.TryGetValue(module.Name, out string? name)) catalogRenames.Add(rows.CurrentPointer, name);
                }
                if (added.Length != 0 && owner == null) throw new NotSupportedException("The native module namespace owner is unavailable.");
                foreach (OfficeVbaModule module in added) {
                    int slot = checked(++nextSlot); slots.Add(module.Name, slot);
                    using (MemoryStream properties = new MemoryStream()) {
                        using (BinaryWriter value = new BinaryWriter(properties, Encoding.Unicode, true)) { value.Write(0); value.Write(2); value.Write(module.Kind == OfficeVbaModuleKind.Class ? 0x10000 : 0); value.Write((byte)0); }
                        replacements["Modules/" + slot.ToString(System.Globalization.CultureInfo.InvariantCulture) + "/PropData"] = properties.ToArray();
                    }
                    object?[] row = new object?[catalog.Columns.Count];
                    SetNativeField(catalog, row, "Id", checked(++nextId)); SetNativeField(catalog, row, "ParentId", namespaceId);
                    SetNativeField(catalog, row, "Name", module.Name); SetNativeField(catalog, row, "Type", (short)-32761);
                    SetNativeField(catalog, row, "DateCreate", _document.CreatedAt.ToOADate()); SetNativeField(catalog, row, "DateUpdate", _document.CreatedAt.ToOADate());
                    SetNativeField(catalog, row, "Owner", owner); SetNativeField(catalog, row, "Flags", 0); additions.Add(row);
                }
                replacements["Modules/\u0003DirData"] = WriteModuleDirectory(slots);
            }
            if (_applicationStreams.TryGetValue("VBA/AcessVBAData", out AccessStorageStream? accessData)) {
                byte[] metadata = accessData.Payload.GetBytes();
                if (metadata.Length != 12 || AccessNativeBinary.I32(metadata, 0) != 1 || AccessNativeBinary.I32(metadata, 4) != 1)
                    throw new NotSupportedException("The Access VBA host metadata is outside its qualified layout.");
                for (int i = 0; i < 4; i++) metadata[8 + i] = (byte)(project.Modules.Count >> (i * 8));
                replacements[accessData.Path] = metadata;
            } else {
                if (_applicationStreams.Keys.Any(x => x.StartsWith(prefix, StringComparison.OrdinalIgnoreCase)))
                    throw new NotSupportedException("The Access VBA host metadata is unavailable.");
                byte[] metadata = new byte[12]; metadata[0] = 1; metadata[4] = 1;
                for (int i = 0; i < 4; i++) metadata[8 + i] = (byte)(project.Modules.Count >> (i * 8));
                replacements["VBA/AcessVBAData"] = metadata;
                if (!_applicationStreams.TryGetValue("PropData", out AccessStorageStream? propertyStream)
                    || !propertyStream.Payload.GetBytes().SequenceEqual(new byte[] { 0,0,0,0,2,0x69,0,0,0,4,0,0,0 }))
                    throw new NotSupportedException("First-project application properties require the qualified empty native layout.");
                using MemoryStream properties = new MemoryStream(); properties.Write(propertyStream.Payload.GetBytes(), 0, 13);
                using (BinaryWriter value = new BinaryWriter(properties, Encoding.Unicode, true)) { value.Write((byte)2); value.Write(0x6a); value.Write(project.CodePage); }
                replacements["PropData"] = properties.ToArray();
            }
            AccessNativeWriter writer = BuildApplicationStreamReplacement(replacements, maximumBytes, cancellation, removals);
            writer.MutateApplicationRows(this, catalog, additions, catalogRemovals, catalogRenames);
            if (additions.Count != 0 || deletedIds.Count != 0) MutateModulePermissions(writer, namespaceId, original.Values.Select(x => x.Id).ToArray(), additions, deletedIds, cancellation);
            return writer;
        }

        private static void SetNativeField(AccessNativeTable table, object?[] row, string name, object? value) {
            int ordinal = table.Columns.FindIndex(x => x.Name.Equals(name, StringComparison.OrdinalIgnoreCase));
            if (ordinal < 0) throw new NotSupportedException("The native application catalog schema is unavailable.");
            row[ordinal] = value;
        }

        internal static byte[] WriteModuleDirectory(IReadOnlyDictionary<string, int> slots) {
            using MemoryStream output = new MemoryStream(); using BinaryWriter writer = new BinaryWriter(output, Encoding.Unicode, true);
            writer.Write(0);
            foreach (var pair in slots.OrderBy(x => x.Key, StringComparer.OrdinalIgnoreCase)) {
                byte[] name = AccessNativeBinary.Unicode.GetBytes(pair.Key); writer.Write((byte)4); writer.Write(checked((byte)(name.Length + 4))); writer.Write(name); writer.Write(pair.Value);
            }
            writer.Flush(); return output.ToArray();
        }
    }
}
