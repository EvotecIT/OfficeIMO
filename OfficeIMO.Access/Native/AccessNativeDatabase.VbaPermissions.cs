namespace OfficeIMO.Access {
    internal sealed partial class AccessNativeDatabase {
        private void MutateModulePermissions(AccessNativeWriter writer, int parentId, int[] originalIds,
            IReadOnlyList<object?[]> moduleAdditions, ISet<int> deletedIds, CancellationToken cancellation) {
            if (!_tables.TryGetValue("MSysACEs", out AccessNativeTable? permissions)) throw new NotSupportedException("The native module permission catalog is unavailable.");
            var parent = new HashSet<string>(StringComparer.Ordinal); var inherited = new List<object?[]>(); var profiles = originalIds.ToDictionary(x => x, _ => new List<object?[]>());
            var removals = new HashSet<uint>();
            int idOrdinal = permissions.Columns.FindIndex(x => x.Name == "ObjectId");
            int sidOrdinal = permissions.Columns.FindIndex(x => x.Name == "SID"), rightsOrdinal = permissions.Columns.FindIndex(x => x.Name == "ACM");
            int inheritOrdinal = permissions.Columns.FindIndex(x => x.Name == "FInheritable");
            if (idOrdinal < 0 || sidOrdinal < 0 || rightsOrdinal < 0 || inheritOrdinal < 0 || permissions.Columns.Count != 4)
                throw new NotSupportedException("The native module permission schema is outside the qualified layout.");
            string Identity(object?[] row) => Convert.ToBase64String((byte[])row[sidOrdinal]!) + ":" + Convert.ToInt32(row[rightsOrdinal]).ToString(System.Globalization.CultureInfo.InvariantCulture);
            using (var cursor = new AccessNativeRowCursor(permissions, cancellation, rowLimit: checked((long)MaxCatalogObjects * 16))) while (cursor.Read(cancellation)) {
                int id = Convert.ToInt32(cursor.GetValue(idOrdinal, cancellation));
                if (deletedIds.Contains(id)) removals.Add(cursor.CurrentPointer);
                if (id != parentId && !profiles.ContainsKey(id)) continue;
                object?[] row = Enumerable.Range(0, 4).Select(i => cursor.GetValue(i, cancellation)).ToArray();
                if (row[sidOrdinal] is not byte[] || row[rightsOrdinal] is not int || row[inheritOrdinal] is not bool)
                    throw new InvalidDataException("A native module permission record is invalid.");
                if (id == parentId && (bool)row[inheritOrdinal]!) { parent.Add(Identity(row)); inherited.Add(row); }
                else if (profiles.TryGetValue(id, out List<object?[]>? profile)) profile.Add(row);
            }
            var additions = new List<object?[]>();
            if (moduleAdditions.Count != 0) {
                List<object?[]>? template = profiles.Values.FirstOrDefault();
                if (profiles.Count == 0) {
                    byte[]? owner = _document.Catalog.Single(x => x.NativeId == parentId).Owner?.GetBytes();
                    if (owner == null || owner.Length != 2 || inherited.Count == 0 || inherited.Any(x => Convert.ToInt32(x[rightsOrdinal]) != 1048575))
                        throw new NotSupportedException("First modules require qualified full-control namespace inheritance and its built-in owner.");
                    // Native Access copies the inheritable principals except its built-in Users group (SID 0x0402).
                    // Owner SID 0x0103 and group SID 0x0402 share the same two-byte native obfuscation mask.
                    template = inherited.Where(x => ((byte[])x[sidOrdinal]!).Length != 2
                        || ((byte[])x[sidOrdinal]!)[0] != (byte)(owner[0] ^ 1) || ((byte[])x[sidOrdinal]!)[1] != (byte)(owner[1] ^ 5))
                        .Select(x => { object?[] row = (object?[])x.Clone(); row[inheritOrdinal] = false; return row; }).ToList();
                }
                if (template == null || template.Count == 0 || template.Any(x => (bool)x[inheritOrdinal]! || !parent.Contains(Identity(x))))
                    throw new NotSupportedException("New modules require an existing noninheritable permission profile contained in the Modules namespace.");
                string[] expected = template.Select(Identity).OrderBy(x => x, StringComparer.Ordinal).ToArray();
                if (profiles.Values.Any(profile => profile.Any(x => (bool)x[inheritOrdinal]!) || !profile.Select(Identity).OrderBy(x => x, StringComparer.Ordinal).SequenceEqual(expected)))
                    throw new NotSupportedException("New modules require uniform qualified existing module permissions.");
                int catalogId = _tables["MSysObjects"].Columns.FindIndex(x => x.Name == "Id");
                foreach (object?[] module in moduleAdditions) foreach (object?[] source in template) {
                    object?[] row = (object?[])source.Clone(); row[idOrdinal] = module[catalogId]; additions.Add(row);
                }
            }
            writer.MutateApplicationRows(this, permissions, additions, removals, rowLimit: checked((long)MaxCatalogObjects * 16));
        }
    }
}
