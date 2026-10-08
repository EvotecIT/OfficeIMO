namespace OfficeIMO.Access {
    internal sealed partial class AccessNativeWriter {
        private void Initialize(AccessDocument document) {
            if (document.Inspection != null || document.Queries.Count != 0) throw new NotSupportedException("Native creation accepts new table models; existing files and saved query authoring require separate codecs.");
            for (int i = 0; i < 6; i++) Allocate();
            _pages[0] = Header(document.Format == AccessFileFormat.Accdb, document.CreatedAt);
            List<Table> users = new List<Table>();
            foreach (AccessTable model in document.Tables) {
                _cancellation.ThrowIfCancellationRequested();
                ValidateName(model.Name);
                foreach (AccessColumn column in model.Columns) ValidateName(column.Name);
                foreach (AccessIndex index in model.Indexes) ValidateName(index.Name);
                if (model.Name.StartsWith("MSys", StringComparison.OrdinalIgnoreCase)) throw new NotSupportedException("Names beginning with MSys are reserved for native system tables.");
                TextKey(model.Name); // Catalog keys require a qualified collation even without user indexes.
                if (model.Columns.Count == 0) throw new NotSupportedException("A native table requires at least one column.");
                if (model.Columns.Count(x => x.DataType == AccessDataType.AutoNumber) > 1) throw new NotSupportedException("Only one sequential AutoNumber column per table is supported.");
                Column[] columns = model.Columns.Select(CreateColumn).ToArray();
                int minimumRow = 2 + (columns.Length + 7) / 8 + columns.Where(c => !c.Variable && c.Type != 1).Sum(c => c.Size);
                if (columns.Any(c => c.Variable)) minimumRow += 4 + columns.Count(c => c.Variable) * 2;
                if (minimumRow > 4060) throw new NotSupportedException("The fixed schema exceeds the qualified native row capacity of 4060 bytes.");
                if (model.Indexes.Count > 32) throw new NotSupportedException("Native Access tables support at most 32 indexes, including relationship indexes.");
                object?[][] rows = new object?[model.Rows.Count][]; int counter = model.Columns.FirstOrDefault(c => c.DataType == AccessDataType.AutoNumber)?.AutoNumberSeed - 1 ?? 0;
                for (int r = 0; r < rows.Length; r++) {
                    _cancellation.ThrowIfCancellationRequested(); rows[r] = new object?[columns.Length];
                    for (int c = 0; c < columns.Length; c++) {
                        bool specified = model.Rows[r].TryGetValue(columns[c].Name, out object? value);
                        if (columns[c].AutoNumber) {
                            if (specified && value == null) throw new NotSupportedException("An explicit null AutoNumber cannot be preserved by sequential native allocation.");
                            if (!specified) value = checked(counter + 1);
                            counter = Math.Max(counter, (int)value!);
                        }
                        if (columns[c].Type == 1) {
                            if (specified && value == null) throw new NotSupportedException("Native Yes/No stores two states; an explicit null cannot be persisted without loss.");
                            value ??= false;
                        }
                        rows[r][c] = value;
                    }
                }
                Table table = new Table(Allocate(), model.Name, columns, rows) { AutoNumberLast = counter, MapPage = Allocate() };
                foreach (AccessIndex index in model.Indexes) {
                    if (index.Descending.Any(x => x)) throw new NotSupportedException("Descending index authoring is not qualified.");
                    if (index.Columns.Any(c => c.DataType != AccessDataType.Byte && c.DataType != AccessDataType.Int16 && c.DataType != AccessDataType.Int32 && c.DataType != AccessDataType.AutoNumber && c.DataType != AccessDataType.ShortText))
                        throw new NotSupportedException("Native index creation currently qualifies Byte, Int16, Int32/AutoNumber and bounded text keys.");
                    Index native = new Index(index.Name, index.Columns.Select(c => model.Columns.Items.IndexOf(c)).ToArray(), (byte)(128 | (index.IsUnique ? 1 : 0) | (index.IsPrimaryKey ? 8 : 0)), (byte)(index.IsPrimaryKey ? 1 : 0)) { RootPage = Allocate() };
                    table.Indexes.Add(native);
                }
                users.Add(table);
            }
            List<object?[]> relations = new List<object?[]>();
            foreach (AccessRelationship relationship in document.Relationships) {
                ValidateName(relationship.Name);
                if (relationship.Fields.Count != 1 || relationship.NativeFlags != 0) throw new NotSupportedException("Native creation currently qualifies enforced single-field relationships without cascades.");
                TextKey(relationship.Name);
                Table parent = users[document.Tables.Items.IndexOf(relationship.Parent.Table)], child = users[document.Tables.Items.IndexOf(relationship.Child.Table)];
                int pc = relationship.Parent.Table.Columns.Items.IndexOf(relationship.Parent), cc = relationship.Child.Table.Columns.Items.IndexOf(relationship.Child);
                Index? unique = parent.Indexes.FirstOrDefault(i => (i.Flags & 1) != 0 && i.Columns.SequenceEqual(new[] { pc }));
                if (unique == null) throw new NotSupportedException("An enforced relationship requires a unique index over the referenced field.");
                if (child.Indexes.Any(i => StringComparer.OrdinalIgnoreCase.Equals(i.Name, relationship.Name))) throw new NotSupportedException("A relationship name conflicts with a child index.");
                Index parentLink = new Index(".r" + relations.Count, unique.Columns, unique.Flags, 2) { RootPage = unique.RootPage, RelatedType = 1, RelatedTable = child.DefinitionPage, RelatedIndex = child.Indexes.Count };
                if (ReferenceEquals(parent, child)) parentLink.RelatedIndex++;
                Index childLink = new Index(relationship.Name, new[] { cc }, 128, 2) { RootPage = Allocate(), RelatedType = 2, RelatedTable = parent.DefinitionPage, RelatedIndex = parent.Indexes.Count };
                parent.Indexes.Add(parentLink); child.Indexes.Add(childLink);
                if (parent.Indexes.Count > 32 || child.Indexes.Count > 32) throw new NotSupportedException("Native Access tables support at most 32 indexes, including relationship indexes.");
                HashSet<string> parentKeys = new HashSet<string>(parent.Rows.Select(row => Convert.ToBase64String(Key(parent, unique, row))));
                foreach (object?[] row in child.Rows) if (row[cc] != null && !parentKeys.Contains(Convert.ToBase64String(Key(child, childLink, row)))) throw new InvalidDataException("A modeled row violates the enforced relationship.");
                relations.Add(new object?[] { relationship.Name, 0, 1, 0, child.Name, relationship.Child.Name, parent.Name, relationship.Parent.Name });
            }
            CreateSystemTables(document, users, relations.ToArray());
            _tables.AddRange(users);
            foreach (Table table in _tables) {
                WriteRows(table);
                foreach (Index index in PhysicalIndexes(table)) WriteIndex(table, index);
                foreach (Index index in table.Indexes) index.UniqueCount = table.Indexes.First(i => i.RootPage == index.RootPage).UniqueCount;
                WriteMaps(table); WriteDefinition(table);
            }
            WriteGlobalMap();
        }
        private static Column CreateColumn(AccessColumn model) {
            Column column = model.DataType switch {
                AccessDataType.Boolean => new Column(model.Name, 1, 1), AccessDataType.Byte => new Column(model.Name, 2, 1),
                AccessDataType.Int16 => new Column(model.Name, 3, 2), AccessDataType.Int32 or AccessDataType.AutoNumber => new Column(model.Name, 4, 4),
                AccessDataType.Currency => new Column(model.Name, 5, 8), AccessDataType.Single => new Column(model.Name, 6, 4),
                AccessDataType.Double => new Column(model.Name, 7, 8), AccessDataType.DateTime => new Column(model.Name, 8, 8),
                AccessDataType.ShortText => new Column(model.Name, 10, model.MaxLength!.Value * 2, true),
                AccessDataType.LongText => new Column(model.Name, 12, 0, true), AccessDataType.Binary => new Column(model.Name, 11, 0, true),
                AccessDataType.Guid => new Column(model.Name, 15, 16), AccessDataType.Decimal => new Column(model.Name, 16, 17),
                _ => throw new NotSupportedException("This field type has no qualified native creation codec.")
            };
            column.AutoNumber = model.DataType == AccessDataType.AutoNumber;
            column.Precision = (byte)(model.Precision ?? 28); column.Scale = (byte)(model.Scale ?? 0);
            return column;
        }
        private static void ValidateName(string name) {
            AccessNativeBinary.Unicode.GetByteCount(name);
            if (name != name.Trim() || name.Any(c => c < 32 || c == '.' || c == '!' || c == '`' || c == '[' || c == ']'))
                throw new NotSupportedException("Native object names cannot have leading/trailing whitespace, control characters, periods, exclamation marks, backticks or square brackets.");
        }
    }
}
