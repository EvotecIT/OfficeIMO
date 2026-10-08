using System.Collections.ObjectModel;
using static OfficeIMO.Access.AccessNativeBinary;

namespace OfficeIMO.Access;

internal sealed partial class AccessNativeDatabase {
    private readonly Dictionary<string, AccessNativeTable> _tables = new Dictionary<string, AccessNativeTable>(StringComparer.OrdinalIgnoreCase);
    private readonly List<NativeCatalogRecord> _catalog = new List<NativeCatalogRecord>();
    internal void LoadCatalog(CancellationToken cancellation) {
        var catalog = Definition(2, "MSysObjects", cancellation);
        RequireFields(catalog, "Id", "Name", "Type", "Flags");
        using (var rows = new AccessNativeRowCursor(catalog, cancellation, rowLimit: MaxCatalogObjects)) {
            while (rows.Read(cancellation)) {
                if (_catalog.Count == MaxCatalogObjects) throw new InvalidDataException("Native Access catalog exceeds MaxCatalogObjects.");
                var record = new NativeCatalogRecord {
                    Id = Convert.ToInt32(RequiredField(catalog, rows, "Id", cancellation)),
                    Name = RequiredName(catalog, rows, "Name", cancellation),
                    Type = Convert.ToInt32(RequiredField(catalog, rows, "Type", cancellation)),
                    Flags = Convert.ToInt32(RequiredField(catalog, rows, "Flags", cancellation)),
                    ParentId = Field(catalog, rows, "ParentId", cancellation) as int?,
                    Properties = Field(catalog, rows, "LvProp", cancellation) as byte[],
                    Source = Field(catalog, rows, "Database", cancellation) as string,
                    ForeignTable = Field(catalog, rows, "ForeignName", cancellation) as string,
                    Connection = Field(catalog, rows, "Connect", cancellation) as string
                };
                _catalog.Add(record);
                // Different catalog namespaces can legitimately reuse a name. Their typed collections check ambiguity independently.
                var entry = new AccessCatalogEntry(_document, record.Name, record.Type, record.Flags) {
                    NativeId = record.Id, NativeParentId = record.ParentId,
                    NativeRecord = new AccessOpaqueValue(0, rows.Current.NativeBytes(), "Exact catalog row, including uninterpreted object metadata. Explicit raw inspection may contain credential-bearing metadata."),
                    Owner = Field(catalog, rows, "Owner", cancellation) is byte[] owner ? new AccessOpaqueValue(9, owner, "Persisted security identifier; no authentication is performed.") : null
                };
                var payloads = new Dictionary<string, AccessOpaqueValue>(StringComparer.OrdinalIgnoreCase);
                foreach (string field in new[] { "Lv", "LvModule", "LvExtra" }) if (Field(catalog, rows, field, cancellation) is byte[] bytes)
                    payloads.Add(field, new AccessOpaqueValue(11, bytes, "Exact catalog application metadata; typed semantics remain unqualified."));
                entry.NativePayloads = new ReadOnlyDictionary<string, AccessOpaqueValue>(payloads);
                _document.Catalog.Items.Add(entry);
            }
        }
        foreach (var record in _catalog.Where(x => x.Type == 2 && x.Name == "MSysDb")) {
            var metadata = new AccessTable(_document, record.Name); LoadProperties(metadata, record.Properties, cancellation);
            _document.Properties = metadata.Properties; _document.NativeProperties = metadata.NativeProperties;
            _document.Diagnostics = metadata.Diagnostics;
        }
        foreach (var record in _catalog.Where(x => x.Type == 1)) {
            cancellation.ThrowIfCancellationRequested();
            if (record.Id <= 0) throw new InvalidDataException("Native Access local table has an invalid definition reference.");
            bool system = (record.Flags & unchecked((int)0x80000002)) != 0;
            if (!system && SelectedTables != null && !SelectedTables.Contains(record.Name)) continue;
            var definition = Definition(record.Id & 0x00ffffff, record.Name, cancellation);
            if (_tables.ContainsKey(record.Name)) throw new InvalidDataException("Native Access table names are ambiguous.");
            _tables.Add(record.Name, definition);
            var model = Model(definition, system);
            LoadProperties(model, record.Properties, cancellation);
            if (system) _document.SystemTables.AddNativeItem(model); else _document.Tables.AddNativeItem(model);
        }
        foreach (var record in _catalog.Where(x => x.Type == 4 || x.Type == 6)) {
            if (SelectedTables != null && !SelectedTables.Contains(record.Name)) continue;
            var table = new AccessTable(_document, record.Name) { LinkedTable = new AccessLinkedTableInfo(record.Source, record.ForeignTable, RedactConnection(record.Connection)) };
            table.Columns.CatalogStatus = table.Indexes.CatalogStatus = AccessCatalogStatus.NotDecoded;
            table.Diagnostics = Array.AsReadOnly(new[] { new AccessDiagnostic("access.linked-table.inert", "Linked-table metadata is inspected without opening its source. External schema and rows are unavailable.", table.Id) });
            _document.Tables.AddNativeItem(table);
        }
        _document.CatalogStatus = AccessCatalogStatus.Decoded;
        _document.Tables.CatalogStatus = _document.SystemTables.CatalogStatus = _document.Catalog.CatalogStatus = AccessCatalogStatus.Decoded;
        LoadComplexDefinitions(cancellation);
        LoadRelationships(cancellation);
        LoadQueries(cancellation);
        if (DecodeApplicationObjects) LoadApplicationObjects(cancellation);
    }
    private AccessTable Model(AccessNativeTable definition, bool system) {
        var table = new AccessTable(_document, definition.Name) { NativeTable = definition, IsSystem = system }; definition.Model = table;
        foreach (var native in definition.Columns) {
            AccessDataType dataType = DataType(native);
            var column = new AccessColumn(table, native.Name, dataType, dataType == AccessDataType.ShortText ? native.Size / 2 : (int?)null) {
                IsAutoNumber = (native.Flags & 0x44) != 0, AutoNumberSeed = null, IsHyperlink = (native.Flags & 0x80) != 0, IsCalculated = native.Calculated,
                Precision = native.Type == 16 ? native.Precision : (int?)null, Scale = native.Type == 16 ? native.Scale : (int?)null
            };
            native.Model = column; table.Columns.AddNativeItem(column);
            native.RedactConnection = definition.Name == "MSysObjects" && native.Name == "Connect";
            string? opaque = native.Calculated ? "calculated" : column.DataType == AccessDataType.Unknown ? "unknown-type" : native.Type == 16 && (native.Precision < 1 || native.Precision > 28 || native.Scale > 28) ? "decimal-precision" : null;
            if (opaque != null) column.Diagnostics = Array.AsReadOnly(new[] { new AccessDiagnostic("access.value.opaque." + opaque, "The field's native payload is retained exactly without evaluation, narrowing or coercion.", column.Id) });
        }
        foreach (var native in definition.Indexes) table.Indexes.AddNativeItem(new AccessIndex(table, native.Name,
            native.Columns.Select(x => x.Model!).ToArray(), native.Type == 1, (native.Flags & 1) != 0, native.Type == 2, native.Descending));
        table.Columns.CatalogStatus = table.Indexes.CatalogStatus = AccessCatalogStatus.Decoded;
        return table;
    }
    private static object? Field(AccessNativeTable table, IAccessRowCursor rows, string name, CancellationToken cancellation) {
        int ordinal = table.Columns.FindIndex(x => StringComparer.OrdinalIgnoreCase.Equals(x.Name, name));
        return ordinal < 0 ? null : rows.GetValue(ordinal, cancellation);
    }
    private static void RequireFields(AccessNativeTable table, params string[] names) {
        foreach (string name in names) if (!table.Columns.Any(x => StringComparer.OrdinalIgnoreCase.Equals(x.Name, name)))
            throw new InvalidDataException($"Native Access system table '{table.Name}' is missing required field '{name}'.");
    }
    private static object RequiredField(AccessNativeTable table, IAccessRowCursor rows, string name, CancellationToken cancellation) =>
        Field(table, rows, name, cancellation) ?? throw new InvalidDataException($"Native Access system field '{table.Name}.{name}' is missing or null.");
    private static string RequiredName(AccessNativeTable table, IAccessRowCursor rows, string name, CancellationToken cancellation) {
        if (!(RequiredField(table, rows, name, cancellation) is string value)) throw new InvalidDataException($"Native Access system field '{table.Name}.{name}' is not a name.");
        try { return AccessNamedObject.ValidateName(value); }
        catch (ArgumentException exception) { throw new InvalidDataException($"Native Access system field '{table.Name}.{name}' has an invalid name.", exception); }
    }
    private void LoadRelationships(CancellationToken cancellation) {
        if (_tables.TryGetValue("MSysRelationships", out var definition)) {
            RequireFields(definition, "szRelationship", "szReferencedObject", "szReferencedColumn", "szObject", "szColumn", "icolumn", "ccolumn", "grbit");
            var records = new Dictionary<string, List<NativeRelationshipField>>(StringComparer.OrdinalIgnoreCase);
            using (var rows = new AccessNativeRowCursor(definition, cancellation, rowLimit: checked((long)MaxCatalogObjects * 10))) while (rows.Read(cancellation)) {
                string name = RequiredName(definition, rows, "szRelationship", cancellation);
                if (!records.TryGetValue(name, out var fields)) records.Add(name, fields = new List<NativeRelationshipField>());
                if (records.Count > MaxCatalogObjects || fields.Count == 10) throw new InvalidDataException("Native Access relationship metadata exceeds its limit.");
                fields.Add(new NativeRelationshipField {
                    ParentTable = RequiredName(definition, rows, "szReferencedObject", cancellation), ParentColumn = RequiredName(definition, rows, "szReferencedColumn", cancellation),
                    ChildTable = RequiredName(definition, rows, "szObject", cancellation), ChildColumn = RequiredName(definition, rows, "szColumn", cancellation),
                    Ordinal = Convert.ToInt32(RequiredField(definition, rows, "icolumn", cancellation)), Count = Convert.ToInt32(RequiredField(definition, rows, "ccolumn", cancellation)),
                    Flags = Convert.ToInt32(RequiredField(definition, rows, "grbit", cancellation))
                });
            }
            foreach (var pair in records) {
                var fields = pair.Value.OrderBy(x => x.Ordinal).ToArray();
                if (fields.Length != fields[0].Count || fields.Where((x, i) => x.Ordinal != i || x.Count != fields.Length || x.Flags != fields[0].Flags
                    || !StringComparer.OrdinalIgnoreCase.Equals(x.ParentTable, fields[0].ParentTable) || !StringComparer.OrdinalIgnoreCase.Equals(x.ChildTable, fields[0].ChildTable)).Any()) throw new InvalidDataException("Native Access relationship fields are incomplete or inconsistent.");
                if (!_catalog.Any(x => (x.Type == 1 || x.Type == 4 || x.Type == 6) && StringComparer.OrdinalIgnoreCase.Equals(x.Name, fields[0].ParentTable))
                    || !_catalog.Any(x => (x.Type == 1 || x.Type == 4 || x.Type == 6) && StringComparer.OrdinalIgnoreCase.Equals(x.Name, fields[0].ChildTable))) throw new InvalidDataException("Native Access relationship refers to a missing table.");
                if (!_tables.TryGetValue(fields[0].ParentTable, out var parent) || parent.Model == null || !_tables.TryGetValue(fields[0].ChildTable, out var child) || child.Model == null) continue;
                var mappings = fields.Select(x => new AccessRelationshipField(parent.Model.Columns[x.ParentColumn!], child.Model.Columns[x.ChildColumn!])).ToArray();
                _document.Relationships.AddNativeItem(new AccessRelationship(_document, pair.Key, mappings, fields[0].Flags));
            }
        }
        _document.Relationships.CatalogStatus = AccessCatalogStatus.Decoded;
    }
    internal static string? RedactConnection(string? connection) {
        if (connection == null) return null;
        // DbConnectionStringBuilder handles quoted semicolons and escaped credential values without provider activation.
        try {
            var builder = new System.Data.Common.DbConnectionStringBuilder();
            string content = connection.TrimStart(';'); int separator = content.IndexOf(';');
            string prefix = separator >= 0 && content.Substring(0, separator).IndexOf('=') < 0 ? content.Substring(0, separator + 1) : string.Empty;
            builder.ConnectionString = content.Substring(prefix.Length);
            foreach (string key in builder.Keys.Cast<string>().ToArray()) {
                if (key.IndexOf(';') >= 0) return "[redacted: unparsed connection metadata]";
                if (key.Equals("PWD", StringComparison.OrdinalIgnoreCase) || key.IndexOf("password", StringComparison.OrdinalIgnoreCase) >= 0 || key.Equals("UID", StringComparison.OrdinalIgnoreCase) || key.Equals("User ID", StringComparison.OrdinalIgnoreCase) || key.IndexOf("token", StringComparison.OrdinalIgnoreCase) >= 0 || key.IndexOf("secret", StringComparison.OrdinalIgnoreCase) >= 0) builder[key] = "[redacted]";
            }
            return prefix + builder.ConnectionString;
        } catch (ArgumentException) { return "[redacted: unparsed connection metadata]"; }
    }
    private sealed class NativeCatalogRecord {
        internal int Id, Type, Flags; internal int? ParentId; internal string Name = string.Empty;
        internal byte[]? Properties; internal string? Source, ForeignTable, Connection;
    }
    private sealed class NativeRelationshipField {
        internal string ParentTable = string.Empty, ParentColumn = string.Empty, ChildTable = string.Empty, ChildColumn = string.Empty; internal int Ordinal, Count, Flags;
    }
}
