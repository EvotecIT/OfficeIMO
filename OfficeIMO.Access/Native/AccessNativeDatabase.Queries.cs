using System.Text;

namespace OfficeIMO.Access;

internal sealed partial class AccessNativeDatabase {
    private void LoadQueries(CancellationToken cancellation) {
        var records = new Dictionary<int, List<AccessQueryRecord>>();
        if (_tables.TryGetValue("MSysQueries", out var table)) {
            using var rows = new AccessNativeRowCursor(table, cancellation, rowLimit: MaxCatalogObjects);
            while (rows.Read(cancellation)) {
                int id = Convert.ToInt32(Field(table, rows, "ObjectId", cancellation));
                if (!records.TryGetValue(id, out var definitions)) records.Add(id, definitions = new List<AccessQueryRecord>());
                definitions.Add(new AccessQueryRecord(Convert.ToByte(Field(table, rows, "Attribute", cancellation)), Field(table, rows, "Name1", cancellation) as string,
                    Field(table, rows, "Name2", cancellation) as string, Field(table, rows, "Expression", cancellation) as string,
                    Field(table, rows, "Flag", cancellation) as short?, Field(table, rows, "LvExtra", cancellation) as int?, Field(table, rows, "Order", cancellation) as byte[], rows.Current.NativeBytes()));
            }
        }
        foreach (var record in _catalog.Where(x => x.Type == 5)) {
            cancellation.ThrowIfCancellationRequested(); records.TryGetValue(record.Id, out var definitions); definitions ??= new List<AccessQueryRecord>();
            var ordered = definitions.OrderBy(x => x.GetOrderBytes(), QueryOrderComparer.Instance).ToArray();
            string? sql = SelectSql(record.Flags, ordered);
            var query = new AccessQueryDefinition(_document, record.Name, sql, ordered) { NativeFlags = record.Flags };
            query.Parameters = Array.AsReadOnly(ordered.Where(x => x.Attribute == 2).Select(x => new AccessQueryParameter(x.Name1 ?? throw new InvalidDataException("Native Access query parameter has no name."), x.Flag ?? 0,
                x.Flag < 0 || x.Flag > 255 ? AccessDataType.Unknown : DataType(new AccessNativeColumn { Type = (byte)(x.Flag ?? 0) }))).ToArray());
            if (sql == null) query.Diagnostics = Array.AsReadOnly(new[] { new AccessDiagnostic("access.query.sql-not-decoded", "The query records are retained exactly. SQL reconstruction for this definition is unqualified; no query is executed.", query.Id) });
            _document.Queries.AddNativeItem(query);
        }
        _document.Queries.CatalogStatus = AccessCatalogStatus.Decoded;
    }
    private static string? SelectSql(int flags, AccessQueryRecord[] records) {
        int kind = flags & 0xf0;
        if (kind == 0x80) {
            if (records.Any(x => x.Attribute != 0 && x.Attribute != 255 && x.Attribute != 1 && x.Attribute != 3 && x.Attribute != 5 && x.Attribute != 11)) return null;
            var unionTypes = records.Where(x => x.Attribute == 1).ToArray(); if (unionTypes.Length != 1 || unionTypes[0].Flag != 9) return null;
            var parts = records.Where(x => x.Attribute == 5 && x.Expression != null).ToArray();
            if (parts.Length != 2 || !parts.Any(x => x.Name2 == "X7YZ_____1") || !parts.Any(x => x.Name2 == "X7YZ_____2")) return null;
            var unionFlags = records.Where(x => x.Attribute == 3).ToArray(); if (unionFlags.Length > 1 || unionFlags.Length == 1 && (unionFlags[0].Flag.GetValueOrDefault() & ~3) != 0) return null;
            bool distinct = ((unionFlags.FirstOrDefault()?.Flag ?? 0) & 2) != 0;
            var ordering = records.Where(x => x.Attribute == 11).ToArray(); if (ordering.Any(x => x.Expression == null || x.Name1 != null && x.Name1 != "D" && x.Name1 != "A")) return null;
            return parts.Single(x => x.Name2 == "X7YZ_____1").Expression!.Trim().TrimEnd(';') + "\nUNION" + (distinct ? "" : " ALL") + "\n" + parts.Single(x => x.Name2 == "X7YZ_____2").Expression!.Trim().TrimEnd(';')
                + (ordering.Length == 0 ? "" : "\nORDER BY " + string.Join(", ", ordering.Select(x => x.Expression + (x.Name1 == "D" ? " DESC" : "")))) + ";";
        }
        if (kind != 0 || records.Any(x => x.Attribute == 4 || x.Attribute == 7 || x.Attribute == 9 || x.Attribute == 10 || x.Attribute == 1 || x.Attribute != 0 && x.Attribute != 255 && x.Attribute != 2 && x.Attribute != 3 && x.Attribute != 5 && x.Attribute != 6 && x.Attribute != 8 && x.Attribute != 11)) return null;
        var flagRows = records.Where(x => x.Attribute == 3).ToArray(); if (flagRows.Length > 1 || flagRows.Length == 1 && (flagRows[0].Flag.GetValueOrDefault() & ~1) != 0) return null;
        var tables = records.Where(x => x.Attribute == 5).ToArray(); if (tables.Length != 1 || tables[0].Name1 == null) return null;
        var parameters = records.Where(x => x.Attribute == 2).ToArray(); var sql = new StringBuilder();
        if (parameters.Length != 0) {
            var declarations = new List<string>();
            foreach (var parameter in parameters) { string? type = ParameterType(parameter.Flag ?? 0); if (type == null || parameter.Name1 == null) return null; declarations.Add(Identifier(parameter.Name1) + " " + type); }
            sql.Append("PARAMETERS ").Append(string.Join(", ", declarations)).Append(";\n");
        }
        var columns = records.Where(x => x.Attribute == 6).ToArray();
        if (columns.Any(x => x.Expression == null || x.Flag.GetValueOrDefault() != 0)) return null;
        string[] expressions = columns.Select(x => x.Expression! + (x.Name1 == null ? "" : " AS " + Identifier(x.Name1))).ToArray();
        bool selectStar = ((flagRows.FirstOrDefault()?.Flag ?? 0) & 1) != 0;
        if (expressions.Length == 0 && !selectStar) return null;
        sql.Append("SELECT ").Append(string.Join(", ", expressions.Concat(selectStar ? new[] { "*" } : Array.Empty<string>()))).Append("\nFROM ").Append(Identifier(tables[0].Name1!));
        if (tables[0].Name2 != null) sql.Append(" AS ").Append(Identifier(tables[0].Name2!));
        var where = records.Where(x => x.Attribute == 8).ToArray(); if (where.Length > 1 || where.Length == 1 && where[0].Expression == null) return null;
        if (where.Length == 1) sql.Append("\nWHERE ").Append(where[0].Expression);
        var order = records.Where(x => x.Attribute == 11).ToArray(); if (order.Any(x => x.Expression == null || x.Name1 != null && x.Name1 != "D" && x.Name1 != "A")) return null;
        if (order.Length != 0) sql.Append("\nORDER BY ").Append(string.Join(", ", order.Select(x => x.Expression + (x.Name1 == "D" ? " DESC" : ""))));
        return sql.Append(';').ToString();
    }
    private static string Identifier(string name) => "[" + name.Replace("]", "]]") + "]";
    private static string? ParameterType(int type) => type switch { 1 => "Bit", 2 => "Byte", 3 => "Short", 4 => "Long", 5 => "Currency", 6 => "IEEESingle", 7 => "IEEEDouble", 8 => "DateTime", 9 => "Binary", 10 => "Text", 11 => "LongBinary", 15 => "Guid", _ => null };
    private sealed class QueryOrderComparer : IComparer<byte[]?> {
        internal static readonly QueryOrderComparer Instance = new QueryOrderComparer();
        public int Compare(byte[]? first, byte[]? second) { if (first == null) return second == null ? 0 : -1; if (second == null) return 1; for (int i = 0; i < Math.Min(first.Length, second.Length); i++) { int comparison = first[i].CompareTo(second[i]); if (comparison != 0) return comparison; } return first.Length.CompareTo(second.Length); }
    }
}
