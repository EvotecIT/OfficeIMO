namespace OfficeIMO.Access;

internal sealed partial class AccessNativeDatabase {
    private void LoadComplexDefinitions(CancellationToken cancellation) {
        if (!_tables.TryGetValue("MSysComplexColumns", out var metadata)) return;
        RequireFields(metadata, "ComplexID", "ConceptualTableID", "FlatTableID", "ComplexTypeObjectID", "ColumnName");
        var definitions = new Dictionary<int, (int Table, int Flat, int Type, string? Column)>();
        using (var rows = new AccessNativeRowCursor(metadata, cancellation, rowLimit: MaxCatalogObjects)) while (rows.Read(cancellation)) {
            int id = Convert.ToInt32(RequiredField(metadata, rows, "ComplexID", cancellation));
            if (definitions.ContainsKey(id) || definitions.Count == MaxCatalogObjects) throw new InvalidDataException("Native Access complex definitions are ambiguous or exceed their limit.");
            definitions.Add(id, (Convert.ToInt32(RequiredField(metadata, rows, "ConceptualTableID", cancellation)), Convert.ToInt32(RequiredField(metadata, rows, "FlatTableID", cancellation)),
                Convert.ToInt32(RequiredField(metadata, rows, "ComplexTypeObjectID", cancellation)), RequiredField(metadata, rows, "ColumnName", cancellation) as string));
        }
        foreach (var table in _tables.Values.ToArray()) foreach (var column in table.Columns.Where(x => x.Type == 18)) {
            cancellation.ThrowIfCancellationRequested();
            if (!definitions.TryGetValue(column.ComplexId, out var definition) || definition.Table != table.DefinitionPage || !StringComparer.OrdinalIgnoreCase.Equals(column.Name, definition.Column)) throw new InvalidDataException("Native Access structured field does not match its complex catalog definition.");
            if (!_definitions.TryGetValue(definition.Flat, out var flat) || !_definitions.TryGetValue(definition.Type, out var type) || flat.Model == null) throw new InvalidDataException("Native Access complex backing schema is missing.");
            var values = flat.Columns.Where(x => type.Columns.Any(t => StringComparer.OrdinalIgnoreCase.Equals(t.Name, x.Name))).ToArray();
            var keys = flat.Columns.Where(x => !values.Contains(x)).ToArray();
            var foreign = keys.Where(x => (x.ExtraFlags & 8) != 0).ToArray();
            if (foreign.Length == 0) foreign = keys.Where(x => x.Type == 4 && (x.Flags & 4) == 0).ToArray();
            if (foreign.Length != 1 || !keys.Any(x => (x.Flags & 4) != 0)) throw new InvalidDataException("Native Access complex key fields are missing or ambiguous.");
            AccessComplexKind kind = type.Name.Equals("MSysComplexType_Attachment", StringComparison.OrdinalIgnoreCase) ? AccessComplexKind.Attachment
                : type.Name.StartsWith("MSysComplexTypeVH_", StringComparison.OrdinalIgnoreCase) ? AccessComplexKind.VersionHistory
                : values.Length == 1 ? AccessComplexKind.MultiValue : AccessComplexKind.Unknown;
            column.ComplexDefinition = new AccessComplexDefinition(flat, flat.Columns.IndexOf(foreign[0]), kind, values.Select(x => x.Model!).ToArray());
            column.Model!.ComplexDefinition = column.ComplexDefinition;
            if (kind == AccessComplexKind.Unknown) column.Model.Diagnostics = Array.AsReadOnly(new[] { new AccessDiagnostic("access.complex.opaque-structure", "The complex backing rows and exact field payloads remain available without coercion.", column.Model.Id) });
        }
    }
}
