using System;
using System.Collections.Generic;
using System.Globalization;

namespace OfficeIMO.Reader;

public static partial class OfficeDocumentStructuredExtractor {
    private static void AddMetadata(OfficeDocumentReadResult document, ExtractionState state) {
        IReadOnlyList<OfficeDocumentMetadataEntry> metadata = document.Metadata ?? Array.Empty<OfficeDocumentMetadataEntry>();
        for (int index = 0; index < metadata.Count; index++) {
            state.CancellationToken.ThrowIfCancellationRequested();
            OfficeDocumentMetadataEntry entry = metadata[index];
            var attributes = CopyAttributes(entry.Attributes);
            if (!string.IsNullOrWhiteSpace(entry.Category)) attributes["metadataCategory"] = entry.Category;
            if (!state.TryAdd(new OfficeDocumentStructuredRecord {
                Id = "metadata-" + index.ToString("D4", CultureInfo.InvariantCulture),
                Category = "metadata",
                Name = entry.Name ?? string.Empty,
                Value = entry.Value,
                ValueType = entry.ValueType,
                SourceObjectId = string.IsNullOrWhiteSpace(entry.SourceObjectId) ? entry.Id : entry.SourceObjectId,
                Location = entry.Location,
                Attributes = attributes
            })) return;
        }
    }

    private static void AddForms(IReadOnlyList<OfficeDocumentFormField> forms, ExtractionState state) {
        for (int index = 0; index < forms.Count; index++) {
            state.CancellationToken.ThrowIfCancellationRequested();
            OfficeDocumentFormField form = forms[index];
            if (!state.TryAdd(new OfficeDocumentStructuredRecord {
                Id = "form-" + index.ToString("D4", CultureInfo.InvariantCulture),
                Category = "form",
                Name = string.IsNullOrWhiteSpace(form.Name) ? form.Id : form.Name!,
                Value = form.Value,
                ValueType = form.Kind,
                SourceObjectId = form.Id,
                Location = form.Location,
                Attributes = new SortedDictionary<string, string>(StringComparer.Ordinal) {
                    ["isReadOnly"] = form.IsReadOnly ? "true" : "false",
                    ["isRequired"] = form.IsRequired ? "true" : "false"
                }
            })) return;
        }
    }

    private static void AddKeyValueRows(OfficeDocumentReadResult document, ExtractionState state) {
        int tableIndex = 0;
        foreach (ReaderTable table in OfficeDocumentModelTraversal.Tables(document)) {
            state.CancellationToken.ThrowIfCancellationRequested();
            IReadOnlyList<string> columns = table.Columns ?? Array.Empty<string>();
            if (IsShapeDataTable(table)) {
                tableIndex++;
                continue;
            }

            int pathIndex = FindColumn(columns, "path");
            int typeIndex = FindColumn(columns, "type");
            int valueIndex = FindColumn(columns, "value");
            bool pathValue = pathIndex >= 0 && valueIndex >= 0;
            bool twoColumn = columns.Count == 2;
            if (!pathValue && !twoColumn) {
                tableIndex++;
                continue;
            }

            IReadOnlyList<IReadOnlyList<string>> rows = table.Rows ?? Array.Empty<IReadOnlyList<string>>();
            for (int rowIndex = 0; rowIndex < rows.Count; rowIndex++) {
                IReadOnlyList<string> row = rows[rowIndex] ?? Array.Empty<string>();
                int nameIndex = pathValue ? pathIndex : 0;
                int scalarIndex = pathValue ? valueIndex : 1;
                string name = GetCell(row, nameIndex).Trim();
                if (name.Length == 0) continue;
                var attributes = BuildTableAttributes(table, tableIndex, rowIndex);
                string category = pathValue ? "structured-value" : "key-value";
                if (!state.TryAdd(new OfficeDocumentStructuredRecord {
                    Id = category + "-" + tableIndex.ToString("D4", CultureInfo.InvariantCulture) + "-" + rowIndex.ToString("D6", CultureInfo.InvariantCulture),
                    Category = category,
                    Name = name,
                    Value = GetCell(row, scalarIndex),
                    ValueType = typeIndex >= 0 ? GetCell(row, typeIndex) : null,
                    Location = table.Location,
                    Attributes = attributes
                })) return;
            }
            tableIndex++;
        }
    }

    private static void AddShapeData(OfficeDocumentReadResult document, ExtractionState state) {
        int tableIndex = 0;
        foreach (ReaderTable table in OfficeDocumentModelTraversal.Tables(document)) {
            state.CancellationToken.ThrowIfCancellationRequested();
            if (!IsShapeDataTable(table)) {
                tableIndex++;
                continue;
            }

            IReadOnlyList<string> columns = table.Columns ?? Array.Empty<string>();
            int ownerTypeIndex = FindColumn(columns, "ownertype");
            int ownerIdIndex = FindColumn(columns, "ownerid");
            int ownerTextIndex = FindColumn(columns, "ownertext");
            int nameIndex = FindColumn(columns, "name");
            int labelIndex = FindColumn(columns, "label");
            int valueIndex = FindColumn(columns, "value");
            int typeIndex = FindColumn(columns, "type");
            int promptIndex = FindColumn(columns, "prompt");
            IReadOnlyList<IReadOnlyList<string>> rows = table.Rows ?? Array.Empty<IReadOnlyList<string>>();
            for (int rowIndex = 0; rowIndex < rows.Count; rowIndex++) {
                IReadOnlyList<string> row = rows[rowIndex] ?? Array.Empty<string>();
                string label = GetCell(row, labelIndex).Trim();
                string name = label.Length == 0 ? GetCell(row, nameIndex).Trim() : label;
                if (name.Length == 0) continue;
                string ownerType = GetCell(row, ownerTypeIndex).Trim();
                string ownerId = GetCell(row, ownerIdIndex).Trim();
                var attributes = BuildTableAttributes(table, tableIndex, rowIndex);
                AddIfPresent(attributes, "ownerType", ownerType);
                AddIfPresent(attributes, "ownerId", ownerId);
                AddIfPresent(attributes, "ownerText", GetCell(row, ownerTextIndex));
                AddIfPresent(attributes, "prompt", GetCell(row, promptIndex));
                if (!state.TryAdd(new OfficeDocumentStructuredRecord {
                    Id = "shape-data-" + tableIndex.ToString("D4", CultureInfo.InvariantCulture) + "-" + rowIndex.ToString("D6", CultureInfo.InvariantCulture),
                    Category = "shape-data",
                    Name = name,
                    Value = GetCell(row, valueIndex),
                    ValueType = GetCell(row, typeIndex),
                    SourceObjectId = ownerType.Length == 0 ? ownerId : ownerType + ":" + ownerId,
                    Location = table.Location,
                    Attributes = attributes
                })) return;
            }
            tableIndex++;
        }
    }

    private static bool IsShapeDataTable(ReaderTable table) =>
        string.Equals(table.Kind?.Trim(), "visio-shape-data", StringComparison.OrdinalIgnoreCase);

    private static int FindColumn(IReadOnlyList<string> columns, string name) {
        for (int index = 0; index < columns.Count; index++) {
            if (string.Equals(columns[index]?.Trim(), name, StringComparison.OrdinalIgnoreCase)) return index;
        }
        return -1;
    }

    private static string GetCell(IReadOnlyList<string> row, int index) =>
        index >= 0 && index < row.Count ? row[index] ?? string.Empty : string.Empty;

    private static SortedDictionary<string, string> BuildTableAttributes(ReaderTable table, int tableIndex, int rowIndex) {
        var attributes = new SortedDictionary<string, string>(StringComparer.Ordinal) {
            ["rowIndex"] = rowIndex.ToString(CultureInfo.InvariantCulture),
            ["tableIndex"] = tableIndex.ToString(CultureInfo.InvariantCulture),
            ["tableTruncated"] = table.Truncated ? "true" : "false",
            ["totalRowCount"] = table.TotalRowCount.ToString(CultureInfo.InvariantCulture)
        };
        AddIfPresent(attributes, "tableKind", table.Kind);
        AddIfPresent(attributes, "tableTitle", table.Title);
        return attributes;
    }

    private static SortedDictionary<string, string> CopyAttributes(IReadOnlyDictionary<string, string>? source) {
        var attributes = new SortedDictionary<string, string>(StringComparer.Ordinal);
        if (source == null) return attributes;
        foreach (KeyValuePair<string, string> item in source) attributes[item.Key] = item.Value;
        return attributes;
    }

    private static void AddIfPresent(IDictionary<string, string> attributes, string name, string? value) {
        if (!string.IsNullOrWhiteSpace(value)) attributes[name] = value!.Trim();
    }
}
