using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;

namespace OfficeIMO.Excel {
    public partial class ExcelSheet {
        internal void CleanupTableArtifacts() {
            var worksheet = WorksheetRoot;
            var usedTableIds = CollectUsedTableIds(excludeWorksheetPart: _worksheetPart);
            var usedTableNames = CollectUsedTableNames(excludeWorksheetPart: _worksheetPart);
            var validRelationshipIds = new List<string>();

            foreach (var tableDefinitionPart in _worksheetPart.TableDefinitionParts.ToList()) {
                string? relationshipId = null;
                try {
                    relationshipId = _worksheetPart.GetIdOfPart(tableDefinitionPart);
                } catch {
                    // If the relationship is already broken, deleting the part is the safest repair.
                }

                var table = tableDefinitionPart.Table;
                if (table == null || !TryNormalizeTable(table, usedTableIds, usedTableNames)) {
                    DeleteTableDefinitionPart(tableDefinitionPart);
                    continue;
                }

                if (!string.IsNullOrWhiteSpace(relationshipId)) {
                    validRelationshipIds.Add(relationshipId!);
                }

                tableDefinitionPart.Table!.Save();
            }

            SyncWorksheetTableParts(worksheet, validRelationshipIds);
            RefreshWorkbookTableNameCache();
        }

        private bool TryNormalizeTable(Table table, HashSet<uint> usedTableIds, HashSet<string> usedTableNames) {
            string? tableRange = table.Reference?.Value;
            if (string.IsNullOrWhiteSpace(tableRange) || !A1.TryParseRange(tableRange!, out int startRow, out int startColumn, out _, out int endColumn)) {
                return false;
            }

            table.Id = NormalizeTableId(table.Id?.Value, usedTableIds);

            string normalizedName = NormalizeTableName(table.Name?.Value ?? table.DisplayName?.Value, table.Id!.Value, usedTableNames);
            table.Name = normalizedName;
            table.DisplayName = normalizedName;

            uint headerRowCount = table.HeaderRowCount?.Value ?? 1U;
            if (headerRowCount > 1U) {
                headerRowCount = 1U;
            }
            table.HeaderRowCount = headerRowCount;

            NormalizeTableColumns(table, startRow, startColumn, endColumn, headerRowCount > 0U);
            return true;
        }

        private void NormalizeTableColumns(Table table, int startRow, int startColumn, int endColumn, bool hasHeaderRow) {
            int width = endColumn - startColumn + 1;
            var existingColumns = table.TableColumns?.Elements<TableColumn>().ToList() ?? new List<TableColumn>();
            bool requiresRebuild = table.TableColumns == null ||
                                   existingColumns.Count != width ||
                                   existingColumns.Any(column => (column.Id?.Value ?? 0U) == 0U ||
                                                                 column.Id!.Value > (uint)width ||
                                                                 string.IsNullOrWhiteSpace(column.Name?.Value)) ||
                                   existingColumns.Select(column => column.Id!.Value).Distinct().Count() != existingColumns.Count;

            if (!requiresRebuild) {
                table.TableColumns!.Count = (uint)width;
                return;
            }

            var rebuiltColumns = new TableColumns { Count = (uint)width };
            var usedNames = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
            for (int index = 0; index < width; index++) {
                var source = index < existingColumns.Count ? existingColumns[index] : null;
                var rebuilt = source != null ? (TableColumn)source.CloneNode(true) : new TableColumn();
                rebuilt.Id = (uint)(index + 1);
                rebuilt.Name = EnsureUniqueColumnName(ResolveTableColumnName(source, startRow, startColumn + index, hasHeaderRow, index), usedNames);
                rebuiltColumns.Append(rebuilt);
            }

            if (table.TableColumns != null) {
                table.ReplaceChild(rebuiltColumns, table.TableColumns);
            } else if (table.TableStyleInfo != null) {
                table.InsertBefore(rebuiltColumns, table.TableStyleInfo);
            } else {
                table.Append(rebuiltColumns);
            }
        }

        private string ResolveTableColumnName(TableColumn? source, int headerRow, int columnIndex, bool hasHeaderRow, int zeroBasedIndex) {
            string? existingName = source?.Name?.Value;
            if (!string.IsNullOrWhiteSpace(existingName)) {
                return existingName!;
            }

            if (hasHeaderRow) {
                string headerCellValue = GetCellText(GetCell(headerRow, columnIndex));
                if (!string.IsNullOrWhiteSpace(headerCellValue)) {
                    return headerCellValue;
                }
            }

            return $"Column{zeroBasedIndex + 1}";
        }

        private static string EnsureUniqueColumnName(string proposedName, HashSet<string> usedNames) {
            string baseName = string.IsNullOrWhiteSpace(proposedName) ? "Column" : proposedName.Trim();
            string candidate = baseName;
            int suffix = 2;
            while (!usedNames.Add(candidate)) {
                candidate = $"{baseName} ({suffix++})";
            }

            return candidate;
        }

        private static uint NormalizeTableId(uint? proposedId, HashSet<uint> usedTableIds) {
            uint id = proposedId.GetValueOrDefault();
            if (id == 0U || !usedTableIds.Add(id)) {
                id = 1U;
                while (!usedTableIds.Add(id)) {
                    id++;
                }
            }

            return id;
        }

        private static string NormalizeTableName(string? proposedName, uint tableId, HashSet<string> usedTableNames) {
            const int MaxLength = 255;
            string fallback = $"Table{tableId}";
            string seed = string.IsNullOrWhiteSpace(proposedName) ? fallback : proposedName!;

            var sanitized = new System.Text.StringBuilder(seed.Length);
            foreach (char character in seed) {
                if (char.IsLetterOrDigit(character) || character == '_') {
                    sanitized.Append(character);
                } else if (char.IsWhiteSpace(character)) {
                    sanitized.Append('_');
                } else {
                    sanitized.Append('_');
                }
            }

            if (sanitized.Length == 0) {
                sanitized.Append(fallback);
            }

            if (char.IsDigit(sanitized[0])) {
                sanitized.Insert(0, '_');
            }

            if (sanitized.Length > MaxLength) {
                sanitized.Length = MaxLength;
            }

            string baseName = sanitized.ToString();
            string candidate = baseName;
            int suffix = 2;
            while (!usedTableNames.Add(candidate)) {
                string suffixText = suffix.ToString(System.Globalization.CultureInfo.InvariantCulture);
                int maxBaseLength = Math.Max(1, MaxLength - suffixText.Length);
                string trimmedBase = baseName.Length > maxBaseLength ? baseName.Substring(0, maxBaseLength) : baseName;
                candidate = trimmedBase + suffixText;
                suffix++;
            }

            return candidate;
        }

        private void SyncWorksheetTableParts(Worksheet worksheet, List<string> validRelationshipIds) {
            var tableParts = worksheet.Elements<TableParts>().FirstOrDefault();
            if (validRelationshipIds.Count == 0) {
                if (tableParts != null) {
                    worksheet.RemoveChild(tableParts);
                }
                return;
            }

            if (tableParts == null) {
                tableParts = new TableParts();
                worksheet.Append(tableParts);
            }

            var desiredIds = new HashSet<string>(validRelationshipIds, StringComparer.Ordinal);
            foreach (var tablePart in tableParts.Elements<TablePart>().ToList()) {
                string? id = tablePart.Id?.Value;
                if (string.IsNullOrWhiteSpace(id) || !desiredIds.Contains(id!)) {
                    tablePart.Remove();
                }
            }

            var existingIds = new HashSet<string>(
                tableParts.Elements<TablePart>().Select(tablePart => tablePart.Id?.Value).Where(id => !string.IsNullOrWhiteSpace(id))!.Cast<string>(),
                StringComparer.Ordinal);

            foreach (string relationshipId in validRelationshipIds) {
                if (existingIds.Add(relationshipId)) {
                    tableParts.Append(new TablePart { Id = relationshipId });
                }
            }

            if (!tableParts.Elements<TablePart>().Any()) {
                worksheet.RemoveChild(tableParts);
            } else {
                tableParts.Count = (uint)tableParts.Elements<TablePart>().Count();
            }
        }

        private void DeleteTableDefinitionPart(TableDefinitionPart tableDefinitionPart) {
            string? tableName = tableDefinitionPart.Table?.Name?.Value;
            if (!string.IsNullOrWhiteSpace(tableName)) {
                _excelDocument.RemoveReservedTableName(tableName!);
            }

            _worksheetPart.DeletePart(tableDefinitionPart);
        }

        private HashSet<uint> CollectUsedTableIds(WorksheetPart excludeWorksheetPart) {
            var usedIds = new HashSet<uint>();
            var workbookPart = WorkbookPartRoot;
            foreach (var worksheetPart in workbookPart.WorksheetParts) {
                if (worksheetPart == excludeWorksheetPart) {
                    continue;
                }

                foreach (var tableDefinitionPart in worksheetPart.TableDefinitionParts) {
                    uint id = tableDefinitionPart.Table?.Id?.Value ?? 0U;
                    if (id > 0U) {
                        usedIds.Add(id);
                    }
                }
            }

            return usedIds;
        }

        private HashSet<string> CollectUsedTableNames(WorksheetPart excludeWorksheetPart) {
            var usedNames = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
            var workbookPart = WorkbookPartRoot;
            foreach (var worksheetPart in workbookPart.WorksheetParts) {
                if (worksheetPart == excludeWorksheetPart) {
                    continue;
                }

                foreach (var tableDefinitionPart in worksheetPart.TableDefinitionParts) {
                    string? name = tableDefinitionPart.Table?.Name?.Value ?? tableDefinitionPart.Table?.DisplayName?.Value;
                    if (!string.IsNullOrWhiteSpace(name)) {
                        usedNames.Add(name!);
                    }
                }
            }

            return usedNames;
        }

        private void RefreshWorkbookTableNameCache() {
            var cache = _excelDocument.GetOrInitTableNameCache();
            cache.Clear();

            var workbookPart = WorkbookPartRoot;
            foreach (var worksheetPart in workbookPart.WorksheetParts) {
                foreach (var tableDefinitionPart in worksheetPart.TableDefinitionParts) {
                    string? name = tableDefinitionPart.Table?.Name?.Value;
                    if (!string.IsNullOrWhiteSpace(name)) {
                        cache.Add(name!);
                    }
                }
            }
        }
    }
}
