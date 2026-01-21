using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using System.Xml.Linq;

namespace OfficeIMO.Excel {
    public partial class ExcelSheet {
        /// <summary>
        /// Returns pivot tables defined on this worksheet.
        /// </summary>
        public IReadOnlyList<ExcelPivotTableInfo> GetPivotTables() {
            return Locking.ExecuteRead(_excelDocument.EnsureLock(), () => {
                var list = new List<ExcelPivotTableInfo>();
                var workbookPart = _spreadSheetDocument.WorkbookPart;
                if (workbookPart == null) return list;

                var cacheMap = BuildPivotCacheMap(workbookPart);
                var sheetIndex = ResolveSheetIndex(workbookPart);

                foreach (var pivotPart in _worksheetPart.PivotTableParts) {
                    var def = pivotPart.PivotTableDefinition;
                    if (def == null) continue;

                    uint cacheId = def.CacheId?.Value ?? 0U;
                    cacheMap.TryGetValue(cacheId, out var cacheDef);
                    var cacheFields = BuildCacheFieldNames(cacheDef);
                    var sourceSheet = cacheDef?.CacheSource?.WorksheetSource?.Sheet?.Value;
                    var sourceRange = cacheDef?.CacheSource?.WorksheetSource?.Reference?.Value;

                    var rowFields = ResolveFieldNames(def.RowFields?.Elements<Field>(), cacheFields);
                    var columnFields = ResolveFieldNames(def.ColumnFields?.Elements<Field>(), cacheFields);
                    var pageFields = ResolvePageFieldNames(def.PageFields?.Elements<PageField>(), cacheFields);
                    var dataFields = ResolveDataFields(def.DataFields?.Elements<DataField>(), cacheFields);

                    var layout = ResolveLayout(def.CompactData, def.OutlineData);

                    list.Add(new ExcelPivotTableInfo(
                        name: def.Name?.Value ?? string.Empty,
                        cacheId: cacheId,
                        location: def.Location?.Reference?.Value,
                        sourceSheet: sourceSheet,
                        sourceRange: sourceRange,
                        sheetName: Name,
                        sheetIndex: sheetIndex,
                        pivotStyle: def.PivotTableStyleName?.Value,
                        layout: layout,
                        dataOnRows: def.DataOnRows?.Value,
                        showHeaders: def.ShowHeaders?.Value,
                        showEmptyRows: def.ShowEmptyRow?.Value,
                        showEmptyColumns: def.ShowEmptyColumn?.Value,
                        showDrill: def.ShowDrill?.Value,
                        rowFields: rowFields,
                        columnFields: columnFields,
                        pageFields: pageFields,
                        dataFields: dataFields));
                }

                return list;
            });
        }

        /// <summary>
        /// Adds a basic pivot table based on a source range and places it at a destination cell.
        /// </summary>
        /// <param name="sourceRange">Source data range (including header row), e.g. "A1:D100".</param>
        /// <param name="destinationCell">Top-left cell for the pivot table (e.g. "F2").</param>
        /// <param name="name">Optional pivot table name. Defaults to "PivotTable1" style.</param>
        /// <param name="rowFields">Optional row fields (header names).</param>
        /// <param name="columnFields">Optional column fields (header names).</param>
        /// <param name="pageFields">Optional page fields (header names) used as filters.</param>
        /// <param name="dataFields">Optional data field definitions. Defaults to last column with Sum.</param>
        /// <param name="showRowGrandTotals">Show row grand totals.</param>
        /// <param name="showColumnGrandTotals">Show column grand totals.</param>
        /// <param name="pivotStyleName">Optional pivot table style name.</param>
        /// <param name="layout">Layout mode (Compact, Outline, Tabular).</param>
        /// <param name="dataOnRows">Whether to show data fields on rows instead of columns.</param>
        /// <param name="showHeaders">Whether to show field headers.</param>
        /// <param name="showEmptyRows">Whether to show empty rows.</param>
        /// <param name="showEmptyColumns">Whether to show empty columns.</param>
        /// <param name="showDrill">Whether to show drill indicators.</param>
        public void AddPivotTable(
            string sourceRange,
            string destinationCell,
            string? name = null,
            IEnumerable<string>? rowFields = null,
            IEnumerable<string>? columnFields = null,
            IEnumerable<string>? pageFields = null,
            IEnumerable<ExcelPivotDataField>? dataFields = null,
            bool showRowGrandTotals = true,
            bool showColumnGrandTotals = true,
            string? pivotStyleName = null,
            ExcelPivotLayout layout = ExcelPivotLayout.Compact,
            bool? dataOnRows = null,
            bool? showHeaders = null,
            bool? showEmptyRows = null,
            bool? showEmptyColumns = null,
            bool? showDrill = null) {
            if (string.IsNullOrWhiteSpace(sourceRange)) throw new ArgumentNullException(nameof(sourceRange));
            if (string.IsNullOrWhiteSpace(destinationCell)) throw new ArgumentNullException(nameof(destinationCell));
            if (!A1.TryParseRange(sourceRange, out int r1, out int c1, out int r2, out int c2)) {
                throw new ArgumentException($"Invalid A1 range '{sourceRange}'.", nameof(sourceRange));
            }

            var (destRow, destCol) = A1.ParseCellRef(destinationCell);
            if (destRow <= 0 || destCol <= 0) {
                throw new ArgumentException($"Invalid destination cell '{destinationCell}'.", nameof(destinationCell));
            }

            WriteLock(() => {
                var headers = BuildPivotHeaders(r1, c1, c2);
                if (headers.Count == 0) {
                    throw new InvalidOperationException("Pivot source range must include at least one header column.");
                }

                var headerIndex = new Dictionary<string, int>(StringComparer.OrdinalIgnoreCase);
                for (int i = 0; i < headers.Count; i++) {
                    headerIndex[headers[i]] = i;
                }

                var dataFieldList = (dataFields ?? Array.Empty<ExcelPivotDataField>()).Where(df => df != null).ToList();
                if (dataFieldList.Count == 0) {
                    dataFieldList.Add(new ExcelPivotDataField(headers[headers.Count - 1], DataConsolidateFunctionValues.Sum));
                }

                var rowFieldIndices = ResolveFieldIndices(rowFields, headerIndex, nameof(rowFields));
                var columnFieldIndices = ResolveFieldIndices(columnFields, headerIndex, nameof(columnFields));
                var pageFieldIndices = ResolveFieldIndices(pageFields, headerIndex, nameof(pageFields));
                if (pageFieldIndices.Count > 0) {
                    rowFieldIndices.RemoveAll(idx => pageFieldIndices.Contains(idx));
                    columnFieldIndices.RemoveAll(idx => pageFieldIndices.Contains(idx));
                }

                // If no row/column fields provided, default to the first non-data field when possible.
                if (rowFields == null && columnFields == null && rowFieldIndices.Count == 0 && columnFieldIndices.Count == 0) {
                    int dataIdx = ResolveFieldIndex(dataFieldList[0].FieldName, headerIndex, nameof(dataFields));
                    int fallback = headers.Count > 1 && dataIdx == 0 ? 1 : 0;
                    if (fallback >= 0 && fallback < headers.Count && fallback != dataIdx) {
                        rowFieldIndices.Add(fallback);
                    }
                }

                var dataFieldIndices = new HashSet<int>();
                foreach (var df in dataFieldList) {
                    int idx = ResolveFieldIndex(df.FieldName, headerIndex, nameof(dataFields));
                    dataFieldIndices.Add(idx);
                }

                var workbookPart = _spreadSheetDocument.WorkbookPart ?? throw new InvalidOperationException("WorkbookPart is null.");
                uint cacheId = NextPivotCacheId(workbookPart);

                var cacheDefPart = workbookPart.AddNewPart<PivotTableCacheDefinitionPart>();
                var cacheDef = new PivotCacheDefinition {
                    CacheSource = new CacheSource {
                        Type = SourceValues.Worksheet,
                        WorksheetSource = new WorksheetSource {
                            Sheet = Name,
                            Reference = sourceRange
                        }
                    },
                    CacheFields = new CacheFields { Count = (uint)headers.Count },
                    RecordCount = 0,
                    RefreshOnLoad = true,
                    SaveData = false
                };

                foreach (var header in headers) {
                    var cacheField = new CacheField { Name = header };
                    cacheField.SharedItems = new SharedItems { Count = 0U };
                    cacheDef.CacheFields.Append(cacheField);
                }

                cacheDefPart.PivotCacheDefinition = cacheDef;
                cacheDefPart.PivotCacheDefinition.Save();

                var cacheRecordsPart = cacheDefPart.AddNewPart<PivotTableCacheRecordsPart>();
                cacheRecordsPart.PivotCacheRecords = new PivotCacheRecords { Count = 0U };
                cacheRecordsPart.PivotCacheRecords.Save();

                var pivotCaches = workbookPart.Workbook.PivotCaches ?? workbookPart.Workbook.AppendChild(new PivotCaches());
                pivotCaches.Append(new PivotCache {
                    CacheId = cacheId,
                    Id = workbookPart.GetIdOfPart(cacheDefPart)
                });
                // Count attribute is optional; OpenXml SDK does not expose a setter for PivotCaches.Count in all targets.

                var existingNames = _worksheetPart.PivotTableParts
                    .Select(p => p.PivotTableDefinition?.Name?.Value)
                    .Where(n => !string.IsNullOrWhiteSpace(n))
                    .Select(n => n!)
                    .ToList();
                string pivotName = EnsureUniquePivotTableName(name, existingNames);

                var pivotPart = _worksheetPart.AddNewPart<PivotTablePart>();
                string pivotRelId = _worksheetPart.GetIdOfPart(pivotPart);

                var pivotFields = new PivotFields { Count = (uint)headers.Count };
                for (int i = 0; i < headers.Count; i++) {
                    var pivotField = new PivotField { ShowAll = true };
                    if (pageFieldIndices.Contains(i)) pivotField.Axis = PivotTableAxisValues.AxisPage;
                    if (rowFieldIndices.Contains(i)) pivotField.Axis = PivotTableAxisValues.AxisRow;
                    if (columnFieldIndices.Contains(i)) pivotField.Axis = PivotTableAxisValues.AxisColumn;
                    if (dataFieldIndices.Contains(i)) pivotField.DataField = true;
                    pivotFields.Append(pivotField);
                }

                var rowFieldsElement = rowFieldIndices.Count > 0 ? new RowFields { Count = (uint)rowFieldIndices.Count } : null;
                if (rowFieldsElement != null) {
                    foreach (int idx in rowFieldIndices) rowFieldsElement.Append(new Field { Index = idx });
                }

                var columnFieldsElement = columnFieldIndices.Count > 0 ? new ColumnFields { Count = (uint)columnFieldIndices.Count } : null;
                if (columnFieldsElement != null) {
                    foreach (int idx in columnFieldIndices) columnFieldsElement.Append(new Field { Index = idx });
                }

                var pageFieldsElement = pageFieldIndices.Count > 0 ? new PageFields { Count = (uint)pageFieldIndices.Count } : null;
                if (pageFieldsElement != null) {
                    foreach (int idx in pageFieldIndices) pageFieldsElement.Append(new PageField { Field = idx });
                }

                var dataFieldsElement = new DataFields { Count = (uint)dataFieldList.Count };
                foreach (var df in dataFieldList) {
                    int idx = ResolveFieldIndex(df.FieldName, headerIndex, nameof(dataFields));
                    string display = df.DisplayName ?? $"{df.Function} of {headers[idx]}";
                    var dataField = new DataField {
                        Name = display,
                        Field = (uint)idx,
                        Subtotal = df.Function
                    };
                    if (df.NumberFormatId.HasValue) dataField.NumberFormatId = df.NumberFormatId.Value;
                    dataFieldsElement.Append(dataField);
                }

                string pivotRef = BuildPivotLocationReference(destRow, destCol, rowFieldIndices.Count + columnFieldIndices.Count + dataFieldList.Count);

                var pivotDefinition = new PivotTableDefinition {
                    Name = pivotName,
                    CacheId = cacheId,
                    ApplyNumberFormats = true,
                    ApplyBorderFormats = true,
                    ApplyAlignmentFormats = true,
                    ApplyWidthHeightFormats = true,
                    ApplyPatternFormats = true,
                    UseAutoFormatting = true,
                    PreserveFormatting = true,
                    RowGrandTotals = showRowGrandTotals,
                    ColumnGrandTotals = showColumnGrandTotals,
                    MultipleFieldFilters = true,
                    DataCaption = "Values",
                    PivotTableStyleName = string.IsNullOrWhiteSpace(pivotStyleName) ? null : pivotStyleName,
                    Location = new Location {
                        Reference = pivotRef,
                        FirstHeaderRow = 1U,
                        FirstDataRow = 2U,
                        FirstDataColumn = 1U
                    },
                    PivotFields = pivotFields,
                    DataFields = dataFieldsElement
                };

                if (rowFieldsElement != null) pivotDefinition.RowFields = rowFieldsElement;
                if (columnFieldsElement != null) pivotDefinition.ColumnFields = columnFieldsElement;
                if (pageFieldsElement != null) pivotDefinition.PageFields = pageFieldsElement;

                switch (layout) {
                    case ExcelPivotLayout.Compact:
                        pivotDefinition.CompactData = true;
                        pivotDefinition.OutlineData = false;
                        break;
                    case ExcelPivotLayout.Outline:
                        pivotDefinition.CompactData = false;
                        pivotDefinition.OutlineData = true;
                        break;
                    case ExcelPivotLayout.Tabular:
                        pivotDefinition.CompactData = false;
                        pivotDefinition.OutlineData = false;
                        break;
                }

                if (dataOnRows.HasValue) pivotDefinition.DataOnRows = dataOnRows.Value;
                if (showHeaders.HasValue) pivotDefinition.ShowHeaders = showHeaders.Value;
                if (showEmptyRows.HasValue) pivotDefinition.ShowEmptyRow = showEmptyRows.Value;
                if (showEmptyColumns.HasValue) pivotDefinition.ShowEmptyColumn = showEmptyColumns.Value;
                if (showDrill.HasValue) pivotDefinition.ShowDrill = showDrill.Value;

                pivotPart.PivotTableDefinition = pivotDefinition;
                pivotPart.PivotTableDefinition.Save();

                EnsurePivotTablePartsElement(pivotRelId);

                _worksheetPart.Worksheet.Save();
                workbookPart.Workbook.Save();
            });
        }

        private List<string> BuildPivotHeaders(int headerRow, int startColumn, int endColumn) {
            var headers = new List<string>();
            var used = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
            for (int col = startColumn; col <= endColumn; col++) {
                string header = string.Empty;
                if (TryGetCellText(headerRow, col, out var text)) {
                    header = text?.Trim() ?? string.Empty;
                }
                if (string.IsNullOrWhiteSpace(header)) {
                    header = $"Column{col}";
                }
                header = EnsureUniqueName(header, used);
                used.Add(header);
                headers.Add(header);
            }
            return headers;
        }

        private static string EnsureUniqueName(string name, HashSet<string> used) {
            string baseName = string.IsNullOrWhiteSpace(name) ? "Column" : name.Trim();
            if (!used.Contains(baseName)) return baseName;
            int i = 2;
            string candidate;
            do {
                candidate = $"{baseName}_{i}";
                i++;
            } while (used.Contains(candidate));
            return candidate;
        }

        private static string EnsureUniquePivotTableName(string? name, IEnumerable<string> existingNames) {
            string baseName = string.IsNullOrWhiteSpace(name) ? "PivotTable" : name!.Trim();
            var existing = new HashSet<string>(existingNames, StringComparer.OrdinalIgnoreCase);
            if (!existing.Contains(baseName)) return baseName;
            int i = 2;
            string candidate;
            do {
                candidate = $"{baseName}{i}";
                i++;
            } while (existing.Contains(candidate));
            return candidate;
        }

        private static List<int> ResolveFieldIndices(IEnumerable<string>? fields, IDictionary<string, int> headerIndex, string paramName) {
            var indices = new List<int>();
            if (fields == null) return indices;
            foreach (var field in fields) {
                if (string.IsNullOrWhiteSpace(field)) continue;
                int idx = ResolveFieldIndex(field, headerIndex, paramName);
                if (!indices.Contains(idx)) indices.Add(idx);
            }
            return indices;
        }

        private static int ResolveFieldIndex(string field, IDictionary<string, int> headerIndex, string paramName) {
            var key = field.Trim();
            if (!headerIndex.TryGetValue(key, out int idx)) {
                throw new ArgumentException($"Field '{field}' was not found in pivot source headers.", paramName);
            }
            return idx;
        }

        private static string BuildPivotLocationReference(int startRow, int startColumn, int columnCount) {
            int width = Math.Max(1, columnCount);
            int endColumn = startColumn + width - 1;
            int endRow = startRow + 1; // header + at least one data row
            string start = $"{A1.ColumnIndexToLetters(startColumn)}{startRow}";
            string end = $"{A1.ColumnIndexToLetters(endColumn)}{endRow}";
            return $"{start}:{end}";
        }

        private void EnsurePivotTablePartsElement(string relId) {
            var worksheet = _worksheetPart.Worksheet;
            const string mainNs = "http://schemas.openxmlformats.org/spreadsheetml/2006/main";
            var existing = worksheet.ChildElements
                .OfType<OpenXmlUnknownElement>()
                .FirstOrDefault(e => e.LocalName == "pivotTableParts" && e.NamespaceUri == mainNs);

            var relIds = new List<string>();
            if (existing != null) {
                try {
                    var xdoc = XDocument.Parse(existing.OuterXml);
                    XNamespace r = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";
                    XNamespace s = mainNs;
                    foreach (var part in xdoc.Root?.Elements(s + "pivotTablePart") ?? Enumerable.Empty<XElement>()) {
                        var id = part.Attribute(r + "id")?.Value;
                        if (!string.IsNullOrWhiteSpace(id)) relIds.Add(id!);
                    }
                } catch {
                    // If parsing fails, fall back to only the new relationship.
                }
            }

            if (!relIds.Contains(relId, StringComparer.OrdinalIgnoreCase)) {
                relIds.Add(relId);
            }

            const string relNs = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";
            var pivotParts = new OpenXmlUnknownElement("pivotTableParts", mainNs);
            pivotParts.SetAttribute(new OpenXmlAttribute("", "count", "", relIds.Count.ToString()));
            pivotParts.AddNamespaceDeclaration("r", relNs);

            foreach (var id in relIds) {
                var part = new OpenXmlUnknownElement("pivotTablePart", mainNs);
                part.SetAttribute(new OpenXmlAttribute("r", "id", relNs, id));
                pivotParts.Append(part);
            }

            var unknown = pivotParts;
            existing?.Remove();

            var ext = worksheet.Elements<ExtensionList>().FirstOrDefault();
            if (ext != null) {
                worksheet.InsertBefore(unknown, ext);
            } else {
                worksheet.Append(unknown);
            }
        }

        private static uint NextPivotCacheId(WorkbookPart workbookPart) {
            var pivotCaches = workbookPart.Workbook.PivotCaches;
            if (pivotCaches == null) return 1;
            uint max = 0;
            foreach (var cache in pivotCaches.Elements<PivotCache>()) {
                if (cache.CacheId != null && cache.CacheId.Value > max) max = cache.CacheId.Value;
            }
            return max + 1;
        }

        private static Dictionary<uint, PivotCacheDefinition> BuildPivotCacheMap(WorkbookPart workbookPart) {
            var map = new Dictionary<uint, PivotCacheDefinition>();
            var pivotCaches = workbookPart.Workbook.PivotCaches;
            if (pivotCaches == null) return map;
            foreach (var cache in pivotCaches.Elements<PivotCache>()) {
                if (cache.CacheId == null) continue;
                var relId = cache.Id?.Value;
                if (relId == null) continue;
                if (relId.Length == 0) continue;
                if (workbookPart.GetPartById(relId) is PivotTableCacheDefinitionPart part && part.PivotCacheDefinition != null) {
                    map[cache.CacheId.Value] = part.PivotCacheDefinition;
                }
            }
            return map;
        }

        private static List<string> BuildCacheFieldNames(PivotCacheDefinition? cacheDef) {
            var names = new List<string>();
            if (cacheDef?.CacheFields == null) return names;
            var used = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
            int i = 0;
            foreach (var field in cacheDef.CacheFields.Elements<CacheField>()) {
                var name = field.Name?.Value ?? $"Field{i + 1}";
                if (string.IsNullOrWhiteSpace(name)) name = $"Field{i + 1}";
                name = EnsureUniqueName(name, used);
                used.Add(name);
                names.Add(name);
                i++;
            }
            return names;
        }

        private static List<string> ResolveFieldNames(IEnumerable<Field>? fields, IReadOnlyList<string> cacheFields) {
            var list = new List<string>();
            if (fields == null) return list;
            foreach (var field in fields) {
                if (field.Index == null) continue;
                list.Add(ResolveFieldName(field.Index.Value, cacheFields));
            }
            return list;
        }

        private static List<string> ResolvePageFieldNames(IEnumerable<PageField>? fields, IReadOnlyList<string> cacheFields) {
            var list = new List<string>();
            if (fields == null) return list;
            foreach (var field in fields) {
                if (field.Field == null) continue;
                list.Add(ResolveFieldName(field.Field.Value, cacheFields));
            }
            return list;
        }

        private static List<ExcelPivotDataFieldInfo> ResolveDataFields(IEnumerable<DataField>? fields, IReadOnlyList<string> cacheFields) {
            var list = new List<ExcelPivotDataFieldInfo>();
            if (fields == null) return list;
            foreach (var field in fields) {
                int idx = field.Field?.Value is uint u ? (int)u : 0;
                var name = ResolveFieldName(idx, cacheFields);
                var fn = field.Subtotal?.Value ?? DataConsolidateFunctionValues.Sum;
                var display = field.Name?.Value;
                list.Add(new ExcelPivotDataFieldInfo(name, fn, display));
            }
            return list;
        }

        private static string ResolveFieldName(int index, IReadOnlyList<string> cacheFields) {
            if (index >= 0 && index < cacheFields.Count) return cacheFields[index];
            return $"Field{index + 1}";
        }

        private int ResolveSheetIndex(WorkbookPart workbookPart) {
            var sheets = workbookPart.Workbook.Sheets?.OfType<Sheet>().ToList();
            if (sheets == null) return -1;
            for (int i = 0; i < sheets.Count; i++) {
                if (ReferenceEquals(sheets[i], _sheet)) return i;
            }
            return -1;
        }

        private static ExcelPivotLayout ResolveLayout(BooleanValue? compactData, BooleanValue? outlineData) {
            if (outlineData != null && outlineData.Value) return ExcelPivotLayout.Outline;
            if (compactData != null && compactData.Value) return ExcelPivotLayout.Compact;
            return ExcelPivotLayout.Tabular;
        }
    }
}
