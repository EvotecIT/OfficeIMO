using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;

namespace OfficeIMO.Excel {
    public partial class ExcelSheet {
        /// <summary>
        /// Changes the worksheet source of an existing pivot table and requests an application refresh on open.
        /// </summary>
        /// <remarks>
        /// Stale cache records are invalidated instead of being silently retained. When the pivot cache is shared,
        /// every pivot that uses it is affected and the update is rejected unless explicitly allowed.
        /// </remarks>
        /// <param name="pivotTableName">Pivot table name on this worksheet.</param>
        /// <param name="sourceSheet">Worksheet that contains the new source data.</param>
        /// <param name="sourceRange">New A1 source range including its header row.</param>
        /// <param name="options">Optional header and shared-cache safety settings.</param>
        /// <returns>Details about the updated cache and affected pivot tables.</returns>
        public ExcelPivotSourceUpdateResult UpdatePivotTableSource(
            string pivotTableName,
            ExcelSheet sourceSheet,
            string sourceRange,
            ExcelPivotSourceUpdateOptions? options = null) {
            if (string.IsNullOrWhiteSpace(pivotTableName)) {
                throw new ArgumentNullException(nameof(pivotTableName));
            }

            if (sourceSheet == null) {
                throw new ArgumentNullException(nameof(sourceSheet));
            }

            if (!ReferenceEquals(_excelDocument, sourceSheet._excelDocument)) {
                throw new ArgumentException("The pivot source worksheet must belong to the same workbook.", nameof(sourceSheet));
            }

            if (string.IsNullOrWhiteSpace(sourceRange)) {
                throw new ArgumentNullException(nameof(sourceRange));
            }

            string sourceRangeToParse = sourceRange.Trim().Replace("$", string.Empty);
            if (!A1.TryParseRange(sourceRangeToParse, out int firstRow, out int firstColumn, out int lastRow, out int lastColumn)) {
                throw new ArgumentException($"Invalid A1 range '{sourceRange}'.", nameof(sourceRange));
            }

            if (lastRow <= firstRow) {
                throw new ArgumentException("A pivot source range must contain a header row and at least one data row.", nameof(sourceRange));
            }

            string normalizedSourceRange = $"{A1.CellReference(firstRow, firstColumn)}:{A1.CellReference(lastRow, lastColumn)}";
            ExcelPivotSourceUpdateResult? result = null;
            WriteLockWorksheetPreparationOnly(() => {
                sourceSheet.MaterializePendingDirectCellValues();

                PivotTablePart pivotTablePart = _worksheetPart.PivotTableParts
                    .FirstOrDefault(part => string.Equals(
                        part.PivotTableDefinition?.Name?.Value,
                        pivotTableName.Trim(),
                        StringComparison.OrdinalIgnoreCase))
                    ?? throw new ArgumentException($"Pivot table '{pivotTableName}' was not found on worksheet '{Name}'.", nameof(pivotTableName));
                PivotTableDefinition pivotDefinition = pivotTablePart.PivotTableDefinition
                    ?? throw new InvalidOperationException($"Pivot table '{pivotTableName}' has no definition.");
                PivotTableCacheDefinitionPart cachePart = pivotTablePart.PivotTableCacheDefinitionPart
                    ?? throw new InvalidOperationException($"Pivot table '{pivotTableName}' has no cache definition part.");
                PivotCacheDefinition cacheDefinition = cachePart.PivotCacheDefinition
                    ?? throw new InvalidOperationException($"Pivot table '{pivotTableName}' has no cache definition.");

                if (cacheDefinition.CacheSource?.Type?.Value != SourceValues.Worksheet) {
                    throw new NotSupportedException("Only worksheet-backed pivot cache sources can be updated.");
                }

                List<string> newHeaders = sourceSheet.BuildPivotHeaders(firstRow, firstColumn, lastColumn);
                List<string> existingHeaders = cacheDefinition.CacheFields?
                    .Elements<CacheField>()
                    .Where(field => field.DatabaseField?.Value != false)
                    .Select(field => field.Name?.Value ?? string.Empty)
                    .ToList() ?? new List<string>();
                if (options?.RequireMatchingHeaders != false
                    && !existingHeaders.SequenceEqual(newHeaders, StringComparer.OrdinalIgnoreCase)) {
                    throw new InvalidOperationException(
                        $"The new pivot source headers ({string.Join(", ", newHeaders)}) do not match the existing cache fields ({string.Join(", ", existingHeaders)}).");
                }

                uint cacheId = pivotDefinition.CacheId?.Value
                    ?? throw new InvalidOperationException($"Pivot table '{pivotTableName}' has no cache identifier.");
                List<string> affectedPivotTables = WorkbookPartRoot.WorksheetParts
                    .SelectMany(part => part.PivotTableParts)
                    .Select(part => part.PivotTableDefinition)
                    .Where(definition => definition?.CacheId?.Value == cacheId)
                    .Select(definition => definition!.Name?.Value ?? string.Empty)
                    .Where(name => name.Length > 0)
                    .Distinct(StringComparer.OrdinalIgnoreCase)
                    .OrderBy(name => name, StringComparer.OrdinalIgnoreCase)
                    .ToList();
                if (affectedPivotTables.Count > 1 && options?.AllowSharedCacheUpdate != true) {
                    throw new InvalidOperationException(
                        $"Pivot cache {cacheId} is shared by {affectedPivotTables.Count} pivot tables ({string.Join(", ", affectedPivotTables)}). " +
                        "Set AllowSharedCacheUpdate to true to update all of them.");
                }

                WorksheetSource worksheetSource = cacheDefinition.CacheSource.WorksheetSource
                    ?? cacheDefinition.CacheSource.AppendChild(new WorksheetSource());
                worksheetSource.Sheet = sourceSheet.Name;
                worksheetSource.Reference = normalizedSourceRange;
                cacheDefinition.RecordCount = (uint)(lastRow - firstRow);
                cacheDefinition.RefreshOnLoad = true;
                cacheDefinition.SaveData = false;

                uint invalidatedRecordCount = 0;
                PivotTableCacheRecordsPart? recordsPart = cachePart.GetPartsOfType<PivotTableCacheRecordsPart>().FirstOrDefault();
                if (recordsPart != null) {
                    invalidatedRecordCount = recordsPart.PivotCacheRecords?.Count?.Value ?? 0U;
                    recordsPart.PivotCacheRecords = new PivotCacheRecords { Count = 0U };
                    ExcelDocument.MarkPivotCacheRecordsPartAsModelWritten(recordsPart);
                    recordsPart.PivotCacheRecords.Save();
                }

                cacheDefinition.Save();
                result = new ExcelPivotSourceUpdateResult(
                    pivotDefinition.Name?.Value ?? pivotTableName.Trim(),
                    cacheId,
                    sourceSheet.Name,
                    normalizedSourceRange,
                    affectedPivotTables,
                    invalidatedRecordCount);
            });

            return result!;
        }
    }
}
