using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;

namespace OfficeIMO.Excel {
    public partial class ExcelSheet {
        private void RemapShiftedPivotSources(int firstAffectedRow, int rowDelta, int? lastDeletedRow) {
            foreach (PivotTableCacheDefinitionPart cachePart in GetWorkbookPivotCacheDefinitionParts()) {
                WorksheetSource? source = cachePart.PivotCacheDefinition?.CacheSource?.WorksheetSource;
                if (source != null
                    && string.IsNullOrWhiteSpace(source.Id?.Value)
                    && string.Equals(source.Sheet?.Value, Name, StringComparison.OrdinalIgnoreCase)
                    && source.Reference?.Value is string sourceReference
                    && TryRemapShiftedReferenceListRows(
                        sourceReference,
                        firstAffectedRow,
                        rowDelta,
                        lastDeletedRow,
                        out List<string> remappedSources)) {
                    source.Reference = remappedSources.Count == 0 ? "#REF!" : remappedSources[0];
                    uint? recordCount = null;
                    if (remappedSources.Count > 0
                        && A1.TryParseRange(
                            remappedSources[0],
                            out int firstRow,
                            out _,
                            out int lastRow,
                            out _)) {
                        recordCount = (uint)Math.Max(0, lastRow - firstRow);
                    }

                    InvalidatePivotCacheAfterStructuralEdit(cachePart, recordCount);
                    continue;
                }

                if (string.IsNullOrWhiteSpace(source?.Id?.Value)
                    && source?.Name?.Value is string sourceName
                    && IsNamedPivotSourceAffected(sourceName, firstAffectedRow)) {
                    InvalidatePivotCacheAfterStructuralEdit(cachePart, recordCount: null);
                    continue;
                }

                bool consolidationChanged = false;
                foreach (RangeSet rangeSet in cachePart.PivotCacheDefinition?.CacheSource?
                    .Consolidation?.RangeSets?.Elements<RangeSet>() ?? Enumerable.Empty<RangeSet>()) {
                    if (!string.IsNullOrWhiteSpace(rangeSet.Id?.Value)
                        || !string.Equals(rangeSet.Sheet?.Value, Name, StringComparison.OrdinalIgnoreCase)
                        || rangeSet.Reference?.Value is not string reference
                        || !TryRemapShiftedReferenceListRows(
                            reference,
                            firstAffectedRow,
                            rowDelta,
                            lastDeletedRow,
                            out List<string> remappedReferences)) {
                        continue;
                    }

                    rangeSet.Reference = remappedReferences.Count == 0
                        ? "#REF!"
                        : remappedReferences[0];
                    consolidationChanged = true;
                }

                if (consolidationChanged) {
                    InvalidatePivotCacheAfterStructuralEdit(cachePart, recordCount: null);
                }
            }
        }

        private IEnumerable<PivotTableCacheDefinitionPart> GetWorkbookPivotCacheDefinitionParts() {
            return WorkbookPartRoot.WorksheetParts
                .SelectMany(worksheetPart => worksheetPart.PivotTableParts)
                .Select(pivotPart => pivotPart.PivotTableCacheDefinitionPart)
                .Where(cachePart => cachePart != null)
                .Cast<PivotTableCacheDefinitionPart>()
                .Concat(WorkbookPartRoot.PivotTableCacheDefinitionParts)
                .Distinct();
        }

        private bool IsNamedPivotSourceAffected(string sourceName, int firstAffectedRow) {
            string? tableRange = GetTableRange(sourceName);
            if (tableRange != null
                && A1.TryParseRange(tableRange.Replace("$", string.Empty), out _, out _, out int tableLastRow, out _)) {
                return tableLastRow >= firstAffectedRow;
            }

            if (TryResolveDefinedNameRange(
                    sourceName,
                    currentRow: null,
                    out ExcelSheet sourceSheet,
                    out _,
                    out _,
                    out int namedLastRow,
                    out _)) {
                bool isMutatedSheet = ReferenceEquals(sourceSheet._worksheetPart, _worksheetPart)
                    || string.Equals(sourceSheet.Name, Name, StringComparison.OrdinalIgnoreCase);
                return isMutatedSheet && namedLastRow >= firstAffectedRow;
            }

            return WorkbookRoot.DefinedNames?.Elements<DefinedName>()
                .Any(name => string.Equals(name.Name?.Value, sourceName, StringComparison.OrdinalIgnoreCase)) == true;
        }

        internal static void InvalidatePivotCacheAfterStructuralEdit(
            PivotTableCacheDefinitionPart cachePart,
            uint? recordCount) {
            PivotCacheDefinition? definition = cachePart.PivotCacheDefinition;
            if (definition == null) {
                return;
            }

            definition.RefreshOnLoad = true;
            definition.SaveData = false;
            if (recordCount.HasValue) {
                definition.RecordCount = recordCount.Value;
            } else {
                definition.RecordCount = 0U;
            }

            PivotTableCacheRecordsPart? recordsPart = cachePart.PivotTableCacheRecordsPart;
            if (recordsPart != null) {
                recordsPart.PivotCacheRecords = new PivotCacheRecords { Count = 0U };
                ExcelDocument.MarkPivotCacheRecordsPartAsModelWritten(recordsPart);
                recordsPart.PivotCacheRecords.Save();
            }

            definition.Save();
        }
    }
}
