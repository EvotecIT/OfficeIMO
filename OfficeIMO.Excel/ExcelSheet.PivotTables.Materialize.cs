using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using DocumentFormat.OpenXml.Validation;
using System.Threading;

namespace OfficeIMO.Excel {
    public partial class ExcelSheet {
        /// <summary>
        /// Generates a saved pivot view and its source cache from current worksheet values.
        /// Supports one measure and at most one ungrouped field on each axis, with deterministic
        /// first-seen key order and optional grand totals. Source formulas use their saved typed caches.
        /// </summary>
        /// <param name="pivotTableName">Pivot definition on this worksheet.</param>
        /// <param name="options">Source, output and rollback limits. Source and output each use MaximumAffectedCells, capped at one million.</param>
        /// <param name="cancellationToken">Cancels preparation or rolls back an interrupted write.</param>
        /// <exception cref="NotSupportedException">The pivot uses an unqualified grouping, filter, calculated field, shared cache or measure layout.</exception>
        /// <exception cref="InvalidOperationException">A budget or destination collision prevents generation.</exception>
        public ExcelPivotMaterializationResult MaterializePivotTable(string pivotTableName, ExcelMutationPlanOptions? options = null, CancellationToken cancellationToken = default) {
            if (string.IsNullOrWhiteSpace(pivotTableName)) throw new ArgumentException("A pivot table name is required.", nameof(pivotTableName));
            var effective = (options ?? new ExcelMutationPlanOptions()).CloneAndValidate();
            ExcelPivotMaterializationResult? result = null;
            Batch(_ => {
                cancellationToken.ThrowIfCancellationRequested();
                var plan = PreparePivotMaterialization(pivotTableName, effective, cancellationToken);
                var mutation = ApplyTransactionalMutation(token => {
                    var recordPart = plan.CachePart.PivotTableCacheRecordsPart ?? plan.CachePart.AddNewPart<PivotTableCacheRecordsPart>();
                    plan.Cache.Id = plan.CachePart.GetIdOfPart(recordPart);
                    recordPart.PivotCacheRecords = plan.Records;
                    ExcelDocument.MarkPivotCacheRecordsPartAsModelWritten(recordPart);
                    plan.CachePart.PivotCacheDefinition = plan.Cache;
                    plan.Part.PivotTableDefinition = plan.Definition;
                    ClearExistingCellFieldsInRange((plan.Top, plan.Left, plan.AffectedBottom, plan.AffectedRight), ExcelClearOptions.Values);
                    ClearHeaderCache();
                    for (int row = plan.Top; row <= plan.AffectedBottom; row++) {
                        token.ThrowIfCancellationRequested();
                        for (int column = plan.Left; column <= plan.AffectedRight; column++) {
                            if (row > plan.Bottom || column > plan.Right) {
                                continue;
                            }
                            var value = plan.Values[row - plan.Top, column - plan.Left];
                            if (value?.Kind == ExcelCellDataKind.Error) CellError(row, column, (string)value.Value!);
                            else if (value != null) CellValue(row, column, value.Value);
                        }
                    }
                    var validator = new OpenXmlValidator(FileFormatVersions.Microsoft365) { MaxNumberOfErrors = 1 };
                    foreach (var changed in new OpenXmlPart[] { _worksheetPart, plan.Part, plan.CachePart, recordPart }) {
                        token.ThrowIfCancellationRequested();
                        var error = validator.Validate(changed, token).FirstOrDefault();
                        if (error != null) throw new InvalidOperationException("Generated pivot metadata is invalid: " + error.Description);
                    }
                    return (plan.AffectedBottom - plan.Top + 1) * (plan.AffectedRight - plan.Left + 1);
                }, effective, cancellationToken);
                result = new ExcelPivotMaterializationResult(plan.Definition.Name!.Value!, plan.Definition.Location!.Reference!.Value!, plan.SourceRecords, mutation);
            });
            return result!;
        }

        private sealed class PivotMaterializationPlan {
            internal PivotTablePart Part = null!;
            internal PivotTableCacheDefinitionPart CachePart = null!;
            internal PivotTableDefinition Definition = null!;
            internal PivotCacheDefinition Cache = null!;
            internal PivotCacheRecords Records = null!;
            internal ExcelCellData?[,] Values = null!;
            internal int Top, Left, Bottom, Right, OldBottom, OldRight, AffectedBottom, AffectedRight, SourceRecords;
        }

        private PivotMaterializationPlan PreparePivotMaterialization(string name, ExcelMutationPlanOptions options, CancellationToken token) {
            var part = _worksheetPart.PivotTableParts.FirstOrDefault(p => string.Equals(p.PivotTableDefinition?.Name?.Value, name, StringComparison.OrdinalIgnoreCase))
                ?? throw new ArgumentException("The pivot table was not found on this worksheet.", nameof(name));
            var definition = part.PivotTableDefinition ?? throw new InvalidOperationException("The pivot definition is missing.");
            var cachePart = part.PivotTableCacheDefinitionPart ?? throw new InvalidOperationException("The pivot cache is missing.");
            var cache = cachePart.PivotCacheDefinition ?? throw new InvalidOperationException("The pivot cache definition is missing.");
            var source = cache.CacheSource?.WorksheetSource;
            if (cache.CacheSource?.Type?.Value != SourceValues.Worksheet || source == null || source.Id != null || source.Name != null
                || string.IsNullOrEmpty(source.Sheet?.Value) || !A1.TryParseRange(source.Reference?.Value ?? "", out int r1, out int c1, out int r2, out int c2))
                throw new NotSupportedException("Materialization requires a local worksheet source range.");
            var sourceSheet = _excelDocument.GetSheetForLockedOperation(source.Sheet!.Value!);
            sourceSheet.MaterializePendingDirectCellValues();
            int fieldCount = c2 - c1 + 1;
            int limit = Math.Min(1_000_000, options.MaximumAffectedCells);
            if (r2 <= r1 || fieldCount > 256 || (long)(r2 - r1) * fieldCount > limit)
                throw new InvalidOperationException("The pivot source exceeds the materialization budget or has no data records.");
            var fields = cache.CacheFields?.Elements<CacheField>().Take(257).ToArray() ?? Array.Empty<CacheField>();
            var pivotFields = definition.PivotFields?.Elements<PivotField>().Take(257).ToArray() ?? Array.Empty<PivotField>();
            var measures = definition.DataFields?.Elements<DataField>().Take(2).ToArray() ?? Array.Empty<DataField>();
            if (fields.Length != fieldCount || pivotFields.Length != fieldCount || measures.Length != 1
                || fields.Any(f => f.FieldGroup != null || f.Formula != null || f.DatabaseField?.Value == false)
                || definition.PageFields?.ChildElements.Count > 0 || definition.PivotFilters?.ChildElements.Count > 0
                || pivotFields.Any(f => f.Items?.Elements<Item>().Any(i => i.Hidden?.Value == true) == true)
                || (measures[0].ShowDataAs?.Value is ShowDataAsValues mode && mode != ShowDataAsValues.Normal))
                throw new NotSupportedException("Materialization requires one ordinary measure, no calculated/grouped fields, and no filters or page fields.");
            if (!sourceSheet.BuildPivotHeaders(r1, c1, c2).SequenceEqual(fields.Select(f => f.Name?.Value ?? ""), StringComparer.OrdinalIgnoreCase))
                throw new InvalidOperationException("The source headers no longer match the pivot cache fields.");
            int rowField = ResolveMaterializationAxis(definition.RowFields);
            int columnField = ResolveMaterializationAxis(definition.ColumnFields);
            if (measures[0].Field == null || measures[0].Field!.Value >= fieldCount)
                throw new NotSupportedException("The pivot measure does not match the source fields.");
            int measureField = (int)measures[0].Field!.Value;
            if (rowField >= fieldCount || columnField >= fieldCount || measureField < 0 || measureField >= fieldCount || (rowField >= 0 && rowField == columnField))
                throw new NotSupportedException("The pivot axes or measure do not match the source fields.");
            if (WorkbookPartRoot.WorksheetParts.SelectMany(p => p.PivotTableParts).Count(p => ReferenceEquals(p.PivotTableCacheDefinitionPart, cachePart)) != 1)
                throw new NotSupportedException("Materialization of a shared pivot cache requires a coordinated refresh of all its views.");
            if (_excelDocument.GetWorkbookSlicerCaches().Concat(_excelDocument.GetWorkbookTimelineCaches())
                .Any(c => string.IsNullOrEmpty(c.PivotTableName) || string.Equals(c.PivotTableName, definition.Name?.Value, StringComparison.OrdinalIgnoreCase)))
                throw new NotSupportedException("Materialization of pivot interaction caches requires coordinated cache updates.");
            if (!A1.TryParseRange(definition.Location?.Reference?.Value ?? "", out int top, out int left, out int oldBottom, out int oldRight))
                throw new InvalidOperationException("The pivot output location is invalid.");
            var grouping = new Dictionary<int, ExcelPivotGrouping>();
            // Every field is saved in the source cache, including fields absent from the axes.
            for (int row = r1 + 1; row <= r2; row++) {
                token.ThrowIfCancellationRequested();
                for (int column = c1; column <= c2; column++) {
                    var cell = sourceSheet.TryGetExistingCell(row, column);
                    if (cell?.CellFormula != null && cell.CellValue == null)
                        throw new InvalidOperationException("Pivot source formulas must have saved cached values before materialization.");
                }
            }
            var collect = Enumerable.Repeat(true, fieldCount).ToArray();
            var maps = sourceSheet.BuildPivotFieldValueMap(fieldCount, r1 + 1, r2, c1, grouping, collect);
            if (maps.Any(m => m.Items.Count > 100_000)) throw new InvalidOperationException("A pivot cache field exceeds 100,000 distinct items.");
            // Date and error item keys need independent axis-layout qualification before generation.
            foreach (int field in new[] { rowField, columnField }.Where(f => f >= 0))
                if (maps[field].Items.Any(v => v.Kind == PivotFieldValueKind.Date || v.Kind == PivotFieldValueKind.Error))
                    throw new NotSupportedException("Date and error pivot keys are not yet supported by this materializer.");
            int rowKeys = rowField < 0 ? 1 : maps[rowField].Items.Count;
            int columnKeys = columnField < 0 ? 1 : maps[columnField].Items.Count;
            bool rowTotal = rowField >= 0 && definition.ColumnGrandTotals?.Value != false;
            bool columnTotal = columnField >= 0 && definition.RowGrandTotals?.Value != false;
            int dataRow = columnField >= 0 ? 2 : 1;
            int dataColumn = rowField >= 0 || columnField >= 0 ? 1 : 0;
            int height = dataRow + rowKeys + (rowTotal ? 1 : 0);
            int width = dataColumn + columnKeys + (columnTotal ? 1 : 0);
            bool hasOldView = definition.RowItems?.ChildElements.Count > 0 && definition.ColumnItems?.ChildElements.Count > 0;
            if (!hasOldView) { oldBottom = top - 1; oldRight = left - 1; }
            long bottom = (long)top + height - 1, right = (long)left + width - 1;
            long affectedBottom = Math.Max(bottom, oldBottom), affectedRight = Math.Max(right, oldRight);
            if (bottom > 1_048_576 || right > 16_384 || (long)height * width > limit
                || (affectedBottom - top + 1) * (affectedRight - left + 1) > limit)
                throw new InvalidOperationException("The pivot output exceeds the worksheet or materialization budget.");
            ValidatePivotMaterializationDestination(part, sourceSheet, r1, c1, r2, c2, top, left, (int)affectedBottom, (int)affectedRight, oldBottom, oldRight, token);
            var plan = new PivotMaterializationPlan { Part = part, CachePart = cachePart, Definition = (PivotTableDefinition)definition.CloneNode(true),
                Cache = (PivotCacheDefinition)cache.CloneNode(true), Top = top, Left = left, Bottom = (int)bottom, Right = (int)right,
                OldBottom = oldBottom, OldRight = oldRight, AffectedBottom = (int)affectedBottom, AffectedRight = (int)affectedRight, SourceRecords = r2 - r1 };
            FillMaterializedPivot(plan, sourceSheet, r1, c1, r2, maps, rowField, columnField, measureField,
                (measures[0].Subtotal?.Value ?? DataConsolidateFunctionValues.Sum).ToOfficeEnum(), rowTotal, columnTotal, dataRow, dataColumn, token);
            var cacheFields = plan.Cache.CacheFields!.Elements<CacheField>().ToArray();
            for (int field = 0; field < fieldCount; field++) cacheFields[field].SharedItems = BuildSharedItems(maps[field], null);
            plan.Records = sourceSheet.BuildPivotCacheRecords(fieldCount, r1 + 1, r2, c1, grouping, maps, collect,
                Array.Empty<GeneratedPivotGroupingField>(), Array.Empty<PivotFieldValues>(), 0);
            plan.Cache.RecordCount = (uint)plan.SourceRecords;
            plan.Cache.SaveData = true;
            plan.Cache.RefreshOnLoad = false;
            return plan;
        }

        private static int ResolveMaterializationAxis(OpenXmlCompositeElement? axis) {
            if (axis == null || axis.ChildElements.Count == 0) return -1;
            if (axis.ChildElements.Count != 1 || axis.FirstChild is not Field field || field.Index == null || field.Index.Value < 0)
                throw new NotSupportedException("Materialization supports at most one real field on each axis.");
            return field.Index.Value;
        }
    }
}
