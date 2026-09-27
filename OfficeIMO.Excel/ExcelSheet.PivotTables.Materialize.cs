using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using DocumentFormat.OpenXml.Validation;
using System.Threading;

namespace OfficeIMO.Excel {
    public partial class ExcelSheet {
        /// <summary>
        /// Generates a saved pivot view and its source cache from current worksheet values.
        /// Supports up to 256 measures and source fields across ungrouped axes, with deterministic
        /// first-seen key order and optional grand totals. Source formulas use their saved typed caches.
        /// </summary>
        /// <param name="pivotTableName">Pivot definition on this worksheet.</param>
        /// <param name="options">Source, measure-input, output and rollback limits. Each cell/work budget uses MaximumAffectedCells, capped at one million.</param>
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
            var measures = definition.DataFields?.Elements<DataField>().Take(257).ToArray() ?? Array.Empty<DataField>();
            if (fields.Length != fieldCount || pivotFields.Length != fieldCount || measures.Length == 0 || measures.Length > 256
                || fields.Any(f => f.FieldGroup != null || f.Formula != null || f.DatabaseField?.Value == false)
                || definition.PageFields?.ChildElements.Count > 0 || definition.PivotFilters?.ChildElements.Count > 0
                || pivotFields.Any(f => f.Items?.Elements<Item>().Any(i => i.Hidden?.Value == true) == true)
                || measures.Any(m => m.ShowDataAs?.Value is ShowDataAsValues mode && mode != ShowDataAsValues.Normal))
                throw new NotSupportedException("Materialization requires ordinary measures, no calculated/grouped fields, and no filters or page fields.");
            if (!sourceSheet.BuildPivotHeaders(r1, c1, c2).SequenceEqual(fields.Select(f => f.Name?.Value ?? ""), StringComparer.OrdinalIgnoreCase))
                throw new InvalidOperationException("The source headers no longer match the pivot cache fields.");
            var rowAxis = ResolveMaterializationAxis(definition.RowFields);
            var columnAxis = ResolveMaterializationAxis(definition.ColumnFields);
            int rowField = rowAxis.RealField;
            int columnField = columnAxis.RealField;
            if (measures.Any(m => m.Field == null || m.Field.Value >= fieldCount))
                throw new NotSupportedException("The pivot measure does not match the source fields.");
            if (measures.Length > 1 && (measures.Any(m => string.IsNullOrWhiteSpace(m.Name?.Value))
                || measures.Select(m => m.Name!.Value!).Distinct(StringComparer.OrdinalIgnoreCase).Count() != measures.Length))
                throw new InvalidOperationException("Multiple pivot measures require unique non-empty captions.");
            int valuesAxes = (rowAxis.HasValues ? 1 : 0) + (columnAxis.HasValues ? 1 : 0);
            if ((measures.Length > 1 && valuesAxes != 1) || (measures.Length == 1 && valuesAxes != 0))
                throw new NotSupportedException("Multiple measures require exactly one Values axis; single measures use real axes only.");
            var realFields = rowAxis.RealFields.Concat(columnAxis.RealFields).ToArray();
            if (realFields.Any(field => field >= fieldCount) || realFields.Distinct().Count() != realFields.Length)
                throw new NotSupportedException("The pivot axes or measure do not match the source fields.");
            bool rowTotal = rowField >= 0 && definition.ColumnGrandTotals?.Value != false;
            bool columnTotal = columnField >= 0 && definition.RowGrandTotals?.Value != false;
            long visits = (long)(r2 - r1) * measures.Length * MaterializedInputLevels(rowAxis, pivotFields)
                * MaterializedInputLevels(columnAxis, pivotFields);
            if (visits > limit) throw new InvalidOperationException("The pivot measure input visits, including intermediate subtotals, exceed the materialization budget.");
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
            foreach (int field in realFields)
                if (maps[field].Items.Any(v => v.Kind == PivotFieldValueKind.Date || v.Kind == PivotFieldValueKind.Error))
                    throw new NotSupportedException("Date and error pivot keys are not yet supported by this materializer.");
            // Lookup indexes both the saved field items and their shared keys. Keep every
            // possible criterion combination usable, rather than accepting an unreadable view.
            foreach (var axis in new[] { rowAxis, columnAxis }) {
                long indexedItems = axis.RealFields.Sum(field => 2L * maps[field].Items.Count
                    + (MaterializedAutomaticSubtotal(pivotFields[field]) ? 1 : 0));
                if (indexedItems > limit)
                    throw new InvalidOperationException("The pivot criteria index exceeds the materialization budget.");
            }
            var rows = BuildMaterializedHierarchyAxis(sourceSheet, r1, r2, c1, maps, rowAxis, pivotFields, measures.Length, rowTotal, token);
            var columns = BuildMaterializedHierarchyAxis(sourceSheet, r1, r2, c1, maps, columnAxis, pivotFields, measures.Length, columnTotal, token);
            int dataRow = columnField >= 0 ? columnAxis.Fields.Length + 1 : 1;
            int dataColumn = rowAxis.Fields.Length > 0 ? rowAxis.Fields.Length : measures.Length == 1 && columnField >= 0 ? 1 : 0;
            int height = dataRow + rows.Entries.Count;
            int width = dataColumn + columns.Entries.Count;
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
            var aggregates = AggregateMaterializedHierarchy(sourceSheet, r1, c1, r2, rows, columns, measures, limit, token);
            FillMaterializedHierarchy(plan, maps, rows, columns, measures, dataRow, dataColumn, aggregates, token);
            var cacheFields = plan.Cache.CacheFields!.Elements<CacheField>().ToArray();
            for (int field = 0; field < fieldCount; field++) cacheFields[field].SharedItems = BuildSharedItems(maps[field], null);
            plan.Records = sourceSheet.BuildPivotCacheRecords(fieldCount, r1 + 1, r2, c1, grouping, maps, collect,
                Array.Empty<GeneratedPivotGroupingField>(), Array.Empty<PivotFieldValues>(), 0);
            plan.Cache.RecordCount = (uint)plan.SourceRecords;
            plan.Cache.SaveData = true;
            plan.Cache.RefreshOnLoad = false;
            return plan;
        }

        private sealed class PivotMaterializationAxis {
            internal int[] Fields = Array.Empty<int>();
            internal int[] RealFields = Array.Empty<int>();
            internal int RealField => RealFields.Length == 0 ? -1 : RealFields[0];
            internal bool HasValues => Fields.Contains(-2);
        }

        private static PivotMaterializationAxis ResolveMaterializationAxis(OpenXmlCompositeElement? axis) {
            if (axis == null || axis.ChildElements.Count == 0) return new PivotMaterializationAxis();
            if (axis.ChildElements.Count > 257 || axis.ChildElements.Any(f => f is not Field))
                throw new NotSupportedException("Materialization supports at most 256 source fields and one Values field across its axes.");
            var indices = axis.Elements<Field>().Select(f => f.Index?.Value ?? int.MinValue).ToArray();
            if (indices.Count(f => f == -2) > 1 || indices.Any(f => f < 0 && f != -2))
                throw new NotSupportedException("Materialization requires ordinary source fields and at most one Values field per axis.");
            return new PivotMaterializationAxis { Fields = indices, RealFields = indices.Where(f => f >= 0).ToArray() };
        }
    }
}
