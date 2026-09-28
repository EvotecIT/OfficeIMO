using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using DocumentFormat.OpenXml.Validation;
using System.Threading;

namespace OfficeIMO.Excel {
    public partial class ExcelSheet {
        /// <summary>
        /// Generates saved pivot views and their shared source cache from current worksheet values.
        /// Supports up to 256 measures and fields across ordinary, numeric-range, derived date and manual text group axes, selected page items,
        /// hidden row, column, or page items, deterministic first-seen key order, and optional grand totals.
        /// Source formulas use their saved typed caches.
        /// </summary>
        /// <param name="pivotTableName">Pivot definition on this worksheet.</param>
        /// <param name="options">Source, measure-input, output and rollback limits. Each cell/work budget uses MaximumAffectedCells, capped at one million.</param>
        /// <param name="cancellationToken">Cancels preparation or rolls back an interrupted write.</param>
        /// <param name="referenceDate">Local calendar date used for relative date filters. Defaults to today's date, captured once for all shared-cache views.</param>
        /// <exception cref="NotSupportedException">A cache view uses an unqualified grouping, filter, calculated field, measure layout, or incompatible cache representation.</exception>
        /// <exception cref="InvalidOperationException">A budget or destination collision prevents generation.</exception>
        public ExcelPivotMaterializationResult MaterializePivotTable(string pivotTableName, ExcelMutationPlanOptions? options = null,
            CancellationToken cancellationToken = default, DateTime? referenceDate = null) {
            if (string.IsNullOrWhiteSpace(pivotTableName)) throw new ArgumentException("A pivot table name is required.", nameof(pivotTableName));
            var effective = (options ?? new ExcelMutationPlanOptions()).CloneAndValidate();
            DateTime pivotReferenceDate = DateTime.SpecifyKind((referenceDate ?? DateTime.Today).Date, DateTimeKind.Unspecified);
            ExcelPivotMaterializationResult? result = null;
            Batch(_ => {
                cancellationToken.ThrowIfCancellationRequested();
                PivotTablePart requestedPart = _worksheetPart.PivotTableParts.FirstOrDefault(p =>
                    string.Equals(p.PivotTableDefinition?.Name?.Value, pivotTableName, StringComparison.OrdinalIgnoreCase))
                    ?? throw new ArgumentException("The pivot table was not found on this worksheet.", nameof(pivotTableName));
                PivotTableCacheDefinitionPart cachePart = requestedPart.PivotTableCacheDefinitionPart
                    ?? throw new InvalidOperationException("The pivot cache is missing.");
                var views = _excelDocument.GetSheetsForLockedOperation()
                    .SelectMany(sheet => sheet._worksheetPart.PivotTableParts
                        .Where(part => ReferenceEquals(part.PivotTableCacheDefinitionPart, cachePart))
                        .Select(part => (Sheet: sheet, Part: part)))
                    .OrderBy(view => ReferenceEquals(view.Part, requestedPart) ? 0 : 1)
                    .ThenBy(view => view.Sheet.Name, StringComparer.OrdinalIgnoreCase)
                    .ThenBy(view => view.Part.PivotTableDefinition?.Name?.Value, StringComparer.OrdinalIgnoreCase)
                    .ToArray();
                var source = cachePart.PivotCacheDefinition?.CacheSource?.WorksheetSource;
                if (views.Length > 1 && A1.TryParseRange(source?.Reference?.Value ?? "",
                        out int firstRow, out int firstColumn, out int lastRow, out int lastColumn)
                    && (long)views.Length * (lastRow - firstRow) * (lastColumn - firstColumn + 1)
                        > Math.Min(1_000_000, effective.MaximumAffectedCells))
                    throw new InvalidOperationException("The combined pivot source work exceeds the materialization budget.");
                int remainingInputVisits = Math.Min(1_000_000, effective.MaximumAffectedCells);
                int remainingOutputCells = remainingInputVisits;
                int[] dateFilterFields = views.SelectMany(view =>
                        view.Part.PivotTableDefinition?.PivotFilters?.Elements<PivotFilter>()
                            ?? Enumerable.Empty<PivotFilter>())
                    .Where(filter => IsMaterializedPivotFixedDateFilter(filter) || IsMaterializedPivotCalendarFilter(filter)
                        || IsMaterializedPivotRelativeDateFilter(filter))
                    .Select(filter => filter.Field?.Value).Where(field => field.HasValue)
                    .Select(field => (int)field!.Value).Distinct().ToArray();
                var prepared = new List<(ExcelSheet Sheet, PivotMaterializationPlan Plan)>();
                foreach (var view in views) {
                    PivotMaterializationPlan current = view.Sheet.PreparePivotMaterialization(
                        view.Part.PivotTableDefinition?.Name?.Value
                            ?? throw new InvalidOperationException("A shared pivot view has no name."),
                        effective, remainingInputVisits, remainingOutputCells, dateFilterFields,
                        pivotReferenceDate, cancellationToken);
                    prepared.Add((view.Sheet, current));
                    remainingInputVisits -= (int)current.MeasureInputVisits;
                    remainingOutputCells -= (current.AffectedBottom - current.Top + 1)
                        * (current.AffectedRight - current.Left + 1);
                }
                var plans = prepared.ToArray();
                PivotMaterializationPlan plan = plans.Single(item => ReferenceEquals(item.Plan.Part, requestedPart)).Plan;
                int affectedCells = ValidateCoordinatedPivotPlans(plans, effective);
                ExcelMutationResult mutation;
                try {
                    mutation = ApplyTransactionalMutation(token => {
                        var recordPart = cachePart.PivotTableCacheRecordsPart ?? cachePart.AddNewPart<PivotTableCacheRecordsPart>();
                        plan.Cache.Id = cachePart.GetIdOfPart(recordPart);
                        recordPart.PivotCacheRecords = plan.Records;
                        ExcelDocument.MarkPivotCacheRecordsPartAsModelWritten(recordPart);
                        cachePart.PivotCacheDefinition = plan.Cache;
                        using (BeginNoLock()) {
                            foreach (var item in plans) {
                                token.ThrowIfCancellationRequested();
                                item.Sheet.WritePivotMaterializationView(item.Plan, token);
                            }
                        }
                        var validator = new OpenXmlValidator(FileFormatVersions.Microsoft365) { MaxNumberOfErrors = 1 };
                        foreach (OpenXmlPart changed in plans.SelectMany(item => new OpenXmlPart[] {
                            item.Sheet._worksheetPart, item.Plan.Part }).Concat(new OpenXmlPart[] { cachePart, recordPart }).Distinct()) {
                            token.ThrowIfCancellationRequested();
                            var error = validator.Validate(changed, token).FirstOrDefault();
                            if (error != null) throw new InvalidOperationException("Generated pivot metadata is invalid: " + error.Description);
                        }
                        return affectedCells;
                    }, effective, cancellationToken);
                } catch {
                    foreach (var item in plans) item.Sheet.ResetMutationCaches();
                    throw;
                }
                result = new ExcelPivotMaterializationResult(plan.Definition.Name!.Value!, plan.Definition.Location!.Reference!.Value!,
                    plan.SourceRecords, mutation, plans.Select(item => item.Plan.Definition.Name!.Value!).ToArray());
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
            internal ExcelDateSystem SourceDateSystem;
            internal HashSet<(int Row, int Column)> DateCells = new();
            internal Dictionary<(int Row, int Column), uint> NumberFormatCells = new();
            internal int Top, Left, Bottom, Right, OldBottom, OldRight, AffectedBottom, AffectedRight, SourceRecords;
            internal long SourceCellVisits, MeasureInputVisits;
        }

        private static int ValidateCoordinatedPivotPlans(
            (ExcelSheet Sheet, PivotMaterializationPlan Plan)[] plans, ExcelMutationPlanOptions options) {
            if (plans.Length == 0) throw new InvalidOperationException("The pivot cache has no views.");
            PivotMaterializationPlan first = plans[0].Plan;
            int limit = Math.Min(1_000_000, options.MaximumAffectedCells);
            long totalCells = 0, sourceCellVisits = 0, measureInputVisits = 0;
            foreach (var item in plans) {
                PivotMaterializationPlan current = item.Plan;
                if (plans.Length > 1 && !ReferenceEquals(current, first)
                    && (!string.Equals(current.Cache.OuterXml, first.Cache.OuterXml, StringComparison.Ordinal)
                        || !string.Equals(current.Records.OuterXml, first.Records.OuterXml, StringComparison.Ordinal)))
                    throw new NotSupportedException("Shared pivot views require identical generated cache fields and records.");
                totalCells += (long)(current.AffectedBottom - current.Top + 1) * (current.AffectedRight - current.Left + 1);
                sourceCellVisits += current.SourceCellVisits;
                measureInputVisits += current.MeasureInputVisits;
                if (totalCells > limit) throw new InvalidOperationException("The combined pivot output exceeds the materialization budget.");
                if (sourceCellVisits > limit || measureInputVisits > limit)
                    throw new InvalidOperationException("The combined pivot source work exceeds the materialization budget.");
            }
            for (int left = 0; left < plans.Length; left++) {
                for (int right = left + 1; right < plans.Length; right++) {
                    if (!ReferenceEquals(plans[left].Sheet._worksheetPart, plans[right].Sheet._worksheetPart)) continue;
                    PivotMaterializationPlan a = plans[left].Plan, b = plans[right].Plan;
                    if (a.Top <= b.AffectedBottom && a.AffectedBottom >= b.Top
                        && a.Left <= b.AffectedRight && a.AffectedRight >= b.Left)
                        throw new InvalidOperationException("The materialized pivot views would overlap.");
                }
            }
            return (int)totalCells;
        }

        private void WritePivotMaterializationView(PivotMaterializationPlan plan, CancellationToken token) {
            plan.Part.PivotTableDefinition = plan.Definition;
            ClearExistingCellFieldsInRange((plan.Top, plan.Left, plan.AffectedBottom, plan.AffectedRight), ExcelClearOptions.Values);
            ClearHeaderCache();
            var dateStyles = new Dictionary<uint, uint>();
            var numberStyles = new Dictionary<(uint BaseStyle, uint Format), uint>();
            var existingDateStyles = plan.DateCells.Count == 0 ? null : StylesCache.Build(_spreadSheetDocument);
            for (int row = plan.Top; row <= plan.AffectedBottom; row++) {
                token.ThrowIfCancellationRequested();
                for (int column = plan.Left; column <= plan.AffectedRight; column++) {
                    if (row > plan.Bottom || column > plan.Right) continue;
                    var value = plan.Values[row - plan.Top, column - plan.Left];
                    if (value?.Kind == ExcelCellDataKind.Error) CellError(row, column, (string)value.Value!);
                    else if (value != null) CellValue(row, column, value.Value);
                    if (plan.NumberFormatCells.TryGetValue((row - plan.Top, column - plan.Left), out uint formatId)) {
                        var cell = GetCell(row, column);
                        uint style = cell.StyleIndex?.Value ?? 0U;
                        if (!numberStyles.TryGetValue((style, formatId), out uint numberStyle))
                            numberStyles.Add((style, formatId), numberStyle = GetOrCreateBuiltInNumberFormatStyleIndex(style, formatId));
                        cell.StyleIndex = numberStyle;
                    } else if (plan.DateCells.Contains((row - plan.Top, column - plan.Left))) {
                        var cell = GetCell(row, column);
                        uint style = cell.StyleIndex?.Value ?? 0U;
                        if (existingDateStyles?.IsDateLike(style) != true) {
                            if (!dateStyles.TryGetValue(style, out uint dateStyle))
                                dateStyles.Add(style, dateStyle = GetOrCreateBuiltInNumberFormatStyleIndex(style, 22));
                            cell.StyleIndex = dateStyle;
                        }
                    }
                }
            }
            WorksheetRoot.Save();
            MarkRequiresSavePreparation();
        }

        private PivotMaterializationPlan PreparePivotMaterialization(string name, ExcelMutationPlanOptions options,
            int remainingInputVisits, int remainingOutputCells, IReadOnlyList<int> dateFilterFields,
            DateTime referenceDate, CancellationToken token) {
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
            sourceSheet._pivotStylesCache = null;
            int fieldCount = c2 - c1 + 1;
            int limit = Math.Min(1_000_000, options.MaximumAffectedCells);
            if (r2 <= r1 || fieldCount > 256 || (long)(r2 - r1) * fieldCount > limit)
                throw new InvalidOperationException("The pivot source exceeds the materialization budget or has no data records.");
            var fields = cache.CacheFields?.Elements<CacheField>().Take(257).ToArray() ?? Array.Empty<CacheField>();
            var pivotFields = definition.PivotFields?.Elements<PivotField>().Take(257).ToArray() ?? Array.Empty<PivotField>();
            var measures = definition.DataFields?.Elements<DataField>().Take(257).ToArray() ?? Array.Empty<DataField>();
            var pages = definition.PageFields?.Elements<PageField>().Take(257).ToArray() ?? Array.Empty<PageField>();
            if (fields.Length < fieldCount || fields.Length > 256 || pivotFields.Length != fields.Length || measures.Length == 0 || measures.Length > 256
                || fields.Any(f => f.Formula != null || f.DatabaseField?.Value == false && !IsDerivedDateGroup(f) && !IsDerivedManualGroup(f))
                || pages.Length > 256 || pages.Length != (definition.PageFields?.ChildElements.Count ?? 0)
                || measures.Any(m => m.ShowDataAs?.Value is ShowDataAsValues mode && mode != ShowDataAsValues.Normal))
                throw new NotSupportedException("Materialization requires ordinary measures and no calculated fields.");
            if (!sourceSheet.BuildPivotHeaders(r1, c1, c2).SequenceEqual(fields.Take(fieldCount).Select(f => f.Name?.Value ?? ""), StringComparer.OrdinalIgnoreCase))
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
            var pageFields = pages.Select(page => page.Field?.Value ?? -1).ToArray();
            if (realFields.Any(field => field >= fields.Length) || pageFields.Any(field => field < 0 || field >= fieldCount)
                || realFields.Concat(pageFields).Distinct().Count() != realFields.Length + pageFields.Length)
                throw new NotSupportedException("The pivot axes or page fields do not match distinct source fields.");
            if (realFields.Any(field => pivotFields[field].ShowAll?.Value == true))
                throw new NotSupportedException("Pivot fields set to show all items can add empty combinations on Excel refresh and are not qualified for headless materialization.");
            var groupings = new Dictionary<int, PivotNumericGrouping>();
            var dateGroupings = new Dictionary<int, PivotDateGrouping>();
            var manualGroupings = new Dictionary<int, PivotManualGrouping>();
            var sourceDateGroupings = new Dictionary<int, ExcelPivotGrouping>();
            for (int field = fieldCount; field < fields.Length; field++) {
                if (!realFields.Contains(field))
                    throw new NotSupportedException("Derived fields require a row or column axis.");
                if (IsDerivedDateGroup(fields[field])) {
                    var date = ReadPivotDateGrouping(fields, field, fieldCount);
                    dateGroupings.Add(field, date);
                    sourceDateGroupings[date.SourceField] = ExcelPivotGrouping.Date(fields[date.SourceField].Name?.Value ?? "", date.GroupBy);
                } else if (IsDerivedManualGroup(fields[field])) {
                    manualGroupings.Add(field, ReadPivotManualGrouping(fields, pivotFields, field, fieldCount));
                } else throw new NotSupportedException("The derived pivot field is not qualified for materialization.");
            }
            for (int field = 0; field < fieldCount; field++) {
                if (fields[field].FieldGroup == null) continue;
                if (!realFields.Contains(field) && sourceDateGroupings.ContainsKey(field)) continue;
                if (manualGroupings.Values.Any(group => group.SourceField == field)) continue;
                if (!realFields.Contains(field) || pageFields.Contains(field))
                    throw new NotSupportedException("Grouped fields require a row or column axis.");
                groupings.Add(field, ReadPivotNumericGrouping(fields[field], field)!);
            }
            if (pivotFields.Where((field, index) => !realFields.Contains(index) && !pageFields.Contains(index))
                .Any(field => field.Items?.Elements<Item>().Any(item => item.Hidden?.Value == true) == true))
                throw new NotSupportedException("Hidden items outside the pivot axes and page fields cannot be materialized.");
            bool rowTotal = rowField >= 0 && definition.ColumnGrandTotals?.Value != false;
            bool columnTotal = columnField >= 0 && definition.RowGrandTotals?.Value != false;
            long visits = (long)(r2 - r1) * measures.Length * MaterializedInputLevels(rowAxis, pivotFields)
                * MaterializedInputLevels(columnAxis, pivotFields);
            if (visits > remainingInputVisits)
                throw new InvalidOperationException("The combined pivot measure input visits, including intermediate subtotals, exceed the materialization budget.");
            if (_excelDocument.GetWorkbookSlicerCaches().Concat(_excelDocument.GetWorkbookTimelineCaches())
                .Any(c => string.IsNullOrEmpty(c.PivotTableName) || string.Equals(c.PivotTableName, definition.Name?.Value, StringComparison.OrdinalIgnoreCase)))
                throw new NotSupportedException("Materialization of pivot interaction caches requires coordinated cache updates.");
            if (!A1.TryParseRange(definition.Location?.Reference?.Value ?? "", out int top, out int left, out int oldBottom, out int oldRight))
                throw new InvalidOperationException("The pivot output location is invalid.");
            var grouping = groupings.ToDictionary(pair => pair.Key, pair => pair.Value.SourceGrouping);
            foreach (var pair in sourceDateGroupings) grouping.Add(pair.Key, pair.Value);
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
            // Shared-cache keys use one chronological order across every view. Each
            // view can retain a different manual or descending axis order below.
            foreach (int field in dateFilterFields) {
                if (field >= fieldCount || fields[field].FieldGroup != null
                    || sourceDateGroupings.ContainsKey(field)
                    || !maps[field].Items.Any(key => key.Kind == PivotFieldValueKind.Date)
                    || maps[field].Items.Any(key => key.Kind != PivotFieldValueKind.Date && key.Kind != PivotFieldValueKind.Blank))
                    continue;
                var sorted = maps[field].Items.OrderBy(key => key.Kind == PivotFieldValueKind.Blank)
                    .ThenBy(key => key.Date).ToArray();
                maps[field] = new PivotFieldValues(sorted);
            }
            var dateFieldOrders = new Dictionary<int, int[]>();
            foreach (int field in dateFilterFields) {
                if (field >= fieldCount || !realFields.Contains(field)
                    || maps[field].Items.Count == 0 || maps[field].Items.Any(key =>
                        key.Kind != PivotFieldValueKind.Date && key.Kind != PivotFieldValueKind.Blank))
                    continue;
                var sort = pivotFields[field].SortType?.Value;
                IReadOnlyList<PivotFieldValue> ordered = sort == FieldSortValues.Descending
                    ? maps[field].Items.Where(key => key.Kind == PivotFieldValueKind.Date).Reverse()
                            .Concat(maps[field].Items.Where(key => key.Kind == PivotFieldValueKind.Blank)).ToArray()
                    : sort == FieldSortValues.Manual || HasSavedPivotDateOrder(pivotFields[field])
                        ? OrderManualDateValues(maps[field], fields[field], pivotFields[field]).Items
                        : maps[field].Items;
                var index = IndexPivotMaterializationKeys(maps[field]);
                dateFieldOrders.Add(field, ordered.Select(key => index[key]).ToArray());
            }
            foreach (int sourceField in manualGroupings.Values.Select(g => g.SourceField).Distinct())
                maps[sourceField] = OrderManualSourceValues(maps[sourceField], fields[sourceField], pivotFields[sourceField]);
            foreach (var manual in manualGroupings.Values) manual.IncludeSourceKeys(maps[manual.SourceField]);
            var displayMaps = maps.ToList();
            foreach (var pair in groupings) displayMaps[pair.Key] = pair.Value.Labels;
            for (int field = fieldCount; field < fields.Length; field++)
                displayMaps.Add(dateGroupings.TryGetValue(field, out var date) ? date.Labels : manualGroupings[field].Labels);
            var captions = ReadPivotMaterializationCaptions(fields, pivotFields, displayMaps,
                realFields.Concat(pageFields),
                groupings, dateGroupings, manualGroupings);
            var visibility = BuildPivotMaterializationVisibility(sourceSheet, fields, pivotFields, pages,
                displayMaps, realFields, groupings, dateGroupings, manualGroupings, r1, r2, c1, limit, token);
            ApplyMaterializedPivotFilters(sourceSheet, definition, fields, displayMaps, captions, measures,
                realFields, groupings, dateGroupings, manualGroupings, visibility, r1, r2, c1, limit,
                referenceDate, token);
            // Lookup indexes both the saved field items and their shared keys. Keep every
            // possible criterion combination usable, rather than accepting an unreadable view.
            foreach (var axis in new[] { rowAxis, columnAxis }) {
                long indexedItems = axis.RealFields.Sum(field => 2L * displayMaps[field].Items.Count
                    + (MaterializedAutomaticSubtotal(pivotFields[field]) ? 1 : 0));
                if (indexedItems > limit)
                    throw new InvalidOperationException("The pivot criteria index exceeds the materialization budget.");
            }
            var rows = BuildMaterializedHierarchyAxis(sourceSheet, r1, r2, c1, displayMaps, rowAxis, pivotFields,
                groupings, dateGroupings, manualGroupings, dateFieldOrders,
                visibility.IncludedRows, measures.Length, rowTotal, token);
            var columns = BuildMaterializedHierarchyAxis(sourceSheet, r1, r2, c1, displayMaps, columnAxis, pivotFields,
                groupings, dateGroupings, manualGroupings, dateFieldOrders,
                visibility.IncludedRows, measures.Length, columnTotal, token);
            int dataRow = columnField >= 0 ? columnAxis.Fields.Length + 1 : 1;
            int dataColumn = rowAxis.Fields.Length > 0 ? rowAxis.Fields.Length : measures.Length == 1 && columnField >= 0 ? 1 : 0;
            int height = dataRow + rows.Entries.Count;
            int width = dataColumn + columns.Entries.Count;
            bool hasOldView = definition.RowItems?.ChildElements.Count > 0 && definition.ColumnItems?.ChildElements.Count > 0;
            if (!hasOldView) { oldBottom = top - 1; oldRight = left - 1; }
            long bottom = (long)top + height - 1, right = (long)left + width - 1;
            long affectedBottom = Math.Max(bottom, oldBottom), affectedRight = Math.Max(right, oldRight);
            if (bottom > 1_048_576 || right > 16_384 || (long)height * width > limit
                || (affectedBottom - top + 1) * (affectedRight - left + 1) > remainingOutputCells)
                throw new InvalidOperationException("The pivot output exceeds the worksheet or materialization budget.");
            ValidatePivotMaterializationDestination(part, sourceSheet, r1, c1, r2, c2, top, left, (int)affectedBottom, (int)affectedRight, oldBottom, oldRight, token);
            var plan = new PivotMaterializationPlan { Part = part, CachePart = cachePart, Definition = (PivotTableDefinition)definition.CloneNode(true),
                Cache = (PivotCacheDefinition)cache.CloneNode(true), Top = top, Left = left, Bottom = (int)bottom, Right = (int)right,
                OldBottom = oldBottom, OldRight = oldRight, AffectedBottom = (int)affectedBottom, AffectedRight = (int)affectedRight, SourceRecords = r2 - r1,
                SourceDateSystem = _excelDocument.DateSystem, SourceCellVisits = (long)(r2 - r1) * fieldCount,
                MeasureInputVisits = visits };
            UpdateMaterializedPivotRelativeDateBounds(plan.Definition, referenceDate, _excelDocument.DateSystem);
            var aggregates = AggregateMaterializedHierarchy(sourceSheet, r1, c1, r2, rows, columns,
                visibility.IncludedRows, measures, limit, token);
            bool dateRowHierarchy = rowAxis.RealFields.Length > 1 && rowAxis.RealFields.All(dateGroupings.ContainsKey);
            bool dateColumnHierarchy = columnAxis.RealFields.Length > 1 && columnAxis.RealFields.All(dateGroupings.ContainsKey);
            bool manualRowHierarchy = rowAxis.RealFields.Length == 2 && manualGroupings.TryGetValue(rowAxis.RealFields[0], out var rowManual)
                && rowManual.SourceField == rowAxis.RealFields[1];
            bool manualColumnHierarchy = columnAxis.RealFields.Length == 2 && manualGroupings.TryGetValue(columnAxis.RealFields[0], out var columnManual)
                && columnManual.SourceField == columnAxis.RealFields[1];
            FillMaterializedHierarchy(plan, displayMaps, captions, rows, columns, visibility, measures, dateFieldOrders,
                dataRow, dataColumn, aggregates,
                dateRowHierarchy, dateColumnHierarchy, manualRowHierarchy, manualColumnHierarchy, token);
            var cacheFields = plan.Cache.CacheFields!.Elements<CacheField>().ToArray();
            var savedPivotFields = plan.Definition.PivotFields!.Elements<PivotField>().ToArray();
            for (int field = 0; field < fieldCount; field++) {
                if (realFields.Contains(field) || pageFields.Contains(field) || savedPivotFields[field].Items == null) continue;
                bool hasDefault = savedPivotFields[field].Items!.Elements<Item>()
                    .Any(item => item.ItemType?.Value == ItemValues.Default);
                savedPivotFields[field].Items = CreateMaterializedFilteredItems(maps[field], null, hasDefault, false);
            }
            for (int field = 0; field < fieldCount; field++) cacheFields[field].SharedItems = BuildSharedItems(maps[field], null);
            foreach (var pair in manualGroupings) pair.Value.RewriteCacheField(cacheFields[pair.Key], maps[pair.Value.SourceField]);
            plan.Records = sourceSheet.BuildPivotCacheRecords(fieldCount, r1 + 1, r2, c1, grouping, maps, collect,
                Array.Empty<GeneratedPivotGroupingField>(), Array.Empty<PivotFieldValues>(), 0);
            plan.Cache.RecordCount = (uint)plan.SourceRecords;
            plan.Cache.SaveData = true;
            plan.Cache.RefreshOnLoad = false;
            return plan;
        }

        private static PivotFieldValues OrderManualDateValues(PivotFieldValues current, CacheField source,
            PivotField pivotField) {
            var saved = source.SharedItems?.ChildElements.ToArray() ?? Array.Empty<OpenXmlElement>();
            var available = new HashSet<PivotFieldValue>(current.Items);
            var ordered = new List<PivotFieldValue>(current.Items.Count);
            foreach (var item in pivotField.Items?.Elements<Item>() ?? Enumerable.Empty<Item>()) {
                if (item.ItemType?.Value == ItemValues.Default) continue;
                PivotFieldValue key = OriginalPivotMaterializationKey(item, saved, false);
                if (available.Remove(key)) ordered.Add(key);
            }
            foreach (var value in current.Items) if (available.Remove(value)) ordered.Add(value);
            return new PivotFieldValues(ordered);
        }

        private static bool HasSavedPivotDateOrder(PivotField field) {
            int position = 0;
            foreach (var item in field.Items?.Elements<Item>() ?? Enumerable.Empty<Item>()) {
                if (item.ItemType?.Value == ItemValues.Default) continue;
                if (item.Index?.Value != position) return true;
                position++;
            }
            return false;
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
