using OfficeIMO.Excel.LegacyXls.Diagnostics;
using OfficeIMO.Excel.LegacyXls.Model;
using System.Text;

namespace OfficeIMO.Excel.LegacyXls {
    /// <summary>
    /// Compact import summary intended for corpus baselines and preflight checks.
    /// </summary>
    public sealed class LegacyXlsImportReport {
        internal LegacyXlsImportReport(LegacyXlsWorkbook workbook) {
            if (workbook == null) throw new ArgumentNullException(nameof(workbook));

            WorksheetCount = workbook.Worksheets.Count;
            UnsupportedSheetCount = workbook.UnsupportedSheets.Count;
            CellCount = workbook.Worksheets.Sum(sheet => sheet.Cells.Count);
            FormulaCellCount = workbook.Worksheets.Sum(sheet => sheet.Cells.Count(cell => cell.IsFormula));
            CommentCount = workbook.Worksheets.Sum(sheet => sheet.Comments.Count);
            HyperlinkCount = workbook.Worksheets.Sum(sheet => sheet.Hyperlinks.Count);
            DataValidationCount = workbook.Worksheets.Sum(sheet => sheet.DataValidations.Count);
            ConditionalFormattingCount = workbook.Worksheets.Sum(sheet => sheet.ConditionalFormattings.Count);
            AutoFilterCriteriaCount = workbook.Worksheets.Sum(sheet => sheet.AutoFilterCriteria.Count);
            DataValidationsByType = CountByCode(workbook.Worksheets.SelectMany(sheet => sheet.DataValidations).Select(validation => validation.Type.ToString()));
            DataValidationsByOperator = CountByCode(workbook.Worksheets.SelectMany(sheet => sheet.DataValidations).Select(validation => validation.Operator.ToString()));
            ConditionalFormattingsByType = CountByCode(workbook.Worksheets.SelectMany(sheet => sheet.ConditionalFormattings).Select(formatting => formatting.Type.ToString()));
            ConditionalFormattingsByOperator = CountByCode(workbook.Worksheets
                .SelectMany(sheet => sheet.ConditionalFormattings)
                .Where(formatting => formatting.Operator.HasValue)
                .Select(formatting => formatting.Operator!.Value.ToString()));
            AutoFilterCriteriaByOperator = CountByCode(workbook.Worksheets
                .SelectMany(sheet => sheet.AutoFilterCriteria)
                .SelectMany(criteria => criteria.Conditions)
                .Select(condition => condition.Operator.ToString()));
            DefinedNameCount = workbook.DefinedNames.Count;
            ExternalReferenceCount = workbook.ExternalReferences.Count;
            ExternalSheetNameCount = workbook.ExternalReferences.Sum(reference => reference.SheetNames.Count);
            ExternalNameCount = workbook.ExternalReferences.Sum(reference => reference.ExternalNames.Count);
            ExternalCellCacheCount = workbook.ExternalReferences.Sum(reference => reference.CachedCellCaches.Count);
            ExternalCachedCellCount = workbook.ExternalReferences.Sum(reference => reference.CachedCellCaches.Sum(cache => cache.Cells.Count));
            PivotTableRecordCount = workbook.PivotTableRecords.Count;
            ChartRecordCount = workbook.ChartRecords.Count;
            DrawingRecordCount = workbook.DrawingRecords.Count;
            CompoundFeatureRecordCount = workbook.CompoundFeatureRecords.Count;
            CompoundFeatureEntryCount = workbook.CompoundFeatureRecords.Sum(record => record.Entries.Count);
            CalculationSettingRecordCount = workbook.CalculationSettings.Records.Count;
            CellStyleRecordCount = workbook.CellStyles.Count;
            WorkbookMetadataRecordCount = workbook.MetadataRecords.Count;
            WorksheetMetadataRecordCount = workbook.Worksheets.Sum(sheet => sheet.MetadataRecords.Count);
            UnsupportedSheetMetadataRecordCount = workbook.UnsupportedSheets.Sum(sheet => sheet.MetadataRecords.Count);
            UnsupportedFeatureCount = workbook.UnsupportedFeatures.Count;
            PreservedFeatureRecordCount = workbook.PreservedFeatureRecords.Count;
            ErrorCount = workbook.Diagnostics.Count(diagnostic => diagnostic.Severity == LegacyXlsDiagnosticSeverity.Error);
            WarningCount = workbook.Diagnostics.Count(diagnostic => diagnostic.Severity == LegacyXlsDiagnosticSeverity.Warning);
            DiagnosticsByCode = CountByCode(workbook.Diagnostics.Select(diagnostic => diagnostic.Code));
            FormulaTokenBlockers = CountByCode(workbook.Diagnostics
                .Where(diagnostic => string.Equals(diagnostic.Code, "XLS-BIFF-FORMULA-TOKENS-UNSUPPORTED", StringComparison.OrdinalIgnoreCase))
                .Select(diagnostic => diagnostic.DetailCode ?? "FormulaUnknown"));
            UnsupportedFeaturesByCode = CountByCode(workbook.UnsupportedFeatures.Select(feature => feature.Code));
            UnsupportedFeaturesByKind = CountByKind(workbook.UnsupportedFeatures);
            UnsupportedFeaturesByRecordType = CountByCode(workbook.UnsupportedFeatures
                .Where(feature => feature.RecordType.HasValue)
                .Select(feature => $"{feature.Kind}|{feature.Code}|0x{feature.RecordType!.Value:X4}"));
            UnsupportedFeaturesByDetail = CountByCode(workbook.UnsupportedFeatures
                .Where(feature => !string.IsNullOrWhiteSpace(feature.DetailCode))
                .Select(feature => $"{feature.Kind}|{feature.Code}|{feature.DetailCode}"));
            UnsupportedFeaturesByLocation = CountByCode(workbook.UnsupportedFeatures
                .Select(GetFeatureLocationKey));
            UnsupportedSheetsByKind = CountUnsupportedSheetsByKind(workbook.UnsupportedSheets);
            UnsupportedSheetsByType = CountByCode(workbook.UnsupportedSheets.Select(sheet => $"0x{sheet.SheetType:X2}|{sheet.Kind}"));
            UnsupportedSheetsByName = CountByCode(workbook.UnsupportedSheets.Select(sheet => sheet.Name));
            ExternalReferencesByKind = CountExternalReferencesByKind(workbook.ExternalReferences);
            ExternalReferencesByTarget = CountByCode(workbook.ExternalReferences.Select(GetExternalReferenceTargetKey));
            ExternalSheetNamesByReferenceKind = CountExternalSheetNamesByReferenceKind(workbook.ExternalReferences);
            ExternalNamesByReferenceKind = CountExternalNamesByReferenceKind(workbook.ExternalReferences);
            ExternalNamesByName = CountByCode(workbook.ExternalReferences.SelectMany(reference => reference.ExternalNames.Select(name => name.Name)));
            ExternalCellCachesBySheetName = CountByCode(workbook.ExternalReferences.SelectMany(reference => reference.CachedCellCaches.Select(GetExternalCellCacheSheetKey)));
            ExternalCachedCellsByValueKind = CountExternalCachedCellsByValueKind(workbook.ExternalReferences);
            PivotTableRecordsByKind = CountPivotTableRecordsByKind(workbook.PivotTableRecords);
            PivotTableRecordsByName = CountByCode(workbook.PivotTableRecords.Select(record => record.RecordName));
            PivotTableDataItemAggregations = CountByCode(workbook.PivotTableRecords
                .Where(record => record.AggregationFunction.HasValue)
                .Select(record => $"AggregationFunction:{record.AggregationFunction!.Value}"));
            PivotTableGroupingKinds = CountByCode(workbook.PivotTableRecords
                .Where(record => record.GroupingKind.HasValue)
                .Select(record => record.GroupingKind!.Value.ToString()));
            ChartRecordsByKind = CountChartRecordsByKind(workbook.ChartRecords);
            ChartRecordsByName = CountByCode(workbook.ChartRecords.Select(record => record.RecordName));
            ChartRecordsByChartType = CountByCode(workbook.ChartRecords
                .Where(record => !string.IsNullOrWhiteSpace(record.ChartTypeName))
                .Select(record => record.ChartTypeName!));
            ChartRecordsByRectangle = CountByCode(workbook.ChartRecords
                .Where(record => record.ChartX.HasValue && record.ChartY.HasValue && record.ChartWidth.HasValue && record.ChartHeight.HasValue)
                .Select(record => $"X:{record.ChartX!.Value};Y:{record.ChartY!.Value};Width:{record.ChartWidth!.Value};Height:{record.ChartHeight!.Value}"));
            ChartRecordsByAxisType = CountByCode(workbook.ChartRecords
                .Where(record => !string.IsNullOrWhiteSpace(record.AxisTypeName))
                .Select(record => record.AxisTypeName!));
            ChartRecordsByLocation = CountByCode(workbook.ChartRecords.Select(GetChartRecordLocationKey));
            DrawingRecordsByKind = CountDrawingRecordsByKind(workbook.DrawingRecords);
            DrawingRecordsByName = CountByCode(workbook.DrawingRecords.Select(record => record.RecordName));
            DrawingRecordsByObjectType = CountByCode(workbook.DrawingRecords
                .Where(record => record.ObjectType.HasValue)
                .Select(record => $"ObjectType:0x{record.ObjectType!.Value:X4}"));
            DrawingRecordsByEscherRecordType = CountByCode(workbook.DrawingRecords
                .Where(record => record.EscherRecordType.HasValue)
                .Select(record => $"EscherRecordType:0x{record.EscherRecordType!.Value:X4}"));
            DrawingRecordsByLocation = CountByCode(workbook.DrawingRecords.Select(GetDrawingRecordLocationKey));
            CompoundFeatureRecordsByKind = CountCompoundFeatureRecordsByKind(workbook.CompoundFeatureRecords);
            CompoundFeatureEntriesByKind = CountCompoundFeatureEntriesByKind(workbook.CompoundFeatureRecords);
            CompoundFeatureEntriesByName = CountByCode(workbook.CompoundFeatureRecords.SelectMany(record => record.Entries));
            CalculationSettingsByKind = CountCalculationSettingsByKind(workbook.CalculationSettings.Records);
            CellStylesByKind = CountByCode(workbook.CellStyles.Select(style => style.IsBuiltIn ? "BuiltIn" : "Custom"));
            WorkbookMetadataRecordsByKind = CountWorkbookMetadataRecordsByKind(workbook.MetadataRecords);
            WorksheetMetadataRecordsByKind = CountWorksheetMetadataRecordsByKind(workbook.Worksheets.SelectMany(sheet => sheet.MetadataRecords));
            UnsupportedSheetMetadataRecordsByKind = CountUnsupportedSheetMetadataRecordsByKind(workbook.UnsupportedSheets.SelectMany(sheet => sheet.MetadataRecords));
            PreservedFeatureRecordsByKind = CountPreservedRecordsByKind(workbook.PreservedFeatureRecords);
            PreservedFeatureRecordsByDetail = CountByCode(workbook.PreservedFeatureRecords
                .Where(record => !string.IsNullOrWhiteSpace(record.DetailCode))
                .Select(record => $"{record.Kind}|{record.Code}|{record.DetailCode}"));
        }

        /// <summary>Gets the number of imported worksheet sheets.</summary>
        public int WorksheetCount { get; }

        /// <summary>Gets the number of sheet entries that were preserved as unsupported metadata.</summary>
        public int UnsupportedSheetCount { get; }

        /// <summary>Gets the number of imported cells, including blank style-only cells.</summary>
        public int CellCount { get; }

        /// <summary>Gets the number of imported formula cells.</summary>
        public int FormulaCellCount { get; }

        /// <summary>Gets the number of imported comments.</summary>
        public int CommentCount { get; }

        /// <summary>Gets the number of imported hyperlinks.</summary>
        public int HyperlinkCount { get; }

        /// <summary>Gets the number of imported data validation rules.</summary>
        public int DataValidationCount { get; }

        /// <summary>Gets the number of imported conditional formatting rules.</summary>
        public int ConditionalFormattingCount { get; }

        /// <summary>Gets the number of imported AutoFilter criteria columns.</summary>
        public int AutoFilterCriteriaCount { get; }

        /// <summary>Gets imported data validations grouped by validation type.</summary>
        public IReadOnlyDictionary<string, int> DataValidationsByType { get; }

        /// <summary>Gets imported data validations grouped by comparison operator.</summary>
        public IReadOnlyDictionary<string, int> DataValidationsByOperator { get; }

        /// <summary>Gets imported conditional formatting rules grouped by rule type.</summary>
        public IReadOnlyDictionary<string, int> ConditionalFormattingsByType { get; }

        /// <summary>Gets imported conditional formatting cell-is rules grouped by comparison operator.</summary>
        public IReadOnlyDictionary<string, int> ConditionalFormattingsByOperator { get; }

        /// <summary>Gets imported AutoFilter conditions grouped by comparison operator.</summary>
        public IReadOnlyDictionary<string, int> AutoFilterCriteriaByOperator { get; }

        /// <summary>Gets the number of imported defined names.</summary>
        public int DefinedNameCount { get; }

        /// <summary>Gets the number of preserved external-reference records.</summary>
        public int ExternalReferenceCount { get; }

        /// <summary>Gets the number of external workbook sheet names declared by supporting links.</summary>
        public int ExternalSheetNameCount { get; }

        /// <summary>Gets the number of external names declared by supporting links.</summary>
        public int ExternalNameCount { get; }

        /// <summary>Gets the number of preserved external cell cache sections.</summary>
        public int ExternalCellCacheCount { get; }

        /// <summary>Gets the number of preserved cached external cell values.</summary>
        public int ExternalCachedCellCount { get; }

        /// <summary>Gets the number of preserve-only PivotTable BIFF records discovered during import.</summary>
        public int PivotTableRecordCount { get; }

        /// <summary>Gets the number of preserve-only chart BIFF records discovered during import.</summary>
        public int ChartRecordCount { get; }

        /// <summary>Gets the number of preserve-only drawing and object BIFF records discovered during import.</summary>
        public int DrawingRecordCount { get; }

        /// <summary>Gets the number of preserve-only compound container features discovered during import.</summary>
        public int CompoundFeatureRecordCount { get; }

        /// <summary>Gets the number of matching compound directory entries behind preserve-only compound features.</summary>
        public int CompoundFeatureEntryCount { get; }

        /// <summary>Gets the number of calculation setting records parsed from BIFF records.</summary>
        public int CalculationSettingRecordCount { get; }

        /// <summary>Gets the number of workbook cell style records parsed from Style records.</summary>
        public int CellStyleRecordCount { get; }

        /// <summary>Gets the number of workbook metadata records parsed from BIFF records.</summary>
        public int WorkbookMetadataRecordCount { get; }

        /// <summary>Gets the number of worksheet metadata records parsed from BIFF records.</summary>
        public int WorksheetMetadataRecordCount { get; }

        /// <summary>Gets the number of metadata records parsed from unsupported sheet substreams.</summary>
        public int UnsupportedSheetMetadataRecordCount { get; }

        /// <summary>Gets the number of unsupported or preserve-only feature findings.</summary>
        public int UnsupportedFeatureCount { get; }

        /// <summary>Gets the number of preserve-only BIFF feature records with typed metadata.</summary>
        public int PreservedFeatureRecordCount { get; }

        /// <summary>Gets the number of error diagnostics produced during import.</summary>
        public int ErrorCount { get; }

        /// <summary>Gets the number of warning diagnostics produced during import.</summary>
        public int WarningCount { get; }

        /// <summary>Gets diagnostic counts grouped by stable diagnostic code.</summary>
        public IReadOnlyDictionary<string, int> DiagnosticsByCode { get; }

        /// <summary>Gets unsupported formula token blockers grouped by stable detail key.</summary>
        public IReadOnlyDictionary<string, int> FormulaTokenBlockers { get; }

        /// <summary>Gets unsupported/preserve-only feature counts grouped by stable feature code.</summary>
        public IReadOnlyDictionary<string, int> UnsupportedFeaturesByCode { get; }

        /// <summary>Gets unsupported/preserve-only feature counts grouped by feature kind.</summary>
        public IReadOnlyDictionary<LegacyXlsUnsupportedFeatureKind, int> UnsupportedFeaturesByKind { get; }

        /// <summary>Gets unsupported/preserve-only feature counts grouped by kind, code, and BIFF record type.</summary>
        public IReadOnlyDictionary<string, int> UnsupportedFeaturesByRecordType { get; }

        /// <summary>Gets unsupported/preserve-only feature counts grouped by kind, code, and stable feature subtype.</summary>
        public IReadOnlyDictionary<string, int> UnsupportedFeaturesByDetail { get; }

        /// <summary>Gets unsupported/preserve-only feature counts grouped by code and workbook or sheet location.</summary>
        public IReadOnlyDictionary<string, int> UnsupportedFeaturesByLocation { get; }

        /// <summary>Gets unsupported sheet entries grouped by decoded sheet kind.</summary>
        public IReadOnlyDictionary<LegacyXlsUnsupportedSheetKind, int> UnsupportedSheetsByKind { get; }

        /// <summary>Gets unsupported sheet entries grouped by raw BoundSheet type and decoded kind.</summary>
        public IReadOnlyDictionary<string, int> UnsupportedSheetsByType { get; }

        /// <summary>Gets unsupported sheet entries grouped by sheet name.</summary>
        public IReadOnlyDictionary<string, int> UnsupportedSheetsByName { get; }

        /// <summary>Gets preserved external references grouped by supporting-link kind.</summary>
        public IReadOnlyDictionary<LegacyXlsExternalReferenceKind, int> ExternalReferencesByKind { get; }

        /// <summary>Gets preserved external references grouped by target path or source.</summary>
        public IReadOnlyDictionary<string, int> ExternalReferencesByTarget { get; }

        /// <summary>Gets external workbook sheet-name counts grouped by supporting-link kind.</summary>
        public IReadOnlyDictionary<LegacyXlsExternalReferenceKind, int> ExternalSheetNamesByReferenceKind { get; }

        /// <summary>Gets external defined-name counts grouped by supporting-link kind.</summary>
        public IReadOnlyDictionary<LegacyXlsExternalReferenceKind, int> ExternalNamesByReferenceKind { get; }

        /// <summary>Gets external defined names grouped by name text.</summary>
        public IReadOnlyDictionary<string, int> ExternalNamesByName { get; }

        /// <summary>Gets external cell cache sections grouped by resolved external sheet name.</summary>
        public IReadOnlyDictionary<string, int> ExternalCellCachesBySheetName { get; }

        /// <summary>Gets cached external cell values grouped by value kind.</summary>
        public IReadOnlyDictionary<LegacyXlsCellValueKind, int> ExternalCachedCellsByValueKind { get; }

        /// <summary>Gets preserve-only PivotTable BIFF records grouped by decoded metadata kind.</summary>
        public IReadOnlyDictionary<LegacyXlsPivotTableRecordKind, int> PivotTableRecordsByKind { get; }

        /// <summary>Gets preserve-only PivotTable BIFF records grouped by record name.</summary>
        public IReadOnlyDictionary<string, int> PivotTableRecordsByName { get; }

        /// <summary>Gets decoded SXDI PivotTable data item records grouped by raw aggregation function identifier.</summary>
        public IReadOnlyDictionary<string, int> PivotTableDataItemAggregations { get; }

        /// <summary>Gets decoded SXRng PivotTable grouping records grouped by grouping kind.</summary>
        public IReadOnlyDictionary<string, int> PivotTableGroupingKinds { get; }

        /// <summary>Gets preserve-only chart BIFF records grouped by shallow category.</summary>
        public IReadOnlyDictionary<LegacyXlsChartRecordKind, int> ChartRecordsByKind { get; }

        /// <summary>Gets preserve-only chart BIFF records grouped by record name.</summary>
        public IReadOnlyDictionary<string, int> ChartRecordsByName { get; }

        /// <summary>Gets preserve-only chart BIFF chart-type records grouped by decoded chart family.</summary>
        public IReadOnlyDictionary<string, int> ChartRecordsByChartType { get; }

        /// <summary>Gets Chart records grouped by decoded chart rectangle.</summary>
        public IReadOnlyDictionary<string, int> ChartRecordsByRectangle { get; }

        /// <summary>Gets Axis records grouped by decoded axis type.</summary>
        public IReadOnlyDictionary<string, int> ChartRecordsByAxisType { get; }

        /// <summary>Gets preserve-only chart BIFF records grouped by workbook or sheet location.</summary>
        public IReadOnlyDictionary<string, int> ChartRecordsByLocation { get; }

        /// <summary>Gets preserve-only drawing and object BIFF records grouped by shallow category.</summary>
        public IReadOnlyDictionary<LegacyXlsDrawingRecordKind, int> DrawingRecordsByKind { get; }

        /// <summary>Gets preserve-only drawing and object BIFF records grouped by record name.</summary>
        public IReadOnlyDictionary<string, int> DrawingRecordsByName { get; }

        /// <summary>Gets OBJ records grouped by decoded common-object type identifier.</summary>
        public IReadOnlyDictionary<string, int> DrawingRecordsByObjectType { get; }

        /// <summary>Gets MsoDrawing records grouped by decoded top-level Escher record type.</summary>
        public IReadOnlyDictionary<string, int> DrawingRecordsByEscherRecordType { get; }

        /// <summary>Gets preserve-only drawing and object BIFF records grouped by workbook or sheet location.</summary>
        public IReadOnlyDictionary<string, int> DrawingRecordsByLocation { get; }

        /// <summary>Gets preserve-only compound feature records grouped by kind.</summary>
        public IReadOnlyDictionary<LegacyXlsCompoundFeatureRecordKind, int> CompoundFeatureRecordsByKind { get; }

        /// <summary>Gets matching compound feature entries grouped by feature kind.</summary>
        public IReadOnlyDictionary<LegacyXlsCompoundFeatureRecordKind, int> CompoundFeatureEntriesByKind { get; }

        /// <summary>Gets matching compound feature entries grouped by compound entry path or name.</summary>
        public IReadOnlyDictionary<string, int> CompoundFeatureEntriesByName { get; }

        /// <summary>Gets parsed calculation setting records grouped by setting kind.</summary>
        public IReadOnlyDictionary<LegacyXlsCalculationSettingKind, int> CalculationSettingsByKind { get; }

        /// <summary>Gets parsed workbook cell styles grouped by built-in/custom kind.</summary>
        public IReadOnlyDictionary<string, int> CellStylesByKind { get; }

        /// <summary>Gets parsed workbook metadata records grouped by metadata kind.</summary>
        public IReadOnlyDictionary<LegacyXlsWorkbookMetadataKind, int> WorkbookMetadataRecordsByKind { get; }

        /// <summary>Gets parsed worksheet metadata records grouped by metadata kind.</summary>
        public IReadOnlyDictionary<LegacyXlsWorksheetMetadataKind, int> WorksheetMetadataRecordsByKind { get; }

        /// <summary>Gets parsed unsupported-sheet metadata records grouped by metadata kind.</summary>
        public IReadOnlyDictionary<LegacyXlsUnsupportedSheetMetadataKind, int> UnsupportedSheetMetadataRecordsByKind { get; }

        /// <summary>Gets preserved feature record counts grouped by feature kind.</summary>
        public IReadOnlyDictionary<LegacyXlsUnsupportedFeatureKind, int> PreservedFeatureRecordsByKind { get; }

        /// <summary>Gets preserved feature record counts grouped by kind, code, and stable feature subtype.</summary>
        public IReadOnlyDictionary<string, int> PreservedFeatureRecordsByDetail { get; }

        /// <summary>Gets whether the import produced error diagnostics.</summary>
        public bool HasImportErrors => ErrorCount > 0;

        /// <summary>Gets whether the import discovered unsupported or preserve-only features.</summary>
        public bool HasUnsupportedFeatures => UnsupportedFeatureCount > 0;

        /// <summary>
        /// Returns a compact Markdown summary suitable for corpus snapshots.
        /// </summary>
        public string ToMarkdown() {
            var builder = new StringBuilder();
            builder.AppendLine("# Legacy XLS Import Report");
            builder.AppendLine();
            builder.AppendLine($"Worksheets: {WorksheetCount}");
            builder.AppendLine($"Unsupported sheets: {UnsupportedSheetCount}");
            builder.AppendLine($"Cells: {CellCount}");
            builder.AppendLine($"Formula cells: {FormulaCellCount}");
            builder.AppendLine($"Comments: {CommentCount}");
            builder.AppendLine($"Hyperlinks: {HyperlinkCount}");
            builder.AppendLine($"Data validations: {DataValidationCount}");
            builder.AppendLine($"Conditional formatting rules: {ConditionalFormattingCount}");
            builder.AppendLine($"AutoFilter criteria columns: {AutoFilterCriteriaCount}");
            builder.AppendLine($"Defined names: {DefinedNameCount}");
            builder.AppendLine($"External references: {ExternalReferenceCount}");
            builder.AppendLine($"External sheet names: {ExternalSheetNameCount}");
            builder.AppendLine($"External names: {ExternalNameCount}");
            builder.AppendLine($"External cell caches: {ExternalCellCacheCount}");
            builder.AppendLine($"External cached cells: {ExternalCachedCellCount}");
            builder.AppendLine($"Pivot table records: {PivotTableRecordCount}");
            builder.AppendLine($"Chart records: {ChartRecordCount}");
            builder.AppendLine($"Drawing records: {DrawingRecordCount}");
            builder.AppendLine($"Compound feature records: {CompoundFeatureRecordCount}");
            builder.AppendLine($"Compound feature entries: {CompoundFeatureEntryCount}");
            builder.AppendLine($"Calculation setting records: {CalculationSettingRecordCount}");
            builder.AppendLine($"Cell style records: {CellStyleRecordCount}");
            builder.AppendLine($"Workbook metadata records: {WorkbookMetadataRecordCount}");
            builder.AppendLine($"Worksheet metadata records: {WorksheetMetadataRecordCount}");
            builder.AppendLine($"Unsupported sheet metadata records: {UnsupportedSheetMetadataRecordCount}");
            builder.AppendLine($"Unsupported features: {UnsupportedFeatureCount}");
            builder.AppendLine($"Preserved feature records: {PreservedFeatureRecordCount}");
            builder.AppendLine($"Errors: {ErrorCount}");
            builder.AppendLine($"Warnings: {WarningCount}");
            AppendDictionary(builder, "Diagnostics By Code", DiagnosticsByCode);
            AppendDictionary(builder, "Formula Token Blockers", FormulaTokenBlockers);
            AppendDictionary(builder, "Data Validations By Type", DataValidationsByType);
            AppendDictionary(builder, "Data Validations By Operator", DataValidationsByOperator);
            AppendDictionary(builder, "Conditional Formatting By Type", ConditionalFormattingsByType);
            AppendDictionary(builder, "Conditional Formatting By Operator", ConditionalFormattingsByOperator);
            AppendDictionary(builder, "AutoFilter Criteria By Operator", AutoFilterCriteriaByOperator);
            AppendDictionary(builder, "Unsupported Features By Code", UnsupportedFeaturesByCode);
            AppendDictionary(builder, "Unsupported Features By Kind", UnsupportedFeaturesByKind.ToDictionary(
                entry => entry.Key.ToString(),
                entry => entry.Value,
                StringComparer.OrdinalIgnoreCase));
            AppendDictionary(builder, "Unsupported Feature Record Types", UnsupportedFeaturesByRecordType);
            AppendDictionary(builder, "Unsupported Feature Details", UnsupportedFeaturesByDetail);
            AppendDictionary(builder, "Unsupported Feature Locations", UnsupportedFeaturesByLocation);
            AppendDictionary(builder, "Unsupported Sheets By Kind", UnsupportedSheetsByKind.ToDictionary(
                entry => entry.Key.ToString(),
                entry => entry.Value,
                StringComparer.OrdinalIgnoreCase));
            AppendDictionary(builder, "Unsupported Sheets By Type", UnsupportedSheetsByType);
            AppendDictionary(builder, "Unsupported Sheets By Name", UnsupportedSheetsByName);
            AppendDictionary(builder, "External References By Kind", ExternalReferencesByKind.ToDictionary(
                entry => entry.Key.ToString(),
                entry => entry.Value,
                StringComparer.OrdinalIgnoreCase));
            AppendDictionary(builder, "External References By Target", ExternalReferencesByTarget);
            AppendDictionary(builder, "External Sheet Names By Reference Kind", ExternalSheetNamesByReferenceKind.ToDictionary(
                entry => entry.Key.ToString(),
                entry => entry.Value,
                StringComparer.OrdinalIgnoreCase));
            AppendDictionary(builder, "External Names By Reference Kind", ExternalNamesByReferenceKind.ToDictionary(
                entry => entry.Key.ToString(),
                entry => entry.Value,
                StringComparer.OrdinalIgnoreCase));
            AppendDictionary(builder, "External Names By Name", ExternalNamesByName);
            AppendDictionary(builder, "External Cell Caches By Sheet Name", ExternalCellCachesBySheetName);
            AppendDictionary(builder, "External Cached Cells By Value Kind", ExternalCachedCellsByValueKind.ToDictionary(
                entry => entry.Key.ToString(),
                entry => entry.Value,
                StringComparer.OrdinalIgnoreCase));
            AppendDictionary(builder, "Pivot Table Records By Kind", PivotTableRecordsByKind.ToDictionary(
                entry => entry.Key.ToString(),
                entry => entry.Value,
                StringComparer.OrdinalIgnoreCase));
            AppendDictionary(builder, "Pivot Table Records By Name", PivotTableRecordsByName);
            AppendDictionary(builder, "Pivot Table Data Item Aggregations", PivotTableDataItemAggregations);
            AppendDictionary(builder, "Pivot Table Grouping Kinds", PivotTableGroupingKinds);
            AppendDictionary(builder, "Chart Records By Kind", ChartRecordsByKind.ToDictionary(
                entry => entry.Key.ToString(),
                entry => entry.Value,
                StringComparer.OrdinalIgnoreCase));
            AppendDictionary(builder, "Chart Records By Name", ChartRecordsByName);
            AppendDictionary(builder, "Chart Records By Chart Type", ChartRecordsByChartType);
            AppendDictionary(builder, "Chart Records By Rectangle", ChartRecordsByRectangle);
            AppendDictionary(builder, "Chart Records By Axis Type", ChartRecordsByAxisType);
            AppendDictionary(builder, "Chart Records By Location", ChartRecordsByLocation);
            AppendDictionary(builder, "Drawing Records By Kind", DrawingRecordsByKind.ToDictionary(
                entry => entry.Key.ToString(),
                entry => entry.Value,
                StringComparer.OrdinalIgnoreCase));
            AppendDictionary(builder, "Drawing Records By Name", DrawingRecordsByName);
            AppendDictionary(builder, "Drawing Records By Object Type", DrawingRecordsByObjectType);
            AppendDictionary(builder, "Drawing Records By Escher Record Type", DrawingRecordsByEscherRecordType);
            AppendDictionary(builder, "Drawing Records By Location", DrawingRecordsByLocation);
            AppendDictionary(builder, "Compound Feature Records By Kind", CompoundFeatureRecordsByKind.ToDictionary(
                entry => entry.Key.ToString(),
                entry => entry.Value,
                StringComparer.OrdinalIgnoreCase));
            AppendDictionary(builder, "Compound Feature Entries By Kind", CompoundFeatureEntriesByKind.ToDictionary(
                entry => entry.Key.ToString(),
                entry => entry.Value,
                StringComparer.OrdinalIgnoreCase));
            AppendDictionary(builder, "Compound Feature Entries By Name", CompoundFeatureEntriesByName);
            AppendDictionary(builder, "Calculation Settings By Kind", CalculationSettingsByKind.ToDictionary(
                entry => entry.Key.ToString(),
                entry => entry.Value,
                StringComparer.OrdinalIgnoreCase));
            AppendDictionary(builder, "Cell Styles By Kind", CellStylesByKind);
            AppendDictionary(builder, "Workbook Metadata Records By Kind", WorkbookMetadataRecordsByKind.ToDictionary(
                entry => entry.Key.ToString(),
                entry => entry.Value,
                StringComparer.OrdinalIgnoreCase));
            AppendDictionary(builder, "Worksheet Metadata Records By Kind", WorksheetMetadataRecordsByKind.ToDictionary(
                entry => entry.Key.ToString(),
                entry => entry.Value,
                StringComparer.OrdinalIgnoreCase));
            AppendDictionary(builder, "Unsupported Sheet Metadata Records By Kind", UnsupportedSheetMetadataRecordsByKind.ToDictionary(
                entry => entry.Key.ToString(),
                entry => entry.Value,
                StringComparer.OrdinalIgnoreCase));
            AppendDictionary(builder, "Preserved Feature Records By Kind", PreservedFeatureRecordsByKind.ToDictionary(
                entry => entry.Key.ToString(),
                entry => entry.Value,
                StringComparer.OrdinalIgnoreCase));
            AppendDictionary(builder, "Preserved Feature Record Details", PreservedFeatureRecordsByDetail);
            return builder.ToString();
        }

        private static IReadOnlyDictionary<string, int> CountByCode(IEnumerable<string> codes) {
            return codes
                .Where(code => !string.IsNullOrWhiteSpace(code))
                .GroupBy(code => code, StringComparer.OrdinalIgnoreCase)
                .OrderBy(group => group.Key, StringComparer.OrdinalIgnoreCase)
                .ToDictionary(group => group.Key, group => group.Count(), StringComparer.OrdinalIgnoreCase);
        }

        private static IReadOnlyDictionary<LegacyXlsUnsupportedFeatureKind, int> CountByKind(IEnumerable<LegacyXlsUnsupportedFeature> features) {
            return features
                .GroupBy(feature => feature.Kind)
                .OrderBy(group => group.Key.ToString(), StringComparer.OrdinalIgnoreCase)
                .ToDictionary(group => group.Key, group => group.Count());
        }

        private static IReadOnlyDictionary<LegacyXlsUnsupportedFeatureKind, int> CountPreservedRecordsByKind(IEnumerable<LegacyXlsPreservedFeatureRecord> records) {
            return records
                .GroupBy(record => record.Kind)
                .OrderBy(group => group.Key.ToString(), StringComparer.OrdinalIgnoreCase)
                .ToDictionary(group => group.Key, group => group.Count());
        }

        private static IReadOnlyDictionary<LegacyXlsUnsupportedSheetKind, int> CountUnsupportedSheetsByKind(IEnumerable<LegacyXlsUnsupportedSheet> sheets) {
            return sheets
                .GroupBy(sheet => sheet.Kind)
                .OrderBy(group => group.Key.ToString(), StringComparer.OrdinalIgnoreCase)
                .ToDictionary(group => group.Key, group => group.Count());
        }

        private static IReadOnlyDictionary<LegacyXlsExternalReferenceKind, int> CountExternalReferencesByKind(IEnumerable<LegacyXlsExternalReference> references) {
            return references
                .GroupBy(reference => reference.Kind)
                .OrderBy(group => group.Key.ToString(), StringComparer.OrdinalIgnoreCase)
                .ToDictionary(group => group.Key, group => group.Count());
        }

        private static IReadOnlyDictionary<LegacyXlsExternalReferenceKind, int> CountExternalSheetNamesByReferenceKind(IEnumerable<LegacyXlsExternalReference> references) {
            return references
                .Where(reference => reference.SheetNames.Count > 0)
                .GroupBy(reference => reference.Kind)
                .OrderBy(group => group.Key.ToString(), StringComparer.OrdinalIgnoreCase)
                .ToDictionary(group => group.Key, group => group.Sum(reference => reference.SheetNames.Count));
        }

        private static IReadOnlyDictionary<LegacyXlsExternalReferenceKind, int> CountExternalNamesByReferenceKind(IEnumerable<LegacyXlsExternalReference> references) {
            return references
                .Where(reference => reference.ExternalNames.Count > 0)
                .GroupBy(reference => reference.Kind)
                .OrderBy(group => group.Key.ToString(), StringComparer.OrdinalIgnoreCase)
                .ToDictionary(group => group.Key, group => group.Sum(reference => reference.ExternalNames.Count));
        }

        private static IReadOnlyDictionary<LegacyXlsCellValueKind, int> CountExternalCachedCellsByValueKind(IEnumerable<LegacyXlsExternalReference> references) {
            return references
                .SelectMany(reference => reference.CachedCellCaches)
                .SelectMany(cache => cache.Cells)
                .GroupBy(cell => cell.Kind)
                .OrderBy(group => group.Key.ToString(), StringComparer.OrdinalIgnoreCase)
                .ToDictionary(group => group.Key, group => group.Count());
        }

        private static IReadOnlyDictionary<LegacyXlsCalculationSettingKind, int> CountCalculationSettingsByKind(IEnumerable<LegacyXlsCalculationSettingRecord> records) {
            return records
                .GroupBy(record => record.Kind)
                .OrderBy(group => group.Key.ToString(), StringComparer.OrdinalIgnoreCase)
                .ToDictionary(group => group.Key, group => group.Count());
        }

        private static IReadOnlyDictionary<LegacyXlsPivotTableRecordKind, int> CountPivotTableRecordsByKind(IEnumerable<LegacyXlsPivotTableRecord> records) {
            return records
                .GroupBy(record => record.Kind)
                .OrderBy(group => group.Key.ToString(), StringComparer.OrdinalIgnoreCase)
                .ToDictionary(group => group.Key, group => group.Count());
        }

        private static IReadOnlyDictionary<LegacyXlsChartRecordKind, int> CountChartRecordsByKind(IEnumerable<LegacyXlsChartRecord> records) {
            return records
                .GroupBy(record => record.Kind)
                .OrderBy(group => group.Key.ToString(), StringComparer.OrdinalIgnoreCase)
                .ToDictionary(group => group.Key, group => group.Count());
        }

        private static IReadOnlyDictionary<LegacyXlsDrawingRecordKind, int> CountDrawingRecordsByKind(IEnumerable<LegacyXlsDrawingRecord> records) {
            return records
                .GroupBy(record => record.Kind)
                .OrderBy(group => group.Key.ToString(), StringComparer.OrdinalIgnoreCase)
                .ToDictionary(group => group.Key, group => group.Count());
        }

        private static IReadOnlyDictionary<LegacyXlsCompoundFeatureRecordKind, int> CountCompoundFeatureRecordsByKind(IEnumerable<LegacyXlsCompoundFeatureRecord> records) {
            return records
                .GroupBy(record => record.Kind)
                .OrderBy(group => group.Key.ToString(), StringComparer.OrdinalIgnoreCase)
                .ToDictionary(group => group.Key, group => group.Count());
        }

        private static IReadOnlyDictionary<LegacyXlsCompoundFeatureRecordKind, int> CountCompoundFeatureEntriesByKind(IEnumerable<LegacyXlsCompoundFeatureRecord> records) {
            return records
                .GroupBy(record => record.Kind)
                .OrderBy(group => group.Key.ToString(), StringComparer.OrdinalIgnoreCase)
                .ToDictionary(group => group.Key, group => group.Sum(record => record.Entries.Count));
        }

        private static IReadOnlyDictionary<LegacyXlsWorkbookMetadataKind, int> CountWorkbookMetadataRecordsByKind(IEnumerable<LegacyXlsWorkbookMetadataRecord> records) {
            return records
                .GroupBy(record => record.Kind)
                .OrderBy(group => group.Key.ToString(), StringComparer.OrdinalIgnoreCase)
                .ToDictionary(group => group.Key, group => group.Count());
        }

        private static IReadOnlyDictionary<LegacyXlsWorksheetMetadataKind, int> CountWorksheetMetadataRecordsByKind(IEnumerable<LegacyXlsWorksheetMetadataRecord> records) {
            return records
                .GroupBy(record => record.Kind)
                .OrderBy(group => group.Key.ToString(), StringComparer.OrdinalIgnoreCase)
                .ToDictionary(group => group.Key, group => group.Count());
        }

        private static IReadOnlyDictionary<LegacyXlsUnsupportedSheetMetadataKind, int> CountUnsupportedSheetMetadataRecordsByKind(IEnumerable<LegacyXlsUnsupportedSheetMetadataRecord> records) {
            return records
                .GroupBy(record => record.Kind)
                .OrderBy(group => group.Key.ToString(), StringComparer.OrdinalIgnoreCase)
                .ToDictionary(group => group.Key, group => group.Count());
        }

        private static string GetFeatureLocationKey(LegacyXlsUnsupportedFeature feature) {
            string location = string.IsNullOrWhiteSpace(feature.SheetName) ? "(workbook)" : feature.SheetName!;
            return feature.Code + "|" + location;
        }

        private static string GetChartRecordLocationKey(LegacyXlsChartRecord record) {
            return string.IsNullOrWhiteSpace(record.SheetName) ? "(workbook)" : record.SheetName!;
        }

        private static string GetDrawingRecordLocationKey(LegacyXlsDrawingRecord record) {
            return string.IsNullOrWhiteSpace(record.SheetName) ? "(workbook)" : record.SheetName!;
        }

        private static string GetExternalReferenceTargetKey(LegacyXlsExternalReference reference) {
            return string.IsNullOrWhiteSpace(reference.Target) ? $"({reference.Kind})" : EscapeControlCharacters(reference.Target!);
        }

        private static string GetExternalCellCacheSheetKey(LegacyXlsExternalCellCache cache) {
            if (!string.IsNullOrWhiteSpace(cache.SheetName)) {
                return cache.SheetName!;
            }

            return cache.SheetIndex.HasValue ? $"SheetIndex:{cache.SheetIndex.Value}" : "(unknown)";
        }

        private static void AppendDictionary(StringBuilder builder, string title, IReadOnlyDictionary<string, int> values) {
            if (values.Count == 0) {
                return;
            }

            builder.AppendLine();
            builder.AppendLine("## " + title);
            builder.AppendLine();
            builder.AppendLine("| Key | Count |");
            builder.AppendLine("| --- | --- |");
            foreach (KeyValuePair<string, int> entry in values) {
                builder.Append("| ");
                builder.Append(EscapeMarkdownCell(entry.Key));
                builder.Append(" | ");
                builder.Append(entry.Value);
                builder.AppendLine(" |");
            }
        }

        private static string EscapeMarkdownCell(string value) {
            return value.Replace("\\", "\\\\").Replace("|", "\\|").Replace("\r", " ").Replace("\n", " ");
        }

        private static string EscapeControlCharacters(string value) {
            var builder = new StringBuilder(value.Length);
            foreach (char character in value) {
                if (!char.IsControl(character)) {
                    builder.Append(character);
                } else if (character <= 0xFF) {
                    builder.Append("\\x");
                    builder.Append(((int)character).ToString("X2"));
                } else {
                    builder.Append("\\u");
                    builder.Append(((int)character).ToString("X4"));
                }
            }

            return builder.ToString();
        }
    }
}
