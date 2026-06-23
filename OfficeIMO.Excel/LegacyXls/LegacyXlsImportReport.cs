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
            DataValidationsByErrorStyle = CountByCode(workbook.Worksheets.SelectMany(sheet => sheet.DataValidations).Select(validation => validation.ErrorStyle.ToString()));
            DataValidationListSourcesByKind = CountByCode(workbook.Worksheets
                .SelectMany(sheet => sheet.DataValidations)
                .Where(validation => validation.Type == LegacyXlsDataValidationType.List)
                .Select(validation => validation.ListSourceKind.ToString()));
            ConditionalFormattingsByType = CountByCode(workbook.Worksheets.SelectMany(sheet => sheet.ConditionalFormattings).Select(formatting => formatting.Type.ToString()));
            ConditionalFormattingsByOperator = CountByCode(workbook.Worksheets
                .SelectMany(sheet => sheet.ConditionalFormattings)
                .Where(formatting => formatting.Operator.HasValue)
                .Select(formatting => formatting.Operator!.Value.ToString()));
            AutoFilterCriteriaByOperator = CountByCode(workbook.Worksheets
                .SelectMany(sheet => sheet.AutoFilterCriteria)
                .SelectMany(criteria => criteria.Conditions)
                .Select(condition => condition.Operator.ToString()));
            AutoFilterCriteriaByValueKind = CountByCode(workbook.Worksheets
                .SelectMany(sheet => sheet.AutoFilterCriteria)
                .SelectMany(criteria => criteria.Conditions)
                .Select(condition => condition.ValueKind.ToString()));
            AutoFilterCriteriaByJoinOperator = CountByCode(workbook.Worksheets
                .SelectMany(sheet => sheet.AutoFilterCriteria)
                .Select(criteria => criteria.JoinOperator.ToString()));
            AutoFilterCriteriaByKind = CountByCode(workbook.Worksheets
                .SelectMany(sheet => sheet.AutoFilterCriteria)
                .Select(criteria => criteria.Kind.ToString()));
            AutoFilterTop10Kinds = CountByCode(workbook.Worksheets
                .SelectMany(sheet => sheet.AutoFilterCriteria)
                .Where(criteria => criteria.IsTop10)
                .Select(GetAutoFilterTop10KindKey));
            AutoFilterTop10Values = CountByCode(workbook.Worksheets
                .SelectMany(sheet => sheet.AutoFilterCriteria)
                .Where(criteria => criteria.IsTop10 && criteria.Top10Value.HasValue)
                .Select(criteria => $"{GetAutoFilterTop10KindKey(criteria)}:{criteria.Top10Value!.Value}"));
            DefinedNameCount = workbook.DefinedNames.Count;
            ExternalReferenceCount = workbook.ExternalReferences.Count;
            ExternalSheetNameCount = workbook.ExternalReferences.Sum(reference => reference.SheetNames.Count);
            ExternalNameCount = workbook.ExternalReferences.Sum(reference => reference.ExternalNames.Count);
            ExternalCellCacheCount = workbook.ExternalReferences.Sum(reference => reference.CachedCellCaches.Count);
            ExternalCachedCellCount = workbook.ExternalReferences.Sum(reference => reference.CachedCellCaches.Sum(cache => cache.Cells.Count));
            PivotTableRecordCount = workbook.PivotTableRecords.Count;
            ChartRecordCount = workbook.ChartRecords.Count;
            DrawingRecordCount = workbook.DrawingRecords.Count;
            DifferentialFormatCount = workbook.DifferentialFormats.Count;
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
            FormulaTokenBlockersByToken = CountByCode(workbook.Diagnostics
                .Where(diagnostic => string.Equals(diagnostic.Code, "XLS-BIFF-FORMULA-TOKENS-UNSUPPORTED", StringComparison.OrdinalIgnoreCase))
                .Where(diagnostic => diagnostic.FormulaToken.HasValue)
                .Select(diagnostic => $"Token:0x{diagnostic.FormulaToken!.Value:X2}"));
            FormulaTokenBlockersByTokenName = CountByCode(workbook.Diagnostics
                .Where(diagnostic => string.Equals(diagnostic.Code, "XLS-BIFF-FORMULA-TOKENS-UNSUPPORTED", StringComparison.OrdinalIgnoreCase))
                .Where(diagnostic => !string.IsNullOrWhiteSpace(diagnostic.FormulaTokenName))
                .Select(diagnostic => diagnostic.FormulaTokenName!));
            FormulaTokenBlockersByOffset = CountByCode(workbook.Diagnostics
                .Where(diagnostic => string.Equals(diagnostic.Code, "XLS-BIFF-FORMULA-TOKENS-UNSUPPORTED", StringComparison.OrdinalIgnoreCase))
                .Where(diagnostic => diagnostic.FormulaTokenOffset.HasValue)
                .Select(diagnostic => $"Offset:{diagnostic.FormulaTokenOffset!.Value}"));
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
            FileFormatBlockers = CountByCode(workbook.UnsupportedFeatures
                .Where(feature => feature.Kind == LegacyXlsUnsupportedFeatureKind.EncryptedWorkbook
                    || feature.Kind == LegacyXlsUnsupportedFeatureKind.UnsupportedBiffVersion)
                .Select(feature => $"{feature.Kind}|{feature.DetailCode ?? feature.Code}"));
            EncryptedWorkbooksByMethod = CountByCode(workbook.UnsupportedFeatures
                .Where(feature => feature.Kind == LegacyXlsUnsupportedFeatureKind.EncryptedWorkbook)
                .Select(GetEncryptionMethodKey));
            UnsupportedBiffVersionsByVersion = CountByCode(workbook.UnsupportedFeatures
                .Where(feature => feature.Kind == LegacyXlsUnsupportedFeatureKind.UnsupportedBiffVersion)
                .Select(GetBiffVersionKey));
            UnsupportedBiffVersionsBySubstream = CountByCode(workbook.UnsupportedFeatures
                .Where(feature => feature.Kind == LegacyXlsUnsupportedFeatureKind.UnsupportedBiffVersion)
                .Select(GetBiffSubstreamKey));
            UnsupportedSheetsByKind = CountUnsupportedSheetsByKind(workbook.UnsupportedSheets);
            UnsupportedSheetsByType = CountByCode(workbook.UnsupportedSheets.Select(sheet => $"0x{sheet.SheetType:X2}|{sheet.Kind}"));
            UnsupportedSheetsByName = CountByCode(workbook.UnsupportedSheets.Select(sheet => sheet.Name));
            UnsupportedChartSheetPrintSizes = CountByCode(workbook.UnsupportedSheets
                .Where(sheet => sheet.Kind == LegacyXlsUnsupportedSheetKind.ChartSheet && sheet.ChartPrintSize.HasValue)
                .Select(sheet => $"PrintSize:{sheet.ChartPrintSize!.Value}"));
            UnsupportedChartSheetTextObjectCounts = CountByCode(workbook.UnsupportedSheets
                .Where(sheet => sheet.Kind == LegacyXlsUnsupportedSheetKind.ChartSheet && sheet.ChartTextObjectCount > 0)
                .Select(sheet => $"TextObjects:{sheet.ChartTextObjectCount}"));
            ExternalReferencesByKind = CountExternalReferencesByKind(workbook.ExternalReferences);
            ExternalReferencesByTarget = CountByCode(workbook.ExternalReferences.Select(GetExternalReferenceTargetKey));
            ExternalSheetNamesByReferenceKind = CountExternalSheetNamesByReferenceKind(workbook.ExternalReferences);
            ExternalNamesByReferenceKind = CountExternalNamesByReferenceKind(workbook.ExternalReferences);
            ExternalNamesByName = CountByCode(workbook.ExternalReferences.SelectMany(reference => reference.ExternalNames.Select(name => name.Name)));
            ExternalNamesByScope = CountByCode(workbook.ExternalReferences
                .SelectMany(reference => reference.ExternalNames)
                .Select(name => name.LocalSheetIndex.HasValue ? "SheetLocal" : "Workbook"));
            ExternalNamesByBuiltInState = CountByCode(workbook.ExternalReferences
                .SelectMany(reference => reference.ExternalNames)
                .Select(name => name.BuiltIn ? "BuiltIn" : "Custom"));
            ExternalCellCachesBySheetName = CountByCode(workbook.ExternalReferences.SelectMany(reference => reference.CachedCellCaches.Select(GetExternalCellCacheSheetKey)));
            ExternalCachedCellsByValueKind = CountExternalCachedCellsByValueKind(workbook.ExternalReferences);
            PivotTableRecordsByKind = CountPivotTableRecordsByKind(workbook.PivotTableRecords);
            PivotTableRecordsByName = CountByCode(workbook.PivotTableRecords.Select(record => record.RecordName));
            PivotTableDataItemAggregations = CountByCode(workbook.PivotTableRecords
                .Where(record => record.AggregationFunction.HasValue)
                .Select(record => $"AggregationFunction:{record.AggregationFunction!.Value}"));
            PivotTableDataItemAggregationKinds = CountByCode(workbook.PivotTableRecords
                .Where(record => !string.IsNullOrWhiteSpace(record.AggregationFunctionName))
                .Select(record => record.AggregationFunctionName!));
            PivotTableDataItemDisplayCalculations = CountByCode(workbook.PivotTableRecords
                .Where(record => !string.IsNullOrWhiteSpace(record.DisplayCalculationName))
                .Select(record => record.DisplayCalculationName!));
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
            ChartRecordsByAxesUsedCount = CountByCode(workbook.ChartRecords
                .Where(record => record.AxesUsedCount.HasValue)
                .Select(record => $"AxesUsed:{record.AxesUsedCount!.Value}"));
            ChartSeriesCategoryDataTypes = CountByCode(workbook.ChartRecords
                .Where(record => !string.IsNullOrWhiteSpace(record.SeriesCategoryDataTypeName))
                .Select(record => record.SeriesCategoryDataTypeName!));
            ChartSeriesValueCounts = CountByCode(workbook.ChartRecords
                .Where(record => record.SeriesCategoryCount.HasValue && record.SeriesValueCount.HasValue && record.SeriesBubbleSizeCount.HasValue)
                .Select(record => $"Categories:{record.SeriesCategoryCount!.Value};Values:{record.SeriesValueCount!.Value};BubbleSizes:{record.SeriesBubbleSizeCount!.Value}"));
            ChartDataFormatTargets = CountByCode(workbook.ChartRecords
                .Where(record => !string.IsNullOrWhiteSpace(record.DataFormatTarget))
                .Select(record => record.DataFormatTarget!));
            ChartDataFormatSeriesIndexes = CountByCode(workbook.ChartRecords
                .Where(record => record.DataFormatSeriesIndex.HasValue)
                .Select(record => $"SeriesIndex:{record.DataFormatSeriesIndex!.Value}"));
            ChartLineFormatStyles = CountByCode(workbook.ChartRecords
                .Where(record => record.LineFormat != null)
                .Select(record => record.LineFormat!.StyleName));
            ChartLineFormatWeights = CountByCode(workbook.ChartRecords
                .Where(record => record.LineFormat != null)
                .Select(record => record.LineFormat!.WeightName));
            ChartAreaFormatPatterns = CountByCode(workbook.ChartRecords
                .Where(record => record.AreaFormat != null)
                .Select(record => record.AreaFormat!.PatternName));
            ChartMarkerFormatTypes = CountByCode(workbook.ChartRecords
                .Where(record => record.MarkerFormat != null)
                .Select(record => record.MarkerFormat!.MarkerTypeName));
            ChartMarkerFormatSizes = CountByCode(workbook.ChartRecords
                .Where(record => record.MarkerFormat != null)
                .Select(record => $"SizeTwips:{record.MarkerFormat!.SizeTwips}"));
            ChartDefaultTextTargets = CountByCode(workbook.ChartRecords
                .Where(record => !string.IsNullOrWhiteSpace(record.DefaultTextTargetName))
                .Select(record => record.DefaultTextTargetName!));
            ChartTextHorizontalAlignments = CountByCode(workbook.ChartRecords
                .Where(record => record.Text != null)
                .Select(record => record.Text!.HorizontalAlignmentName));
            ChartTextVerticalAlignments = CountByCode(workbook.ChartRecords
                .Where(record => record.Text != null)
                .Select(record => record.Text!.VerticalAlignmentName));
            ChartTextDataLabelPositions = CountByCode(workbook.ChartRecords
                .Where(record => record.Text != null)
                .Select(record => record.Text!.DataLabelPositionName));
            ChartTextFlags = CountByCode(workbook.ChartRecords
                .Where(record => record.Text != null)
                .SelectMany(record => record.Text!.FlagNames));
            ChartObjectLinkTargets = CountByCode(workbook.ChartRecords
                .Where(record => record.ObjectLink != null)
                .Select(record => record.ObjectLink!.LinkedObjectName));
            ChartLegendLayouts = CountByCode(workbook.ChartRecords
                .Where(record => record.Legend != null)
                .Select(record => record.Legend!.Vertical ? "Vertical" : "MultiColumnOrManual"));
            ChartTickMajorLocations = CountByCode(workbook.ChartRecords
                .Where(record => record.Tick != null)
                .Select(record => record.Tick!.MajorTickLocationName));
            ChartTickLabelLocations = CountByCode(workbook.ChartRecords
                .Where(record => record.Tick != null)
                .Select(record => record.Tick!.LabelLocationName));
            ChartRecordsByLocation = CountByCode(workbook.ChartRecords.Select(GetChartRecordLocationKey));
            DrawingRecordsByKind = CountDrawingRecordsByKind(workbook.DrawingRecords);
            DrawingRecordsByName = CountByCode(workbook.DrawingRecords.Select(record => record.RecordName));
            DrawingRecordsByObjectType = CountByCode(workbook.DrawingRecords
                .Where(record => record.ObjectType.HasValue)
                .Select(record => $"ObjectType:0x{record.ObjectType!.Value:X4}"));
            DrawingRecordsByObjectTypeName = CountByCode(workbook.DrawingRecords
                .Where(record => !string.IsNullOrWhiteSpace(record.ObjectTypeName))
                .Select(record => record.ObjectTypeName!));
            DrawingRecordsByObjectFlags = CountByCode(workbook.DrawingRecords
                .Where(record => record.ObjectFlags.HasValue)
                .Select(record => $"ObjectFlags:0x{record.ObjectFlags!.Value:X4}"));
            DrawingRecordsByObjectFlagName = CountByCode(workbook.DrawingRecords
                .SelectMany(record => record.ObjectFlagNames));
            DrawingRecordsByEscherRecordType = CountByCode(workbook.DrawingRecords
                .Where(record => record.EscherRecordType.HasValue)
                .Select(record => $"EscherRecordType:0x{record.EscherRecordType!.Value:X4}"));
            DrawingRecordsByEscherRecordTypeName = CountByCode(workbook.DrawingRecords
                .Where(record => !string.IsNullOrWhiteSpace(record.EscherRecordTypeName))
                .Select(record => record.EscherRecordTypeName!));
            DrawingBlipStoreEntriesByType = CountByCode(workbook.DrawingRecords
                .SelectMany(record => record.BlipStoreEntries)
                .Select(entry => entry.RecordInstanceBlipTypeName));
            DrawingBlipStoreEntriesByEmbeddedRecordType = CountByCode(workbook.DrawingRecords
                .SelectMany(record => record.BlipStoreEntries)
                .Where(entry => !string.IsNullOrWhiteSpace(entry.EmbeddedBlipRecordTypeName))
                .Select(entry => entry.EmbeddedBlipRecordTypeName!));
            DrawingBlipStoreEntriesBySize = CountByCode(workbook.DrawingRecords
                .SelectMany(record => record.BlipStoreEntries)
                .Where(entry => entry.SizeBytes.HasValue)
                .Select(entry => $"SizeBytes:{entry.SizeBytes!.Value}"));
            DrawingBlipStoreEntriesByReferenceCount = CountByCode(workbook.DrawingRecords
                .SelectMany(record => record.BlipStoreEntries)
                .Where(entry => entry.ReferenceCount.HasValue)
                .Select(entry => $"References:{entry.ReferenceCount!.Value}"));
            DrawingRecordsByLocation = CountByCode(workbook.DrawingRecords.Select(GetDrawingRecordLocationKey));
            CompoundFeatureRecordsByKind = CountCompoundFeatureRecordsByKind(workbook.CompoundFeatureRecords);
            CompoundFeatureEntriesByKind = CountCompoundFeatureEntriesByKind(workbook.CompoundFeatureRecords);
            CompoundFeatureEntriesByName = CountByCode(workbook.CompoundFeatureRecords.SelectMany(record => record.Entries));
            CompoundFeatureEntriesByRole = CountByCode(workbook.CompoundFeatureRecords
                .SelectMany(record => record.EntryRoles.Values)
                .Select(role => role.ToString()));
            CompoundFeatureEntriesByKindAndRole = CountByCode(workbook.CompoundFeatureRecords
                .SelectMany(record => record.EntryRoles.Values.Select(role => $"{record.Kind}|{role}")));
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

        /// <summary>Gets imported data validations grouped by error alert style.</summary>
        public IReadOnlyDictionary<string, int> DataValidationsByErrorStyle { get; }

        /// <summary>Gets imported list data validations grouped by source shape.</summary>
        public IReadOnlyDictionary<string, int> DataValidationListSourcesByKind { get; }

        /// <summary>Gets imported conditional formatting rules grouped by rule type.</summary>
        public IReadOnlyDictionary<string, int> ConditionalFormattingsByType { get; }

        /// <summary>Gets imported conditional formatting cell-is rules grouped by comparison operator.</summary>
        public IReadOnlyDictionary<string, int> ConditionalFormattingsByOperator { get; }

        /// <summary>Gets imported AutoFilter conditions grouped by comparison operator.</summary>
        public IReadOnlyDictionary<string, int> AutoFilterCriteriaByOperator { get; }

        /// <summary>Gets imported AutoFilter conditions grouped by BIFF operand kind.</summary>
        public IReadOnlyDictionary<string, int> AutoFilterCriteriaByValueKind { get; }

        /// <summary>Gets imported AutoFilter criteria grouped by condition join operator.</summary>
        public IReadOnlyDictionary<string, int> AutoFilterCriteriaByJoinOperator { get; }

        /// <summary>Gets imported AutoFilter criteria grouped by criteria kind.</summary>
        public IReadOnlyDictionary<string, int> AutoFilterCriteriaByKind { get; }

        /// <summary>Gets imported Top/Bottom AutoFilter criteria grouped by top/bottom and items/percent shape.</summary>
        public IReadOnlyDictionary<string, int> AutoFilterTop10Kinds { get; }

        /// <summary>Gets imported Top/Bottom AutoFilter criteria grouped by shape and value.</summary>
        public IReadOnlyDictionary<string, int> AutoFilterTop10Values { get; }

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

        /// <summary>Gets the number of parsed differential formats discovered during import.</summary>
        public int DifferentialFormatCount { get; }

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

        /// <summary>Gets unsupported formula token blockers grouped by raw formula token byte.</summary>
        public IReadOnlyDictionary<string, int> FormulaTokenBlockersByToken { get; }

        /// <summary>Gets unsupported formula token blockers grouped by BIFF parsed-formula token name.</summary>
        public IReadOnlyDictionary<string, int> FormulaTokenBlockersByTokenName { get; }

        /// <summary>Gets unsupported formula token blockers grouped by zero-based parsed-expression token offset.</summary>
        public IReadOnlyDictionary<string, int> FormulaTokenBlockersByOffset { get; }

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

        /// <summary>Gets hard file-format blockers grouped by kind and detail.</summary>
        public IReadOnlyDictionary<string, int> FileFormatBlockers { get; }

        /// <summary>Gets encrypted workbook blockers grouped by FilePass encryption method.</summary>
        public IReadOnlyDictionary<string, int> EncryptedWorkbooksByMethod { get; }

        /// <summary>Gets unsupported BIFF blockers grouped by BIFF version.</summary>
        public IReadOnlyDictionary<string, int> UnsupportedBiffVersionsByVersion { get; }

        /// <summary>Gets unsupported BIFF blockers grouped by BOF substream.</summary>
        public IReadOnlyDictionary<string, int> UnsupportedBiffVersionsBySubstream { get; }

        /// <summary>Gets unsupported sheet entries grouped by decoded sheet kind.</summary>
        public IReadOnlyDictionary<LegacyXlsUnsupportedSheetKind, int> UnsupportedSheetsByKind { get; }

        /// <summary>Gets unsupported sheet entries grouped by raw BoundSheet type and decoded kind.</summary>
        public IReadOnlyDictionary<string, int> UnsupportedSheetsByType { get; }

        /// <summary>Gets unsupported sheet entries grouped by sheet name.</summary>
        public IReadOnlyDictionary<string, int> UnsupportedSheetsByName { get; }

        /// <summary>Gets unsupported chart sheets grouped by raw PrintSize value.</summary>
        public IReadOnlyDictionary<string, int> UnsupportedChartSheetPrintSizes { get; }

        /// <summary>Gets unsupported chart sheets grouped by chart text object count.</summary>
        public IReadOnlyDictionary<string, int> UnsupportedChartSheetTextObjectCounts { get; }

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

        /// <summary>Gets external defined names grouped by workbook or sheet-local scope.</summary>
        public IReadOnlyDictionary<string, int> ExternalNamesByScope { get; }

        /// <summary>Gets external defined names grouped by built-in or custom state.</summary>
        public IReadOnlyDictionary<string, int> ExternalNamesByBuiltInState { get; }

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

        /// <summary>Gets decoded SXDI PivotTable data item records grouped by aggregation function name.</summary>
        public IReadOnlyDictionary<string, int> PivotTableDataItemAggregationKinds { get; }

        /// <summary>Gets decoded SXDI PivotTable data item records grouped by display calculation name.</summary>
        public IReadOnlyDictionary<string, int> PivotTableDataItemDisplayCalculations { get; }

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

        /// <summary>Gets AxesUsed records grouped by decoded axis group count.</summary>
        public IReadOnlyDictionary<string, int> ChartRecordsByAxesUsedCount { get; }

        /// <summary>Gets Series records grouped by decoded category data type.</summary>
        public IReadOnlyDictionary<string, int> ChartSeriesCategoryDataTypes { get; }

        /// <summary>Gets Series records grouped by category, value, and bubble-size counts.</summary>
        public IReadOnlyDictionary<string, int> ChartSeriesValueCounts { get; }

        /// <summary>Gets DataFormat records grouped by whether formatting targets a series or point.</summary>
        public IReadOnlyDictionary<string, int> ChartDataFormatTargets { get; }

        /// <summary>Gets DataFormat records grouped by raw series index.</summary>
        public IReadOnlyDictionary<string, int> ChartDataFormatSeriesIndexes { get; }

        /// <summary>Gets LineFormat records grouped by decoded line style.</summary>
        public IReadOnlyDictionary<string, int> ChartLineFormatStyles { get; }

        /// <summary>Gets LineFormat records grouped by decoded line weight.</summary>
        public IReadOnlyDictionary<string, int> ChartLineFormatWeights { get; }

        /// <summary>Gets AreaFormat records grouped by decoded fill pattern.</summary>
        public IReadOnlyDictionary<string, int> ChartAreaFormatPatterns { get; }

        /// <summary>Gets MarkerFormat records grouped by decoded marker type.</summary>
        public IReadOnlyDictionary<string, int> ChartMarkerFormatTypes { get; }

        /// <summary>Gets MarkerFormat records grouped by marker size in twips.</summary>
        public IReadOnlyDictionary<string, int> ChartMarkerFormatSizes { get; }

        /// <summary>Gets DefaultText records grouped by decoded target scope.</summary>
        public IReadOnlyDictionary<string, int> ChartDefaultTextTargets { get; }

        /// <summary>Gets Text records grouped by decoded horizontal alignment.</summary>
        public IReadOnlyDictionary<string, int> ChartTextHorizontalAlignments { get; }

        /// <summary>Gets Text records grouped by decoded vertical alignment.</summary>
        public IReadOnlyDictionary<string, int> ChartTextVerticalAlignments { get; }

        /// <summary>Gets Text records grouped by decoded data-label position.</summary>
        public IReadOnlyDictionary<string, int> ChartTextDataLabelPositions { get; }

        /// <summary>Gets Text records grouped by decoded flag name.</summary>
        public IReadOnlyDictionary<string, int> ChartTextFlags { get; }

        /// <summary>Gets ObjectLink records grouped by decoded linked chart object.</summary>
        public IReadOnlyDictionary<string, int> ChartObjectLinkTargets { get; }

        /// <summary>Gets Legend records grouped by decoded layout.</summary>
        public IReadOnlyDictionary<string, int> ChartLegendLayouts { get; }

        /// <summary>Gets Tick records grouped by decoded major tick-mark location.</summary>
        public IReadOnlyDictionary<string, int> ChartTickMajorLocations { get; }

        /// <summary>Gets Tick records grouped by decoded axis-label location.</summary>
        public IReadOnlyDictionary<string, int> ChartTickLabelLocations { get; }

        /// <summary>Gets preserve-only chart BIFF records grouped by workbook or sheet location.</summary>
        public IReadOnlyDictionary<string, int> ChartRecordsByLocation { get; }

        /// <summary>Gets preserve-only drawing and object BIFF records grouped by shallow category.</summary>
        public IReadOnlyDictionary<LegacyXlsDrawingRecordKind, int> DrawingRecordsByKind { get; }

        /// <summary>Gets preserve-only drawing and object BIFF records grouped by record name.</summary>
        public IReadOnlyDictionary<string, int> DrawingRecordsByName { get; }

        /// <summary>Gets OBJ records grouped by decoded common-object type identifier.</summary>
        public IReadOnlyDictionary<string, int> DrawingRecordsByObjectType { get; }

        /// <summary>Gets OBJ records grouped by decoded common-object type name.</summary>
        public IReadOnlyDictionary<string, int> DrawingRecordsByObjectTypeName { get; }

        /// <summary>Gets OBJ records grouped by decoded common-object flag bitfield.</summary>
        public IReadOnlyDictionary<string, int> DrawingRecordsByObjectFlags { get; }

        /// <summary>Gets OBJ records grouped by decoded common-object flag name.</summary>
        public IReadOnlyDictionary<string, int> DrawingRecordsByObjectFlagName { get; }

        /// <summary>Gets MsoDrawing records grouped by decoded top-level Escher record type.</summary>
        public IReadOnlyDictionary<string, int> DrawingRecordsByEscherRecordType { get; }

        /// <summary>Gets MsoDrawing records grouped by decoded top-level Escher record type name.</summary>
        public IReadOnlyDictionary<string, int> DrawingRecordsByEscherRecordTypeName { get; }

        /// <summary>Gets OfficeArt FBSE image-store entries grouped by decoded BLIP type.</summary>
        public IReadOnlyDictionary<string, int> DrawingBlipStoreEntriesByType { get; }

        /// <summary>Gets OfficeArt FBSE image-store entries grouped by embedded BLIP record type.</summary>
        public IReadOnlyDictionary<string, int> DrawingBlipStoreEntriesByEmbeddedRecordType { get; }

        /// <summary>Gets OfficeArt FBSE image-store entries grouped by stored byte size.</summary>
        public IReadOnlyDictionary<string, int> DrawingBlipStoreEntriesBySize { get; }

        /// <summary>Gets OfficeArt FBSE image-store entries grouped by reference count.</summary>
        public IReadOnlyDictionary<string, int> DrawingBlipStoreEntriesByReferenceCount { get; }

        /// <summary>Gets preserve-only drawing and object BIFF records grouped by workbook or sheet location.</summary>
        public IReadOnlyDictionary<string, int> DrawingRecordsByLocation { get; }

        /// <summary>Gets preserve-only compound feature records grouped by kind.</summary>
        public IReadOnlyDictionary<LegacyXlsCompoundFeatureRecordKind, int> CompoundFeatureRecordsByKind { get; }

        /// <summary>Gets matching compound feature entries grouped by feature kind.</summary>
        public IReadOnlyDictionary<LegacyXlsCompoundFeatureRecordKind, int> CompoundFeatureEntriesByKind { get; }

        /// <summary>Gets matching compound feature entries grouped by compound entry path or name.</summary>
        public IReadOnlyDictionary<string, int> CompoundFeatureEntriesByName { get; }

        /// <summary>Gets matching compound feature entries grouped by preserve-only entry role.</summary>
        public IReadOnlyDictionary<string, int> CompoundFeatureEntriesByRole { get; }

        /// <summary>Gets matching compound feature entries grouped by feature kind and entry role.</summary>
        public IReadOnlyDictionary<string, int> CompoundFeatureEntriesByKindAndRole { get; }

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
            builder.AppendLine($"Differential formats: {DifferentialFormatCount}");
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
            AppendDictionary(builder, "Formula Token Blockers By Token", FormulaTokenBlockersByToken);
            AppendDictionary(builder, "Formula Token Blockers By Token Name", FormulaTokenBlockersByTokenName);
            AppendDictionary(builder, "Formula Token Blockers By Offset", FormulaTokenBlockersByOffset);
            AppendDictionary(builder, "Data Validations By Type", DataValidationsByType);
            AppendDictionary(builder, "Data Validations By Operator", DataValidationsByOperator);
            AppendDictionary(builder, "Data Validations By Error Style", DataValidationsByErrorStyle);
            AppendDictionary(builder, "Data Validation List Sources By Kind", DataValidationListSourcesByKind);
            AppendDictionary(builder, "Conditional Formatting By Type", ConditionalFormattingsByType);
            AppendDictionary(builder, "Conditional Formatting By Operator", ConditionalFormattingsByOperator);
            AppendDictionary(builder, "AutoFilter Criteria By Kind", AutoFilterCriteriaByKind);
            AppendDictionary(builder, "AutoFilter Criteria By Operator", AutoFilterCriteriaByOperator);
            AppendDictionary(builder, "AutoFilter Criteria By Value Kind", AutoFilterCriteriaByValueKind);
            AppendDictionary(builder, "AutoFilter Criteria By Join Operator", AutoFilterCriteriaByJoinOperator);
            AppendDictionary(builder, "AutoFilter Top10 Kinds", AutoFilterTop10Kinds);
            AppendDictionary(builder, "AutoFilter Top10 Values", AutoFilterTop10Values);
            AppendDictionary(builder, "Unsupported Features By Code", UnsupportedFeaturesByCode);
            AppendDictionary(builder, "Unsupported Features By Kind", UnsupportedFeaturesByKind.ToDictionary(
                entry => entry.Key.ToString(),
                entry => entry.Value,
                StringComparer.OrdinalIgnoreCase));
            AppendDictionary(builder, "Unsupported Feature Record Types", UnsupportedFeaturesByRecordType);
            AppendDictionary(builder, "Unsupported Feature Details", UnsupportedFeaturesByDetail);
            AppendDictionary(builder, "Unsupported Feature Locations", UnsupportedFeaturesByLocation);
            AppendDictionary(builder, "File Format Blockers", FileFormatBlockers);
            AppendDictionary(builder, "Encrypted Workbooks By Method", EncryptedWorkbooksByMethod);
            AppendDictionary(builder, "Unsupported BIFF Versions By Version", UnsupportedBiffVersionsByVersion);
            AppendDictionary(builder, "Unsupported BIFF Versions By Substream", UnsupportedBiffVersionsBySubstream);
            AppendDictionary(builder, "Unsupported Sheets By Kind", UnsupportedSheetsByKind.ToDictionary(
                entry => entry.Key.ToString(),
                entry => entry.Value,
                StringComparer.OrdinalIgnoreCase));
            AppendDictionary(builder, "Unsupported Sheets By Type", UnsupportedSheetsByType);
            AppendDictionary(builder, "Unsupported Sheets By Name", UnsupportedSheetsByName);
            AppendDictionary(builder, "Unsupported Chart Sheet Print Sizes", UnsupportedChartSheetPrintSizes);
            AppendDictionary(builder, "Unsupported Chart Sheet Text Object Counts", UnsupportedChartSheetTextObjectCounts);
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
            AppendDictionary(builder, "External Names By Scope", ExternalNamesByScope);
            AppendDictionary(builder, "External Names By Built-In State", ExternalNamesByBuiltInState);
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
            AppendDictionary(builder, "Pivot Table Data Item Aggregation Kinds", PivotTableDataItemAggregationKinds);
            AppendDictionary(builder, "Pivot Table Data Item Display Calculations", PivotTableDataItemDisplayCalculations);
            AppendDictionary(builder, "Pivot Table Grouping Kinds", PivotTableGroupingKinds);
            AppendDictionary(builder, "Chart Records By Kind", ChartRecordsByKind.ToDictionary(
                entry => entry.Key.ToString(),
                entry => entry.Value,
                StringComparer.OrdinalIgnoreCase));
            AppendDictionary(builder, "Chart Records By Name", ChartRecordsByName);
            AppendDictionary(builder, "Chart Records By Chart Type", ChartRecordsByChartType);
            AppendDictionary(builder, "Chart Records By Rectangle", ChartRecordsByRectangle);
            AppendDictionary(builder, "Chart Records By Axis Type", ChartRecordsByAxisType);
            AppendDictionary(builder, "Chart Records By Axes Used Count", ChartRecordsByAxesUsedCount);
            AppendDictionary(builder, "Chart Series Category Data Types", ChartSeriesCategoryDataTypes);
            AppendDictionary(builder, "Chart Series Value Counts", ChartSeriesValueCounts);
            AppendDictionary(builder, "Chart DataFormat Targets", ChartDataFormatTargets);
            AppendDictionary(builder, "Chart DataFormat Series Indexes", ChartDataFormatSeriesIndexes);
            AppendDictionary(builder, "Chart LineFormat Styles", ChartLineFormatStyles);
            AppendDictionary(builder, "Chart LineFormat Weights", ChartLineFormatWeights);
            AppendDictionary(builder, "Chart AreaFormat Patterns", ChartAreaFormatPatterns);
            AppendDictionary(builder, "Chart MarkerFormat Types", ChartMarkerFormatTypes);
            AppendDictionary(builder, "Chart MarkerFormat Sizes", ChartMarkerFormatSizes);
            AppendDictionary(builder, "Chart DefaultText Targets", ChartDefaultTextTargets);
            AppendDictionary(builder, "Chart Text Horizontal Alignments", ChartTextHorizontalAlignments);
            AppendDictionary(builder, "Chart Text Vertical Alignments", ChartTextVerticalAlignments);
            AppendDictionary(builder, "Chart Text Data Label Positions", ChartTextDataLabelPositions);
            AppendDictionary(builder, "Chart Text Flags", ChartTextFlags);
            AppendDictionary(builder, "Chart ObjectLink Targets", ChartObjectLinkTargets);
            AppendDictionary(builder, "Chart Legend Layouts", ChartLegendLayouts);
            AppendDictionary(builder, "Chart Tick Major Locations", ChartTickMajorLocations);
            AppendDictionary(builder, "Chart Tick Label Locations", ChartTickLabelLocations);
            AppendDictionary(builder, "Chart Records By Location", ChartRecordsByLocation);
            AppendDictionary(builder, "Drawing Records By Kind", DrawingRecordsByKind.ToDictionary(
                entry => entry.Key.ToString(),
                entry => entry.Value,
                StringComparer.OrdinalIgnoreCase));
            AppendDictionary(builder, "Drawing Records By Name", DrawingRecordsByName);
            AppendDictionary(builder, "Drawing Records By Object Type", DrawingRecordsByObjectType);
            AppendDictionary(builder, "Drawing Records By Object Type Name", DrawingRecordsByObjectTypeName);
            AppendDictionary(builder, "Drawing Records By Object Flags", DrawingRecordsByObjectFlags);
            AppendDictionary(builder, "Drawing Records By Object Flag Name", DrawingRecordsByObjectFlagName);
            AppendDictionary(builder, "Drawing Records By Escher Record Type", DrawingRecordsByEscherRecordType);
            AppendDictionary(builder, "Drawing Records By Escher Record Type Name", DrawingRecordsByEscherRecordTypeName);
            AppendDictionary(builder, "Drawing BLIP Store Entries By Type", DrawingBlipStoreEntriesByType);
            AppendDictionary(builder, "Drawing BLIP Store Entries By Embedded Record Type", DrawingBlipStoreEntriesByEmbeddedRecordType);
            AppendDictionary(builder, "Drawing BLIP Store Entries By Size", DrawingBlipStoreEntriesBySize);
            AppendDictionary(builder, "Drawing BLIP Store Entries By Reference Count", DrawingBlipStoreEntriesByReferenceCount);
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
            AppendDictionary(builder, "Compound Feature Entries By Role", CompoundFeatureEntriesByRole);
            AppendDictionary(builder, "Compound Feature Entries By Kind And Role", CompoundFeatureEntriesByKindAndRole);
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

        private static string GetEncryptionMethodKey(LegacyXlsUnsupportedFeature feature) {
            const string prefix = "Encryption:FilePass:";
            return feature.DetailCode != null && feature.DetailCode.StartsWith(prefix, StringComparison.OrdinalIgnoreCase)
                ? feature.DetailCode.Substring(prefix.Length)
                : feature.DetailCode ?? feature.Code;
        }

        private static string GetBiffVersionKey(LegacyXlsUnsupportedFeature feature) {
            string[] parts = (feature.DetailCode ?? string.Empty).Split(':');
            return parts.Length >= 2 && !string.IsNullOrWhiteSpace(parts[1]) ? parts[1] : feature.DetailCode ?? feature.Code;
        }

        private static string GetBiffSubstreamKey(LegacyXlsUnsupportedFeature feature) {
            string[] parts = (feature.DetailCode ?? string.Empty).Split(':');
            return parts.Length >= 3 && !string.IsNullOrWhiteSpace(parts[2]) ? parts[2] : feature.DetailCode ?? feature.Code;
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

        private static string GetAutoFilterTop10KindKey(LegacyXlsAutoFilterCriteria criteria) {
            string rank = criteria.Top10IsTop ? "Top" : "Bottom";
            string unit = criteria.Top10IsPercent ? "Percent" : "Items";
            return rank + unit;
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
