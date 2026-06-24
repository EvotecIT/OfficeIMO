using OfficeIMO.Excel.LegacyXls.Diagnostics;
using OfficeIMO.Excel.LegacyXls.Model;
using System.Globalization;
using System.Text;

namespace OfficeIMO.Excel.LegacyXls {
    /// <summary>
    /// Compact import summary intended for corpus baselines and preflight checks.
    /// </summary>
    public sealed class LegacyXlsImportReport {
        internal LegacyXlsImportReport(LegacyXlsWorkbook workbook) {
            if (workbook == null) throw new ArgumentNullException(nameof(workbook));

            LegacyXlsDataValidation[] dataValidations = workbook.Worksheets.SelectMany(sheet => sheet.DataValidations).ToArray();
            LegacyXlsConditionalFormatting[] conditionalFormattings = workbook.Worksheets.SelectMany(sheet => sheet.ConditionalFormattings).ToArray();
            WorksheetCount = workbook.Worksheets.Count;
            UnsupportedSheetCount = workbook.UnsupportedSheets.Count;
            CellCount = workbook.Worksheets.Sum(sheet => sheet.Cells.Count);
            FormulaCellCount = workbook.Worksheets.Sum(sheet => sheet.Cells.Count(cell => cell.IsFormula));
            CommentCount = workbook.Worksheets.Sum(sheet => sheet.Comments.Count);
            HyperlinkCount = workbook.Worksheets.Sum(sheet => sheet.Hyperlinks.Count);
            DataValidationCount = dataValidations.Length;
            ConditionalFormattingCount = conditionalFormattings.Length;
            AutoFilterCriteriaCount = workbook.Worksheets.Sum(sheet => sheet.AutoFilterCriteria.Count);
            DataValidationsByType = CountByCode(dataValidations.Select(validation => validation.Type.ToString()));
            DataValidationsByOperator = CountByCode(dataValidations.Select(validation => validation.Operator.ToString()));
            DataValidationsByErrorStyle = CountByCode(dataValidations.Select(validation => validation.ErrorStyle.ToString()));
            DataValidationsByAllowBlankState = CountByCode(dataValidations.Select(validation => validation.AllowBlank ? "AllowBlank" : "RejectBlank"));
            DataValidationsByInputMessageState = CountByCode(dataValidations.Select(validation => validation.ShowInputMessage ? "ShowInputMessage" : "HideInputMessage"));
            DataValidationsByErrorMessageState = CountByCode(dataValidations.Select(validation => validation.ShowErrorMessage ? "ShowErrorMessage" : "HideErrorMessage"));
            DataValidationsByPromptTextState = CountByCode(dataValidations.Select(validation => validation.PromptTitle != null || validation.Prompt != null ? "Present" : "Missing"));
            DataValidationsByErrorTextState = CountByCode(dataValidations.Select(validation => validation.ErrorTitle != null || validation.Error != null ? "Present" : "Missing"));
            DataValidationsByDropDownState = CountByCode(dataValidations.Select(GetDataValidationDropDownState));
            DataValidationsByRangeCount = CountByCode(dataValidations.Select(validation => $"Ranges:{validation.RangeCount}"));
            DataValidationsByRange = CountByCode(dataValidations.SelectMany(validation => validation.Ranges));
            DataValidationsByFormula1State = CountByCode(dataValidations.Select(validation => GetFormulaStateKey(validation.Formula1)));
            DataValidationsByFormula2State = CountByCode(dataValidations.Select(validation => GetFormulaStateKey(validation.Formula2)));
            DataValidationsByFormulaPairState = CountByCode(dataValidations.Select(validation => GetFormulaPairStateKey(validation.Formula1, validation.Formula2)));
            DataValidationListSourcesByKind = CountByCode(dataValidations
                .Where(validation => validation.Type == LegacyXlsDataValidationType.List)
                .Select(validation => validation.ListSourceKind.ToString()));
            DataValidationListSourcesByItemCount = CountByCode(dataValidations
                .Where(validation => validation.Type == LegacyXlsDataValidationType.List)
                .Select(validation => $"Items:{validation.ListItems.Count.ToString(CultureInfo.InvariantCulture)}"));
            DataValidationListSourcesByRange = CountByCode(dataValidations
                .Where(validation => validation.Type == LegacyXlsDataValidationType.List && !string.IsNullOrWhiteSpace(validation.ListSourceRange))
                .Select(validation => validation.ListSourceRange!));
            DataValidationListSourcesByName = CountByCode(dataValidations
                .Where(validation => validation.Type == LegacyXlsDataValidationType.List && !string.IsNullOrWhiteSpace(validation.ListSourceName))
                .Select(validation => validation.ListSourceName!));
            DataValidationListSourcesBySheetName = CountByCode(dataValidations
                .Where(validation => validation.Type == LegacyXlsDataValidationType.List && !string.IsNullOrWhiteSpace(validation.ListSourceSheetName))
                .Select(validation => validation.ListSourceSheetName!));
            ConditionalFormattingsByType = CountByCode(conditionalFormattings.Select(formatting => formatting.Type.ToString()));
            ConditionalFormattingsByOperator = CountByCode(conditionalFormattings
                .Where(formatting => formatting.Operator.HasValue)
                .Select(formatting => formatting.Operator!.Value.ToString()));
            ConditionalFormattingsByRangeCount = CountByCode(conditionalFormattings.Select(formatting => $"Ranges:{formatting.RangeCount}"));
            ConditionalFormattingsByRange = CountByCode(conditionalFormattings.SelectMany(formatting => formatting.Ranges));
            ConditionalFormattingsByFormula1State = CountByCode(conditionalFormattings.Select(formatting => GetFormulaStateKey(formatting.Formula1)));
            ConditionalFormattingsByFormula2State = CountByCode(conditionalFormattings.Select(formatting => GetFormulaStateKey(formatting.Formula2)));
            ConditionalFormattingsByFormulaPairState = CountByCode(conditionalFormattings.Select(formatting => GetFormulaPairStateKey(formatting.Formula1, formatting.Formula2)));
            ConditionalFormattingsByPriorityState = CountByCode(conditionalFormattings.Select(formatting => formatting.Priority.HasValue ? "Present" : "Missing"));
            ConditionalFormattingsByPriority = CountByCode(conditionalFormattings
                .Where(formatting => formatting.Priority.HasValue)
                .Select(formatting => $"Priority:{formatting.Priority!.Value}"));
            ConditionalFormattingsByStopIfTrueState = CountByCode(conditionalFormattings.Select(formatting => formatting.StopIfTrue ? "StopIfTrue" : "Continue"));
            ConditionalFormattingsByDifferentialFormatState = CountByCode(conditionalFormattings.Select(formatting => formatting.DifferentialFormat == null ? "Missing" : "Present"));
            ConditionalFormattingsByDifferentialFill = CountByCode(conditionalFormattings.SelectMany(GetConditionalFormattingDifferentialFillKeys));
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
            AutoFilterCriteriaByColumn = CountByCode(workbook.Worksheets
                .SelectMany(sheet => sheet.AutoFilterCriteria)
                .Select(criteria => $"Column:{criteria.ColumnId.ToString(CultureInfo.InvariantCulture)}"));
            AutoFilterCriteriaByConditionCount = CountByCode(workbook.Worksheets
                .SelectMany(sheet => sheet.AutoFilterCriteria)
                .Select(criteria => $"Conditions:{criteria.Conditions.Count.ToString(CultureInfo.InvariantCulture)}"));
            AutoFilterTop10Kinds = CountByCode(workbook.Worksheets
                .SelectMany(sheet => sheet.AutoFilterCriteria)
                .Where(criteria => criteria.IsTop10)
                .Select(GetAutoFilterTop10KindKey));
            AutoFilterTop10Values = CountByCode(workbook.Worksheets
                .SelectMany(sheet => sheet.AutoFilterCriteria)
                .Where(criteria => criteria.IsTop10 && criteria.Top10Value.HasValue)
                .Select(criteria => $"{GetAutoFilterTop10KindKey(criteria)}:{criteria.Top10Value!.Value}"));
            AutoFilterTop10Directions = CountByCode(workbook.Worksheets
                .SelectMany(sheet => sheet.AutoFilterCriteria)
                .Where(criteria => criteria.IsTop10)
                .Select(criteria => criteria.Top10IsTop ? "Top" : "Bottom"));
            AutoFilterTop10Units = CountByCode(workbook.Worksheets
                .SelectMany(sheet => sheet.AutoFilterCriteria)
                .Where(criteria => criteria.IsTop10)
                .Select(criteria => criteria.Top10IsPercent ? "Percent" : "Items"));
            WorksheetsByVisibility = CountByCode(workbook.Worksheets.Select(sheet => sheet.VisibilityName));
            WorkbookCodeNameStates = CountByCode(new[] { string.IsNullOrWhiteSpace(workbook.CodeName) ? "Missing" : "Present" });
            WorkbookCodeNames = CountByCode(string.IsNullOrWhiteSpace(workbook.CodeName) ? Array.Empty<string>() : new[] { workbook.CodeName! });
            WorkbookOptionStates = CountByCode(GetWorkbookOptionStateKeys(workbook));
            WorkbookBuiltInFunctionGroupCounts = CountByCode(workbook.BuiltInFunctionGroupCount.HasValue
                ? new[] { $"Count:{workbook.BuiltInFunctionGroupCount.Value}" }
                : Array.Empty<string>());
            WorksheetCodeNameStates = CountByCode(workbook.Worksheets.Select(sheet => string.IsNullOrWhiteSpace(sheet.CodeName) ? "Missing" : "Present"));
            WorksheetCodeNames = CountByCode(workbook.Worksheets
                .Where(sheet => !string.IsNullOrWhiteSpace(sheet.CodeName))
                .Select(sheet => sheet.CodeName!));
            DefinedNameCount = workbook.DefinedNames.Count;
            ExternalReferenceCount = workbook.ExternalReferences.Count;
            ExternalSheetNameCount = workbook.ExternalReferences.Sum(reference => reference.SheetNames.Count);
            ExternalNameCount = workbook.ExternalReferences.Sum(reference => reference.ExternalNames.Count);
            ExternalCellCacheCount = workbook.ExternalReferences.Sum(reference => reference.CachedCellCaches.Count);
            ExternalCachedCellCount = workbook.ExternalReferences.Sum(reference => reference.CachedCellCaches.Sum(cache => cache.Cells.Count));
            DataConsolidationReferenceCount = workbook.DataConsolidationReferences.Count;
            PivotTableRecordCount = workbook.PivotTableRecords.Count;
            ChartRecordCount = workbook.ChartRecords.Count;
            DrawingRecordCount = workbook.DrawingRecords.Count;
            ThemeRecordCount = workbook.ThemeRecords.Count;
            DrawingOfficeArtRecordCount = workbook.DrawingRecords.Sum(record => record.OfficeArtRecords.Count);
            DrawingShapePropertyCount = workbook.DrawingRecords.Sum(record => record.ShapeProperties.Count);
            DifferentialFormatCount = workbook.DifferentialFormats.Count;
            CompoundFeatureRecordCount = workbook.CompoundFeatureRecords.Count;
            CompoundFeatureEntryCount = workbook.CompoundFeatureRecords.Sum(record => record.Entries.Count);
            CompoundVbaModuleCount = workbook.CompoundFeatureRecords.Sum(record => record.VbaModuleCount);
            CompoundFeatureEntryByteCount = workbook.CompoundFeatureRecords.Sum(record => record.EntryByteCount);
            CompoundVbaModuleByteCount = workbook.CompoundFeatureRecords.Sum(record => record.VbaModuleByteCount);
            CalculationSettingRecordCount = workbook.CalculationSettings.Records.Count;
            CellStyleRecordCount = workbook.CellStyles.Count;
            CellStyleExtensionRecordCount = workbook.CellStyleExtensions.Count;
            FormulaTokenRecordCount = workbook.FormulaTokenRecords.Count;
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
            FormulaTokensByName = CountByCode(workbook.FormulaTokenRecords.Select(record => record.TokenName));
            FormulaTokensByContext = CountByCode(workbook.FormulaTokenRecords.Select(record => record.Context));
            FormulaTokensByRecordType = CountByCode(workbook.FormulaTokenRecords.Select(record => $"0x{record.RecordType:X4}|{record.TokenName}"));
            FormulaFunctionsById = CountByCode(workbook.FormulaTokenRecords
                .Where(record => record.FunctionId.HasValue)
                .Select(record => $"Function:0x{record.FunctionId!.Value:X4}"));
            FormulaFunctionsByName = CountByCode(workbook.FormulaTokenRecords
                .Where(record => !string.IsNullOrWhiteSpace(record.FunctionName))
                .Select(record => record.FunctionName!));
            FormulaAttributesByName = CountByCode(workbook.FormulaTokenRecords
                .Where(record => !string.IsNullOrWhiteSpace(record.AttributeName))
                .Select(record => record.AttributeName!));
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
            UnsupportedSheetsByVisibility = CountByCode(workbook.UnsupportedSheets.Select(sheet => sheet.VisibilityName));
            UnsupportedSheetsByKindAndVisibility = CountByCode(workbook.UnsupportedSheets.Select(sheet => $"{sheet.Kind}|{sheet.VisibilityName}"));
            UnsupportedChartSheetPrintSizes = CountByCode(workbook.UnsupportedSheets
                .Where(sheet => sheet.Kind == LegacyXlsUnsupportedSheetKind.ChartSheet && sheet.ChartPrintSize.HasValue)
                .Select(sheet => $"PrintSize:{sheet.ChartPrintSize!.Value}"));
            UnsupportedChartSheetPrintSizeKinds = CountByCode(workbook.UnsupportedSheets
                .Where(sheet => sheet.Kind == LegacyXlsUnsupportedSheetKind.ChartSheet && !string.IsNullOrWhiteSpace(sheet.ChartPrintSizeName))
                .Select(sheet => sheet.ChartPrintSizeName!));
            UnsupportedChartSheetTextObjectCounts = CountByCode(workbook.UnsupportedSheets
                .Where(sheet => sheet.Kind == LegacyXlsUnsupportedSheetKind.ChartSheet && sheet.ChartTextObjectCount > 0)
                .Select(sheet => $"TextObjects:{sheet.ChartTextObjectCount}"));
            UnsupportedChartSheetChartRecordCounts = CountByCode(workbook.UnsupportedSheets
                .Where(sheet => sheet.Kind == LegacyXlsUnsupportedSheetKind.ChartSheet && sheet.ChartRecordCount > 0)
                .Select(sheet => $"ChartRecords:{sheet.ChartRecordCount}"));
            UnsupportedChartSheetChartRecordKinds = CountUnsupportedChartSheetChartRecordKinds(workbook.UnsupportedSheets);
            UnsupportedChartSheetChartTypes = CountUnsupportedChartSheetChartTypes(workbook.UnsupportedSheets);
            ExternalReferencesByKind = CountExternalReferencesByKind(workbook.ExternalReferences);
            ExternalReferencesByTarget = CountByCode(workbook.ExternalReferences.Select(GetExternalReferenceTargetKey));
            ExternalReferencesByShape = CountByCode(workbook.ExternalReferences.Select(GetExternalReferenceShapeKey));
            ExternalReferencesByDeclaredSheetCount = CountByCode(workbook.ExternalReferences.Select(reference => $"DeclaredSheets:{reference.SheetCount}"));
            ExternalReferencesBySheetNameCount = CountByCode(workbook.ExternalReferences.Select(reference => $"Sheets:{reference.SheetNameCount}"));
            ExternalReferencesBySheetTableState = CountByCode(workbook.ExternalReferences.Select(GetExternalReferenceSheetTableStateKey));
            ExternalReferencesByExternalNameCount = CountByCode(workbook.ExternalReferences.Select(reference => $"Names:{reference.ExternalNameCount}"));
            ExternalReferencesByCacheCount = CountByCode(workbook.ExternalReferences.Select(reference => $"Caches:{reference.CachedCellCacheCount}"));
            ExternalReferencesByCachedCellCount = CountByCode(workbook.ExternalReferences.Select(reference => $"CachedCells:{reference.CachedCellCount}"));
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
            ExternalCellCachesByCellRange = CountByCode(workbook.ExternalReferences.SelectMany(reference => reference.CachedCellCaches.Select(GetExternalCellCacheRangeKey)));
            ExternalCellCachesByCellCount = CountByCode(workbook.ExternalReferences.SelectMany(reference => reference.CachedCellCaches.Select(cache => $"Cells:{cache.Cells.Count}")));
            ExternalCellCachesByRowSpan = CountByCode(workbook.ExternalReferences.SelectMany(reference => reference.CachedCellCaches.Select(cache => cache.RowSpan.HasValue ? $"Rows:{cache.RowSpan.Value}" : "(empty)")));
            ExternalCellCachesByColumnSpan = CountByCode(workbook.ExternalReferences.SelectMany(reference => reference.CachedCellCaches.Select(cache => cache.ColumnSpan.HasValue ? $"Columns:{cache.ColumnSpan.Value}" : "(empty)")));
            ExternalCellCachesByLinkState = CountByCode(workbook.ExternalReferences.SelectMany(reference => reference.CachedCellCaches.Select(cache => cache.LinkValid ? "ValidLink" : "InvalidLink")));
            ExternalCachedCellsByValueKind = CountExternalCachedCellsByValueKind(workbook.ExternalReferences);
            DataConsolidationReferencesBySourceKind = CountByCode(workbook.DataConsolidationReferences.Select(reference => reference.SourceKind.ToString()));
            DataConsolidationReferencesBySource = CountByCode(workbook.DataConsolidationReferences.Select(reference => reference.Source));
            DataConsolidationReferencesByRange = CountByCode(workbook.DataConsolidationReferences.Select(reference => reference.CellRange));
            DataConsolidationReferencesByUnusedByteCount = CountByCode(workbook.DataConsolidationReferences.Select(reference => $"UnusedBytes:{reference.UnusedByteCount}"));
            ThemeRecordsByVersion = CountByCode(workbook.ThemeRecords.Select(record => record.ThemeVersionName));
            ThemeRecordsByRawVersion = CountByCode(workbook.ThemeRecords.Select(record => $"Version:{record.ThemeVersion}"));
            ThemeRecordsByContentState = CountByCode(workbook.ThemeRecords.Select(record => record.HasThemeBytes ? "EmbeddedThemeBytes" : "NoEmbeddedThemeBytes"));
            ThemeRecordsByContentLength = CountByCode(workbook.ThemeRecords.Select(record => $"Bytes:{record.ThemeByteCount}"));
            PivotTableRecordsByKind = CountPivotTableRecordsByKind(workbook.PivotTableRecords);
            PivotTableRecordsByName = CountByCode(workbook.PivotTableRecords.Select(record => record.RecordName));
            PivotTableDataItemAggregations = CountByCode(workbook.PivotTableRecords
                .Where(record => record.AggregationFunction.HasValue)
                .Select(record => $"AggregationFunction:{record.AggregationFunction!.Value}"));
            PivotTableDataItemAggregationKinds = CountByCode(workbook.PivotTableRecords
                .Where(record => !string.IsNullOrWhiteSpace(record.AggregationFunctionName))
                .Select(record => record.AggregationFunctionName!));
            PivotTableDataItemFieldIndexes = CountByCode(workbook.PivotTableRecords
                .Where(record => record.DataItemFieldIndex.HasValue)
                .Select(record => $"FieldIndex:{record.DataItemFieldIndex!.Value}"));
            PivotTableDataItemDisplayCalculations = CountByCode(workbook.PivotTableRecords
                .Where(record => !string.IsNullOrWhiteSpace(record.DisplayCalculationName))
                .Select(record => record.DisplayCalculationName!));
            PivotTableDataItemDisplayCalculationFieldIndexes = CountByCode(workbook.PivotTableRecords
                .Where(record => record.DisplayCalculationFieldIndex.HasValue)
                .Select(record => $"FieldIndex:{record.DisplayCalculationFieldIndex!.Value}"));
            PivotTableDataItemDisplayCalculationItemIndexes = CountByCode(workbook.PivotTableRecords
                .Where(record => record.DisplayCalculationItemIndex.HasValue)
                .Select(record => $"ItemIndex:{record.DisplayCalculationItemIndex!.Value}"));
            PivotTableDataItemNumberFormats = CountByCode(workbook.PivotTableRecords
                .Where(record => record.NumberFormatId.HasValue)
                .Select(record => $"NumberFormatId:{record.NumberFormatId!.Value}"));
            PivotTableDataItemNames = CountByCode(workbook.PivotTableRecords
                .Where(record => !string.IsNullOrWhiteSpace(record.Name))
                .Select(record => record.Name!));
            PivotTableGroupingKinds = CountByCode(workbook.PivotTableRecords
                .Where(record => record.GroupingKind.HasValue)
                .Select(record => record.GroupingKind!.Value.ToString()));
            PivotTableGroupingBoundaryStates = CountByCode(workbook.PivotTableRecords
                .Where(record => record.AutoStart.HasValue && record.AutoEnd.HasValue)
                .Select(record => $"AutoStart:{record.AutoStart!.Value};AutoEnd:{record.AutoEnd!.Value}"));
            PivotTableGroupingNumericRanges = CountByCode(workbook.PivotTableRecords
                .Where(record => record.GroupingNumericStart.HasValue && record.GroupingNumericEnd.HasValue && record.GroupingNumericInterval.HasValue)
                .Select(GetPivotTableGroupingNumericRangeKey));
            PivotTableGroupingDateRanges = CountByCode(workbook.PivotTableRecords
                .Where(record => record.GroupingDateStart != null && record.GroupingDateEnd != null && record.GroupingDateInterval.HasValue)
                .Select(GetPivotTableGroupingDateRangeKey));
            PivotTableExtendedFieldStates = CountByCode(workbook.PivotTableRecords.SelectMany(GetPivotTableExtendedFieldStateKeys));
            PivotTableAdditionalClasses = CountByCode(workbook.PivotTableRecords
                .Where(record => !string.IsNullOrWhiteSpace(record.AdditionalClassName))
                .Select(record => record.AdditionalClassName!));
            PivotTableAdditionalTypes = CountByCode(workbook.PivotTableRecords
                .Where(record => !string.IsNullOrWhiteSpace(record.AdditionalTypeName))
                .Select(record => record.AdditionalTypeName!));
            PivotTableAdditionalClassTypes = CountByCode(workbook.PivotTableRecords
                .Where(record => !string.IsNullOrWhiteSpace(record.AdditionalClassName) && !string.IsNullOrWhiteSpace(record.AdditionalTypeName))
                .Select(GetPivotTableAdditionalClassTypeKey));
            PivotTableAdditionalCacheIds = CountByCode(workbook.PivotTableRecords
                .Where(record => record.AdditionalCacheId.HasValue)
                .Select(record => $"CacheId:{record.AdditionalCacheId!.Value}"));
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
            ChartSeriesValueDataTypes = CountByCode(workbook.ChartRecords
                .Where(record => !string.IsNullOrWhiteSpace(record.SeriesValueDataTypeName))
                .Select(record => record.SeriesValueDataTypeName!));
            ChartSeriesBubbleSizeDataTypes = CountByCode(workbook.ChartRecords
                .Where(record => !string.IsNullOrWhiteSpace(record.SeriesBubbleSizeDataTypeName))
                .Select(record => record.SeriesBubbleSizeDataTypeName!));
            ChartSeriesValueCounts = CountByCode(workbook.ChartRecords
                .Where(record => record.SeriesCategoryCount.HasValue && record.SeriesValueCount.HasValue && record.SeriesBubbleSizeCount.HasValue)
                .Select(record => $"Categories:{record.SeriesCategoryCount!.Value};Values:{record.SeriesValueCount!.Value};BubbleSizes:{record.SeriesBubbleSizeCount!.Value}"));
            ChartDataFormatTargets = CountByCode(workbook.ChartRecords
                .Where(record => !string.IsNullOrWhiteSpace(record.DataFormatTarget))
                .Select(record => record.DataFormatTarget!));
            ChartDataFormatSeriesIndexes = CountByCode(workbook.ChartRecords
                .Where(record => record.DataFormatSeriesIndex.HasValue)
                .Select(record => $"SeriesIndex:{record.DataFormatSeriesIndex!.Value}"));
            ChartNumberFormatIds = CountByCode(workbook.ChartRecords
                .Where(record => record.NumberFormatId.HasValue)
                .Select(record => $"NumberFormatId:{record.NumberFormatId!.Value}"));
            ChartFontIndexes = CountByCode(workbook.ChartRecords
                .Where(record => record.FontIndex.HasValue)
                .Select(record => $"FontIndex:{record.FontIndex!.Value}"));
            ChartDataTableOptions = CountByCode(workbook.ChartRecords
                .Where(record => record.DataTableOptions != null)
                .Select(record => record.DataTableOptions!)
                .Select(options => $"HorizontalBorders:{options.HasHorizontalBorders};VerticalBorders:{options.HasVerticalBorders};Outline:{options.HasOutlineBorder};SeriesKeys:{options.ShowSeriesKeys}"));
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
            ChartValueRangeScales = CountByCode(workbook.ChartRecords
                .Where(record => record.ValueRange != null)
                .Select(GetChartValueRangeScaleKey));
            ChartValueRangeStates = CountByCode(workbook.ChartRecords
                .Where(record => record.ValueRange != null)
                .Select(GetChartValueRangeStateKey));
            ChartPositionModePairs = CountByCode(workbook.ChartRecords
                .Where(record => record.Position != null)
                .Select(record => $"{record.Position!.TopLeftModeName}/{record.Position.BottomRightModeName}"));
            ChartPositionRectangles = CountByCode(workbook.ChartRecords
                .Where(record => record.Position != null)
                .Select(record => $"X1:{record.Position!.X1};Y1:{record.Position.Y1};X2:{record.Position.X2};Y2:{record.Position.Y2}"));
            ChartFrameTypes = CountByCode(workbook.ChartRecords
                .Where(record => record.Frame != null)
                .Select(record => record.Frame!.FrameTypeName));
            ChartFrameAutoStates = CountByCode(workbook.ChartRecords
                .Where(record => record.Frame != null)
                .Select(record => $"AutoSize:{record.Frame!.AutomaticSize};AutoPosition:{record.Frame.AutomaticPosition}"));
            ChartPlotGrowthFactors = CountByCode(workbook.ChartRecords
                .Where(record => record.PlotGrowth != null)
                .Select(record => $"Horizontal:{FormatDouble(record.PlotGrowth!.HorizontalGrowthPoints)};Vertical:{FormatDouble(record.PlotGrowth.VerticalGrowthPoints)}"));
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
            DrawingObjectSubRecordsByType = CountByCode(workbook.DrawingRecords
                .SelectMany(record => record.ObjectSubRecords)
                .Select(subRecord => subRecord.SubRecordTypeKey));
            DrawingObjectSubRecordsByName = CountByCode(workbook.DrawingRecords
                .SelectMany(record => record.ObjectSubRecords)
                .Select(subRecord => subRecord.SubRecordName));
            DrawingObjectSubRecordsByDeclaredLength = CountByCode(workbook.DrawingRecords
                .SelectMany(record => record.ObjectSubRecords)
                .Select(subRecord => $"DeclaredBytes:{subRecord.DeclaredLength}"));
            DrawingObjectSubRecordsByCompleteness = CountByCode(workbook.DrawingRecords
                .SelectMany(record => record.ObjectSubRecords)
                .Select(subRecord => subRecord.IsComplete ? "Complete" : "Truncated"));
            DrawingRecordsByEscherRecordType = CountByCode(workbook.DrawingRecords
                .Where(record => record.EscherRecordType.HasValue)
                .Select(record => $"EscherRecordType:0x{record.EscherRecordType!.Value:X4}"));
            DrawingRecordsByEscherRecordTypeName = CountByCode(workbook.DrawingRecords
                .Where(record => !string.IsNullOrWhiteSpace(record.EscherRecordTypeName))
                .Select(record => record.EscherRecordTypeName!));
            DrawingOfficeArtRecordsByType = CountByCode(workbook.DrawingRecords
                .SelectMany(record => record.OfficeArtRecords)
                .Select(record => $"EscherRecordType:0x{record.RecordType:X4}"));
            DrawingOfficeArtRecordsByTypeName = CountByCode(workbook.DrawingRecords
                .SelectMany(record => record.OfficeArtRecords)
                .Select(record => record.RecordTypeName));
            DrawingOfficeArtRecordsByDepth = CountByCode(workbook.DrawingRecords
                .SelectMany(record => record.OfficeArtRecords)
                .Select(record => $"Depth:{record.Depth}"));
            DrawingOfficeArtRecordsByContainerState = CountByCode(workbook.DrawingRecords
                .SelectMany(record => record.OfficeArtRecords)
                .Select(record => record.IsContainer ? "Container" : "Leaf"));
            DrawingOfficeArtRecordsByPayloadLength = CountByCode(workbook.DrawingRecords
                .SelectMany(record => record.OfficeArtRecords)
                .Select(record => $"PayloadLength:{record.PayloadLength}"));
            DrawingShapePropertiesById = CountByCode(workbook.DrawingRecords
                .SelectMany(record => record.ShapeProperties)
                .Select(property => property.PropertyIdKey));
            DrawingShapePropertiesByName = CountByCode(workbook.DrawingRecords
                .SelectMany(record => record.ShapeProperties)
                .Select(property => property.PropertyName));
            DrawingShapePropertiesByGroup = CountByCode(workbook.DrawingRecords
                .SelectMany(record => record.ShapeProperties)
                .Select(property => property.PropertyGroupName));
            DrawingShapePropertiesByFlagState = CountByCode(workbook.DrawingRecords
                .SelectMany(record => record.ShapeProperties)
                .Select(GetShapePropertyFlagState));
            DrawingShapePropertiesByValue = CountByCode(workbook.DrawingRecords
                .SelectMany(record => record.ShapeProperties)
                .Where(property => !property.IsComplex)
                .Select(property => $"{property.PropertyIdKey};Value:0x{property.Value:X8}"));
            DrawingShapeComplexPropertiesByDeclaredLength = CountByCode(workbook.DrawingRecords
                .SelectMany(record => record.ShapeProperties)
                .Where(property => property.DeclaredComplexDataLength.HasValue)
                .Select(property => $"{property.PropertyIdKey};DeclaredBytes:{property.DeclaredComplexDataLength!.Value}"));
            DrawingShapeComplexPropertiesByAvailableLength = CountByCode(workbook.DrawingRecords
                .SelectMany(record => record.ShapeProperties)
                .Where(property => property.AvailableComplexDataLength.HasValue)
                .Select(property => $"{property.PropertyIdKey};AvailableBytes:{property.AvailableComplexDataLength!.Value}"));
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
            DrawingShapeEntriesByType = CountByCode(workbook.DrawingRecords
                .SelectMany(record => record.ShapeEntries)
                .Select(shape => shape.ShapeTypeName));
            DrawingShapeEntriesById = CountByCode(workbook.DrawingRecords
                .SelectMany(record => record.ShapeEntries)
                .Select(shape => $"ShapeId:{shape.ShapeId}"));
            DrawingShapeEntriesByFlags = CountByCode(workbook.DrawingRecords
                .SelectMany(record => record.ShapeEntries)
                .Select(shape => $"Flags:0x{shape.Flags:X8}"));
            DrawingShapeEntriesByFlagName = CountByCode(workbook.DrawingRecords
                .SelectMany(record => record.ShapeEntries)
                .SelectMany(shape => shape.FlagNames));
            DrawingAnchorEntriesByRange = CountByCode(workbook.DrawingRecords
                .SelectMany(record => record.AnchorEntries)
                .Select(anchor => $"R{anchor.StartRow}C{anchor.StartColumn}:R{anchor.EndRow}C{anchor.EndColumn}"));
            DrawingAnchorEntriesByOffset = CountByCode(workbook.DrawingRecords
                .SelectMany(record => record.AnchorEntries)
                .Select(anchor => $"StartDx:{anchor.StartDx};StartDy:{anchor.StartDy};EndDx:{anchor.EndDx};EndDy:{anchor.EndDy}"));
            DrawingAnchorEntriesByFlags = CountByCode(workbook.DrawingRecords
                .SelectMany(record => record.AnchorEntries)
                .Select(anchor => $"Flags:0x{anchor.Flags:X4}"));
            DrawingChildAnchorEntriesByRectangle = CountByCode(workbook.DrawingRecords
                .SelectMany(record => record.ChildAnchorEntries)
                .Select(anchor => $"Left:{anchor.Left};Top:{anchor.Top};Right:{anchor.Right};Bottom:{anchor.Bottom}"));
            DrawingChildAnchorEntriesBySize = CountByCode(workbook.DrawingRecords
                .SelectMany(record => record.ChildAnchorEntries)
                .Select(anchor => $"Width:{anchor.Width};Height:{anchor.Height}"));
            DrawingRecordsByLocation = CountByCode(workbook.DrawingRecords.Select(GetDrawingRecordLocationKey));
            CompoundFeatureRecordsByKind = CountCompoundFeatureRecordsByKind(workbook.CompoundFeatureRecords);
            CompoundFeatureEntriesByKind = CountCompoundFeatureEntriesByKind(workbook.CompoundFeatureRecords);
            CompoundFeatureEntriesByName = CountByCode(workbook.CompoundFeatureRecords.SelectMany(record => record.Entries));
            CompoundFeatureEntriesByRole = CountByCode(workbook.CompoundFeatureRecords
                .SelectMany(record => record.EntryDetails)
                .Select(entry => entry.Role.ToString()));
            CompoundFeatureEntriesByKindAndRole = CountByCode(workbook.CompoundFeatureRecords
                .SelectMany(record => record.EntryDetails.Select(entry => $"{record.Kind}|{entry.Role}")));
            CompoundFeatureEntriesByObjectType = CountByCode(workbook.CompoundFeatureRecords
                .SelectMany(record => record.EntryDetails)
                .Select(entry => entry.ObjectType.ToString()));
            CompoundFeatureEntriesByRoleAndObjectType = CountByCode(workbook.CompoundFeatureRecords
                .SelectMany(record => record.EntryDetails)
                .Select(entry => $"{entry.Role}|{entry.ObjectType}"));
            CompoundFeatureEntriesBySize = CountByCode(workbook.CompoundFeatureRecords
                .SelectMany(record => record.EntryDetails)
                .Select(GetCompoundFeatureEntrySizeKey));
            CompoundFeatureEntriesByRoleAndSize = CountByCode(workbook.CompoundFeatureRecords
                .SelectMany(record => record.EntryDetails)
                .Select(entry => $"{entry.Role}|{GetCompoundFeatureEntrySizeKey(entry)}"));
            CompoundVbaModulesByName = CountByCode(workbook.CompoundFeatureRecords.SelectMany(record => record.VbaModuleNames));
            CompoundVbaModulesBySize = CountByCode(workbook.CompoundFeatureRecords
                .SelectMany(record => record.EntryDetails)
                .Where(entry => entry.Role == LegacyXlsCompoundFeatureEntryRole.VbaModuleStream)
                .Select(GetCompoundFeatureEntrySizeKey));
            CompoundVbaModulesByCodeNameMatch = CountByCode(GetCompoundVbaModuleCodeNameMatchKeys(workbook));
            CompoundVbaProjectsByModuleCount = CountByCode(workbook.CompoundFeatureRecords
                .Where(record => record.Kind == LegacyXlsCompoundFeatureRecordKind.VbaProject)
                .Select(record => $"Modules:{record.VbaModuleCount}"));
            CompoundVbaProjectsByModuleByteCount = CountByCode(workbook.CompoundFeatureRecords
                .Where(record => record.Kind == LegacyXlsCompoundFeatureRecordKind.VbaProject)
                .Select(record => $"Bytes:{record.VbaModuleByteCount.ToString(CultureInfo.InvariantCulture)}"));
            CalculationSettingsByKind = CountCalculationSettingsByKind(workbook.CalculationSettings.Records);
            CellStylesByKind = CountByCode(workbook.CellStyles.Select(style => style.IsBuiltIn ? "BuiltIn" : "Custom"));
            CellStyleExtensionsByRecordName = CountByCode(workbook.CellStyleExtensions.Select(extension => extension.RecordName));
            CellStyleExtensionsByFormatIndex = CountByCode(workbook.CellStyleExtensions
                .Where(extension => extension.HasFormatIndex)
                .Select(extension => $"FormatIndex:{extension.FormatIndex}"));
            CellStyleExtensionsByExtensionCount = CountByCode(workbook.CellStyleExtensions
                .Where(extension => extension.HasExtensionCount)
                .Select(extension => $"Extensions:{extension.ExtensionCount}"));
            CellStyleExtensionsByStyleCategory = CountByCode(workbook.CellStyleExtensions
                .Where(extension => extension.StyleCategoryName is not null)
                .Select(extension => extension.StyleCategoryName!));
            CellStyleExtensionsByStyleFlags = CountByCode(workbook.CellStyleExtensions
                .Where(extension => extension.IsBuiltInStyle.HasValue && extension.IsHidden.HasValue && extension.IsCustom.HasValue)
                .Select(extension => $"BuiltIn:{extension.IsBuiltInStyle!.Value};Hidden:{extension.IsHidden!.Value};Custom:{extension.IsCustom!.Value}"));
            CellStyleExtensionsByStyleName = CountByCode(workbook.CellStyleExtensions
                .Where(extension => !string.IsNullOrWhiteSpace(extension.StyleName))
                .Select(extension => extension.StyleName!));
            CellStyleExtensionsByXfRecordCount = CountByCode(workbook.CellStyleExtensions
                .Where(extension => extension.XfRecordCount.HasValue)
                .Select(extension => $"XFs:{extension.XfRecordCount!.Value}"));
            CellStyleExtensionsByChecksum = CountByCode(workbook.CellStyleExtensions
                .Where(extension => extension.Checksum.HasValue)
                .Select(extension => $"Checksum:0x{extension.Checksum!.Value:X8}"));
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

        /// <summary>Gets imported data validations grouped by blank-value handling.</summary>
        public IReadOnlyDictionary<string, int> DataValidationsByAllowBlankState { get; }

        /// <summary>Gets imported data validations grouped by input-prompt display state.</summary>
        public IReadOnlyDictionary<string, int> DataValidationsByInputMessageState { get; }

        /// <summary>Gets imported data validations grouped by error-alert display state.</summary>
        public IReadOnlyDictionary<string, int> DataValidationsByErrorMessageState { get; }

        /// <summary>Gets imported data validations grouped by input prompt text presence.</summary>
        public IReadOnlyDictionary<string, int> DataValidationsByPromptTextState { get; }

        /// <summary>Gets imported data validations grouped by error alert text presence.</summary>
        public IReadOnlyDictionary<string, int> DataValidationsByErrorTextState { get; }

        /// <summary>Gets imported list data validations grouped by in-cell dropdown behavior.</summary>
        public IReadOnlyDictionary<string, int> DataValidationsByDropDownState { get; }

        /// <summary>Gets imported data validations grouped by number of covered ranges.</summary>
        public IReadOnlyDictionary<string, int> DataValidationsByRangeCount { get; }

        /// <summary>Gets imported data validations grouped by covered A1 range.</summary>
        public IReadOnlyDictionary<string, int> DataValidationsByRange { get; }

        /// <summary>Gets imported data validations grouped by first-formula presence.</summary>
        public IReadOnlyDictionary<string, int> DataValidationsByFormula1State { get; }

        /// <summary>Gets imported data validations grouped by second-formula presence.</summary>
        public IReadOnlyDictionary<string, int> DataValidationsByFormula2State { get; }

        /// <summary>Gets imported data validations grouped by combined first/second formula presence.</summary>
        public IReadOnlyDictionary<string, int> DataValidationsByFormulaPairState { get; }

        /// <summary>Gets imported list data validations grouped by source shape.</summary>
        public IReadOnlyDictionary<string, int> DataValidationListSourcesByKind { get; }

        /// <summary>Gets imported list data validations grouped by inline item count.</summary>
        public IReadOnlyDictionary<string, int> DataValidationListSourcesByItemCount { get; }

        /// <summary>Gets imported list data validations grouped by source range.</summary>
        public IReadOnlyDictionary<string, int> DataValidationListSourcesByRange { get; }

        /// <summary>Gets imported list data validations grouped by source defined name.</summary>
        public IReadOnlyDictionary<string, int> DataValidationListSourcesByName { get; }

        /// <summary>Gets imported list data validations grouped by source sheet name.</summary>
        public IReadOnlyDictionary<string, int> DataValidationListSourcesBySheetName { get; }

        /// <summary>Gets imported conditional formatting rules grouped by rule type.</summary>
        public IReadOnlyDictionary<string, int> ConditionalFormattingsByType { get; }

        /// <summary>Gets imported conditional formatting cell-is rules grouped by comparison operator.</summary>
        public IReadOnlyDictionary<string, int> ConditionalFormattingsByOperator { get; }

        /// <summary>Gets imported conditional formatting rules grouped by number of covered ranges.</summary>
        public IReadOnlyDictionary<string, int> ConditionalFormattingsByRangeCount { get; }

        /// <summary>Gets imported conditional formatting rules grouped by covered A1 range.</summary>
        public IReadOnlyDictionary<string, int> ConditionalFormattingsByRange { get; }

        /// <summary>Gets imported conditional formatting rules grouped by first-formula presence.</summary>
        public IReadOnlyDictionary<string, int> ConditionalFormattingsByFormula1State { get; }

        /// <summary>Gets imported conditional formatting rules grouped by second-formula presence.</summary>
        public IReadOnlyDictionary<string, int> ConditionalFormattingsByFormula2State { get; }

        /// <summary>Gets imported conditional formatting rules grouped by combined first/second formula presence.</summary>
        public IReadOnlyDictionary<string, int> ConditionalFormattingsByFormulaPairState { get; }

        /// <summary>Gets imported conditional formatting rules grouped by whether an extension priority was decoded.</summary>
        public IReadOnlyDictionary<string, int> ConditionalFormattingsByPriorityState { get; }

        /// <summary>Gets imported conditional formatting extension priorities grouped by priority value.</summary>
        public IReadOnlyDictionary<string, int> ConditionalFormattingsByPriority { get; }

        /// <summary>Gets imported conditional formatting rules grouped by stop-if-true behavior.</summary>
        public IReadOnlyDictionary<string, int> ConditionalFormattingsByStopIfTrueState { get; }

        /// <summary>Gets imported conditional formatting rules grouped by whether a differential format was attached.</summary>
        public IReadOnlyDictionary<string, int> ConditionalFormattingsByDifferentialFormatState { get; }

        /// <summary>Gets imported conditional formatting differential formats grouped by decoded fill shape.</summary>
        public IReadOnlyDictionary<string, int> ConditionalFormattingsByDifferentialFill { get; }

        /// <summary>Gets imported AutoFilter conditions grouped by comparison operator.</summary>
        public IReadOnlyDictionary<string, int> AutoFilterCriteriaByOperator { get; }

        /// <summary>Gets imported AutoFilter conditions grouped by BIFF operand kind.</summary>
        public IReadOnlyDictionary<string, int> AutoFilterCriteriaByValueKind { get; }

        /// <summary>Gets imported AutoFilter criteria grouped by condition join operator.</summary>
        public IReadOnlyDictionary<string, int> AutoFilterCriteriaByJoinOperator { get; }

        /// <summary>Gets imported AutoFilter criteria grouped by criteria kind.</summary>
        public IReadOnlyDictionary<string, int> AutoFilterCriteriaByKind { get; }

        /// <summary>Gets imported AutoFilter criteria grouped by zero-based column id.</summary>
        public IReadOnlyDictionary<string, int> AutoFilterCriteriaByColumn { get; }

        /// <summary>Gets imported AutoFilter criteria grouped by condition count.</summary>
        public IReadOnlyDictionary<string, int> AutoFilterCriteriaByConditionCount { get; }

        /// <summary>Gets imported Top/Bottom AutoFilter criteria grouped by top/bottom and items/percent shape.</summary>
        public IReadOnlyDictionary<string, int> AutoFilterTop10Kinds { get; }

        /// <summary>Gets imported Top/Bottom AutoFilter criteria grouped by shape and value.</summary>
        public IReadOnlyDictionary<string, int> AutoFilterTop10Values { get; }

        /// <summary>Gets imported Top/Bottom AutoFilter criteria grouped by top or bottom direction.</summary>
        public IReadOnlyDictionary<string, int> AutoFilterTop10Directions { get; }

        /// <summary>Gets imported Top/Bottom AutoFilter criteria grouped by item-count or percentage unit.</summary>
        public IReadOnlyDictionary<string, int> AutoFilterTop10Units { get; }

        /// <summary>Gets imported worksheets grouped by decoded BoundSheet visibility state.</summary>
        public IReadOnlyDictionary<string, int> WorksheetsByVisibility { get; }

        /// <summary>Gets whether the workbook CodeName record was present.</summary>
        public IReadOnlyDictionary<string, int> WorkbookCodeNameStates { get; }

        /// <summary>Gets workbook CodeName values grouped by name.</summary>
        public IReadOnlyDictionary<string, int> WorkbookCodeNames { get; }

        /// <summary>Gets decoded workbook option states from Backup and BookBool records.</summary>
        public IReadOnlyDictionary<string, int> WorkbookOptionStates { get; }

        /// <summary>Gets decoded BuiltInFnGroupCount values grouped by observed function category count.</summary>
        public IReadOnlyDictionary<string, int> WorkbookBuiltInFunctionGroupCounts { get; }

        /// <summary>Gets imported worksheets grouped by CodeName record presence.</summary>
        public IReadOnlyDictionary<string, int> WorksheetCodeNameStates { get; }

        /// <summary>Gets worksheet CodeName values grouped by name.</summary>
        public IReadOnlyDictionary<string, int> WorksheetCodeNames { get; }

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

        /// <summary>Gets the number of preserve-only DConRef source range records decoded during import.</summary>
        public int DataConsolidationReferenceCount { get; }

        /// <summary>Gets the number of preserve-only PivotTable BIFF records discovered during import.</summary>
        public int PivotTableRecordCount { get; }

        /// <summary>Gets the number of preserve-only chart BIFF records discovered during import.</summary>
        public int ChartRecordCount { get; }

        /// <summary>Gets the number of preserve-only drawing and object BIFF records discovered during import.</summary>
        public int DrawingRecordCount { get; }

        /// <summary>Gets the number of preserve-only workbook Theme records discovered during import.</summary>
        public int ThemeRecordCount { get; }

        /// <summary>Gets the number of OfficeArt record headers discovered under preserve-only drawing records.</summary>
        public int DrawingOfficeArtRecordCount { get; }

        /// <summary>Gets the number of OfficeArtFOPT shape property entries discovered under preserve-only drawing records.</summary>
        public int DrawingShapePropertyCount { get; }

        /// <summary>Gets the number of parsed differential formats discovered during import.</summary>
        public int DifferentialFormatCount { get; }

        /// <summary>Gets the number of preserve-only compound container features discovered during import.</summary>
        public int CompoundFeatureRecordCount { get; }

        /// <summary>Gets the number of matching compound directory entries behind preserve-only compound features.</summary>
        public int CompoundFeatureEntryCount { get; }

        /// <summary>Gets the number of VBA module streams discovered in preserve-only compound features.</summary>
        public int CompoundVbaModuleCount { get; }

        /// <summary>Gets the total declared byte size of matching preserve-only compound entries with known sizes.</summary>
        public long CompoundFeatureEntryByteCount { get; }

        /// <summary>Gets the total declared byte size of discovered VBA module streams with known sizes.</summary>
        public long CompoundVbaModuleByteCount { get; }

        /// <summary>Gets the number of calculation setting records parsed from BIFF records.</summary>
        public int CalculationSettingRecordCount { get; }

        /// <summary>Gets the number of workbook cell style records parsed from Style records.</summary>
        public int CellStyleRecordCount { get; }

        /// <summary>Gets the number of preserve-only style extension records parsed from XFExt records.</summary>
        public int CellStyleExtensionRecordCount { get; }

        /// <summary>Gets the number of parsed-formula token observations captured during import.</summary>
        public int FormulaTokenRecordCount { get; }

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

        /// <summary>Gets observed parsed-formula tokens grouped by BIFF token name.</summary>
        public IReadOnlyDictionary<string, int> FormulaTokensByName { get; }

        /// <summary>Gets observed parsed-formula tokens grouped by formula source context.</summary>
        public IReadOnlyDictionary<string, int> FormulaTokensByContext { get; }

        /// <summary>Gets observed parsed-formula tokens grouped by BIFF record type and token name.</summary>
        public IReadOnlyDictionary<string, int> FormulaTokensByRecordType { get; }

        /// <summary>Gets observed built-in formula function tokens grouped by raw function id.</summary>
        public IReadOnlyDictionary<string, int> FormulaFunctionsById { get; }

        /// <summary>Gets observed built-in formula function tokens grouped by function name when known.</summary>
        public IReadOnlyDictionary<string, int> FormulaFunctionsByName { get; }

        /// <summary>Gets observed PtgAttr formula tokens grouped by attribute name.</summary>
        public IReadOnlyDictionary<string, int> FormulaAttributesByName { get; }

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

        /// <summary>Gets unsupported sheet entries grouped by decoded BoundSheet visibility state.</summary>
        public IReadOnlyDictionary<string, int> UnsupportedSheetsByVisibility { get; }

        /// <summary>Gets unsupported sheet entries grouped by sheet kind and decoded visibility state.</summary>
        public IReadOnlyDictionary<string, int> UnsupportedSheetsByKindAndVisibility { get; }

        /// <summary>Gets unsupported chart sheets grouped by raw PrintSize value.</summary>
        public IReadOnlyDictionary<string, int> UnsupportedChartSheetPrintSizes { get; }

        /// <summary>Gets unsupported chart sheets grouped by decoded PrintSize mode name.</summary>
        public IReadOnlyDictionary<string, int> UnsupportedChartSheetPrintSizeKinds { get; }

        /// <summary>Gets unsupported chart sheets grouped by chart text object count.</summary>
        public IReadOnlyDictionary<string, int> UnsupportedChartSheetTextObjectCounts { get; }

        /// <summary>Gets unsupported chart sheets grouped by preserve-only chart record count.</summary>
        public IReadOnlyDictionary<string, int> UnsupportedChartSheetChartRecordCounts { get; }

        /// <summary>Gets unsupported chart sheet preserve-only chart records grouped by shallow category.</summary>
        public IReadOnlyDictionary<string, int> UnsupportedChartSheetChartRecordKinds { get; }

        /// <summary>Gets unsupported chart sheet preserve-only chart type records grouped by decoded chart family.</summary>
        public IReadOnlyDictionary<string, int> UnsupportedChartSheetChartTypes { get; }

        /// <summary>Gets preserved external references grouped by supporting-link kind.</summary>
        public IReadOnlyDictionary<LegacyXlsExternalReferenceKind, int> ExternalReferencesByKind { get; }

        /// <summary>Gets preserved external references grouped by target path or source.</summary>
        public IReadOnlyDictionary<string, int> ExternalReferencesByTarget { get; }

        /// <summary>Gets preserved external references grouped by their sheet/name/cache/cached-cell shape.</summary>
        public IReadOnlyDictionary<string, int> ExternalReferencesByShape { get; }

        /// <summary>Gets preserved external references grouped by declared SupBook sheet count.</summary>
        public IReadOnlyDictionary<string, int> ExternalReferencesByDeclaredSheetCount { get; }

        /// <summary>Gets preserved external references grouped by sheet-name count.</summary>
        public IReadOnlyDictionary<string, int> ExternalReferencesBySheetNameCount { get; }

        /// <summary>Gets preserved external references grouped by parsed sheet-name table completeness.</summary>
        public IReadOnlyDictionary<string, int> ExternalReferencesBySheetTableState { get; }

        /// <summary>Gets preserved external references grouped by external-name count.</summary>
        public IReadOnlyDictionary<string, int> ExternalReferencesByExternalNameCount { get; }

        /// <summary>Gets preserved external references grouped by cached cell section count.</summary>
        public IReadOnlyDictionary<string, int> ExternalReferencesByCacheCount { get; }

        /// <summary>Gets preserved external references grouped by cached cell value count.</summary>
        public IReadOnlyDictionary<string, int> ExternalReferencesByCachedCellCount { get; }

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

        /// <summary>Gets external cell cache sections grouped by occupied zero-based row/column range.</summary>
        public IReadOnlyDictionary<string, int> ExternalCellCachesByCellRange { get; }

        /// <summary>Gets external cell cache sections grouped by cached value count.</summary>
        public IReadOnlyDictionary<string, int> ExternalCellCachesByCellCount { get; }

        /// <summary>Gets external cell cache sections grouped by occupied row span.</summary>
        public IReadOnlyDictionary<string, int> ExternalCellCachesByRowSpan { get; }

        /// <summary>Gets external cell cache sections grouped by occupied column span.</summary>
        public IReadOnlyDictionary<string, int> ExternalCellCachesByColumnSpan { get; }

        /// <summary>Gets external cell cache sections grouped by XCT link-valid state.</summary>
        public IReadOnlyDictionary<string, int> ExternalCellCachesByLinkState { get; }

        /// <summary>Gets cached external cell values grouped by value kind.</summary>
        public IReadOnlyDictionary<LegacyXlsCellValueKind, int> ExternalCachedCellsByValueKind { get; }

        /// <summary>Gets DConRef records grouped by decoded DConFile source kind.</summary>
        public IReadOnlyDictionary<string, int> DataConsolidationReferencesBySourceKind { get; }

        /// <summary>Gets DConRef records grouped by decoded source path or sheet name.</summary>
        public IReadOnlyDictionary<string, int> DataConsolidationReferencesBySource { get; }

        /// <summary>Gets DConRef records grouped by decoded source range.</summary>
        public IReadOnlyDictionary<string, int> DataConsolidationReferencesByRange { get; }

        /// <summary>Gets DConRef records grouped by trailing unused byte count.</summary>
        public IReadOnlyDictionary<string, int> DataConsolidationReferencesByUnusedByteCount { get; }

        /// <summary>Gets Theme records grouped by decoded theme version.</summary>
        public IReadOnlyDictionary<string, int> ThemeRecordsByVersion { get; }

        /// <summary>Gets Theme records grouped by raw theme version value.</summary>
        public IReadOnlyDictionary<string, int> ThemeRecordsByRawVersion { get; }

        /// <summary>Gets Theme records grouped by whether embedded theme content bytes were present.</summary>
        public IReadOnlyDictionary<string, int> ThemeRecordsByContentState { get; }

        /// <summary>Gets Theme records grouped by embedded theme content byte length.</summary>
        public IReadOnlyDictionary<string, int> ThemeRecordsByContentLength { get; }

        /// <summary>Gets preserve-only PivotTable BIFF records grouped by decoded metadata kind.</summary>
        public IReadOnlyDictionary<LegacyXlsPivotTableRecordKind, int> PivotTableRecordsByKind { get; }

        /// <summary>Gets preserve-only PivotTable BIFF records grouped by record name.</summary>
        public IReadOnlyDictionary<string, int> PivotTableRecordsByName { get; }

        /// <summary>Gets decoded SXDI PivotTable data item records grouped by raw aggregation function identifier.</summary>
        public IReadOnlyDictionary<string, int> PivotTableDataItemAggregations { get; }

        /// <summary>Gets decoded SXDI PivotTable data item records grouped by aggregation function name.</summary>
        public IReadOnlyDictionary<string, int> PivotTableDataItemAggregationKinds { get; }

        /// <summary>Gets decoded SXDI PivotTable data item records grouped by pivot field index.</summary>
        public IReadOnlyDictionary<string, int> PivotTableDataItemFieldIndexes { get; }

        /// <summary>Gets decoded SXDI PivotTable data item records grouped by display calculation name.</summary>
        public IReadOnlyDictionary<string, int> PivotTableDataItemDisplayCalculations { get; }

        /// <summary>Gets decoded SXDI PivotTable data item records grouped by display-calculation field index.</summary>
        public IReadOnlyDictionary<string, int> PivotTableDataItemDisplayCalculationFieldIndexes { get; }

        /// <summary>Gets decoded SXDI PivotTable data item records grouped by display-calculation item index.</summary>
        public IReadOnlyDictionary<string, int> PivotTableDataItemDisplayCalculationItemIndexes { get; }

        /// <summary>Gets decoded SXDI PivotTable data item records grouped by number format identifier.</summary>
        public IReadOnlyDictionary<string, int> PivotTableDataItemNumberFormats { get; }

        /// <summary>Gets decoded SXDI PivotTable data item records grouped by custom data item name.</summary>
        public IReadOnlyDictionary<string, int> PivotTableDataItemNames { get; }

        /// <summary>Gets decoded SXRng PivotTable grouping records grouped by grouping kind.</summary>
        public IReadOnlyDictionary<string, int> PivotTableGroupingKinds { get; }

        /// <summary>Gets decoded SXRng PivotTable grouping records grouped by automatic boundary state.</summary>
        public IReadOnlyDictionary<string, int> PivotTableGroupingBoundaryStates { get; }

        /// <summary>Gets decoded SXRng numeric grouping records grouped by start, end, and interval values.</summary>
        public IReadOnlyDictionary<string, int> PivotTableGroupingNumericRanges { get; }

        /// <summary>Gets decoded SXRng date grouping records grouped by start, end, and interval values.</summary>
        public IReadOnlyDictionary<string, int> PivotTableGroupingDateRanges { get; }

        /// <summary>Gets decoded SXVDEx PivotTable field flags grouped by flag state.</summary>
        public IReadOnlyDictionary<string, int> PivotTableExtendedFieldStates { get; }

        /// <summary>Gets decoded SXAddl records grouped by PivotTable extension class.</summary>
        public IReadOnlyDictionary<string, int> PivotTableAdditionalClasses { get; }

        /// <summary>Gets decoded SXAddl records grouped by PivotTable extension detail type.</summary>
        public IReadOnlyDictionary<string, int> PivotTableAdditionalTypes { get; }

        /// <summary>Gets decoded SXAddl records grouped by class and detail type.</summary>
        public IReadOnlyDictionary<string, int> PivotTableAdditionalClassTypes { get; }

        /// <summary>Gets decoded SXAddl SxcCache/SXDId records grouped by PivotCache identifier.</summary>
        public IReadOnlyDictionary<string, int> PivotTableAdditionalCacheIds { get; }

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

        /// <summary>Gets Series records grouped by decoded value data type.</summary>
        public IReadOnlyDictionary<string, int> ChartSeriesValueDataTypes { get; }

        /// <summary>Gets Series records grouped by decoded bubble-size data type.</summary>
        public IReadOnlyDictionary<string, int> ChartSeriesBubbleSizeDataTypes { get; }

        /// <summary>Gets Series records grouped by category, value, and bubble-size counts.</summary>
        public IReadOnlyDictionary<string, int> ChartSeriesValueCounts { get; }

        /// <summary>Gets DataFormat records grouped by whether formatting targets a series or point.</summary>
        public IReadOnlyDictionary<string, int> ChartDataFormatTargets { get; }

        /// <summary>Gets DataFormat records grouped by raw series index.</summary>
        public IReadOnlyDictionary<string, int> ChartDataFormatSeriesIndexes { get; }

        /// <summary>Gets IFmtRecord records grouped by raw number format identifier.</summary>
        public IReadOnlyDictionary<string, int> ChartNumberFormatIds { get; }

        /// <summary>Gets FontX records grouped by raw font index.</summary>
        public IReadOnlyDictionary<string, int> ChartFontIndexes { get; }

        /// <summary>Gets Dat records grouped by decoded data-table display options.</summary>
        public IReadOnlyDictionary<string, int> ChartDataTableOptions { get; }

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

        /// <summary>Gets ValueRange records grouped by decoded value-axis scale fields.</summary>
        public IReadOnlyDictionary<string, int> ChartValueRangeScales { get; }

        /// <summary>Gets ValueRange records grouped by decoded automatic scale and axis-direction flags.</summary>
        public IReadOnlyDictionary<string, int> ChartValueRangeStates { get; }

        /// <summary>Gets Pos records grouped by decoded upper-left and lower-right position modes.</summary>
        public IReadOnlyDictionary<string, int> ChartPositionModePairs { get; }

        /// <summary>Gets Pos records grouped by decoded coordinate and size fields.</summary>
        public IReadOnlyDictionary<string, int> ChartPositionRectangles { get; }

        /// <summary>Gets Frame records grouped by decoded frame type.</summary>
        public IReadOnlyDictionary<string, int> ChartFrameTypes { get; }

        /// <summary>Gets Frame records grouped by automatic size and position flags.</summary>
        public IReadOnlyDictionary<string, int> ChartFrameAutoStates { get; }

        /// <summary>Gets PlotGrowth records grouped by decoded horizontal and vertical growth factors.</summary>
        public IReadOnlyDictionary<string, int> ChartPlotGrowthFactors { get; }

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

        /// <summary>Gets OBJ subrecords grouped by raw subrecord type.</summary>
        public IReadOnlyDictionary<string, int> DrawingObjectSubRecordsByType { get; }

        /// <summary>Gets OBJ subrecords grouped by decoded subrecord name.</summary>
        public IReadOnlyDictionary<string, int> DrawingObjectSubRecordsByName { get; }

        /// <summary>Gets OBJ subrecords grouped by declared payload length.</summary>
        public IReadOnlyDictionary<string, int> DrawingObjectSubRecordsByDeclaredLength { get; }

        /// <summary>Gets OBJ subrecords grouped by whether the declared payload was fully available.</summary>
        public IReadOnlyDictionary<string, int> DrawingObjectSubRecordsByCompleteness { get; }

        /// <summary>Gets MsoDrawing records grouped by decoded top-level Escher record type.</summary>
        public IReadOnlyDictionary<string, int> DrawingRecordsByEscherRecordType { get; }

        /// <summary>Gets MsoDrawing records grouped by decoded top-level Escher record type name.</summary>
        public IReadOnlyDictionary<string, int> DrawingRecordsByEscherRecordTypeName { get; }

        /// <summary>Gets nested OfficeArt records grouped by raw record type.</summary>
        public IReadOnlyDictionary<string, int> DrawingOfficeArtRecordsByType { get; }

        /// <summary>Gets nested OfficeArt records grouped by decoded record type name.</summary>
        public IReadOnlyDictionary<string, int> DrawingOfficeArtRecordsByTypeName { get; }

        /// <summary>Gets nested OfficeArt records grouped by traversal depth.</summary>
        public IReadOnlyDictionary<string, int> DrawingOfficeArtRecordsByDepth { get; }

        /// <summary>Gets nested OfficeArt records grouped by container or leaf state.</summary>
        public IReadOnlyDictionary<string, int> DrawingOfficeArtRecordsByContainerState { get; }

        /// <summary>Gets nested OfficeArt records grouped by declared payload length.</summary>
        public IReadOnlyDictionary<string, int> DrawingOfficeArtRecordsByPayloadLength { get; }

        /// <summary>Gets OfficeArtFOPT shape properties grouped by property identifier.</summary>
        public IReadOnlyDictionary<string, int> DrawingShapePropertiesById { get; }

        /// <summary>Gets OfficeArtFOPT shape properties grouped by decoded property name.</summary>
        public IReadOnlyDictionary<string, int> DrawingShapePropertiesByName { get; }

        /// <summary>Gets OfficeArtFOPT shape properties grouped by decoded property family.</summary>
        public IReadOnlyDictionary<string, int> DrawingShapePropertiesByGroup { get; }

        /// <summary>Gets OfficeArtFOPT shape properties grouped by complex and BLIP flag state.</summary>
        public IReadOnlyDictionary<string, int> DrawingShapePropertiesByFlagState { get; }

        /// <summary>Gets simple OfficeArtFOPT shape properties grouped by raw value.</summary>
        public IReadOnlyDictionary<string, int> DrawingShapePropertiesByValue { get; }

        /// <summary>Gets complex OfficeArtFOPT shape properties grouped by declared complex byte length.</summary>
        public IReadOnlyDictionary<string, int> DrawingShapeComplexPropertiesByDeclaredLength { get; }

        /// <summary>Gets complex OfficeArtFOPT shape properties grouped by available complex byte length.</summary>
        public IReadOnlyDictionary<string, int> DrawingShapeComplexPropertiesByAvailableLength { get; }

        /// <summary>Gets OfficeArt FBSE image-store entries grouped by decoded BLIP type.</summary>
        public IReadOnlyDictionary<string, int> DrawingBlipStoreEntriesByType { get; }

        /// <summary>Gets OfficeArt FBSE image-store entries grouped by embedded BLIP record type.</summary>
        public IReadOnlyDictionary<string, int> DrawingBlipStoreEntriesByEmbeddedRecordType { get; }

        /// <summary>Gets OfficeArt FBSE image-store entries grouped by stored byte size.</summary>
        public IReadOnlyDictionary<string, int> DrawingBlipStoreEntriesBySize { get; }

        /// <summary>Gets OfficeArt FBSE image-store entries grouped by reference count.</summary>
        public IReadOnlyDictionary<string, int> DrawingBlipStoreEntriesByReferenceCount { get; }

        /// <summary>Gets OfficeArt shape entries grouped by decoded shape type.</summary>
        public IReadOnlyDictionary<string, int> DrawingShapeEntriesByType { get; }

        /// <summary>Gets OfficeArt shape entries grouped by shape identifier.</summary>
        public IReadOnlyDictionary<string, int> DrawingShapeEntriesById { get; }

        /// <summary>Gets OfficeArt shape entries grouped by raw flag bitfield.</summary>
        public IReadOnlyDictionary<string, int> DrawingShapeEntriesByFlags { get; }

        /// <summary>Gets OfficeArt shape entries grouped by decoded flag name.</summary>
        public IReadOnlyDictionary<string, int> DrawingShapeEntriesByFlagName { get; }

        /// <summary>Gets OfficeArt client anchors grouped by start and end cell.</summary>
        public IReadOnlyDictionary<string, int> DrawingAnchorEntriesByRange { get; }

        /// <summary>Gets OfficeArt client anchors grouped by start and end offsets.</summary>
        public IReadOnlyDictionary<string, int> DrawingAnchorEntriesByOffset { get; }

        /// <summary>Gets OfficeArt client anchors grouped by raw flag bitfield.</summary>
        public IReadOnlyDictionary<string, int> DrawingAnchorEntriesByFlags { get; }

        /// <summary>Gets OfficeArt child anchors grouped by decoded rectangle.</summary>
        public IReadOnlyDictionary<string, int> DrawingChildAnchorEntriesByRectangle { get; }

        /// <summary>Gets OfficeArt child anchors grouped by decoded width and height.</summary>
        public IReadOnlyDictionary<string, int> DrawingChildAnchorEntriesBySize { get; }

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

        /// <summary>Gets matching compound feature entries grouped by OLE compound object type.</summary>
        public IReadOnlyDictionary<string, int> CompoundFeatureEntriesByObjectType { get; }

        /// <summary>Gets matching compound feature entries grouped by role and OLE compound object type.</summary>
        public IReadOnlyDictionary<string, int> CompoundFeatureEntriesByRoleAndObjectType { get; }

        /// <summary>Gets matching compound feature entries grouped by declared byte size.</summary>
        public IReadOnlyDictionary<string, int> CompoundFeatureEntriesBySize { get; }

        /// <summary>Gets matching compound feature entries grouped by role and declared byte size.</summary>
        public IReadOnlyDictionary<string, int> CompoundFeatureEntriesByRoleAndSize { get; }

        /// <summary>Gets VBA module streams grouped by module name.</summary>
        public IReadOnlyDictionary<string, int> CompoundVbaModulesByName { get; }

        /// <summary>Gets VBA module streams grouped by declared byte size.</summary>
        public IReadOnlyDictionary<string, int> CompoundVbaModulesBySize { get; }

        /// <summary>Gets VBA module streams grouped by whether they match workbook or worksheet CodeName records.</summary>
        public IReadOnlyDictionary<string, int> CompoundVbaModulesByCodeNameMatch { get; }

        /// <summary>Gets VBA project compound features grouped by discovered module count.</summary>
        public IReadOnlyDictionary<string, int> CompoundVbaProjectsByModuleCount { get; }

        /// <summary>Gets VBA project compound features grouped by total declared module stream bytes.</summary>
        public IReadOnlyDictionary<string, int> CompoundVbaProjectsByModuleByteCount { get; }

        /// <summary>Gets parsed calculation setting records grouped by setting kind.</summary>
        public IReadOnlyDictionary<LegacyXlsCalculationSettingKind, int> CalculationSettingsByKind { get; }

        /// <summary>Gets parsed workbook cell styles grouped by built-in/custom kind.</summary>
        public IReadOnlyDictionary<string, int> CellStylesByKind { get; }

        /// <summary>Gets preserve-only style extension records grouped by BIFF record name.</summary>
        public IReadOnlyDictionary<string, int> CellStyleExtensionsByRecordName { get; }

        /// <summary>Gets preserve-only style extension records grouped by extended XF index.</summary>
        public IReadOnlyDictionary<string, int> CellStyleExtensionsByFormatIndex { get; }

        /// <summary>Gets preserve-only style extension records grouped by declared extension-property count.</summary>
        public IReadOnlyDictionary<string, int> CellStyleExtensionsByExtensionCount { get; }

        /// <summary>Gets StyleExt records grouped by style category.</summary>
        public IReadOnlyDictionary<string, int> CellStyleExtensionsByStyleCategory { get; }

        /// <summary>Gets StyleExt records grouped by declared flag state.</summary>
        public IReadOnlyDictionary<string, int> CellStyleExtensionsByStyleFlags { get; }

        /// <summary>Gets StyleExt records grouped by style name.</summary>
        public IReadOnlyDictionary<string, int> CellStyleExtensionsByStyleName { get; }

        /// <summary>Gets XFCRC records grouped by declared XF record count.</summary>
        public IReadOnlyDictionary<string, int> CellStyleExtensionsByXfRecordCount { get; }

        /// <summary>Gets XFCRC records grouped by declared checksum.</summary>
        public IReadOnlyDictionary<string, int> CellStyleExtensionsByChecksum { get; }

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
            builder.AppendLine($"Data consolidation references: {DataConsolidationReferenceCount}");
            builder.AppendLine($"Pivot table records: {PivotTableRecordCount}");
            builder.AppendLine($"Chart records: {ChartRecordCount}");
            builder.AppendLine($"Drawing records: {DrawingRecordCount}");
            builder.AppendLine($"Theme records: {ThemeRecordCount}");
            builder.AppendLine($"Drawing OfficeArt records: {DrawingOfficeArtRecordCount}");
            builder.AppendLine($"Drawing shape properties: {DrawingShapePropertyCount}");
            builder.AppendLine($"Differential formats: {DifferentialFormatCount}");
            builder.AppendLine($"Compound feature records: {CompoundFeatureRecordCount}");
            builder.AppendLine($"Compound feature entries: {CompoundFeatureEntryCount}");
            builder.AppendLine($"Compound VBA modules: {CompoundVbaModuleCount}");
            builder.AppendLine($"Compound feature entry bytes: {CompoundFeatureEntryByteCount}");
            builder.AppendLine($"Compound VBA module bytes: {CompoundVbaModuleByteCount}");
            builder.AppendLine($"Calculation setting records: {CalculationSettingRecordCount}");
            builder.AppendLine($"Cell style records: {CellStyleRecordCount}");
            builder.AppendLine($"Cell style extension records: {CellStyleExtensionRecordCount}");
            builder.AppendLine($"Formula token records: {FormulaTokenRecordCount}");
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
            AppendDictionary(builder, "Formula Tokens By Name", FormulaTokensByName);
            AppendDictionary(builder, "Formula Tokens By Context", FormulaTokensByContext);
            AppendDictionary(builder, "Formula Tokens By Record Type", FormulaTokensByRecordType);
            AppendDictionary(builder, "Formula Functions By Id", FormulaFunctionsById);
            AppendDictionary(builder, "Formula Functions By Name", FormulaFunctionsByName);
            AppendDictionary(builder, "Formula Attributes By Name", FormulaAttributesByName);
            AppendDictionary(builder, "Data Validations By Type", DataValidationsByType);
            AppendDictionary(builder, "Data Validations By Operator", DataValidationsByOperator);
            AppendDictionary(builder, "Data Validations By Error Style", DataValidationsByErrorStyle);
            AppendDictionary(builder, "Data Validations By Allow Blank State", DataValidationsByAllowBlankState);
            AppendDictionary(builder, "Data Validations By Input Message State", DataValidationsByInputMessageState);
            AppendDictionary(builder, "Data Validations By Error Message State", DataValidationsByErrorMessageState);
            AppendDictionary(builder, "Data Validations By Prompt Text State", DataValidationsByPromptTextState);
            AppendDictionary(builder, "Data Validations By Error Text State", DataValidationsByErrorTextState);
            AppendDictionary(builder, "Data Validations By Drop Down State", DataValidationsByDropDownState);
            AppendDictionary(builder, "Data Validations By Range Count", DataValidationsByRangeCount);
            AppendDictionary(builder, "Data Validations By Range", DataValidationsByRange);
            AppendDictionary(builder, "Data Validations By Formula1 State", DataValidationsByFormula1State);
            AppendDictionary(builder, "Data Validations By Formula2 State", DataValidationsByFormula2State);
            AppendDictionary(builder, "Data Validations By Formula Pair State", DataValidationsByFormulaPairState);
            AppendDictionary(builder, "Data Validation List Sources By Kind", DataValidationListSourcesByKind);
            AppendDictionary(builder, "Data Validation List Sources By Item Count", DataValidationListSourcesByItemCount);
            AppendDictionary(builder, "Data Validation List Sources By Range", DataValidationListSourcesByRange);
            AppendDictionary(builder, "Data Validation List Sources By Name", DataValidationListSourcesByName);
            AppendDictionary(builder, "Data Validation List Sources By Sheet Name", DataValidationListSourcesBySheetName);
            AppendDictionary(builder, "Conditional Formatting By Type", ConditionalFormattingsByType);
            AppendDictionary(builder, "Conditional Formatting By Operator", ConditionalFormattingsByOperator);
            AppendDictionary(builder, "Conditional Formatting By Range Count", ConditionalFormattingsByRangeCount);
            AppendDictionary(builder, "Conditional Formatting By Range", ConditionalFormattingsByRange);
            AppendDictionary(builder, "Conditional Formatting By Formula1 State", ConditionalFormattingsByFormula1State);
            AppendDictionary(builder, "Conditional Formatting By Formula2 State", ConditionalFormattingsByFormula2State);
            AppendDictionary(builder, "Conditional Formatting By Formula Pair State", ConditionalFormattingsByFormulaPairState);
            AppendDictionary(builder, "Conditional Formatting By Priority State", ConditionalFormattingsByPriorityState);
            AppendDictionary(builder, "Conditional Formatting By Priority", ConditionalFormattingsByPriority);
            AppendDictionary(builder, "Conditional Formatting By Stop If True State", ConditionalFormattingsByStopIfTrueState);
            AppendDictionary(builder, "Conditional Formatting By Differential Format State", ConditionalFormattingsByDifferentialFormatState);
            AppendDictionary(builder, "Conditional Formatting By Differential Fill", ConditionalFormattingsByDifferentialFill);
            AppendDictionary(builder, "AutoFilter Criteria By Kind", AutoFilterCriteriaByKind);
            AppendDictionary(builder, "AutoFilter Criteria By Operator", AutoFilterCriteriaByOperator);
            AppendDictionary(builder, "AutoFilter Criteria By Value Kind", AutoFilterCriteriaByValueKind);
            AppendDictionary(builder, "AutoFilter Criteria By Join Operator", AutoFilterCriteriaByJoinOperator);
            AppendDictionary(builder, "AutoFilter Criteria By Column", AutoFilterCriteriaByColumn);
            AppendDictionary(builder, "AutoFilter Criteria By Condition Count", AutoFilterCriteriaByConditionCount);
            AppendDictionary(builder, "AutoFilter Top10 Kinds", AutoFilterTop10Kinds);
            AppendDictionary(builder, "AutoFilter Top10 Values", AutoFilterTop10Values);
            AppendDictionary(builder, "AutoFilter Top10 Directions", AutoFilterTop10Directions);
            AppendDictionary(builder, "AutoFilter Top10 Units", AutoFilterTop10Units);
            AppendDictionary(builder, "Worksheets By Visibility", WorksheetsByVisibility);
            AppendDictionary(builder, "Workbook CodeName States", WorkbookCodeNameStates);
            AppendDictionary(builder, "Workbook CodeNames", WorkbookCodeNames);
            AppendDictionary(builder, "Workbook Option States", WorkbookOptionStates);
            AppendDictionary(builder, "Workbook Built-In Function Group Counts", WorkbookBuiltInFunctionGroupCounts);
            AppendDictionary(builder, "Worksheet CodeName States", WorksheetCodeNameStates);
            AppendDictionary(builder, "Worksheet CodeNames", WorksheetCodeNames);
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
            AppendDictionary(builder, "Unsupported Sheets By Visibility", UnsupportedSheetsByVisibility);
            AppendDictionary(builder, "Unsupported Sheets By Kind And Visibility", UnsupportedSheetsByKindAndVisibility);
            AppendDictionary(builder, "Unsupported Chart Sheet Print Sizes", UnsupportedChartSheetPrintSizes);
            AppendDictionary(builder, "Unsupported Chart Sheet Print Size Kinds", UnsupportedChartSheetPrintSizeKinds);
            AppendDictionary(builder, "Unsupported Chart Sheet Text Object Counts", UnsupportedChartSheetTextObjectCounts);
            AppendDictionary(builder, "Unsupported Chart Sheet Chart Record Counts", UnsupportedChartSheetChartRecordCounts);
            AppendDictionary(builder, "Unsupported Chart Sheet Chart Record Kinds", UnsupportedChartSheetChartRecordKinds);
            AppendDictionary(builder, "Unsupported Chart Sheet Chart Types", UnsupportedChartSheetChartTypes);
            AppendDictionary(builder, "External References By Kind", ExternalReferencesByKind.ToDictionary(
                entry => entry.Key.ToString(),
                entry => entry.Value,
                StringComparer.OrdinalIgnoreCase));
            AppendDictionary(builder, "External References By Target", ExternalReferencesByTarget);
            AppendDictionary(builder, "External References By Shape", ExternalReferencesByShape);
            AppendDictionary(builder, "External References By Declared Sheet Count", ExternalReferencesByDeclaredSheetCount);
            AppendDictionary(builder, "External References By Sheet Name Count", ExternalReferencesBySheetNameCount);
            AppendDictionary(builder, "External References By Sheet Table State", ExternalReferencesBySheetTableState);
            AppendDictionary(builder, "External References By External Name Count", ExternalReferencesByExternalNameCount);
            AppendDictionary(builder, "External References By Cache Count", ExternalReferencesByCacheCount);
            AppendDictionary(builder, "External References By Cached Cell Count", ExternalReferencesByCachedCellCount);
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
            AppendDictionary(builder, "External Cell Caches By Cell Range", ExternalCellCachesByCellRange);
            AppendDictionary(builder, "External Cell Caches By Cell Count", ExternalCellCachesByCellCount);
            AppendDictionary(builder, "External Cell Caches By Row Span", ExternalCellCachesByRowSpan);
            AppendDictionary(builder, "External Cell Caches By Column Span", ExternalCellCachesByColumnSpan);
            AppendDictionary(builder, "External Cell Caches By Link State", ExternalCellCachesByLinkState);
            AppendDictionary(builder, "External Cached Cells By Value Kind", ExternalCachedCellsByValueKind.ToDictionary(
                entry => entry.Key.ToString(),
                entry => entry.Value,
                StringComparer.OrdinalIgnoreCase));
            AppendDictionary(builder, "Data Consolidation References By Source Kind", DataConsolidationReferencesBySourceKind);
            AppendDictionary(builder, "Data Consolidation References By Source", DataConsolidationReferencesBySource);
            AppendDictionary(builder, "Data Consolidation References By Range", DataConsolidationReferencesByRange);
            AppendDictionary(builder, "Data Consolidation References By Unused Byte Count", DataConsolidationReferencesByUnusedByteCount);
            AppendDictionary(builder, "Theme Records By Version", ThemeRecordsByVersion);
            AppendDictionary(builder, "Theme Records By Raw Version", ThemeRecordsByRawVersion);
            AppendDictionary(builder, "Theme Records By Content State", ThemeRecordsByContentState);
            AppendDictionary(builder, "Theme Records By Content Length", ThemeRecordsByContentLength);
            AppendDictionary(builder, "Pivot Table Records By Kind", PivotTableRecordsByKind.ToDictionary(
                entry => entry.Key.ToString(),
                entry => entry.Value,
                StringComparer.OrdinalIgnoreCase));
            AppendDictionary(builder, "Pivot Table Records By Name", PivotTableRecordsByName);
            AppendDictionary(builder, "Pivot Table Data Item Aggregations", PivotTableDataItemAggregations);
            AppendDictionary(builder, "Pivot Table Data Item Aggregation Kinds", PivotTableDataItemAggregationKinds);
            AppendDictionary(builder, "Pivot Table Data Item Field Indexes", PivotTableDataItemFieldIndexes);
            AppendDictionary(builder, "Pivot Table Data Item Display Calculations", PivotTableDataItemDisplayCalculations);
            AppendDictionary(builder, "Pivot Table Data Item Display Calculation Field Indexes", PivotTableDataItemDisplayCalculationFieldIndexes);
            AppendDictionary(builder, "Pivot Table Data Item Display Calculation Item Indexes", PivotTableDataItemDisplayCalculationItemIndexes);
            AppendDictionary(builder, "Pivot Table Data Item Number Formats", PivotTableDataItemNumberFormats);
            AppendDictionary(builder, "Pivot Table Data Item Names", PivotTableDataItemNames);
            AppendDictionary(builder, "Pivot Table Grouping Kinds", PivotTableGroupingKinds);
            AppendDictionary(builder, "Pivot Table Grouping Boundary States", PivotTableGroupingBoundaryStates);
            AppendDictionary(builder, "Pivot Table Grouping Numeric Ranges", PivotTableGroupingNumericRanges);
            AppendDictionary(builder, "Pivot Table Grouping Date Ranges", PivotTableGroupingDateRanges);
            AppendDictionary(builder, "Pivot Table Extended Field States", PivotTableExtendedFieldStates);
            AppendDictionary(builder, "Pivot Table Additional Classes", PivotTableAdditionalClasses);
            AppendDictionary(builder, "Pivot Table Additional Types", PivotTableAdditionalTypes);
            AppendDictionary(builder, "Pivot Table Additional Class Types", PivotTableAdditionalClassTypes);
            AppendDictionary(builder, "Pivot Table Additional Cache Ids", PivotTableAdditionalCacheIds);
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
            AppendDictionary(builder, "Chart Series Value Data Types", ChartSeriesValueDataTypes);
            AppendDictionary(builder, "Chart Series Bubble Size Data Types", ChartSeriesBubbleSizeDataTypes);
            AppendDictionary(builder, "Chart Series Value Counts", ChartSeriesValueCounts);
            AppendDictionary(builder, "Chart DataFormat Targets", ChartDataFormatTargets);
            AppendDictionary(builder, "Chart DataFormat Series Indexes", ChartDataFormatSeriesIndexes);
            AppendDictionary(builder, "Chart Number Format Ids", ChartNumberFormatIds);
            AppendDictionary(builder, "Chart Font Indexes", ChartFontIndexes);
            AppendDictionary(builder, "Chart DataTable Options", ChartDataTableOptions);
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
            AppendDictionary(builder, "Chart ValueRange Scales", ChartValueRangeScales);
            AppendDictionary(builder, "Chart ValueRange States", ChartValueRangeStates);
            AppendDictionary(builder, "Chart Position Mode Pairs", ChartPositionModePairs);
            AppendDictionary(builder, "Chart Position Rectangles", ChartPositionRectangles);
            AppendDictionary(builder, "Chart Frame Types", ChartFrameTypes);
            AppendDictionary(builder, "Chart Frame Auto States", ChartFrameAutoStates);
            AppendDictionary(builder, "Chart PlotGrowth Factors", ChartPlotGrowthFactors);
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
            AppendDictionary(builder, "Drawing Object Subrecords By Type", DrawingObjectSubRecordsByType);
            AppendDictionary(builder, "Drawing Object Subrecords By Name", DrawingObjectSubRecordsByName);
            AppendDictionary(builder, "Drawing Object Subrecords By Declared Length", DrawingObjectSubRecordsByDeclaredLength);
            AppendDictionary(builder, "Drawing Object Subrecords By Completeness", DrawingObjectSubRecordsByCompleteness);
            AppendDictionary(builder, "Drawing Records By Escher Record Type", DrawingRecordsByEscherRecordType);
            AppendDictionary(builder, "Drawing Records By Escher Record Type Name", DrawingRecordsByEscherRecordTypeName);
            AppendDictionary(builder, "Drawing OfficeArt Records By Type", DrawingOfficeArtRecordsByType);
            AppendDictionary(builder, "Drawing OfficeArt Records By Type Name", DrawingOfficeArtRecordsByTypeName);
            AppendDictionary(builder, "Drawing OfficeArt Records By Depth", DrawingOfficeArtRecordsByDepth);
            AppendDictionary(builder, "Drawing OfficeArt Records By Container State", DrawingOfficeArtRecordsByContainerState);
            AppendDictionary(builder, "Drawing OfficeArt Records By Payload Length", DrawingOfficeArtRecordsByPayloadLength);
            AppendDictionary(builder, "Drawing Shape Properties By Id", DrawingShapePropertiesById);
            AppendDictionary(builder, "Drawing Shape Properties By Name", DrawingShapePropertiesByName);
            AppendDictionary(builder, "Drawing Shape Properties By Group", DrawingShapePropertiesByGroup);
            AppendDictionary(builder, "Drawing Shape Properties By Flag State", DrawingShapePropertiesByFlagState);
            AppendDictionary(builder, "Drawing Shape Properties By Value", DrawingShapePropertiesByValue);
            AppendDictionary(builder, "Drawing Shape Complex Properties By Declared Length", DrawingShapeComplexPropertiesByDeclaredLength);
            AppendDictionary(builder, "Drawing Shape Complex Properties By Available Length", DrawingShapeComplexPropertiesByAvailableLength);
            AppendDictionary(builder, "Drawing BLIP Store Entries By Type", DrawingBlipStoreEntriesByType);
            AppendDictionary(builder, "Drawing BLIP Store Entries By Embedded Record Type", DrawingBlipStoreEntriesByEmbeddedRecordType);
            AppendDictionary(builder, "Drawing BLIP Store Entries By Size", DrawingBlipStoreEntriesBySize);
            AppendDictionary(builder, "Drawing BLIP Store Entries By Reference Count", DrawingBlipStoreEntriesByReferenceCount);
            AppendDictionary(builder, "Drawing Shape Entries By Type", DrawingShapeEntriesByType);
            AppendDictionary(builder, "Drawing Shape Entries By Id", DrawingShapeEntriesById);
            AppendDictionary(builder, "Drawing Shape Entries By Flags", DrawingShapeEntriesByFlags);
            AppendDictionary(builder, "Drawing Shape Entries By Flag Name", DrawingShapeEntriesByFlagName);
            AppendDictionary(builder, "Drawing Anchor Entries By Range", DrawingAnchorEntriesByRange);
            AppendDictionary(builder, "Drawing Anchor Entries By Offset", DrawingAnchorEntriesByOffset);
            AppendDictionary(builder, "Drawing Anchor Entries By Flags", DrawingAnchorEntriesByFlags);
            AppendDictionary(builder, "Drawing Child Anchor Entries By Rectangle", DrawingChildAnchorEntriesByRectangle);
            AppendDictionary(builder, "Drawing Child Anchor Entries By Size", DrawingChildAnchorEntriesBySize);
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
            AppendDictionary(builder, "Compound Feature Entries By Object Type", CompoundFeatureEntriesByObjectType);
            AppendDictionary(builder, "Compound Feature Entries By Role And Object Type", CompoundFeatureEntriesByRoleAndObjectType);
            AppendDictionary(builder, "Compound Feature Entries By Size", CompoundFeatureEntriesBySize);
            AppendDictionary(builder, "Compound Feature Entries By Role And Size", CompoundFeatureEntriesByRoleAndSize);
            AppendDictionary(builder, "Compound VBA Modules By Name", CompoundVbaModulesByName);
            AppendDictionary(builder, "Compound VBA Modules By Size", CompoundVbaModulesBySize);
            AppendDictionary(builder, "Compound VBA Modules By CodeName Match", CompoundVbaModulesByCodeNameMatch);
            AppendDictionary(builder, "Compound VBA Projects By Module Count", CompoundVbaProjectsByModuleCount);
            AppendDictionary(builder, "Compound VBA Projects By Module Byte Count", CompoundVbaProjectsByModuleByteCount);
            AppendDictionary(builder, "Calculation Settings By Kind", CalculationSettingsByKind.ToDictionary(
                entry => entry.Key.ToString(),
                entry => entry.Value,
                StringComparer.OrdinalIgnoreCase));
            AppendDictionary(builder, "Cell Styles By Kind", CellStylesByKind);
            AppendDictionary(builder, "Cell Style Extensions By Record Name", CellStyleExtensionsByRecordName);
            AppendDictionary(builder, "Cell Style Extensions By Format Index", CellStyleExtensionsByFormatIndex);
            AppendDictionary(builder, "Cell Style Extensions By Extension Count", CellStyleExtensionsByExtensionCount);
            AppendDictionary(builder, "Cell Style Extensions By Style Category", CellStyleExtensionsByStyleCategory);
            AppendDictionary(builder, "Cell Style Extensions By Style Flags", CellStyleExtensionsByStyleFlags);
            AppendDictionary(builder, "Cell Style Extensions By Style Name", CellStyleExtensionsByStyleName);
            AppendDictionary(builder, "Cell Style Extensions By XF Record Count", CellStyleExtensionsByXfRecordCount);
            AppendDictionary(builder, "Cell Style Extensions By Checksum", CellStyleExtensionsByChecksum);
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

        private static IEnumerable<string> GetWorkbookOptionStateKeys(LegacyXlsWorkbook workbook) {
            if (workbook.SaveBackup.HasValue) {
                yield return $"SaveBackup:{workbook.SaveBackup.Value}";
            }

            if (workbook.DoNotSaveExternalLinkValues.HasValue) {
                yield return $"DoNotSaveExternalLinkValues:{workbook.DoNotSaveExternalLinkValues.Value}";
            }

            if (workbook.HasEnvelope.HasValue) {
                yield return $"HasEnvelope:{workbook.HasEnvelope.Value}";
            }

            if (workbook.EnvelopeVisible.HasValue) {
                yield return $"EnvelopeVisible:{workbook.EnvelopeVisible.Value}";
            }

            if (workbook.EnvelopeInitialized.HasValue) {
                yield return $"EnvelopeInitialized:{workbook.EnvelopeInitialized.Value}";
            }

            if (workbook.ExternalLinkUpdateMode.HasValue) {
                yield return $"ExternalLinkUpdateMode:{workbook.ExternalLinkUpdateMode.Value}";
            }

            if (workbook.HideBordersForInactiveTables.HasValue) {
                yield return $"HideBordersForInactiveTables:{workbook.HideBordersForInactiveTables.Value}";
            }
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

        private static IReadOnlyDictionary<string, int> CountUnsupportedChartSheetChartRecordKinds(IEnumerable<LegacyXlsUnsupportedSheet> sheets) {
            return sheets
                .Where(sheet => sheet.Kind == LegacyXlsUnsupportedSheetKind.ChartSheet)
                .SelectMany(sheet => sheet.ChartRecordsByKind)
                .GroupBy(entry => entry.Key.ToString(), StringComparer.OrdinalIgnoreCase)
                .OrderBy(group => group.Key, StringComparer.OrdinalIgnoreCase)
                .ToDictionary(group => group.Key, group => group.Sum(entry => entry.Value), StringComparer.OrdinalIgnoreCase);
        }

        private static IReadOnlyDictionary<string, int> CountUnsupportedChartSheetChartTypes(IEnumerable<LegacyXlsUnsupportedSheet> sheets) {
            return sheets
                .Where(sheet => sheet.Kind == LegacyXlsUnsupportedSheetKind.ChartSheet)
                .SelectMany(sheet => sheet.ChartRecordsByChartType)
                .GroupBy(entry => entry.Key, StringComparer.OrdinalIgnoreCase)
                .OrderBy(group => group.Key, StringComparer.OrdinalIgnoreCase)
                .ToDictionary(group => group.Key, group => group.Sum(entry => entry.Value), StringComparer.OrdinalIgnoreCase);
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

        private static string GetCompoundFeatureEntrySizeKey(LegacyXlsCompoundFeatureEntryInfo entry) {
            return entry.SizeBytes.HasValue
                ? $"Bytes:{entry.SizeBytes.Value.ToString(CultureInfo.InvariantCulture)}"
                : "Bytes:Unknown";
        }

        private static IEnumerable<string> GetCompoundVbaModuleCodeNameMatchKeys(LegacyXlsWorkbook workbook) {
            var workbookCodeNames = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
            if (!string.IsNullOrWhiteSpace(workbook.CodeName)) {
                workbookCodeNames.Add(workbook.CodeName!);
            }

            var worksheetCodeNames = new HashSet<string>(
                workbook.Worksheets
                    .Where(sheet => !string.IsNullOrWhiteSpace(sheet.CodeName))
                    .Select(sheet => sheet.CodeName!),
                StringComparer.OrdinalIgnoreCase);

            foreach (string moduleName in workbook.CompoundFeatureRecords.SelectMany(record => record.VbaModuleNames)) {
                if (workbookCodeNames.Contains(moduleName)) {
                    yield return "WorkbookCodeName";
                } else if (worksheetCodeNames.Contains(moduleName)) {
                    yield return "WorksheetCodeName";
                } else {
                    yield return "UnmatchedCodeName";
                }
            }
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

        private static string GetChartValueRangeScaleKey(LegacyXlsChartRecord record) {
            LegacyXlsChartValueRange valueRange = record.ValueRange!;
            return $"Min:{FormatDouble(valueRange.Minimum)};Max:{FormatDouble(valueRange.Maximum)};Major:{FormatDouble(valueRange.MajorUnit)};Minor:{FormatDouble(valueRange.MinorUnit)};Cross:{FormatDouble(valueRange.CrossingValue)}";
        }

        private static string GetChartValueRangeStateKey(LegacyXlsChartRecord record) {
            LegacyXlsChartValueRange valueRange = record.ValueRange!;
            return $"AutoMin:{valueRange.AutoMinimum};AutoMax:{valueRange.AutoMaximum};AutoMajor:{valueRange.AutoMajorUnit};AutoMinor:{valueRange.AutoMinorUnit};AutoCross:{valueRange.AutoCrossingValue};Log:{valueRange.LogarithmicScale};Reversed:{valueRange.Reversed};MaxCross:{valueRange.MaximumCrossing}";
        }

        private static string FormatDouble(double value) {
            return value.ToString("G15", CultureInfo.InvariantCulture);
        }

        private static string GetDrawingRecordLocationKey(LegacyXlsDrawingRecord record) {
            return string.IsNullOrWhiteSpace(record.SheetName) ? "(workbook)" : record.SheetName!;
        }

        private static string GetShapePropertyFlagState(LegacyXlsDrawingShapeProperty property) {
            if (property.IsComplex && property.IsBlipId) {
                return "Complex|Blip";
            }

            if (property.IsComplex) {
                return "Complex";
            }

            return property.IsBlipId ? "Blip" : "Simple";
        }

        private static string GetExternalReferenceTargetKey(LegacyXlsExternalReference reference) {
            return string.IsNullOrWhiteSpace(reference.Target)
                ? $"({reference.Kind})"
                : EscapeControlCharacters(NormalizeExternalReferenceTarget(reference.Target!));
        }

        private static string NormalizeExternalReferenceTarget(string target) {
            return target.Length > 0 && target[0] == '\u0001'
                ? target.Substring(1)
                : target;
        }

        private static string GetExternalReferenceShapeKey(LegacyXlsExternalReference reference) {
            return $"{reference.Kind}|Sheets:{reference.SheetNameCount}|Names:{reference.ExternalNameCount}|Caches:{reference.CachedCellCacheCount}|CachedCells:{reference.CachedCellCount}";
        }

        private static string GetExternalReferenceSheetTableStateKey(LegacyXlsExternalReference reference) {
            if (reference.SheetCount == reference.SheetNameCount) {
                return $"Matched:{reference.SheetCount}";
            }

            return $"Declared:{reference.SheetCount};Parsed:{reference.SheetNameCount}";
        }

        private static string GetExternalCellCacheSheetKey(LegacyXlsExternalCellCache cache) {
            if (!string.IsNullOrWhiteSpace(cache.SheetName)) {
                return cache.SheetName!;
            }

            return cache.SheetIndex.HasValue ? $"SheetIndex:{cache.SheetIndex.Value}" : "(unknown)";
        }

        private static string GetExternalCellCacheRangeKey(LegacyXlsExternalCellCache cache) {
            return string.IsNullOrWhiteSpace(cache.CellRange) ? "(empty)" : cache.CellRange!;
        }

        private static string GetPivotTableGroupingNumericRangeKey(LegacyXlsPivotTableRecord record) {
            return $"Start:{FormatDouble(record.GroupingNumericStart!.Value)};End:{FormatDouble(record.GroupingNumericEnd!.Value)};Interval:{FormatDouble(record.GroupingNumericInterval!.Value)}";
        }

        private static string GetPivotTableGroupingDateRangeKey(LegacyXlsPivotTableRecord record) {
            return $"Start:{record.GroupingDateStart};End:{record.GroupingDateEnd};Interval:{record.GroupingDateInterval!.Value}";
        }

        private static string GetAutoFilterTop10KindKey(LegacyXlsAutoFilterCriteria criteria) {
            string rank = criteria.Top10IsTop ? "Top" : "Bottom";
            string unit = criteria.Top10IsPercent ? "Percent" : "Items";
            return rank + unit;
        }

        private static string GetFormulaStateKey(string? formula) {
            return string.IsNullOrWhiteSpace(formula) ? "Missing" : "Present";
        }

        private static string GetFormulaPairStateKey(string? formula1, string? formula2) {
            return $"Formula1:{GetFormulaStateKey(formula1)}|Formula2:{GetFormulaStateKey(formula2)}";
        }

        private static string GetDataValidationDropDownState(LegacyXlsDataValidation validation) {
            if (validation.Type != LegacyXlsDataValidationType.List) {
                return "NotList";
            }

            return validation.SuppressDropDown ? "Suppressed" : "Visible";
        }

        private static IEnumerable<string> GetPivotTableExtendedFieldStateKeys(LegacyXlsPivotTableRecord record) {
            if (!record.ShowAllItems.HasValue) {
                yield break;
            }

            yield return $"ShowAllItems:{record.ShowAllItems.Value}";
            yield return $"CanDragToRow:{record.CanDragToRow!.Value}";
            yield return $"CanDragToColumn:{record.CanDragToColumn!.Value}";
            yield return $"CanDragToPage:{record.CanDragToPage!.Value}";
            yield return $"CanDragToHide:{record.CanDragToHide!.Value}";
            yield return $"PreventDragToData:{record.PreventDragToData!.Value}";
            yield return $"ServerBased:{record.ServerBased!.Value}";
        }

        private static string GetPivotTableAdditionalClassTypeKey(LegacyXlsPivotTableRecord record) {
            return $"{record.AdditionalClassName}|{record.AdditionalTypeName}";
        }

        private static IEnumerable<string> GetConditionalFormattingDifferentialFillKeys(LegacyXlsConditionalFormatting formatting) {
            LegacyXlsDifferentialFormat? format = formatting.DifferentialFormat;
            if (format == null) {
                yield break;
            }

            if (format.FillPattern.HasValue) {
                yield return $"Pattern:{format.FillPattern.Value}";
            }

            if (!string.IsNullOrWhiteSpace(format.FillForegroundColor)) {
                yield return $"Foreground:{format.FillForegroundColor}";
            }

            if (!string.IsNullOrWhiteSpace(format.FillBackgroundColor)) {
                yield return $"Background:{format.FillBackgroundColor}";
            }
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
