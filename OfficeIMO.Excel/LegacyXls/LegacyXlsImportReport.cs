using OfficeIMO.Excel.LegacyXls.Diagnostics;
using OfficeIMO.Excel.LegacyXls.Biff;
using OfficeIMO.Excel.LegacyXls.Model;
using System.Globalization;
using System.Text;

namespace OfficeIMO.Excel.LegacyXls {
    /// <summary>
    /// Compact public import report for preflight checks and diagnostics.
    /// Detailed corpus aggregations are kept internal to avoid turning parser telemetry into a public compatibility surface.
    /// </summary>
    public sealed class LegacyXlsImportReport {
        internal LegacyXlsImportReport(LegacyXlsWorkbook workbook) {
            if (workbook == null) throw new ArgumentNullException(nameof(workbook));

            LegacyXlsComment[] comments = workbook.Worksheets.SelectMany(sheet => sheet.Comments).ToArray();
            LegacyXlsDataValidationCollectionRecord[] dataValidationCollections = workbook.Worksheets.SelectMany(sheet => sheet.DataValidationCollections).ToArray();
            LegacyXlsDataValidation[] dataValidations = workbook.Worksheets.SelectMany(sheet => sheet.DataValidations).ToArray();
            var phoneticSettings = workbook.Worksheets
                .Where(sheet => sheet.PhoneticSettings != null)
                .Select(sheet => new { SheetName = sheet.Name, Settings = sheet.PhoneticSettings! })
                .ToArray();
            var arrayFormulaRecords = workbook.Worksheets
                .SelectMany(sheet => sheet.ArrayFormulaRecords.Select(record => new { SheetName = sheet.Name, Record = record }))
                .ToArray();
            LegacyXlsConditionalFormatting[] conditionalFormattings = workbook.Worksheets.SelectMany(sheet => sheet.ConditionalFormattings).ToArray();
            LegacyXlsConditionalFormattingExtensionRecord[] conditionalFormattingExtensions = workbook.Worksheets.SelectMany(sheet => sheet.ConditionalFormattingExtensions).ToArray();
            LegacyXlsUnsupportedFeature[] unsupportedProjectionGaps = GetUnsupportedProjectionGaps(workbook);
            WorksheetCount = workbook.Worksheets.Count;
            ChartSheetCount = workbook.ChartSheets.Count;
            UnsupportedSheetCount = workbook.UnsupportedSheets.Count;
            CellCount = workbook.Worksheets.Sum(sheet => sheet.Cells.Count);
            FormulaCellCount = workbook.Worksheets.Sum(sheet => sheet.Cells.Count(cell => cell.IsFormula));
            CommentCount = comments.Length;
            HyperlinkCount = workbook.Worksheets.Sum(sheet => sheet.Hyperlinks.Count);
            CommentsByObjectType = CountByCode(comments
                .Where(comment => comment.ObjectType.HasValue)
                .Select(comment => $"ObjectType:0x{comment.ObjectType!.Value:X4}"));
            CommentsByObjectTypeName = CountByCode(comments
                .Where(comment => !string.IsNullOrWhiteSpace(comment.ObjectTypeName))
                .Select(comment => comment.ObjectTypeName!));
            CommentsByObjectFlags = CountByCode(comments
                .Where(comment => comment.ObjectFlags.HasValue)
                .Select(comment => $"ObjectFlags:0x{comment.ObjectFlags!.Value:X4}"));
            CommentsByObjectFlagName = CountByCode(comments.SelectMany(comment => comment.ObjectFlagNames));
            CommentsByAnchorRange = CountByCode(comments
                .Where(comment => comment.Anchor != null)
                .Select(comment => GetDrawingAnchorRangeKey(comment.Anchor!)));
            CommentsByAnchorOffset = CountByCode(comments
                .Where(comment => comment.Anchor != null)
                .Select(comment => GetDrawingAnchorOffsetKey(comment.Anchor!)));
            CommentsByAnchorFlags = CountByCode(comments
                .Where(comment => comment.Anchor != null)
                .Select(comment => GetDrawingAnchorFlagsKey(comment.Anchor!)));
            DataValidationCount = dataValidations.Length;
            DataValidationCollectionRecordCount = dataValidationCollections.Length;
            ConditionalFormattingCount = conditionalFormattings.Length;
            AutoFilterCriteriaCount = workbook.Worksheets.Sum(sheet => sheet.AutoFilterCriteria.Count);
            WorksheetFeatureStates = CountByCode(workbook.Worksheets.SelectMany(GetWorksheetFeatureStateKeys));
            WorksheetProtectionObjectStates = CountByCode(workbook.Worksheets
                .Where(sheet => sheet.Protection?.ProtectObjects.HasValue == true)
                .Select(sheet => sheet.Protection!.ProtectObjects!.Value ? "Protected" : "Unprotected"));
            WorksheetProtectionScenarioStates = CountByCode(workbook.Worksheets
                .Where(sheet => sheet.Protection?.ProtectScenarios.HasValue == true)
                .Select(sheet => sheet.Protection!.ProtectScenarios!.Value ? "Protected" : "Unprotected"));
            WorksheetPhoneticSettingsBySheet = CountByCode(phoneticSettings.Select(item => item.SheetName));
            WorksheetPhoneticSettingsByType = CountByCode(phoneticSettings.Select(item => item.Settings.Type.ToString()));
            WorksheetPhoneticSettingsByAlignment = CountByCode(phoneticSettings.Select(item => item.Settings.Alignment.ToString()));
            WorksheetPhoneticSettingsByFontId = CountByCode(phoneticSettings.Select(item => $"Font:{item.Settings.FontId.ToString(CultureInfo.InvariantCulture)}"));
            WorksheetPhoneticSettingsByRangeCount = CountByCode(phoneticSettings.Select(item => $"Ranges:{item.Settings.Ranges.Count.ToString(CultureInfo.InvariantCulture)}"));
            WorksheetPhoneticRangesBySheet = CountByCode(phoneticSettings
                .SelectMany(item => item.Settings.Ranges.Select(_ => item.SheetName)));
            WorksheetPhoneticRangesBySheetAndRange = CountByCode(phoneticSettings
                .SelectMany(item => item.Settings.Ranges.Select(range => $"{item.SheetName}!{range}")));
            DataValidationCollectionsBySheet = CountByCode(dataValidationCollections.Select(record => record.SheetName));
            DataValidationCollectionsByDeclaredCount = CountByCode(dataValidationCollections.Select(record => $"Declared:{record.DeclaredValidationCount.ToString(CultureInfo.InvariantCulture)}"));
            DataValidationCollectionStates = CountByCode(workbook.Worksheets.SelectMany(GetDataValidationCollectionStateKeys));
            DataValidationsByType = CountByCode(dataValidations.Select(validation => validation.Type.ToString()));
            DataValidationsByOperator = CountByCode(dataValidations.Select(validation => validation.Operator.ToString()));
            DataValidationsByErrorStyle = CountByCode(dataValidations.Select(validation => validation.ErrorStyle.ToString()));
            DataValidationsByAllowBlankState = CountByCode(dataValidations.Select(validation => validation.AllowBlank ? "AllowBlank" : "RejectBlank"));
            DataValidationsByInputMessageState = CountByCode(dataValidations.Select(validation => validation.ShowInputMessage ? "ShowInputMessage" : "HideInputMessage"));
            DataValidationsByErrorMessageState = CountByCode(dataValidations.Select(validation => validation.ShowErrorMessage ? "ShowErrorMessage" : "HideErrorMessage"));
            DataValidationsByPromptTextState = CountByCode(dataValidations.Select(validation => validation.PromptTitle != null || validation.Prompt != null ? "Present" : "Missing"));
            DataValidationsByErrorTextState = CountByCode(dataValidations.Select(validation => validation.ErrorTitle != null || validation.Error != null ? "Present" : "Missing"));
            DataValidationsByDropDownState = CountByCode(dataValidations.Select(GetDataValidationDropDownState));
            DataValidationsBySheet = CountByCode(workbook.Worksheets
                .SelectMany(sheet => sheet.DataValidations.Select(_ => sheet.Name)));
            DataValidationsByRangeCount = CountByCode(dataValidations.Select(validation => $"Ranges:{validation.RangeCount}"));
            DataValidationsByRange = CountByCode(dataValidations.SelectMany(validation => validation.Ranges));
            DataValidationsBySheetAndRange = CountByCode(workbook.Worksheets
                .SelectMany(sheet => sheet.DataValidations
                    .SelectMany(validation => validation.Ranges.Select(range => $"{sheet.Name}!{range}"))));
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
            ArrayFormulaRecordCount = arrayFormulaRecords.Length;
            ArrayFormulasBySheet = CountByCode(arrayFormulaRecords.Select(item => item.SheetName));
            ArrayFormulasByRange = CountByCode(arrayFormulaRecords.Select(item => item.Record.Range));
            ArrayFormulasBySheetAndRange = CountByCode(arrayFormulaRecords.Select(item => $"{item.SheetName}!{item.Record.Range}"));
            ArrayFormulasByDeclaredCellCount = CountByCode(arrayFormulaRecords.Select(item => $"Cells:{item.Record.DeclaredCellCount.ToString(CultureInfo.InvariantCulture)}"));
            ArrayFormulasByMatchedFormulaCellCount = CountByCode(arrayFormulaRecords.Select(item => $"Matched:{item.Record.MatchedFormulaCellCount.ToString(CultureInfo.InvariantCulture)}"));
            ArrayFormulasByAlwaysCalculateState = CountByCode(arrayFormulaRecords.Select(item => item.Record.AlwaysCalculate ? "AlwaysCalculate" : "NormalCalculation"));
            ArrayFormulasByProjectionState = CountByCode(arrayFormulaRecords.Select(item => item.Record.FormulaTextProjected ? "FormulaTextProjected" : "CachedOnly"));
            ArrayFormulasByTokenByteCount = CountByCode(arrayFormulaRecords.Select(item => $"TokenBytes:{item.Record.FormulaTokenByteCount.ToString(CultureInfo.InvariantCulture)}"));
            ArrayFormulasByExtraByteCount = CountByCode(arrayFormulaRecords.Select(item => $"ExtraBytes:{item.Record.FormulaExtraByteCount.ToString(CultureInfo.InvariantCulture)}"));
            ConditionalFormattingsByType = CountByCode(conditionalFormattings.Select(formatting => formatting.Type.ToString()));
            ConditionalFormattingsByOperator = CountByCode(conditionalFormattings
                .Where(formatting => formatting.Operator.HasValue)
                .Select(formatting => formatting.Operator!.Value.ToString()));
            ConditionalFormattingsBySheet = CountByCode(workbook.Worksheets
                .SelectMany(sheet => sheet.ConditionalFormattings.Select(_ => sheet.Name)));
            ConditionalFormattingsByRangeCount = CountByCode(conditionalFormattings.Select(formatting => $"Ranges:{formatting.RangeCount}"));
            ConditionalFormattingsByRange = CountByCode(conditionalFormattings.SelectMany(formatting => formatting.Ranges));
            ConditionalFormattingsBySheetAndRange = CountByCode(workbook.Worksheets
                .SelectMany(sheet => sheet.ConditionalFormattings
                    .SelectMany(formatting => formatting.Ranges.Select(range => $"{sheet.Name}!{range}"))));
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
            ConditionalFormattingsByDifferentialFont = CountByCode(conditionalFormattings.SelectMany(GetConditionalFormattingDifferentialFontKeys));
            ConditionalFormattingsByDifferentialBorder = CountByCode(conditionalFormattings.SelectMany(GetConditionalFormattingDifferentialBorderKeys));
            ConditionalFormattingsByDifferentialNumberFormat = CountByCode(conditionalFormattings.SelectMany(GetConditionalFormattingDifferentialNumberFormatKeys));
            ConditionalFormattingExtensionRecordCount = conditionalFormattingExtensions.Length;
            ConditionalFormattingExtensionsBySheet = CountByCode(conditionalFormattingExtensions.Select(record => record.SheetName));
            ConditionalFormattingExtensionsByRecordType = CountByCode(conditionalFormattingExtensions.Select(record => $"0x{record.RecordType:X4}"));
            ConditionalFormattingExtensionStates = CountByCode(conditionalFormattingExtensions.Select(GetConditionalFormattingExtensionStateKey));
            ConditionalFormattingExtensionPriorities = CountByCode(conditionalFormattingExtensions
                .Where(record => record.Priority.HasValue)
                .Select(record => $"Priority:{record.Priority!.Value}"));
            ConditionalFormattingExtensionStopIfTrueStates = CountByCode(conditionalFormattingExtensions.Select(GetConditionalFormattingExtensionStopIfTrueStateKey));
            ConditionalFormattingExtensionInlineFormattingByteCounts = CountByCode(conditionalFormattingExtensions
                .Where(record => record.InlineFormattingByteCount.HasValue)
                .Select(record => $"Bytes:{record.InlineFormattingByteCount!.Value.ToString(CultureInfo.InvariantCulture)}"));
            ConditionalFormattingExtensionDxfProjectionStates = CountByCode(conditionalFormattingExtensions
                .Select(record => GetConditionalFormattingExtensionDxfProjectionStateKey(record, workbook.DifferentialFormats.Count)));
            AutoFilterCriteriaBySheet = CountByCode(workbook.Worksheets
                .SelectMany(sheet => sheet.AutoFilterCriteria.Select(_ => sheet.Name)));
            AutoFilterCriteriaByOperator = CountByCode(workbook.Worksheets
                .SelectMany(sheet => sheet.AutoFilterCriteria)
                .SelectMany(criteria => criteria.Conditions)
                .Select(condition => condition.Operator.ToString()));
            AutoFilterCriteriaByValueKind = CountByCode(workbook.Worksheets
                .SelectMany(sheet => sheet.AutoFilterCriteria)
                .SelectMany(criteria => criteria.Conditions)
                .Select(condition => condition.ValueKind.ToString()));
            AutoFilterCriteriaByTextPattern = CountByCode(workbook.Worksheets
                .SelectMany(sheet => sheet.AutoFilterCriteria)
                .SelectMany(criteria => criteria.Conditions)
                .Where(condition => condition.TextPatternKind != LegacyXlsAutoFilterTextPatternKind.None)
                .Select(condition => condition.TextPatternKind.ToString()));
            AutoFilterCriteriaByJoinOperator = CountByCode(workbook.Worksheets
                .SelectMany(sheet => sheet.AutoFilterCriteria)
                .Select(criteria => criteria.JoinOperator.ToString()));
            AutoFilterCriteriaByKind = CountByCode(workbook.Worksheets
                .SelectMany(sheet => sheet.AutoFilterCriteria)
                .Select(criteria => criteria.Kind.ToString()));
            AutoFilterCriteriaByColumn = CountByCode(workbook.Worksheets
                .SelectMany(sheet => sheet.AutoFilterCriteria)
                .Select(criteria => $"Column:{criteria.ColumnId.ToString(CultureInfo.InvariantCulture)}"));
            AutoFilterCriteriaBySheetAndColumn = CountByCode(workbook.Worksheets
                .SelectMany(sheet => sheet.AutoFilterCriteria
                    .Select(criteria => $"{sheet.Name}!Column:{criteria.ColumnId.ToString(CultureInfo.InvariantCulture)}")));
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
            WorkbookFutureMetadataRecordsByKind = CountByCode(workbook.FutureMetadataRecords.Select(record => record.Kind.ToString()));
            WorkbookFutureMetadataRecordsByRecordType = CountByCode(workbook.FutureMetadataRecords.Select(record => $"0x{record.RecordType:X4}"));
            WorkbookFutureMetadataRecordsByRecordName = CountByCode(workbook.FutureMetadataRecords.Select(record => BiffUnsupportedRecordDiagnostics.GetBiffRecordName(record.RecordType)));
            WorkbookFutureMetadataRecordsByHeaderState = CountByCode(workbook.FutureMetadataRecords.Select(record => record.HeaderState));
            WorkbookFutureMetadataRecordsByHeaderRecordType = CountByCode(workbook.FutureMetadataRecords
                .Where(record => record.HeaderRecordType.HasValue)
                .Select(record => $"0x{record.HeaderRecordType!.Value:X4}"));
            WorkbookFutureMetadataRecordsByHeaderFlags = CountByCode(workbook.FutureMetadataRecords
                .Where(record => record.HeaderFlags.HasValue)
                .Select(record => $"Flags:0x{record.HeaderFlags!.Value:X4}"));
            WorkbookFutureMetadataRecordsByPayloadLength = CountByCode(workbook.FutureMetadataRecords.Select(record => $"Bytes:{record.PayloadLength}"));
            WorkbookFutureMetadataRecordsByBodyByteCount = CountByCode(workbook.FutureMetadataRecords.Select(record => $"Bytes:{record.BodyByteCount}"));
            WorksheetFutureMetadataRecordsByKind = CountByCode(workbook.Worksheets
                .SelectMany(sheet => sheet.FutureMetadataRecords)
                .Select(record => record.Kind.ToString()));
            WorksheetFutureMetadataRecordsBySheet = CountByCode(workbook.Worksheets
                .SelectMany(sheet => sheet.FutureMetadataRecords.Select(_ => sheet.Name)));
            WorksheetFutureMetadataRecordsBySheetAndKind = CountByCode(workbook.Worksheets
                .SelectMany(sheet => sheet.FutureMetadataRecords.Select(record => $"{sheet.Name}|{record.Kind}")));
            WorksheetFutureMetadataRecordsByRecordType = CountByCode(workbook.Worksheets
                .SelectMany(sheet => sheet.FutureMetadataRecords)
                .Select(record => $"0x{record.RecordType:X4}"));
            WorksheetFutureMetadataRecordsByRecordName = CountByCode(workbook.Worksheets
                .SelectMany(sheet => sheet.FutureMetadataRecords)
                .Select(record => BiffUnsupportedRecordDiagnostics.GetBiffRecordName(record.RecordType)));
            WorksheetFutureMetadataRecordsByHeaderState = CountByCode(workbook.Worksheets
                .SelectMany(sheet => sheet.FutureMetadataRecords)
                .Select(record => record.HeaderState));
            WorksheetFutureMetadataRecordsByHeaderRecordType = CountByCode(workbook.Worksheets
                .SelectMany(sheet => sheet.FutureMetadataRecords)
                .Where(record => record.HeaderRecordType.HasValue)
                .Select(record => $"0x{record.HeaderRecordType!.Value:X4}"));
            WorksheetFutureMetadataRecordsByHeaderFlags = CountByCode(workbook.Worksheets
                .SelectMany(sheet => sheet.FutureMetadataRecords)
                .Where(record => record.HeaderFlags.HasValue)
                .Select(record => $"Flags:0x{record.HeaderFlags!.Value:X4}"));
            WorksheetFutureMetadataRecordsByPayloadLength = CountByCode(workbook.Worksheets
                .SelectMany(sheet => sheet.FutureMetadataRecords)
                .Select(record => $"Bytes:{record.PayloadLength}"));
            WorksheetFutureMetadataRecordsByBodyByteCount = CountByCode(workbook.Worksheets
                .SelectMany(sheet => sheet.FutureMetadataRecords)
                .Select(record => $"Bytes:{record.BodyByteCount}"));
            UnsupportedSheetFutureMetadataRecordsByKind = CountByCode(workbook.UnsupportedSheets
                .SelectMany(sheet => sheet.FutureMetadataRecords)
                .Select(record => record.Kind.ToString()));
            UnsupportedSheetFutureMetadataRecordsBySheet = CountByCode(workbook.UnsupportedSheets
                .SelectMany(sheet => sheet.FutureMetadataRecords.Select(_ => sheet.Name)));
            UnsupportedSheetFutureMetadataRecordsBySheetAndKind = CountByCode(workbook.UnsupportedSheets
                .SelectMany(sheet => sheet.FutureMetadataRecords.Select(record => $"{sheet.Name}|{record.Kind}")));
            UnsupportedSheetFutureMetadataRecordsByRecordType = CountByCode(workbook.UnsupportedSheets
                .SelectMany(sheet => sheet.FutureMetadataRecords)
                .Select(record => $"0x{record.RecordType:X4}"));
            UnsupportedSheetFutureMetadataRecordsByRecordName = CountByCode(workbook.UnsupportedSheets
                .SelectMany(sheet => sheet.FutureMetadataRecords)
                .Select(record => BiffUnsupportedRecordDiagnostics.GetBiffRecordName(record.RecordType)));
            UnsupportedSheetFutureMetadataRecordsByHeaderState = CountByCode(workbook.UnsupportedSheets
                .SelectMany(sheet => sheet.FutureMetadataRecords)
                .Select(record => record.HeaderState));
            UnsupportedSheetFutureMetadataRecordsByHeaderRecordType = CountByCode(workbook.UnsupportedSheets
                .SelectMany(sheet => sheet.FutureMetadataRecords)
                .Where(record => record.HeaderRecordType.HasValue)
                .Select(record => $"0x{record.HeaderRecordType!.Value:X4}"));
            UnsupportedSheetFutureMetadataRecordsByHeaderFlags = CountByCode(workbook.UnsupportedSheets
                .SelectMany(sheet => sheet.FutureMetadataRecords)
                .Where(record => record.HeaderFlags.HasValue)
                .Select(record => $"Flags:0x{record.HeaderFlags!.Value:X4}"));
            UnsupportedSheetFutureMetadataRecordsByPayloadLength = CountByCode(workbook.UnsupportedSheets
                .SelectMany(sheet => sheet.FutureMetadataRecords)
                .Select(record => $"Bytes:{record.PayloadLength}"));
            UnsupportedSheetFutureMetadataRecordsByBodyByteCount = CountByCode(workbook.UnsupportedSheets
                .SelectMany(sheet => sheet.FutureMetadataRecords)
                .Select(record => $"Bytes:{record.BodyByteCount}"));
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
            ExternalQueryConnectionCount = workbook.ExternalQueryConnections.Count;
            DataConsolidationReferenceCount = workbook.DataConsolidationReferences.Count;
            DataConsolidationNameCount = workbook.DataConsolidationNames.Count;
            PivotTableRecordCount = workbook.PivotTableRecords.Count;
            ChartRecordCount = workbook.ChartRecords.Count;
            ChartSheetMetadataRecordCount = workbook.ChartSheets.Sum(sheet => sheet.MetadataRecords.Count);
            ChartSheetFutureMetadataRecordCount = workbook.ChartSheets.Sum(sheet => sheet.FutureMetadataRecords.Count);
            DrawingRecordCount = workbook.DrawingRecords.Count;
            ThemeRecordCount = workbook.ThemeRecords.Count;
            DrawingOfficeArtRecordCount = workbook.DrawingRecords.Sum(record => record.OfficeArtRecords.Count);
            DrawingGroupBlockCount = workbook.DrawingRecords.Sum(record => record.DrawingGroupBlocks.Count);
            DrawingGroupInfoCount = workbook.DrawingRecords.Sum(record => record.DrawingGroupInfos.Count);
            DrawingIdentifierClusterCount = workbook.DrawingRecords.Sum(record => record.DrawingGroupBlocks.Sum(block => block.IdentifierClusters.Count));
            DrawingShapePropertyCount = workbook.DrawingRecords.Sum(record => record.ShapeProperties.Count);
            DifferentialFormatCount = workbook.DifferentialFormats.Count;
            DifferentialFormatsByRecordType = CountByCode(workbook.DifferentialFormats
                .Select(format => $"RecordType:0x{format.RecordType:X4}"));
            DifferentialFormatsByContentState = CountByCode(workbook.DifferentialFormats
                .Select(GetDifferentialFormatContentStateKey));
            DifferentialFormatsByFill = CountByCode(workbook.DifferentialFormats.SelectMany(GetDifferentialFormatFillKeys));
            DifferentialFormatsByFont = CountByCode(workbook.DifferentialFormats.SelectMany(GetDifferentialFormatFontKeys));
            DifferentialFormatsByBorder = CountByCode(workbook.DifferentialFormats.SelectMany(GetDifferentialFormatBorderKeys));
            DifferentialFormatsByNumberFormat = CountByCode(workbook.DifferentialFormats.SelectMany(GetDifferentialFormatNumberFormatKeys));
            TableStyleCollectionRecordCount = workbook.TableStyleCollections.Count;
            TableStyleDefinitionCount = workbook.TableStyles.Count;
            TableStyleElementRecordCount = workbook.TableStyles.Sum(style => style.Elements.Count);
            TableStyleCollectionsByDefaultTableStyle = CountByCode(workbook.TableStyleCollections
                .Where(collection => !string.IsNullOrWhiteSpace(collection.DefaultTableStyleName))
                .Select(collection => collection.DefaultTableStyleName!));
            TableStyleCollectionsByDefaultPivotStyle = CountByCode(workbook.TableStyleCollections
                .Where(collection => !string.IsNullOrWhiteSpace(collection.DefaultPivotStyleName))
                .Select(collection => collection.DefaultPivotStyleName!));
            TableStyleCollectionsByTotalStyleCount = CountByCode(workbook.TableStyleCollections.Select(collection => $"Styles:{collection.TotalStyleCount}"));
            TableStylesByName = CountByCode(workbook.TableStyles.Select(style => style.Name));
            TableStylesByApplicability = CountByCode(workbook.TableStyles.Select(style =>
                style.AppliesToTables && style.AppliesToPivotTables
                    ? "TableAndPivot"
                    : style.AppliesToTables
                        ? "Table"
                        : style.AppliesToPivotTables
                            ? "Pivot"
                            : "None"));
            TableStylesByDeclaredElementCount = CountByCode(workbook.TableStyles.Select(style => $"Declared:{style.DeclaredElementCount}"));
            TableStylesByParsedElementCount = CountByCode(workbook.TableStyles.Select(style => $"Parsed:{style.Elements.Count}"));
            TableStyleElementsByType = CountByCode(workbook.TableStyles.SelectMany(style => style.Elements).Select(element => element.ElementTypeName));
            TableStyleElementsByDifferentialFormatIndex = CountByCode(workbook.TableStyles.SelectMany(style => style.Elements).Select(element => $"Dxf:{element.DifferentialFormatIndex}"));
            TableStyleElementsByStripeSize = CountByCode(workbook.TableStyles
                .SelectMany(style => style.Elements)
                .Where(element => element.ElementTypeName.Contains("Stripe", StringComparison.Ordinal))
                .Select(element => $"Size:{element.StripeSize}"));
            CompoundFeatureRecordCount = workbook.CompoundFeatureRecords.Count;
            CompoundFeatureEntryCount = workbook.CompoundFeatureRecords.Sum(record => record.Entries.Count);
            CompoundVbaModuleCount = workbook.CompoundFeatureRecords.Sum(record => record.VbaModuleCount);
            CompoundFeatureEntryByteCount = workbook.CompoundFeatureRecords.Sum(record => record.EntryByteCount);
            CompoundVbaModuleByteCount = workbook.CompoundFeatureRecords.Sum(record => record.VbaModuleByteCount);
            CalculationSettingRecordCount = workbook.CalculationSettings.Records.Count;
            CellStyleRecordCount = workbook.CellStyles.Count;
            CellStyleExtensionRecordCount = workbook.CellStyleExtensions.Count;
            FormulaTokenRecordCount = workbook.FormulaTokenRecords.Count;
            FutureFunctionAliasCount = workbook.FutureFunctionAliases.Count;
            WorkbookMetadataRecordCount = workbook.MetadataRecords.Count;
            WorkbookFutureMetadataRecordCount = workbook.FutureMetadataRecords.Count;
            WorksheetMetadataRecordCount = workbook.Worksheets.Sum(sheet => sheet.MetadataRecords.Count);
            WorksheetFutureMetadataRecordCount = workbook.Worksheets.Sum(sheet => sheet.FutureMetadataRecords.Count);
            UnsupportedSheetMetadataRecordCount = workbook.UnsupportedSheets.Sum(sheet => sheet.MetadataRecords.Count);
            UnsupportedSheetFutureMetadataRecordCount = workbook.UnsupportedSheets.Sum(sheet => sheet.FutureMetadataRecords.Count);
            UnsupportedFeatureCount = workbook.UnsupportedFeatures.Count;
            UnsupportedProjectionGapCount = unsupportedProjectionGaps.Length;
            PreservedFeatureRecordCount = workbook.PreservedFeatureRecords.Count;
            ErrorCount = workbook.Diagnostics.Count(diagnostic => diagnostic.Severity == LegacyXlsDiagnosticSeverity.Error);
            WarningCount = workbook.Diagnostics.Count(diagnostic => diagnostic.Severity == LegacyXlsDiagnosticSeverity.Warning);
            DiagnosticsByCode = CountByCode(workbook.Diagnostics.Select(diagnostic => diagnostic.Code));
            LegacyXlsImportDiagnostic[] formulaTokenDiagnostics = workbook.Diagnostics
                .Where(diagnostic => string.Equals(diagnostic.Code, "XLS-BIFF-FORMULA-TOKENS-UNSUPPORTED", StringComparison.OrdinalIgnoreCase))
                .ToArray();
            FormulaTokenBlockers = CountByCode(formulaTokenDiagnostics
                .Select(diagnostic => diagnostic.DetailCode ?? "FormulaUnknown"));
            FormulaTokenBlockersByToken = CountByCode(formulaTokenDiagnostics
                .Where(diagnostic => diagnostic.FormulaToken.HasValue)
                .Select(diagnostic => $"Token:0x{diagnostic.FormulaToken!.Value:X2}"));
            FormulaTokenBlockersByTokenName = CountByCode(formulaTokenDiagnostics
                .Where(diagnostic => !string.IsNullOrWhiteSpace(diagnostic.FormulaTokenName))
                .Select(diagnostic => diagnostic.FormulaTokenName!));
            FormulaTokenBlockersByOffset = CountByCode(formulaTokenDiagnostics
                .Where(diagnostic => diagnostic.FormulaTokenOffset.HasValue)
                .Select(diagnostic => $"Offset:{diagnostic.FormulaTokenOffset!.Value}"));
            FormulaTokenBlockersBySheet = CountByCode(formulaTokenDiagnostics
                .Select(GetDiagnosticSheetKey));
            FormulaTokenBlockersByContext = CountByCode(formulaTokenDiagnostics.Select(GetDiagnosticFormulaContextKey));
            FormulaTokenBlockersByContextAndToken = CountByCode(formulaTokenDiagnostics
                .Where(diagnostic => diagnostic.FormulaToken.HasValue)
                .Select(diagnostic => $"{GetDiagnosticFormulaContextKey(diagnostic)}|Token:0x{diagnostic.FormulaToken!.Value:X2}"));
            FormulaTokenBlockersByContextAndTokenName = CountByCode(formulaTokenDiagnostics
                .Where(diagnostic => !string.IsNullOrWhiteSpace(diagnostic.FormulaTokenName))
                .Select(diagnostic => $"{GetDiagnosticFormulaContextKey(diagnostic)}|{diagnostic.FormulaTokenName!}"));
            FormulaTokenBlockersByContextAndDetail = CountByCode(formulaTokenDiagnostics
                .Select(diagnostic => $"{GetDiagnosticFormulaContextKey(diagnostic)}|{diagnostic.DetailCode ?? "FormulaUnknown"}"));
            FormulaTokensByName = CountByCode(workbook.FormulaTokenRecords.Select(record => record.TokenName));
            FormulaTokensByContext = CountByCode(workbook.FormulaTokenRecords.Select(record => record.Context));
            FormulaTokensBySheet = CountByCode(workbook.FormulaTokenRecords.Select(GetFormulaTokenSheetKey));
            FormulaTokensByContextAndSheet = CountByCode(workbook.FormulaTokenRecords.Select(record => record.Context + "|" + GetFormulaTokenSheetKey(record)));
            FormulaTokensByContextAndOperandKind = CountByCode(workbook.FormulaTokenRecords
                .Where(record => !string.IsNullOrWhiteSpace(record.OperandKind))
                .Select(record => $"{record.Context}|{record.OperandKind!}"));
            FormulaTokensByRecordType = CountByCode(workbook.FormulaTokenRecords.Select(record => $"0x{record.RecordType:X4}|{record.TokenName}"));
            FormulaTokensByClass = CountByCode(workbook.FormulaTokenRecords
                .Where(record => !string.IsNullOrWhiteSpace(record.TokenClassName))
                .Select(record => record.TokenClassName!));
            FormulaTokensByNameAndClass = CountByCode(workbook.FormulaTokenRecords
                .Where(record => !string.IsNullOrWhiteSpace(record.TokenClassName))
                .Select(record => $"{record.TokenName}|{record.TokenClassName}"));
            FormulaTokensByOperandByteCount = CountByCode(workbook.FormulaTokenRecords
                .Where(record => record.OperandByteCount.HasValue)
                .Select(record => $"{record.TokenName}|Bytes:{record.OperandByteCount!.Value}"));
            FormulaTokensByOperandKind = CountByCode(workbook.FormulaTokenRecords
                .Where(record => !string.IsNullOrWhiteSpace(record.OperandKind))
                .Select(record => record.OperandKind!));
            FormulaTokensByNameAndOperandKind = CountByCode(workbook.FormulaTokenRecords
                .Where(record => !string.IsNullOrWhiteSpace(record.OperandKind))
                .Select(record => $"{record.TokenName}|{record.OperandKind!}"));
            FormulaTokensByOperandKindAndText = CountByCode(workbook.FormulaTokenRecords
                .Where(record => !string.IsNullOrWhiteSpace(record.OperandKind) && !string.IsNullOrWhiteSpace(record.OperandText))
                .Select(record => $"{record.OperandKind!}|{record.OperandText!}"));
            FormulaTokensByNameAndOperandText = CountByCode(workbook.FormulaTokenRecords
                .Where(record => !string.IsNullOrWhiteSpace(record.OperandText))
                .Select(record => $"{record.TokenName}|{record.OperandText!}"));
            FormulaTokensBySequenceIndex = CountByCode(workbook.FormulaTokenRecords
                .Where(record => record.SequenceIndex.HasValue)
                .Select(record => $"Index:{record.SequenceIndex!.Value}"));
            FormulaFunctionsById = CountByCode(workbook.FormulaTokenRecords
                .Where(record => record.FunctionId.HasValue)
                .Select(record => $"Function:0x{record.FunctionId!.Value:X4}"));
            FormulaFunctionsByName = CountByCode(workbook.FormulaTokenRecords
                .Where(record => !string.IsNullOrWhiteSpace(record.FunctionName))
                .Select(record => record.FunctionName!));
            FormulaFunctionsByParameterCount = CountByCode(workbook.FormulaTokenRecords
                .Where(record => record.FunctionParameterCount.HasValue)
                .Select(record => $"{record.FunctionName ?? $"Function:0x{record.FunctionId!.GetValueOrDefault():X4}"}|Args:{record.FunctionParameterCount!.Value}"));
            FormulaFunctionsByCetabState = CountByCode(workbook.FormulaTokenRecords
                .Where(record => record.FunctionIsCetab.HasValue)
                .Select(record => record.FunctionIsCetab!.Value ? "Cetab" : "BuiltIn"));
            FormulaAttributesByName = CountByCode(workbook.FormulaTokenRecords
                .Where(record => !string.IsNullOrWhiteSpace(record.AttributeName))
                .Select(record => record.AttributeName!));
            FutureFunctionAliasesByName = CountByCode(workbook.FutureFunctionAliases.Select(alias => alias.Name));
            FutureFunctionAliasesByFunction = CountByCode(workbook.FutureFunctionAliases.Select(alias => alias.FunctionName));
            FutureFunctionAliasesByTokenName = CountByCode(workbook.FutureFunctionAliases
                .Where(alias => !string.IsNullOrWhiteSpace(alias.FormulaTokenName))
                .Select(alias => alias.FormulaTokenName!));
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
            UnsupportedProjectionGapsByKind = CountByKind(unsupportedProjectionGaps);
            UnsupportedProjectionGapsByRecordType = CountByCode(unsupportedProjectionGaps
                .Where(feature => feature.RecordType.HasValue)
                .Select(feature => $"{feature.Kind}|{feature.Code}|0x{feature.RecordType!.Value:X4}"));
            UnsupportedProjectionGapsByDetail = CountByCode(unsupportedProjectionGaps
                .Where(feature => !string.IsNullOrWhiteSpace(feature.DetailCode))
                .Select(feature => $"{feature.Kind}|{feature.Code}|{feature.DetailCode}"));
            FileFormatStates = CountByCode(GetFileFormatStateKeys(workbook));
            FileFormatBlockers = CountByCode(workbook.UnsupportedFeatures
                .Where(IsFileFormatBlocker)
                .Select(feature => $"{feature.Kind}|{feature.DetailCode ?? feature.Code}"));
            FileFormatBlockersByRecordType = CountByCode(workbook.UnsupportedFeatures
                .Where(IsFileFormatBlocker)
                .Where(feature => feature.RecordType.HasValue)
                .Select(feature => $"{feature.Kind}|0x{feature.RecordType!.Value:X4}"));
            FileFormatBlockersByRecordName = CountByCode(workbook.UnsupportedFeatures
                .Where(IsFileFormatBlocker)
                .Where(feature => feature.RecordType.HasValue)
                .Select(feature => $"{feature.Kind}|{BiffUnsupportedRecordDiagnostics.GetBiffRecordName(feature.RecordType!.Value)}"));
            FileFormatBlockersByLocation = CountByCode(workbook.UnsupportedFeatures
                .Where(IsFileFormatBlocker)
                .Select(GetFeatureLocationKey));
            EncryptedWorkbooksByMethod = CountByCode(workbook.UnsupportedFeatures
                .Where(feature => feature.Kind == LegacyXlsUnsupportedFeatureKind.EncryptedWorkbook)
                .Select(GetEncryptionMethodKey));
            UnsupportedBiffVersionsByVersion = CountByCode(workbook.UnsupportedFeatures
                .Where(feature => feature.Kind == LegacyXlsUnsupportedFeatureKind.UnsupportedBiffVersion)
                .Select(GetBiffVersionKey));
            UnsupportedBiffVersionsBySubstream = CountByCode(workbook.UnsupportedFeatures
                .Where(feature => feature.Kind == LegacyXlsUnsupportedFeatureKind.UnsupportedBiffVersion)
                .Select(GetBiffSubstreamKey));
            UnsupportedBiffVersionsByVersionAndSubstream = CountByCode(workbook.UnsupportedFeatures
                .Where(feature => feature.Kind == LegacyXlsUnsupportedFeatureKind.UnsupportedBiffVersion)
                .Select(feature => $"{GetBiffVersionKey(feature)}|{GetBiffSubstreamKey(feature)}"));
            UnsupportedSheetsByKind = CountUnsupportedSheetsByKind(workbook.UnsupportedSheets);
            UnsupportedSheetsByType = CountByCode(workbook.UnsupportedSheets.Select(sheet => $"0x{sheet.SheetType:X2}|{sheet.Kind}"));
            UnsupportedSheetsByName = CountByCode(workbook.UnsupportedSheets.Select(sheet => sheet.Name));
            UnsupportedSheetsByVisibility = CountByCode(workbook.UnsupportedSheets.Select(sheet => sheet.VisibilityName));
            UnsupportedSheetsByKindAndVisibility = CountByCode(workbook.UnsupportedSheets.Select(sheet => $"{sheet.Kind}|{sheet.VisibilityName}"));
            ChartSheetsByType = CountByCode(workbook.ChartSheets.Select(sheet => $"0x{sheet.SheetType:X2}|ChartSheet"));
            ChartSheetsByName = CountByCode(workbook.ChartSheets.Select(sheet => sheet.Name));
            ChartSheetsByVisibility = CountByCode(workbook.ChartSheets.Select(sheet => sheet.VisibilityName));
            ChartSheetMetadataRecordsByKind = CountByCode(workbook.ChartSheets
                .SelectMany(sheet => sheet.MetadataRecords)
                .Select(record => record.Kind.ToString()));
            ChartSheetFutureMetadataRecordsByRecordType = CountByCode(workbook.ChartSheets
                .SelectMany(sheet => sheet.FutureMetadataRecords)
                .Select(record => $"0x{record.RecordType:X4}"));
            ChartSheetPrintSizes = CountByCode(workbook.ChartSheets
                .Where(sheet => sheet.ChartPrintSize.HasValue)
                .Select(sheet => $"PrintSize:{sheet.ChartPrintSize!.Value}"));
            ChartSheetPrintSizeKinds = CountByCode(workbook.ChartSheets
                .Where(sheet => !string.IsNullOrWhiteSpace(sheet.ChartPrintSizeName))
                .Select(sheet => sheet.ChartPrintSizeName!));
            ChartSheetTextObjectCounts = CountByCode(workbook.ChartSheets
                .Where(sheet => sheet.ChartTextObjectCount > 0)
                .Select(sheet => $"TextObjects:{sheet.ChartTextObjectCount}"));
            ChartSheetChartRecordCounts = CountByCode(workbook.ChartSheets
                .Where(sheet => sheet.ChartRecordCount > 0)
                .Select(sheet => $"ChartRecords:{sheet.ChartRecordCount}"));
            ChartSheetChartRecordCountsBySheet = CountByCode(workbook.ChartSheets
                .Where(sheet => sheet.ChartRecordCount > 0)
                .Select(sheet => $"Sheet:{sheet.Name};ChartRecords:{sheet.ChartRecordCount}"));
            ChartSheetChartRecordKinds = CountChartSheetChartRecordKinds(workbook.ChartSheets);
            ChartSheetChartRecordKindsBySheet = CountChartSheetChartRecordKindsBySheet(workbook.ChartSheets);
            ChartSheetChartTypes = CountChartSheetChartTypes(workbook.ChartSheets);
            ChartSheetChartTypesBySheet = CountChartSheetChartTypesBySheet(workbook.ChartSheets);
            ChartSheetStates = CountByCode(workbook.ChartSheets.Select(GetChartSheetStateKey));
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
            UnsupportedChartSheetChartRecordCountsBySheet = CountByCode(workbook.UnsupportedSheets
                .Where(sheet => sheet.Kind == LegacyXlsUnsupportedSheetKind.ChartSheet && sheet.ChartRecordCount > 0)
                .Select(sheet => $"Sheet:{sheet.Name};ChartRecords:{sheet.ChartRecordCount}"));
            UnsupportedChartSheetChartRecordKinds = CountUnsupportedChartSheetChartRecordKinds(workbook.UnsupportedSheets);
            UnsupportedChartSheetChartRecordKindsBySheet = CountUnsupportedChartSheetChartRecordKindsBySheet(workbook.UnsupportedSheets);
            UnsupportedChartSheetChartTypes = CountUnsupportedChartSheetChartTypes(workbook.UnsupportedSheets);
            UnsupportedChartSheetChartTypesBySheet = CountUnsupportedChartSheetChartTypesBySheet(workbook.UnsupportedSheets);
            UnsupportedChartSheetStates = CountByCode(workbook.UnsupportedSheets
                .Where(sheet => sheet.Kind == LegacyXlsUnsupportedSheetKind.ChartSheet)
                .Select(GetUnsupportedChartSheetStateKey));
            ExternalReferencesByKind = CountExternalReferencesByKind(workbook.ExternalReferences);
            ExternalReferencesByTarget = CountByCode(workbook.ExternalReferences.Select(GetExternalReferenceTargetKey));
            ExternalReferencesByShape = CountByCode(workbook.ExternalReferences.Select(GetExternalReferenceShapeKey));
            ExternalReferenceWorkbookStates = CountByCode(GetExternalReferenceWorkbookStateKeys(workbook.ExternalReferences));
            ExternalReferencesByDeclaredSheetCount = CountByCode(workbook.ExternalReferences.Select(reference => $"DeclaredSheets:{reference.SheetCount}"));
            ExternalReferencesBySheetNameCount = CountByCode(workbook.ExternalReferences.Select(reference => $"Sheets:{reference.SheetNameCount}"));
            ExternalReferencesBySheetTableState = CountByCode(workbook.ExternalReferences.Select(GetExternalReferenceSheetTableStateKey));
            ExternalReferencesByExternalNameCount = CountByCode(workbook.ExternalReferences.Select(reference => $"Names:{reference.ExternalNameCount}"));
            ExternalReferencesByCacheCount = CountByCode(workbook.ExternalReferences.Select(reference => $"Caches:{reference.CachedCellCacheCount}"));
            ExternalReferencesByCachedCellCount = CountByCode(workbook.ExternalReferences.Select(reference => $"CachedCells:{reference.CachedCellCount}"));
            ExternalSheetNamesByReferenceKind = CountExternalSheetNamesByReferenceKind(workbook.ExternalReferences);
            ExternalSheetNamesByTarget = CountByCode(workbook.ExternalReferences
                .SelectMany(reference => reference.SheetNames.Select(sheetName => $"{GetExternalReferenceTargetKey(reference)}!{sheetName}")));
            ExternalNamesByReferenceKind = CountExternalNamesByReferenceKind(workbook.ExternalReferences);
            ExternalNamesByName = CountByCode(workbook.ExternalReferences.SelectMany(reference => reference.ExternalNames.Select(name => name.Name)));
            ExternalNamesByScope = CountByCode(workbook.ExternalReferences
                .SelectMany(reference => reference.ExternalNames)
                .Select(name => name.LocalSheetIndex.HasValue ? "SheetLocal" : "Workbook"));
            ExternalNamesByBuiltInState = CountByCode(workbook.ExternalReferences
                .SelectMany(reference => reference.ExternalNames)
                .Select(name => name.BuiltIn ? "BuiltIn" : "Custom"));
            ExternalNamesByBodyKind = CountByCode(workbook.ExternalReferences
                .SelectMany(reference => reference.ExternalNames)
                .Select(name => name.BodyKind.ToString()));
            ExternalNamesByCachedClipboardFormat = CountByCode(workbook.ExternalReferences
                .SelectMany(reference => reference.ExternalNames)
                .Select(name => $"{name.CachedClipboardFormatName}:{name.CachedClipboardFormat}"));
            ExternalNamesByAdviseState = CountByCode(workbook.ExternalReferences
                .SelectMany(reference => reference.ExternalNames)
                .Select(name => name.WantsAdvise ? "Present" : "Missing"));
            ExternalNamesByPictureState = CountByCode(workbook.ExternalReferences
                .SelectMany(reference => reference.ExternalNames)
                .Select(name => name.WantsPicture ? "Present" : "Missing"));
            ExternalNamesByOleState = CountByCode(workbook.ExternalReferences
                .SelectMany(reference => reference.ExternalNames)
                .Select(name => name.Ole ? "Present" : "Missing"));
            ExternalNamesByOleLinkState = CountByCode(workbook.ExternalReferences
                .SelectMany(reference => reference.ExternalNames)
                .Select(name => name.OleLink ? "Present" : "Missing"));
            ExternalNamesByIconState = CountByCode(workbook.ExternalReferences
                .SelectMany(reference => reference.ExternalNames)
                .Select(name => name.Icon ? "Present" : "Missing"));
            ExternalNamesByFlagShape = CountByCode(workbook.ExternalReferences
                .SelectMany(reference => reference.ExternalNames)
                .Select(GetExternalNameFlagShapeKey));
            ExternalCellCachesByTarget = CountByCode(workbook.ExternalReferences
                .SelectMany(reference => reference.CachedCellCaches.Select(_ => GetExternalReferenceTargetKey(reference))));
            ExternalCellCachesBySheetName = CountByCode(workbook.ExternalReferences.SelectMany(reference => reference.CachedCellCaches.Select(GetExternalCellCacheSheetKey)));
            ExternalCellCachesByTargetAndSheetName = CountByCode(workbook.ExternalReferences
                .SelectMany(reference => reference.CachedCellCaches
                    .Select(cache => $"{GetExternalReferenceTargetKey(reference)}!{GetExternalCellCacheSheetKey(cache)}")));
            ExternalCellCachesByCellRange = CountByCode(workbook.ExternalReferences.SelectMany(reference => reference.CachedCellCaches.Select(GetExternalCellCacheRangeKey)));
            ExternalCellCachesByTargetAndCellRange = CountByCode(workbook.ExternalReferences
                .SelectMany(reference => reference.CachedCellCaches
                    .Select(cache => $"{GetExternalReferenceTargetKey(reference)}!{GetExternalCellCacheRangeKey(cache)}")));
            ExternalCellCachesByCellCount = CountByCode(workbook.ExternalReferences.SelectMany(reference => reference.CachedCellCaches.Select(cache => $"Cells:{cache.Cells.Count}")));
            ExternalCellCachesByRowSpan = CountByCode(workbook.ExternalReferences.SelectMany(reference => reference.CachedCellCaches.Select(cache => cache.RowSpan.HasValue ? $"Rows:{cache.RowSpan.Value}" : "(empty)")));
            ExternalCellCachesByColumnSpan = CountByCode(workbook.ExternalReferences.SelectMany(reference => reference.CachedCellCaches.Select(cache => cache.ColumnSpan.HasValue ? $"Columns:{cache.ColumnSpan.Value}" : "(empty)")));
            ExternalCellCachesByLinkState = CountByCode(workbook.ExternalReferences.SelectMany(reference => reference.CachedCellCaches.Select(cache => cache.LinkValid ? "ValidLink" : "InvalidLink")));
            ExternalCachedCellsByValueKind = CountExternalCachedCellsByValueKind(workbook.ExternalReferences);
            ExternalCachedCellsByTargetSheetAndValueKind = CountByCode(workbook.ExternalReferences
                .SelectMany(reference => reference.CachedCellCaches
                    .SelectMany(cache => cache.Cells
                        .Select(cell => $"{GetExternalReferenceTargetKey(reference)}!{GetExternalCellCacheSheetKey(cache)}|{cell.Kind}"))));
            ExternalQueryConnectionsBySourceType = CountByCode(workbook.ExternalQueryConnections.Select(connection => connection.SourceTypeName));
            ExternalQueryConnectionsByState = CountByCode(workbook.ExternalQueryConnections.Select(GetExternalQueryConnectionStateKey));
            ExternalQueryConnectionsByConnectionFlag = CountByCode(workbook.ExternalQueryConnections.SelectMany(GetExternalQueryConnectionFlagKeys));
            ExternalQueryConnectionsByQueryOption = CountByCode(workbook.ExternalQueryConnections.SelectMany(GetExternalQueryConnectionOptionKeys));
            ExternalQueryConnectionsByParameterFlagCount = CountByCode(workbook.ExternalQueryConnections.Select(connection => $"Parameters:{connection.ParameterFlagCount}"));
            ExternalQueryConnectionsByParameterFlagByteCount = CountByCode(workbook.ExternalQueryConnections.Select(connection => $"Bytes:{connection.ParameterFlagByteCount}"));
            ExternalQueryConnectionsByParameterFlagState = CountByCode(workbook.ExternalQueryConnections.Select(connection => connection.HasCompleteParameterFlags ? "Complete" : "Mismatched"));
            ExternalQueryConnectionsByFutureByteCount = CountByCode(workbook.ExternalQueryConnections.Select(connection => $"Bytes:{connection.FutureByteCount}"));
            ExternalQueryConnectionsByRefreshInterval = CountByCode(workbook.ExternalQueryConnections.Select(connection => connection.RefreshIntervalMinutes == 0 ? "Off" : $"Minutes:{connection.RefreshIntervalMinutes}"));
            ExternalQueryConnectionsByOleDbConnectionCount = CountByCode(workbook.ExternalQueryConnections.Select(connection => $"OleDbConnections:{connection.OleDbConnectionCount}"));
            ExternalQueryConnectionsByHtmlFormat = CountByCode(workbook.ExternalQueryConnections.Select(connection => $"HtmlFormat:0x{connection.HtmlFormat:X4}"));
            ExternalQueryConnectionsByVersionTriplet = CountByCode(workbook.ExternalQueryConnections.Select(connection => $"Edit:{connection.EditVersion};Refreshed:{connection.RefreshedVersion};RefreshableMin:{connection.RefreshableMinimumVersion}"));
            ExternalQueryConnectionsBySourceSpecificFlags = CountByCode(workbook.ExternalQueryConnections.Select(connection => $"Flags:0x{connection.SourceSpecificFlags:X4}"));
            DataConsolidationReferencesBySourceKind = CountByCode(workbook.DataConsolidationReferences.Select(reference => reference.SourceKind.ToString()));
            DataConsolidationReferencesBySourcePrefix = CountByCode(workbook.DataConsolidationReferences.Select(GetDataConsolidationReferenceSourcePrefixKey));
            DataConsolidationReferencesBySource = CountByCode(workbook.DataConsolidationReferences.Select(reference => reference.Source));
            DataConsolidationReferencesByRange = CountByCode(workbook.DataConsolidationReferences.Select(reference => reference.CellRange));
            DataConsolidationReferencesByShape = CountByCode(workbook.DataConsolidationReferences.Select(reference => $"Rows:{reference.RowSpan};Columns:{reference.ColumnSpan}"));
            DataConsolidationReferencesBySourceAndRange = CountByCode(workbook.DataConsolidationReferences.Select(reference => $"{reference.Source}|{reference.CellRange}"));
            DataConsolidationReferencesByUnusedByteCount = CountByCode(workbook.DataConsolidationReferences.Select(reference => $"UnusedBytes:{reference.UnusedByteCount}"));
            DataConsolidationNamesBySourceKind = CountByCode(workbook.DataConsolidationNames.Select(name => name.SourceKind.ToString()));
            DataConsolidationNamesByName = CountByCode(workbook.DataConsolidationNames.Select(name => name.Name));
            DataConsolidationNamesBySource = CountByCode(workbook.DataConsolidationNames.Select(GetDataConsolidationNameSourceKey));
            DataConsolidationNamesByNameAndSource = CountByCode(workbook.DataConsolidationNames.Select(name => $"{name.Name}|{GetDataConsolidationNameSourceKey(name)}"));
            DataConsolidationNamesByUnusedByteCount = CountByCode(workbook.DataConsolidationNames.Select(name => $"UnusedBytes:{name.UnusedByteCount}"));
            ThemeRecordsByVersion = CountByCode(workbook.ThemeRecords.Select(record => record.ThemeVersionName));
            ThemeRecordsByRawVersion = CountByCode(workbook.ThemeRecords.Select(record => $"Version:{record.ThemeVersion}"));
            ThemeRecordsByContentState = CountByCode(workbook.ThemeRecords.Select(record => record.HasThemeBytes ? "EmbeddedThemeBytes" : "NoEmbeddedThemeBytes"));
            ThemeRecordsByContentLength = CountByCode(workbook.ThemeRecords.Select(record => $"Bytes:{record.ThemeByteCount}"));
            PivotTableRecordsByKind = CountPivotTableRecordsByKind(workbook.PivotTableRecords);
            PivotTableRecordsByName = CountByCode(workbook.PivotTableRecords.Select(record => record.RecordName));
            PivotTableRecordsByLocation = CountByCode(workbook.PivotTableRecords.Select(GetPivotTableRecordLocationKey));
            PivotTableRecordsByKindAndLocation = CountByCode(workbook.PivotTableRecords
                .Select(record => $"{record.Kind}|{GetPivotTableRecordLocationKey(record)}"));
            PivotTableRecordsByNameAndLocation = CountByCode(workbook.PivotTableRecords
                .Select(record => $"{record.RecordName}|{GetPivotTableRecordLocationKey(record)}"));
            PivotTableWorkbookStates = CountByCode(GetPivotTableWorkbookStateKeys(workbook.PivotTableRecords));
            PivotTableViewRanges = CountByCode(workbook.PivotTableRecords
                .Where(record => !string.IsNullOrWhiteSpace(record.ViewRange))
                .Select(record => record.ViewRange!));
            PivotTableViewNames = CountByCode(workbook.PivotTableRecords
                .Where(record => !string.IsNullOrWhiteSpace(record.ViewTableName))
                .Select(record => record.ViewTableName!));
            PivotTableViewDataNames = CountByCode(workbook.PivotTableRecords
                .Where(record => !string.IsNullOrWhiteSpace(record.ViewDataName))
                .Select(record => record.ViewDataName!));
            PivotTableViewFieldCounts = CountByCode(workbook.PivotTableRecords
                .Where(record => record.ViewFieldCount.HasValue)
                .Select(record => $"Fields:{record.ViewFieldCount!.Value};Rows:{record.ViewRowFieldCount!.Value};Columns:{record.ViewColumnFieldCount!.Value};Pages:{record.ViewPageFieldCount!.Value};Data:{record.ViewDataFieldCount!.Value}"));
            PivotTableViewLineCounts = CountByCode(workbook.PivotTableRecords
                .Where(record => record.ViewRowLineCount.HasValue && record.ViewColumnLineCount.HasValue)
                .Select(record => $"Rows:{record.ViewRowLineCount!.Value};Columns:{record.ViewColumnLineCount!.Value}"));
            PivotTableViewDataAxes = CountByCode(workbook.PivotTableRecords
                .Where(record => !string.IsNullOrWhiteSpace(record.ViewDataAxisName))
                .Select(record => record.ViewDataAxisName!));
            PivotTableViewDataPositions = CountByCode(workbook.PivotTableRecords
                .Where(record => record.ViewDataPosition.HasValue)
                .Select(record => $"Position:{record.ViewDataPosition!.Value}"));
            PivotTableViewCacheIndexes = CountByCode(workbook.PivotTableRecords
                .Where(record => record.ViewCacheIndex.HasValue)
                .Select(record => $"CacheIndex:{record.ViewCacheIndex!.Value}"));
            PivotTableViewGrandTotalStates = CountByCode(workbook.PivotTableRecords
                .Where(record => record.ViewRowGrandTotals.HasValue && record.ViewColumnGrandTotals.HasValue)
                .Select(record => $"Rows:{record.ViewRowGrandTotals!.Value};Columns:{record.ViewColumnGrandTotals!.Value}"));
            PivotTableViewAutoFormatStates = CountByCode(workbook.PivotTableRecords
                .Where(record => record.ViewAutoFormat.HasValue && record.ViewAutoFormatId.HasValue)
                .Select(record => $"AutoFormat:{record.ViewAutoFormat!.Value};Id:{record.ViewAutoFormatId!.Value}"));
            PivotTableFieldAxes = CountByCode(workbook.PivotTableRecords
                .Where(record => !string.IsNullOrWhiteSpace(record.FieldAxisName))
                .Select(record => record.FieldAxisName!));
            PivotTableFieldItemCounts = CountByCode(workbook.PivotTableRecords
                .Where(record => record.FieldItemCount.HasValue)
                .Select(record => $"Items:{record.FieldItemCount!.Value}"));
            PivotTableFieldSubtotalCounts = CountByCode(workbook.PivotTableRecords
                .Where(record => record.FieldSubtotalCount.HasValue && record.FieldSubtotalFlags.HasValue)
                .Select(record => $"Subtotals:{record.FieldSubtotalCount!.Value};Flags:0x{record.FieldSubtotalFlags!.Value:X4}"));
            PivotTableFieldSubtotalFunctions = CountByCode(workbook.PivotTableRecords
                .SelectMany(record => record.FieldSubtotalFunctionNames));
            PivotTableFieldNames = CountByCode(workbook.PivotTableRecords
                .Where(record => !string.IsNullOrWhiteSpace(record.FieldName))
                .Select(record => record.FieldName!));
            PivotTableFieldIndexListLengths = CountByCode(workbook.PivotTableRecords
                .Where(record => record.FieldIndexReferences.Count > 0)
                .Select(record => $"Indexes:{record.FieldIndexReferences.Count}"));
            PivotTableFieldIndexReferences = CountByCode(workbook.PivotTableRecords
                .SelectMany(record => record.FieldIndexReferences)
                .Select(index => $"FieldIndex:{index}"));
            PivotTableFieldIndexSequences = CountByCode(workbook.PivotTableRecords
                .Where(record => record.FieldIndexReferences.Count > 0)
                .Select(record => "FieldIndexes:" + string.Join(",", record.FieldIndexReferences)));
            PivotTableLineItemCounts = CountByCode(workbook.PivotTableRecords
                .Where(record => record.LineItems.Count > 0)
                .Select(record => $"LineItems:{record.LineItems.Count}"));
            PivotTableLineItemTypes = CountByCode(workbook.PivotTableRecords
                .SelectMany(record => record.LineItems)
                .Select(item => $"LineItemType:{item.ItemType}"));
            PivotTableLineItemTypeKinds = CountByCode(workbook.PivotTableRecords
                .SelectMany(record => record.LineItems)
                .Select(item => item.ItemTypeName));
            PivotTableLineItemEntryCounts = CountByCode(workbook.PivotTableRecords
                .SelectMany(record => record.LineItems)
                .Select(item => $"Entries:{item.EntryCount}"));
            PivotTableLineItemEntrySlotCounts = CountByCode(workbook.PivotTableRecords
                .SelectMany(record => record.LineItems)
                .Select(item => $"Slots:{item.EntrySlotCount};Entries:{item.EntryCount}"));
            PivotTableLineItemEntryIndexes = CountByCode(workbook.PivotTableRecords
                .SelectMany(record => record.LineItems)
                .SelectMany(item => item.EntryIndexNames));
            PivotTableLineItemDataIndexes = CountByCode(workbook.PivotTableRecords
                .SelectMany(record => record.LineItems)
                .Select(item => $"DataIndex:{item.DataIndex}"));
            PivotTableLineItemFlagStates = CountByCode(workbook.PivotTableRecords
                .SelectMany(record => record.LineItems)
                .Select(item => $"Subtotal:{item.Subtotal};Block:{item.BlockTotal};Grand:{item.GrandTotal};MultiDataName:{item.MultiDataName};MultiDataOnAxis:{item.MultiDataOnAxis}"));
            PivotTableLineItemSequences = CountByCode(workbook.PivotTableRecords
                .Where(record => record.LineItems.Count > 0)
                .Select(record => "LineItems:" + string.Join(",", record.LineItems.Select(item => $"Type:{item.ItemTypeName};Entries:{string.Join(",", item.EntryIndexNames)}"))));
            PivotTablePageItemCounts = CountByCode(workbook.PivotTableRecords
                .Where(record => record.PageItems.Count > 0)
                .Select(record => $"PageItems:{record.PageItems.Count}"));
            PivotTablePageItemFieldIndexes = CountByCode(workbook.PivotTableRecords
                .SelectMany(record => record.PageItems)
                .Select(item => $"FieldIndex:{item.FieldIndex}"));
            PivotTablePageItemIndexes = CountByCode(workbook.PivotTableRecords
                .SelectMany(record => record.PageItems)
                .Select(item => item.ItemIndexName));
            PivotTablePageItemObjectIds = CountByCode(workbook.PivotTableRecords
                .SelectMany(record => record.PageItems)
                .Select(item => $"ObjectId:{item.ObjectId}"));
            PivotTablePageItemSequences = CountByCode(workbook.PivotTableRecords
                .Where(record => record.PageItems.Count > 0)
                .Select(record => "PageItems:" + string.Join(",", record.PageItems.Select(item => $"FieldIndex:{item.FieldIndex};{item.ItemIndexName};ObjectId:{item.ObjectId}"))));
            PivotTableItemTypes = CountByCode(workbook.PivotTableRecords
                .Where(record => record.ItemType.HasValue)
                .Select(record => $"ItemType:{record.ItemType!.Value}"));
            PivotTableItemTypeKinds = CountByCode(workbook.PivotTableRecords
                .Where(record => !string.IsNullOrWhiteSpace(record.ItemTypeName))
                .Select(record => record.ItemTypeName!));
            PivotTableItemCacheIndexes = CountByCode(workbook.PivotTableRecords
                .Where(record => !string.IsNullOrWhiteSpace(record.ItemCacheIndexName))
                .Select(record => record.ItemCacheIndexName!));
            PivotTableItemFlagStates = CountByCode(workbook.PivotTableRecords
                .Where(record => record.ItemHidden.HasValue && record.ItemHideDetail.HasValue && record.ItemFormula.HasValue && record.ItemMissing.HasValue)
                .Select(record => $"Hidden:{record.ItemHidden!.Value};HideDetail:{record.ItemHideDetail!.Value};Formula:{record.ItemFormula!.Value};Missing:{record.ItemMissing!.Value}"));
            PivotTableItemNames = CountByCode(workbook.PivotTableRecords
                .Where(record => !string.IsNullOrWhiteSpace(record.ItemName))
                .Select(record => record.ItemName!));
            PivotTableFormulaPayloadLengths = CountByCode(workbook.PivotTableRecords
                .Where(record => record.Kind == LegacyXlsPivotTableRecordKind.Formula)
                .Select(record => $"{record.RecordName}|Bytes:{record.PayloadLength}"));
            PivotTableFormulaPayloadKinds = CountByCode(workbook.PivotTableRecords
                .Where(record => !string.IsNullOrWhiteSpace(record.CalculatedItemFormulaPayloadKind))
                .Select(record => record.CalculatedItemFormulaPayloadKind!));
            PivotTableFormulaTokenByteCounts = CountByCode(workbook.PivotTableRecords
                .Where(record => record.CalculatedItemFormulaTokenByteCount.HasValue)
                .Select(record => $"TokenBytes:{record.CalculatedItemFormulaTokenByteCount!.Value}"));
            PivotTableCalculatedFieldFormulaTokenByteCounts = CountByCode(workbook.PivotTableRecords
                .Where(record => record.CalculatedFieldFormulaTokenByteCount.HasValue)
                .Select(record => $"TokenBytes:{record.CalculatedFieldFormulaTokenByteCount!.Value}"));
            PivotTableFormulaTrailingByteCounts = CountByCode(workbook.PivotTableRecords
                .Where(record => record.CalculatedItemFormulaTrailingByteCount.HasValue)
                .Select(record => $"TrailingBytes:{record.CalculatedItemFormulaTrailingByteCount!.Value}"));
            PivotTableRuleAxes = CountByCode(workbook.PivotTableRecords
                .Where(record => !string.IsNullOrWhiteSpace(record.RuleAxisName))
                .Select(record => record.RuleAxisName!));
            PivotTableRuleTypes = CountByCode(workbook.PivotTableRecords
                .Where(record => !string.IsNullOrWhiteSpace(record.RuleTypeName))
                .Select(record => record.RuleTypeName!));
            PivotTableRuleFieldReferences = CountByCode(workbook.PivotTableRecords
                .Where(record => !string.IsNullOrWhiteSpace(record.RuleFieldReferenceName))
                .Select(record => record.RuleFieldReferenceName!));
            PivotTableRuleFilterCounts = CountByCode(workbook.PivotTableRecords
                .Where(record => record.RuleFilterCount.HasValue)
                .Select(record => $"Filters:{record.RuleFilterCount!.Value}"));
            PivotTableRuleOptionStates = CountByCode(workbook.PivotTableRecords
                .Where(record => record.RulePartialArea.HasValue)
                .Select(record => $"Partial:{record.RulePartialArea!.Value};DataOnly:{record.RuleDataOnly!.Value};LabelOnly:{record.RuleLabelOnly!.Value};CacheBased:{record.RuleCacheBased!.Value}"));
            PivotTableRulePartialAreas = CountByCode(workbook.PivotTableRecords
                .Where(record => !string.IsNullOrWhiteSpace(record.RulePartialAreaRange))
                .Select(record => record.RulePartialAreaRange!));
            PivotTableRuleFilterEntryCounts = CountByCode(workbook.PivotTableRecords
                .Where(record => record.RuleFilters.Count > 0)
                .Select(record => $"Filters:{record.RuleFilters.Count}"));
            PivotTableRuleFilterAxes = CountByCode(workbook.PivotTableRecords
                .SelectMany(record => record.RuleFilters)
                .Select(filter => filter.AxisName));
            PivotTableRuleFilterFieldPositions = CountByCode(workbook.PivotTableRecords
                .SelectMany(record => record.RuleFilters)
                .Select(filter => $"Position:{filter.FieldPosition}"));
            PivotTableRuleFilterFieldReferences = CountByCode(workbook.PivotTableRecords
                .SelectMany(record => record.RuleFilters)
                .Select(filter => filter.FieldReferenceName));
            PivotTableRuleFilterSelectedStates = CountByCode(workbook.PivotTableRecords
                .SelectMany(record => record.RuleFilters)
                .Select(filter => $"Selected:{filter.Selected}"));
            PivotTableRuleFilterSubtotalFlags = CountByCode(workbook.PivotTableRecords
                .SelectMany(record => record.RuleFilters)
                .Select(filter => $"Flags:0x{filter.SubtotalFlags:X4}"));
            PivotTableRuleFilterSubtotalFunctions = CountByCode(workbook.PivotTableRecords
                .SelectMany(record => record.RuleFilters)
                .SelectMany(filter => filter.SubtotalFunctionNames));
            PivotTableRuleFilterItemIndexCounts = CountByCode(workbook.PivotTableRecords
                .SelectMany(record => record.RuleFilters)
                .Select(filter => $"Indexes:{filter.ItemIndexCount}"));
            PivotTableRuleFilterStates = CountByCode(workbook.PivotTableRecords
                .SelectMany(record => record.RuleFilters)
                .Select(GetPivotTableRuleFilterStateKey));
            PivotTableCacheItemKinds = CountByCode(workbook.PivotTableRecords
                .Where(record => !string.IsNullOrWhiteSpace(record.CacheItemKindName))
                .Select(record => record.CacheItemKindName!));
            PivotTableCacheItemValueStates = CountByCode(workbook.PivotTableRecords
                .Where(record => record.CacheItemKind.HasValue)
                .Select(record => record.IsEmptyCacheItem ? "Empty" : "HasValue"));
            PivotTableCacheItemStringLengths = CountByCode(workbook.PivotTableRecords
                .Where(record => record.CacheItemKind == LegacyXlsPivotCacheItemKind.String)
                .Select(record => record.CacheItemStringValue == null ? "NoStringSegment" : $"Characters:{record.CacheItemStringValue.Length}"));
            PivotTableCacheItemErrorCodes = CountByCode(workbook.PivotTableRecords
                .Where(record => record.CacheItemErrorCode.HasValue)
                .Select(record => $"ErrorCode:0x{record.CacheItemErrorCode!.Value:X2}"));
            PivotTableCacheItemBooleanValues = CountByCode(workbook.PivotTableRecords
                .Where(record => record.CacheItemBooleanValue.HasValue)
                .Select(record => record.CacheItemBooleanValue!.Value ? "True" : "False"));
            PivotTableCacheStreamNames = CountByCode(workbook.PivotTableRecords
                .Where(record => !string.IsNullOrWhiteSpace(record.CacheStreamName))
                .Select(record => $"{record.RecordName}|{record.CacheStreamName}"));
            PivotTableCacheSourceTypes = CountByCode(workbook.PivotTableRecords
                .Where(record => !string.IsNullOrWhiteSpace(record.CacheSourceTypeName))
                .Select(record => $"{record.RecordName}|{record.CacheSourceTypeName}"));
            PivotTableCacheRecordCounts = CountByCode(workbook.PivotTableRecords
                .Where(record => record.CacheRecordCount.HasValue)
                .Select(record => $"Records:{record.CacheRecordCount!.Value}"));
            PivotTableCacheFieldCounts = CountByCode(workbook.PivotTableRecords
                .Where(record => record.CacheSourceFieldCount.HasValue && record.CacheTotalFieldCount.HasValue)
                .Select(record => $"SourceFields:{record.CacheSourceFieldCount!.Value};TotalFields:{record.CacheTotalFieldCount!.Value}"));
            PivotTableCacheUsedRecordCounts = CountByCode(workbook.PivotTableRecords
                .Where(record => record.CacheUsedRecordCount.HasValue)
                .Select(record => $"UsedRecords:{record.CacheUsedRecordCount!.Value}"));
            PivotTableCachePropertyFlags = CountByCode(workbook.PivotTableRecords.SelectMany(GetPivotTableCachePropertyFlagKeys));
            PivotTableCacheRefreshUserStates = CountByCode(workbook.PivotTableRecords
                .Where(record => record.CacheRecordCount.HasValue)
                .Select(record => string.IsNullOrWhiteSpace(record.CacheRefreshedBy) ? "NoRefreshUser" : "HasRefreshUser"));
            PivotTableQueryTagTargets = CountByCode(workbook.PivotTableRecords
                .Where(record => !string.IsNullOrWhiteSpace(record.QueryTableTagTargetName))
                .Select(record => record.QueryTableTagTargetName!));
            PivotTableQueryTagNames = CountByCode(workbook.PivotTableRecords
                .Where(record => !string.IsNullOrWhiteSpace(record.QueryTableTagName))
                .Select(record => record.QueryTableTagName!));
            PivotTableQueryTagRefreshStates = CountByCode(workbook.PivotTableRecords
                .Where(record => record.QueryTableTagRefreshEnabled.HasValue && record.QueryTableTagCacheInvalid.HasValue && record.QueryTableTagTensorEx.HasValue)
                .Select(record => $"RefreshEnabled:{record.QueryTableTagRefreshEnabled!.Value};CacheInvalid:{record.QueryTableTagCacheInvalid!.Value};TensorEx:{record.QueryTableTagTensorEx!.Value}"));
            PivotTableQueryTagVersions = CountByCode(workbook.PivotTableRecords
                .Where(record => record.QueryTableTagLastUpdatedVersion.HasValue && record.QueryTableTagUpdatableMinimumVersion.HasValue)
                .Select(record => $"LastUpdated:{record.QueryTableTagLastUpdatedVersion!.Value};UpdatableMin:{record.QueryTableTagUpdatableMinimumVersion!.Value}"));
            PivotTableQueryTagFutureOptions = CountByCode(workbook.PivotTableRecords
                .Where(record => record.QueryTableTagFutureOptions.HasValue)
                .Select(record => $"Options:0x{record.QueryTableTagFutureOptions!.Value:X8}"));
            PivotTableQueryTagUnusedValues = CountByCode(workbook.PivotTableRecords
                .Where(record => record.QueryTableTagUnused.HasValue)
                .Select(record => $"Unused:0x{record.QueryTableTagUnused!.Value:X4}"));
            PivotTableDataItemAggregations = CountByCode(workbook.PivotTableRecords
                .Where(record => record.AggregationFunction.HasValue)
                .Select(record => $"AggregationFunction:{record.AggregationFunction!.Value}"));
            PivotTableDataItemAggregationKinds = CountByCode(workbook.PivotTableRecords
                .Where(record => !string.IsNullOrWhiteSpace(record.AggregationFunctionName))
                .Select(record => record.AggregationFunctionName!));
            PivotTableDataItemFieldIndexes = CountByCode(workbook.PivotTableRecords
                .Where(record => record.DataItemFieldIndex.HasValue)
                .Select(record => $"FieldIndex:{record.DataItemFieldIndex!.Value}"));
            PivotTableDataItemDisplayCalculationIds = CountByCode(workbook.PivotTableRecords
                .Where(record => record.DisplayCalculation.HasValue)
                .Select(record => $"DisplayCalculation:{record.DisplayCalculation!.Value}"));
            PivotTableDataItemDisplayCalculations = CountByCode(workbook.PivotTableRecords
                .Where(record => !string.IsNullOrWhiteSpace(record.DisplayCalculationName))
                .Select(record => record.DisplayCalculationName!));
            PivotTableDataItemDisplayCalculationReferenceStates = CountByCode(workbook.PivotTableRecords
                .Where(record => !string.IsNullOrWhiteSpace(record.DisplayCalculationName)
                    && !string.IsNullOrWhiteSpace(record.DisplayCalculationFieldReferenceName)
                    && !string.IsNullOrWhiteSpace(record.DisplayCalculationItemReferenceName))
                .Select(GetPivotTableDataItemDisplayCalculationReferenceStateKey));
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
            PivotTableGroupingCompletionStates = CountByCode(workbook.PivotTableRecords
                .Where(record => record.GroupingKind.HasValue)
                .Select(GetPivotTableGroupingCompletionStateKey));
            PivotTableGroupingStates = CountByCode(workbook.PivotTableRecords
                .Where(record => record.GroupingKind.HasValue)
                .Select(GetPivotTableGroupingStateKey));
            PivotTableGroupingNumericRanges = CountByCode(workbook.PivotTableRecords
                .Where(record => record.GroupingNumericStart.HasValue && record.GroupingNumericEnd.HasValue && record.GroupingNumericInterval.HasValue)
                .Select(GetPivotTableGroupingNumericRangeKey));
            PivotTableGroupingDateRanges = CountByCode(workbook.PivotTableRecords
                .Where(record => record.GroupingDateStart != null && record.GroupingDateEnd != null && record.GroupingDateInterval.HasValue)
                .Select(GetPivotTableGroupingDateRangeKey));
            PivotTableFormulaScopes = CountByCode(workbook.PivotTableRecords
                .Where(record => !string.IsNullOrWhiteSpace(record.CalculatedItemFormulaScopeName))
                .Select(record => record.CalculatedItemFormulaScopeName!));
            PivotTableFormulaCacheFieldIndexes = CountByCode(workbook.PivotTableRecords
                .Where(record => record.CalculatedItemFormulaCacheFieldIndex.HasValue)
                .Select(record => $"CacheField:{record.CalculatedItemFormulaCacheFieldIndex!.Value}"));
            PivotTableFormulaReservedValues = CountByCode(workbook.PivotTableRecords
                .Where(record => record.CalculatedItemFormulaReserved.HasValue)
                .Select(record => $"Reserved:0x{record.CalculatedItemFormulaReserved!.Value:X4}"));
            PivotTableExtendedFieldStates = CountByCode(workbook.PivotTableRecords.SelectMany(GetPivotTableExtendedFieldStateKeys));
            PivotTableExtendedFieldPermissionStates = CountByCode(workbook.PivotTableRecords
                .Where(record => record.ShowAllItems.HasValue)
                .Select(GetPivotTableExtendedFieldPermissionStateKey));
            PivotTableAdditionalClasses = CountByCode(workbook.PivotTableRecords
                .Where(record => !string.IsNullOrWhiteSpace(record.AdditionalClassName))
                .Select(record => record.AdditionalClassName!));
            PivotTableAdditionalTypes = CountByCode(workbook.PivotTableRecords
                .Where(record => !string.IsNullOrWhiteSpace(record.AdditionalTypeName))
                .Select(record => record.AdditionalTypeName!));
            PivotTableAdditionalClassTypes = CountByCode(workbook.PivotTableRecords
                .Where(record => !string.IsNullOrWhiteSpace(record.AdditionalClassName) && !string.IsNullOrWhiteSpace(record.AdditionalTypeName))
                .Select(GetPivotTableAdditionalClassTypeKey));
            PivotTableAdditionalFutureRecordTypes = CountByCode(workbook.PivotTableRecords
                .Where(record => record.AdditionalFutureRecordType.HasValue)
                .Select(record => $"FrtType:0x{record.AdditionalFutureRecordType!.Value:X4}"));
            PivotTableAdditionalFutureFlags = CountByCode(workbook.PivotTableRecords
                .Where(record => record.AdditionalFutureFlags.HasValue)
                .Select(record => $"Flags:0x{record.AdditionalFutureFlags!.Value:X4}"));
            PivotTableAdditionalSequenceIndexes = CountByCode(workbook.PivotTableRecords
                .Where(record => record.AdditionalSequenceIndex.HasValue)
                .Select(record => $"Index:{record.AdditionalSequenceIndex!.Value}"));
            PivotTableAdditionalPayloadLengthsByClassType = CountByCode(workbook.PivotTableRecords
                .Where(record => !string.IsNullOrWhiteSpace(record.AdditionalClassName) && !string.IsNullOrWhiteSpace(record.AdditionalTypeName))
                .Select(record => $"{GetPivotTableAdditionalClassTypeKey(record)}|Bytes:{record.PayloadLength}"));
            PivotTableAdditionalCacheIds = CountByCode(workbook.PivotTableRecords
                .Where(record => record.AdditionalCacheId.HasValue)
                .Select(record => $"CacheId:{record.AdditionalCacheId!.Value}"));
            PivotTableAdditionalClassDepthsBefore = CountByCode(workbook.PivotTableRecords
                .Where(record => record.AdditionalClassDepthBefore.HasValue)
                .Select(record => $"Depth:{record.AdditionalClassDepthBefore!.Value}"));
            PivotTableAdditionalClassDepthsAfter = CountByCode(workbook.PivotTableRecords
                .Where(record => record.AdditionalClassDepthAfter.HasValue)
                .Select(record => $"Depth:{record.AdditionalClassDepthAfter!.Value}"));
            PivotTableAdditionalClassTransitions = CountByCode(workbook.PivotTableRecords
                .Where(record => !string.IsNullOrWhiteSpace(record.AdditionalClassTransition))
                .Select(record => record.AdditionalClassTransition!));
            PivotTableAdditionalClassTransitionsByClassType = CountByCode(workbook.PivotTableRecords
                .Where(record => !string.IsNullOrWhiteSpace(record.AdditionalClassName)
                    && !string.IsNullOrWhiteSpace(record.AdditionalTypeName)
                    && !string.IsNullOrWhiteSpace(record.AdditionalClassTransition))
                .Select(record => $"{GetPivotTableAdditionalClassTypeKey(record)}|{record.AdditionalClassTransition}"));
            ChartRecordsByKind = CountChartRecordsByKind(workbook.ChartRecords);
            ChartRecordsByName = CountByCode(workbook.ChartRecords.Select(record => record.RecordName));
            ChartRecordsByNameAndPayloadLength = CountByCode(workbook.ChartRecords
                .Select(record => $"{record.RecordName}|Bytes:{record.PayloadLength}"));
            ChartWorkbookStates = CountByCode(GetChartWorkbookStateKeys(workbook.ChartRecords, workbook.ChartSheets));
            ChartRecordsByContainerDepthBefore = CountByCode(workbook.ChartRecords
                .Where(record => record.ContainerDepthBefore.HasValue)
                .Select(record => $"Depth:{record.ContainerDepthBefore!.Value}"));
            ChartRecordsByContainerDepthAfter = CountByCode(workbook.ChartRecords
                .Where(record => record.ContainerDepthAfter.HasValue)
                .Select(record => $"Depth:{record.ContainerDepthAfter!.Value}"));
            ChartRecordsByContainerTransition = CountByCode(workbook.ChartRecords
                .Where(record => !string.IsNullOrWhiteSpace(record.ContainerTransition))
                .Select(record => record.ContainerTransition!));
            ChartRecordsByNameAndContainerDepth = CountByCode(workbook.ChartRecords
                .Where(record => record.ContainerDepthBefore.HasValue)
                .Select(record => $"{record.RecordName}|Depth:{record.ContainerDepthBefore!.Value}"));
            ChartRecordsByNameAndContainerTransition = CountByCode(workbook.ChartRecords
                .Where(record => !string.IsNullOrWhiteSpace(record.ContainerTransition))
                .Select(record => $"{record.RecordName}|{record.ContainerTransition}"));
            ChartRecordsByChartType = CountByCode(workbook.ChartRecords
                .Where(record => !string.IsNullOrWhiteSpace(record.ChartTypeName))
                .Select(record => record.ChartTypeName!));
            ChartRecordsByRectangle = CountByCode(workbook.ChartRecords
                .Where(record => record.ChartX.HasValue && record.ChartY.HasValue && record.ChartWidth.HasValue && record.ChartHeight.HasValue)
                .Select(record => $"X:{record.ChartX!.Value};Y:{record.ChartY!.Value};Width:{record.ChartWidth!.Value};Height:{record.ChartHeight!.Value}"));
            ChartRecordsByAxisType = CountByCode(workbook.ChartRecords
                .Where(record => !string.IsNullOrWhiteSpace(record.AxisTypeName))
                .Select(record => record.AxisTypeName!));
            ChartGroupVariedColorStates = CountByCode(workbook.ChartRecords
                .Where(record => record.ChartGroupOptions != null)
                .Select(record => record.ChartGroupOptions!.VariedDataPointColors ? "VariedDataPointColors" : "UniformDataPointColors"));
            ChartGroupDrawingOrders = CountByCode(workbook.ChartRecords
                .Where(record => record.ChartGroupOptions != null)
                .Select(record => $"DrawingOrder:{record.ChartGroupOptions!.DrawingOrder}"));
            ChartRecordsByAxesUsedCount = CountByCode(workbook.ChartRecords
                .Where(record => record.AxesUsedCount.HasValue)
                .Select(record => $"AxesUsed:{record.AxesUsedCount!.Value}"));
            ChartCategorySeriesRangeIntervals = CountByCode(workbook.ChartRecords
                .Where(record => record.CategorySeriesRange != null)
                .Select(GetChartCategorySeriesRangeIntervalKey));
            ChartCategorySeriesRangeStates = CountByCode(workbook.ChartRecords
                .Where(record => record.CategorySeriesRange != null)
                .Select(GetChartCategorySeriesRangeStateKey));
            ChartAxisExtensionDateRanges = CountByCode(workbook.ChartRecords
                .Where(record => record.AxisExtension != null)
                .Select(GetChartAxisExtensionDateRangeKey));
            ChartAxisExtensionDateUnits = CountByCode(workbook.ChartRecords
                .Where(record => record.AxisExtension != null)
                .Select(GetChartAxisExtensionDateUnitKey));
            ChartAxisExtensionStates = CountByCode(workbook.ChartRecords
                .Where(record => record.AxisExtension != null)
                .Select(GetChartAxisExtensionStateKey));
            ChartAxisExtensionReservedStates = CountByCode(workbook.ChartRecords
                .Where(record => record.AxisExtension != null)
                .Select(record => record.AxisExtension!.HasZeroReservedByte ? "ReservedZero" : "ReservedNonZero"));
            ChartCategoryLabelAlignments = CountByCode(workbook.ChartRecords
                .Where(record => record.CategoryLabelOptions != null)
                .Select(record => record.CategoryLabelOptions!.AlignmentName));
            ChartCategoryLabelOffsets = CountByCode(workbook.ChartRecords
                .Where(record => record.CategoryLabelOptions != null)
                .Select(record => $"Offset:{record.CategoryLabelOptions!.OffsetPercentage}%"));
            ChartCategoryLabelCountStates = CountByCode(workbook.ChartRecords
                .Where(record => record.CategoryLabelOptions != null)
                .Select(record => record.CategoryLabelOptions!.UseAutomaticLabelCount ? "AutomaticLabelCount" : "CatSerRangeLabelCount"));
            ChartAxisLineFormatTargets = CountByCode(workbook.ChartRecords
                .Where(record => record.AxisLineFormat != null)
                .Select(record => record.AxisLineFormat!.TargetName));
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
            ChartSeriesChartGroupIndexes = CountByCode(workbook.ChartRecords
                .Where(record => record.SeriesChartGroupReference != null)
                .Select(record => $"ChartGroupIndex:{record.SeriesChartGroupReference!.ChartGroupIndex}"));
            ChartSeriesListDeclaredCounts = CountByCode(workbook.ChartRecords
                .Where(record => record.SeriesList != null)
                .Select(record => $"Declared:{record.SeriesList!.DeclaredSeriesCount}"));
            ChartSeriesListDecodedCounts = CountByCode(workbook.ChartRecords
                .Where(record => record.SeriesList != null)
                .Select(record => $"Decoded:{record.SeriesList!.DecodedSeriesCount}"));
            ChartSeriesListCompletenessStates = CountByCode(workbook.ChartRecords
                .Where(record => record.SeriesList != null)
                .Select(record => record.SeriesList!.HasCompleteSeriesIndexList ? "Complete" : "Truncated"));
            ChartSeriesListIndexValidityStates = CountByCode(workbook.ChartRecords
                .Where(record => record.SeriesList != null)
                .Select(record => record.SeriesList!.HasOnlyValidSeriesIndexes ? "AllValid" : "ContainsInvalid"));
            ChartPivotViewReferences = CountByCode(workbook.ChartRecords
                .Where(record => record.PivotViewReference != null)
                .Select(record => record.PivotViewReference!.Reference));
            ChartSeriesDataCacheIndexes = CountByCode(workbook.ChartRecords
                .Where(record => record.SeriesDataCacheIndex.HasValue)
                .Select(record => $"Index:{record.SeriesDataCacheIndex!.Value}"));
            ChartSeriesDataCacheTypes = CountByCode(workbook.ChartRecords
                .Where(record => !string.IsNullOrWhiteSpace(record.SeriesDataCacheIndexName))
                .Select(record => record.SeriesDataCacheIndexName!));
            ChartDataSourceIds = CountByCode(workbook.ChartRecords
                .Where(record => record.DataSource != null)
                .Select(record => record.DataSource!.SourceIdName));
            ChartDataSourceReferenceTypes = CountByCode(workbook.ChartRecords
                .Where(record => record.DataSource != null)
                .Select(record => record.DataSource!.ReferenceTypeName));
            ChartDataSourceNumberFormatIds = CountByCode(workbook.ChartRecords
                .Where(record => record.DataSource != null)
                .Select(record => $"NumberFormatId:{record.DataSource!.NumberFormatId}"));
            ChartDataSourceFormulaByteCounts = CountByCode(workbook.ChartRecords
                .Where(record => record.DataSource != null)
                .Select(record => $"FormulaBytes:{record.DataSource!.FormulaByteCount}"));
            ChartDataSourceFormulaProjectionStates = CountByCode(workbook.ChartRecords
                .Where(record => record.DataSource != null)
                .Select(GetChartDataSourceFormulaProjectionStateKey));
            ChartDataSourceFormulaTexts = CountByCode(workbook.ChartRecords
                .Where(record => !string.IsNullOrWhiteSpace(record.DataSource?.FormulaText))
                .Select(record => record.DataSource!.FormulaText!));
            ChartDataSourceFormulaProjectionFailures = CountByCode(workbook.ChartRecords
                .Where(record => record.DataSource?.HasFormulaProjectionFailure == true)
                .Select(record => record.DataSource!.FormulaProjectionFailureCode!));
            ChartDataSourceFormulaProjectionFailuresByToken = CountByCode(workbook.ChartRecords
                .Where(record => record.DataSource?.FormulaProjectionFailureToken != null)
                .Select(record => $"Token:0x{record.DataSource!.FormulaProjectionFailureToken!.Value:X2}"));
            ChartDataSourceFormulaProjectionFailuresByTokenName = CountByCode(workbook.ChartRecords
                .Where(record => !string.IsNullOrWhiteSpace(record.DataSource?.FormulaProjectionFailureTokenName))
                .Select(record => record.DataSource!.FormulaProjectionFailureTokenName!));
            ChartDataSourceFormulaProjectionFailuresByOffset = CountByCode(workbook.ChartRecords
                .Where(record => record.DataSource?.FormulaProjectionFailureTokenOffset != null)
                .Select(record => $"Offset:{record.DataSource!.FormulaProjectionFailureTokenOffset!.Value}"));
            ChartDataSourceStates = CountByCode(workbook.ChartRecords
                .Where(record => record.DataSource != null)
                .Select(GetChartDataSourceStateKey));
            ChartDataFormatTargets = CountByCode(workbook.ChartRecords
                .Where(record => !string.IsNullOrWhiteSpace(record.DataFormatTarget))
                .Select(record => record.DataFormatTarget!));
            ChartDataFormatSeriesIndexes = CountByCode(workbook.ChartRecords
                .Where(record => record.DataFormatSeriesIndex.HasValue)
                .Select(record => $"SeriesIndex:{record.DataFormatSeriesIndex!.Value}"));
            ChartDataFormatPointIndexes = CountByCode(workbook.ChartRecords
                .Where(record => record.DataFormatPointIndex.HasValue)
                .Select(record => $"PointIndex:{record.DataFormatPointIndex!.Value}"));
            ChartDataFormatOrders = CountByCode(workbook.ChartRecords
                .Where(record => record.DataFormatOrder.HasValue)
                .Select(record => $"Order:{record.DataFormatOrder!.Value}"));
            ChartDataFormatStates = CountByCode(workbook.ChartRecords
                .Where(record => record.DataFormatPointIndex.HasValue
                    || record.DataFormatSeriesIndex.HasValue
                    || record.DataFormatOrder.HasValue
                    || !string.IsNullOrWhiteSpace(record.DataFormatTarget))
                .Select(GetChartDataFormatStateKey));
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
            ChartDataTableReservedStates = CountByCode(workbook.ChartRecords
                .Where(record => record.DataTableOptions != null)
                .Select(record => record.DataTableOptions!.ReservedState));
            ChartErrorBarDirections = CountByCode(workbook.ChartRecords
                .Where(record => record.ErrorBarOptions != null)
                .Select(record => record.ErrorBarOptions!.DirectionName));
            ChartErrorBarValueSources = CountByCode(workbook.ChartRecords
                .Where(record => record.ErrorBarOptions != null)
                .Select(record => record.ErrorBarOptions!.ValueSourceName));
            ChartErrorBarValues = CountByCode(workbook.ChartRecords
                .Where(record => record.ErrorBarOptions != null)
                .Select(record => $"Value:{FormatDouble(record.ErrorBarOptions!.Value)};CustomCount:{record.ErrorBarOptions.CustomValueCount}"));
            ChartErrorBarStates = CountByCode(workbook.ChartRecords
                .Where(record => record.ErrorBarOptions != null)
                .Select(GetChartErrorBarStateKey));
            ChartErrorBarReservedStates = CountByCode(workbook.ChartRecords
                .Where(record => record.ErrorBarOptions != null)
                .Select(record => record.ErrorBarOptions!.HasExpectedReservedValue ? "ReservedExpected" : "ReservedUnexpected"));
            ChartBarOverlapPercentages = CountByCode(workbook.ChartRecords
                .Where(record => record.BarOptions != null)
                .Select(record => $"Overlap:{record.BarOptions!.OverlapPercentage}"));
            ChartBarGapWidths = CountByCode(workbook.ChartRecords
                .Where(record => record.BarOptions != null)
                .Select(record => $"Gap:{record.BarOptions!.GapWidthPercentage}"));
            ChartBarStates = CountByCode(workbook.ChartRecords
                .Where(record => record.BarOptions != null)
                .Select(GetChartBarStateKey));
            ChartLineStates = CountByCode(workbook.ChartRecords
                .Where(record => record.LineOptions != null)
                .Select(GetChartLineStateKey));
            ChartLineReservedStates = CountByCode(workbook.ChartRecords
                .Where(record => record.LineOptions != null)
                .Select(record => record.LineOptions!.HasZeroReservedBits ? "ReservedZero" : "ReservedNonZero"));
            ChartLinePercentStackedStates = CountByCode(workbook.ChartRecords
                .Where(record => record.LineOptions != null)
                .Select(record => record.LineOptions!.HasValidPercentStackedState ? "ValidPercentState" : "InvalidPercentWithoutStacking"));
            ChartAreaStates = CountByCode(workbook.ChartRecords
                .Where(record => record.AreaOptions != null)
                .Select(GetChartAreaStateKey));
            ChartAreaReservedStates = CountByCode(workbook.ChartRecords
                .Where(record => record.AreaOptions != null)
                .Select(record => record.AreaOptions!.HasZeroReservedBits ? "ReservedZero" : "ReservedNonZero"));
            ChartAreaPercentStackedStates = CountByCode(workbook.ChartRecords
                .Where(record => record.AreaOptions != null)
                .Select(record => record.AreaOptions!.HasValidPercentStackedState ? "ValidPercentState" : "InvalidPercentWithoutStacking"));
            ChartBopPopSubtypes = CountByCode(workbook.ChartRecords
                .Where(record => record.BopPopOptions != null)
                .Select(record => record.BopPopOptions!.SubtypeName));
            ChartBopPopSplitTypes = CountByCode(workbook.ChartRecords
                .Where(record => record.BopPopOptions != null)
                .Select(record => record.BopPopOptions!.SplitName));
            ChartBopPopSplitValues = CountByCode(workbook.ChartRecords
                .Where(record => record.BopPopOptions != null)
                .Select(record => $"Position:{record.BopPopOptions!.SplitPosition};Percent:{record.BopPopOptions.SplitPercent};Size:{record.BopPopOptions.SecondaryPieSizePercent};Gap:{record.BopPopOptions.GapPercent};Value:{FormatDouble(record.BopPopOptions.SplitValue)}"));
            ChartBopPopStates = CountByCode(workbook.ChartRecords
                .Where(record => record.BopPopOptions != null)
                .Select(GetChartBopPopStateKey));
            ChartBopPopReservedStates = CountByCode(workbook.ChartRecords
                .Where(record => record.BopPopOptions != null)
                .Select(record => record.BopPopOptions!.HasZeroReservedBits ? "ReservedZero" : "ReservedNonZero"));
            ChartBopPopCustomDataPointCounts = CountByCode(workbook.ChartRecords
                .Where(record => record.BopPopCustomSplit != null)
                .Select(record => $"DataPoints:{record.BopPopCustomSplit!.DataPointCount}"));
            ChartBopPopCustomSecondaryCounts = CountByCode(workbook.ChartRecords
                .Where(record => record.BopPopCustomSplit != null)
                .Select(record => $"Secondary:{record.BopPopCustomSplit!.SecondaryDataPointIndexes.Count}"));
            ChartBopPopCustomSecondaryIndexes = CountByCode(workbook.ChartRecords
                .Where(record => record.BopPopCustomSplit != null)
                .Select(record => record.BopPopCustomSplit!.SecondaryDataPointIndexes.Count == 0 ? "Secondary:None" : "Secondary:" + string.Join(",", record.BopPopCustomSplit.SecondaryDataPointIndexes)));
            ChartBopPopCustomCompletionStates = CountByCode(workbook.ChartRecords
                .Where(record => record.BopPopCustomSplit != null)
                .Select(record => record.BopPopCustomSplit!.HasCompleteBitmap ? "Complete" : "Truncated"));
            ChartBopPopCustomStates = CountByCode(workbook.ChartRecords
                .Where(record => record.BopPopCustomSplit != null)
                .Select(GetChartBopPopCustomStateKey));
            ChartThreeDimensionalViewAngles = CountByCode(workbook.ChartRecords
                .Where(record => record.ThreeDimensionalOptions != null)
                .Select(record => $"Rotation:{record.ThreeDimensionalOptions!.RotationDegrees};Elevation:{record.ThreeDimensionalOptions.ElevationDegrees}"));
            ChartThreeDimensionalScaleValues = CountByCode(workbook.ChartRecords
                .Where(record => record.ThreeDimensionalOptions != null)
                .Select(record => $"FieldOfView:{record.ThreeDimensionalOptions!.FieldOfViewDegrees};Height:{record.ThreeDimensionalOptions.HeightPercent};Depth:{record.ThreeDimensionalOptions.DepthPercent};Gap:{record.ThreeDimensionalOptions.GapWidthPercent}"));
            ChartThreeDimensionalStates = CountByCode(workbook.ChartRecords
                .Where(record => record.ThreeDimensionalOptions != null)
                .Select(GetChartThreeDimensionalStateKey));
            ChartThreeDimensionalReservedStates = CountByCode(workbook.ChartRecords
                .Where(record => record.ThreeDimensionalOptions != null)
                .Select(record => record.ThreeDimensionalOptions!.HasZeroReservedBits ? "ReservedZero" : "ReservedNonZero"));
            ChartThreeDimensionalBarShapeRisers = CountByCode(workbook.ChartRecords
                .Where(record => record.ThreeDimensionalBarShapeOptions != null)
                .Select(record => record.ThreeDimensionalBarShapeOptions!.RiserName));
            ChartThreeDimensionalBarShapeTapers = CountByCode(workbook.ChartRecords
                .Where(record => record.ThreeDimensionalBarShapeOptions != null)
                .Select(record => record.ThreeDimensionalBarShapeOptions!.TaperName));
            ChartThreeDimensionalBarShapeStates = CountByCode(workbook.ChartRecords
                .Where(record => record.ThreeDimensionalBarShapeOptions != null)
                .Select(GetChartThreeDimensionalBarShapeStateKey));
            ChartScatterBubbleSizeRatios = CountByCode(workbook.ChartRecords
                .Where(record => record.ScatterOptions != null)
                .Select(record => $"Ratio:{record.ScatterOptions!.BubbleSizeRatio}"));
            ChartScatterBubbleSizeRepresentations = CountByCode(workbook.ChartRecords
                .Where(record => record.ScatterOptions != null)
                .Select(record => record.ScatterOptions!.BubbleSizeRepresentationName));
            ChartScatterBubbleSizeRatioStates = CountByCode(workbook.ChartRecords
                .Where(record => record.ScatterOptions != null)
                .Select(record => record.ScatterOptions!.HasValidBubbleSizeRatio ? "Valid" : "Invalid"));
            ChartScatterStates = CountByCode(workbook.ChartRecords
                .Where(record => record.ScatterOptions != null)
                .Select(GetChartScatterStateKey));
            ChartFontBasisScaleBasis = CountByCode(workbook.ChartRecords
                .Where(record => record.FontBasisOptions != null)
                .Select(record => record.FontBasisOptions!.ScaleBasisName));
            ChartFontBasisFontIndexes = CountByCode(workbook.ChartRecords
                .Where(record => record.FontBasisOptions != null)
                .Select(record => $"FontIndex:{record.FontBasisOptions!.FontIndex}"));
            ChartFontBasisStates = CountByCode(workbook.ChartRecords
                .Where(record => record.FontBasisOptions != null)
                .Select(GetChartFontBasisStateKey));
            ChartLayout12ModePairs = CountByCode(workbook.ChartRecords
                .Where(record => record.Layout12 != null)
                .Select(record => $"X:{record.Layout12!.XModeName};Y:{record.Layout12.YModeName};Width:{record.Layout12.WidthModeName};Height:{record.Layout12.HeightModeName}"));
            ChartLayout12AutoLayoutTypes = CountByCode(workbook.ChartRecords
                .Where(record => record.Layout12 != null)
                .Select(record => record.Layout12!.AutomaticLayoutTypeName));
            ChartLayout12Checksums = CountByCode(workbook.ChartRecords
                .Where(record => record.Layout12 != null)
                .Select(record => $"Checksum:0x{record.Layout12!.Checksum:X8}"));
            ChartLayout12Rectangles = CountByCode(workbook.ChartRecords
                .Where(record => record.Layout12 != null)
                .Select(record => $"X:{FormatDouble(record.Layout12!.X)};Y:{FormatDouble(record.Layout12.Y)};Width:{FormatDouble(record.Layout12.Width)};Height:{FormatDouble(record.Layout12.Height)}"));
            ChartPlotAreaLayout12Targets = CountByCode(workbook.ChartRecords
                .Where(record => record.PlotAreaLayout12 != null)
                .Select(record => record.PlotAreaLayout12!.TargetName));
            ChartPlotAreaLayout12ModePairs = CountByCode(workbook.ChartRecords
                .Where(record => record.PlotAreaLayout12 != null)
                .Select(record => $"X:{record.PlotAreaLayout12!.XModeName};Y:{record.PlotAreaLayout12.YModeName};Width:{record.PlotAreaLayout12.WidthModeName};Height:{record.PlotAreaLayout12.HeightModeName}"));
            ChartPlotAreaLayout12Checksums = CountByCode(workbook.ChartRecords
                .Where(record => record.PlotAreaLayout12 != null)
                .Select(record => $"Checksum:0x{record.PlotAreaLayout12!.Checksum:X8}"));
            ChartPlotAreaLayout12Bounds = CountByCode(workbook.ChartRecords
                .Where(record => record.PlotAreaLayout12 != null)
                .Select(record => $"X:{record.PlotAreaLayout12!.UpperLeftX};Y:{record.PlotAreaLayout12.UpperLeftY};Width:{record.PlotAreaLayout12.WidthSprc};Height:{record.PlotAreaLayout12.HeightSprc}"));
            ChartPlotAreaLayout12Rectangles = CountByCode(workbook.ChartRecords
                .Where(record => record.PlotAreaLayout12 != null)
                .Select(record => $"X:{FormatDouble(record.PlotAreaLayout12!.X)};Y:{FormatDouble(record.PlotAreaLayout12.Y)};Width:{FormatDouble(record.PlotAreaLayout12.Width)};Height:{FormatDouble(record.PlotAreaLayout12.Height)}"));
            ChartFutureRecordInfoVersions = CountByCode(workbook.ChartRecords
                .Where(record => record.FutureRecordInfo != null)
                .Select(record => $"Originator:{record.FutureRecordInfo!.OriginatorVersionName};Writer:{record.FutureRecordInfo.WriterVersionName}"));
            ChartFutureRecordInfoRangeCounts = CountByCode(workbook.ChartRecords
                .Where(record => record.FutureRecordInfo != null)
                .Select(record => $"Ranges:{record.FutureRecordInfo!.Ranges.Count}"));
            ChartFutureRecordInfoRanges = CountByCode(workbook.ChartRecords
                .Where(record => record.FutureRecordInfo != null)
                .SelectMany(record => record.FutureRecordInfo!.Ranges.Select(range => range.RangeKey)));
            ChartFutureBlockDirections = CountByCode(workbook.ChartRecords
                .Where(record => record.FutureBlock != null)
                .Select(record => record.FutureBlock!.DirectionName));
            ChartFutureBlockObjectKinds = CountByCode(workbook.ChartRecords
                .Where(record => record.FutureBlock != null)
                .Select(record => record.FutureBlock!.ObjectKindName));
            ChartFutureBlockScopes = CountByCode(workbook.ChartRecords
                .Where(record => record.FutureBlock != null)
                .Select(record => record.FutureBlock!.ScopeKey));
            ChartUnitsReservedValues = CountByCode(workbook.ChartRecords
                .Where(record => record.Units != null)
                .Select(record => $"Reserved:0x{record.Units!.Reserved:X4}"));
            ChartUnitsReservedStates = CountByCode(workbook.ChartRecords
                .Where(record => record.Units != null)
                .Select(record => record.Units!.HasZeroReservedValue ? "ReservedZero" : "ReservedNonZero"));
            ChartXmlTokenChainDeclaredByteCounts = CountByCode(workbook.ChartRecords
                .Where(record => record.XmlTokenChain != null)
                .Select(record => $"DeclaredBytes:{record.XmlTokenChain!.DeclaredByteCount}"));
            ChartXmlTokenChainFirstSegmentByteCounts = CountByCode(workbook.ChartRecords
                .Where(record => record.XmlTokenChain != null)
                .Select(record => $"FirstSegmentBytes:{record.XmlTokenChain!.FirstSegmentByteCount}"));
            ChartXmlTokenChainCompletionStates = CountByCode(workbook.ChartRecords
                .Where(record => record.XmlTokenChain != null)
                .Select(record => record.XmlTokenChain!.IsCompleteInRecord ? "CompleteInRecord" : "RequiresContinuation"));
            ChartXmlTokenChainTrailingStates = CountByCode(workbook.ChartRecords
                .Where(record => record.XmlTokenChain != null)
                .Select(record => record.XmlTokenChain!.HasZeroTrailingUnusedValue ? "TrailingUnusedZero" : "TrailingUnusedNonZero"));
            ChartSheetPropertyEmptyCellModes = CountByCode(workbook.ChartRecords
                .Where(record => record.SheetProperties != null && record.SheetProperties.HasKnownEmptyCellPlottingMode)
                .Select(record => record.SheetProperties!.EmptyCellPlottingModeName));
            ChartSheetPropertyStates = CountByCode(workbook.ChartRecords
                .Where(record => record.SheetProperties != null)
                .Select(GetChartSheetPropertyStateKey));
            ChartLineFormatStyles = CountByCode(workbook.ChartRecords
                .Where(record => record.LineFormat != null)
                .Select(record => record.LineFormat!.StyleName));
            ChartLineFormatWeights = CountByCode(workbook.ChartRecords
                .Where(record => record.LineFormat != null)
                .Select(record => record.LineFormat!.WeightName));
            ChartLineFormatColors = CountByCode(workbook.ChartRecords
                .Where(record => record.LineFormat != null)
                .Select(record => record.LineFormat!.RgbHex));
            ChartLineFormatColorIndexes = CountByCode(workbook.ChartRecords
                .Where(record => record.LineFormat != null)
                .Select(record => $"ColorIndex:{record.LineFormat!.ColorIndex}"));
            ChartLineFormatStates = CountByCode(workbook.ChartRecords
                .Where(record => record.LineFormat != null)
                .Select(GetChartLineFormatStateKey));
            ChartAreaFormatPatterns = CountByCode(workbook.ChartRecords
                .Where(record => record.AreaFormat != null)
                .Select(record => record.AreaFormat!.PatternName));
            ChartAreaFormatColors = CountByCode(workbook.ChartRecords
                .Where(record => record.AreaFormat != null)
                .SelectMany(record => GetChartAreaFormatColorKeys(record.AreaFormat!)));
            ChartAreaFormatColorIndexes = CountByCode(workbook.ChartRecords
                .Where(record => record.AreaFormat != null)
                .SelectMany(record => GetChartAreaFormatColorIndexKeys(record.AreaFormat!)));
            ChartAreaFormatStates = CountByCode(workbook.ChartRecords
                .Where(record => record.AreaFormat != null)
                .Select(GetChartAreaFormatStateKey));
            ChartMarkerFormatTypes = CountByCode(workbook.ChartRecords
                .Where(record => record.MarkerFormat != null)
                .Select(record => record.MarkerFormat!.MarkerTypeName));
            ChartMarkerFormatSizes = CountByCode(workbook.ChartRecords
                .Where(record => record.MarkerFormat != null)
                .Select(record => $"SizeTwips:{record.MarkerFormat!.SizeTwips}"));
            ChartMarkerFormatColors = CountByCode(workbook.ChartRecords
                .Where(record => record.MarkerFormat != null)
                .SelectMany(record => GetChartMarkerFormatColorKeys(record.MarkerFormat!)));
            ChartMarkerFormatColorIndexes = CountByCode(workbook.ChartRecords
                .Where(record => record.MarkerFormat != null)
                .SelectMany(record => GetChartMarkerFormatColorIndexKeys(record.MarkerFormat!)));
            ChartMarkerFormatStates = CountByCode(workbook.ChartRecords
                .Where(record => record.MarkerFormat != null)
                .Select(GetChartMarkerFormatStateKey));
            ChartPieFormatExplosions = CountByCode(workbook.ChartRecords
                .Where(record => record.PieFormat != null)
                .Select(record => $"ExplosionPercent:{record.PieFormat!.ExplosionPercentage}"));
            ChartSeriesFormatFlags = CountByCode(workbook.ChartRecords
                .Where(record => record.SeriesFormat != null)
                .SelectMany(record => record.SeriesFormat!.FlagNames));
            ChartSeriesFormatStates = CountByCode(workbook.ChartRecords
                .Where(record => record.SeriesFormat != null)
                .Select(GetChartSeriesFormatStateKey));
            ChartSeriesFormatReservedValues = CountByCode(workbook.ChartRecords
                .Where(record => record.SeriesFormat != null)
                .Select(record => $"Reserved:0x{record.SeriesFormat!.Reserved:X4}"));
            ChartSeriesFormatReservedStates = CountByCode(workbook.ChartRecords
                .Where(record => record.SeriesFormat != null)
                .Select(record => record.SeriesFormat!.HasZeroReservedBits ? "ReservedZero" : "ReservedNonZero"));
            ChartClientColorPaletteDeclaredCounts = CountByCode(workbook.ChartRecords
                .Where(record => record.ClientColorPalette != null)
                .Select(record => $"Declared:{record.ClientColorPalette!.DeclaredColorCount}"));
            ChartClientColorPaletteDecodedCounts = CountByCode(workbook.ChartRecords
                .Where(record => record.ClientColorPalette != null)
                .Select(record => $"Decoded:{record.ClientColorPalette!.DecodedColorCount}"));
            ChartClientColorPaletteCompletenessStates = CountByCode(workbook.ChartRecords
                .Where(record => record.ClientColorPalette != null)
                .Select(record => record.ClientColorPalette!.HasCompleteColorList ? "Complete" : "Truncated"));
            ChartClientColorPaletteExpectedCountStates = CountByCode(workbook.ChartRecords
                .Where(record => record.ClientColorPalette != null)
                .Select(record => record.ClientColorPalette!.HasExpectedColorCount ? "ExpectedThreeColors" : "UnexpectedColorCount"));
            ChartClientColorPaletteColors = CountByCode(workbook.ChartRecords
                .Where(record => record.ClientColorPalette != null)
                .SelectMany(record => GetChartClientColorPaletteColorKeys(record.ClientColorPalette!)));
            ChartGelFrameOfficeArtRecordsByType = CountByCode(workbook.ChartRecords
                .Where(record => record.GelFrame != null)
                .SelectMany(record => record.GelFrame!.OfficeArtRecords.Select(officeArtRecord => officeArtRecord.RecordTypeName)));
            ChartGelFrameOfficeArtRecordsByContainerState = CountByCode(workbook.ChartRecords
                .Where(record => record.GelFrame != null)
                .SelectMany(record => record.GelFrame!.OfficeArtRecords.Select(officeArtRecord => officeArtRecord.IsContainer ? "Container" : "Leaf")));
            ChartGelFrameShapePropertyCounts = CountByCode(workbook.ChartRecords
                .Where(record => record.GelFrame != null)
                .Select(record => $"Properties:{record.GelFrame!.ShapePropertyCount}"));
            ChartGelFrameShapePropertiesByName = CountByCode(workbook.ChartRecords
                .Where(record => record.GelFrame != null)
                .SelectMany(record => record.GelFrame!.ShapeProperties.Select(property => property.PropertyName)));
            ChartGelFrameShapePropertiesByGroup = CountByCode(workbook.ChartRecords
                .Where(record => record.GelFrame != null)
                .SelectMany(record => record.GelFrame!.ShapeProperties.Select(property => property.PropertyGroupName)));
            ChartGelFrameShapePropertiesByFlagState = CountByCode(workbook.ChartRecords
                .Where(record => record.GelFrame != null)
                .SelectMany(record => record.GelFrame!.ShapeProperties.Select(property => $"Complex:{property.IsComplex};Blip:{property.IsBlipId}")));
            ChartGelFrameShapePropertiesByValue = CountByCode(workbook.ChartRecords
                .Where(record => record.GelFrame != null)
                .SelectMany(record => record.GelFrame!.ShapeProperties.Select(property => $"{property.PropertyIdKey};Value:0x{property.Value:X8}")));
            ChartAttachedLabelFlags = CountByCode(workbook.ChartRecords
                .Where(record => record.AttachedLabel != null)
                .SelectMany(record => record.AttachedLabel!.FlagNames));
            ChartAttachedLabelStates = CountByCode(workbook.ChartRecords
                .Where(record => record.AttachedLabel != null)
                .Select(GetChartAttachedLabelStateKey));
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
            ChartLegendSpacingStates = CountByCode(workbook.ChartRecords
                .Where(record => record.Legend != null)
                .Select(record => record.Legend!.HasExpectedSpacing ? "ExpectedSpacing" : "UnexpectedSpacing"));
            ChartLegendReservedStates = CountByCode(workbook.ChartRecords
                .Where(record => record.Legend != null)
                .Select(record => record.Legend!.HasValidReservedBits ? "ReservedExpected" : "ReservedUnexpected"));
            ChartLegendAutoPositionStates = CountByCode(workbook.ChartRecords
                .Where(record => record.Legend != null)
                .Select(record => record.Legend!.HasValidAutoPositionState ? "AutoPositionConsistent" : "AutoPositionInconsistent"));
            ChartLegendDataTableStates = CountByCode(workbook.ChartRecords
                .Where(record => record.Legend != null)
                .Select(record => record.Legend!.HasValidDataTableState ? "DataTableConsistent" : "DataTableWithoutVerticalLayout"));
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
            ChartPositionSemanticTypes = CountByCode(workbook.ChartRecords
                .Where(record => record.Position != null)
                .Select(record => record.Position!.SemanticTypeName));
            ChartPositionCoordinateMeanings = CountByCode(workbook.ChartRecords
                .Where(record => record.Position != null)
                .Select(record => $"X1Y1:{record.Position!.X1Y1MeaningName};X2Y2:{record.Position.X2Y2MeaningName}"));
            ChartPositionIgnoredCoordinateStates = CountByCode(workbook.ChartRecords
                .Where(record => record.Position != null)
                .Select(record => record.Position!.IgnoredCoordinateStateName));
            ChartPositionKnownSemanticStates = CountByCode(workbook.ChartRecords
                .Where(record => record.Position != null)
                .Select(record => $"Known:{record.Position!.HasKnownSemanticCombination}"));
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
                .Select(subRecord => subRecord.CompletionState));
            DrawingFutureRecordWrappedTypes = CountByCode(workbook.DrawingRecords
                .Where(record => record.FutureRecordHeader != null)
                .Select(record => $"{record.RecordName}|0x{record.FutureRecordHeader!.WrappedRecordType:X4}"));
            DrawingFutureRecordFlags = CountByCode(workbook.DrawingRecords
                .Where(record => record.FutureRecordHeader != null)
                .Select(record => $"{record.RecordName}|Flags:0x{record.FutureRecordHeader!.Flags:X4}"));
            DrawingFutureRecordReferenceStates = CountByCode(workbook.DrawingRecords
                .Where(record => record.FutureRecordHeader != null)
                .Select(record => record.FutureRecordHeader!.HasRange ? $"{record.RecordName}|HasRange" : $"{record.RecordName}|NoRange"));
            DrawingFutureRecordRanges = CountByCode(workbook.DrawingRecords
                .Where(record => record.FutureRecordHeader?.HasRange == true
                    && record.FutureRecordHeader.FirstRow.HasValue
                    && record.FutureRecordHeader.LastRow.HasValue
                    && record.FutureRecordHeader.FirstColumn.HasValue
                    && record.FutureRecordHeader.LastColumn.HasValue)
                .Select(GetDrawingFutureRecordRangeKey));
            DrawingFutureRecordStreamByteCounts = CountByCode(workbook.DrawingRecords
                .Where(record => record.FutureRecordHeader != null)
                .Select(record => $"{record.RecordName}|StreamBytes:{record.FutureRecordHeader!.StreamByteCount}"));
            DrawingHeaderFooterPictureHeaderStates = CountByCode(workbook.DrawingRecords
                .Where(record => record.HeaderFooterPicture != null)
                .Select(record => record.HeaderFooterPicture!.HeaderState));
            DrawingHeaderFooterPictureDrawingKinds = CountByCode(workbook.DrawingRecords
                .Where(record => record.HeaderFooterPicture != null)
                .Select(record => record.HeaderFooterPicture!.DrawingKindName));
            DrawingHeaderFooterPictureContinuationStates = CountByCode(workbook.DrawingRecords
                .Where(record => record.HeaderFooterPicture != null)
                .Select(record => record.HeaderFooterPicture!.ContinuationState));
            DrawingHeaderFooterPictureFutureRecordFlags = CountByCode(workbook.DrawingRecords
                .Where(record => record.HeaderFooterPicture != null)
                .Select(record => $"Flags:0x{record.HeaderFooterPicture!.FutureRecordFlags:X4}"));
            DrawingHeaderFooterPictureDrawingByteCounts = CountByCode(workbook.DrawingRecords
                .Where(record => record.HeaderFooterPicture != null)
                .Select(record => $"DrawingBytes:{record.HeaderFooterPicture!.DrawingByteCount}"));
            DrawingTextObjectAlignments = CountByCode(workbook.DrawingRecords
                .Where(record => record.TextObject != null)
                .Select(record => $"Horizontal:{record.TextObject!.HorizontalAlignmentName};Vertical:{record.TextObject.VerticalAlignmentName}"));
            DrawingTextObjectRotations = CountByCode(workbook.DrawingRecords
                .Where(record => record.TextObject != null)
                .Select(record => record.TextObject!.RotationName));
            DrawingTextObjectTextLengths = CountByCode(workbook.DrawingRecords
                .Where(record => record.TextObject != null)
                .Select(record => $"Characters:{record.TextObject!.TextCharacterCount}"));
            DrawingTextObjectFormattingRunByteCounts = CountByCode(workbook.DrawingRecords
                .Where(record => record.TextObject != null)
                .Select(record => $"RunBytes:{record.TextObject!.FormattingRunByteCount}"));
            DrawingTextObjectFormulaByteCounts = CountByCode(workbook.DrawingRecords
                .Where(record => record.TextObject != null)
                .Select(record => $"FormulaBytes:{record.TextObject!.FormulaByteCount}"));
            DrawingTextObjectFlags = CountByCode(workbook.DrawingRecords
                .Where(record => record.TextObject != null)
                .SelectMany(record => GetDrawingTextObjectFlagKeys(record.TextObject!)));
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
            DrawingGroupBlocksByMaxShapeId = CountByCode(workbook.DrawingRecords
                .SelectMany(record => record.DrawingGroupBlocks)
                .Select(block => $"MaxShapeId:{block.MaxShapeId}"));
            DrawingGroupBlocksByDeclaredIdentifierClusterCount = CountByCode(workbook.DrawingRecords
                .SelectMany(record => record.DrawingGroupBlocks)
                .Select(block => $"DeclaredIdentifierClusters:{block.DeclaredIdentifierClusterCount}"));
            DrawingGroupBlocksByDecodedIdentifierClusterCount = CountByCode(workbook.DrawingRecords
                .SelectMany(record => record.DrawingGroupBlocks)
                .Select(block => $"DecodedIdentifierClusters:{block.IdentifierClusters.Count}"));
            DrawingGroupBlocksBySavedShapeCount = CountByCode(workbook.DrawingRecords
                .SelectMany(record => record.DrawingGroupBlocks)
                .Select(block => $"SavedShapes:{block.SavedShapeCount}"));
            DrawingGroupBlocksBySavedDrawingCount = CountByCode(workbook.DrawingRecords
                .SelectMany(record => record.DrawingGroupBlocks)
                .Select(block => $"SavedDrawings:{block.SavedDrawingCount}"));
            DrawingIdentifierClustersByDrawingId = CountByCode(workbook.DrawingRecords
                .SelectMany(record => record.DrawingGroupBlocks)
                .SelectMany(block => block.IdentifierClusters)
                .Select(cluster => $"DrawingId:{cluster.DrawingId}"));
            DrawingIdentifierClustersByCurrentShapeId = CountByCode(workbook.DrawingRecords
                .SelectMany(record => record.DrawingGroupBlocks)
                .SelectMany(block => block.IdentifierClusters)
                .Select(cluster => $"CurrentShapeId:{cluster.CurrentShapeId}"));
            DrawingGroupInfosByDrawingId = CountByCode(workbook.DrawingRecords
                .SelectMany(record => record.DrawingGroupInfos)
                .Select(info => $"DrawingId:{info.DrawingId}"));
            DrawingGroupInfosByShapeCount = CountByCode(workbook.DrawingRecords
                .SelectMany(record => record.DrawingGroupInfos)
                .Select(info => $"Shapes:{info.ShapeCount}"));
            DrawingGroupInfosByLastShapeId = CountByCode(workbook.DrawingRecords
                .SelectMany(record => record.DrawingGroupInfos)
                .Select(info => $"LastShapeId:{info.LastShapeId}"));
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
            DrawingShapeComplexPropertiesByText = CountByCode(workbook.DrawingRecords
                .SelectMany(record => record.ShapeProperties)
                .Where(property => !string.IsNullOrWhiteSpace(property.ComplexText))
                .Select(property => $"{property.PropertyName}:{property.ComplexText!}"));
            DrawingBlipStoreEntriesByType = CountByCode(workbook.DrawingRecords
                .SelectMany(record => record.BlipStoreEntries)
                .Select(entry => entry.RecordInstanceBlipTypeName));
            DrawingBlipStoreEntriesByLocation = CountByCode(workbook.DrawingRecords
                .SelectMany(record => record.BlipStoreEntries.Select(_ => GetDrawingRecordLocationKey(record))));
            DrawingBlipStoreEntriesByTypeAndLocation = CountByCode(workbook.DrawingRecords
                .SelectMany(record => record.BlipStoreEntries
                    .Select(entry => $"{GetDrawingRecordLocationKey(record)}|{entry.RecordInstanceBlipTypeName}")));
            DrawingBlipStoreEntriesByUid = CountByCode(workbook.DrawingRecords
                .SelectMany(record => record.BlipStoreEntries)
                .Where(entry => !string.IsNullOrWhiteSpace(entry.UidHex))
                .Select(entry => entry.UidHex!));
            DrawingBlipStoreEntriesByEmbeddedRecordType = CountByCode(workbook.DrawingRecords
                .SelectMany(record => record.BlipStoreEntries)
                .Where(entry => !string.IsNullOrWhiteSpace(entry.EmbeddedBlipRecordTypeName))
                .Select(entry => entry.EmbeddedBlipRecordTypeName!));
            DrawingBlipStoreEntriesByEmbeddedPayloadAvailableLength = CountByCode(workbook.DrawingRecords
                .SelectMany(record => record.BlipStoreEntries)
                .Where(entry => entry.EmbeddedBlipPayloadAvailableLength.HasValue)
                .Select(entry => $"AvailableBytes:{entry.EmbeddedBlipPayloadAvailableLength!.Value}"));
            DrawingBlipStoreEntriesByEmbeddedPayloadHash = CountByCode(workbook.DrawingRecords
                .SelectMany(record => record.BlipStoreEntries)
                .Where(entry => !string.IsNullOrWhiteSpace(entry.EmbeddedBlipPayloadSha256))
                .Select(entry => entry.EmbeddedBlipPayloadSha256!));
            DrawingBlipStoreEntriesBySize = CountByCode(workbook.DrawingRecords
                .SelectMany(record => record.BlipStoreEntries)
                .Where(entry => entry.SizeBytes.HasValue)
                .Select(entry => $"SizeBytes:{entry.SizeBytes!.Value}"));
            DrawingBlipStoreEntriesByReferenceCount = CountByCode(workbook.DrawingRecords
                .SelectMany(record => record.BlipStoreEntries)
                .Where(entry => entry.ReferenceCount.HasValue)
                .Select(entry => $"References:{entry.ReferenceCount!.Value}"));
            DrawingShapeBlipPropertiesByLocation = CountByCode(workbook.DrawingRecords
                .SelectMany(record => record.ShapeProperties
                    .Where(IsBlipShapeProperty)
                    .Select(_ => GetDrawingRecordLocationKey(record))));
            DrawingShapeBlipPropertiesByNameAndValue = CountByCode(workbook.DrawingRecords
                .SelectMany(record => record.ShapeProperties)
                .Where(IsBlipShapeProperty)
                .Select(property => $"{property.PropertyName};Value:0x{property.Value:X8}"));
            DrawingPictureBlipReferencesByLocation = CountByCode(workbook.DrawingRecords
                .SelectMany(record => record.ShapeProperties
                    .Where(IsPictureBlipReferenceProperty)
                    .Select(_ => GetDrawingRecordLocationKey(record))));
            DrawingPictureBlipReferencesByValue = CountByCode(workbook.DrawingRecords
                .SelectMany(record => record.ShapeProperties)
                .Where(IsPictureBlipReferenceProperty)
                .Select(property => $"BlipId:{property.Value}"));
            DrawingPictureStates = CountByCode(GetDrawingPictureStateKeys(workbook));
            DrawingPictureCountStates = CountByCode(GetDrawingPictureCountStateKeys(workbook));
            DrawingShapeEntriesByType = CountByCode(workbook.DrawingRecords
                .SelectMany(record => record.ShapeEntries)
                .Select(shape => shape.ShapeTypeName));
            DrawingShapeEntriesById = CountByCode(workbook.DrawingRecords
                .SelectMany(record => record.ShapeEntries)
                .Select(shape => $"ShapeId:{shape.ShapeId}"));
            DrawingShapeEntriesByFlags = CountByCode(workbook.DrawingRecords
                .SelectMany(record => record.ShapeEntries)
                .Select(shape => $"Flags:0x{shape.Flags:X8}"));
            DrawingShapeEntriesByReservedState = CountByCode(workbook.DrawingRecords
                .SelectMany(record => record.ShapeEntries)
                .Select(shape => shape.ReservedState));
            DrawingShapeEntriesByFlagName = CountByCode(workbook.DrawingRecords
                .SelectMany(record => record.ShapeEntries)
                .SelectMany(shape => shape.FlagNames));
            DrawingAnchorEntriesByRange = CountByCode(workbook.DrawingRecords
                .SelectMany(record => record.AnchorEntries)
                .Select(GetDrawingAnchorRangeKey));
            DrawingAnchorEntriesByOffset = CountByCode(workbook.DrawingRecords
                .SelectMany(record => record.AnchorEntries)
                .Select(GetDrawingAnchorOffsetKey));
            DrawingAnchorEntriesByFlags = CountByCode(workbook.DrawingRecords
                .SelectMany(record => record.AnchorEntries)
                .Select(GetDrawingAnchorFlagsKey));
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
            CompoundFeatureEntriesByContentKind = CountByCode(workbook.CompoundFeatureRecords
                .SelectMany(record => record.EntryDetails)
                .Select(entry => entry.ContentKind.ToString()));
            CompoundFeatureEntriesByRoleAndContentKind = CountByCode(workbook.CompoundFeatureRecords
                .SelectMany(record => record.EntryDetails)
                .Select(entry => $"{entry.Role}|{entry.ContentKind}"));
            CompoundFeatureEntriesBySize = CountByCode(workbook.CompoundFeatureRecords
                .SelectMany(record => record.EntryDetails)
                .Select(GetCompoundFeatureEntrySizeKey));
            CompoundFeatureEntriesByRoleAndSize = CountByCode(workbook.CompoundFeatureRecords
                .SelectMany(record => record.EntryDetails)
                .Select(entry => $"{entry.Role}|{GetCompoundFeatureEntrySizeKey(entry)}"));
            CompoundVbaModulesByName = CountByCode(workbook.CompoundFeatureRecords.SelectMany(record => record.VbaModuleNames));
            CompoundVbaModulesByPath = CountByCode(workbook.CompoundFeatureRecords
                .SelectMany(record => record.EntryDetails)
                .Where(entry => entry.Role == LegacyXlsCompoundFeatureEntryRole.VbaModuleStream)
                .Select(entry => entry.Path));
            CompoundVbaModulesBySize = CountByCode(workbook.CompoundFeatureRecords
                .SelectMany(record => record.EntryDetails)
                .Where(entry => entry.Role == LegacyXlsCompoundFeatureEntryRole.VbaModuleStream)
                .Select(GetCompoundFeatureEntrySizeKey));
            CompoundVbaModulesByNameAndSize = CountByCode(workbook.CompoundFeatureRecords
                .SelectMany(record => record.EntryDetails)
                .Where(entry => entry.Role == LegacyXlsCompoundFeatureEntryRole.VbaModuleStream)
                .Select(entry => $"{GetCompoundFeatureEntryLeafName(entry.Path)}|{GetCompoundFeatureEntrySizeKey(entry)}"));
            CompoundVbaModulesByContentKind = CountByCode(workbook.CompoundFeatureRecords
                .SelectMany(record => record.EntryDetails)
                .Where(entry => entry.Role == LegacyXlsCompoundFeatureEntryRole.VbaModuleStream)
                .Select(entry => entry.ContentKind.ToString()));
            CompoundVbaModulesByNameAndContentKind = CountByCode(workbook.CompoundFeatureRecords
                .SelectMany(record => record.EntryDetails)
                .Where(entry => entry.Role == LegacyXlsCompoundFeatureEntryRole.VbaModuleStream)
                .Select(entry => $"{GetCompoundFeatureEntryLeafName(entry.Path)}|{entry.ContentKind}"));
            CompoundVbaModulesByCodeNameMatch = CountByCode(GetCompoundVbaModuleCodeNameMatchKeys(workbook));
            CompoundVbaModulesByCodeNameMatchAndName = CountByCode(GetCompoundVbaModuleCodeNameMatchAndNameKeys(workbook));
            CompoundVbaProjectsByModuleCount = CountByCode(workbook.CompoundFeatureRecords
                .Where(record => record.Kind == LegacyXlsCompoundFeatureRecordKind.VbaProject)
                .Select(record => $"Modules:{record.VbaModuleCount}"));
            CompoundVbaProjectsByModuleByteCount = CountByCode(workbook.CompoundFeatureRecords
                .Where(record => record.Kind == LegacyXlsCompoundFeatureRecordKind.VbaProject)
                .Select(record => $"Bytes:{record.VbaModuleByteCount.ToString(CultureInfo.InvariantCulture)}"));
            CompoundVbaProjectsByStructure = CountByCode(workbook.CompoundFeatureRecords
                .Where(record => record.Kind == LegacyXlsCompoundFeatureRecordKind.VbaProject)
                .Select(GetCompoundVbaProjectStructureKey));
            VbaProjectWorkbookStates = CountByCode(GetVbaProjectWorkbookStateKeys(workbook));
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
            CellStyleExtensionPropertiesByType = CountByCode(workbook.CellStyleExtensions
                .SelectMany(extension => extension.Properties)
                .Select(property => $"0x{property.PropertyType:X4}"));
            CellStyleExtensionPropertiesByName = CountByCode(workbook.CellStyleExtensions
                .SelectMany(extension => extension.Properties)
                .Select(property => property.PropertyTypeName));
            CellStyleExtensionPropertiesByDataByteCount = CountByCode(workbook.CellStyleExtensions
                .SelectMany(extension => extension.Properties)
                .Select(property => $"Bytes:{property.DataByteCount}"));
            CellStyleExtensionPropertiesByNumericValue = CountByCode(workbook.CellStyleExtensions
                .SelectMany(extension => extension.Properties)
                .Where(property => property.NumericValue.HasValue)
                .Select(property => $"{property.PropertyTypeName}:{property.NumericValue!.Value}"));
            CellStyleExtensionPropertiesByNumericValueName = CountByCode(workbook.CellStyleExtensions
                .SelectMany(extension => extension.Properties)
                .Where(property => !string.IsNullOrWhiteSpace(property.NumericValueName))
                .Select(property => $"{property.PropertyTypeName}:{property.NumericValueName}"));
            CellStyleExtensionPropertiesByColorType = CountByCode(workbook.CellStyleExtensions
                .SelectMany(extension => extension.Properties)
                .Where(property => !string.IsNullOrWhiteSpace(property.ColorTypeName))
                .Select(property => $"{property.PropertyTypeName}:{property.ColorTypeName}"));
            CellStyleExtensionPropertiesByColorTintShade = CountByCode(workbook.CellStyleExtensions
                .SelectMany(extension => extension.Properties)
                .Where(property => property.ColorTintShade.HasValue)
                .Select(property => $"{property.PropertyTypeName}:TintShade:{property.ColorTintShade!.Value}"));
            CellStyleExtensionPropertiesByColorValue = CountByCode(workbook.CellStyleExtensions
                .SelectMany(extension => extension.Properties)
                .Where(property => !string.IsNullOrWhiteSpace(property.ColorValueHex))
                .Select(property => $"{property.PropertyTypeName}:{property.ColorValueHex}"));
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

        /// <summary>Gets the number of decoded legacy chart sheets.</summary>
        public int ChartSheetCount { get; }

        /// <summary>Gets the number of sheet entries that were preserved as unsupported metadata.</summary>
        public int UnsupportedSheetCount { get; }

        /// <summary>Gets the number of imported cells, including blank style-only cells.</summary>
        public int CellCount { get; }

        /// <summary>Gets the number of imported formula cells.</summary>
        public int FormulaCellCount { get; }

        /// <summary>Gets the number of imported comments.</summary>
        public int CommentCount { get; }

        /// <summary>Gets parsed comments grouped by decoded OBJ common-object type id.</summary>
        internal IReadOnlyDictionary<string, int> CommentsByObjectType { get; }

        /// <summary>Gets parsed comments grouped by decoded OBJ common-object type name.</summary>
        internal IReadOnlyDictionary<string, int> CommentsByObjectTypeName { get; }

        /// <summary>Gets parsed comments grouped by decoded OBJ common-object flag bitfield.</summary>
        internal IReadOnlyDictionary<string, int> CommentsByObjectFlags { get; }

        /// <summary>Gets parsed comments grouped by decoded OBJ common-object flag name.</summary>
        internal IReadOnlyDictionary<string, int> CommentsByObjectFlagName { get; }

        /// <summary>Gets parsed comments grouped by preserved OfficeArt client-anchor start and end cell.</summary>
        internal IReadOnlyDictionary<string, int> CommentsByAnchorRange { get; }

        /// <summary>Gets parsed comments grouped by preserved OfficeArt client-anchor offsets.</summary>
        internal IReadOnlyDictionary<string, int> CommentsByAnchorOffset { get; }

        /// <summary>Gets parsed comments grouped by preserved OfficeArt client-anchor flag bitfield.</summary>
        internal IReadOnlyDictionary<string, int> CommentsByAnchorFlags { get; }

        /// <summary>Gets the number of imported hyperlinks.</summary>
        public int HyperlinkCount { get; }

        /// <summary>Gets the number of imported data validation rules.</summary>
        public int DataValidationCount { get; }

        /// <summary>Gets the number of parsed DVal data-validation collection headers.</summary>
        internal int DataValidationCollectionRecordCount { get; }

        /// <summary>Gets the number of imported conditional formatting rules.</summary>
        public int ConditionalFormattingCount { get; }

        /// <summary>Gets the number of imported AutoFilter criteria columns.</summary>
        public int AutoFilterCriteriaCount { get; }

        /// <summary>Gets non-empty worksheet feature bundles grouped by data-validation, conditional-formatting, and AutoFilter state.</summary>
        internal IReadOnlyDictionary<string, int> WorksheetFeatureStates { get; }

        /// <summary>Gets worksheet object protection states declared by ObjProtect records.</summary>
        internal IReadOnlyDictionary<string, int> WorksheetProtectionObjectStates { get; }

        /// <summary>Gets worksheet scenario protection states declared by ScenarioProtect records.</summary>
        internal IReadOnlyDictionary<string, int> WorksheetProtectionScenarioStates { get; }

        /// <summary>Gets worksheet phonetic settings grouped by worksheet name.</summary>
        internal IReadOnlyDictionary<string, int> WorksheetPhoneticSettingsBySheet { get; }

        /// <summary>Gets worksheet phonetic settings grouped by conversion type.</summary>
        internal IReadOnlyDictionary<string, int> WorksheetPhoneticSettingsByType { get; }

        /// <summary>Gets worksheet phonetic settings grouped by alignment.</summary>
        internal IReadOnlyDictionary<string, int> WorksheetPhoneticSettingsByAlignment { get; }

        /// <summary>Gets worksheet phonetic settings grouped by BIFF font id.</summary>
        internal IReadOnlyDictionary<string, int> WorksheetPhoneticSettingsByFontId { get; }

        /// <summary>Gets worksheet phonetic settings grouped by attached range count.</summary>
        internal IReadOnlyDictionary<string, int> WorksheetPhoneticSettingsByRangeCount { get; }

        /// <summary>Gets worksheet phonetic ranges grouped by worksheet name.</summary>
        internal IReadOnlyDictionary<string, int> WorksheetPhoneticRangesBySheet { get; }

        /// <summary>Gets worksheet phonetic ranges grouped by worksheet-qualified A1 range.</summary>
        internal IReadOnlyDictionary<string, int> WorksheetPhoneticRangesBySheetAndRange { get; }

        /// <summary>Gets parsed DVal collection headers grouped by worksheet name.</summary>
        internal IReadOnlyDictionary<string, int> DataValidationCollectionsBySheet { get; }

        /// <summary>Gets parsed DVal collection headers grouped by declared validation count.</summary>
        internal IReadOnlyDictionary<string, int> DataValidationCollectionsByDeclaredCount { get; }

        /// <summary>Gets parsed DVal collection headers grouped by declared-vs-parsed validation state.</summary>
        internal IReadOnlyDictionary<string, int> DataValidationCollectionStates { get; }

        /// <summary>Gets imported data validations grouped by validation type.</summary>
        internal IReadOnlyDictionary<string, int> DataValidationsByType { get; }

        /// <summary>Gets imported data validations grouped by comparison operator.</summary>
        internal IReadOnlyDictionary<string, int> DataValidationsByOperator { get; }

        /// <summary>Gets imported data validations grouped by error alert style.</summary>
        internal IReadOnlyDictionary<string, int> DataValidationsByErrorStyle { get; }

        /// <summary>Gets imported data validations grouped by blank-value handling.</summary>
        internal IReadOnlyDictionary<string, int> DataValidationsByAllowBlankState { get; }

        /// <summary>Gets imported data validations grouped by input-prompt display state.</summary>
        internal IReadOnlyDictionary<string, int> DataValidationsByInputMessageState { get; }

        /// <summary>Gets imported data validations grouped by error-alert display state.</summary>
        internal IReadOnlyDictionary<string, int> DataValidationsByErrorMessageState { get; }

        /// <summary>Gets imported data validations grouped by input prompt text presence.</summary>
        internal IReadOnlyDictionary<string, int> DataValidationsByPromptTextState { get; }

        /// <summary>Gets imported data validations grouped by error alert text presence.</summary>
        internal IReadOnlyDictionary<string, int> DataValidationsByErrorTextState { get; }

        /// <summary>Gets imported list data validations grouped by in-cell dropdown behavior.</summary>
        internal IReadOnlyDictionary<string, int> DataValidationsByDropDownState { get; }

        /// <summary>Gets imported data validations grouped by worksheet name.</summary>
        internal IReadOnlyDictionary<string, int> DataValidationsBySheet { get; }

        /// <summary>Gets imported data validations grouped by number of covered ranges.</summary>
        internal IReadOnlyDictionary<string, int> DataValidationsByRangeCount { get; }

        /// <summary>Gets imported data validations grouped by covered A1 range.</summary>
        internal IReadOnlyDictionary<string, int> DataValidationsByRange { get; }

        /// <summary>Gets imported data validations grouped by worksheet-qualified covered A1 range.</summary>
        internal IReadOnlyDictionary<string, int> DataValidationsBySheetAndRange { get; }

        /// <summary>Gets imported data validations grouped by first-formula presence.</summary>
        internal IReadOnlyDictionary<string, int> DataValidationsByFormula1State { get; }

        /// <summary>Gets imported data validations grouped by second-formula presence.</summary>
        internal IReadOnlyDictionary<string, int> DataValidationsByFormula2State { get; }

        /// <summary>Gets imported data validations grouped by combined first/second formula presence.</summary>
        internal IReadOnlyDictionary<string, int> DataValidationsByFormulaPairState { get; }

        /// <summary>Gets imported list data validations grouped by source shape.</summary>
        internal IReadOnlyDictionary<string, int> DataValidationListSourcesByKind { get; }

        /// <summary>Gets imported list data validations grouped by inline item count.</summary>
        internal IReadOnlyDictionary<string, int> DataValidationListSourcesByItemCount { get; }

        /// <summary>Gets imported list data validations grouped by source range.</summary>
        internal IReadOnlyDictionary<string, int> DataValidationListSourcesByRange { get; }

        /// <summary>Gets imported list data validations grouped by source defined name.</summary>
        internal IReadOnlyDictionary<string, int> DataValidationListSourcesByName { get; }

        /// <summary>Gets imported list data validations grouped by source sheet name.</summary>
        internal IReadOnlyDictionary<string, int> DataValidationListSourcesBySheetName { get; }

        /// <summary>Gets the number of Array formula records decoded during import.</summary>
        internal int ArrayFormulaRecordCount { get; }

        /// <summary>Gets Array formula records grouped by worksheet name.</summary>
        internal IReadOnlyDictionary<string, int> ArrayFormulasBySheet { get; }

        /// <summary>Gets Array formula records grouped by covered A1 range.</summary>
        internal IReadOnlyDictionary<string, int> ArrayFormulasByRange { get; }

        /// <summary>Gets Array formula records grouped by worksheet-qualified A1 range.</summary>
        internal IReadOnlyDictionary<string, int> ArrayFormulasBySheetAndRange { get; }

        /// <summary>Gets Array formula records grouped by declared cell count.</summary>
        internal IReadOnlyDictionary<string, int> ArrayFormulasByDeclaredCellCount { get; }

        /// <summary>Gets Array formula records grouped by the number of matched cached formula cells.</summary>
        internal IReadOnlyDictionary<string, int> ArrayFormulasByMatchedFormulaCellCount { get; }

        /// <summary>Gets Array formula records grouped by recalculation flag state.</summary>
        internal IReadOnlyDictionary<string, int> ArrayFormulasByAlwaysCalculateState { get; }

        /// <summary>Gets Array formula records grouped by whether formula text was projected onto matched cells.</summary>
        internal IReadOnlyDictionary<string, int> ArrayFormulasByProjectionState { get; }

        /// <summary>Gets Array formula records grouped by parsed token byte count.</summary>
        internal IReadOnlyDictionary<string, int> ArrayFormulasByTokenByteCount { get; }

        /// <summary>Gets Array formula records grouped by parsed-formula ancillary byte count.</summary>
        internal IReadOnlyDictionary<string, int> ArrayFormulasByExtraByteCount { get; }

        /// <summary>Gets imported conditional formatting rules grouped by rule type.</summary>
        internal IReadOnlyDictionary<string, int> ConditionalFormattingsByType { get; }

        /// <summary>Gets imported conditional formatting cell-is rules grouped by comparison operator.</summary>
        internal IReadOnlyDictionary<string, int> ConditionalFormattingsByOperator { get; }

        /// <summary>Gets imported conditional formatting rules grouped by worksheet name.</summary>
        internal IReadOnlyDictionary<string, int> ConditionalFormattingsBySheet { get; }

        /// <summary>Gets imported conditional formatting rules grouped by number of covered ranges.</summary>
        internal IReadOnlyDictionary<string, int> ConditionalFormattingsByRangeCount { get; }

        /// <summary>Gets imported conditional formatting rules grouped by covered A1 range.</summary>
        internal IReadOnlyDictionary<string, int> ConditionalFormattingsByRange { get; }

        /// <summary>Gets imported conditional formatting rules grouped by worksheet-qualified covered A1 range.</summary>
        internal IReadOnlyDictionary<string, int> ConditionalFormattingsBySheetAndRange { get; }

        /// <summary>Gets imported conditional formatting rules grouped by first-formula presence.</summary>
        internal IReadOnlyDictionary<string, int> ConditionalFormattingsByFormula1State { get; }

        /// <summary>Gets imported conditional formatting rules grouped by second-formula presence.</summary>
        internal IReadOnlyDictionary<string, int> ConditionalFormattingsByFormula2State { get; }

        /// <summary>Gets imported conditional formatting rules grouped by combined first/second formula presence.</summary>
        internal IReadOnlyDictionary<string, int> ConditionalFormattingsByFormulaPairState { get; }

        /// <summary>Gets imported conditional formatting rules grouped by whether an extension priority was decoded.</summary>
        internal IReadOnlyDictionary<string, int> ConditionalFormattingsByPriorityState { get; }

        /// <summary>Gets imported conditional formatting extension priorities grouped by priority value.</summary>
        internal IReadOnlyDictionary<string, int> ConditionalFormattingsByPriority { get; }

        /// <summary>Gets imported conditional formatting rules grouped by stop-if-true behavior.</summary>
        internal IReadOnlyDictionary<string, int> ConditionalFormattingsByStopIfTrueState { get; }

        /// <summary>Gets imported conditional formatting rules grouped by whether a differential format was attached.</summary>
        internal IReadOnlyDictionary<string, int> ConditionalFormattingsByDifferentialFormatState { get; }

        /// <summary>Gets imported conditional formatting differential formats grouped by decoded fill shape.</summary>
        internal IReadOnlyDictionary<string, int> ConditionalFormattingsByDifferentialFill { get; }

        /// <summary>Gets imported conditional formatting differential formats grouped by decoded font shape.</summary>
        internal IReadOnlyDictionary<string, int> ConditionalFormattingsByDifferentialFont { get; }

        /// <summary>Gets imported conditional formatting differential formats grouped by decoded border shape.</summary>
        internal IReadOnlyDictionary<string, int> ConditionalFormattingsByDifferentialBorder { get; }

        /// <summary>Gets imported conditional formatting differential formats grouped by decoded number format shape.</summary>
        internal IReadOnlyDictionary<string, int> ConditionalFormattingsByDifferentialNumberFormat { get; }

        /// <summary>Gets the number of preserve-only conditional-formatting extension records discovered during import.</summary>
        internal int ConditionalFormattingExtensionRecordCount { get; }

        /// <summary>Gets preserve-only conditional-formatting extension records grouped by worksheet name.</summary>
        internal IReadOnlyDictionary<string, int> ConditionalFormattingExtensionsBySheet { get; }

        /// <summary>Gets preserve-only conditional-formatting extension records grouped by BIFF record type.</summary>
        internal IReadOnlyDictionary<string, int> ConditionalFormattingExtensionsByRecordType { get; }

        /// <summary>Gets preserve-only conditional-formatting extension records grouped by decoded state.</summary>
        internal IReadOnlyDictionary<string, int> ConditionalFormattingExtensionStates { get; }

        /// <summary>Gets preserve-only conditional-formatting extension records grouped by decoded priority.</summary>
        internal IReadOnlyDictionary<string, int> ConditionalFormattingExtensionPriorities { get; }

        /// <summary>Gets preserve-only conditional-formatting extension records grouped by decoded stop-if-true state.</summary>
        internal IReadOnlyDictionary<string, int> ConditionalFormattingExtensionStopIfTrueStates { get; }

        /// <summary>Gets preserve-only conditional-formatting extension records grouped by declared inline formatting byte count.</summary>
        internal IReadOnlyDictionary<string, int> ConditionalFormattingExtensionInlineFormattingByteCounts { get; }

        /// <summary>Gets preserve-only conditional-formatting extension records grouped by Dxf projection state.</summary>
        internal IReadOnlyDictionary<string, int> ConditionalFormattingExtensionDxfProjectionStates { get; }

        /// <summary>Gets imported AutoFilter criteria grouped by worksheet name.</summary>
        internal IReadOnlyDictionary<string, int> AutoFilterCriteriaBySheet { get; }

        /// <summary>Gets imported AutoFilter conditions grouped by comparison operator.</summary>
        internal IReadOnlyDictionary<string, int> AutoFilterCriteriaByOperator { get; }

        /// <summary>Gets imported AutoFilter conditions grouped by BIFF operand kind.</summary>
        internal IReadOnlyDictionary<string, int> AutoFilterCriteriaByValueKind { get; }

        /// <summary>Gets imported AutoFilter text conditions grouped by wildcard-pattern shape.</summary>
        internal IReadOnlyDictionary<string, int> AutoFilterCriteriaByTextPattern { get; }

        /// <summary>Gets imported AutoFilter criteria grouped by condition join operator.</summary>
        internal IReadOnlyDictionary<string, int> AutoFilterCriteriaByJoinOperator { get; }

        /// <summary>Gets imported AutoFilter criteria grouped by criteria kind.</summary>
        internal IReadOnlyDictionary<string, int> AutoFilterCriteriaByKind { get; }

        /// <summary>Gets imported AutoFilter criteria grouped by zero-based column id.</summary>
        internal IReadOnlyDictionary<string, int> AutoFilterCriteriaByColumn { get; }

        /// <summary>Gets imported AutoFilter criteria grouped by worksheet-qualified zero-based column id.</summary>
        internal IReadOnlyDictionary<string, int> AutoFilterCriteriaBySheetAndColumn { get; }

        /// <summary>Gets imported AutoFilter criteria grouped by condition count.</summary>
        internal IReadOnlyDictionary<string, int> AutoFilterCriteriaByConditionCount { get; }

        /// <summary>Gets imported Top/Bottom AutoFilter criteria grouped by top/bottom and items/percent shape.</summary>
        internal IReadOnlyDictionary<string, int> AutoFilterTop10Kinds { get; }

        /// <summary>Gets imported Top/Bottom AutoFilter criteria grouped by shape and value.</summary>
        internal IReadOnlyDictionary<string, int> AutoFilterTop10Values { get; }

        /// <summary>Gets imported Top/Bottom AutoFilter criteria grouped by top or bottom direction.</summary>
        internal IReadOnlyDictionary<string, int> AutoFilterTop10Directions { get; }

        /// <summary>Gets imported Top/Bottom AutoFilter criteria grouped by item-count or percentage unit.</summary>
        internal IReadOnlyDictionary<string, int> AutoFilterTop10Units { get; }

        /// <summary>Gets imported worksheets grouped by decoded BoundSheet visibility state.</summary>
        internal IReadOnlyDictionary<string, int> WorksheetsByVisibility { get; }

        /// <summary>Gets whether the workbook CodeName record was present.</summary>
        internal IReadOnlyDictionary<string, int> WorkbookCodeNameStates { get; }

        /// <summary>Gets workbook CodeName values grouped by name.</summary>
        internal IReadOnlyDictionary<string, int> WorkbookCodeNames { get; }

        /// <summary>Gets decoded workbook option states from Backup and BookBool records.</summary>
        internal IReadOnlyDictionary<string, int> WorkbookOptionStates { get; }

        /// <summary>Gets decoded BuiltInFnGroupCount values grouped by observed function category count.</summary>
        internal IReadOnlyDictionary<string, int> WorkbookBuiltInFunctionGroupCounts { get; }

        /// <summary>Gets imported worksheets grouped by CodeName record presence.</summary>
        internal IReadOnlyDictionary<string, int> WorksheetCodeNameStates { get; }

        /// <summary>Gets worksheet CodeName values grouped by name.</summary>
        internal IReadOnlyDictionary<string, int> WorksheetCodeNames { get; }

        /// <summary>Gets the number of imported defined names.</summary>
        public int DefinedNameCount { get; }

        /// <summary>Gets the number of preserved external-reference records.</summary>
        public int ExternalReferenceCount { get; }

        /// <summary>Gets the number of external workbook sheet names declared by supporting links.</summary>
        internal int ExternalSheetNameCount { get; }

        /// <summary>Gets the number of external names declared by supporting links.</summary>
        internal int ExternalNameCount { get; }

        /// <summary>Gets the number of preserved external cell cache sections.</summary>
        internal int ExternalCellCacheCount { get; }

        /// <summary>Gets the number of preserved cached external cell values.</summary>
        internal int ExternalCachedCellCount { get; }

        /// <summary>Gets the number of preserve-only DBQueryExt external query connection records decoded during import.</summary>
        internal int ExternalQueryConnectionCount { get; }

        /// <summary>Gets the number of preserve-only DConRef source range records decoded during import.</summary>
        internal int DataConsolidationReferenceCount { get; }

        /// <summary>Gets the number of DConName named consolidation source records decoded during import.</summary>
        internal int DataConsolidationNameCount { get; }

        /// <summary>Gets the number of preserve-only PivotTable BIFF records discovered during import.</summary>
        public int PivotTableRecordCount { get; }

        /// <summary>Gets the number of preserve-only chart BIFF records discovered during import.</summary>
        public int ChartRecordCount { get; }

        /// <summary>Gets the number of metadata records parsed from chart-sheet substreams.</summary>
        internal int ChartSheetMetadataRecordCount { get; }

        /// <summary>Gets the number of future metadata records parsed from chart-sheet substreams.</summary>
        internal int ChartSheetFutureMetadataRecordCount { get; }

        /// <summary>Gets the number of preserve-only drawing and object BIFF records discovered during import.</summary>
        public int DrawingRecordCount { get; }

        /// <summary>Gets the number of workbook Theme records discovered during import.</summary>
        internal int ThemeRecordCount { get; }

        /// <summary>Gets the number of OfficeArt record headers discovered under preserve-only drawing records.</summary>
        internal int DrawingOfficeArtRecordCount { get; }

        /// <summary>Gets the number of OfficeArtFDGGBlock records decoded under preserve-only drawing records.</summary>
        internal int DrawingGroupBlockCount { get; }

        /// <summary>Gets the number of OfficeArtFDG records decoded under preserve-only drawing records.</summary>
        internal int DrawingGroupInfoCount { get; }

        /// <summary>Gets the number of OfficeArtIDCL clusters decoded under OfficeArtFDGGBlock records.</summary>
        internal int DrawingIdentifierClusterCount { get; }

        /// <summary>Gets the number of OfficeArtFOPT shape property entries discovered under preserve-only drawing records.</summary>
        internal int DrawingShapePropertyCount { get; }

        /// <summary>Gets the number of parsed differential formats discovered during import.</summary>
        internal int DifferentialFormatCount { get; }

        /// <summary>Gets parsed differential formats grouped by source BIFF record type.</summary>
        internal IReadOnlyDictionary<string, int> DifferentialFormatsByRecordType { get; }

        /// <summary>Gets parsed differential formats grouped by decoded content state.</summary>
        internal IReadOnlyDictionary<string, int> DifferentialFormatsByContentState { get; }

        /// <summary>Gets parsed differential formats grouped by decoded fill shape.</summary>
        internal IReadOnlyDictionary<string, int> DifferentialFormatsByFill { get; }

        /// <summary>Gets parsed differential formats grouped by decoded font shape.</summary>
        internal IReadOnlyDictionary<string, int> DifferentialFormatsByFont { get; }

        /// <summary>Gets parsed differential formats grouped by decoded border shape.</summary>
        internal IReadOnlyDictionary<string, int> DifferentialFormatsByBorder { get; }

        /// <summary>Gets parsed differential formats grouped by decoded number format shape.</summary>
        internal IReadOnlyDictionary<string, int> DifferentialFormatsByNumberFormat { get; }

        /// <summary>Gets the number of parsed TableStyles collection records.</summary>
        internal int TableStyleCollectionRecordCount { get; }

        /// <summary>Gets the number of parsed user-defined TableStyle records.</summary>
        internal int TableStyleDefinitionCount { get; }

        /// <summary>Gets the number of parsed TableStyleElement records.</summary>
        internal int TableStyleElementRecordCount { get; }

        /// <summary>Gets table style collection records grouped by default table style name.</summary>
        internal IReadOnlyDictionary<string, int> TableStyleCollectionsByDefaultTableStyle { get; }

        /// <summary>Gets table style collection records grouped by default PivotTable style name.</summary>
        internal IReadOnlyDictionary<string, int> TableStyleCollectionsByDefaultPivotStyle { get; }

        /// <summary>Gets table style collection records grouped by declared total style count.</summary>
        internal IReadOnlyDictionary<string, int> TableStyleCollectionsByTotalStyleCount { get; }

        /// <summary>Gets user-defined table styles grouped by style name.</summary>
        internal IReadOnlyDictionary<string, int> TableStylesByName { get; }

        /// <summary>Gets user-defined table styles grouped by table and PivotTable applicability.</summary>
        internal IReadOnlyDictionary<string, int> TableStylesByApplicability { get; }

        /// <summary>Gets user-defined table styles grouped by declared element count.</summary>
        internal IReadOnlyDictionary<string, int> TableStylesByDeclaredElementCount { get; }

        /// <summary>Gets user-defined table styles grouped by parsed element count.</summary>
        internal IReadOnlyDictionary<string, int> TableStylesByParsedElementCount { get; }

        /// <summary>Gets table style elements grouped by element type.</summary>
        internal IReadOnlyDictionary<string, int> TableStyleElementsByType { get; }

        /// <summary>Gets table style elements grouped by referenced differential format index.</summary>
        internal IReadOnlyDictionary<string, int> TableStyleElementsByDifferentialFormatIndex { get; }

        /// <summary>Gets stripe table style elements grouped by stripe size.</summary>
        internal IReadOnlyDictionary<string, int> TableStyleElementsByStripeSize { get; }

        /// <summary>Gets the number of preserve-only compound container features discovered during import.</summary>
        internal int CompoundFeatureRecordCount { get; }

        /// <summary>Gets the number of matching compound directory entries behind preserve-only compound features.</summary>
        internal int CompoundFeatureEntryCount { get; }

        /// <summary>Gets the number of VBA module streams discovered in preserve-only compound features.</summary>
        internal int CompoundVbaModuleCount { get; }

        /// <summary>Gets the total declared byte size of matching preserve-only compound entries with known sizes.</summary>
        internal long CompoundFeatureEntryByteCount { get; }

        /// <summary>Gets the total declared byte size of discovered VBA module streams with known sizes.</summary>
        internal long CompoundVbaModuleByteCount { get; }

        /// <summary>Gets the number of calculation setting records parsed from BIFF records.</summary>
        internal int CalculationSettingRecordCount { get; }

        /// <summary>Gets the number of workbook cell style records parsed from Style records.</summary>
        internal int CellStyleRecordCount { get; }

        /// <summary>Gets the number of preserve-only style extension records parsed from XFExt records.</summary>
        internal int CellStyleExtensionRecordCount { get; }

        /// <summary>Gets the number of parsed-formula token observations captured during import.</summary>
        internal int FormulaTokenRecordCount { get; }

        /// <summary>Gets the number of Excel future-function aliases discovered from defined-name records.</summary>
        internal int FutureFunctionAliasCount { get; }

        /// <summary>Gets the number of workbook metadata records parsed from BIFF records.</summary>
        internal int WorkbookMetadataRecordCount { get; }

        /// <summary>Gets the number of preserve-only workbook future metadata records parsed from BIFF records.</summary>
        internal int WorkbookFutureMetadataRecordCount { get; }

        /// <summary>Gets the number of worksheet metadata records parsed from BIFF records.</summary>
        internal int WorksheetMetadataRecordCount { get; }

        /// <summary>Gets the number of preserve-only worksheet future metadata records parsed from BIFF records.</summary>
        internal int WorksheetFutureMetadataRecordCount { get; }

        /// <summary>Gets the number of metadata records parsed from unsupported sheet substreams.</summary>
        internal int UnsupportedSheetMetadataRecordCount { get; }

        /// <summary>Gets the number of preserve-only future metadata records parsed from unsupported sheet substreams.</summary>
        internal int UnsupportedSheetFutureMetadataRecordCount { get; }

        /// <summary>Gets the number of unsupported or preserve-only feature findings.</summary>
        public int UnsupportedFeatureCount { get; }

        /// <summary>Gets the number of unsupported features after subtracting records that are already preserve-modeled.</summary>
        public int UnsupportedProjectionGapCount { get; }

        /// <summary>Gets the number of preserve-only BIFF feature records with typed metadata.</summary>
        public int PreservedFeatureRecordCount { get; }

        /// <summary>Gets the number of error diagnostics produced during import.</summary>
        public int ErrorCount { get; }

        /// <summary>Gets the number of warning diagnostics produced during import.</summary>
        public int WarningCount { get; }

        /// <summary>Gets diagnostic counts grouped by stable diagnostic code.</summary>
        internal IReadOnlyDictionary<string, int> DiagnosticsByCode { get; }

        /// <summary>Gets unsupported formula token blockers grouped by stable detail key.</summary>
        internal IReadOnlyDictionary<string, int> FormulaTokenBlockers { get; }

        /// <summary>Gets unsupported formula token blockers grouped by raw formula token byte.</summary>
        internal IReadOnlyDictionary<string, int> FormulaTokenBlockersByToken { get; }

        /// <summary>Gets unsupported formula token blockers grouped by BIFF parsed-formula token name.</summary>
        internal IReadOnlyDictionary<string, int> FormulaTokenBlockersByTokenName { get; }

        /// <summary>Gets unsupported formula token blockers grouped by zero-based parsed-expression token offset.</summary>
        internal IReadOnlyDictionary<string, int> FormulaTokenBlockersByOffset { get; }

        /// <summary>Gets unsupported formula token blockers grouped by worksheet name.</summary>
        internal IReadOnlyDictionary<string, int> FormulaTokenBlockersBySheet { get; }

        /// <summary>Gets unsupported formula token blockers grouped by formula source context.</summary>
        internal IReadOnlyDictionary<string, int> FormulaTokenBlockersByContext { get; }

        /// <summary>Gets unsupported formula token blockers grouped by formula source context and raw formula token byte.</summary>
        internal IReadOnlyDictionary<string, int> FormulaTokenBlockersByContextAndToken { get; }

        /// <summary>Gets unsupported formula token blockers grouped by formula source context and BIFF parsed-formula token name.</summary>
        internal IReadOnlyDictionary<string, int> FormulaTokenBlockersByContextAndTokenName { get; }

        /// <summary>Gets unsupported formula token blockers grouped by formula source context and stable detail key.</summary>
        internal IReadOnlyDictionary<string, int> FormulaTokenBlockersByContextAndDetail { get; }

        /// <summary>Gets observed parsed-formula tokens grouped by BIFF token name.</summary>
        internal IReadOnlyDictionary<string, int> FormulaTokensByName { get; }

        /// <summary>Gets observed parsed-formula tokens grouped by formula source context.</summary>
        internal IReadOnlyDictionary<string, int> FormulaTokensByContext { get; }

        /// <summary>Gets observed parsed-formula tokens grouped by worksheet name.</summary>
        internal IReadOnlyDictionary<string, int> FormulaTokensBySheet { get; }

        /// <summary>Gets observed parsed-formula tokens grouped by formula source context and worksheet name.</summary>
        internal IReadOnlyDictionary<string, int> FormulaTokensByContextAndSheet { get; }

        /// <summary>Gets observed parsed-formula tokens grouped by formula source context and decoded operand category.</summary>
        internal IReadOnlyDictionary<string, int> FormulaTokensByContextAndOperandKind { get; }

        /// <summary>Gets observed parsed-formula tokens grouped by BIFF record type and token name.</summary>
        internal IReadOnlyDictionary<string, int> FormulaTokensByRecordType { get; }

        /// <summary>Gets observed parsed-formula tokens grouped by BIFF token class.</summary>
        internal IReadOnlyDictionary<string, int> FormulaTokensByClass { get; }

        /// <summary>Gets observed parsed-formula tokens grouped by token name and BIFF token class.</summary>
        internal IReadOnlyDictionary<string, int> FormulaTokensByNameAndClass { get; }

        /// <summary>Gets observed parsed-formula tokens grouped by token name and operand byte count.</summary>
        internal IReadOnlyDictionary<string, int> FormulaTokensByOperandByteCount { get; }

        /// <summary>Gets observed parsed-formula tokens grouped by decoded operand category.</summary>
        internal IReadOnlyDictionary<string, int> FormulaTokensByOperandKind { get; }

        /// <summary>Gets observed parsed-formula tokens grouped by token name and decoded operand category.</summary>
        internal IReadOnlyDictionary<string, int> FormulaTokensByNameAndOperandKind { get; }

        /// <summary>Gets observed parsed-formula tokens grouped by decoded operand category and operand text.</summary>
        internal IReadOnlyDictionary<string, int> FormulaTokensByOperandKindAndText { get; }

        /// <summary>Gets observed parsed-formula tokens grouped by token name and decoded operand text.</summary>
        internal IReadOnlyDictionary<string, int> FormulaTokensByNameAndOperandText { get; }

        /// <summary>Gets observed parsed-formula tokens grouped by sequence index within each expression.</summary>
        internal IReadOnlyDictionary<string, int> FormulaTokensBySequenceIndex { get; }

        /// <summary>Gets observed built-in formula function tokens grouped by raw function id.</summary>
        internal IReadOnlyDictionary<string, int> FormulaFunctionsById { get; }

        /// <summary>Gets observed built-in formula function tokens grouped by function name when known.</summary>
        internal IReadOnlyDictionary<string, int> FormulaFunctionsByName { get; }

        /// <summary>Gets observed formula function tokens grouped by function name and argument count.</summary>
        internal IReadOnlyDictionary<string, int> FormulaFunctionsByParameterCount { get; }

        /// <summary>Gets observed variable formula function tokens grouped by built-in versus CETAB state.</summary>
        internal IReadOnlyDictionary<string, int> FormulaFunctionsByCetabState { get; }

        /// <summary>Gets observed PtgAttr formula tokens grouped by attribute name.</summary>
        internal IReadOnlyDictionary<string, int> FormulaAttributesByName { get; }

        /// <summary>Gets Excel future-function aliases grouped by defined-name text.</summary>
        internal IReadOnlyDictionary<string, int> FutureFunctionAliasesByName { get; }

        /// <summary>Gets Excel future-function aliases grouped by function name without the _xlfn. prefix.</summary>
        internal IReadOnlyDictionary<string, int> FutureFunctionAliasesByFunction { get; }

        /// <summary>Gets Excel future-function aliases grouped by BIFF parsed-formula token name.</summary>
        internal IReadOnlyDictionary<string, int> FutureFunctionAliasesByTokenName { get; }

        /// <summary>Gets unsupported/preserve-only feature counts grouped by stable feature code.</summary>
        internal IReadOnlyDictionary<string, int> UnsupportedFeaturesByCode { get; }

        /// <summary>Gets unsupported/preserve-only feature counts grouped by feature kind.</summary>
        internal IReadOnlyDictionary<LegacyXlsUnsupportedFeatureKind, int> UnsupportedFeaturesByKind { get; }

        /// <summary>Gets unsupported/preserve-only feature counts grouped by kind, code, and BIFF record type.</summary>
        internal IReadOnlyDictionary<string, int> UnsupportedFeaturesByRecordType { get; }

        /// <summary>Gets unsupported/preserve-only feature counts grouped by kind, code, and stable feature subtype.</summary>
        internal IReadOnlyDictionary<string, int> UnsupportedFeaturesByDetail { get; }

        /// <summary>Gets unsupported/preserve-only feature counts grouped by code and workbook or sheet location.</summary>
        internal IReadOnlyDictionary<string, int> UnsupportedFeaturesByLocation { get; }

        /// <summary>Gets unsupported feature counts after subtracting records that are already preserve-modeled.</summary>
        internal IReadOnlyDictionary<LegacyXlsUnsupportedFeatureKind, int> UnsupportedProjectionGapsByKind { get; }

        /// <summary>Gets unsupported feature counts after subtracting preserve-modeled records, grouped by kind, code, and BIFF record type.</summary>
        internal IReadOnlyDictionary<string, int> UnsupportedProjectionGapsByRecordType { get; }

        /// <summary>Gets unsupported feature counts after subtracting preserve-modeled records, grouped by kind, code, and stable feature subtype.</summary>
        internal IReadOnlyDictionary<string, int> UnsupportedProjectionGapsByDetail { get; }

        /// <summary>Gets high-level workbook file-format states suitable for preflight and corpus comparison.</summary>
        internal IReadOnlyDictionary<string, int> FileFormatStates { get; }

        /// <summary>Gets hard file-format blockers grouped by kind and detail.</summary>
        internal IReadOnlyDictionary<string, int> FileFormatBlockers { get; }

        /// <summary>Gets hard file-format blockers grouped by kind and BIFF record type.</summary>
        internal IReadOnlyDictionary<string, int> FileFormatBlockersByRecordType { get; }

        /// <summary>Gets hard file-format blockers grouped by kind and BIFF record name.</summary>
        internal IReadOnlyDictionary<string, int> FileFormatBlockersByRecordName { get; }

        /// <summary>Gets hard file-format blockers grouped by code and workbook or sheet location.</summary>
        internal IReadOnlyDictionary<string, int> FileFormatBlockersByLocation { get; }

        /// <summary>Gets encrypted workbook blockers grouped by FilePass encryption method.</summary>
        internal IReadOnlyDictionary<string, int> EncryptedWorkbooksByMethod { get; }

        /// <summary>Gets unsupported BIFF blockers grouped by BIFF version.</summary>
        internal IReadOnlyDictionary<string, int> UnsupportedBiffVersionsByVersion { get; }

        /// <summary>Gets unsupported BIFF blockers grouped by BOF substream.</summary>
        internal IReadOnlyDictionary<string, int> UnsupportedBiffVersionsBySubstream { get; }

        /// <summary>Gets unsupported BIFF blockers grouped by BIFF version and BOF substream.</summary>
        internal IReadOnlyDictionary<string, int> UnsupportedBiffVersionsByVersionAndSubstream { get; }

        /// <summary>Gets unsupported sheet entries grouped by decoded sheet kind.</summary>
        internal IReadOnlyDictionary<LegacyXlsUnsupportedSheetKind, int> UnsupportedSheetsByKind { get; }

        /// <summary>Gets unsupported sheet entries grouped by raw BoundSheet type and decoded kind.</summary>
        internal IReadOnlyDictionary<string, int> UnsupportedSheetsByType { get; }

        /// <summary>Gets unsupported sheet entries grouped by sheet name.</summary>
        internal IReadOnlyDictionary<string, int> UnsupportedSheetsByName { get; }

        /// <summary>Gets unsupported sheet entries grouped by decoded BoundSheet visibility state.</summary>
        internal IReadOnlyDictionary<string, int> UnsupportedSheetsByVisibility { get; }

        /// <summary>Gets unsupported sheet entries grouped by sheet kind and decoded visibility state.</summary>
        internal IReadOnlyDictionary<string, int> UnsupportedSheetsByKindAndVisibility { get; }

        /// <summary>Gets decoded chart sheets grouped by raw BoundSheet type.</summary>
        internal IReadOnlyDictionary<string, int> ChartSheetsByType { get; }

        /// <summary>Gets decoded chart sheets grouped by sheet name.</summary>
        internal IReadOnlyDictionary<string, int> ChartSheetsByName { get; }

        /// <summary>Gets decoded chart sheets grouped by visibility state.</summary>
        internal IReadOnlyDictionary<string, int> ChartSheetsByVisibility { get; }

        /// <summary>Gets chart-sheet metadata records grouped by decoded kind.</summary>
        internal IReadOnlyDictionary<string, int> ChartSheetMetadataRecordsByKind { get; }

        /// <summary>Gets chart-sheet future metadata records grouped by BIFF record type.</summary>
        internal IReadOnlyDictionary<string, int> ChartSheetFutureMetadataRecordsByRecordType { get; }

        /// <summary>Gets decoded chart sheets grouped by raw PrintSize value.</summary>
        internal IReadOnlyDictionary<string, int> ChartSheetPrintSizes { get; }

        /// <summary>Gets decoded chart sheets grouped by PrintSize mode name.</summary>
        internal IReadOnlyDictionary<string, int> ChartSheetPrintSizeKinds { get; }

        /// <summary>Gets decoded chart sheets grouped by chart text object count.</summary>
        internal IReadOnlyDictionary<string, int> ChartSheetTextObjectCounts { get; }

        /// <summary>Gets decoded chart sheets grouped by chart record count.</summary>
        internal IReadOnlyDictionary<string, int> ChartSheetChartRecordCounts { get; }

        /// <summary>Gets decoded chart sheets grouped by sheet name and chart record count.</summary>
        internal IReadOnlyDictionary<string, int> ChartSheetChartRecordCountsBySheet { get; }

        /// <summary>Gets chart-sheet chart records grouped by shallow category.</summary>
        internal IReadOnlyDictionary<string, int> ChartSheetChartRecordKinds { get; }

        /// <summary>Gets chart-sheet chart records grouped by sheet name and shallow category.</summary>
        internal IReadOnlyDictionary<string, int> ChartSheetChartRecordKindsBySheet { get; }

        /// <summary>Gets chart-sheet chart type records grouped by decoded chart family.</summary>
        internal IReadOnlyDictionary<string, int> ChartSheetChartTypes { get; }

        /// <summary>Gets chart-sheet chart type records grouped by sheet name and decoded chart family.</summary>
        internal IReadOnlyDictionary<string, int> ChartSheetChartTypesBySheet { get; }

        /// <summary>Gets decoded chart sheets grouped by metadata shape.</summary>
        internal IReadOnlyDictionary<string, int> ChartSheetStates { get; }

        /// <summary>Gets unsupported chart sheets grouped by raw PrintSize value.</summary>
        internal IReadOnlyDictionary<string, int> UnsupportedChartSheetPrintSizes { get; }

        /// <summary>Gets unsupported chart sheets grouped by decoded PrintSize mode name.</summary>
        internal IReadOnlyDictionary<string, int> UnsupportedChartSheetPrintSizeKinds { get; }

        /// <summary>Gets unsupported chart sheets grouped by chart text object count.</summary>
        internal IReadOnlyDictionary<string, int> UnsupportedChartSheetTextObjectCounts { get; }

        /// <summary>Gets unsupported chart sheets grouped by preserve-only chart record count.</summary>
        internal IReadOnlyDictionary<string, int> UnsupportedChartSheetChartRecordCounts { get; }

        /// <summary>Gets unsupported chart sheets grouped by sheet name and preserve-only chart record count.</summary>
        internal IReadOnlyDictionary<string, int> UnsupportedChartSheetChartRecordCountsBySheet { get; }

        /// <summary>Gets unsupported chart sheet preserve-only chart records grouped by shallow category.</summary>
        internal IReadOnlyDictionary<string, int> UnsupportedChartSheetChartRecordKinds { get; }

        /// <summary>Gets unsupported chart sheet preserve-only chart records grouped by sheet name and shallow category.</summary>
        internal IReadOnlyDictionary<string, int> UnsupportedChartSheetChartRecordKindsBySheet { get; }

        /// <summary>Gets unsupported chart sheet preserve-only chart type records grouped by decoded chart family.</summary>
        internal IReadOnlyDictionary<string, int> UnsupportedChartSheetChartTypes { get; }

        /// <summary>Gets unsupported chart sheet preserve-only chart type records grouped by sheet name and decoded chart family.</summary>
        internal IReadOnlyDictionary<string, int> UnsupportedChartSheetChartTypesBySheet { get; }

        /// <summary>Gets unsupported chart sheets grouped by preserve-only chart metadata shape.</summary>
        internal IReadOnlyDictionary<string, int> UnsupportedChartSheetStates { get; }

        /// <summary>Gets preserved external references grouped by supporting-link kind.</summary>
        internal IReadOnlyDictionary<LegacyXlsExternalReferenceKind, int> ExternalReferencesByKind { get; }

        /// <summary>Gets preserved external references grouped by target path or source.</summary>
        internal IReadOnlyDictionary<string, int> ExternalReferencesByTarget { get; }

        /// <summary>Gets preserved external references grouped by their sheet/name/cache/cached-cell shape.</summary>
        internal IReadOnlyDictionary<string, int> ExternalReferencesByShape { get; }

        /// <summary>Gets workbook-level external-reference model-shape states derived from supporting-link metadata.</summary>
        internal IReadOnlyDictionary<string, int> ExternalReferenceWorkbookStates { get; }

        /// <summary>Gets preserved external references grouped by declared SupBook sheet count.</summary>
        internal IReadOnlyDictionary<string, int> ExternalReferencesByDeclaredSheetCount { get; }

        /// <summary>Gets preserved external references grouped by sheet-name count.</summary>
        internal IReadOnlyDictionary<string, int> ExternalReferencesBySheetNameCount { get; }

        /// <summary>Gets preserved external references grouped by parsed sheet-name table completeness.</summary>
        internal IReadOnlyDictionary<string, int> ExternalReferencesBySheetTableState { get; }

        /// <summary>Gets preserved external references grouped by external-name count.</summary>
        internal IReadOnlyDictionary<string, int> ExternalReferencesByExternalNameCount { get; }

        /// <summary>Gets preserved external references grouped by cached cell section count.</summary>
        internal IReadOnlyDictionary<string, int> ExternalReferencesByCacheCount { get; }

        /// <summary>Gets preserved external references grouped by cached cell value count.</summary>
        internal IReadOnlyDictionary<string, int> ExternalReferencesByCachedCellCount { get; }

        /// <summary>Gets external workbook sheet-name counts grouped by supporting-link kind.</summary>
        internal IReadOnlyDictionary<LegacyXlsExternalReferenceKind, int> ExternalSheetNamesByReferenceKind { get; }

        /// <summary>Gets external workbook sheet names grouped by normalized target and sheet name.</summary>
        internal IReadOnlyDictionary<string, int> ExternalSheetNamesByTarget { get; }

        /// <summary>Gets external defined-name counts grouped by supporting-link kind.</summary>
        internal IReadOnlyDictionary<LegacyXlsExternalReferenceKind, int> ExternalNamesByReferenceKind { get; }

        /// <summary>Gets external defined names grouped by name text.</summary>
        internal IReadOnlyDictionary<string, int> ExternalNamesByName { get; }

        /// <summary>Gets external defined names grouped by workbook or sheet-local scope.</summary>
        internal IReadOnlyDictionary<string, int> ExternalNamesByScope { get; }

        /// <summary>Gets external defined names grouped by built-in or custom state.</summary>
        internal IReadOnlyDictionary<string, int> ExternalNamesByBuiltInState { get; }

        /// <summary>Gets external defined names grouped by decoded ExternName body kind.</summary>
        internal IReadOnlyDictionary<string, int> ExternalNamesByBodyKind { get; }

        /// <summary>Gets external defined names grouped by decoded cached clipboard format.</summary>
        internal IReadOnlyDictionary<string, int> ExternalNamesByCachedClipboardFormat { get; }

        /// <summary>Gets external defined names grouped by ExternName advise-update flag state.</summary>
        internal IReadOnlyDictionary<string, int> ExternalNamesByAdviseState { get; }

        /// <summary>Gets external defined names grouped by ExternName picture-format flag state.</summary>
        internal IReadOnlyDictionary<string, int> ExternalNamesByPictureState { get; }

        /// <summary>Gets external defined names grouped by ExternName OLE flag state.</summary>
        internal IReadOnlyDictionary<string, int> ExternalNamesByOleState { get; }

        /// <summary>Gets external defined names grouped by ExternName OLE-link flag state.</summary>
        internal IReadOnlyDictionary<string, int> ExternalNamesByOleLinkState { get; }

        /// <summary>Gets external defined names grouped by ExternName icon-display flag state.</summary>
        internal IReadOnlyDictionary<string, int> ExternalNamesByIconState { get; }

        /// <summary>Gets external defined names grouped by decoded ExternName flag shape.</summary>
        internal IReadOnlyDictionary<string, int> ExternalNamesByFlagShape { get; }

        /// <summary>Gets external cell cache sections grouped by normalized target path or source.</summary>
        internal IReadOnlyDictionary<string, int> ExternalCellCachesByTarget { get; }

        /// <summary>Gets external cell cache sections grouped by resolved external sheet name.</summary>
        internal IReadOnlyDictionary<string, int> ExternalCellCachesBySheetName { get; }

        /// <summary>Gets external cell cache sections grouped by normalized target and resolved external sheet name.</summary>
        internal IReadOnlyDictionary<string, int> ExternalCellCachesByTargetAndSheetName { get; }

        /// <summary>Gets external cell cache sections grouped by occupied zero-based row/column range.</summary>
        internal IReadOnlyDictionary<string, int> ExternalCellCachesByCellRange { get; }

        /// <summary>Gets external cell cache sections grouped by normalized target and occupied zero-based row/column range.</summary>
        internal IReadOnlyDictionary<string, int> ExternalCellCachesByTargetAndCellRange { get; }

        /// <summary>Gets external cell cache sections grouped by cached value count.</summary>
        internal IReadOnlyDictionary<string, int> ExternalCellCachesByCellCount { get; }

        /// <summary>Gets external cell cache sections grouped by occupied row span.</summary>
        internal IReadOnlyDictionary<string, int> ExternalCellCachesByRowSpan { get; }

        /// <summary>Gets external cell cache sections grouped by occupied column span.</summary>
        internal IReadOnlyDictionary<string, int> ExternalCellCachesByColumnSpan { get; }

        /// <summary>Gets external cell cache sections grouped by XCT link-valid state.</summary>
        internal IReadOnlyDictionary<string, int> ExternalCellCachesByLinkState { get; }

        /// <summary>Gets cached external cell values grouped by value kind.</summary>
        internal IReadOnlyDictionary<LegacyXlsCellValueKind, int> ExternalCachedCellsByValueKind { get; }

        /// <summary>Gets cached external cell values grouped by normalized target, resolved external sheet name, and value kind.</summary>
        internal IReadOnlyDictionary<string, int> ExternalCachedCellsByTargetSheetAndValueKind { get; }

        /// <summary>Gets DBQueryExt records grouped by decoded data source type.</summary>
        internal IReadOnlyDictionary<string, int> ExternalQueryConnectionsBySourceType { get; }

        /// <summary>Gets DBQueryExt records grouped by shallow connection state.</summary>
        internal IReadOnlyDictionary<string, int> ExternalQueryConnectionsByState { get; }

        /// <summary>Gets DBQueryExt records grouped by enabled connection flag.</summary>
        internal IReadOnlyDictionary<string, int> ExternalQueryConnectionsByConnectionFlag { get; }

        /// <summary>Gets DBQueryExt records grouped by enabled query option flag.</summary>
        internal IReadOnlyDictionary<string, int> ExternalQueryConnectionsByQueryOption { get; }

        /// <summary>Gets DBQueryExt records grouped by declared PBT parameter flag count.</summary>
        internal IReadOnlyDictionary<string, int> ExternalQueryConnectionsByParameterFlagCount { get; }

        /// <summary>Gets DBQueryExt records grouped by decoded PBT parameter flag byte count.</summary>
        internal IReadOnlyDictionary<string, int> ExternalQueryConnectionsByParameterFlagByteCount { get; }

        /// <summary>Gets DBQueryExt records grouped by whether decoded PBT bytes match the declared count.</summary>
        internal IReadOnlyDictionary<string, int> ExternalQueryConnectionsByParameterFlagState { get; }

        /// <summary>Gets DBQueryExt records grouped by declared future-byte count.</summary>
        internal IReadOnlyDictionary<string, int> ExternalQueryConnectionsByFutureByteCount { get; }

        /// <summary>Gets DBQueryExt records grouped by automatic refresh interval.</summary>
        internal IReadOnlyDictionary<string, int> ExternalQueryConnectionsByRefreshInterval { get; }

        /// <summary>Gets DBQueryExt records grouped by declared OleDbConn follow-up record count.</summary>
        internal IReadOnlyDictionary<string, int> ExternalQueryConnectionsByOleDbConnectionCount { get; }

        /// <summary>Gets DBQueryExt records grouped by Web query HTML formatting mode.</summary>
        internal IReadOnlyDictionary<string, int> ExternalQueryConnectionsByHtmlFormat { get; }

        /// <summary>Gets DBQueryExt records grouped by data functionality version triplet.</summary>
        internal IReadOnlyDictionary<string, int> ExternalQueryConnectionsByVersionTriplet { get; }

        /// <summary>Gets DBQueryExt records grouped by raw source-specific ConnGrbitDbt flags.</summary>
        internal IReadOnlyDictionary<string, int> ExternalQueryConnectionsBySourceSpecificFlags { get; }

        /// <summary>Gets DConRef records grouped by decoded DConFile source kind.</summary>
        internal IReadOnlyDictionary<string, int> DataConsolidationReferencesBySourceKind { get; }

        /// <summary>Gets DConRef records grouped by raw DConFile source prefix.</summary>
        internal IReadOnlyDictionary<string, int> DataConsolidationReferencesBySourcePrefix { get; }

        /// <summary>Gets DConRef records grouped by decoded source path or sheet name.</summary>
        internal IReadOnlyDictionary<string, int> DataConsolidationReferencesBySource { get; }

        /// <summary>Gets DConRef records grouped by decoded source range.</summary>
        internal IReadOnlyDictionary<string, int> DataConsolidationReferencesByRange { get; }

        /// <summary>Gets DConRef records grouped by decoded source range shape.</summary>
        internal IReadOnlyDictionary<string, int> DataConsolidationReferencesByShape { get; }

        /// <summary>Gets DConRef records grouped by decoded source and range.</summary>
        internal IReadOnlyDictionary<string, int> DataConsolidationReferencesBySourceAndRange { get; }

        /// <summary>Gets DConRef records grouped by trailing unused byte count.</summary>
        internal IReadOnlyDictionary<string, int> DataConsolidationReferencesByUnusedByteCount { get; }

        /// <summary>Gets DConName records grouped by decoded source kind.</summary>
        internal IReadOnlyDictionary<string, int> DataConsolidationNamesBySourceKind { get; }

        /// <summary>Gets DConName records grouped by defined-name reference.</summary>
        internal IReadOnlyDictionary<string, int> DataConsolidationNamesByName { get; }

        /// <summary>Gets DConName records grouped by external source, or self-reference marker.</summary>
        internal IReadOnlyDictionary<string, int> DataConsolidationNamesBySource { get; }

        /// <summary>Gets DConName records grouped by defined name and source.</summary>
        internal IReadOnlyDictionary<string, int> DataConsolidationNamesByNameAndSource { get; }

        /// <summary>Gets DConName records grouped by trailing unused byte count.</summary>
        internal IReadOnlyDictionary<string, int> DataConsolidationNamesByUnusedByteCount { get; }

        /// <summary>Gets Theme records grouped by decoded theme version.</summary>
        internal IReadOnlyDictionary<string, int> ThemeRecordsByVersion { get; }

        /// <summary>Gets Theme records grouped by raw theme version value.</summary>
        internal IReadOnlyDictionary<string, int> ThemeRecordsByRawVersion { get; }

        /// <summary>Gets Theme records grouped by whether embedded theme content bytes were present.</summary>
        internal IReadOnlyDictionary<string, int> ThemeRecordsByContentState { get; }

        /// <summary>Gets Theme records grouped by embedded theme content byte length.</summary>
        internal IReadOnlyDictionary<string, int> ThemeRecordsByContentLength { get; }

        /// <summary>Gets preserve-only PivotTable BIFF records grouped by decoded metadata kind.</summary>
        internal IReadOnlyDictionary<LegacyXlsPivotTableRecordKind, int> PivotTableRecordsByKind { get; }

        /// <summary>Gets preserve-only PivotTable BIFF records grouped by record name.</summary>
        internal IReadOnlyDictionary<string, int> PivotTableRecordsByName { get; }

        /// <summary>Gets preserve-only PivotTable BIFF records grouped by workbook or worksheet location.</summary>
        internal IReadOnlyDictionary<string, int> PivotTableRecordsByLocation { get; }

        /// <summary>Gets preserve-only PivotTable BIFF records grouped by decoded metadata kind and workbook or worksheet location.</summary>
        internal IReadOnlyDictionary<string, int> PivotTableRecordsByKindAndLocation { get; }

        /// <summary>Gets preserve-only PivotTable BIFF records grouped by BIFF record name and workbook or worksheet location.</summary>
        internal IReadOnlyDictionary<string, int> PivotTableRecordsByNameAndLocation { get; }

        /// <summary>Gets workbook-level PivotTable model-shape states derived from preserve-only PivotTable BIFF records.</summary>
        internal IReadOnlyDictionary<string, int> PivotTableWorkbookStates { get; }

        /// <summary>Gets decoded SxView PivotTable views grouped by covered A1 range.</summary>
        internal IReadOnlyDictionary<string, int> PivotTableViewRanges { get; }

        /// <summary>Gets decoded SxView PivotTable views grouped by PivotTable name.</summary>
        internal IReadOnlyDictionary<string, int> PivotTableViewNames { get; }

        /// <summary>Gets decoded SxView PivotTable views grouped by data field caption.</summary>
        internal IReadOnlyDictionary<string, int> PivotTableViewDataNames { get; }

        /// <summary>Gets decoded SxView PivotTable views grouped by field counts.</summary>
        internal IReadOnlyDictionary<string, int> PivotTableViewFieldCounts { get; }

        /// <summary>Gets decoded SxView PivotTable views grouped by row and column line counts.</summary>
        internal IReadOnlyDictionary<string, int> PivotTableViewLineCounts { get; }

        /// <summary>Gets decoded SxView PivotTable views grouped by default data axis.</summary>
        internal IReadOnlyDictionary<string, int> PivotTableViewDataAxes { get; }

        /// <summary>Gets decoded SxView PivotTable views grouped by data field position.</summary>
        internal IReadOnlyDictionary<string, int> PivotTableViewDataPositions { get; }

        /// <summary>Gets decoded SxView PivotTable views grouped by PivotCache index.</summary>
        internal IReadOnlyDictionary<string, int> PivotTableViewCacheIndexes { get; }

        /// <summary>Gets decoded SxView PivotTable views grouped by row and column grand total state.</summary>
        internal IReadOnlyDictionary<string, int> PivotTableViewGrandTotalStates { get; }

        /// <summary>Gets decoded SxView PivotTable views grouped by AutoFormat state and identifier.</summary>
        internal IReadOnlyDictionary<string, int> PivotTableViewAutoFormatStates { get; }

        /// <summary>Gets decoded Sxvd PivotTable fields grouped by axis.</summary>
        internal IReadOnlyDictionary<string, int> PivotTableFieldAxes { get; }

        /// <summary>Gets decoded Sxvd PivotTable fields grouped by declared item count.</summary>
        internal IReadOnlyDictionary<string, int> PivotTableFieldItemCounts { get; }

        /// <summary>Gets decoded Sxvd PivotTable fields grouped by declared subtotal count and raw subtotal flags.</summary>
        internal IReadOnlyDictionary<string, int> PivotTableFieldSubtotalCounts { get; }

        /// <summary>Gets decoded Sxvd PivotTable fields grouped by subtotal function name.</summary>
        internal IReadOnlyDictionary<string, int> PivotTableFieldSubtotalFunctions { get; }

        /// <summary>Gets decoded Sxvd PivotTable fields grouped by explicit caption.</summary>
        internal IReadOnlyDictionary<string, int> PivotTableFieldNames { get; }

        /// <summary>Gets decoded SxIvd PivotTable field-index lists grouped by list length.</summary>
        internal IReadOnlyDictionary<string, int> PivotTableFieldIndexListLengths { get; }

        /// <summary>Gets decoded SxIvd PivotTable field-index lists grouped by referenced pivot field index.</summary>
        internal IReadOnlyDictionary<string, int> PivotTableFieldIndexReferences { get; }

        /// <summary>Gets decoded SxIvd PivotTable field-index lists grouped by full index sequence.</summary>
        internal IReadOnlyDictionary<string, int> PivotTableFieldIndexSequences { get; }

        /// <summary>Gets decoded SXLI PivotTable line item records grouped by line item count.</summary>
        internal IReadOnlyDictionary<string, int> PivotTableLineItemCounts { get; }

        /// <summary>Gets decoded SXLI PivotTable line items grouped by raw item type.</summary>
        internal IReadOnlyDictionary<string, int> PivotTableLineItemTypes { get; }

        /// <summary>Gets decoded SXLI PivotTable line items grouped by item type name.</summary>
        internal IReadOnlyDictionary<string, int> PivotTableLineItemTypeKinds { get; }

        /// <summary>Gets decoded SXLI PivotTable line items grouped by declared entry count.</summary>
        internal IReadOnlyDictionary<string, int> PivotTableLineItemEntryCounts { get; }

        /// <summary>Gets decoded SXLI PivotTable line items grouped by physical entry slots and declared entry count.</summary>
        internal IReadOnlyDictionary<string, int> PivotTableLineItemEntrySlotCounts { get; }

        /// <summary>Gets decoded SXLI PivotTable line items grouped by entry index name.</summary>
        internal IReadOnlyDictionary<string, int> PivotTableLineItemEntryIndexes { get; }

        /// <summary>Gets decoded SXLI PivotTable line items grouped by data item index.</summary>
        internal IReadOnlyDictionary<string, int> PivotTableLineItemDataIndexes { get; }

        /// <summary>Gets decoded SXLI PivotTable line items grouped by subtotal, block total, grand total, and data-axis flags.</summary>
        internal IReadOnlyDictionary<string, int> PivotTableLineItemFlagStates { get; }

        /// <summary>Gets decoded SXLI PivotTable line item records grouped by full line entry sequence.</summary>
        internal IReadOnlyDictionary<string, int> PivotTableLineItemSequences { get; }

        /// <summary>Gets decoded SXPI PivotTable page item selectors grouped by selector count.</summary>
        internal IReadOnlyDictionary<string, int> PivotTablePageItemCounts { get; }

        /// <summary>Gets decoded SXPI PivotTable page item selectors grouped by page-axis field index.</summary>
        internal IReadOnlyDictionary<string, int> PivotTablePageItemFieldIndexes { get; }

        /// <summary>Gets decoded SXPI PivotTable page item selectors grouped by selected item index.</summary>
        internal IReadOnlyDictionary<string, int> PivotTablePageItemIndexes { get; }

        /// <summary>Gets decoded SXPI PivotTable page item selectors grouped by drop-down object identifier.</summary>
        internal IReadOnlyDictionary<string, int> PivotTablePageItemObjectIds { get; }

        /// <summary>Gets decoded SXPI PivotTable page item selectors grouped by full selector sequence.</summary>
        internal IReadOnlyDictionary<string, int> PivotTablePageItemSequences { get; }

        /// <summary>Gets decoded SXVI PivotTable items grouped by raw item type.</summary>
        internal IReadOnlyDictionary<string, int> PivotTableItemTypes { get; }

        /// <summary>Gets decoded SXVI PivotTable items grouped by item type name.</summary>
        internal IReadOnlyDictionary<string, int> PivotTableItemTypeKinds { get; }

        /// <summary>Gets decoded SXVI PivotTable items grouped by referenced PivotCache item index.</summary>
        internal IReadOnlyDictionary<string, int> PivotTableItemCacheIndexes { get; }

        /// <summary>Gets decoded SXVI PivotTable items grouped by visibility and formula flags.</summary>
        internal IReadOnlyDictionary<string, int> PivotTableItemFlagStates { get; }

        /// <summary>Gets decoded SXVI PivotTable items grouped by explicit caption.</summary>
        internal IReadOnlyDictionary<string, int> PivotTableItemNames { get; }

        /// <summary>Gets PivotTable formula records grouped by BIFF record name and payload length.</summary>
        internal IReadOnlyDictionary<string, int> PivotTableFormulaPayloadLengths { get; }

        /// <summary>Gets PivotTable formula records grouped by decoded payload shape.</summary>
        internal IReadOnlyDictionary<string, int> PivotTableFormulaPayloadKinds { get; }

        /// <summary>Gets SXFormula token streams grouped by parsed-expression byte count.</summary>
        internal IReadOnlyDictionary<string, int> PivotTableFormulaTokenByteCounts { get; }

        /// <summary>Gets calculated-field SXFormula token streams grouped by parsed-expression byte count.</summary>
        internal IReadOnlyDictionary<string, int> PivotTableCalculatedFieldFormulaTokenByteCounts { get; }

        /// <summary>Gets SXFormula token streams grouped by trailing byte count after the parsed expression.</summary>
        internal IReadOnlyDictionary<string, int> PivotTableFormulaTrailingByteCounts { get; }

        /// <summary>Gets decoded SxRule records grouped by axis.</summary>
        internal IReadOnlyDictionary<string, int> PivotTableRuleAxes { get; }

        /// <summary>Gets decoded SxRule records grouped by rule area type.</summary>
        internal IReadOnlyDictionary<string, int> PivotTableRuleTypes { get; }

        /// <summary>Gets decoded SxRule records grouped by field reference.</summary>
        internal IReadOnlyDictionary<string, int> PivotTableRuleFieldReferences { get; }

        /// <summary>Gets decoded SxRule records grouped by following SxFilt count.</summary>
        internal IReadOnlyDictionary<string, int> PivotTableRuleFilterCounts { get; }

        /// <summary>Gets decoded SxRule records grouped by option flags.</summary>
        internal IReadOnlyDictionary<string, int> PivotTableRuleOptionStates { get; }

        /// <summary>Gets decoded partial-area SxRule records grouped by relative area bounds.</summary>
        internal IReadOnlyDictionary<string, int> PivotTableRulePartialAreas { get; }

        /// <summary>Gets decoded SxFilt records grouped by contained rule-filter entry count.</summary>
        internal IReadOnlyDictionary<string, int> PivotTableRuleFilterEntryCounts { get; }

        /// <summary>Gets decoded SxFilt rule-filter entries grouped by axis.</summary>
        internal IReadOnlyDictionary<string, int> PivotTableRuleFilterAxes { get; }

        /// <summary>Gets decoded SxFilt rule-filter entries grouped by field position.</summary>
        internal IReadOnlyDictionary<string, int> PivotTableRuleFilterFieldPositions { get; }

        /// <summary>Gets decoded SxFilt rule-filter entries grouped by referenced field.</summary>
        internal IReadOnlyDictionary<string, int> PivotTableRuleFilterFieldReferences { get; }

        /// <summary>Gets decoded SxFilt rule-filter entries grouped by selected-header state.</summary>
        internal IReadOnlyDictionary<string, int> PivotTableRuleFilterSelectedStates { get; }

        /// <summary>Gets decoded SxFilt rule-filter entries grouped by raw subtotal flags.</summary>
        internal IReadOnlyDictionary<string, int> PivotTableRuleFilterSubtotalFlags { get; }

        /// <summary>Gets decoded SxFilt rule-filter entries grouped by subtotal function.</summary>
        internal IReadOnlyDictionary<string, int> PivotTableRuleFilterSubtotalFunctions { get; }

        /// <summary>Gets decoded SxFilt rule-filter entries grouped by following SxItm index count.</summary>
        internal IReadOnlyDictionary<string, int> PivotTableRuleFilterItemIndexCounts { get; }

        /// <summary>Gets decoded SxFilt rule-filter entries grouped by compact filter state.</summary>
        internal IReadOnlyDictionary<string, int> PivotTableRuleFilterStates { get; }

        /// <summary>Gets decoded PivotCache item records grouped by value kind.</summary>
        internal IReadOnlyDictionary<string, int> PivotTableCacheItemKinds { get; }

        /// <summary>Gets decoded PivotCache item records grouped by empty or value-bearing state.</summary>
        internal IReadOnlyDictionary<string, int> PivotTableCacheItemValueStates { get; }

        /// <summary>Gets decoded PivotCache string item records grouped by character count.</summary>
        internal IReadOnlyDictionary<string, int> PivotTableCacheItemStringLengths { get; }

        /// <summary>Gets decoded PivotCache error item records grouped by error code.</summary>
        internal IReadOnlyDictionary<string, int> PivotTableCacheItemErrorCodes { get; }

        /// <summary>Gets decoded PivotCache Boolean item records grouped by value.</summary>
        internal IReadOnlyDictionary<string, int> PivotTableCacheItemBooleanValues { get; }

        /// <summary>Gets decoded PivotCache stream identifiers grouped by BIFF record name and stream name.</summary>
        internal IReadOnlyDictionary<string, int> PivotTableCacheStreamNames { get; }

        /// <summary>Gets decoded PivotCache source-data types grouped by BIFF record name and source type.</summary>
        internal IReadOnlyDictionary<string, int> PivotTableCacheSourceTypes { get; }

        /// <summary>Gets decoded SXDB PivotCache records grouped by declared cache record count.</summary>
        internal IReadOnlyDictionary<string, int> PivotTableCacheRecordCounts { get; }

        /// <summary>Gets decoded SXDB PivotCache records grouped by source and total field counts.</summary>
        internal IReadOnlyDictionary<string, int> PivotTableCacheFieldCounts { get; }

        /// <summary>Gets decoded SXDB PivotCache records grouped by used cache record count.</summary>
        internal IReadOnlyDictionary<string, int> PivotTableCacheUsedRecordCounts { get; }

        /// <summary>Gets decoded SXDB PivotCache records grouped by cache property flag state.</summary>
        internal IReadOnlyDictionary<string, int> PivotTableCachePropertyFlags { get; }

        /// <summary>Gets decoded SXDB PivotCache records grouped by whether the last-refresh user was present.</summary>
        internal IReadOnlyDictionary<string, int> PivotTableCacheRefreshUserStates { get; }

        /// <summary>Gets decoded QsiSXTag records grouped by target type.</summary>
        internal IReadOnlyDictionary<string, int> PivotTableQueryTagTargets { get; }

        /// <summary>Gets decoded QsiSXTag records grouped by query table or PivotTable view name.</summary>
        internal IReadOnlyDictionary<string, int> PivotTableQueryTagNames { get; }

        /// <summary>Gets decoded QsiSXTag records grouped by refresh/cache validity flags.</summary>
        internal IReadOnlyDictionary<string, int> PivotTableQueryTagRefreshStates { get; }

        /// <summary>Gets decoded QsiSXTag records grouped by functionality version pair.</summary>
        internal IReadOnlyDictionary<string, int> PivotTableQueryTagVersions { get; }

        /// <summary>Gets decoded QsiSXTag records grouped by raw future option flags.</summary>
        internal IReadOnlyDictionary<string, int> PivotTableQueryTagFutureOptions { get; }

        /// <summary>Gets decoded QsiSXTag records grouped by trailing unused field value.</summary>
        internal IReadOnlyDictionary<string, int> PivotTableQueryTagUnusedValues { get; }

        /// <summary>Gets decoded SXDI PivotTable data item records grouped by raw aggregation function identifier.</summary>
        internal IReadOnlyDictionary<string, int> PivotTableDataItemAggregations { get; }

        /// <summary>Gets decoded SXDI PivotTable data item records grouped by aggregation function name.</summary>
        internal IReadOnlyDictionary<string, int> PivotTableDataItemAggregationKinds { get; }

        /// <summary>Gets decoded SXDI PivotTable data item records grouped by pivot field index.</summary>
        internal IReadOnlyDictionary<string, int> PivotTableDataItemFieldIndexes { get; }

        /// <summary>Gets decoded SXDI PivotTable data item records grouped by raw display calculation identifier.</summary>
        internal IReadOnlyDictionary<string, int> PivotTableDataItemDisplayCalculationIds { get; }

        /// <summary>Gets decoded SXDI PivotTable data item records grouped by display calculation name.</summary>
        internal IReadOnlyDictionary<string, int> PivotTableDataItemDisplayCalculations { get; }

        /// <summary>Gets decoded SXDI PivotTable data item records grouped by display calculation and reference state.</summary>
        internal IReadOnlyDictionary<string, int> PivotTableDataItemDisplayCalculationReferenceStates { get; }

        /// <summary>Gets decoded SXDI PivotTable data item records grouped by display-calculation field index.</summary>
        internal IReadOnlyDictionary<string, int> PivotTableDataItemDisplayCalculationFieldIndexes { get; }

        /// <summary>Gets decoded SXDI PivotTable data item records grouped by display-calculation item index.</summary>
        internal IReadOnlyDictionary<string, int> PivotTableDataItemDisplayCalculationItemIndexes { get; }

        /// <summary>Gets decoded SXDI PivotTable data item records grouped by number format identifier.</summary>
        internal IReadOnlyDictionary<string, int> PivotTableDataItemNumberFormats { get; }

        /// <summary>Gets decoded SXDI PivotTable data item records grouped by custom data item name.</summary>
        internal IReadOnlyDictionary<string, int> PivotTableDataItemNames { get; }

        /// <summary>Gets decoded SXRng PivotTable grouping records grouped by grouping kind.</summary>
        internal IReadOnlyDictionary<string, int> PivotTableGroupingKinds { get; }

        /// <summary>Gets decoded SXRng PivotTable grouping records grouped by automatic boundary state.</summary>
        internal IReadOnlyDictionary<string, int> PivotTableGroupingBoundaryStates { get; }

        /// <summary>Gets decoded SXRng PivotTable grouping records grouped by whether the expected companion values were present.</summary>
        internal IReadOnlyDictionary<string, int> PivotTableGroupingCompletionStates { get; }

        /// <summary>Gets decoded SXRng PivotTable grouping records grouped by grouping kind, boundary mode, and completion state.</summary>
        internal IReadOnlyDictionary<string, int> PivotTableGroupingStates { get; }

        /// <summary>Gets decoded SXRng numeric grouping records grouped by start, end, and interval values.</summary>
        internal IReadOnlyDictionary<string, int> PivotTableGroupingNumericRanges { get; }

        /// <summary>Gets decoded SXRng date grouping records grouped by start, end, and interval values.</summary>
        internal IReadOnlyDictionary<string, int> PivotTableGroupingDateRanges { get; }

        /// <summary>Gets decoded SXFormula calculated-item formula records grouped by cache-field scope.</summary>
        internal IReadOnlyDictionary<string, int> PivotTableFormulaScopes { get; }

        /// <summary>Gets decoded SXFormula calculated-item formula records grouped by raw cache field index.</summary>
        internal IReadOnlyDictionary<string, int> PivotTableFormulaCacheFieldIndexes { get; }

        /// <summary>Gets decoded SXFormula calculated-item formula records grouped by reserved value.</summary>
        internal IReadOnlyDictionary<string, int> PivotTableFormulaReservedValues { get; }

        /// <summary>Gets decoded SXVDEx PivotTable field flags grouped by flag state.</summary>
        internal IReadOnlyDictionary<string, int> PivotTableExtendedFieldStates { get; }

        /// <summary>Gets decoded SXVDEx PivotTable field flags grouped by complete interaction-permission shape.</summary>
        internal IReadOnlyDictionary<string, int> PivotTableExtendedFieldPermissionStates { get; }

        /// <summary>Gets decoded SXAddl records grouped by PivotTable extension class.</summary>
        internal IReadOnlyDictionary<string, int> PivotTableAdditionalClasses { get; }

        /// <summary>Gets decoded SXAddl records grouped by PivotTable extension detail type.</summary>
        internal IReadOnlyDictionary<string, int> PivotTableAdditionalTypes { get; }

        /// <summary>Gets decoded SXAddl records grouped by class and detail type.</summary>
        internal IReadOnlyDictionary<string, int> PivotTableAdditionalClassTypes { get; }

        /// <summary>Gets decoded SXAddl records grouped by future-record type.</summary>
        internal IReadOnlyDictionary<string, int> PivotTableAdditionalFutureRecordTypes { get; }

        /// <summary>Gets decoded SXAddl records grouped by future-record flags.</summary>
        internal IReadOnlyDictionary<string, int> PivotTableAdditionalFutureFlags { get; }

        /// <summary>Gets decoded SXAddl records grouped by sequence index in the scanned PivotTable scope.</summary>
        internal IReadOnlyDictionary<string, int> PivotTableAdditionalSequenceIndexes { get; }

        /// <summary>Gets decoded SXAddl records grouped by class, detail type, and payload length.</summary>
        internal IReadOnlyDictionary<string, int> PivotTableAdditionalPayloadLengthsByClassType { get; }

        /// <summary>Gets decoded SXAddl SxcCache/SXDId records grouped by PivotCache identifier.</summary>
        internal IReadOnlyDictionary<string, int> PivotTableAdditionalCacheIds { get; }

        /// <summary>Gets decoded SXAddl records grouped by class nesting depth before each record is applied.</summary>
        internal IReadOnlyDictionary<string, int> PivotTableAdditionalClassDepthsBefore { get; }

        /// <summary>Gets decoded SXAddl records grouped by class nesting depth after each record is applied.</summary>
        internal IReadOnlyDictionary<string, int> PivotTableAdditionalClassDepthsAfter { get; }

        /// <summary>Gets decoded SXAddl records grouped by shallow class-stack transition.</summary>
        internal IReadOnlyDictionary<string, int> PivotTableAdditionalClassTransitions { get; }

        /// <summary>Gets decoded SXAddl records grouped by class, detail type, and shallow class-stack transition.</summary>
        internal IReadOnlyDictionary<string, int> PivotTableAdditionalClassTransitionsByClassType { get; }

        /// <summary>Gets preserve-only chart BIFF records grouped by shallow category.</summary>
        internal IReadOnlyDictionary<LegacyXlsChartRecordKind, int> ChartRecordsByKind { get; }

        /// <summary>Gets preserve-only chart BIFF records grouped by record name.</summary>
        internal IReadOnlyDictionary<string, int> ChartRecordsByName { get; }

        /// <summary>Gets preserve-only chart BIFF records grouped by record name and payload length.</summary>
        internal IReadOnlyDictionary<string, int> ChartRecordsByNameAndPayloadLength { get; }

        /// <summary>Gets workbook-level chart model-shape states derived from preserve-only chart BIFF records.</summary>
        internal IReadOnlyDictionary<string, int> ChartWorkbookStates { get; }

        /// <summary>Gets chart BIFF records grouped by container nesting depth before each record is applied.</summary>
        internal IReadOnlyDictionary<string, int> ChartRecordsByContainerDepthBefore { get; }

        /// <summary>Gets chart BIFF records grouped by container nesting depth after each record is applied.</summary>
        internal IReadOnlyDictionary<string, int> ChartRecordsByContainerDepthAfter { get; }

        /// <summary>Gets chart BIFF records grouped by shallow container transition state.</summary>
        internal IReadOnlyDictionary<string, int> ChartRecordsByContainerTransition { get; }

        /// <summary>Gets chart BIFF records grouped by record name and container depth before each record.</summary>
        internal IReadOnlyDictionary<string, int> ChartRecordsByNameAndContainerDepth { get; }

        /// <summary>Gets chart BIFF records grouped by record name and shallow container transition state.</summary>
        internal IReadOnlyDictionary<string, int> ChartRecordsByNameAndContainerTransition { get; }

        /// <summary>Gets preserve-only chart BIFF chart-type records grouped by decoded chart family.</summary>
        internal IReadOnlyDictionary<string, int> ChartRecordsByChartType { get; }

        /// <summary>Gets Chart records grouped by decoded chart rectangle.</summary>
        internal IReadOnlyDictionary<string, int> ChartRecordsByRectangle { get; }

        /// <summary>Gets Axis records grouped by decoded axis type.</summary>
        internal IReadOnlyDictionary<string, int> ChartRecordsByAxisType { get; }

        /// <summary>Gets ChartFormat records grouped by varied data-point color state.</summary>
        internal IReadOnlyDictionary<string, int> ChartGroupVariedColorStates { get; }

        /// <summary>Gets ChartFormat records grouped by chart-group drawing order.</summary>
        internal IReadOnlyDictionary<string, int> ChartGroupDrawingOrders { get; }

        /// <summary>Gets AxesUsed records grouped by decoded axis group count.</summary>
        internal IReadOnlyDictionary<string, int> ChartRecordsByAxesUsedCount { get; }

        /// <summary>Gets CatSerRange records grouped by decoded crossing and interval values.</summary>
        internal IReadOnlyDictionary<string, int> ChartCategorySeriesRangeIntervals { get; }

        /// <summary>Gets CatSerRange records grouped by decoded axis state flags.</summary>
        internal IReadOnlyDictionary<string, int> ChartCategorySeriesRangeStates { get; }

        /// <summary>Gets AxcExt records grouped by decoded minimum, maximum, and crossing dates.</summary>
        internal IReadOnlyDictionary<string, int> ChartAxisExtensionDateRanges { get; }

        /// <summary>Gets AxcExt records grouped by decoded major, minor, and base date units.</summary>
        internal IReadOnlyDictionary<string, int> ChartAxisExtensionDateUnits { get; }

        /// <summary>Gets AxcExt records grouped by automatic date-axis flags.</summary>
        internal IReadOnlyDictionary<string, int> ChartAxisExtensionStates { get; }

        /// <summary>Gets AxcExt records grouped by reserved byte state.</summary>
        internal IReadOnlyDictionary<string, int> ChartAxisExtensionReservedStates { get; }

        /// <summary>Gets CatLab records grouped by decoded axis-label alignment.</summary>
        internal IReadOnlyDictionary<string, int> ChartCategoryLabelAlignments { get; }

        /// <summary>Gets CatLab records grouped by axis-label offset percentage.</summary>
        internal IReadOnlyDictionary<string, int> ChartCategoryLabelOffsets { get; }

        /// <summary>Gets CatLab records grouped by automatic category-label count state.</summary>
        internal IReadOnlyDictionary<string, int> ChartCategoryLabelCountStates { get; }

        /// <summary>Gets AxisLineFormat records grouped by decoded formatted axis component.</summary>
        internal IReadOnlyDictionary<string, int> ChartAxisLineFormatTargets { get; }

        /// <summary>Gets Series records grouped by decoded category data type.</summary>
        internal IReadOnlyDictionary<string, int> ChartSeriesCategoryDataTypes { get; }

        /// <summary>Gets Series records grouped by decoded value data type.</summary>
        internal IReadOnlyDictionary<string, int> ChartSeriesValueDataTypes { get; }

        /// <summary>Gets Series records grouped by decoded bubble-size data type.</summary>
        internal IReadOnlyDictionary<string, int> ChartSeriesBubbleSizeDataTypes { get; }

        /// <summary>Gets Series records grouped by category, value, and bubble-size counts.</summary>
        internal IReadOnlyDictionary<string, int> ChartSeriesValueCounts { get; }

        /// <summary>Gets SerToCrt records grouped by referenced ChartFormat index.</summary>
        internal IReadOnlyDictionary<string, int> ChartSeriesChartGroupIndexes { get; }

        /// <summary>Gets SeriesList records grouped by declared series index count.</summary>
        internal IReadOnlyDictionary<string, int> ChartSeriesListDeclaredCounts { get; }

        /// <summary>Gets SeriesList records grouped by decoded series index count.</summary>
        internal IReadOnlyDictionary<string, int> ChartSeriesListDecodedCounts { get; }

        /// <summary>Gets SeriesList records grouped by whether all declared indexes were present.</summary>
        internal IReadOnlyDictionary<string, int> ChartSeriesListCompletenessStates { get; }

        /// <summary>Gets SeriesList records grouped by decoded one-based index validity.</summary>
        internal IReadOnlyDictionary<string, int> ChartSeriesListIndexValidityStates { get; }

        /// <summary>Gets SBaseRef records grouped by decoded referenced PivotTable-view range.</summary>
        internal IReadOnlyDictionary<string, int> ChartPivotViewReferences { get; }

        /// <summary>Gets SIIndex records grouped by raw chart data-cache sequence index.</summary>
        internal IReadOnlyDictionary<string, int> ChartSeriesDataCacheIndexes { get; }

        /// <summary>Gets SIIndex records grouped by decoded chart data-cache sequence type.</summary>
        internal IReadOnlyDictionary<string, int> ChartSeriesDataCacheTypes { get; }

        /// <summary>Gets BRAI records grouped by decoded chart source part.</summary>
        internal IReadOnlyDictionary<string, int> ChartDataSourceIds { get; }

        /// <summary>Gets BRAI records grouped by decoded referenced data type.</summary>
        internal IReadOnlyDictionary<string, int> ChartDataSourceReferenceTypes { get; }

        /// <summary>Gets BRAI records grouped by raw number format identifier.</summary>
        internal IReadOnlyDictionary<string, int> ChartDataSourceNumberFormatIds { get; }

        /// <summary>Gets BRAI records grouped by declared ChartParsedFormula byte count.</summary>
        internal IReadOnlyDictionary<string, int> ChartDataSourceFormulaByteCounts { get; }

        /// <summary>Gets BRAI records grouped by ChartParsedFormula text projection state.</summary>
        internal IReadOnlyDictionary<string, int> ChartDataSourceFormulaProjectionStates { get; }

        /// <summary>Gets BRAI records grouped by projected ChartParsedFormula text.</summary>
        internal IReadOnlyDictionary<string, int> ChartDataSourceFormulaTexts { get; }

        /// <summary>Gets BRAI formula projection failures grouped by stable failure code.</summary>
        internal IReadOnlyDictionary<string, int> ChartDataSourceFormulaProjectionFailures { get; }

        /// <summary>Gets BRAI formula projection failures grouped by unsupported token byte.</summary>
        internal IReadOnlyDictionary<string, int> ChartDataSourceFormulaProjectionFailuresByToken { get; }

        /// <summary>Gets BRAI formula projection failures grouped by unsupported token name.</summary>
        internal IReadOnlyDictionary<string, int> ChartDataSourceFormulaProjectionFailuresByTokenName { get; }

        /// <summary>Gets BRAI formula projection failures grouped by token offset.</summary>
        internal IReadOnlyDictionary<string, int> ChartDataSourceFormulaProjectionFailuresByOffset { get; }

        /// <summary>Gets BRAI records grouped by decoded source, reference, number format, and formula-byte state.</summary>
        internal IReadOnlyDictionary<string, int> ChartDataSourceStates { get; }

        /// <summary>Gets DataFormat records grouped by whether formatting targets a series or point.</summary>
        internal IReadOnlyDictionary<string, int> ChartDataFormatTargets { get; }

        /// <summary>Gets DataFormat records grouped by raw series index.</summary>
        internal IReadOnlyDictionary<string, int> ChartDataFormatSeriesIndexes { get; }

        /// <summary>Gets DataFormat records grouped by raw point index.</summary>
        internal IReadOnlyDictionary<string, int> ChartDataFormatPointIndexes { get; }

        /// <summary>Gets DataFormat records grouped by raw chart-format order.</summary>
        internal IReadOnlyDictionary<string, int> ChartDataFormatOrders { get; }

        /// <summary>Gets DataFormat records grouped by target, point, series, and order.</summary>
        internal IReadOnlyDictionary<string, int> ChartDataFormatStates { get; }

        /// <summary>Gets IFmtRecord records grouped by raw number format identifier.</summary>
        internal IReadOnlyDictionary<string, int> ChartNumberFormatIds { get; }

        /// <summary>Gets FontX records grouped by raw font index.</summary>
        internal IReadOnlyDictionary<string, int> ChartFontIndexes { get; }

        /// <summary>Gets Dat records grouped by decoded data-table display options.</summary>
        internal IReadOnlyDictionary<string, int> ChartDataTableOptions { get; }

        /// <summary>Gets Dat records grouped by reserved-bit state.</summary>
        internal IReadOnlyDictionary<string, int> ChartDataTableReservedStates { get; }

        /// <summary>Gets SerAuxErrBar records grouped by decoded error-bar direction.</summary>
        internal IReadOnlyDictionary<string, int> ChartErrorBarDirections { get; }

        /// <summary>Gets SerAuxErrBar records grouped by decoded error-bar value source.</summary>
        internal IReadOnlyDictionary<string, int> ChartErrorBarValueSources { get; }

        /// <summary>Gets SerAuxErrBar records grouped by fixed value and custom value count.</summary>
        internal IReadOnlyDictionary<string, int> ChartErrorBarValues { get; }

        /// <summary>Gets SerAuxErrBar records grouped by direction, value source, tee-top, and meaningful-field states.</summary>
        internal IReadOnlyDictionary<string, int> ChartErrorBarStates { get; }

        /// <summary>Gets SerAuxErrBar records grouped by reserved-byte state.</summary>
        internal IReadOnlyDictionary<string, int> ChartErrorBarReservedStates { get; }

        /// <summary>Gets Bar records grouped by decoded overlap percentage.</summary>
        internal IReadOnlyDictionary<string, int> ChartBarOverlapPercentages { get; }

        /// <summary>Gets Bar records grouped by decoded category gap width percentage.</summary>
        internal IReadOnlyDictionary<string, int> ChartBarGapWidths { get; }

        /// <summary>Gets Bar records grouped by decoded orientation, stacking, percentage, and shadow flags.</summary>
        internal IReadOnlyDictionary<string, int> ChartBarStates { get; }

        /// <summary>Gets Line records grouped by decoded stacking, percentage, and shadow flags.</summary>
        internal IReadOnlyDictionary<string, int> ChartLineStates { get; }

        /// <summary>Gets Line records grouped by reserved-bit state.</summary>
        internal IReadOnlyDictionary<string, int> ChartLineReservedStates { get; }

        /// <summary>Gets Line records grouped by percentage-stacking validity.</summary>
        internal IReadOnlyDictionary<string, int> ChartLinePercentStackedStates { get; }

        /// <summary>Gets Area records grouped by decoded stacking, percentage, and shadow flags.</summary>
        internal IReadOnlyDictionary<string, int> ChartAreaStates { get; }

        /// <summary>Gets Area records grouped by reserved-bit state.</summary>
        internal IReadOnlyDictionary<string, int> ChartAreaReservedStates { get; }

        /// <summary>Gets Area records grouped by percentage-stacking validity.</summary>
        internal IReadOnlyDictionary<string, int> ChartAreaPercentStackedStates { get; }

        /// <summary>Gets BopPop records grouped by decoded chart group subtype.</summary>
        internal IReadOnlyDictionary<string, int> ChartBopPopSubtypes { get; }

        /// <summary>Gets BopPop records grouped by decoded split mode.</summary>
        internal IReadOnlyDictionary<string, int> ChartBopPopSplitTypes { get; }

        /// <summary>Gets BopPop records grouped by decoded split position, percentage, size, gap, and value.</summary>
        internal IReadOnlyDictionary<string, int> ChartBopPopSplitValues { get; }

        /// <summary>Gets BopPop records grouped by decoded subtype, split, automatic split, and shadow state.</summary>
        internal IReadOnlyDictionary<string, int> ChartBopPopStates { get; }

        /// <summary>Gets BopPop records grouped by reserved-bit state.</summary>
        internal IReadOnlyDictionary<string, int> ChartBopPopReservedStates { get; }

        /// <summary>Gets BopPopCustom records grouped by declared data point count.</summary>
        internal IReadOnlyDictionary<string, int> ChartBopPopCustomDataPointCounts { get; }

        /// <summary>Gets BopPopCustom records grouped by secondary bar/pie data point count.</summary>
        internal IReadOnlyDictionary<string, int> ChartBopPopCustomSecondaryCounts { get; }

        /// <summary>Gets BopPopCustom records grouped by secondary bar/pie data point indexes.</summary>
        internal IReadOnlyDictionary<string, int> ChartBopPopCustomSecondaryIndexes { get; }

        /// <summary>Gets BopPopCustom records grouped by bitmap completion state.</summary>
        internal IReadOnlyDictionary<string, int> ChartBopPopCustomCompletionStates { get; }

        /// <summary>Gets BopPopCustom records grouped by decoded count, marker, and consistency state.</summary>
        internal IReadOnlyDictionary<string, int> ChartBopPopCustomStates { get; }

        /// <summary>Gets Chart3d records grouped by rotation and elevation angles.</summary>
        internal IReadOnlyDictionary<string, int> ChartThreeDimensionalViewAngles { get; }

        /// <summary>Gets Chart3d records grouped by field-of-view, height, depth, and gap values.</summary>
        internal IReadOnlyDictionary<string, int> ChartThreeDimensionalScaleValues { get; }

        /// <summary>Gets Chart3d records grouped by decoded perspective, clustering, scaling, pie, and wall flags.</summary>
        internal IReadOnlyDictionary<string, int> ChartThreeDimensionalStates { get; }

        /// <summary>Gets Chart3d records grouped by reserved-bit state.</summary>
        internal IReadOnlyDictionary<string, int> ChartThreeDimensionalReservedStates { get; }

        /// <summary>Gets Chart3DBarShape records grouped by decoded data-point base shape.</summary>
        internal IReadOnlyDictionary<string, int> ChartThreeDimensionalBarShapeRisers { get; }

        /// <summary>Gets Chart3DBarShape records grouped by decoded tapering mode.</summary>
        internal IReadOnlyDictionary<string, int> ChartThreeDimensionalBarShapeTapers { get; }

        /// <summary>Gets Chart3DBarShape records grouped by decoded base-shape and tapering mode.</summary>
        internal IReadOnlyDictionary<string, int> ChartThreeDimensionalBarShapeStates { get; }

        /// <summary>Gets Scatter records grouped by decoded bubble-size ratio.</summary>
        internal IReadOnlyDictionary<string, int> ChartScatterBubbleSizeRatios { get; }

        /// <summary>Gets Scatter records grouped by decoded bubble-size representation.</summary>
        internal IReadOnlyDictionary<string, int> ChartScatterBubbleSizeRepresentations { get; }

        /// <summary>Gets Scatter records grouped by bubble-size ratio validity.</summary>
        internal IReadOnlyDictionary<string, int> ChartScatterBubbleSizeRatioStates { get; }

        /// <summary>Gets Scatter records grouped by decoded bubble, negative-bubble, shadow, and size-representation state.</summary>
        internal IReadOnlyDictionary<string, int> ChartScatterStates { get; }

        /// <summary>Gets Fbi or Fbi2 records grouped by decoded chart font scale basis.</summary>
        internal IReadOnlyDictionary<string, int> ChartFontBasisScaleBasis { get; }

        /// <summary>Gets Fbi or Fbi2 records grouped by referenced chart font index.</summary>
        internal IReadOnlyDictionary<string, int> ChartFontBasisFontIndexes { get; }

        /// <summary>Gets Fbi or Fbi2 records grouped by decoded basis, default font height, scale basis, and font index.</summary>
        internal IReadOnlyDictionary<string, int> ChartFontBasisStates { get; }

        /// <summary>Gets CrtLayout12 records grouped by decoded layout modes.</summary>
        internal IReadOnlyDictionary<string, int> ChartLayout12ModePairs { get; }

        /// <summary>Gets CrtLayout12 records grouped by decoded automatic legend layout type.</summary>
        internal IReadOnlyDictionary<string, int> ChartLayout12AutoLayoutTypes { get; }

        /// <summary>Gets CrtLayout12 records grouped by checksum value.</summary>
        internal IReadOnlyDictionary<string, int> ChartLayout12Checksums { get; }

        /// <summary>Gets CrtLayout12 records grouped by decoded layout rectangle values.</summary>
        internal IReadOnlyDictionary<string, int> ChartLayout12Rectangles { get; }

        /// <summary>Gets CrtLayout12A records grouped by decoded plot-area target.</summary>
        internal IReadOnlyDictionary<string, int> ChartPlotAreaLayout12Targets { get; }

        /// <summary>Gets CrtLayout12A records grouped by decoded layout modes.</summary>
        internal IReadOnlyDictionary<string, int> ChartPlotAreaLayout12ModePairs { get; }

        /// <summary>Gets CrtLayout12A records grouped by checksum value.</summary>
        internal IReadOnlyDictionary<string, int> ChartPlotAreaLayout12Checksums { get; }

        /// <summary>Gets CrtLayout12A records grouped by SPRC plot-area bounds.</summary>
        internal IReadOnlyDictionary<string, int> ChartPlotAreaLayout12Bounds { get; }

        /// <summary>Gets CrtLayout12A records grouped by decoded layout rectangle values.</summary>
        internal IReadOnlyDictionary<string, int> ChartPlotAreaLayout12Rectangles { get; }

        /// <summary>Gets ChartFrtInfo records grouped by originator and writer application version.</summary>
        internal IReadOnlyDictionary<string, int> ChartFutureRecordInfoVersions { get; }

        /// <summary>Gets ChartFrtInfo records grouped by decoded future-record range count.</summary>
        internal IReadOnlyDictionary<string, int> ChartFutureRecordInfoRangeCounts { get; }

        /// <summary>Gets ChartFrtInfo records grouped by declared future-record id range.</summary>
        internal IReadOnlyDictionary<string, int> ChartFutureRecordInfoRanges { get; }

        /// <summary>Gets StartBlock and EndBlock records grouped by block direction.</summary>
        internal IReadOnlyDictionary<string, int> ChartFutureBlockDirections { get; }

        /// <summary>Gets StartBlock and EndBlock records grouped by decoded object kind.</summary>
        internal IReadOnlyDictionary<string, int> ChartFutureBlockObjectKinds { get; }

        /// <summary>Gets StartBlock and EndBlock records grouped by compact scope key.</summary>
        internal IReadOnlyDictionary<string, int> ChartFutureBlockScopes { get; }

        /// <summary>Gets Units records grouped by reserved field value.</summary>
        internal IReadOnlyDictionary<string, int> ChartUnitsReservedValues { get; }

        /// <summary>Gets Units records grouped by whether the reserved field is zero.</summary>
        internal IReadOnlyDictionary<string, int> ChartUnitsReservedStates { get; }

        /// <summary>Gets CrtMlFrt records grouped by declared XmlTkChain byte count.</summary>
        internal IReadOnlyDictionary<string, int> ChartXmlTokenChainDeclaredByteCounts { get; }

        /// <summary>Gets CrtMlFrt records grouped by XmlTkChain bytes present in the first record segment.</summary>
        internal IReadOnlyDictionary<string, int> ChartXmlTokenChainFirstSegmentByteCounts { get; }

        /// <summary>Gets CrtMlFrt records grouped by whether the declared XmlTkChain is complete in the first record segment.</summary>
        internal IReadOnlyDictionary<string, int> ChartXmlTokenChainCompletionStates { get; }

        /// <summary>Gets CrtMlFrt records grouped by ignored trailing field state.</summary>
        internal IReadOnlyDictionary<string, int> ChartXmlTokenChainTrailingStates { get; }

        /// <summary>Gets ShtProps records grouped by decoded empty-cell plotting mode.</summary>
        internal IReadOnlyDictionary<string, int> ChartSheetPropertyEmptyCellModes { get; }

        /// <summary>Gets ShtProps records grouped by decoded chart property flags.</summary>
        internal IReadOnlyDictionary<string, int> ChartSheetPropertyStates { get; }

        /// <summary>Gets LineFormat records grouped by decoded line style.</summary>
        internal IReadOnlyDictionary<string, int> ChartLineFormatStyles { get; }

        /// <summary>Gets LineFormat records grouped by decoded line weight.</summary>
        internal IReadOnlyDictionary<string, int> ChartLineFormatWeights { get; }

        /// <summary>Gets LineFormat records grouped by decoded RGB color.</summary>
        internal IReadOnlyDictionary<string, int> ChartLineFormatColors { get; }

        /// <summary>Gets LineFormat records grouped by chart color index.</summary>
        internal IReadOnlyDictionary<string, int> ChartLineFormatColorIndexes { get; }

        /// <summary>Gets LineFormat records grouped by decoded flag state.</summary>
        internal IReadOnlyDictionary<string, int> ChartLineFormatStates { get; }

        /// <summary>Gets AreaFormat records grouped by decoded fill pattern.</summary>
        internal IReadOnlyDictionary<string, int> ChartAreaFormatPatterns { get; }

        /// <summary>Gets AreaFormat records grouped by decoded foreground and background RGB color.</summary>
        internal IReadOnlyDictionary<string, int> ChartAreaFormatColors { get; }

        /// <summary>Gets AreaFormat records grouped by chart foreground and background color index.</summary>
        internal IReadOnlyDictionary<string, int> ChartAreaFormatColorIndexes { get; }

        /// <summary>Gets AreaFormat records grouped by decoded flag state.</summary>
        internal IReadOnlyDictionary<string, int> ChartAreaFormatStates { get; }

        /// <summary>Gets MarkerFormat records grouped by decoded marker type.</summary>
        internal IReadOnlyDictionary<string, int> ChartMarkerFormatTypes { get; }

        /// <summary>Gets MarkerFormat records grouped by marker size in twips.</summary>
        internal IReadOnlyDictionary<string, int> ChartMarkerFormatSizes { get; }

        /// <summary>Gets MarkerFormat records grouped by decoded foreground and background RGB color.</summary>
        internal IReadOnlyDictionary<string, int> ChartMarkerFormatColors { get; }

        /// <summary>Gets MarkerFormat records grouped by chart foreground and background color index.</summary>
        internal IReadOnlyDictionary<string, int> ChartMarkerFormatColorIndexes { get; }

        /// <summary>Gets MarkerFormat records grouped by decoded flag state.</summary>
        internal IReadOnlyDictionary<string, int> ChartMarkerFormatStates { get; }

        /// <summary>Gets PieFormat records grouped by decoded explosion distance percentage.</summary>
        internal IReadOnlyDictionary<string, int> ChartPieFormatExplosions { get; }

        /// <summary>Gets SerFmt records grouped by enabled formatting flag.</summary>
        internal IReadOnlyDictionary<string, int> ChartSeriesFormatFlags { get; }

        /// <summary>Gets SerFmt records grouped by full decoded flag state.</summary>
        internal IReadOnlyDictionary<string, int> ChartSeriesFormatStates { get; }

        /// <summary>Gets SerFmt records grouped by reserved bit value.</summary>
        internal IReadOnlyDictionary<string, int> ChartSeriesFormatReservedValues { get; }

        /// <summary>Gets SerFmt records grouped by whether reserved bits are zero.</summary>
        internal IReadOnlyDictionary<string, int> ChartSeriesFormatReservedStates { get; }

        /// <summary>Gets ClrtClient records grouped by declared color count.</summary>
        internal IReadOnlyDictionary<string, int> ChartClientColorPaletteDeclaredCounts { get; }

        /// <summary>Gets ClrtClient records grouped by decoded color count.</summary>
        internal IReadOnlyDictionary<string, int> ChartClientColorPaletteDecodedCounts { get; }

        /// <summary>Gets ClrtClient records grouped by whether all declared colors were present.</summary>
        internal IReadOnlyDictionary<string, int> ChartClientColorPaletteCompletenessStates { get; }

        /// <summary>Gets ClrtClient records grouped by whether the expected three colors are present.</summary>
        internal IReadOnlyDictionary<string, int> ChartClientColorPaletteExpectedCountStates { get; }

        /// <summary>Gets ClrtClient records grouped by decoded role-specific color.</summary>
        internal IReadOnlyDictionary<string, int> ChartClientColorPaletteColors { get; }

        /// <summary>Gets GelFrame OfficeArt records grouped by record type.</summary>
        internal IReadOnlyDictionary<string, int> ChartGelFrameOfficeArtRecordsByType { get; }

        /// <summary>Gets GelFrame OfficeArt records grouped by container or leaf state.</summary>
        internal IReadOnlyDictionary<string, int> ChartGelFrameOfficeArtRecordsByContainerState { get; }

        /// <summary>Gets GelFrame records grouped by OfficeArtFOPT property count.</summary>
        internal IReadOnlyDictionary<string, int> ChartGelFrameShapePropertyCounts { get; }

        /// <summary>Gets GelFrame OfficeArtFOPT properties grouped by property name.</summary>
        internal IReadOnlyDictionary<string, int> ChartGelFrameShapePropertiesByName { get; }

        /// <summary>Gets GelFrame OfficeArtFOPT properties grouped by property family.</summary>
        internal IReadOnlyDictionary<string, int> ChartGelFrameShapePropertiesByGroup { get; }

        /// <summary>Gets GelFrame OfficeArtFOPT properties grouped by complex and BLIP flag state.</summary>
        internal IReadOnlyDictionary<string, int> ChartGelFrameShapePropertiesByFlagState { get; }

        /// <summary>Gets GelFrame OfficeArtFOPT properties grouped by raw property value.</summary>
        internal IReadOnlyDictionary<string, int> ChartGelFrameShapePropertiesByValue { get; }

        /// <summary>Gets AttachedLabel records grouped by decoded displayed data-label element.</summary>
        internal IReadOnlyDictionary<string, int> ChartAttachedLabelFlags { get; }

        /// <summary>Gets AttachedLabel records grouped by full decoded flag state.</summary>
        internal IReadOnlyDictionary<string, int> ChartAttachedLabelStates { get; }

        /// <summary>Gets DefaultText records grouped by decoded target scope.</summary>
        internal IReadOnlyDictionary<string, int> ChartDefaultTextTargets { get; }

        /// <summary>Gets Text records grouped by decoded horizontal alignment.</summary>
        internal IReadOnlyDictionary<string, int> ChartTextHorizontalAlignments { get; }

        /// <summary>Gets Text records grouped by decoded vertical alignment.</summary>
        internal IReadOnlyDictionary<string, int> ChartTextVerticalAlignments { get; }

        /// <summary>Gets Text records grouped by decoded data-label position.</summary>
        internal IReadOnlyDictionary<string, int> ChartTextDataLabelPositions { get; }

        /// <summary>Gets Text records grouped by decoded flag name.</summary>
        internal IReadOnlyDictionary<string, int> ChartTextFlags { get; }

        /// <summary>Gets ObjectLink records grouped by decoded linked chart object.</summary>
        internal IReadOnlyDictionary<string, int> ChartObjectLinkTargets { get; }

        /// <summary>Gets Legend records grouped by decoded layout.</summary>
        internal IReadOnlyDictionary<string, int> ChartLegendLayouts { get; }

        /// <summary>Gets Legend records grouped by entry-spacing validity.</summary>
        internal IReadOnlyDictionary<string, int> ChartLegendSpacingStates { get; }

        /// <summary>Gets Legend records grouped by reserved-bit validity.</summary>
        internal IReadOnlyDictionary<string, int> ChartLegendReservedStates { get; }

        /// <summary>Gets Legend records grouped by automatic-position consistency.</summary>
        internal IReadOnlyDictionary<string, int> ChartLegendAutoPositionStates { get; }

        /// <summary>Gets Legend records grouped by data-table layout consistency.</summary>
        internal IReadOnlyDictionary<string, int> ChartLegendDataTableStates { get; }

        /// <summary>Gets Tick records grouped by decoded major tick-mark location.</summary>
        internal IReadOnlyDictionary<string, int> ChartTickMajorLocations { get; }

        /// <summary>Gets Tick records grouped by decoded axis-label location.</summary>
        internal IReadOnlyDictionary<string, int> ChartTickLabelLocations { get; }

        /// <summary>Gets ValueRange records grouped by decoded value-axis scale fields.</summary>
        internal IReadOnlyDictionary<string, int> ChartValueRangeScales { get; }

        /// <summary>Gets ValueRange records grouped by decoded automatic scale and axis-direction flags.</summary>
        internal IReadOnlyDictionary<string, int> ChartValueRangeStates { get; }

        /// <summary>Gets Pos records grouped by decoded upper-left and lower-right position modes.</summary>
        internal IReadOnlyDictionary<string, int> ChartPositionModePairs { get; }

        /// <summary>Gets Pos records grouped by decoded coordinate and size fields.</summary>
        internal IReadOnlyDictionary<string, int> ChartPositionRectangles { get; }

        /// <summary>Gets Pos records grouped by semantic object type inferred from position modes.</summary>
        internal IReadOnlyDictionary<string, int> ChartPositionSemanticTypes { get; }

        /// <summary>Gets Pos records grouped by decoded coordinate meaning.</summary>
        internal IReadOnlyDictionary<string, int> ChartPositionCoordinateMeanings { get; }

        /// <summary>Gets Pos records grouped by ignored coordinate state.</summary>
        internal IReadOnlyDictionary<string, int> ChartPositionIgnoredCoordinateStates { get; }

        /// <summary>Gets Pos records grouped by whether the position mode pair is a known semantic combination.</summary>
        internal IReadOnlyDictionary<string, int> ChartPositionKnownSemanticStates { get; }

        /// <summary>Gets Frame records grouped by decoded frame type.</summary>
        internal IReadOnlyDictionary<string, int> ChartFrameTypes { get; }

        /// <summary>Gets Frame records grouped by automatic size and position flags.</summary>
        internal IReadOnlyDictionary<string, int> ChartFrameAutoStates { get; }

        /// <summary>Gets PlotGrowth records grouped by decoded horizontal and vertical growth factors.</summary>
        internal IReadOnlyDictionary<string, int> ChartPlotGrowthFactors { get; }

        /// <summary>Gets preserve-only chart BIFF records grouped by workbook or sheet location.</summary>
        internal IReadOnlyDictionary<string, int> ChartRecordsByLocation { get; }

        /// <summary>Gets preserve-only drawing and object BIFF records grouped by shallow category.</summary>
        internal IReadOnlyDictionary<LegacyXlsDrawingRecordKind, int> DrawingRecordsByKind { get; }

        /// <summary>Gets preserve-only drawing and object BIFF records grouped by record name.</summary>
        internal IReadOnlyDictionary<string, int> DrawingRecordsByName { get; }

        /// <summary>Gets OBJ records grouped by decoded common-object type identifier.</summary>
        internal IReadOnlyDictionary<string, int> DrawingRecordsByObjectType { get; }

        /// <summary>Gets OBJ records grouped by decoded common-object type name.</summary>
        internal IReadOnlyDictionary<string, int> DrawingRecordsByObjectTypeName { get; }

        /// <summary>Gets OBJ records grouped by decoded common-object flag bitfield.</summary>
        internal IReadOnlyDictionary<string, int> DrawingRecordsByObjectFlags { get; }

        /// <summary>Gets OBJ records grouped by decoded common-object flag name.</summary>
        internal IReadOnlyDictionary<string, int> DrawingRecordsByObjectFlagName { get; }

        /// <summary>Gets OBJ subrecords grouped by raw subrecord type.</summary>
        internal IReadOnlyDictionary<string, int> DrawingObjectSubRecordsByType { get; }

        /// <summary>Gets OBJ subrecords grouped by decoded subrecord name.</summary>
        internal IReadOnlyDictionary<string, int> DrawingObjectSubRecordsByName { get; }

        /// <summary>Gets OBJ subrecords grouped by declared payload length.</summary>
        internal IReadOnlyDictionary<string, int> DrawingObjectSubRecordsByDeclaredLength { get; }

        /// <summary>Gets OBJ subrecords grouped by whether the declared payload was fully available.</summary>
        internal IReadOnlyDictionary<string, int> DrawingObjectSubRecordsByCompleteness { get; }

        /// <summary>Gets drawing future-record streams grouped by wrapped BIFF record type.</summary>
        internal IReadOnlyDictionary<string, int> DrawingFutureRecordWrappedTypes { get; }

        /// <summary>Gets drawing future-record streams grouped by raw future-record flags.</summary>
        internal IReadOnlyDictionary<string, int> DrawingFutureRecordFlags { get; }

        /// <summary>Gets drawing future-record streams grouped by whether a cell range reference is present.</summary>
        internal IReadOnlyDictionary<string, int> DrawingFutureRecordReferenceStates { get; }

        /// <summary>Gets drawing future-record streams grouped by decoded cell range reference.</summary>
        internal IReadOnlyDictionary<string, int> DrawingFutureRecordRanges { get; }

        /// <summary>Gets drawing future-record streams grouped by remaining stream byte count.</summary>
        internal IReadOnlyDictionary<string, int> DrawingFutureRecordStreamByteCounts { get; }

        /// <summary>Gets HFPicture records grouped by decoded header state.</summary>
        internal IReadOnlyDictionary<string, int> DrawingHeaderFooterPictureHeaderStates { get; }

        /// <summary>Gets HFPicture records grouped by declared OfficeArt drawing kind.</summary>
        internal IReadOnlyDictionary<string, int> DrawingHeaderFooterPictureDrawingKinds { get; }

        /// <summary>Gets HFPicture records grouped by continuation state.</summary>
        internal IReadOnlyDictionary<string, int> DrawingHeaderFooterPictureContinuationStates { get; }

        /// <summary>Gets HFPicture records grouped by future-record flags.</summary>
        internal IReadOnlyDictionary<string, int> DrawingHeaderFooterPictureFutureRecordFlags { get; }

        /// <summary>Gets HFPicture records grouped by embedded OfficeArt byte count.</summary>
        internal IReadOnlyDictionary<string, int> DrawingHeaderFooterPictureDrawingByteCounts { get; }

        /// <summary>Gets TxO text-object records grouped by decoded horizontal and vertical alignment.</summary>
        internal IReadOnlyDictionary<string, int> DrawingTextObjectAlignments { get; }

        /// <summary>Gets TxO text-object records grouped by decoded rotation.</summary>
        internal IReadOnlyDictionary<string, int> DrawingTextObjectRotations { get; }

        /// <summary>Gets TxO text-object records grouped by declared text character count.</summary>
        internal IReadOnlyDictionary<string, int> DrawingTextObjectTextLengths { get; }

        /// <summary>Gets TxO text-object records grouped by declared formatting-run byte count.</summary>
        internal IReadOnlyDictionary<string, int> DrawingTextObjectFormattingRunByteCounts { get; }

        /// <summary>Gets TxO text-object records grouped by optional formula byte count.</summary>
        internal IReadOnlyDictionary<string, int> DrawingTextObjectFormulaByteCounts { get; }

        /// <summary>Gets TxO text-object records grouped by decoded flag state.</summary>
        internal IReadOnlyDictionary<string, int> DrawingTextObjectFlags { get; }

        /// <summary>Gets MsoDrawing records grouped by decoded top-level Escher record type.</summary>
        internal IReadOnlyDictionary<string, int> DrawingRecordsByEscherRecordType { get; }

        /// <summary>Gets MsoDrawing records grouped by decoded top-level Escher record type name.</summary>
        internal IReadOnlyDictionary<string, int> DrawingRecordsByEscherRecordTypeName { get; }

        /// <summary>Gets nested OfficeArt records grouped by raw record type.</summary>
        internal IReadOnlyDictionary<string, int> DrawingOfficeArtRecordsByType { get; }

        /// <summary>Gets nested OfficeArt records grouped by decoded record type name.</summary>
        internal IReadOnlyDictionary<string, int> DrawingOfficeArtRecordsByTypeName { get; }

        /// <summary>Gets nested OfficeArt records grouped by traversal depth.</summary>
        internal IReadOnlyDictionary<string, int> DrawingOfficeArtRecordsByDepth { get; }

        /// <summary>Gets nested OfficeArt records grouped by container or leaf state.</summary>
        internal IReadOnlyDictionary<string, int> DrawingOfficeArtRecordsByContainerState { get; }

        /// <summary>Gets nested OfficeArt records grouped by declared payload length.</summary>
        internal IReadOnlyDictionary<string, int> DrawingOfficeArtRecordsByPayloadLength { get; }

        /// <summary>Gets OfficeArtFDGGBlock records grouped by maximum shape identifier.</summary>
        internal IReadOnlyDictionary<string, int> DrawingGroupBlocksByMaxShapeId { get; }

        /// <summary>Gets OfficeArtFDGGBlock records grouped by declared identifier-cluster count.</summary>
        internal IReadOnlyDictionary<string, int> DrawingGroupBlocksByDeclaredIdentifierClusterCount { get; }

        /// <summary>Gets OfficeArtFDGGBlock records grouped by decoded identifier-cluster count.</summary>
        internal IReadOnlyDictionary<string, int> DrawingGroupBlocksByDecodedIdentifierClusterCount { get; }

        /// <summary>Gets OfficeArtFDGGBlock records grouped by saved shape count.</summary>
        internal IReadOnlyDictionary<string, int> DrawingGroupBlocksBySavedShapeCount { get; }

        /// <summary>Gets OfficeArtFDGGBlock records grouped by saved drawing count.</summary>
        internal IReadOnlyDictionary<string, int> DrawingGroupBlocksBySavedDrawingCount { get; }

        /// <summary>Gets OfficeArtIDCL clusters grouped by drawing identifier.</summary>
        internal IReadOnlyDictionary<string, int> DrawingIdentifierClustersByDrawingId { get; }

        /// <summary>Gets OfficeArtIDCL clusters grouped by current shape identifier.</summary>
        internal IReadOnlyDictionary<string, int> DrawingIdentifierClustersByCurrentShapeId { get; }

        /// <summary>Gets OfficeArtFDG records grouped by drawing identifier.</summary>
        internal IReadOnlyDictionary<string, int> DrawingGroupInfosByDrawingId { get; }

        /// <summary>Gets OfficeArtFDG records grouped by drawing shape count.</summary>
        internal IReadOnlyDictionary<string, int> DrawingGroupInfosByShapeCount { get; }

        /// <summary>Gets OfficeArtFDG records grouped by last shape identifier.</summary>
        internal IReadOnlyDictionary<string, int> DrawingGroupInfosByLastShapeId { get; }

        /// <summary>Gets OfficeArtFOPT shape properties grouped by property identifier.</summary>
        internal IReadOnlyDictionary<string, int> DrawingShapePropertiesById { get; }

        /// <summary>Gets OfficeArtFOPT shape properties grouped by decoded property name.</summary>
        internal IReadOnlyDictionary<string, int> DrawingShapePropertiesByName { get; }

        /// <summary>Gets OfficeArtFOPT shape properties grouped by decoded property family.</summary>
        internal IReadOnlyDictionary<string, int> DrawingShapePropertiesByGroup { get; }

        /// <summary>Gets OfficeArtFOPT shape properties grouped by complex and BLIP flag state.</summary>
        internal IReadOnlyDictionary<string, int> DrawingShapePropertiesByFlagState { get; }

        /// <summary>Gets simple OfficeArtFOPT shape properties grouped by raw value.</summary>
        internal IReadOnlyDictionary<string, int> DrawingShapePropertiesByValue { get; }

        /// <summary>Gets complex OfficeArtFOPT shape properties grouped by declared complex byte length.</summary>
        internal IReadOnlyDictionary<string, int> DrawingShapeComplexPropertiesByDeclaredLength { get; }

        /// <summary>Gets complex OfficeArtFOPT shape properties grouped by available complex byte length.</summary>
        internal IReadOnlyDictionary<string, int> DrawingShapeComplexPropertiesByAvailableLength { get; }

        /// <summary>Gets complex OfficeArtFOPT shape properties grouped by decoded text payload.</summary>
        internal IReadOnlyDictionary<string, int> DrawingShapeComplexPropertiesByText { get; }

        /// <summary>Gets OfficeArt FBSE image-store entries grouped by decoded BLIP type.</summary>
        internal IReadOnlyDictionary<string, int> DrawingBlipStoreEntriesByType { get; }

        /// <summary>Gets OfficeArt FBSE image-store entries grouped by workbook or sheet location.</summary>
        internal IReadOnlyDictionary<string, int> DrawingBlipStoreEntriesByLocation { get; }

        /// <summary>Gets OfficeArt FBSE image-store entries grouped by workbook/sheet location and decoded BLIP type.</summary>
        internal IReadOnlyDictionary<string, int> DrawingBlipStoreEntriesByTypeAndLocation { get; }

        /// <summary>Gets OfficeArt FBSE image-store entries grouped by image UID.</summary>
        internal IReadOnlyDictionary<string, int> DrawingBlipStoreEntriesByUid { get; }

        /// <summary>Gets OfficeArt FBSE image-store entries grouped by embedded BLIP record type.</summary>
        internal IReadOnlyDictionary<string, int> DrawingBlipStoreEntriesByEmbeddedRecordType { get; }

        /// <summary>Gets OfficeArt FBSE image-store entries grouped by available embedded payload byte length.</summary>
        internal IReadOnlyDictionary<string, int> DrawingBlipStoreEntriesByEmbeddedPayloadAvailableLength { get; }

        /// <summary>Gets OfficeArt FBSE image-store entries grouped by embedded payload SHA-256 hash.</summary>
        internal IReadOnlyDictionary<string, int> DrawingBlipStoreEntriesByEmbeddedPayloadHash { get; }

        /// <summary>Gets OfficeArt FBSE image-store entries grouped by stored byte size.</summary>
        internal IReadOnlyDictionary<string, int> DrawingBlipStoreEntriesBySize { get; }

        /// <summary>Gets OfficeArt FBSE image-store entries grouped by reference count.</summary>
        internal IReadOnlyDictionary<string, int> DrawingBlipStoreEntriesByReferenceCount { get; }

        /// <summary>Gets OfficeArtFOPT BLIP-family shape properties grouped by workbook or sheet location.</summary>
        internal IReadOnlyDictionary<string, int> DrawingShapeBlipPropertiesByLocation { get; }

        /// <summary>Gets OfficeArtFOPT BLIP-family shape properties grouped by property name and raw value.</summary>
        internal IReadOnlyDictionary<string, int> DrawingShapeBlipPropertiesByNameAndValue { get; }

        /// <summary>Gets OfficeArtFOPT picture BLIP reference properties grouped by workbook or sheet location.</summary>
        internal IReadOnlyDictionary<string, int> DrawingPictureBlipReferencesByLocation { get; }

        /// <summary>Gets OfficeArtFOPT picture BLIP references grouped by referenced image-store id.</summary>
        internal IReadOnlyDictionary<string, int> DrawingPictureBlipReferencesByValue { get; }

        /// <summary>Gets picture drawing states grouped by object, image-store, BLIP reference, and reference-resolution presence.</summary>
        internal IReadOnlyDictionary<string, int> DrawingPictureStates { get; }

        /// <summary>Gets picture drawing states grouped by object, BLIP, reference, and decoded picture-frame counts.</summary>
        internal IReadOnlyDictionary<string, int> DrawingPictureCountStates { get; }

        /// <summary>Gets OfficeArt shape entries grouped by decoded shape type.</summary>
        internal IReadOnlyDictionary<string, int> DrawingShapeEntriesByType { get; }

        /// <summary>Gets OfficeArt shape entries grouped by shape identifier.</summary>
        internal IReadOnlyDictionary<string, int> DrawingShapeEntriesById { get; }

        /// <summary>Gets OfficeArt shape entries grouped by raw flag bitfield.</summary>
        internal IReadOnlyDictionary<string, int> DrawingShapeEntriesByFlags { get; }

        /// <summary>Gets OfficeArt shape entries grouped by whether raw flags contain reserved bits.</summary>
        internal IReadOnlyDictionary<string, int> DrawingShapeEntriesByReservedState { get; }

        /// <summary>Gets OfficeArt shape entries grouped by decoded flag name.</summary>
        internal IReadOnlyDictionary<string, int> DrawingShapeEntriesByFlagName { get; }

        /// <summary>Gets OfficeArt client anchors grouped by start and end cell.</summary>
        internal IReadOnlyDictionary<string, int> DrawingAnchorEntriesByRange { get; }

        /// <summary>Gets OfficeArt client anchors grouped by start and end offsets.</summary>
        internal IReadOnlyDictionary<string, int> DrawingAnchorEntriesByOffset { get; }

        /// <summary>Gets OfficeArt client anchors grouped by raw flag bitfield.</summary>
        internal IReadOnlyDictionary<string, int> DrawingAnchorEntriesByFlags { get; }

        /// <summary>Gets OfficeArt child anchors grouped by decoded rectangle.</summary>
        internal IReadOnlyDictionary<string, int> DrawingChildAnchorEntriesByRectangle { get; }

        /// <summary>Gets OfficeArt child anchors grouped by decoded width and height.</summary>
        internal IReadOnlyDictionary<string, int> DrawingChildAnchorEntriesBySize { get; }

        /// <summary>Gets preserve-only drawing and object BIFF records grouped by workbook or sheet location.</summary>
        internal IReadOnlyDictionary<string, int> DrawingRecordsByLocation { get; }

        /// <summary>Gets preserve-only compound feature records grouped by kind.</summary>
        internal IReadOnlyDictionary<LegacyXlsCompoundFeatureRecordKind, int> CompoundFeatureRecordsByKind { get; }

        /// <summary>Gets matching compound feature entries grouped by feature kind.</summary>
        internal IReadOnlyDictionary<LegacyXlsCompoundFeatureRecordKind, int> CompoundFeatureEntriesByKind { get; }

        /// <summary>Gets matching compound feature entries grouped by compound entry path or name.</summary>
        internal IReadOnlyDictionary<string, int> CompoundFeatureEntriesByName { get; }

        /// <summary>Gets matching compound feature entries grouped by preserve-only entry role.</summary>
        internal IReadOnlyDictionary<string, int> CompoundFeatureEntriesByRole { get; }

        /// <summary>Gets matching compound feature entries grouped by feature kind and entry role.</summary>
        internal IReadOnlyDictionary<string, int> CompoundFeatureEntriesByKindAndRole { get; }

        /// <summary>Gets matching compound feature entries grouped by OLE compound object type.</summary>
        internal IReadOnlyDictionary<string, int> CompoundFeatureEntriesByObjectType { get; }

        /// <summary>Gets matching compound feature entries grouped by role and OLE compound object type.</summary>
        internal IReadOnlyDictionary<string, int> CompoundFeatureEntriesByRoleAndObjectType { get; }

        /// <summary>Gets matching compound feature entries grouped by preserve-only content shape.</summary>
        internal IReadOnlyDictionary<string, int> CompoundFeatureEntriesByContentKind { get; }

        /// <summary>Gets matching compound feature entries grouped by role and preserve-only content shape.</summary>
        internal IReadOnlyDictionary<string, int> CompoundFeatureEntriesByRoleAndContentKind { get; }

        /// <summary>Gets matching compound feature entries grouped by declared byte size.</summary>
        internal IReadOnlyDictionary<string, int> CompoundFeatureEntriesBySize { get; }

        /// <summary>Gets matching compound feature entries grouped by role and declared byte size.</summary>
        internal IReadOnlyDictionary<string, int> CompoundFeatureEntriesByRoleAndSize { get; }

        /// <summary>Gets VBA module streams grouped by module name.</summary>
        internal IReadOnlyDictionary<string, int> CompoundVbaModulesByName { get; }

        /// <summary>Gets VBA module streams grouped by compound entry path.</summary>
        internal IReadOnlyDictionary<string, int> CompoundVbaModulesByPath { get; }

        /// <summary>Gets VBA module streams grouped by declared byte size.</summary>
        internal IReadOnlyDictionary<string, int> CompoundVbaModulesBySize { get; }

        /// <summary>Gets VBA module streams grouped by module name and declared byte size.</summary>
        internal IReadOnlyDictionary<string, int> CompoundVbaModulesByNameAndSize { get; }

        /// <summary>Gets VBA module streams grouped by preserve-only content shape.</summary>
        internal IReadOnlyDictionary<string, int> CompoundVbaModulesByContentKind { get; }

        /// <summary>Gets VBA module streams grouped by module name and preserve-only content shape.</summary>
        internal IReadOnlyDictionary<string, int> CompoundVbaModulesByNameAndContentKind { get; }

        /// <summary>Gets VBA module streams grouped by whether they match workbook or worksheet CodeName records.</summary>
        internal IReadOnlyDictionary<string, int> CompoundVbaModulesByCodeNameMatch { get; }

        /// <summary>Gets VBA module streams grouped by CodeName match type and module name.</summary>
        internal IReadOnlyDictionary<string, int> CompoundVbaModulesByCodeNameMatchAndName { get; }

        /// <summary>Gets VBA project compound features grouped by discovered module count.</summary>
        internal IReadOnlyDictionary<string, int> CompoundVbaProjectsByModuleCount { get; }

        /// <summary>Gets VBA project compound features grouped by total declared module stream bytes.</summary>
        internal IReadOnlyDictionary<string, int> CompoundVbaProjectsByModuleByteCount { get; }

        /// <summary>Gets VBA project compound features grouped by module, dir stream, and project stream counts.</summary>
        internal IReadOnlyDictionary<string, int> CompoundVbaProjectsByStructure { get; }

        /// <summary>Gets VBA project workbook states grouped by marker, compound storage, and module presence.</summary>
        internal IReadOnlyDictionary<string, int> VbaProjectWorkbookStates { get; }

        /// <summary>Gets parsed calculation setting records grouped by setting kind.</summary>
        internal IReadOnlyDictionary<LegacyXlsCalculationSettingKind, int> CalculationSettingsByKind { get; }

        /// <summary>Gets parsed workbook cell styles grouped by built-in/custom kind.</summary>
        internal IReadOnlyDictionary<string, int> CellStylesByKind { get; }

        /// <summary>Gets preserve-only style extension records grouped by BIFF record name.</summary>
        internal IReadOnlyDictionary<string, int> CellStyleExtensionsByRecordName { get; }

        /// <summary>Gets preserve-only style extension records grouped by extended XF index.</summary>
        internal IReadOnlyDictionary<string, int> CellStyleExtensionsByFormatIndex { get; }

        /// <summary>Gets preserve-only style extension records grouped by declared extension-property count.</summary>
        internal IReadOnlyDictionary<string, int> CellStyleExtensionsByExtensionCount { get; }

        /// <summary>Gets StyleExt records grouped by style category.</summary>
        internal IReadOnlyDictionary<string, int> CellStyleExtensionsByStyleCategory { get; }

        /// <summary>Gets StyleExt records grouped by declared flag state.</summary>
        internal IReadOnlyDictionary<string, int> CellStyleExtensionsByStyleFlags { get; }

        /// <summary>Gets StyleExt records grouped by style name.</summary>
        internal IReadOnlyDictionary<string, int> CellStyleExtensionsByStyleName { get; }

        /// <summary>Gets XFCRC records grouped by declared XF record count.</summary>
        internal IReadOnlyDictionary<string, int> CellStyleExtensionsByXfRecordCount { get; }

        /// <summary>Gets XFCRC records grouped by declared checksum.</summary>
        internal IReadOnlyDictionary<string, int> CellStyleExtensionsByChecksum { get; }

        /// <summary>Gets XFExt properties grouped by raw ExtProp type identifier.</summary>
        internal IReadOnlyDictionary<string, int> CellStyleExtensionPropertiesByType { get; }

        /// <summary>Gets XFExt properties grouped by decoded ExtProp type name.</summary>
        internal IReadOnlyDictionary<string, int> CellStyleExtensionPropertiesByName { get; }

        /// <summary>Gets XFExt properties grouped by ExtProp data payload byte count.</summary>
        internal IReadOnlyDictionary<string, int> CellStyleExtensionPropertiesByDataByteCount { get; }

        /// <summary>Gets XFExt properties grouped by decoded simple numeric value.</summary>
        internal IReadOnlyDictionary<string, int> CellStyleExtensionPropertiesByNumericValue { get; }

        /// <summary>Gets XFExt properties grouped by decoded simple numeric value name.</summary>
        internal IReadOnlyDictionary<string, int> CellStyleExtensionPropertiesByNumericValueName { get; }

        /// <summary>Gets XFExt color properties grouped by decoded FullColorExt color type.</summary>
        internal IReadOnlyDictionary<string, int> CellStyleExtensionPropertiesByColorType { get; }

        /// <summary>Gets XFExt color properties grouped by FullColorExt tint/shade value.</summary>
        internal IReadOnlyDictionary<string, int> CellStyleExtensionPropertiesByColorTintShade { get; }

        /// <summary>Gets XFExt color properties grouped by raw FullColorExt color value.</summary>
        internal IReadOnlyDictionary<string, int> CellStyleExtensionPropertiesByColorValue { get; }

        /// <summary>Gets parsed workbook metadata records grouped by metadata kind.</summary>
        internal IReadOnlyDictionary<LegacyXlsWorkbookMetadataKind, int> WorkbookMetadataRecordsByKind { get; }

        /// <summary>Gets preserve-only workbook future metadata records grouped by metadata kind.</summary>
        internal IReadOnlyDictionary<string, int> WorkbookFutureMetadataRecordsByKind { get; }

        /// <summary>Gets preserve-only workbook future metadata records grouped by BIFF record type.</summary>
        internal IReadOnlyDictionary<string, int> WorkbookFutureMetadataRecordsByRecordType { get; }

        /// <summary>Gets preserve-only workbook future metadata records grouped by BIFF record name.</summary>
        internal IReadOnlyDictionary<string, int> WorkbookFutureMetadataRecordsByRecordName { get; }

        /// <summary>Gets preserve-only workbook future metadata records grouped by decoded header state.</summary>
        internal IReadOnlyDictionary<string, int> WorkbookFutureMetadataRecordsByHeaderState { get; }

        /// <summary>Gets preserve-only workbook future metadata records grouped by future-record header type.</summary>
        internal IReadOnlyDictionary<string, int> WorkbookFutureMetadataRecordsByHeaderRecordType { get; }

        /// <summary>Gets preserve-only workbook future metadata records grouped by future-record header flags.</summary>
        internal IReadOnlyDictionary<string, int> WorkbookFutureMetadataRecordsByHeaderFlags { get; }

        /// <summary>Gets preserve-only workbook future metadata records grouped by BIFF payload length.</summary>
        internal IReadOnlyDictionary<string, int> WorkbookFutureMetadataRecordsByPayloadLength { get; }

        /// <summary>Gets preserve-only workbook future metadata records grouped by body byte count after future-record header decoding.</summary>
        internal IReadOnlyDictionary<string, int> WorkbookFutureMetadataRecordsByBodyByteCount { get; }

        /// <summary>Gets parsed worksheet metadata records grouped by metadata kind.</summary>
        internal IReadOnlyDictionary<LegacyXlsWorksheetMetadataKind, int> WorksheetMetadataRecordsByKind { get; }

        /// <summary>Gets preserve-only worksheet future metadata records grouped by metadata kind.</summary>
        internal IReadOnlyDictionary<string, int> WorksheetFutureMetadataRecordsByKind { get; }

        /// <summary>Gets preserve-only worksheet future metadata records grouped by sheet name.</summary>
        internal IReadOnlyDictionary<string, int> WorksheetFutureMetadataRecordsBySheet { get; }

        /// <summary>Gets preserve-only worksheet future metadata records grouped by sheet and metadata kind.</summary>
        internal IReadOnlyDictionary<string, int> WorksheetFutureMetadataRecordsBySheetAndKind { get; }

        /// <summary>Gets preserve-only worksheet future metadata records grouped by BIFF record type.</summary>
        internal IReadOnlyDictionary<string, int> WorksheetFutureMetadataRecordsByRecordType { get; }

        /// <summary>Gets preserve-only worksheet future metadata records grouped by BIFF record name.</summary>
        internal IReadOnlyDictionary<string, int> WorksheetFutureMetadataRecordsByRecordName { get; }

        /// <summary>Gets preserve-only worksheet future metadata records grouped by decoded header state.</summary>
        internal IReadOnlyDictionary<string, int> WorksheetFutureMetadataRecordsByHeaderState { get; }

        /// <summary>Gets preserve-only worksheet future metadata records grouped by future-record header type.</summary>
        internal IReadOnlyDictionary<string, int> WorksheetFutureMetadataRecordsByHeaderRecordType { get; }

        /// <summary>Gets preserve-only worksheet future metadata records grouped by future-record header flags.</summary>
        internal IReadOnlyDictionary<string, int> WorksheetFutureMetadataRecordsByHeaderFlags { get; }

        /// <summary>Gets preserve-only worksheet future metadata records grouped by BIFF payload length.</summary>
        internal IReadOnlyDictionary<string, int> WorksheetFutureMetadataRecordsByPayloadLength { get; }

        /// <summary>Gets preserve-only worksheet future metadata records grouped by body byte count after future-record header decoding.</summary>
        internal IReadOnlyDictionary<string, int> WorksheetFutureMetadataRecordsByBodyByteCount { get; }

        /// <summary>Gets parsed unsupported-sheet metadata records grouped by metadata kind.</summary>
        internal IReadOnlyDictionary<LegacyXlsUnsupportedSheetMetadataKind, int> UnsupportedSheetMetadataRecordsByKind { get; }

        /// <summary>Gets preserve-only unsupported-sheet future metadata records grouped by metadata kind.</summary>
        internal IReadOnlyDictionary<string, int> UnsupportedSheetFutureMetadataRecordsByKind { get; }

        /// <summary>Gets preserve-only unsupported-sheet future metadata records grouped by sheet name.</summary>
        internal IReadOnlyDictionary<string, int> UnsupportedSheetFutureMetadataRecordsBySheet { get; }

        /// <summary>Gets preserve-only unsupported-sheet future metadata records grouped by sheet and metadata kind.</summary>
        internal IReadOnlyDictionary<string, int> UnsupportedSheetFutureMetadataRecordsBySheetAndKind { get; }

        /// <summary>Gets preserve-only unsupported-sheet future metadata records grouped by BIFF record type.</summary>
        internal IReadOnlyDictionary<string, int> UnsupportedSheetFutureMetadataRecordsByRecordType { get; }

        /// <summary>Gets preserve-only unsupported-sheet future metadata records grouped by BIFF record name.</summary>
        internal IReadOnlyDictionary<string, int> UnsupportedSheetFutureMetadataRecordsByRecordName { get; }

        /// <summary>Gets preserve-only unsupported-sheet future metadata records grouped by decoded header state.</summary>
        internal IReadOnlyDictionary<string, int> UnsupportedSheetFutureMetadataRecordsByHeaderState { get; }

        /// <summary>Gets preserve-only unsupported-sheet future metadata records grouped by future-record header type.</summary>
        internal IReadOnlyDictionary<string, int> UnsupportedSheetFutureMetadataRecordsByHeaderRecordType { get; }

        /// <summary>Gets preserve-only unsupported-sheet future metadata records grouped by future-record header flags.</summary>
        internal IReadOnlyDictionary<string, int> UnsupportedSheetFutureMetadataRecordsByHeaderFlags { get; }

        /// <summary>Gets preserve-only unsupported-sheet future metadata records grouped by BIFF payload length.</summary>
        internal IReadOnlyDictionary<string, int> UnsupportedSheetFutureMetadataRecordsByPayloadLength { get; }

        /// <summary>Gets preserve-only unsupported-sheet future metadata records grouped by body byte count after future-record header decoding.</summary>
        internal IReadOnlyDictionary<string, int> UnsupportedSheetFutureMetadataRecordsByBodyByteCount { get; }

        /// <summary>Gets preserved feature record counts grouped by feature kind.</summary>
        internal IReadOnlyDictionary<LegacyXlsUnsupportedFeatureKind, int> PreservedFeatureRecordsByKind { get; }

        /// <summary>Gets preserved feature record counts grouped by kind, code, and stable feature subtype.</summary>
        internal IReadOnlyDictionary<string, int> PreservedFeatureRecordsByDetail { get; }

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
            builder.AppendLine($"Chart sheets: {ChartSheetCount}");
            builder.AppendLine($"Unsupported sheets: {UnsupportedSheetCount}");
            builder.AppendLine($"Cells: {CellCount}");
            builder.AppendLine($"Formula cells: {FormulaCellCount}");
            builder.AppendLine($"Comments: {CommentCount}");
            builder.AppendLine($"Hyperlinks: {HyperlinkCount}");
            builder.AppendLine($"Data validations: {DataValidationCount}");
            builder.AppendLine($"Data validation collection headers: {DataValidationCollectionRecordCount}");
            builder.AppendLine($"Conditional formatting rules: {ConditionalFormattingCount}");
            builder.AppendLine($"Conditional formatting extension records: {ConditionalFormattingExtensionRecordCount}");
            builder.AppendLine($"AutoFilter criteria columns: {AutoFilterCriteriaCount}");
            builder.AppendLine($"Defined names: {DefinedNameCount}");
            builder.AppendLine($"External references: {ExternalReferenceCount}");
            builder.AppendLine($"External sheet names: {ExternalSheetNameCount}");
            builder.AppendLine($"External names: {ExternalNameCount}");
            builder.AppendLine($"External cell caches: {ExternalCellCacheCount}");
            builder.AppendLine($"External cached cells: {ExternalCachedCellCount}");
            builder.AppendLine($"External query connections: {ExternalQueryConnectionCount}");
            builder.AppendLine($"Data consolidation references: {DataConsolidationReferenceCount}");
            builder.AppendLine($"Data consolidation named sources: {DataConsolidationNameCount}");
            builder.AppendLine($"Pivot table records: {PivotTableRecordCount}");
            builder.AppendLine($"Chart records: {ChartRecordCount}");
            builder.AppendLine($"Chart sheet metadata records: {ChartSheetMetadataRecordCount}");
            builder.AppendLine($"Chart sheet future metadata records: {ChartSheetFutureMetadataRecordCount}");
            builder.AppendLine($"Drawing records: {DrawingRecordCount}");
            builder.AppendLine($"Theme records: {ThemeRecordCount}");
            builder.AppendLine($"Drawing OfficeArt records: {DrawingOfficeArtRecordCount}");
            builder.AppendLine($"Drawing group blocks: {DrawingGroupBlockCount}");
            builder.AppendLine($"Drawing group infos: {DrawingGroupInfoCount}");
            builder.AppendLine($"Drawing identifier clusters: {DrawingIdentifierClusterCount}");
            builder.AppendLine($"Drawing shape properties: {DrawingShapePropertyCount}");
            builder.AppendLine($"Differential formats: {DifferentialFormatCount}");
            builder.AppendLine($"Table style collection records: {TableStyleCollectionRecordCount}");
            builder.AppendLine($"Table style definitions: {TableStyleDefinitionCount}");
            builder.AppendLine($"Table style element records: {TableStyleElementRecordCount}");
            builder.AppendLine($"Compound feature records: {CompoundFeatureRecordCount}");
            builder.AppendLine($"Compound feature entries: {CompoundFeatureEntryCount}");
            builder.AppendLine($"Compound VBA modules: {CompoundVbaModuleCount}");
            builder.AppendLine($"Compound feature entry bytes: {CompoundFeatureEntryByteCount}");
            builder.AppendLine($"Compound VBA module bytes: {CompoundVbaModuleByteCount}");
            builder.AppendLine($"Calculation setting records: {CalculationSettingRecordCount}");
            builder.AppendLine($"Cell style records: {CellStyleRecordCount}");
            builder.AppendLine($"Cell style extension records: {CellStyleExtensionRecordCount}");
            builder.AppendLine($"Formula token records: {FormulaTokenRecordCount}");
            builder.AppendLine($"Array formula records: {ArrayFormulaRecordCount}");
            builder.AppendLine($"Future function aliases: {FutureFunctionAliasCount}");
            builder.AppendLine($"Workbook metadata records: {WorkbookMetadataRecordCount}");
            builder.AppendLine($"Workbook future metadata records: {WorkbookFutureMetadataRecordCount}");
            builder.AppendLine($"Worksheet metadata records: {WorksheetMetadataRecordCount}");
            builder.AppendLine($"Worksheet future metadata records: {WorksheetFutureMetadataRecordCount}");
            builder.AppendLine($"Unsupported sheet metadata records: {UnsupportedSheetMetadataRecordCount}");
            builder.AppendLine($"Unsupported sheet future metadata records: {UnsupportedSheetFutureMetadataRecordCount}");
            builder.AppendLine($"Unsupported features: {UnsupportedFeatureCount}");
            builder.AppendLine($"Unsupported projection gaps: {UnsupportedProjectionGapCount}");
            builder.AppendLine($"Preserved feature records: {PreservedFeatureRecordCount}");
            builder.AppendLine($"Errors: {ErrorCount}");
            builder.AppendLine($"Warnings: {WarningCount}");
            AppendDictionary(builder, "Diagnostics By Code", DiagnosticsByCode);
            AppendDictionary(builder, "Formula Token Blockers", FormulaTokenBlockers);
            AppendDictionary(builder, "Formula Token Blockers By Token", FormulaTokenBlockersByToken);
            AppendDictionary(builder, "Formula Token Blockers By Token Name", FormulaTokenBlockersByTokenName);
            AppendDictionary(builder, "Formula Token Blockers By Offset", FormulaTokenBlockersByOffset);
            AppendDictionary(builder, "Formula Token Blockers By Sheet", FormulaTokenBlockersBySheet);
            AppendDictionary(builder, "Formula Token Blockers By Context", FormulaTokenBlockersByContext);
            AppendDictionary(builder, "Formula Token Blockers By Context And Token", FormulaTokenBlockersByContextAndToken);
            AppendDictionary(builder, "Formula Token Blockers By Context And Token Name", FormulaTokenBlockersByContextAndTokenName);
            AppendDictionary(builder, "Formula Token Blockers By Context And Detail", FormulaTokenBlockersByContextAndDetail);
            AppendDictionary(builder, "Formula Tokens By Name", FormulaTokensByName);
            AppendDictionary(builder, "Formula Tokens By Context", FormulaTokensByContext);
            AppendDictionary(builder, "Formula Tokens By Sheet", FormulaTokensBySheet);
            AppendDictionary(builder, "Formula Tokens By Context And Sheet", FormulaTokensByContextAndSheet);
            AppendDictionary(builder, "Formula Tokens By Context And Operand Kind", FormulaTokensByContextAndOperandKind);
            AppendDictionary(builder, "Formula Tokens By Record Type", FormulaTokensByRecordType);
            AppendDictionary(builder, "Formula Tokens By Class", FormulaTokensByClass);
            AppendDictionary(builder, "Formula Tokens By Name And Class", FormulaTokensByNameAndClass);
            AppendDictionary(builder, "Formula Tokens By Operand Byte Count", FormulaTokensByOperandByteCount);
            AppendDictionary(builder, "Formula Tokens By Operand Kind", FormulaTokensByOperandKind);
            AppendDictionary(builder, "Formula Tokens By Name And Operand Kind", FormulaTokensByNameAndOperandKind);
            AppendDictionary(builder, "Formula Tokens By Operand Kind And Text", FormulaTokensByOperandKindAndText);
            AppendDictionary(builder, "Formula Tokens By Name And Operand Text", FormulaTokensByNameAndOperandText);
            AppendDictionary(builder, "Formula Tokens By Sequence Index", FormulaTokensBySequenceIndex);
            AppendDictionary(builder, "Formula Functions By Id", FormulaFunctionsById);
            AppendDictionary(builder, "Formula Functions By Name", FormulaFunctionsByName);
            AppendDictionary(builder, "Formula Functions By Parameter Count", FormulaFunctionsByParameterCount);
            AppendDictionary(builder, "Formula Functions By Cetab State", FormulaFunctionsByCetabState);
            AppendDictionary(builder, "Formula Attributes By Name", FormulaAttributesByName);
            AppendDictionary(builder, "Array Formulas By Sheet", ArrayFormulasBySheet);
            AppendDictionary(builder, "Array Formulas By Range", ArrayFormulasByRange);
            AppendDictionary(builder, "Array Formulas By Sheet And Range", ArrayFormulasBySheetAndRange);
            AppendDictionary(builder, "Array Formulas By Declared Cell Count", ArrayFormulasByDeclaredCellCount);
            AppendDictionary(builder, "Array Formulas By Matched Formula Cell Count", ArrayFormulasByMatchedFormulaCellCount);
            AppendDictionary(builder, "Array Formulas By Always Calculate State", ArrayFormulasByAlwaysCalculateState);
            AppendDictionary(builder, "Array Formulas By Projection State", ArrayFormulasByProjectionState);
            AppendDictionary(builder, "Array Formulas By Token Byte Count", ArrayFormulasByTokenByteCount);
            AppendDictionary(builder, "Array Formulas By Extra Byte Count", ArrayFormulasByExtraByteCount);
            AppendDictionary(builder, "Future Function Aliases By Name", FutureFunctionAliasesByName);
            AppendDictionary(builder, "Future Function Aliases By Function", FutureFunctionAliasesByFunction);
            AppendDictionary(builder, "Future Function Aliases By Token Name", FutureFunctionAliasesByTokenName);
            AppendDictionary(builder, "Worksheet Feature States", WorksheetFeatureStates);
            AppendDictionary(builder, "Worksheet Protection Object States", WorksheetProtectionObjectStates);
            AppendDictionary(builder, "Worksheet Protection Scenario States", WorksheetProtectionScenarioStates);
            AppendDictionary(builder, "Data Validations By Type", DataValidationsByType);
            AppendDictionary(builder, "Data Validations By Operator", DataValidationsByOperator);
            AppendDictionary(builder, "Data Validations By Error Style", DataValidationsByErrorStyle);
            AppendDictionary(builder, "Data Validations By Allow Blank State", DataValidationsByAllowBlankState);
            AppendDictionary(builder, "Data Validations By Input Message State", DataValidationsByInputMessageState);
            AppendDictionary(builder, "Data Validations By Error Message State", DataValidationsByErrorMessageState);
            AppendDictionary(builder, "Data Validations By Prompt Text State", DataValidationsByPromptTextState);
            AppendDictionary(builder, "Data Validations By Error Text State", DataValidationsByErrorTextState);
            AppendDictionary(builder, "Data Validations By Drop Down State", DataValidationsByDropDownState);
            AppendDictionary(builder, "Data Validation Collections By Sheet", DataValidationCollectionsBySheet);
            AppendDictionary(builder, "Data Validation Collections By Declared Count", DataValidationCollectionsByDeclaredCount);
            AppendDictionary(builder, "Data Validation Collection States", DataValidationCollectionStates);
            AppendDictionary(builder, "Data Validations By Sheet", DataValidationsBySheet);
            AppendDictionary(builder, "Data Validations By Range Count", DataValidationsByRangeCount);
            AppendDictionary(builder, "Data Validations By Range", DataValidationsByRange);
            AppendDictionary(builder, "Data Validations By Sheet And Range", DataValidationsBySheetAndRange);
            AppendDictionary(builder, "Data Validations By Formula1 State", DataValidationsByFormula1State);
            AppendDictionary(builder, "Data Validations By Formula2 State", DataValidationsByFormula2State);
            AppendDictionary(builder, "Data Validations By Formula Pair State", DataValidationsByFormulaPairState);
            AppendDictionary(builder, "Data Validation List Sources By Kind", DataValidationListSourcesByKind);
            AppendDictionary(builder, "Data Validation List Sources By Item Count", DataValidationListSourcesByItemCount);
            AppendDictionary(builder, "Data Validation List Sources By Range", DataValidationListSourcesByRange);
            AppendDictionary(builder, "Data Validation List Sources By Name", DataValidationListSourcesByName);
            AppendDictionary(builder, "Data Validation List Sources By Sheet Name", DataValidationListSourcesBySheetName);
            AppendDictionary(builder, "Comments By Object Type", CommentsByObjectType);
            AppendDictionary(builder, "Comments By Object Type Name", CommentsByObjectTypeName);
            AppendDictionary(builder, "Comments By Object Flags", CommentsByObjectFlags);
            AppendDictionary(builder, "Comments By Object Flag Name", CommentsByObjectFlagName);
            AppendDictionary(builder, "Comments By Anchor Range", CommentsByAnchorRange);
            AppendDictionary(builder, "Comments By Anchor Offset", CommentsByAnchorOffset);
            AppendDictionary(builder, "Comments By Anchor Flags", CommentsByAnchorFlags);
            AppendDictionary(builder, "Conditional Formatting By Type", ConditionalFormattingsByType);
            AppendDictionary(builder, "Conditional Formatting By Operator", ConditionalFormattingsByOperator);
            AppendDictionary(builder, "Conditional Formatting By Sheet", ConditionalFormattingsBySheet);
            AppendDictionary(builder, "Conditional Formatting By Range Count", ConditionalFormattingsByRangeCount);
            AppendDictionary(builder, "Conditional Formatting By Range", ConditionalFormattingsByRange);
            AppendDictionary(builder, "Conditional Formatting By Sheet And Range", ConditionalFormattingsBySheetAndRange);
            AppendDictionary(builder, "Conditional Formatting By Formula1 State", ConditionalFormattingsByFormula1State);
            AppendDictionary(builder, "Conditional Formatting By Formula2 State", ConditionalFormattingsByFormula2State);
            AppendDictionary(builder, "Conditional Formatting By Formula Pair State", ConditionalFormattingsByFormulaPairState);
            AppendDictionary(builder, "Conditional Formatting By Priority State", ConditionalFormattingsByPriorityState);
            AppendDictionary(builder, "Conditional Formatting By Priority", ConditionalFormattingsByPriority);
            AppendDictionary(builder, "Conditional Formatting By Stop If True State", ConditionalFormattingsByStopIfTrueState);
            AppendDictionary(builder, "Conditional Formatting By Differential Format State", ConditionalFormattingsByDifferentialFormatState);
            AppendDictionary(builder, "Conditional Formatting By Differential Fill", ConditionalFormattingsByDifferentialFill);
            AppendDictionary(builder, "Conditional Formatting By Differential Font", ConditionalFormattingsByDifferentialFont);
            AppendDictionary(builder, "Conditional Formatting By Differential Border", ConditionalFormattingsByDifferentialBorder);
            AppendDictionary(builder, "Conditional Formatting By Differential Number Format", ConditionalFormattingsByDifferentialNumberFormat);
            AppendDictionary(builder, "Conditional Formatting Extensions By Sheet", ConditionalFormattingExtensionsBySheet);
            AppendDictionary(builder, "Conditional Formatting Extensions By Record Type", ConditionalFormattingExtensionsByRecordType);
            AppendDictionary(builder, "Conditional Formatting Extension States", ConditionalFormattingExtensionStates);
            AppendDictionary(builder, "Conditional Formatting Extension Priorities", ConditionalFormattingExtensionPriorities);
            AppendDictionary(builder, "Conditional Formatting Extension Stop If True States", ConditionalFormattingExtensionStopIfTrueStates);
            AppendDictionary(builder, "Conditional Formatting Extension Inline Formatting Byte Counts", ConditionalFormattingExtensionInlineFormattingByteCounts);
            AppendDictionary(builder, "Conditional Formatting Extension Dxf Projection States", ConditionalFormattingExtensionDxfProjectionStates);
            AppendDictionary(builder, "Differential Formats By Record Type", DifferentialFormatsByRecordType);
            AppendDictionary(builder, "Differential Formats By Content State", DifferentialFormatsByContentState);
            AppendDictionary(builder, "Differential Formats By Fill", DifferentialFormatsByFill);
            AppendDictionary(builder, "Differential Formats By Font", DifferentialFormatsByFont);
            AppendDictionary(builder, "Differential Formats By Border", DifferentialFormatsByBorder);
            AppendDictionary(builder, "Differential Formats By Number Format", DifferentialFormatsByNumberFormat);
            AppendDictionary(builder, "Table Style Collections By Default Table Style", TableStyleCollectionsByDefaultTableStyle);
            AppendDictionary(builder, "Table Style Collections By Default Pivot Style", TableStyleCollectionsByDefaultPivotStyle);
            AppendDictionary(builder, "Table Style Collections By Total Style Count", TableStyleCollectionsByTotalStyleCount);
            AppendDictionary(builder, "Table Styles By Name", TableStylesByName);
            AppendDictionary(builder, "Table Styles By Applicability", TableStylesByApplicability);
            AppendDictionary(builder, "Table Styles By Declared Element Count", TableStylesByDeclaredElementCount);
            AppendDictionary(builder, "Table Styles By Parsed Element Count", TableStylesByParsedElementCount);
            AppendDictionary(builder, "Table Style Elements By Type", TableStyleElementsByType);
            AppendDictionary(builder, "Table Style Elements By Differential Format Index", TableStyleElementsByDifferentialFormatIndex);
            AppendDictionary(builder, "Table Style Elements By Stripe Size", TableStyleElementsByStripeSize);
            AppendDictionary(builder, "AutoFilter Criteria By Sheet", AutoFilterCriteriaBySheet);
            AppendDictionary(builder, "AutoFilter Criteria By Kind", AutoFilterCriteriaByKind);
            AppendDictionary(builder, "AutoFilter Criteria By Operator", AutoFilterCriteriaByOperator);
            AppendDictionary(builder, "AutoFilter Criteria By Value Kind", AutoFilterCriteriaByValueKind);
            AppendDictionary(builder, "AutoFilter Criteria By Text Pattern", AutoFilterCriteriaByTextPattern);
            AppendDictionary(builder, "AutoFilter Criteria By Join Operator", AutoFilterCriteriaByJoinOperator);
            AppendDictionary(builder, "AutoFilter Criteria By Column", AutoFilterCriteriaByColumn);
            AppendDictionary(builder, "AutoFilter Criteria By Sheet And Column", AutoFilterCriteriaBySheetAndColumn);
            AppendDictionary(builder, "AutoFilter Criteria By Condition Count", AutoFilterCriteriaByConditionCount);
            AppendDictionary(builder, "AutoFilter Top10 Kinds", AutoFilterTop10Kinds);
            AppendDictionary(builder, "AutoFilter Top10 Values", AutoFilterTop10Values);
            AppendDictionary(builder, "AutoFilter Top10 Directions", AutoFilterTop10Directions);
            AppendDictionary(builder, "AutoFilter Top10 Units", AutoFilterTop10Units);
            AppendDictionary(builder, "Worksheet Phonetic Settings By Sheet", WorksheetPhoneticSettingsBySheet);
            AppendDictionary(builder, "Worksheet Phonetic Settings By Type", WorksheetPhoneticSettingsByType);
            AppendDictionary(builder, "Worksheet Phonetic Settings By Alignment", WorksheetPhoneticSettingsByAlignment);
            AppendDictionary(builder, "Worksheet Phonetic Settings By Font Id", WorksheetPhoneticSettingsByFontId);
            AppendDictionary(builder, "Worksheet Phonetic Settings By Range Count", WorksheetPhoneticSettingsByRangeCount);
            AppendDictionary(builder, "Worksheet Phonetic Ranges By Sheet", WorksheetPhoneticRangesBySheet);
            AppendDictionary(builder, "Worksheet Phonetic Ranges By Sheet And Range", WorksheetPhoneticRangesBySheetAndRange);
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
            AppendDictionary(builder, "Unsupported Projection Gaps By Kind", UnsupportedProjectionGapsByKind.ToDictionary(
                entry => entry.Key.ToString(),
                entry => entry.Value,
                StringComparer.OrdinalIgnoreCase));
            AppendDictionary(builder, "Unsupported Projection Gap Record Types", UnsupportedProjectionGapsByRecordType);
            AppendDictionary(builder, "Unsupported Projection Gap Details", UnsupportedProjectionGapsByDetail);
            AppendDictionary(builder, "File Format States", FileFormatStates);
            AppendDictionary(builder, "File Format Blockers", FileFormatBlockers);
            AppendDictionary(builder, "File Format Blockers By Record Type", FileFormatBlockersByRecordType);
            AppendDictionary(builder, "File Format Blockers By Record Name", FileFormatBlockersByRecordName);
            AppendDictionary(builder, "File Format Blockers By Location", FileFormatBlockersByLocation);
            AppendDictionary(builder, "Encrypted Workbooks By Method", EncryptedWorkbooksByMethod);
            AppendDictionary(builder, "Unsupported BIFF Versions By Version", UnsupportedBiffVersionsByVersion);
            AppendDictionary(builder, "Unsupported BIFF Versions By Substream", UnsupportedBiffVersionsBySubstream);
            AppendDictionary(builder, "Unsupported BIFF Versions By Version And Substream", UnsupportedBiffVersionsByVersionAndSubstream);
            AppendDictionary(builder, "Unsupported Sheets By Kind", UnsupportedSheetsByKind.ToDictionary(
                entry => entry.Key.ToString(),
                entry => entry.Value,
                StringComparer.OrdinalIgnoreCase));
            AppendDictionary(builder, "Unsupported Sheets By Type", UnsupportedSheetsByType);
            AppendDictionary(builder, "Unsupported Sheets By Name", UnsupportedSheetsByName);
            AppendDictionary(builder, "Unsupported Sheets By Visibility", UnsupportedSheetsByVisibility);
            AppendDictionary(builder, "Unsupported Sheets By Kind And Visibility", UnsupportedSheetsByKindAndVisibility);
            AppendDictionary(builder, "Chart Sheets By Type", ChartSheetsByType);
            AppendDictionary(builder, "Chart Sheets By Name", ChartSheetsByName);
            AppendDictionary(builder, "Chart Sheets By Visibility", ChartSheetsByVisibility);
            AppendDictionary(builder, "Chart Sheet Metadata Records By Kind", ChartSheetMetadataRecordsByKind);
            AppendDictionary(builder, "Chart Sheet Future Metadata Records By Record Type", ChartSheetFutureMetadataRecordsByRecordType);
            AppendDictionary(builder, "Chart Sheet Print Sizes", ChartSheetPrintSizes);
            AppendDictionary(builder, "Chart Sheet Print Size Kinds", ChartSheetPrintSizeKinds);
            AppendDictionary(builder, "Chart Sheet Text Object Counts", ChartSheetTextObjectCounts);
            AppendDictionary(builder, "Chart Sheet Chart Record Counts", ChartSheetChartRecordCounts);
            AppendDictionary(builder, "Chart Sheet Chart Record Counts By Sheet", ChartSheetChartRecordCountsBySheet);
            AppendDictionary(builder, "Chart Sheet Chart Record Kinds", ChartSheetChartRecordKinds);
            AppendDictionary(builder, "Chart Sheet Chart Record Kinds By Sheet", ChartSheetChartRecordKindsBySheet);
            AppendDictionary(builder, "Chart Sheet Chart Types", ChartSheetChartTypes);
            AppendDictionary(builder, "Chart Sheet Chart Types By Sheet", ChartSheetChartTypesBySheet);
            AppendDictionary(builder, "Chart Sheet States", ChartSheetStates);
            AppendDictionary(builder, "Unsupported Chart Sheet Print Sizes", UnsupportedChartSheetPrintSizes);
            AppendDictionary(builder, "Unsupported Chart Sheet Print Size Kinds", UnsupportedChartSheetPrintSizeKinds);
            AppendDictionary(builder, "Unsupported Chart Sheet Text Object Counts", UnsupportedChartSheetTextObjectCounts);
            AppendDictionary(builder, "Unsupported Chart Sheet Chart Record Counts", UnsupportedChartSheetChartRecordCounts);
            AppendDictionary(builder, "Unsupported Chart Sheet Chart Record Counts By Sheet", UnsupportedChartSheetChartRecordCountsBySheet);
            AppendDictionary(builder, "Unsupported Chart Sheet Chart Record Kinds", UnsupportedChartSheetChartRecordKinds);
            AppendDictionary(builder, "Unsupported Chart Sheet Chart Record Kinds By Sheet", UnsupportedChartSheetChartRecordKindsBySheet);
            AppendDictionary(builder, "Unsupported Chart Sheet Chart Types", UnsupportedChartSheetChartTypes);
            AppendDictionary(builder, "Unsupported Chart Sheet Chart Types By Sheet", UnsupportedChartSheetChartTypesBySheet);
            AppendDictionary(builder, "Unsupported Chart Sheet States", UnsupportedChartSheetStates);
            AppendDictionary(builder, "External References By Kind", ExternalReferencesByKind.ToDictionary(
                entry => entry.Key.ToString(),
                entry => entry.Value,
                StringComparer.OrdinalIgnoreCase));
            AppendDictionary(builder, "External References By Target", ExternalReferencesByTarget);
            AppendDictionary(builder, "External References By Shape", ExternalReferencesByShape);
            AppendDictionary(builder, "External Reference Workbook States", ExternalReferenceWorkbookStates);
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
            AppendDictionary(builder, "External Sheet Names By Target", ExternalSheetNamesByTarget);
            AppendDictionary(builder, "External Names By Reference Kind", ExternalNamesByReferenceKind.ToDictionary(
                entry => entry.Key.ToString(),
                entry => entry.Value,
                StringComparer.OrdinalIgnoreCase));
            AppendDictionary(builder, "External Names By Name", ExternalNamesByName);
            AppendDictionary(builder, "External Names By Scope", ExternalNamesByScope);
            AppendDictionary(builder, "External Names By Built-In State", ExternalNamesByBuiltInState);
            AppendDictionary(builder, "External Names By Body Kind", ExternalNamesByBodyKind);
            AppendDictionary(builder, "External Names By Cached Clipboard Format", ExternalNamesByCachedClipboardFormat);
            AppendDictionary(builder, "External Names By Advise State", ExternalNamesByAdviseState);
            AppendDictionary(builder, "External Names By Picture State", ExternalNamesByPictureState);
            AppendDictionary(builder, "External Names By OLE State", ExternalNamesByOleState);
            AppendDictionary(builder, "External Names By OLE Link State", ExternalNamesByOleLinkState);
            AppendDictionary(builder, "External Names By Icon State", ExternalNamesByIconState);
            AppendDictionary(builder, "External Names By Flag Shape", ExternalNamesByFlagShape);
            AppendDictionary(builder, "External Cell Caches By Target", ExternalCellCachesByTarget);
            AppendDictionary(builder, "External Cell Caches By Sheet Name", ExternalCellCachesBySheetName);
            AppendDictionary(builder, "External Cell Caches By Target And Sheet Name", ExternalCellCachesByTargetAndSheetName);
            AppendDictionary(builder, "External Cell Caches By Cell Range", ExternalCellCachesByCellRange);
            AppendDictionary(builder, "External Cell Caches By Target And Cell Range", ExternalCellCachesByTargetAndCellRange);
            AppendDictionary(builder, "External Cell Caches By Cell Count", ExternalCellCachesByCellCount);
            AppendDictionary(builder, "External Cell Caches By Row Span", ExternalCellCachesByRowSpan);
            AppendDictionary(builder, "External Cell Caches By Column Span", ExternalCellCachesByColumnSpan);
            AppendDictionary(builder, "External Cell Caches By Link State", ExternalCellCachesByLinkState);
            AppendDictionary(builder, "External Cached Cells By Value Kind", ExternalCachedCellsByValueKind.ToDictionary(
                entry => entry.Key.ToString(),
                entry => entry.Value,
                StringComparer.OrdinalIgnoreCase));
            AppendDictionary(builder, "External Cached Cells By Target Sheet And Value Kind", ExternalCachedCellsByTargetSheetAndValueKind);
            AppendDictionary(builder, "External Query Connections By Source Type", ExternalQueryConnectionsBySourceType);
            AppendDictionary(builder, "External Query Connections By State", ExternalQueryConnectionsByState);
            AppendDictionary(builder, "External Query Connections By Connection Flag", ExternalQueryConnectionsByConnectionFlag);
            AppendDictionary(builder, "External Query Connections By Query Option", ExternalQueryConnectionsByQueryOption);
            AppendDictionary(builder, "External Query Connections By Parameter Flag Count", ExternalQueryConnectionsByParameterFlagCount);
            AppendDictionary(builder, "External Query Connections By Parameter Flag Byte Count", ExternalQueryConnectionsByParameterFlagByteCount);
            AppendDictionary(builder, "External Query Connections By Parameter Flag State", ExternalQueryConnectionsByParameterFlagState);
            AppendDictionary(builder, "External Query Connections By Future Byte Count", ExternalQueryConnectionsByFutureByteCount);
            AppendDictionary(builder, "External Query Connections By Refresh Interval", ExternalQueryConnectionsByRefreshInterval);
            AppendDictionary(builder, "External Query Connections By OleDb Connection Count", ExternalQueryConnectionsByOleDbConnectionCount);
            AppendDictionary(builder, "External Query Connections By Html Format", ExternalQueryConnectionsByHtmlFormat);
            AppendDictionary(builder, "External Query Connections By Version Triplet", ExternalQueryConnectionsByVersionTriplet);
            AppendDictionary(builder, "External Query Connections By Source Specific Flags", ExternalQueryConnectionsBySourceSpecificFlags);
            AppendDictionary(builder, "Data Consolidation References By Source Kind", DataConsolidationReferencesBySourceKind);
            AppendDictionary(builder, "Data Consolidation References By Source Prefix", DataConsolidationReferencesBySourcePrefix);
            AppendDictionary(builder, "Data Consolidation References By Source", DataConsolidationReferencesBySource);
            AppendDictionary(builder, "Data Consolidation References By Range", DataConsolidationReferencesByRange);
            AppendDictionary(builder, "Data Consolidation References By Shape", DataConsolidationReferencesByShape);
            AppendDictionary(builder, "Data Consolidation References By Source And Range", DataConsolidationReferencesBySourceAndRange);
            AppendDictionary(builder, "Data Consolidation References By Unused Byte Count", DataConsolidationReferencesByUnusedByteCount);
            AppendDictionary(builder, "Data Consolidation Names By Source Kind", DataConsolidationNamesBySourceKind);
            AppendDictionary(builder, "Data Consolidation Names By Name", DataConsolidationNamesByName);
            AppendDictionary(builder, "Data Consolidation Names By Source", DataConsolidationNamesBySource);
            AppendDictionary(builder, "Data Consolidation Names By Name And Source", DataConsolidationNamesByNameAndSource);
            AppendDictionary(builder, "Data Consolidation Names By Unused Byte Count", DataConsolidationNamesByUnusedByteCount);
            AppendDictionary(builder, "Theme Records By Version", ThemeRecordsByVersion);
            AppendDictionary(builder, "Theme Records By Raw Version", ThemeRecordsByRawVersion);
            AppendDictionary(builder, "Theme Records By Content State", ThemeRecordsByContentState);
            AppendDictionary(builder, "Theme Records By Content Length", ThemeRecordsByContentLength);
            AppendDictionary(builder, "Pivot Table Records By Kind", PivotTableRecordsByKind.ToDictionary(
                entry => entry.Key.ToString(),
                entry => entry.Value,
                StringComparer.OrdinalIgnoreCase));
            AppendDictionary(builder, "Pivot Table Records By Name", PivotTableRecordsByName);
            AppendDictionary(builder, "Pivot Table Records By Location", PivotTableRecordsByLocation);
            AppendDictionary(builder, "Pivot Table Records By Kind And Location", PivotTableRecordsByKindAndLocation);
            AppendDictionary(builder, "Pivot Table Records By Name And Location", PivotTableRecordsByNameAndLocation);
            AppendDictionary(builder, "Pivot Table Workbook States", PivotTableWorkbookStates);
            AppendDictionary(builder, "Pivot Table View Ranges", PivotTableViewRanges);
            AppendDictionary(builder, "Pivot Table View Names", PivotTableViewNames);
            AppendDictionary(builder, "Pivot Table View Data Names", PivotTableViewDataNames);
            AppendDictionary(builder, "Pivot Table View Field Counts", PivotTableViewFieldCounts);
            AppendDictionary(builder, "Pivot Table View Line Counts", PivotTableViewLineCounts);
            AppendDictionary(builder, "Pivot Table View Data Axes", PivotTableViewDataAxes);
            AppendDictionary(builder, "Pivot Table View Data Positions", PivotTableViewDataPositions);
            AppendDictionary(builder, "Pivot Table View Cache Indexes", PivotTableViewCacheIndexes);
            AppendDictionary(builder, "Pivot Table View Grand Total States", PivotTableViewGrandTotalStates);
            AppendDictionary(builder, "Pivot Table View AutoFormat States", PivotTableViewAutoFormatStates);
            AppendDictionary(builder, "Pivot Table Field Axes", PivotTableFieldAxes);
            AppendDictionary(builder, "Pivot Table Field Item Counts", PivotTableFieldItemCounts);
            AppendDictionary(builder, "Pivot Table Field Subtotal Counts", PivotTableFieldSubtotalCounts);
            AppendDictionary(builder, "Pivot Table Field Subtotal Functions", PivotTableFieldSubtotalFunctions);
            AppendDictionary(builder, "Pivot Table Field Names", PivotTableFieldNames);
            AppendDictionary(builder, "Pivot Table Field Index List Lengths", PivotTableFieldIndexListLengths);
            AppendDictionary(builder, "Pivot Table Field Index References", PivotTableFieldIndexReferences);
            AppendDictionary(builder, "Pivot Table Field Index Sequences", PivotTableFieldIndexSequences);
            AppendDictionary(builder, "Pivot Table Line Item Counts", PivotTableLineItemCounts);
            AppendDictionary(builder, "Pivot Table Line Item Types", PivotTableLineItemTypes);
            AppendDictionary(builder, "Pivot Table Line Item Type Kinds", PivotTableLineItemTypeKinds);
            AppendDictionary(builder, "Pivot Table Line Item Entry Counts", PivotTableLineItemEntryCounts);
            AppendDictionary(builder, "Pivot Table Line Item Entry Slot Counts", PivotTableLineItemEntrySlotCounts);
            AppendDictionary(builder, "Pivot Table Line Item Entry Indexes", PivotTableLineItemEntryIndexes);
            AppendDictionary(builder, "Pivot Table Line Item Data Indexes", PivotTableLineItemDataIndexes);
            AppendDictionary(builder, "Pivot Table Line Item Flag States", PivotTableLineItemFlagStates);
            AppendDictionary(builder, "Pivot Table Line Item Sequences", PivotTableLineItemSequences);
            AppendDictionary(builder, "Pivot Table Page Item Counts", PivotTablePageItemCounts);
            AppendDictionary(builder, "Pivot Table Page Item Field Indexes", PivotTablePageItemFieldIndexes);
            AppendDictionary(builder, "Pivot Table Page Item Indexes", PivotTablePageItemIndexes);
            AppendDictionary(builder, "Pivot Table Page Item Object Ids", PivotTablePageItemObjectIds);
            AppendDictionary(builder, "Pivot Table Page Item Sequences", PivotTablePageItemSequences);
            AppendDictionary(builder, "Pivot Table Item Types", PivotTableItemTypes);
            AppendDictionary(builder, "Pivot Table Item Type Kinds", PivotTableItemTypeKinds);
            AppendDictionary(builder, "Pivot Table Item Cache Indexes", PivotTableItemCacheIndexes);
            AppendDictionary(builder, "Pivot Table Item Flag States", PivotTableItemFlagStates);
            AppendDictionary(builder, "Pivot Table Item Names", PivotTableItemNames);
            AppendDictionary(builder, "Pivot Table Formula Payload Lengths", PivotTableFormulaPayloadLengths);
            AppendDictionary(builder, "Pivot Table Formula Payload Kinds", PivotTableFormulaPayloadKinds);
            AppendDictionary(builder, "Pivot Table Formula Token Byte Counts", PivotTableFormulaTokenByteCounts);
            AppendDictionary(builder, "Pivot Table Calculated Field Formula Token Byte Counts", PivotTableCalculatedFieldFormulaTokenByteCounts);
            AppendDictionary(builder, "Pivot Table Formula Trailing Byte Counts", PivotTableFormulaTrailingByteCounts);
            AppendDictionary(builder, "Pivot Table Rule Axes", PivotTableRuleAxes);
            AppendDictionary(builder, "Pivot Table Rule Types", PivotTableRuleTypes);
            AppendDictionary(builder, "Pivot Table Rule Field References", PivotTableRuleFieldReferences);
            AppendDictionary(builder, "Pivot Table Rule Filter Counts", PivotTableRuleFilterCounts);
            AppendDictionary(builder, "Pivot Table Rule Option States", PivotTableRuleOptionStates);
            AppendDictionary(builder, "Pivot Table Rule Partial Areas", PivotTableRulePartialAreas);
            AppendDictionary(builder, "Pivot Table Rule Filter Entry Counts", PivotTableRuleFilterEntryCounts);
            AppendDictionary(builder, "Pivot Table Rule Filter Axes", PivotTableRuleFilterAxes);
            AppendDictionary(builder, "Pivot Table Rule Filter Field Positions", PivotTableRuleFilterFieldPositions);
            AppendDictionary(builder, "Pivot Table Rule Filter Field References", PivotTableRuleFilterFieldReferences);
            AppendDictionary(builder, "Pivot Table Rule Filter Selected States", PivotTableRuleFilterSelectedStates);
            AppendDictionary(builder, "Pivot Table Rule Filter Subtotal Flags", PivotTableRuleFilterSubtotalFlags);
            AppendDictionary(builder, "Pivot Table Rule Filter Subtotal Functions", PivotTableRuleFilterSubtotalFunctions);
            AppendDictionary(builder, "Pivot Table Rule Filter Item Index Counts", PivotTableRuleFilterItemIndexCounts);
            AppendDictionary(builder, "Pivot Table Rule Filter States", PivotTableRuleFilterStates);
            AppendDictionary(builder, "Pivot Table Cache Item Kinds", PivotTableCacheItemKinds);
            AppendDictionary(builder, "Pivot Table Cache Item Value States", PivotTableCacheItemValueStates);
            AppendDictionary(builder, "Pivot Table Cache Item String Lengths", PivotTableCacheItemStringLengths);
            AppendDictionary(builder, "Pivot Table Cache Item Error Codes", PivotTableCacheItemErrorCodes);
            AppendDictionary(builder, "Pivot Table Cache Item Boolean Values", PivotTableCacheItemBooleanValues);
            AppendDictionary(builder, "Pivot Table Cache Stream Names", PivotTableCacheStreamNames);
            AppendDictionary(builder, "Pivot Table Cache Source Types", PivotTableCacheSourceTypes);
            AppendDictionary(builder, "Pivot Table Cache Record Counts", PivotTableCacheRecordCounts);
            AppendDictionary(builder, "Pivot Table Cache Field Counts", PivotTableCacheFieldCounts);
            AppendDictionary(builder, "Pivot Table Cache Used Record Counts", PivotTableCacheUsedRecordCounts);
            AppendDictionary(builder, "Pivot Table Cache Property Flags", PivotTableCachePropertyFlags);
            AppendDictionary(builder, "Pivot Table Cache Refresh User States", PivotTableCacheRefreshUserStates);
            AppendDictionary(builder, "Pivot Table Query Tag Targets", PivotTableQueryTagTargets);
            AppendDictionary(builder, "Pivot Table Query Tag Names", PivotTableQueryTagNames);
            AppendDictionary(builder, "Pivot Table Query Tag Refresh States", PivotTableQueryTagRefreshStates);
            AppendDictionary(builder, "Pivot Table Query Tag Versions", PivotTableQueryTagVersions);
            AppendDictionary(builder, "Pivot Table Query Tag Future Options", PivotTableQueryTagFutureOptions);
            AppendDictionary(builder, "Pivot Table Query Tag Unused Values", PivotTableQueryTagUnusedValues);
            AppendDictionary(builder, "Pivot Table Data Item Aggregations", PivotTableDataItemAggregations);
            AppendDictionary(builder, "Pivot Table Data Item Aggregation Kinds", PivotTableDataItemAggregationKinds);
            AppendDictionary(builder, "Pivot Table Data Item Field Indexes", PivotTableDataItemFieldIndexes);
            AppendDictionary(builder, "Pivot Table Data Item Display Calculation Ids", PivotTableDataItemDisplayCalculationIds);
            AppendDictionary(builder, "Pivot Table Data Item Display Calculations", PivotTableDataItemDisplayCalculations);
            AppendDictionary(builder, "Pivot Table Data Item Display Calculation Reference States", PivotTableDataItemDisplayCalculationReferenceStates);
            AppendDictionary(builder, "Pivot Table Data Item Display Calculation Field Indexes", PivotTableDataItemDisplayCalculationFieldIndexes);
            AppendDictionary(builder, "Pivot Table Data Item Display Calculation Item Indexes", PivotTableDataItemDisplayCalculationItemIndexes);
            AppendDictionary(builder, "Pivot Table Data Item Number Formats", PivotTableDataItemNumberFormats);
            AppendDictionary(builder, "Pivot Table Data Item Names", PivotTableDataItemNames);
            AppendDictionary(builder, "Pivot Table Grouping Kinds", PivotTableGroupingKinds);
            AppendDictionary(builder, "Pivot Table Grouping Boundary States", PivotTableGroupingBoundaryStates);
            AppendDictionary(builder, "Pivot Table Grouping Completion States", PivotTableGroupingCompletionStates);
            AppendDictionary(builder, "Pivot Table Grouping States", PivotTableGroupingStates);
            AppendDictionary(builder, "Pivot Table Grouping Numeric Ranges", PivotTableGroupingNumericRanges);
            AppendDictionary(builder, "Pivot Table Grouping Date Ranges", PivotTableGroupingDateRanges);
            AppendDictionary(builder, "Pivot Table Formula Scopes", PivotTableFormulaScopes);
            AppendDictionary(builder, "Pivot Table Formula Cache Field Indexes", PivotTableFormulaCacheFieldIndexes);
            AppendDictionary(builder, "Pivot Table Formula Reserved Values", PivotTableFormulaReservedValues);
            AppendDictionary(builder, "Pivot Table Extended Field States", PivotTableExtendedFieldStates);
            AppendDictionary(builder, "Pivot Table Extended Field Permission States", PivotTableExtendedFieldPermissionStates);
            AppendDictionary(builder, "Pivot Table Additional Classes", PivotTableAdditionalClasses);
            AppendDictionary(builder, "Pivot Table Additional Types", PivotTableAdditionalTypes);
            AppendDictionary(builder, "Pivot Table Additional Class Types", PivotTableAdditionalClassTypes);
            AppendDictionary(builder, "Pivot Table Additional Future Record Types", PivotTableAdditionalFutureRecordTypes);
            AppendDictionary(builder, "Pivot Table Additional Future Flags", PivotTableAdditionalFutureFlags);
            AppendDictionary(builder, "Pivot Table Additional Sequence Indexes", PivotTableAdditionalSequenceIndexes);
            AppendDictionary(builder, "Pivot Table Additional Payload Lengths By Class Type", PivotTableAdditionalPayloadLengthsByClassType);
            AppendDictionary(builder, "Pivot Table Additional Cache Ids", PivotTableAdditionalCacheIds);
            AppendDictionary(builder, "Pivot Table Additional Class Depths Before", PivotTableAdditionalClassDepthsBefore);
            AppendDictionary(builder, "Pivot Table Additional Class Depths After", PivotTableAdditionalClassDepthsAfter);
            AppendDictionary(builder, "Pivot Table Additional Class Transitions", PivotTableAdditionalClassTransitions);
            AppendDictionary(builder, "Pivot Table Additional Class Transitions By Class Type", PivotTableAdditionalClassTransitionsByClassType);
            AppendDictionary(builder, "Chart Records By Kind", ChartRecordsByKind.ToDictionary(
                entry => entry.Key.ToString(),
                entry => entry.Value,
                StringComparer.OrdinalIgnoreCase));
            AppendDictionary(builder, "Chart Records By Name", ChartRecordsByName);
            AppendDictionary(builder, "Chart Records By Name And Payload Length", ChartRecordsByNameAndPayloadLength);
            AppendDictionary(builder, "Chart Workbook States", ChartWorkbookStates);
            AppendDictionary(builder, "Chart Records By Container Depth Before", ChartRecordsByContainerDepthBefore);
            AppendDictionary(builder, "Chart Records By Container Depth After", ChartRecordsByContainerDepthAfter);
            AppendDictionary(builder, "Chart Records By Container Transition", ChartRecordsByContainerTransition);
            AppendDictionary(builder, "Chart Records By Name And Container Depth", ChartRecordsByNameAndContainerDepth);
            AppendDictionary(builder, "Chart Records By Name And Container Transition", ChartRecordsByNameAndContainerTransition);
            AppendDictionary(builder, "Chart Records By Chart Type", ChartRecordsByChartType);
            AppendDictionary(builder, "Chart Records By Rectangle", ChartRecordsByRectangle);
            AppendDictionary(builder, "Chart Records By Axis Type", ChartRecordsByAxisType);
            AppendDictionary(builder, "Chart Group Varied Color States", ChartGroupVariedColorStates);
            AppendDictionary(builder, "Chart Group Drawing Orders", ChartGroupDrawingOrders);
            AppendDictionary(builder, "Chart Records By Axes Used Count", ChartRecordsByAxesUsedCount);
            AppendDictionary(builder, "Chart CatSerRange Intervals", ChartCategorySeriesRangeIntervals);
            AppendDictionary(builder, "Chart CatSerRange States", ChartCategorySeriesRangeStates);
            AppendDictionary(builder, "Chart AxcExt Date Ranges", ChartAxisExtensionDateRanges);
            AppendDictionary(builder, "Chart AxcExt Date Units", ChartAxisExtensionDateUnits);
            AppendDictionary(builder, "Chart AxcExt States", ChartAxisExtensionStates);
            AppendDictionary(builder, "Chart AxcExt Reserved States", ChartAxisExtensionReservedStates);
            AppendDictionary(builder, "Chart CatLab Alignments", ChartCategoryLabelAlignments);
            AppendDictionary(builder, "Chart CatLab Offsets", ChartCategoryLabelOffsets);
            AppendDictionary(builder, "Chart CatLab Count States", ChartCategoryLabelCountStates);
            AppendDictionary(builder, "Chart AxisLineFormat Targets", ChartAxisLineFormatTargets);
            AppendDictionary(builder, "Chart Series Category Data Types", ChartSeriesCategoryDataTypes);
            AppendDictionary(builder, "Chart Series Value Data Types", ChartSeriesValueDataTypes);
            AppendDictionary(builder, "Chart Series Bubble Size Data Types", ChartSeriesBubbleSizeDataTypes);
            AppendDictionary(builder, "Chart Series Value Counts", ChartSeriesValueCounts);
            AppendDictionary(builder, "Chart Series Chart Group Indexes", ChartSeriesChartGroupIndexes);
            AppendDictionary(builder, "Chart SeriesList Declared Counts", ChartSeriesListDeclaredCounts);
            AppendDictionary(builder, "Chart SeriesList Decoded Counts", ChartSeriesListDecodedCounts);
            AppendDictionary(builder, "Chart SeriesList Completeness States", ChartSeriesListCompletenessStates);
            AppendDictionary(builder, "Chart SeriesList Index Validity States", ChartSeriesListIndexValidityStates);
            AppendDictionary(builder, "Chart Pivot View References", ChartPivotViewReferences);
            AppendDictionary(builder, "Chart Series Data Cache Indexes", ChartSeriesDataCacheIndexes);
            AppendDictionary(builder, "Chart Series Data Cache Types", ChartSeriesDataCacheTypes);
            AppendDictionary(builder, "Chart DataSource Ids", ChartDataSourceIds);
            AppendDictionary(builder, "Chart DataSource Reference Types", ChartDataSourceReferenceTypes);
            AppendDictionary(builder, "Chart DataSource Number Format Ids", ChartDataSourceNumberFormatIds);
            AppendDictionary(builder, "Chart DataSource Formula Byte Counts", ChartDataSourceFormulaByteCounts);
            AppendDictionary(builder, "Chart DataSource Formula Projection States", ChartDataSourceFormulaProjectionStates);
            AppendDictionary(builder, "Chart DataSource Formula Texts", ChartDataSourceFormulaTexts);
            AppendDictionary(builder, "Chart DataSource Formula Projection Failures", ChartDataSourceFormulaProjectionFailures);
            AppendDictionary(builder, "Chart DataSource Formula Projection Failures By Token", ChartDataSourceFormulaProjectionFailuresByToken);
            AppendDictionary(builder, "Chart DataSource Formula Projection Failures By Token Name", ChartDataSourceFormulaProjectionFailuresByTokenName);
            AppendDictionary(builder, "Chart DataSource Formula Projection Failures By Offset", ChartDataSourceFormulaProjectionFailuresByOffset);
            AppendDictionary(builder, "Chart DataSource States", ChartDataSourceStates);
            AppendDictionary(builder, "Chart DataFormat Targets", ChartDataFormatTargets);
            AppendDictionary(builder, "Chart DataFormat Series Indexes", ChartDataFormatSeriesIndexes);
            AppendDictionary(builder, "Chart DataFormat Point Indexes", ChartDataFormatPointIndexes);
            AppendDictionary(builder, "Chart DataFormat Orders", ChartDataFormatOrders);
            AppendDictionary(builder, "Chart DataFormat States", ChartDataFormatStates);
            AppendDictionary(builder, "Chart Number Format Ids", ChartNumberFormatIds);
            AppendDictionary(builder, "Chart Font Indexes", ChartFontIndexes);
            AppendDictionary(builder, "Chart DataTable Options", ChartDataTableOptions);
            AppendDictionary(builder, "Chart DataTable Reserved States", ChartDataTableReservedStates);
            AppendDictionary(builder, "Chart Error Bar Directions", ChartErrorBarDirections);
            AppendDictionary(builder, "Chart Error Bar Value Sources", ChartErrorBarValueSources);
            AppendDictionary(builder, "Chart Error Bar Values", ChartErrorBarValues);
            AppendDictionary(builder, "Chart Error Bar States", ChartErrorBarStates);
            AppendDictionary(builder, "Chart Error Bar Reserved States", ChartErrorBarReservedStates);
            AppendDictionary(builder, "Chart Bar Overlap Percentages", ChartBarOverlapPercentages);
            AppendDictionary(builder, "Chart Bar Gap Widths", ChartBarGapWidths);
            AppendDictionary(builder, "Chart Bar States", ChartBarStates);
            AppendDictionary(builder, "Chart Line States", ChartLineStates);
            AppendDictionary(builder, "Chart Line Reserved States", ChartLineReservedStates);
            AppendDictionary(builder, "Chart Line Percent Stacked States", ChartLinePercentStackedStates);
            AppendDictionary(builder, "Chart Area States", ChartAreaStates);
            AppendDictionary(builder, "Chart Area Reserved States", ChartAreaReservedStates);
            AppendDictionary(builder, "Chart Area Percent Stacked States", ChartAreaPercentStackedStates);
            AppendDictionary(builder, "Chart BopPop Subtypes", ChartBopPopSubtypes);
            AppendDictionary(builder, "Chart BopPop Split Types", ChartBopPopSplitTypes);
            AppendDictionary(builder, "Chart BopPop Split Values", ChartBopPopSplitValues);
            AppendDictionary(builder, "Chart BopPop States", ChartBopPopStates);
            AppendDictionary(builder, "Chart BopPop Reserved States", ChartBopPopReservedStates);
            AppendDictionary(builder, "Chart BopPopCustom Data Point Counts", ChartBopPopCustomDataPointCounts);
            AppendDictionary(builder, "Chart BopPopCustom Secondary Counts", ChartBopPopCustomSecondaryCounts);
            AppendDictionary(builder, "Chart BopPopCustom Secondary Indexes", ChartBopPopCustomSecondaryIndexes);
            AppendDictionary(builder, "Chart BopPopCustom Completion States", ChartBopPopCustomCompletionStates);
            AppendDictionary(builder, "Chart BopPopCustom States", ChartBopPopCustomStates);
            AppendDictionary(builder, "Chart 3D View Angles", ChartThreeDimensionalViewAngles);
            AppendDictionary(builder, "Chart 3D Scale Values", ChartThreeDimensionalScaleValues);
            AppendDictionary(builder, "Chart 3D States", ChartThreeDimensionalStates);
            AppendDictionary(builder, "Chart 3D Reserved States", ChartThreeDimensionalReservedStates);
            AppendDictionary(builder, "Chart 3D Bar Shape Risers", ChartThreeDimensionalBarShapeRisers);
            AppendDictionary(builder, "Chart 3D Bar Shape Tapers", ChartThreeDimensionalBarShapeTapers);
            AppendDictionary(builder, "Chart 3D Bar Shape States", ChartThreeDimensionalBarShapeStates);
            AppendDictionary(builder, "Chart Scatter Bubble Size Ratios", ChartScatterBubbleSizeRatios);
            AppendDictionary(builder, "Chart Scatter Bubble Size Representations", ChartScatterBubbleSizeRepresentations);
            AppendDictionary(builder, "Chart Scatter Bubble Size Ratio States", ChartScatterBubbleSizeRatioStates);
            AppendDictionary(builder, "Chart Scatter States", ChartScatterStates);
            AppendDictionary(builder, "Chart FontBasis Scale Basis", ChartFontBasisScaleBasis);
            AppendDictionary(builder, "Chart FontBasis Font Indexes", ChartFontBasisFontIndexes);
            AppendDictionary(builder, "Chart FontBasis States", ChartFontBasisStates);
            AppendDictionary(builder, "Chart CrtLayout12 Mode Pairs", ChartLayout12ModePairs);
            AppendDictionary(builder, "Chart CrtLayout12 Auto Layout Types", ChartLayout12AutoLayoutTypes);
            AppendDictionary(builder, "Chart CrtLayout12 Checksums", ChartLayout12Checksums);
            AppendDictionary(builder, "Chart CrtLayout12 Rectangles", ChartLayout12Rectangles);
            AppendDictionary(builder, "Chart CrtLayout12A Targets", ChartPlotAreaLayout12Targets);
            AppendDictionary(builder, "Chart CrtLayout12A Mode Pairs", ChartPlotAreaLayout12ModePairs);
            AppendDictionary(builder, "Chart CrtLayout12A Checksums", ChartPlotAreaLayout12Checksums);
            AppendDictionary(builder, "Chart CrtLayout12A Bounds", ChartPlotAreaLayout12Bounds);
            AppendDictionary(builder, "Chart CrtLayout12A Rectangles", ChartPlotAreaLayout12Rectangles);
            AppendDictionary(builder, "Chart Future Record Info Versions", ChartFutureRecordInfoVersions);
            AppendDictionary(builder, "Chart Future Record Info Range Counts", ChartFutureRecordInfoRangeCounts);
            AppendDictionary(builder, "Chart Future Record Info Ranges", ChartFutureRecordInfoRanges);
            AppendDictionary(builder, "Chart Future Block Directions", ChartFutureBlockDirections);
            AppendDictionary(builder, "Chart Future Block Object Kinds", ChartFutureBlockObjectKinds);
            AppendDictionary(builder, "Chart Future Block Scopes", ChartFutureBlockScopes);
            AppendDictionary(builder, "Chart Units Reserved Values", ChartUnitsReservedValues);
            AppendDictionary(builder, "Chart Units Reserved States", ChartUnitsReservedStates);
            AppendDictionary(builder, "Chart XmlTkChain Declared Byte Counts", ChartXmlTokenChainDeclaredByteCounts);
            AppendDictionary(builder, "Chart XmlTkChain First Segment Byte Counts", ChartXmlTokenChainFirstSegmentByteCounts);
            AppendDictionary(builder, "Chart XmlTkChain Completion States", ChartXmlTokenChainCompletionStates);
            AppendDictionary(builder, "Chart XmlTkChain Trailing States", ChartXmlTokenChainTrailingStates);
            AppendDictionary(builder, "Chart Sheet Property Empty Cell Modes", ChartSheetPropertyEmptyCellModes);
            AppendDictionary(builder, "Chart Sheet Property States", ChartSheetPropertyStates);
            AppendDictionary(builder, "Chart LineFormat Styles", ChartLineFormatStyles);
            AppendDictionary(builder, "Chart LineFormat Weights", ChartLineFormatWeights);
            AppendDictionary(builder, "Chart LineFormat Colors", ChartLineFormatColors);
            AppendDictionary(builder, "Chart LineFormat Color Indexes", ChartLineFormatColorIndexes);
            AppendDictionary(builder, "Chart LineFormat States", ChartLineFormatStates);
            AppendDictionary(builder, "Chart AreaFormat Patterns", ChartAreaFormatPatterns);
            AppendDictionary(builder, "Chart AreaFormat Colors", ChartAreaFormatColors);
            AppendDictionary(builder, "Chart AreaFormat Color Indexes", ChartAreaFormatColorIndexes);
            AppendDictionary(builder, "Chart AreaFormat States", ChartAreaFormatStates);
            AppendDictionary(builder, "Chart MarkerFormat Types", ChartMarkerFormatTypes);
            AppendDictionary(builder, "Chart MarkerFormat Sizes", ChartMarkerFormatSizes);
            AppendDictionary(builder, "Chart MarkerFormat Colors", ChartMarkerFormatColors);
            AppendDictionary(builder, "Chart MarkerFormat Color Indexes", ChartMarkerFormatColorIndexes);
            AppendDictionary(builder, "Chart MarkerFormat States", ChartMarkerFormatStates);
            AppendDictionary(builder, "Chart PieFormat Explosions", ChartPieFormatExplosions);
            AppendDictionary(builder, "Chart SerFmt Flags", ChartSeriesFormatFlags);
            AppendDictionary(builder, "Chart SerFmt States", ChartSeriesFormatStates);
            AppendDictionary(builder, "Chart SerFmt Reserved Values", ChartSeriesFormatReservedValues);
            AppendDictionary(builder, "Chart SerFmt Reserved States", ChartSeriesFormatReservedStates);
            AppendDictionary(builder, "Chart ClrtClient Declared Counts", ChartClientColorPaletteDeclaredCounts);
            AppendDictionary(builder, "Chart ClrtClient Decoded Counts", ChartClientColorPaletteDecodedCounts);
            AppendDictionary(builder, "Chart ClrtClient Completeness States", ChartClientColorPaletteCompletenessStates);
            AppendDictionary(builder, "Chart ClrtClient Expected Count States", ChartClientColorPaletteExpectedCountStates);
            AppendDictionary(builder, "Chart ClrtClient Colors", ChartClientColorPaletteColors);
            AppendDictionary(builder, "Chart GelFrame OfficeArt Records By Type", ChartGelFrameOfficeArtRecordsByType);
            AppendDictionary(builder, "Chart GelFrame OfficeArt Records By Container State", ChartGelFrameOfficeArtRecordsByContainerState);
            AppendDictionary(builder, "Chart GelFrame Shape Property Counts", ChartGelFrameShapePropertyCounts);
            AppendDictionary(builder, "Chart GelFrame Shape Properties By Name", ChartGelFrameShapePropertiesByName);
            AppendDictionary(builder, "Chart GelFrame Shape Properties By Group", ChartGelFrameShapePropertiesByGroup);
            AppendDictionary(builder, "Chart GelFrame Shape Properties By Flag State", ChartGelFrameShapePropertiesByFlagState);
            AppendDictionary(builder, "Chart GelFrame Shape Properties By Value", ChartGelFrameShapePropertiesByValue);
            AppendDictionary(builder, "Chart AttachedLabel Flags", ChartAttachedLabelFlags);
            AppendDictionary(builder, "Chart AttachedLabel States", ChartAttachedLabelStates);
            AppendDictionary(builder, "Chart DefaultText Targets", ChartDefaultTextTargets);
            AppendDictionary(builder, "Chart Text Horizontal Alignments", ChartTextHorizontalAlignments);
            AppendDictionary(builder, "Chart Text Vertical Alignments", ChartTextVerticalAlignments);
            AppendDictionary(builder, "Chart Text Data Label Positions", ChartTextDataLabelPositions);
            AppendDictionary(builder, "Chart Text Flags", ChartTextFlags);
            AppendDictionary(builder, "Chart ObjectLink Targets", ChartObjectLinkTargets);
            AppendDictionary(builder, "Chart Legend Layouts", ChartLegendLayouts);
            AppendDictionary(builder, "Chart Legend Spacing States", ChartLegendSpacingStates);
            AppendDictionary(builder, "Chart Legend Reserved States", ChartLegendReservedStates);
            AppendDictionary(builder, "Chart Legend Auto Position States", ChartLegendAutoPositionStates);
            AppendDictionary(builder, "Chart Legend Data Table States", ChartLegendDataTableStates);
            AppendDictionary(builder, "Chart Tick Major Locations", ChartTickMajorLocations);
            AppendDictionary(builder, "Chart Tick Label Locations", ChartTickLabelLocations);
            AppendDictionary(builder, "Chart ValueRange Scales", ChartValueRangeScales);
            AppendDictionary(builder, "Chart ValueRange States", ChartValueRangeStates);
            AppendDictionary(builder, "Chart Position Mode Pairs", ChartPositionModePairs);
            AppendDictionary(builder, "Chart Position Rectangles", ChartPositionRectangles);
            AppendDictionary(builder, "Chart Position Semantic Types", ChartPositionSemanticTypes);
            AppendDictionary(builder, "Chart Position Coordinate Meanings", ChartPositionCoordinateMeanings);
            AppendDictionary(builder, "Chart Position Ignored Coordinate States", ChartPositionIgnoredCoordinateStates);
            AppendDictionary(builder, "Chart Position Known Semantic States", ChartPositionKnownSemanticStates);
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
            AppendDictionary(builder, "Drawing Future Record Wrapped Types", DrawingFutureRecordWrappedTypes);
            AppendDictionary(builder, "Drawing Future Record Flags", DrawingFutureRecordFlags);
            AppendDictionary(builder, "Drawing Future Record Reference States", DrawingFutureRecordReferenceStates);
            AppendDictionary(builder, "Drawing Future Record Ranges", DrawingFutureRecordRanges);
            AppendDictionary(builder, "Drawing Future Record Stream Byte Counts", DrawingFutureRecordStreamByteCounts);
            AppendDictionary(builder, "Drawing Header Footer Picture Header States", DrawingHeaderFooterPictureHeaderStates);
            AppendDictionary(builder, "Drawing Header Footer Picture Drawing Kinds", DrawingHeaderFooterPictureDrawingKinds);
            AppendDictionary(builder, "Drawing Header Footer Picture Continuation States", DrawingHeaderFooterPictureContinuationStates);
            AppendDictionary(builder, "Drawing Header Footer Picture Future Record Flags", DrawingHeaderFooterPictureFutureRecordFlags);
            AppendDictionary(builder, "Drawing Header Footer Picture Drawing Byte Counts", DrawingHeaderFooterPictureDrawingByteCounts);
            AppendDictionary(builder, "Drawing Text Object Alignments", DrawingTextObjectAlignments);
            AppendDictionary(builder, "Drawing Text Object Rotations", DrawingTextObjectRotations);
            AppendDictionary(builder, "Drawing Text Object Text Lengths", DrawingTextObjectTextLengths);
            AppendDictionary(builder, "Drawing Text Object Formatting Run Byte Counts", DrawingTextObjectFormattingRunByteCounts);
            AppendDictionary(builder, "Drawing Text Object Formula Byte Counts", DrawingTextObjectFormulaByteCounts);
            AppendDictionary(builder, "Drawing Text Object Flags", DrawingTextObjectFlags);
            AppendDictionary(builder, "Drawing Records By Escher Record Type", DrawingRecordsByEscherRecordType);
            AppendDictionary(builder, "Drawing Records By Escher Record Type Name", DrawingRecordsByEscherRecordTypeName);
            AppendDictionary(builder, "Drawing OfficeArt Records By Type", DrawingOfficeArtRecordsByType);
            AppendDictionary(builder, "Drawing OfficeArt Records By Type Name", DrawingOfficeArtRecordsByTypeName);
            AppendDictionary(builder, "Drawing OfficeArt Records By Depth", DrawingOfficeArtRecordsByDepth);
            AppendDictionary(builder, "Drawing OfficeArt Records By Container State", DrawingOfficeArtRecordsByContainerState);
            AppendDictionary(builder, "Drawing OfficeArt Records By Payload Length", DrawingOfficeArtRecordsByPayloadLength);
            AppendDictionary(builder, "Drawing Group Blocks By Max Shape Id", DrawingGroupBlocksByMaxShapeId);
            AppendDictionary(builder, "Drawing Group Blocks By Declared Identifier Cluster Count", DrawingGroupBlocksByDeclaredIdentifierClusterCount);
            AppendDictionary(builder, "Drawing Group Blocks By Decoded Identifier Cluster Count", DrawingGroupBlocksByDecodedIdentifierClusterCount);
            AppendDictionary(builder, "Drawing Group Blocks By Saved Shape Count", DrawingGroupBlocksBySavedShapeCount);
            AppendDictionary(builder, "Drawing Group Blocks By Saved Drawing Count", DrawingGroupBlocksBySavedDrawingCount);
            AppendDictionary(builder, "Drawing Identifier Clusters By Drawing Id", DrawingIdentifierClustersByDrawingId);
            AppendDictionary(builder, "Drawing Identifier Clusters By Current Shape Id", DrawingIdentifierClustersByCurrentShapeId);
            AppendDictionary(builder, "Drawing Group Infos By Drawing Id", DrawingGroupInfosByDrawingId);
            AppendDictionary(builder, "Drawing Group Infos By Shape Count", DrawingGroupInfosByShapeCount);
            AppendDictionary(builder, "Drawing Group Infos By Last Shape Id", DrawingGroupInfosByLastShapeId);
            AppendDictionary(builder, "Drawing Shape Properties By Id", DrawingShapePropertiesById);
            AppendDictionary(builder, "Drawing Shape Properties By Name", DrawingShapePropertiesByName);
            AppendDictionary(builder, "Drawing Shape Properties By Group", DrawingShapePropertiesByGroup);
            AppendDictionary(builder, "Drawing Shape Properties By Flag State", DrawingShapePropertiesByFlagState);
            AppendDictionary(builder, "Drawing Shape Properties By Value", DrawingShapePropertiesByValue);
            AppendDictionary(builder, "Drawing Shape Complex Properties By Declared Length", DrawingShapeComplexPropertiesByDeclaredLength);
            AppendDictionary(builder, "Drawing Shape Complex Properties By Available Length", DrawingShapeComplexPropertiesByAvailableLength);
            AppendDictionary(builder, "Drawing Shape Complex Properties By Text", DrawingShapeComplexPropertiesByText);
            AppendDictionary(builder, "Drawing BLIP Store Entries By Type", DrawingBlipStoreEntriesByType);
            AppendDictionary(builder, "Drawing BLIP Store Entries By Location", DrawingBlipStoreEntriesByLocation);
            AppendDictionary(builder, "Drawing BLIP Store Entries By Type And Location", DrawingBlipStoreEntriesByTypeAndLocation);
            AppendDictionary(builder, "Drawing BLIP Store Entries By UID", DrawingBlipStoreEntriesByUid);
            AppendDictionary(builder, "Drawing BLIP Store Entries By Embedded Record Type", DrawingBlipStoreEntriesByEmbeddedRecordType);
            AppendDictionary(builder, "Drawing BLIP Store Entries By Embedded Payload Available Length", DrawingBlipStoreEntriesByEmbeddedPayloadAvailableLength);
            AppendDictionary(builder, "Drawing BLIP Store Entries By Embedded Payload Hash", DrawingBlipStoreEntriesByEmbeddedPayloadHash);
            AppendDictionary(builder, "Drawing BLIP Store Entries By Size", DrawingBlipStoreEntriesBySize);
            AppendDictionary(builder, "Drawing BLIP Store Entries By Reference Count", DrawingBlipStoreEntriesByReferenceCount);
            AppendDictionary(builder, "Drawing Shape BLIP Properties By Location", DrawingShapeBlipPropertiesByLocation);
            AppendDictionary(builder, "Drawing Shape BLIP Properties By Name And Value", DrawingShapeBlipPropertiesByNameAndValue);
            AppendDictionary(builder, "Drawing Picture BLIP References By Location", DrawingPictureBlipReferencesByLocation);
            AppendDictionary(builder, "Drawing Picture BLIP References By Value", DrawingPictureBlipReferencesByValue);
            AppendDictionary(builder, "Drawing Picture States", DrawingPictureStates);
            AppendDictionary(builder, "Drawing Picture Count States", DrawingPictureCountStates);
            AppendDictionary(builder, "Drawing Shape Entries By Type", DrawingShapeEntriesByType);
            AppendDictionary(builder, "Drawing Shape Entries By Id", DrawingShapeEntriesById);
            AppendDictionary(builder, "Drawing Shape Entries By Flags", DrawingShapeEntriesByFlags);
            AppendDictionary(builder, "Drawing Shape Entries By Reserved State", DrawingShapeEntriesByReservedState);
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
            AppendDictionary(builder, "Compound Feature Entries By Content Kind", CompoundFeatureEntriesByContentKind);
            AppendDictionary(builder, "Compound Feature Entries By Role And Content Kind", CompoundFeatureEntriesByRoleAndContentKind);
            AppendDictionary(builder, "Compound Feature Entries By Size", CompoundFeatureEntriesBySize);
            AppendDictionary(builder, "Compound Feature Entries By Role And Size", CompoundFeatureEntriesByRoleAndSize);
            AppendDictionary(builder, "Compound VBA Modules By Name", CompoundVbaModulesByName);
            AppendDictionary(builder, "Compound VBA Modules By Path", CompoundVbaModulesByPath);
            AppendDictionary(builder, "Compound VBA Modules By Size", CompoundVbaModulesBySize);
            AppendDictionary(builder, "Compound VBA Modules By Name And Size", CompoundVbaModulesByNameAndSize);
            AppendDictionary(builder, "Compound VBA Modules By Content Kind", CompoundVbaModulesByContentKind);
            AppendDictionary(builder, "Compound VBA Modules By Name And Content Kind", CompoundVbaModulesByNameAndContentKind);
            AppendDictionary(builder, "Compound VBA Modules By CodeName Match", CompoundVbaModulesByCodeNameMatch);
            AppendDictionary(builder, "Compound VBA Modules By CodeName Match And Name", CompoundVbaModulesByCodeNameMatchAndName);
            AppendDictionary(builder, "Compound VBA Projects By Module Count", CompoundVbaProjectsByModuleCount);
            AppendDictionary(builder, "Compound VBA Projects By Module Byte Count", CompoundVbaProjectsByModuleByteCount);
            AppendDictionary(builder, "Compound VBA Projects By Structure", CompoundVbaProjectsByStructure);
            AppendDictionary(builder, "VBA Project Workbook States", VbaProjectWorkbookStates);
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
            AppendDictionary(builder, "Cell Style Extension Properties By Type", CellStyleExtensionPropertiesByType);
            AppendDictionary(builder, "Cell Style Extension Properties By Name", CellStyleExtensionPropertiesByName);
            AppendDictionary(builder, "Cell Style Extension Properties By Data Byte Count", CellStyleExtensionPropertiesByDataByteCount);
            AppendDictionary(builder, "Cell Style Extension Properties By Numeric Value", CellStyleExtensionPropertiesByNumericValue);
            AppendDictionary(builder, "Cell Style Extension Properties By Numeric Value Name", CellStyleExtensionPropertiesByNumericValueName);
            AppendDictionary(builder, "Cell Style Extension Properties By Color Type", CellStyleExtensionPropertiesByColorType);
            AppendDictionary(builder, "Cell Style Extension Properties By Color Tint Shade", CellStyleExtensionPropertiesByColorTintShade);
            AppendDictionary(builder, "Cell Style Extension Properties By Color Value", CellStyleExtensionPropertiesByColorValue);
            AppendDictionary(builder, "Workbook Metadata Records By Kind", WorkbookMetadataRecordsByKind.ToDictionary(
                entry => entry.Key.ToString(),
                entry => entry.Value,
                StringComparer.OrdinalIgnoreCase));
            AppendDictionary(builder, "Workbook Future Metadata Records By Kind", WorkbookFutureMetadataRecordsByKind);
            AppendDictionary(builder, "Workbook Future Metadata Records By Record Type", WorkbookFutureMetadataRecordsByRecordType);
            AppendDictionary(builder, "Workbook Future Metadata Records By Record Name", WorkbookFutureMetadataRecordsByRecordName);
            AppendDictionary(builder, "Workbook Future Metadata Records By Header State", WorkbookFutureMetadataRecordsByHeaderState);
            AppendDictionary(builder, "Workbook Future Metadata Records By Header Record Type", WorkbookFutureMetadataRecordsByHeaderRecordType);
            AppendDictionary(builder, "Workbook Future Metadata Records By Header Flags", WorkbookFutureMetadataRecordsByHeaderFlags);
            AppendDictionary(builder, "Workbook Future Metadata Records By Payload Length", WorkbookFutureMetadataRecordsByPayloadLength);
            AppendDictionary(builder, "Workbook Future Metadata Records By Body Byte Count", WorkbookFutureMetadataRecordsByBodyByteCount);
            AppendDictionary(builder, "Worksheet Metadata Records By Kind", WorksheetMetadataRecordsByKind.ToDictionary(
                entry => entry.Key.ToString(),
                entry => entry.Value,
                StringComparer.OrdinalIgnoreCase));
            AppendDictionary(builder, "Worksheet Future Metadata Records By Kind", WorksheetFutureMetadataRecordsByKind);
            AppendDictionary(builder, "Worksheet Future Metadata Records By Sheet", WorksheetFutureMetadataRecordsBySheet);
            AppendDictionary(builder, "Worksheet Future Metadata Records By Sheet And Kind", WorksheetFutureMetadataRecordsBySheetAndKind);
            AppendDictionary(builder, "Worksheet Future Metadata Records By Record Type", WorksheetFutureMetadataRecordsByRecordType);
            AppendDictionary(builder, "Worksheet Future Metadata Records By Record Name", WorksheetFutureMetadataRecordsByRecordName);
            AppendDictionary(builder, "Worksheet Future Metadata Records By Header State", WorksheetFutureMetadataRecordsByHeaderState);
            AppendDictionary(builder, "Worksheet Future Metadata Records By Header Record Type", WorksheetFutureMetadataRecordsByHeaderRecordType);
            AppendDictionary(builder, "Worksheet Future Metadata Records By Header Flags", WorksheetFutureMetadataRecordsByHeaderFlags);
            AppendDictionary(builder, "Worksheet Future Metadata Records By Payload Length", WorksheetFutureMetadataRecordsByPayloadLength);
            AppendDictionary(builder, "Worksheet Future Metadata Records By Body Byte Count", WorksheetFutureMetadataRecordsByBodyByteCount);
            AppendDictionary(builder, "Unsupported Sheet Metadata Records By Kind", UnsupportedSheetMetadataRecordsByKind.ToDictionary(
                entry => entry.Key.ToString(),
                entry => entry.Value,
                StringComparer.OrdinalIgnoreCase));
            AppendDictionary(builder, "Unsupported Sheet Future Metadata Records By Kind", UnsupportedSheetFutureMetadataRecordsByKind);
            AppendDictionary(builder, "Unsupported Sheet Future Metadata Records By Sheet", UnsupportedSheetFutureMetadataRecordsBySheet);
            AppendDictionary(builder, "Unsupported Sheet Future Metadata Records By Sheet And Kind", UnsupportedSheetFutureMetadataRecordsBySheetAndKind);
            AppendDictionary(builder, "Unsupported Sheet Future Metadata Records By Record Type", UnsupportedSheetFutureMetadataRecordsByRecordType);
            AppendDictionary(builder, "Unsupported Sheet Future Metadata Records By Record Name", UnsupportedSheetFutureMetadataRecordsByRecordName);
            AppendDictionary(builder, "Unsupported Sheet Future Metadata Records By Header State", UnsupportedSheetFutureMetadataRecordsByHeaderState);
            AppendDictionary(builder, "Unsupported Sheet Future Metadata Records By Header Record Type", UnsupportedSheetFutureMetadataRecordsByHeaderRecordType);
            AppendDictionary(builder, "Unsupported Sheet Future Metadata Records By Header Flags", UnsupportedSheetFutureMetadataRecordsByHeaderFlags);
            AppendDictionary(builder, "Unsupported Sheet Future Metadata Records By Payload Length", UnsupportedSheetFutureMetadataRecordsByPayloadLength);
            AppendDictionary(builder, "Unsupported Sheet Future Metadata Records By Body Byte Count", UnsupportedSheetFutureMetadataRecordsByBodyByteCount);
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

        private static LegacyXlsUnsupportedFeature[] GetUnsupportedProjectionGaps(LegacyXlsWorkbook workbook) {
            var projectedFeatureKeys = new HashSet<string>(
                GetProjectedFeatureKeys(workbook),
                StringComparer.Ordinal);

            return workbook.UnsupportedFeatures
                .Where(feature => !projectedFeatureKeys.Contains(GetUnsupportedFeatureKey(feature)))
                .ToArray();
        }

        private static IEnumerable<string> GetProjectedFeatureKeys(LegacyXlsWorkbook workbook) {
            foreach (LegacyXlsPreservedFeatureRecord record in workbook.PreservedFeatureRecords) {
                yield return GetPreservedFeatureRecordKey(record);
            }

            foreach (LegacyXlsCompoundFeatureRecord record in workbook.CompoundFeatureRecords) {
                string? key = GetCompoundFeatureRecordKey(record);
                if (key != null) {
                    yield return key;
                }
            }
        }

        private static string GetUnsupportedFeatureKey(LegacyXlsUnsupportedFeature feature) {
            return GetProjectedFeatureKey(
                feature.Kind,
                feature.Code,
                feature.SheetName,
                feature.RecordOffset,
                feature.RecordType,
                feature.DetailCode);
        }

        private static string GetPreservedFeatureRecordKey(LegacyXlsPreservedFeatureRecord record) {
            return GetProjectedFeatureKey(
                record.Kind,
                record.Code,
                record.SheetName,
                record.RecordOffset,
                record.RecordType,
                record.DetailCode);
        }

        private static string? GetCompoundFeatureRecordKey(LegacyXlsCompoundFeatureRecord record) {
            return record.Kind switch {
                LegacyXlsCompoundFeatureRecordKind.VbaProject => GetProjectedFeatureKey(
                    LegacyXlsUnsupportedFeatureKind.VbaProject,
                    "XLS-COMPOUND-FEATURE-VBA-PROJECT-PRESERVED",
                    detailCode: "Compound:VbaProjectStorage"),
                LegacyXlsCompoundFeatureRecordKind.OleObject => GetProjectedFeatureKey(
                    LegacyXlsUnsupportedFeatureKind.OleObject,
                    "XLS-COMPOUND-FEATURE-OLE-OBJECT-PRESERVED",
                    detailCode: "Compound:OleObjectStorage"),
                LegacyXlsCompoundFeatureRecordKind.DigitalSignature => GetProjectedFeatureKey(
                    LegacyXlsUnsupportedFeatureKind.DigitalSignature,
                    "XLS-COMPOUND-FEATURE-DIGITAL-SIGNATURE-DIAGNOSED",
                    detailCode: "Compound:DigitalSignature"),
                _ => null
            };
        }

        private static string GetProjectedFeatureKey(
            LegacyXlsUnsupportedFeatureKind kind,
            string code,
            string? sheetName = null,
            int? recordOffset = null,
            ushort? recordType = null,
            string? detailCode = null) {
            return string.Join("|",
                kind,
                code,
                sheetName ?? string.Empty,
                recordOffset?.ToString(CultureInfo.InvariantCulture) ?? string.Empty,
                recordType?.ToString("X4", CultureInfo.InvariantCulture) ?? string.Empty,
                detailCode ?? string.Empty);
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

        private static IReadOnlyDictionary<string, int> CountChartSheetChartRecordKinds(IEnumerable<LegacyXlsChartSheet> sheets) {
            return sheets
                .SelectMany(sheet => sheet.ChartRecordsByKind)
                .GroupBy(entry => entry.Key.ToString(), StringComparer.OrdinalIgnoreCase)
                .OrderBy(group => group.Key, StringComparer.OrdinalIgnoreCase)
                .ToDictionary(group => group.Key, group => group.Sum(entry => entry.Value), StringComparer.OrdinalIgnoreCase);
        }

        private static IReadOnlyDictionary<string, int> CountChartSheetChartRecordKindsBySheet(IEnumerable<LegacyXlsChartSheet> sheets) {
            return sheets
                .SelectMany(sheet => sheet.ChartRecordsByKind.Select(entry => new {
                    Key = $"Sheet:{sheet.Name};Kind:{entry.Key}",
                    entry.Value
                }))
                .GroupBy(entry => entry.Key, StringComparer.OrdinalIgnoreCase)
                .OrderBy(group => group.Key, StringComparer.OrdinalIgnoreCase)
                .ToDictionary(group => group.Key, group => group.Sum(entry => entry.Value), StringComparer.OrdinalIgnoreCase);
        }

        private static IReadOnlyDictionary<string, int> CountChartSheetChartTypes(IEnumerable<LegacyXlsChartSheet> sheets) {
            return sheets
                .SelectMany(sheet => sheet.ChartRecordsByChartType)
                .GroupBy(entry => entry.Key, StringComparer.OrdinalIgnoreCase)
                .OrderBy(group => group.Key, StringComparer.OrdinalIgnoreCase)
                .ToDictionary(group => group.Key, group => group.Sum(entry => entry.Value), StringComparer.OrdinalIgnoreCase);
        }

        private static IReadOnlyDictionary<string, int> CountChartSheetChartTypesBySheet(IEnumerable<LegacyXlsChartSheet> sheets) {
            return sheets
                .SelectMany(sheet => sheet.ChartRecordsByChartType.Select(entry => new {
                    Key = $"Sheet:{sheet.Name};ChartType:{entry.Key}",
                    entry.Value
                }))
                .GroupBy(entry => entry.Key, StringComparer.OrdinalIgnoreCase)
                .OrderBy(group => group.Key, StringComparer.OrdinalIgnoreCase)
                .ToDictionary(group => group.Key, group => group.Sum(entry => entry.Value), StringComparer.OrdinalIgnoreCase);
        }

        private static IReadOnlyDictionary<string, int> CountUnsupportedChartSheetChartRecordKinds(IEnumerable<LegacyXlsUnsupportedSheet> sheets) {
            return sheets
                .Where(sheet => sheet.Kind == LegacyXlsUnsupportedSheetKind.ChartSheet)
                .SelectMany(sheet => sheet.ChartRecordsByKind)
                .GroupBy(entry => entry.Key.ToString(), StringComparer.OrdinalIgnoreCase)
                .OrderBy(group => group.Key, StringComparer.OrdinalIgnoreCase)
                .ToDictionary(group => group.Key, group => group.Sum(entry => entry.Value), StringComparer.OrdinalIgnoreCase);
        }

        private static IReadOnlyDictionary<string, int> CountUnsupportedChartSheetChartRecordKindsBySheet(IEnumerable<LegacyXlsUnsupportedSheet> sheets) {
            return sheets
                .Where(sheet => sheet.Kind == LegacyXlsUnsupportedSheetKind.ChartSheet)
                .SelectMany(sheet => sheet.ChartRecordsByKind.Select(entry => new {
                    Key = $"Sheet:{sheet.Name};Kind:{entry.Key}",
                    entry.Value
                }))
                .GroupBy(entry => entry.Key, StringComparer.OrdinalIgnoreCase)
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

        private static IReadOnlyDictionary<string, int> CountUnsupportedChartSheetChartTypesBySheet(IEnumerable<LegacyXlsUnsupportedSheet> sheets) {
            return sheets
                .Where(sheet => sheet.Kind == LegacyXlsUnsupportedSheetKind.ChartSheet)
                .SelectMany(sheet => sheet.ChartRecordsByChartType.Select(entry => new {
                    Key = $"Sheet:{sheet.Name};ChartType:{entry.Key}",
                    entry.Value
                }))
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

        private static string GetExternalQueryConnectionStateKey(LegacyXlsExternalQueryConnection connection) {
            return connection.SourceTypeName
                + $"|Flags:{FormatJoinedState(connection.ConnectionFlagNames)}"
                + $"|Options:{FormatJoinedState(connection.QueryOptionNames)}"
                + $"|Parameters:{connection.ParameterFlagCount}"
                + $"|ParameterBytes:{connection.ParameterFlagByteCount}"
                + $"|FutureBytes:{connection.FutureByteCount}";
        }

        private static IEnumerable<string> GetExternalQueryConnectionFlagKeys(LegacyXlsExternalQueryConnection connection) {
            return connection.ConnectionFlagNames.Count == 0
                ? new[] { "None" }
                : connection.ConnectionFlagNames;
        }

        private static IEnumerable<string> GetExternalQueryConnectionOptionKeys(LegacyXlsExternalQueryConnection connection) {
            return connection.QueryOptionNames.Count == 0
                ? new[] { "None" }
                : connection.QueryOptionNames;
        }

        private static string FormatJoinedState(IReadOnlyList<string> values) {
            return values.Count == 0 ? "None" : string.Join("+", values);
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

        private static string GetCompoundFeatureEntryLeafName(string entry) {
            int slashIndex = entry.LastIndexOf('/');
            int backslashIndex = entry.LastIndexOf('\\');
            int separatorIndex = Math.Max(slashIndex, backslashIndex);
            return separatorIndex >= 0 && separatorIndex + 1 < entry.Length ? entry.Substring(separatorIndex + 1) : entry;
        }

        private static IEnumerable<string> GetCompoundVbaModuleCodeNameMatchKeys(LegacyXlsWorkbook workbook) {
            foreach ((string match, _) in GetCompoundVbaModuleCodeNameMatches(workbook)) {
                yield return match;
            }
        }

        private static IEnumerable<string> GetCompoundVbaModuleCodeNameMatchAndNameKeys(LegacyXlsWorkbook workbook) {
            foreach ((string match, string moduleName) in GetCompoundVbaModuleCodeNameMatches(workbook)) {
                yield return $"{match}|{moduleName}";
            }
        }

        private static IEnumerable<(string Match, string ModuleName)> GetCompoundVbaModuleCodeNameMatches(LegacyXlsWorkbook workbook) {
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
                    yield return ("WorkbookCodeName", moduleName);
                } else if (worksheetCodeNames.Contains(moduleName)) {
                    yield return ("WorksheetCodeName", moduleName);
                } else {
                    yield return ("UnmatchedCodeName", moduleName);
                }
            }
        }

        private static string GetCompoundVbaProjectStructureKey(LegacyXlsCompoundFeatureRecord record) {
            int dirStreams = record.EntryDetails.Count(entry => entry.Role == LegacyXlsCompoundFeatureEntryRole.VbaDirStream);
            int projectStreams = record.EntryDetails.Count(entry => entry.Role == LegacyXlsCompoundFeatureEntryRole.VbaProjectStream);
            int storageEntries = record.EntryDetails.Count(entry => entry.Role == LegacyXlsCompoundFeatureEntryRole.VbaProjectStorage
                || entry.Role == LegacyXlsCompoundFeatureEntryRole.VbaStorage);
            return $"Modules:{record.VbaModuleCount}|DirStreams:{dirStreams}|ProjectStreams:{projectStreams}|Storages:{storageEntries}";
        }

        private static IEnumerable<string> GetVbaProjectWorkbookStateKeys(LegacyXlsWorkbook workbook) {
            LegacyXlsCompoundFeatureRecord[] vbaProjectRecords = workbook.CompoundFeatureRecords
                .Where(record => record.Kind == LegacyXlsCompoundFeatureRecordKind.VbaProject)
                .ToArray();
            bool hasCompoundProject = vbaProjectRecords.Length > 0;
            int moduleCount = vbaProjectRecords.Sum(record => record.VbaModuleCount);
            if (!workbook.HasVbaProjectMarker && !workbook.HasVbaProjectWithoutMacros && !hasCompoundProject && moduleCount == 0) {
                yield break;
            }

            yield return $"BiffMarker:{GetPresenceKey(workbook.HasVbaProjectMarker)}"
                + $"|NoMacrosMarker:{GetPresenceKey(workbook.HasVbaProjectWithoutMacros)}"
                + $"|CompoundProject:{GetPresenceKey(hasCompoundProject)}"
                + $"|Modules:{GetPresenceKey(moduleCount > 0)}";
        }

        private static string GetPresenceKey(bool value) {
            return value ? "Present" : "Missing";
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

        private static IEnumerable<string> GetFileFormatStateKeys(LegacyXlsWorkbook workbook) {
            bool encrypted = workbook.UnsupportedFeatures.Any(feature => feature.Kind == LegacyXlsUnsupportedFeatureKind.EncryptedWorkbook);
            bool unsupportedBiff = workbook.UnsupportedFeatures.Any(feature => feature.Kind == LegacyXlsUnsupportedFeatureKind.UnsupportedBiffVersion);
            bool malformedBof = workbook.Diagnostics.Any(IsBofDiagnostic);
            bool truncatedStream = workbook.Diagnostics.Any(IsBiffTruncationDiagnostic);

            yield return $"Encryption:{GetPresenceKey(encrypted)}";
            yield return $"UnsupportedBiffVersion:{GetPresenceKey(unsupportedBiff)}";
            yield return $"MalformedBof:{GetPresenceKey(malformedBof)}";
            yield return $"TruncatedStream:{GetPresenceKey(truncatedStream)}";

            if (!encrypted && !unsupportedBiff && !malformedBof && !truncatedStream) {
                yield return "WorkbookFormat:SupportedBiff8";
                yield break;
            }

            if (encrypted) {
                yield return "WorkbookFormat:Encrypted";
            }

            if (unsupportedBiff) {
                yield return "WorkbookFormat:UnsupportedBiff";
            }

            if (malformedBof) {
                yield return "WorkbookFormat:MalformedBof";
            }

            if (truncatedStream) {
                yield return "WorkbookFormat:TruncatedStream";
            }
        }

        private static bool IsBofDiagnostic(LegacyXlsImportDiagnostic diagnostic) {
            return diagnostic.Code.IndexOf("-BOF-", StringComparison.Ordinal) >= 0;
        }

        private static bool IsBiffTruncationDiagnostic(LegacyXlsImportDiagnostic diagnostic) {
            return diagnostic.Code.StartsWith("XLS-BIFF-", StringComparison.Ordinal)
                && diagnostic.Code.IndexOf("TRUNCATED", StringComparison.Ordinal) >= 0;
        }

        private static bool IsFileFormatBlocker(LegacyXlsUnsupportedFeature feature) {
            return feature.Kind == LegacyXlsUnsupportedFeatureKind.EncryptedWorkbook
                || feature.Kind == LegacyXlsUnsupportedFeatureKind.UnsupportedBiffVersion;
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

        private static string GetUnsupportedChartSheetStateKey(LegacyXlsUnsupportedSheet sheet) {
            return $"PrintSize:{GetPresenceKey(sheet.ChartPrintSize.HasValue)}"
                + $"|TextObjects:{GetPresenceKey(sheet.ChartTextObjectCount > 0)}"
                + $"|ChartRecords:{GetPresenceKey(sheet.ChartRecordCount > 0)}"
                + $"|ChartTypes:{GetPresenceKey(sheet.ChartRecordsByChartType.Count > 0)}";
        }

        private static string GetChartSheetStateKey(LegacyXlsChartSheet sheet) {
            return $"PrintSize:{GetPresenceKey(sheet.ChartPrintSize.HasValue)}"
                + $"|TextObjects:{GetPresenceKey(sheet.ChartTextObjectCount > 0)}"
                + $"|ChartRecords:{GetPresenceKey(sheet.ChartRecordCount > 0)}"
                + $"|ChartTypes:{GetPresenceKey(sheet.ChartRecordsByChartType.Count > 0)}";
        }

        private static IEnumerable<string> GetChartWorkbookStateKeys(
            IReadOnlyCollection<LegacyXlsChartRecord> records,
            IReadOnlyCollection<LegacyXlsChartSheet> chartSheets) {
            if (records.Count == 0) {
                yield break;
            }

            var chartSheetNames = new HashSet<string>(
                chartSheets.Select(sheet => sheet.Name),
                StringComparer.OrdinalIgnoreCase);

            yield return $"Containers:{GetPresenceKey(records.Any(record => record.Kind == LegacyXlsChartRecordKind.Container))}"
                + $"|ChartTypes:{GetPresenceKey(records.Any(record => record.Kind == LegacyXlsChartRecordKind.ChartType))}"
                + $"|Series:{GetPresenceKey(records.Any(record => record.Kind == LegacyXlsChartRecordKind.Series))}"
                + $"|Axes:{GetPresenceKey(records.Any(record => record.Kind == LegacyXlsChartRecordKind.Axis))}"
                + $"|Text:{GetPresenceKey(records.Any(record => record.Kind == LegacyXlsChartRecordKind.Text))}"
                + $"|Formatting:{GetPresenceKey(records.Any(record => record.Kind == LegacyXlsChartRecordKind.Formatting))}"
                + $"|Layout:{GetPresenceKey(records.Any(record => record.Kind == LegacyXlsChartRecordKind.Layout))}"
                + $"|Future:{GetPresenceKey(records.Any(record => record.Kind == LegacyXlsChartRecordKind.FutureMetadata))}"
                + $"|PreserveOnly:{GetPresenceKey(records.Any(record => record.Kind == LegacyXlsChartRecordKind.PreserveOnly))}"
                + $"|Scopes:{GetChartRecordScopeKey(records, chartSheetNames)}";
        }

        private static string GetChartRecordScopeKey(
            IEnumerable<LegacyXlsChartRecord> records,
            ISet<string> chartSheetNames) {
            bool hasWorkbookRecords = false;
            bool hasWorksheetRecords = false;
            bool hasChartSheetRecords = false;
            foreach (LegacyXlsChartRecord record in records) {
                if (string.IsNullOrWhiteSpace(record.SheetName)) {
                    hasWorkbookRecords = true;
                } else if (chartSheetNames.Contains(record.SheetName!)) {
                    hasChartSheetRecords = true;
                } else {
                    hasWorksheetRecords = true;
                }
            }

            if (hasWorkbookRecords || (hasWorksheetRecords && hasChartSheetRecords)) {
                return $"Workbook:{GetPresenceKey(hasWorkbookRecords)};Worksheets:{GetPresenceKey(hasWorksheetRecords)};ChartSheets:{GetPresenceKey(hasChartSheetRecords)}";
            }

            return hasChartSheetRecords ? "ChartSheetsOnly" : "WorksheetsOnly";
        }

        private static string GetPivotTableRecordLocationKey(LegacyXlsPivotTableRecord record) {
            return string.IsNullOrWhiteSpace(record.SheetName) ? "(workbook)" : record.SheetName!;
        }

        private static IEnumerable<string> GetPivotTableWorkbookStateKeys(IReadOnlyCollection<LegacyXlsPivotTableRecord> records) {
            if (records.Count == 0) {
                yield break;
            }

            yield return $"View:{GetPresenceKey(records.Any(record => record.Kind == LegacyXlsPivotTableRecordKind.View))}"
                + $"|Cache:{GetPresenceKey(records.Any(record => record.Kind == LegacyXlsPivotTableRecordKind.Cache))}"
                + $"|CacheSource:{GetPresenceKey(records.Any(record => record.Kind == LegacyXlsPivotTableRecordKind.CacheSource))}"
                + $"|CacheItems:{GetPresenceKey(records.Any(record => record.Kind == LegacyXlsPivotTableRecordKind.CacheItem))}"
                + $"|Fields:{GetPresenceKey(records.Any(record => record.Kind == LegacyXlsPivotTableRecordKind.Field))}"
                + $"|Items:{GetPresenceKey(records.Any(record => record.Kind == LegacyXlsPivotTableRecordKind.Item))}"
                + $"|DataItems:{GetPresenceKey(records.Any(record => record.Kind == LegacyXlsPivotTableRecordKind.DataItem))}"
                + $"|Grouping:{GetPresenceKey(records.Any(record => record.Kind == LegacyXlsPivotTableRecordKind.GroupingRange))}"
                + $"|Formulas:{GetPresenceKey(records.Any(record => record.Kind == LegacyXlsPivotTableRecordKind.Formula))}"
                + $"|Additional:{GetPresenceKey(records.Any(record => record.Kind == LegacyXlsPivotTableRecordKind.Additional))}"
                + $"|Locations:{GetPivotTableLocationScopeKey(records)}";
        }

        private static string GetPivotTableLocationScopeKey(IEnumerable<LegacyXlsPivotTableRecord> records) {
            bool hasWorkbookRecords = false;
            bool hasWorksheetRecords = false;
            foreach (LegacyXlsPivotTableRecord record in records) {
                if (string.IsNullOrWhiteSpace(record.SheetName)) {
                    hasWorkbookRecords = true;
                } else {
                    hasWorksheetRecords = true;
                }
            }

            if (hasWorkbookRecords && hasWorksheetRecords) {
                return "WorkbookAndSheets";
            }

            return hasWorkbookRecords ? "WorkbookOnly" : "SheetsOnly";
        }

        private static string GetChartValueRangeScaleKey(LegacyXlsChartRecord record) {
            LegacyXlsChartValueRange valueRange = record.ValueRange!;
            return $"Min:{FormatDouble(valueRange.Minimum)};Max:{FormatDouble(valueRange.Maximum)};Major:{FormatDouble(valueRange.MajorUnit)};Minor:{FormatDouble(valueRange.MinorUnit)};Cross:{FormatDouble(valueRange.CrossingValue)}";
        }

        private static string GetChartValueRangeStateKey(LegacyXlsChartRecord record) {
            LegacyXlsChartValueRange valueRange = record.ValueRange!;
            return $"AutoMin:{valueRange.AutoMinimum};AutoMax:{valueRange.AutoMaximum};AutoMajor:{valueRange.AutoMajorUnit};AutoMinor:{valueRange.AutoMinorUnit};AutoCross:{valueRange.AutoCrossingValue};Log:{valueRange.LogarithmicScale};Reversed:{valueRange.Reversed};MaxCross:{valueRange.MaximumCrossing}";
        }

        private static string GetChartCategorySeriesRangeIntervalKey(LegacyXlsChartRecord record) {
            LegacyXlsChartCategorySeriesRange range = record.CategorySeriesRange!;
            return $"Cross:{range.CrossingCategory};Labels:{range.LabelInterval};Ticks:{range.TickInterval}";
        }

        private static string GetChartCategorySeriesRangeStateKey(LegacyXlsChartRecord record) {
            LegacyXlsChartCategorySeriesRange range = record.CategorySeriesRange!;
            return $"Between:{range.CrossesBetweenTickMarks};MaxCross:{range.CrossesAtMaximum};Reversed:{range.Reversed}";
        }

        private static string GetChartAxisExtensionDateRangeKey(LegacyXlsChartRecord record) {
            LegacyXlsChartAxisExtension extension = record.AxisExtension!;
            return $"Min:{extension.MinimumDate};Max:{extension.MaximumDate};Cross:{extension.CrossingDate}";
        }

        private static string GetChartAxisExtensionDateUnitKey(LegacyXlsChartRecord record) {
            LegacyXlsChartAxisExtension extension = record.AxisExtension!;
            return $"Major:{extension.MajorInterval} {extension.MajorUnitName};Minor:{extension.MinorInterval} {extension.MinorUnitName};Base:{extension.BaseUnitName}";
        }

        private static string GetChartAxisExtensionStateKey(LegacyXlsChartRecord record) {
            LegacyXlsChartAxisExtension extension = record.AxisExtension!;
            return $"AutoMin:{extension.AutoMinimum};AutoMax:{extension.AutoMaximum};AutoMajor:{extension.AutoMajor};AutoMinor:{extension.AutoMinor};DateAxis:{extension.DateAxis};AutoBase:{extension.AutoBase};AutoCross:{extension.AutoCrossing};AutoDate:{extension.AutoDateAxis}";
        }

        private static string GetChartDataSourceStateKey(LegacyXlsChartRecord record) {
            LegacyXlsChartDataSource dataSource = record.DataSource!;
            return $"Source:{dataSource.SourceIdName};Reference:{dataSource.ReferenceTypeName};CustomNumberFormat:{dataSource.UsesCustomNumberFormat};FormulaBytes:{dataSource.FormulaByteCount};FormulaComplete:{dataSource.FormulaByteCountFitsPayload};FormulaTextProjected:{dataSource.FormulaTextProjected}";
        }

        private static string GetChartDataSourceFormulaProjectionStateKey(LegacyXlsChartRecord record) {
            LegacyXlsChartDataSource dataSource = record.DataSource!;
            if (dataSource.FormulaTextProjected) {
                return "FormulaTextProjected";
            }

            return dataSource.FormulaByteCount == 0 ? "NoFormulaBytes" : "FormulaTextUnsupported";
        }

        private static string GetChartSheetPropertyStateKey(LegacyXlsChartRecord record) {
            LegacyXlsChartSheetProperties properties = record.SheetProperties!;
            return $"AutoSeries:{properties.AutomaticallyAllocateSeries};VisibleOnly:{properties.PlotVisibleCellsOnly};DoNotSizeWithWindow:{properties.DoNotSizeWithWindow};ManualPlotArea:{properties.ManualPlotArea};AlwaysAutoPlotArea:{properties.AlwaysAutoPlotArea}";
        }

        private static string GetChartBarStateKey(LegacyXlsChartRecord record) {
            LegacyXlsChartBarOptions options = record.BarOptions!;
            return $"Transposed:{options.IsTransposed};Stacked:{options.IsStacked};Percent:{options.IsPercentStacked};Shadow:{options.HasShadow}";
        }

        private static string GetChartErrorBarStateKey(LegacyXlsChartRecord record) {
            LegacyXlsChartErrorBarOptions options = record.ErrorBarOptions!;
            return $"Direction:{options.DirectionName};Source:{options.ValueSourceName};Tee:{options.HasTeeTop};UsesValue:{options.UsesValue};UsesCustomCount:{options.UsesCustomValueCount}";
        }

        private static string GetChartLineStateKey(LegacyXlsChartRecord record) {
            LegacyXlsChartLineOptions options = record.LineOptions!;
            return $"Stacked:{options.IsStacked};Percent:{options.IsPercentStacked};Shadow:{options.HasShadow}";
        }

        private static string GetChartAreaStateKey(LegacyXlsChartRecord record) {
            LegacyXlsChartAreaOptions options = record.AreaOptions!;
            return $"Stacked:{options.IsStacked};Percent:{options.IsPercentStacked};Shadow:{options.HasShadow}";
        }

        private static string GetChartBopPopStateKey(LegacyXlsChartRecord record) {
            LegacyXlsChartBopPopOptions options = record.BopPopOptions!;
            return $"Subtype:{options.SubtypeName};Split:{options.SplitName};Auto:{options.AutomaticSplit};Shadow:{options.HasShadow}";
        }

        private static string GetChartBopPopCustomStateKey(LegacyXlsChartRecord record) {
            LegacyXlsChartBopPopCustomSplit split = record.BopPopCustomSplit!;
            return $"DataPoints:{split.DataPointCount};Secondary:{split.SecondaryDataPointIndexes.Count};NoSecondary:{split.NoSecondaryDataPointsMarker};Consistent:{split.HasConsistentNoSecondaryDataPointsMarker}";
        }

        private static string GetChartThreeDimensionalBarShapeStateKey(LegacyXlsChartRecord record) {
            LegacyXlsChart3DBarShapeOptions options = record.ThreeDimensionalBarShapeOptions!;
            return $"Riser:{options.RiserName};Taper:{options.TaperName}";
        }

        private static string GetChartThreeDimensionalStateKey(LegacyXlsChartRecord record) {
            LegacyXlsChart3DOptions options = record.ThreeDimensionalOptions!;
            return $"Perspective:{options.UsesPerspective};Clustered:{options.IsClustered};AutoScale:{options.UsesAutomaticScaling};Shape:{options.ChartGroupShapeName};Walls2D:{options.UsesTwoDimensionalWalls}";
        }

        private static string GetChartScatterStateKey(LegacyXlsChartRecord record) {
            LegacyXlsChartScatterOptions options = record.ScatterOptions!;
            return $"Bubble:{options.IsBubbleChart};NegativeBubbles:{options.ShowNegativeBubbles};Shadow:{options.HasShadow};Size:{options.BubbleSizeRepresentationName}";
        }

        private static string GetChartFontBasisStateKey(LegacyXlsChartRecord record) {
            LegacyXlsChartFontBasisOptions options = record.FontBasisOptions!;
            return $"Basis:{options.WidthTwipsBasis}x{options.HeightTwipsBasis};HeightTwips:{options.FontHeightTwips};Scale:{options.ScaleBasisName};FontIndex:{options.FontIndex}";
        }

        private static string GetChartAttachedLabelStateKey(LegacyXlsChartRecord record) {
            LegacyXlsChartAttachedLabel attachedLabel = record.AttachedLabel!;
            return $"ShowValue:{attachedLabel.ShowValue};ShowPercent:{attachedLabel.ShowPercent};ShowLabelAndPercent:{attachedLabel.ShowLabelAndPercent};ShowLabel:{attachedLabel.ShowLabel};ShowBubbleSizes:{attachedLabel.ShowBubbleSizes};ShowSeriesName:{attachedLabel.ShowSeriesName}";
        }

        private static string GetChartDataFormatStateKey(LegacyXlsChartRecord record) {
            return $"Target:{record.DataFormatTarget ?? "Unknown"};PointIndex:{record.DataFormatPointIndex?.ToString(CultureInfo.InvariantCulture) ?? "Missing"};SeriesIndex:{record.DataFormatSeriesIndex?.ToString(CultureInfo.InvariantCulture) ?? "Missing"};Order:{record.DataFormatOrder?.ToString(CultureInfo.InvariantCulture) ?? "Missing"}";
        }

        private static string GetChartLineFormatStateKey(LegacyXlsChartRecord record) {
            LegacyXlsChartLineFormat lineFormat = record.LineFormat!;
            return $"Style:{lineFormat.StyleName};Weight:{lineFormat.WeightName};Automatic:{lineFormat.Automatic};AxisVisible:{lineFormat.AxisVisible};AutomaticColor:{lineFormat.AutomaticColor}";
        }

        private static IEnumerable<string> GetChartAreaFormatColorKeys(LegacyXlsChartAreaFormat areaFormat) {
            yield return $"Foreground:{areaFormat.ForegroundRgbHex}";
            yield return $"Background:{areaFormat.BackgroundRgbHex}";
        }

        private static IEnumerable<string> GetChartAreaFormatColorIndexKeys(LegacyXlsChartAreaFormat areaFormat) {
            yield return $"ForegroundIndex:{areaFormat.ForegroundColorIndex}";
            yield return $"BackgroundIndex:{areaFormat.BackgroundColorIndex}";
        }

        private static string GetChartAreaFormatStateKey(LegacyXlsChartRecord record) {
            LegacyXlsChartAreaFormat areaFormat = record.AreaFormat!;
            return $"Pattern:{areaFormat.PatternName};Automatic:{areaFormat.Automatic};InvertNegative:{areaFormat.InvertNegative}";
        }

        private static IEnumerable<string> GetChartMarkerFormatColorKeys(LegacyXlsChartMarkerFormat markerFormat) {
            yield return $"Foreground:{markerFormat.ForegroundRgbHex}";
            yield return $"Background:{markerFormat.BackgroundRgbHex}";
        }

        private static IEnumerable<string> GetChartMarkerFormatColorIndexKeys(LegacyXlsChartMarkerFormat markerFormat) {
            yield return $"ForegroundIndex:{markerFormat.ForegroundColorIndex}";
            yield return $"BackgroundIndex:{markerFormat.BackgroundColorIndex}";
        }

        private static string GetChartMarkerFormatStateKey(LegacyXlsChartRecord record) {
            LegacyXlsChartMarkerFormat markerFormat = record.MarkerFormat!;
            return $"Type:{markerFormat.MarkerTypeName};Automatic:{markerFormat.Automatic};InteriorHidden:{markerFormat.InteriorHidden};BorderHidden:{markerFormat.BorderHidden}";
        }

        private static string GetChartSeriesFormatStateKey(LegacyXlsChartRecord record) {
            LegacyXlsChartSeriesFormat seriesFormat = record.SeriesFormat!;
            return $"SmoothLine:{seriesFormat.SmoothLine};ThreeDimensionalBubbles:{seriesFormat.ThreeDimensionalBubbles};Shadow:{seriesFormat.Shadow}";
        }

        private static IEnumerable<string> GetChartClientColorPaletteColorKeys(LegacyXlsChartClientColorPalette palette) {
            if (!string.IsNullOrWhiteSpace(palette.ForegroundColor)) {
                yield return $"Foreground:{palette.ForegroundColor}";
            }

            if (!string.IsNullOrWhiteSpace(palette.BackgroundColor)) {
                yield return $"Background:{palette.BackgroundColor}";
            }

            if (!string.IsNullOrWhiteSpace(palette.NeutralColor)) {
                yield return $"Neutral:{palette.NeutralColor}";
            }
        }

        private static string FormatDouble(double value) {
            return value.ToString("G15", CultureInfo.InvariantCulture);
        }

        private static string GetDiagnosticSheetKey(LegacyXlsImportDiagnostic diagnostic) {
            return string.IsNullOrWhiteSpace(diagnostic.SheetName) ? "(workbook)" : diagnostic.SheetName!;
        }

        private static string GetDiagnosticFormulaContextKey(LegacyXlsImportDiagnostic diagnostic) {
            return string.IsNullOrWhiteSpace(diagnostic.FormulaContext) ? "Unknown" : diagnostic.FormulaContext!;
        }

        private static string GetFormulaTokenSheetKey(LegacyXlsFormulaTokenRecord record) {
            return string.IsNullOrWhiteSpace(record.SheetName) ? "(workbook)" : record.SheetName!;
        }

        private static string GetDrawingRecordLocationKey(LegacyXlsDrawingRecord record) {
            return string.IsNullOrWhiteSpace(record.SheetName) ? "(workbook)" : record.SheetName!;
        }

        private static string GetDrawingFutureRecordRangeKey(LegacyXlsDrawingRecord record) {
            LegacyXlsDrawingFutureRecordHeader header = record.FutureRecordHeader!;
            string start = A1.CellReference(header.FirstRow!.Value + 1, header.FirstColumn!.Value + 1);
            string end = A1.CellReference(header.LastRow!.Value + 1, header.LastColumn!.Value + 1);
            return $"{record.RecordName}|{start}:{end}";
        }

        private static IEnumerable<string> GetDrawingTextObjectFlagKeys(LegacyXlsDrawingTextObject textObject) {
            yield return $"TextInContinueRecords:{GetPresenceKey(textObject.HasTextInContinueRecords)}";
            yield return $"FormattingRunsInContinueRecords:{GetPresenceKey(textObject.HasFormattingRunsInContinueRecords)}";
            yield return $"DecodedText:{GetPresenceKey(textObject.HasDecodedText)}";
            yield return $"DecodedFormattingRuns:{textObject.FormattingRuns.Count}";
            yield return $"LockedText:{textObject.LockedText}";
            yield return $"JustifyLastLine:{textObject.JustifyLastLine}";
            yield return $"SecretEdit:{textObject.SecretEdit}";
        }

        private static string GetDrawingAnchorRangeKey(LegacyXlsDrawingAnchor anchor) {
            return $"R{anchor.StartRow}C{anchor.StartColumn}:R{anchor.EndRow}C{anchor.EndColumn}";
        }

        private static string GetDrawingAnchorOffsetKey(LegacyXlsDrawingAnchor anchor) {
            return $"StartDx:{anchor.StartDx};StartDy:{anchor.StartDy};EndDx:{anchor.EndDx};EndDy:{anchor.EndDy}";
        }

        private static string GetDrawingAnchorFlagsKey(LegacyXlsDrawingAnchor anchor) {
            return $"Flags:0x{anchor.Flags:X4}";
        }

        private static bool IsBlipShapeProperty(LegacyXlsDrawingShapeProperty property) {
            return string.Equals(property.PropertyGroupName, "Blip", StringComparison.OrdinalIgnoreCase);
        }

        private static bool IsPictureBlipReferenceProperty(LegacyXlsDrawingShapeProperty property) {
            return string.Equals(property.PropertyName, "pib", StringComparison.OrdinalIgnoreCase);
        }

        private static IEnumerable<string> GetDrawingPictureStateKeys(LegacyXlsWorkbook workbook) {
            bool hasPictureObjects = workbook.DrawingRecords.Any(record => record.ObjectTypeKind == LegacyXlsDrawingObjectType.Picture);
            int blipStoreEntryCount = workbook.DrawingRecords.Sum(record => record.BlipStoreEntries.Count);
            uint[] pictureBlipReferences = workbook.DrawingRecords
                .SelectMany(record => record.ShapeProperties)
                .Where(IsPictureBlipReferenceProperty)
                .Select(property => property.Value)
                .ToArray();
            bool hasPictureBlipReferences = pictureBlipReferences.Length > 0;
            if (!hasPictureObjects && blipStoreEntryCount == 0 && !hasPictureBlipReferences) {
                yield break;
            }

            yield return $"PictureObjects:{GetPresenceKey(hasPictureObjects)}"
                + $"|BlipStore:{GetPresenceKey(blipStoreEntryCount > 0)}"
                + $"|PictureBlipReferences:{GetPresenceKey(hasPictureBlipReferences)}"
                + $"|ReferencedBlips:{GetDrawingPictureBlipResolutionKey(pictureBlipReferences, blipStoreEntryCount)}";
        }

        private static IEnumerable<string> GetDrawingPictureCountStateKeys(LegacyXlsWorkbook workbook) {
            int pictureObjectCount = workbook.DrawingRecords.Count(record => record.ObjectTypeKind == LegacyXlsDrawingObjectType.Picture);
            int blipStoreEntryCount = workbook.DrawingRecords.Sum(record => record.BlipStoreEntries.Count);
            uint[] pictureBlipReferences = workbook.DrawingRecords
                .SelectMany(record => record.ShapeProperties)
                .Where(IsPictureBlipReferenceProperty)
                .Select(property => property.Value)
                .ToArray();
            int pictureFrameCount = workbook.DrawingRecords
                .SelectMany(record => record.ShapeEntries)
                .Count(shape => string.Equals(shape.ShapeTypeName, "PictureFrame", StringComparison.Ordinal));
            if (pictureObjectCount == 0 && blipStoreEntryCount == 0 && pictureBlipReferences.Length == 0 && pictureFrameCount == 0) {
                yield break;
            }

            yield return $"PictureObjects:{pictureObjectCount}"
                + $"|BlipStoreEntries:{blipStoreEntryCount}"
                + $"|PictureBlipReferences:{pictureBlipReferences.Length}"
                + $"|PictureFrames:{pictureFrameCount}"
                + $"|ObjectBlipParity:{GetCountParityKey(pictureObjectCount, blipStoreEntryCount, "Objects", "Blips")}"
                + $"|ObjectFrameCoverage:{GetCountParityKey(pictureObjectCount, pictureFrameCount, "Objects", "Frames")}"
                + $"|ReferencedBlips:{GetDrawingPictureBlipResolutionKey(pictureBlipReferences, blipStoreEntryCount)}";
        }

        private static string GetDrawingPictureBlipResolutionKey(IReadOnlyCollection<uint> pictureBlipReferences, int blipStoreEntryCount) {
            if (pictureBlipReferences.Count == 0) {
                return "None";
            }

            bool hasResolvedReference = false;
            bool hasMissingReference = false;
            foreach (uint reference in pictureBlipReferences) {
                bool isResolved = reference >= 1 && reference <= blipStoreEntryCount;
                hasResolvedReference |= isResolved;
                hasMissingReference |= !isResolved;
            }

            if (hasResolvedReference && hasMissingReference) {
                return "Partial";
            }

            return hasResolvedReference ? "Resolved" : "Missing";
        }

        private static string GetCountParityKey(int primaryCount, int secondaryCount, string primaryName, string secondaryName) {
            if (primaryCount == secondaryCount) {
                return "Balanced";
            }

            return primaryCount > secondaryCount
                ? $"{primaryName}Exceed{secondaryName}"
                : $"{secondaryName}Exceed{primaryName}";
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

        private static string GetDataConsolidationReferenceSourcePrefixKey(LegacyXlsDataConsolidationReference reference) {
            return reference.SourcePrefix.HasValue
                ? $"Prefix:0x{reference.SourcePrefix.Value:X2}"
                : "Prefix:None";
        }

        private static string GetDataConsolidationNameSourceKey(LegacyXlsDataConsolidationName name) {
            return string.IsNullOrWhiteSpace(name.Source)
                ? "(self)"
                : name.Source;
        }

        private static string NormalizeExternalReferenceTarget(string target) {
            return target.Length > 0 && target[0] == '\u0001'
                ? target.Substring(1)
                : target;
        }

        private static string GetExternalReferenceShapeKey(LegacyXlsExternalReference reference) {
            return $"{reference.Kind}|Sheets:{reference.SheetNameCount}|Names:{reference.ExternalNameCount}|Caches:{reference.CachedCellCacheCount}|CachedCells:{reference.CachedCellCount}";
        }

        private static string GetExternalNameFlagShapeKey(LegacyXlsExternalName name) {
            return $"Body:{name.BodyKind}"
                + $"|BuiltIn:{GetPresenceKey(name.BuiltIn)}"
                + $"|Advise:{GetPresenceKey(name.WantsAdvise)}"
                + $"|Picture:{GetPresenceKey(name.WantsPicture)}"
                + $"|Ole:{GetPresenceKey(name.Ole)}"
                + $"|OleLink:{GetPresenceKey(name.OleLink)}"
                + $"|Icon:{GetPresenceKey(name.Icon)}";
        }

        private static IEnumerable<string> GetExternalReferenceWorkbookStateKeys(IReadOnlyCollection<LegacyXlsExternalReference> references) {
            if (references.Count == 0) {
                yield break;
            }

            yield return $"ExternalWorkbooks:{GetPresenceKey(references.Any(reference => reference.Kind == LegacyXlsExternalReferenceKind.ExternalWorkbook))}"
                + $"|Self:{GetPresenceKey(references.Any(reference => reference.Kind == LegacyXlsExternalReferenceKind.Self))}"
                + $"|AddIns:{GetPresenceKey(references.Any(reference => reference.Kind == LegacyXlsExternalReferenceKind.AddIn))}"
                + $"|DdeOle:{GetPresenceKey(references.Any(reference => reference.Kind == LegacyXlsExternalReferenceKind.DdeOrOle))}"
                + $"|SheetTables:{GetPresenceKey(references.Any(reference => reference.SheetNameCount > 0))}"
                + $"|ExternalNames:{GetPresenceKey(references.Any(reference => reference.ExternalNameCount > 0))}"
                + $"|CellCaches:{GetPresenceKey(references.Any(reference => reference.CachedCellCacheCount > 0))}"
                + $"|CachedCells:{GetPresenceKey(references.Any(reference => reference.CachedCellCount > 0))}"
                + $"|CacheLinks:{GetExternalReferenceCacheLinkStateKey(references)}";
        }

        private static string GetExternalReferenceCacheLinkStateKey(IEnumerable<LegacyXlsExternalReference> references) {
            LegacyXlsExternalCellCache[] caches = references
                .SelectMany(reference => reference.CachedCellCaches)
                .ToArray();
            if (caches.Length == 0) {
                return "None";
            }

            bool hasValid = caches.Any(cache => cache.LinkValid);
            bool hasInvalid = caches.Any(cache => !cache.LinkValid);
            if (hasValid && hasInvalid) {
                return "Mixed";
            }

            return hasValid ? "AllValid" : "AllInvalid";
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

        private static string GetPivotTableDataItemDisplayCalculationReferenceStateKey(LegacyXlsPivotTableRecord record) {
            return $"{record.DisplayCalculationName}|Field:{record.DisplayCalculationFieldReferenceName}|Item:{record.DisplayCalculationItemReferenceName}";
        }

        private static string GetPivotTableRuleFilterStateKey(LegacyXlsPivotRuleFilter filter) {
            return $"Axis:{filter.AxisName}"
                + $"|Position:{filter.FieldPosition}"
                + $"|Field:{filter.FieldReferenceName}"
                + $"|Selected:{filter.Selected}"
                + $"|Subtotals:0x{filter.SubtotalFlags:X4}"
                + $"|Indexes:{filter.ItemIndexCount}";
        }

        private static string GetPivotTableGroupingNumericRangeKey(LegacyXlsPivotTableRecord record) {
            return $"Start:{FormatDouble(record.GroupingNumericStart!.Value)};End:{FormatDouble(record.GroupingNumericEnd!.Value)};Interval:{FormatDouble(record.GroupingNumericInterval!.Value)}";
        }

        private static string GetPivotTableGroupingDateRangeKey(LegacyXlsPivotTableRecord record) {
            return $"Start:{record.GroupingDateStart};End:{record.GroupingDateEnd};Interval:{record.GroupingDateInterval!.Value}";
        }

        private static string GetPivotTableGroupingCompletionStateKey(LegacyXlsPivotTableRecord record) {
            if (record.GroupingKind == LegacyXlsPivotGroupingKind.Numeric) {
                return record.GroupingNumericStart.HasValue
                    && record.GroupingNumericEnd.HasValue
                    && record.GroupingNumericInterval.HasValue
                        ? "CompleteNumericRange"
                        : "IncompleteNumericRange";
            }

            return record.GroupingDateStart != null
                && record.GroupingDateEnd != null
                && record.GroupingDateInterval.HasValue
                    ? "CompleteDateRange"
                    : "IncompleteDateRange";
        }

        private static string GetPivotTableGroupingStateKey(LegacyXlsPivotTableRecord record) {
            return $"Kind:{record.GroupingKind!.Value}"
                + $"|AutoStart:{record.AutoStart!.Value}"
                + $"|AutoEnd:{record.AutoEnd!.Value}"
                + $"|{GetPivotTableGroupingCompletionStateKey(record)}";
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

        private static IEnumerable<string> GetWorksheetFeatureStateKeys(LegacyXlsWorksheet sheet) {
            if (sheet.DataValidations.Count == 0
                && sheet.ConditionalFormattings.Count == 0
                && sheet.AutoFilterCriteria.Count == 0
                && !sheet.AutoFilterDropDownCount.HasValue) {
                yield break;
            }

            string dropDownCount = sheet.AutoFilterDropDownCount.HasValue
                ? sheet.AutoFilterDropDownCount.Value.ToString(CultureInfo.InvariantCulture)
                : "Missing";
            yield return $"DataValidations:{sheet.DataValidations.Count.ToString(CultureInfo.InvariantCulture)}"
                + $"|ConditionalFormatting:{sheet.ConditionalFormattings.Count.ToString(CultureInfo.InvariantCulture)}"
                + $"|AutoFilterCriteria:{sheet.AutoFilterCriteria.Count.ToString(CultureInfo.InvariantCulture)}"
                + $"|AutoFilterDropDowns:{dropDownCount}";
        }

        private static IEnumerable<string> GetDataValidationCollectionStateKeys(LegacyXlsWorksheet sheet) {
            if (sheet.DataValidationCollections.Count == 0) {
                yield break;
            }

            int parsedCount = sheet.DataValidations.Count;
            foreach (LegacyXlsDataValidationCollectionRecord collection in sheet.DataValidationCollections) {
                string matchState = collection.DeclaredValidationCount == parsedCount ? "Matched" : "Mismatched";
                yield return $"Declared:{collection.DeclaredValidationCount.ToString(CultureInfo.InvariantCulture)}"
                    + $"|Parsed:{parsedCount.ToString(CultureInfo.InvariantCulture)}"
                    + $"|{matchState}";
            }
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

        private static string GetPivotTableExtendedFieldPermissionStateKey(LegacyXlsPivotTableRecord record) {
            return $"ShowAllItems:{record.ShowAllItems!.Value}"
                + $"|Row:{record.CanDragToRow!.Value}"
                + $"|Column:{record.CanDragToColumn!.Value}"
                + $"|Page:{record.CanDragToPage!.Value}"
                + $"|Hide:{record.CanDragToHide!.Value}"
                + $"|PreventData:{record.PreventDragToData!.Value}"
                + $"|ServerBased:{record.ServerBased!.Value}";
        }

        private static IEnumerable<string> GetPivotTableCachePropertyFlagKeys(LegacyXlsPivotTableRecord record) {
            if (!record.CacheHasRecords.HasValue) {
                yield break;
            }

            yield return $"HasRecords:{record.CacheHasRecords.Value}";
            yield return $"Invalid:{record.CacheInvalid!.Value}";
            yield return $"RefreshOnLoad:{record.CacheRefreshOnLoad!.Value}";
            yield return $"OptimizeMemory:{record.CacheOptimizeMemory!.Value}";
            yield return $"BackgroundQuery:{record.CacheBackgroundQuery!.Value}";
            yield return $"EnableRefresh:{record.CacheEnableRefresh!.Value}";
        }

        private static string GetPivotTableAdditionalClassTypeKey(LegacyXlsPivotTableRecord record) {
            return $"{record.AdditionalClassName}|{record.AdditionalTypeName}";
        }

        private static IEnumerable<string> GetConditionalFormattingDifferentialFillKeys(LegacyXlsConditionalFormatting formatting) {
            LegacyXlsDifferentialFormat? format = formatting.DifferentialFormat;
            if (format == null) {
                yield break;
            }

            foreach (string key in GetDifferentialFormatFillKeys(format)) {
                yield return key;
            }
        }

        private static IEnumerable<string> GetDifferentialFormatFillKeys(LegacyXlsDifferentialFormat format) {
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

        private static IEnumerable<string> GetConditionalFormattingDifferentialFontKeys(LegacyXlsConditionalFormatting formatting) {
            LegacyXlsDifferentialFormat? format = formatting.DifferentialFormat;
            if (format == null) {
                yield break;
            }

            foreach (string key in GetDifferentialFormatFontKeys(format)) {
                yield return key;
            }
        }

        private static IEnumerable<string> GetDifferentialFormatFontKeys(LegacyXlsDifferentialFormat format) {
            if (!string.IsNullOrWhiteSpace(format.FontColor)) {
                yield return $"Color:{format.FontColor}";
            }

            if (format.FontBold.HasValue) {
                yield return format.FontBold.Value ? "Bold" : "NotBold";
            }

            if (format.FontItalic.HasValue) {
                yield return format.FontItalic.Value ? "Italic" : "NotItalic";
            }
        }

        private static IEnumerable<string> GetConditionalFormattingDifferentialBorderKeys(LegacyXlsConditionalFormatting formatting) {
            LegacyXlsDifferentialFormat? format = formatting.DifferentialFormat;
            if (format == null) {
                yield break;
            }

            foreach (string key in GetDifferentialFormatBorderKeys(format)) {
                yield return key;
            }
        }

        private static IEnumerable<string> GetDifferentialFormatBorderKeys(LegacyXlsDifferentialFormat format) {
            if (format.Border == null) {
                yield break;
            }

            foreach (string key in GetDifferentialBorderSideKeys("Top", format.Border.Top)) {
                yield return key;
            }

            foreach (string key in GetDifferentialBorderSideKeys("Bottom", format.Border.Bottom)) {
                yield return key;
            }

            foreach (string key in GetDifferentialBorderSideKeys("Left", format.Border.Left)) {
                yield return key;
            }

            foreach (string key in GetDifferentialBorderSideKeys("Right", format.Border.Right)) {
                yield return key;
            }
        }

        private static IEnumerable<string> GetDifferentialBorderSideKeys(string sideName, LegacyXlsDifferentialBorderSide? side) {
            if (side == null) {
                yield break;
            }

            if (side.HasStyle) {
                yield return $"{sideName}:Style:{side.Style}";
            }

            if (!string.IsNullOrWhiteSpace(side.Color)) {
                yield return $"{sideName}:Color:{side.Color}";
            }
        }

        private static IEnumerable<string> GetConditionalFormattingDifferentialNumberFormatKeys(LegacyXlsConditionalFormatting formatting) {
            LegacyXlsDifferentialFormat? format = formatting.DifferentialFormat;
            if (format == null) {
                yield break;
            }

            foreach (string key in GetDifferentialFormatNumberFormatKeys(format)) {
                yield return key;
            }
        }

        private static IEnumerable<string> GetDifferentialFormatNumberFormatKeys(LegacyXlsDifferentialFormat format) {
            if (format.NumberFormatId.HasValue) {
                yield return $"Id:{format.NumberFormatId.Value}";
            }

            if (!string.IsNullOrWhiteSpace(format.NumberFormatCode)) {
                yield return $"Code:{format.NumberFormatCode}";
            }
        }

        private static string GetConditionalFormattingExtensionStateKey(LegacyXlsConditionalFormattingExtensionRecord record) {
            return $"Cf12:{GetPresenceKey(record.IsCf12)}"
                + $"|UnprojectedFormatting:{GetPresenceKey(record.HasUnprojectedFormatting)}"
                + $"|MatchedRule:{GetPresenceKey(record.MatchedRule)}"
                + $"|Priority:{GetPresenceKey(record.Priority.HasValue)}"
                + $"|StopIfTrue:{GetConditionalFormattingExtensionStopIfTrueStateKey(record)}";
        }

        private static string GetConditionalFormattingExtensionStopIfTrueStateKey(LegacyXlsConditionalFormattingExtensionRecord record) {
            if (!record.StopIfTrue.HasValue) {
                return "Missing";
            }

            return record.StopIfTrue.Value ? "StopIfTrue" : "Continue";
        }

        private static string GetConditionalFormattingExtensionDxfProjectionStateKey(
            LegacyXlsConditionalFormattingExtensionRecord record,
            int differentialFormatCount) {
            if (record.IsCf12) {
                return "UnprojectedCf12";
            }

            if (record.HasProjectedFormatting) {
                if (record.InlineFormattingByteCount.HasValue) {
                    return $"ProjectedInlineDxfBytes:{record.InlineFormattingByteCount.Value.ToString(CultureInfo.InvariantCulture)}";
                }

                return "ProjectedSingleDxf";
            }

            if (!record.HasUnprojectedFormatting) {
                return "NoDxfRequested";
            }

            if (!record.MatchedRule) {
                return "UnprojectedUnmatchedRule";
            }

            if (differentialFormatCount == 0) {
                return "UnprojectedMissingDxf";
            }

            if (differentialFormatCount == 1) {
                return "ProjectedSingleDxf";
            }

            if (record.InlineFormattingByteCount.HasValue) {
                return $"UnprojectedInlineDxfBytes:{record.InlineFormattingByteCount.Value.ToString(CultureInfo.InvariantCulture)}";
            }

            return $"UnprojectedMultipleDxfCandidates:{differentialFormatCount.ToString(CultureInfo.InvariantCulture)}";
        }

        private static string GetDifferentialFormatContentStateKey(LegacyXlsDifferentialFormat format) {
            bool hasFill = format.FillPattern.HasValue
                || !string.IsNullOrWhiteSpace(format.FillForegroundColor)
                || !string.IsNullOrWhiteSpace(format.FillBackgroundColor);
            bool hasFont = !string.IsNullOrWhiteSpace(format.FontColor)
                || format.FontBold.HasValue
                || format.FontItalic.HasValue;
            bool hasBorder = format.Border?.HasAnySide == true;
            bool hasNumberFormat = format.NumberFormatId.HasValue
                || !string.IsNullOrWhiteSpace(format.NumberFormatCode);

            if (hasFill && hasFont && hasBorder && hasNumberFormat) {
                return "FillFontBorderAndNumberFormat";
            }

            if (hasFill && hasFont && hasBorder) {
                return "FillFontAndBorder";
            }

            if (hasFill && hasFont && hasNumberFormat) {
                return "FillFontAndNumberFormat";
            }

            if (hasFill && hasBorder && hasNumberFormat) {
                return "FillBorderAndNumberFormat";
            }

            if (hasFont && hasBorder && hasNumberFormat) {
                return "FontBorderAndNumberFormat";
            }

            if (hasFill && hasFont) {
                return "FillAndFont";
            }

            if (hasFill && hasBorder) {
                return "FillAndBorder";
            }

            if (hasFill && hasNumberFormat) {
                return "FillAndNumberFormat";
            }

            if (hasFont && hasBorder) {
                return "FontAndBorder";
            }

            if (hasFont && hasNumberFormat) {
                return "FontAndNumberFormat";
            }

            if (hasBorder && hasNumberFormat) {
                return "BorderAndNumberFormat";
            }

            if (hasFill) {
                return "FillOnly";
            }

            if (hasFont) {
                return "FontOnly";
            }

            if (hasBorder) {
                return "BorderOnly";
            }

            return hasNumberFormat ? "NumberFormatOnly" : "DecodedNoContent";
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
