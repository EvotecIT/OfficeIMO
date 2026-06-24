using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using OfficeIMO.Excel;
using OfficeIMO.Excel.LegacyXls;
using OfficeIMO.Excel.LegacyXls.Diagnostics;
using OfficeIMO.Excel.LegacyXls.Model;
using Xunit;

namespace OfficeIMO.Tests {
    public partial class Excel {
        [Fact]
        public void LegacyXls_Corpus_Fixtures_MatchApprovedImportReports() {
            string corpusDirectory = Path.Combine(GetTestsProjectRoot(), "Documents", "LegacyXlsCorpus");
            AssertLegacyXlsCorpusBaselines(corpusDirectory);
        }

        [Fact]
        public void LegacyXls_DiagnosticCorpus_Fixtures_MatchApprovedImportReports() {
            string corpusDirectory = Path.Combine(GetTestsProjectRoot(), "Documents", "LegacyXlsDiagnosticCorpus");
            AssertLegacyXlsCorpusBaselines(corpusDirectory);
        }

        [Fact]
        public void LegacyXls_Corpus_OpenPreserveValid_PreservesVbaProjectShape() {
            string workbookPath = Path.Combine(
                GetTestsProjectRoot(),
                "Documents",
                "LegacyXlsCorpus",
                "openpreserve-format-corpus",
                "valid.xls");

            using LegacyXlsLoadResult result = ExcelDocument.LoadLegacyXlsWithReport(workbookPath, new LegacyXlsImportOptions {
                ReportUnsupportedRecords = true
            });

            Assert.DoesNotContain(result.Diagnostics, diagnostic => diagnostic.Severity == LegacyXlsDiagnosticSeverity.Error);
            Assert.Equal(8, result.ImportReport.CompoundVbaModuleCount);
            Assert.Equal(7664, result.ImportReport.CompoundVbaModuleByteCount);
            Assert.Equal(1, result.ImportReport.CompoundVbaModulesByPath["_VBA_PROJECT_CUR/VBA/ThisWorkbook"]);
            Assert.Equal(1, result.ImportReport.CompoundVbaModulesByNameAndSize["ThisWorkbook|Bytes:965"]);
            Assert.Equal(1, result.ImportReport.CompoundVbaModulesByNameAndSize["Sheet8|Bytes:957"]);
            Assert.Equal(1, result.ImportReport.CompoundVbaModulesByCodeNameMatchAndName["WorkbookCodeName|ThisWorkbook"]);
            Assert.Equal(1, result.ImportReport.CompoundVbaModulesByCodeNameMatchAndName["WorksheetCodeName|Sheet1"]);
            Assert.Equal(1, result.ImportReport.CompoundVbaProjectsByStructure["Modules:8|DirStreams:1|ProjectStreams:2|Storages:2"]);
        }

        [Fact]
        public void LegacyXls_Corpus_ExcelExternalLinks_PreserveFormulaAndCacheModel() {
            string workbookPath = Path.Combine(
                GetTestsProjectRoot(),
                "Documents",
                "LegacyXlsCorpus",
                "excel-com-generated",
                "external-links.xls");

            using LegacyXlsLoadResult result = ExcelDocument.LoadLegacyXlsWithReport(workbookPath, new LegacyXlsImportOptions {
                ReportUnsupportedRecords = true
            });

            Assert.DoesNotContain(result.Diagnostics, diagnostic => diagnostic.Severity == LegacyXlsDiagnosticSeverity.Error);
            Assert.DoesNotContain(result.Diagnostics, diagnostic => diagnostic.Code == "XLS-BIFF-FORMULA-TOKENS-UNSUPPORTED");
            Assert.Equal(1, result.ImportReport.ExternalReferencesByTarget["external-source.xls"]);
            Assert.Equal(1, result.ImportReport.ExternalSheetNamesByTarget["external-source.xls!Data"]);
            Assert.Equal(1, result.ImportReport.ExternalCellCachesByTarget["external-source.xls"]);
            Assert.Equal(1, result.ImportReport.ExternalCellCachesByTargetAndSheetName["external-source.xls!Data"]);
            Assert.Equal(1, result.ImportReport.ExternalCellCachesByTargetAndCellRange["external-source.xls!R1C1:R3C1"]);
            Assert.Equal(3, result.ImportReport.ExternalCachedCellsByTargetSheetAndValueKind["external-source.xls!Data|Number"]);

            LegacyXlsExternalReference externalReference = Assert.Single(
                result.Workbook.ExternalReferences,
                reference => reference.Kind == LegacyXlsExternalReferenceKind.ExternalWorkbook);
            Assert.Equal("\u0001external-source.xls", externalReference.Target);
            Assert.Equal(new[] { "Data" }, externalReference.SheetNames);

            LegacyXlsExternalCellCache cache = Assert.Single(externalReference.CachedCellCaches);
            Assert.True(cache.LinkValid);
            Assert.Equal("Data", cache.SheetName);
            Assert.Equal("R1C1:R3C1", cache.CellRange);
            Assert.Collection(cache.Cells,
                cell => {
                    Assert.Equal(LegacyXlsCellValueKind.Number, cell.Kind);
                    Assert.Equal(125d, cell.Value);
                },
                cell => {
                    Assert.Equal(LegacyXlsCellValueKind.Number, cell.Kind);
                    Assert.Equal(80d, cell.Value);
                },
                cell => {
                    Assert.Equal(LegacyXlsCellValueKind.Number, cell.Kind);
                    Assert.Equal(210d, cell.Value);
                });

            LegacyXlsWorksheet sheet = Assert.Single(result.Workbook.Worksheets);
            AssertCorpusFormula(sheet, 1, 2, 125d, "'[external-source.xls]Data'!$B$2");
            AssertCorpusFormula(sheet, 2, 2, 80d, "'[external-source.xls]Data'!$B$3");
            AssertCorpusFormula(sheet, 3, 2, 45d, "B1-B2");
            AssertCorpusFormula(sheet, 5, 2, 415d, "SUM('[external-source.xls]Data'!$B$2:$B$4)");
        }

        [Fact]
        public void LegacyXls_Corpus_ExcelObjects_PreserveDrawingObjectSubrecordModel() {
            string workbookPath = Path.Combine(
                GetTestsProjectRoot(),
                "Documents",
                "LegacyXlsCorpus",
                "excel-com-generated",
                "objects.xls");

            using LegacyXlsLoadResult result = ExcelDocument.LoadLegacyXlsWithReport(workbookPath, new LegacyXlsImportOptions {
                ReportUnsupportedRecords = true
            });

            Assert.DoesNotContain(result.Diagnostics, diagnostic => diagnostic.Severity == LegacyXlsDiagnosticSeverity.Error);
            Assert.True(result.ImportReport.DrawingRecordsByObjectTypeName["Picture"] >= 1);
            Assert.True(result.ImportReport.DrawingRecordsByObjectTypeName["Button"] >= 1);
            Assert.True(result.ImportReport.DrawingRecordsByObjectTypeName["Checkbox"] >= 1);
            Assert.True(result.ImportReport.DrawingRecordsByObjectTypeName["DropdownList"] >= 1);
            Assert.True(result.ImportReport.DrawingObjectSubRecordsByName["FtCmo"] >= 1);
            Assert.True(result.ImportReport.DrawingObjectSubRecordsByName["FtEnd"] >= 1);
            Assert.True(result.ImportReport.DrawingObjectSubRecordsByName["FtCblsData"] >= 1);
            Assert.True(result.ImportReport.DrawingObjectSubRecordsByName["FtLbsData"] >= 1);
            Assert.True(result.ImportReport.DrawingObjectSubRecordsByCompleteness["Complete"] >= 1);
            Assert.True(result.ImportReport.DrawingObjectSubRecordsByCompleteness["Truncated"] >= 1);
            Assert.Equal(1, result.ImportReport.DrawingBlipStoreEntriesByEmbeddedRecordType["OfficeArtBlipPNG"]);
            Assert.Equal(1, result.ImportReport.DrawingBlipStoreEntriesByLocation["(workbook)"]);
            Assert.Equal(1, result.ImportReport.DrawingBlipStoreEntriesByTypeAndLocation["(workbook)|Png"]);
            Assert.Equal(1, result.ImportReport.DrawingShapePropertiesByName["pib"]);
            Assert.Equal(1, result.ImportReport.DrawingShapePropertiesByName["BlipBooleanProperties"]);
            Assert.Equal(1, result.ImportReport.DrawingShapePropertiesByName["ShapeBooleanProperties"]);
            Assert.Equal(2, result.ImportReport.DrawingShapePropertiesByName["wzName"]);
            Assert.Equal(2, result.ImportReport.DrawingShapePropertiesByGroup["Blip"]);
            Assert.Equal(2, result.ImportReport.DrawingShapeBlipPropertiesByLocation["Objects"]);
            Assert.Equal(1, result.ImportReport.DrawingShapeBlipPropertiesByNameAndValue["pib;Value:0x00000001"]);
            Assert.Equal(1, result.ImportReport.DrawingShapeBlipPropertiesByNameAndValue["BlipBooleanProperties;Value:0x00060000"]);
            Assert.Equal(1, result.ImportReport.DrawingPictureBlipReferencesByLocation["Objects"]);
            Assert.Equal(1, result.ImportReport.DrawingPictureBlipReferencesByValue["BlipId:1"]);
            Assert.Equal(1, result.ImportReport.DrawingShapePropertiesByGroup["Shape"]);
            Assert.False(result.ImportReport.DrawingShapePropertiesByGroup.ContainsKey("Unknown"));
            Assert.Equal(1, result.ImportReport.DrawingShapeComplexPropertiesByText["wzName:Chart 5"]);
            Assert.Equal(1, result.ImportReport.DrawingShapeComplexPropertiesByText["wzName:Picture 4"]);
            Assert.Contains(result.Workbook.DrawingRecords.SelectMany(record => record.ShapeProperties),
                property => property.PropertyName == "wzName" && property.ComplexText == "Chart 5");
            Assert.Contains(result.Workbook.DrawingRecords.SelectMany(record => record.ShapeProperties),
                property => property.PropertyName == "wzName" && property.ComplexText == "Picture 4");
            Assert.Contains(result.Workbook.DrawingRecords, record => record.ObjectTypeName == "Picture" && record.HasObjectSubRecords);
            Assert.Contains(result.Workbook.DrawingRecords, record => record.ObjectTypeName == "Button" && record.HasObjectSubRecords);
            Assert.Contains(result.Workbook.DrawingRecords, record => record.ObjectTypeName == "Checkbox" && record.ObjectSubRecords.Any(subRecord => subRecord.SubRecordName == "FtCblsData"));
        }

        [Fact]
        public void LegacyXls_Corpus_AutoFilterShapes_PreservesExcelAuthoredCriteria() {
            string workbookPath = Path.Combine(
                GetTestsProjectRoot(),
                "Documents",
                "LegacyXlsCorpus",
                "excel-com-generated",
                "auto-filter-shapes.xls");

            using LegacyXlsLoadResult result = ExcelDocument.LoadLegacyXlsWithReport(workbookPath, new LegacyXlsImportOptions {
                ReportUnsupportedRecords = true
            });

            Assert.DoesNotContain(result.Diagnostics, diagnostic => diagnostic.Severity == LegacyXlsDiagnosticSeverity.Error);
            Assert.Equal(4, result.ImportReport.AutoFilterCriteriaCount);
            Assert.Equal(1, result.ImportReport.AutoFilterCriteriaByKind["Blanks"]);
            Assert.Equal(1, result.ImportReport.AutoFilterCriteriaByKind["NonBlanks"]);
            Assert.Equal(1, result.ImportReport.AutoFilterCriteriaByKind["Top10"]);
            Assert.Equal(1, result.ImportReport.AutoFilterCriteriaByKind["Custom"]);
            Assert.Equal(1, result.ImportReport.AutoFilterCriteriaByJoinOperator["And"]);
            Assert.Equal(2, result.ImportReport.AutoFilterCriteriaByJoinOperator["Single"]);
            Assert.Equal(1, result.ImportReport.AutoFilterCriteriaByJoinOperator["None"]);
            Assert.Equal(1, result.ImportReport.AutoFilterTop10Values["TopItems:3"]);

            LegacyXlsWorksheet sheet = Assert.Single(result.Workbook.Worksheets);
            LegacyXlsAutoFilterCriteria blankCriteria = Assert.Single(sheet.AutoFilterCriteria, criteria => criteria.Kind == LegacyXlsAutoFilterKind.Blanks);
            Assert.Equal(0U, blankCriteria.ColumnId);
            Assert.Equal(LegacyXlsAutoFilterValueKind.Blank, Assert.Single(blankCriteria.Conditions).ValueKind);

            LegacyXlsAutoFilterCriteria nonBlankCriteria = Assert.Single(sheet.AutoFilterCriteria, criteria => criteria.Kind == LegacyXlsAutoFilterKind.NonBlanks);
            Assert.Equal(1U, nonBlankCriteria.ColumnId);
            Assert.Equal(LegacyXlsAutoFilterValueKind.NonBlank, Assert.Single(nonBlankCriteria.Conditions).ValueKind);

            LegacyXlsAutoFilterCriteria topCriteria = Assert.Single(sheet.AutoFilterCriteria, criteria => criteria.Kind == LegacyXlsAutoFilterKind.Top10);
            Assert.Equal(2U, topCriteria.ColumnId);
            Assert.Equal((ushort)3, topCriteria.Top10Value);
            Assert.True(topCriteria.Top10IsTop);
            Assert.False(topCriteria.Top10IsPercent);

            LegacyXlsAutoFilterCriteria amountCriteria = Assert.Single(sheet.AutoFilterCriteria, criteria => criteria.ColumnId == 3);
            Assert.Equal(LegacyXlsAutoFilterKind.Custom, amountCriteria.Kind);
            Assert.True(amountCriteria.MatchAll);
            Assert.Equal(LegacyXlsAutoFilterJoinOperator.And, amountCriteria.JoinOperator);
            Assert.Collection(amountCriteria.Conditions,
                condition => {
                    Assert.Equal(LegacyXlsAutoFilterOperator.GreaterThanOrEqual, condition.Operator);
                    Assert.Equal("10", condition.Value);
                },
                condition => {
                    Assert.Equal(LegacyXlsAutoFilterOperator.LessThanOrEqual, condition.Operator);
                    Assert.Equal("30", condition.Value);
                });

            using var output = new MemoryStream();
            result.Document.Save(output);
            using SpreadsheetDocument spreadsheet = SpreadsheetDocument.Open(new MemoryStream(output.ToArray()), false);
            WorksheetPart worksheetPart = spreadsheet.WorkbookPart!.WorksheetParts.Single();
            AutoFilter autoFilter = Assert.Single(worksheetPart.Worksheet.Elements<AutoFilter>());
            Assert.Equal("A1:D8", autoFilter.Reference!.Value);
            List<FilterColumn> columns = autoFilter.Elements<FilterColumn>().OrderBy(column => column.ColumnId!.Value).ToList();
            Assert.Equal(4, columns.Count);
            Assert.True(Assert.Single(columns[0].Elements<Filters>()).Blank!.Value);
            Assert.Equal(FilterOperatorValues.NotEqual, Assert.Single(columns[1].GetFirstChild<CustomFilters>()!.Elements<CustomFilter>()).Operator!.Value);
            Assert.Equal(3d, Assert.Single(columns[2].Elements<Top10>()).Val!.Value);
            CustomFilters amountFilters = Assert.Single(columns[3].Elements<CustomFilters>());
            Assert.True(amountFilters.And!.Value);
        }

        [Fact]
        public void LegacyXls_Corpus_DataValidationShapes_PreservesExcelAuthoredRules() {
            string workbookPath = Path.Combine(
                GetTestsProjectRoot(),
                "Documents",
                "LegacyXlsCorpus",
                "excel-com-generated",
                "data-validation-shapes.xls");

            using LegacyXlsLoadResult result = ExcelDocument.LoadLegacyXlsWithReport(workbookPath, new LegacyXlsImportOptions {
                ReportUnsupportedRecords = true
            });

            Assert.DoesNotContain(result.Diagnostics, diagnostic => diagnostic.Severity == LegacyXlsDiagnosticSeverity.Error);
            Assert.Equal(9, result.ImportReport.DataValidationCount);
            Assert.Equal(1, result.ImportReport.DataValidationsByType["WholeNumber"]);
            Assert.Equal(1, result.ImportReport.DataValidationsByType["Decimal"]);
            Assert.Equal(1, result.ImportReport.DataValidationsByType["Date"]);
            Assert.Equal(1, result.ImportReport.DataValidationsByType["Time"]);
            Assert.Equal(1, result.ImportReport.DataValidationsByType["TextLength"]);
            Assert.Equal(3, result.ImportReport.DataValidationsByType["List"]);
            Assert.Equal(1, result.ImportReport.DataValidationsByType["Custom"]);
            Assert.Equal(1, result.ImportReport.DataValidationListSourcesByKind["InlineList"]);
            Assert.Equal(1, result.ImportReport.DataValidationListSourcesByKind["Range"]);
            Assert.Equal(1, result.ImportReport.DataValidationListSourcesByKind["DefinedName"]);
            Assert.Equal(1, result.ImportReport.DataValidationListSourcesByRange["K2:K4"]);
            Assert.Equal(1, result.ImportReport.DataValidationListSourcesByName["NamedOptions"]);

            LegacyXlsWorksheet validationSheet = Assert.Single(result.Workbook.Worksheets, sheet => sheet.Name == "Validation");
            Assert.Single(result.Workbook.Worksheets, sheet => sheet.Name == "Lookup");

            AssertValidation(validationSheet, LegacyXlsDataValidationType.WholeNumber, "A2:A6", "1", "10");
            AssertValidation(validationSheet, LegacyXlsDataValidationType.Decimal, "B2:B6", "1.5", null);
            AssertValidation(validationSheet, LegacyXlsDataValidationType.Date, "C2:C6", "46023", "46387");
            AssertValidation(validationSheet, LegacyXlsDataValidationType.Time, "D2:D6", "0.375", "0.708333333333333");
            AssertValidation(validationSheet, LegacyXlsDataValidationType.TextLength, "E2:E6", "8", null);

            LegacyXlsDataValidation inlineList = AssertValidation(validationSheet, LegacyXlsDataValidationType.List, "F2:F6", "\"Open,Closed,Pending\"", null);
            Assert.Equal(LegacyXlsDataValidationListSourceKind.InlineList, inlineList.ListSourceKind);
            Assert.Equal(new[] { "Open", "Closed", "Pending" }, inlineList.ListItems);

            LegacyXlsDataValidation rangeList = AssertValidation(validationSheet, LegacyXlsDataValidationType.List, "G2:G6", "$K$2:$K$4", null);
            Assert.Equal(LegacyXlsDataValidationListSourceKind.Range, rangeList.ListSourceKind);
            Assert.Equal("K2:K4", rangeList.ListSourceRange);

            LegacyXlsDataValidation namedList = AssertValidation(validationSheet, LegacyXlsDataValidationType.List, "H2:H6", "NamedOptions", null);
            Assert.Equal(LegacyXlsDataValidationListSourceKind.DefinedName, namedList.ListSourceKind);
            Assert.Equal("NamedOptions", namedList.ListSourceName);

            LegacyXlsDataValidation custom = AssertValidation(validationSheet, LegacyXlsDataValidationType.Custom, "I2:I6", "LEN(I2)>0", null);
            Assert.Equal(LegacyXlsDataValidationOperator.Between, custom.Operator);

            using var output = new MemoryStream();
            result.Document.Save(output);
            using SpreadsheetDocument spreadsheet = SpreadsheetDocument.Open(new MemoryStream(output.ToArray()), false);
            WorksheetPart worksheetPart = spreadsheet.WorkbookPart!.WorksheetParts.Single(part => part.Worksheet.Descendants<DataValidation>().Any());
            List<DataValidation> validations = worksheetPart.Worksheet.Descendants<DataValidation>().ToList();
            Assert.Equal(9, validations.Count);
            Assert.Contains(validations, validation => validation.Type!.Value == DataValidationValues.Decimal && validation.SequenceOfReferences!.InnerText == "B2:B6" && validation.GetFirstChild<Formula1>()!.Text == "1.5");
            Assert.Contains(validations, validation => validation.Type!.Value == DataValidationValues.Custom && validation.SequenceOfReferences!.InnerText == "I2:I6" && validation.GetFirstChild<Formula1>()!.Text == "LEN(I2)>0");
            Assert.Contains(validations, validation => validation.Type!.Value == DataValidationValues.List && validation.SequenceOfReferences!.InnerText == "G2:G6" && validation.GetFirstChild<Formula1>()!.Text == "=K2:K4");
            Assert.Contains(validations, validation => validation.Type!.Value == DataValidationValues.List && validation.SequenceOfReferences!.InnerText == "H2:H6" && validation.GetFirstChild<Formula1>()!.Text == "=NamedOptions");
        }

        [Fact]
        public void LegacyXls_Corpus_ConditionalFormattingShapes_PreservesExcelAuthoredRules() {
            string workbookPath = Path.Combine(
                GetTestsProjectRoot(),
                "Documents",
                "LegacyXlsCorpus",
                "excel-com-generated",
                "conditional-format-shapes.xls");

            using LegacyXlsLoadResult result = ExcelDocument.LoadLegacyXlsWithReport(workbookPath, new LegacyXlsImportOptions {
                ReportUnsupportedRecords = true
            });

            Assert.DoesNotContain(result.Diagnostics, diagnostic => diagnostic.Severity == LegacyXlsDiagnosticSeverity.Error);
            Assert.DoesNotContain(result.Diagnostics, diagnostic => diagnostic.Code == "XLS-BIFF-FORMULA-TOKENS-UNSUPPORTED");
            Assert.Equal(8, result.ImportReport.ConditionalFormattingCount);
            Assert.Equal(7, result.ImportReport.ConditionalFormattingsByType["CellIs"]);
            Assert.Equal(1, result.ImportReport.ConditionalFormattingsByType["Formula"]);
            foreach (string operatorName in new[] {
                "Between",
                "Equal",
                "GreaterThan",
                "GreaterThanOrEqual",
                "LessThanOrEqual",
                "NotBetween",
                "NotEqual"
            }) {
                Assert.Equal(1, result.ImportReport.ConditionalFormattingsByOperator[operatorName]);
            }

            Assert.Equal(2, result.ImportReport.ConditionalFormattingsByFormulaPairState["Formula1:Present|Formula2:Present"]);
            Assert.Equal(6, result.ImportReport.ConditionalFormattingsByFormulaPairState["Formula1:Present|Formula2:Missing"]);
            Assert.Equal(4, result.ImportReport.ConditionalFormattingsByPriorityState["Present"]);
            Assert.Equal(4, result.ImportReport.ConditionalFormattingsByPriorityState["Missing"]);
            Assert.Equal(4, result.ImportReport.ConditionalFormattingsByStopIfTrueState["StopIfTrue"]);
            Assert.Equal(4, result.ImportReport.ConditionalFormattingsByStopIfTrueState["Continue"]);
            Assert.Equal(1, result.ImportReport.UnsupportedFeaturesByDetail["ConditionalFormatting|XLS-BIFF-FEATURE-CONDITIONAL-FORMATTING-UNSUPPORTED|ConditionalFormatting:Dxf"]);

            LegacyXlsWorksheet sheet = Assert.Single(result.Workbook.Worksheets);
            AssertConditionalFormatting(sheet, LegacyXlsConditionalFormattingType.CellIs, LegacyXlsConditionalFormattingOperator.GreaterThan, "A2:A6", "50", null);
            AssertConditionalFormatting(sheet, LegacyXlsConditionalFormattingType.CellIs, LegacyXlsConditionalFormattingOperator.Between, "B2:B6", "10", "20");
            AssertConditionalFormatting(sheet, LegacyXlsConditionalFormattingType.CellIs, LegacyXlsConditionalFormattingOperator.NotBetween, "C2:C6", "5", "15");
            AssertConditionalFormatting(sheet, LegacyXlsConditionalFormattingType.Formula, null, "D2:D6", "LEN(D2)>0", null);
            AssertConditionalFormatting(sheet, LegacyXlsConditionalFormattingType.CellIs, LegacyXlsConditionalFormattingOperator.Equal, "E2:E6", "5", null);
            AssertConditionalFormatting(sheet, LegacyXlsConditionalFormattingType.CellIs, LegacyXlsConditionalFormattingOperator.NotEqual, "F2:F6", "\"\"\"B\"\"\"", null);
            AssertConditionalFormatting(sheet, LegacyXlsConditionalFormattingType.CellIs, LegacyXlsConditionalFormattingOperator.GreaterThanOrEqual, "G2:G6", "20", null);
            AssertConditionalFormatting(sheet, LegacyXlsConditionalFormattingType.CellIs, LegacyXlsConditionalFormattingOperator.LessThanOrEqual, "H2:H6", "70", null);

            using var output = new MemoryStream();
            result.Document.Save(output);
            using SpreadsheetDocument spreadsheet = SpreadsheetDocument.Open(new MemoryStream(output.ToArray()), false);
            WorksheetPart worksheetPart = spreadsheet.WorkbookPart!.WorksheetParts.Single(part => part.Worksheet.Descendants<ConditionalFormatting>().Any());
            List<ConditionalFormatting> conditionalFormattings = worksheetPart.Worksheet.Elements<ConditionalFormatting>().ToList();
            Assert.Equal(8, conditionalFormattings.Count);
            Assert.Contains(conditionalFormattings, formatting => formatting.SequenceOfReferences!.InnerText == "A2:A6"
                && Assert.Single(formatting.Elements<ConditionalFormattingRule>()).Operator!.Value == ConditionalFormattingOperatorValues.GreaterThan);
            Assert.Contains(conditionalFormattings, formatting => formatting.SequenceOfReferences!.InnerText == "B2:B6"
                && Assert.Single(formatting.Elements<ConditionalFormattingRule>()).Operator!.Value == ConditionalFormattingOperatorValues.Between
                && formatting.Descendants<Formula>().Select(formula => formula.Text).SequenceEqual(new[] { "10", "20" }));
            Assert.Contains(conditionalFormattings, formatting => formatting.SequenceOfReferences!.InnerText == "D2:D6"
                && Assert.Single(formatting.Elements<ConditionalFormattingRule>()).Type!.Value == ConditionalFormatValues.Expression
                && Assert.Single(formatting.Descendants<Formula>()).Text == "LEN(D2)>0");
        }

        [Fact]
        public void LegacyXls_Corpus_Analytics_PreservesChartSheetProperties() {
            string workbookPath = Path.Combine(
                GetTestsProjectRoot(),
                "Documents",
                "LegacyXlsCorpus",
                "excel-com-generated",
                "analytics.xls");

            using LegacyXlsLoadResult result = ExcelDocument.LoadLegacyXlsWithReport(workbookPath, new LegacyXlsImportOptions {
                ReportUnsupportedRecords = true
            });

            Assert.DoesNotContain(result.Diagnostics, diagnostic => diagnostic.Severity == LegacyXlsDiagnosticSeverity.Error);
            LegacyXlsUnsupportedSheet chartSheet = Assert.Single(result.Workbook.UnsupportedSheets, sheet => sheet.Name == "RevenueChart");
            Assert.Equal(LegacyXlsUnsupportedSheetKind.ChartSheet, chartSheet.Kind);
            LegacyXlsChartRecord sheetPropertiesRecord = Assert.Single(result.Workbook.ChartRecords, record => record.RecordName == "ShtProps" && record.SheetName == "RevenueChart");
            Assert.NotNull(sheetPropertiesRecord.SheetProperties);
            Assert.Equal(0, sheetPropertiesRecord.SheetProperties!.Flags);
            Assert.False(sheetPropertiesRecord.SheetProperties.AutomaticallyAllocateSeries);
            Assert.False(sheetPropertiesRecord.SheetProperties.PlotVisibleCellsOnly);
            Assert.False(sheetPropertiesRecord.SheetProperties.DoNotSizeWithWindow);
            Assert.False(sheetPropertiesRecord.SheetProperties.ManualPlotArea);
            Assert.False(sheetPropertiesRecord.SheetProperties.AlwaysAutoPlotArea);
            Assert.False(sheetPropertiesRecord.SheetProperties.HasKnownEmptyCellPlottingMode);
            Assert.Equal(1, result.ImportReport.ChartSheetPropertyStates["AutoSeries:False;VisibleOnly:False;DoNotSizeWithWindow:False;ManualPlotArea:False;AlwaysAutoPlotArea:False"]);
        }

        [Fact]
        public void LegacyXls_Corpus_FormulaStress_ProjectsSupportedFormulaTokens() {
            string workbookPath = Path.Combine(
                GetTestsProjectRoot(),
                "Documents",
                "LegacyXlsCorpus",
                "excel-com-generated",
                "formula-stress.xls");

            using LegacyXlsLoadResult result = ExcelDocument.LoadLegacyXlsWithReport(workbookPath, new LegacyXlsImportOptions {
                ReportUnsupportedRecords = true
            });

            Assert.DoesNotContain(result.Diagnostics, diagnostic => diagnostic.Severity == LegacyXlsDiagnosticSeverity.Error);
            Assert.DoesNotContain(result.Diagnostics, diagnostic => diagnostic.Code == "XLS-BIFF-FORMULA-TOKENS-UNSUPPORTED");
            Assert.DoesNotContain(result.Diagnostics, diagnostic => diagnostic.Code == "XLS-BIFF-FORMULA-SHARED-UNRESOLVED");
            Assert.Empty(result.ImportReport.FormulaTokenBlockers);
            Assert.Equal(4, result.ImportReport.FormulaTokensByContext["ArrayFormula"]);
            Assert.Equal(4, result.ImportReport.FormulaTokensByContextAndSheet["ArrayFormula|Formulas"]);
            Assert.Contains("ROUND", result.ImportReport.FormulaFunctionsByName.Keys);
            Assert.Contains("IF", result.ImportReport.FormulaFunctionsByName.Keys);
            Assert.Contains("If", result.ImportReport.FormulaAttributesByName.Keys);
            Assert.Contains("Sum", result.ImportReport.FormulaAttributesByName.Keys);

            foreach (string tokenName in new[] {
                "PtgAdd",
                "PtgArea",
                "PtgArray",
                "PtgAttr",
                "PtgConcat",
                "PtgDiv",
                "PtgFunc",
                "PtgFuncVar",
                "PtgGt",
                "PtgLe",
                "PtgNe",
                "PtgPercent",
                "PtgPower",
                "PtgStr",
                "PtgUminus",
                "PtgUplus"
            }) {
                Assert.Contains(tokenName, result.ImportReport.FormulaTokensByName.Keys);
            }

            LegacyXlsWorksheet sheet = Assert.Single(result.Workbook.Worksheets);
            AssertCorpusFormula(sheet, 1, 2, 100d, "A1^2");
            AssertCorpusFormula(sheet, 2, 2, "North-Q1", "A4&\"-\"&A5");
            AssertCorpusFormula(sheet, 3, 2, -3d, "-A2");
            AssertCorpusFormula(sheet, 4, 2, 5d, "+A3");
            AssertCorpusFormula(sheet, 5, 2, 0.03d, "A2%");
            AssertCorpusFormula(sheet, 6, 2, true, "A1>A2");
            AssertCorpusFormula(sheet, 7, 2, true, "A2<=A3");
            AssertCorpusFormula(sheet, 8, 2, true, "A1<>A3");
            AssertCorpusFormula(sheet, 9, 2, 6d, "SUM({1,2,3})");
            AssertCorpusFormula(sheet, 10, 2, 3.33d, "ROUND(A1/A2,2)");
            AssertCorpusFormula(sheet, 11, 2, "yes", "IF(A1>A3,\"yes\",\"no\")");
            AssertCorpusFormula(sheet, 12, 2, 18d, "SUM(A1:A3)");
            AssertCorpusFormula(sheet, 1, 3, 103.33d, "B1+B10");
            AssertCorpusFormula(sheet, 1, 4, 31d, "SUM(A1:A3*{1;2;3})");
        }

        [Fact]
        public void LegacyXls_Corpus_Protection_PreservesSheetProtectionAndCellProtection() {
            string workbookPath = Path.Combine(
                GetTestsProjectRoot(),
                "Documents",
                "LegacyXlsCorpus",
                "excel-com-generated",
                "protection.xls");

            using LegacyXlsLoadResult result = ExcelDocument.LoadLegacyXlsWithReport(workbookPath, new LegacyXlsImportOptions {
                ReportUnsupportedRecords = true
            });

            Assert.DoesNotContain(result.Diagnostics, diagnostic => diagnostic.Severity == LegacyXlsDiagnosticSeverity.Error);
            Assert.DoesNotContain(result.Diagnostics, diagnostic => diagnostic.Code == "XLS-BIFF-FILEPASS-UNSUPPORTED");
            Assert.Contains("PtgMul", result.ImportReport.FormulaTokensByName.Keys);

            Assert.NotNull(result.Workbook.Protection);
            Assert.True(result.Workbook.Protection!.IsProtected);
            Assert.Null(result.Workbook.Protection.LegacyPasswordHash);

            LegacyXlsWorksheet sheet = Assert.Single(result.Workbook.Worksheets);
            Assert.NotNull(sheet.Protection);
            Assert.True(sheet.Protection!.IsProtected);
            Assert.Matches("^[0-9A-F]{4}$", sheet.Protection.LegacyPasswordHash ?? string.Empty);
            AssertCorpusFormula(sheet, 2, 2, 14d, "A2*2");

            LegacyXlsCell inputCell = Assert.Single(sheet.Cells, cell => cell.Row == 2 && cell.Column == 1);
            LegacyXlsCell formulaCell = Assert.Single(sheet.Cells, cell => cell.Row == 2 && cell.Column == 2);
            LegacyXlsCellFormat inputFormat = Assert.Single(result.Workbook.CellFormats, format => format.StyleIndex == inputCell.StyleIndex);
            LegacyXlsCellFormat formulaFormat = Assert.Single(result.Workbook.CellFormats, format => format.StyleIndex == formulaCell.StyleIndex);
            Assert.True(inputFormat.ApplyProtection);
            Assert.False(inputFormat.Locked);
            Assert.False(inputFormat.FormulaHidden);
            Assert.True(formulaFormat.ApplyProtection);
            Assert.True(formulaFormat.Locked);
            Assert.True(formulaFormat.FormulaHidden);
        }

        [Fact]
        public void LegacyXls_DiagnosticCorpus_EncryptedWorkbook_ReportsFilePassBlocker() {
            string workbookPath = Path.Combine(
                GetTestsProjectRoot(),
                "Documents",
                "LegacyXlsDiagnosticCorpus",
                "excel-com-generated",
                "encrypted-password.xls");

            LegacyXlsWorkbook workbook = LegacyXlsWorkbook.Load(workbookPath, new LegacyXlsImportOptions {
                ReportUnsupportedRecords = true
            });
            LegacyXlsImportReport report = workbook.CreateImportReport();

            Assert.Empty(workbook.Worksheets);
            Assert.Contains(workbook.Diagnostics, diagnostic =>
                diagnostic.Severity == LegacyXlsDiagnosticSeverity.Error
                && diagnostic.Code == "XLS-BIFF-FILEPASS-UNSUPPORTED"
                && diagnostic.DetailCode == "Encryption:FilePass:Rc4");
            LegacyXlsUnsupportedFeature feature = Assert.Single(workbook.UnsupportedFeatures);
            Assert.Equal(LegacyXlsUnsupportedFeatureKind.EncryptedWorkbook, feature.Kind);
            Assert.Equal("XLS-BIFF-FILEPASS-UNSUPPORTED", feature.Code);
            Assert.True(report.HasImportErrors);
            Assert.Equal(1, report.UnsupportedFeaturesByKind[LegacyXlsUnsupportedFeatureKind.EncryptedWorkbook]);
            Assert.Equal(1, report.EncryptedWorkbooksByMethod["Rc4"]);
            Assert.Equal(1, report.FileFormatBlockers["EncryptedWorkbook|Encryption:FilePass:Rc4"]);
        }

        [Fact]
        public void LegacyXls_DiagnosticCorpus_Biff5Workbook_ReportsUnsupportedVersionBlocker() {
            string workbookPath = Path.Combine(
                GetTestsProjectRoot(),
                "Documents",
                "LegacyXlsDiagnosticCorpus",
                "excel-com-generated",
                "biff5-workbook.xls");

            LegacyXlsWorkbook workbook = LegacyXlsWorkbook.Load(workbookPath, new LegacyXlsImportOptions {
                ReportUnsupportedRecords = true
            });
            LegacyXlsImportReport report = workbook.CreateImportReport();

            Assert.Empty(workbook.Worksheets);
            Assert.Contains(workbook.Diagnostics, diagnostic =>
                diagnostic.Severity == LegacyXlsDiagnosticSeverity.Error
                && diagnostic.Code == "XLS-BIFF-VERSION-UNSUPPORTED"
                && diagnostic.DetailCode == "BiffVersion:BIFF5:WorkbookGlobals");
            LegacyXlsUnsupportedFeature feature = Assert.Single(workbook.UnsupportedFeatures);
            Assert.Equal(LegacyXlsUnsupportedFeatureKind.UnsupportedBiffVersion, feature.Kind);
            Assert.Equal("XLS-BIFF-VERSION-UNSUPPORTED", feature.Code);
            Assert.True(report.HasImportErrors);
            Assert.Equal(1, report.UnsupportedFeaturesByKind[LegacyXlsUnsupportedFeatureKind.UnsupportedBiffVersion]);
            Assert.Equal(1, report.UnsupportedBiffVersionsByVersion["BIFF5"]);
            Assert.Equal(1, report.UnsupportedBiffVersionsBySubstream["WorkbookGlobals"]);
            Assert.Equal(1, report.FileFormatBlockers["UnsupportedBiffVersion|BiffVersion:BIFF5:WorkbookGlobals"]);
        }

        private static bool IsLegacyXlsCorpusBaselineUpdateRequested() {
            string? value = Environment.GetEnvironmentVariable("OFFICEIMO_UPDATE_LEGACY_XLS_CORPUS_BASELINES");
            return string.Equals(value, "1", StringComparison.Ordinal)
                || string.Equals(value, "true", StringComparison.OrdinalIgnoreCase);
        }

        private static void AssertLegacyXlsCorpusBaselines(string corpusDirectory) {
            if (!Directory.Exists(corpusDirectory)) {
                return;
            }

            string[] workbookPaths = Directory.GetFiles(corpusDirectory, "*.xls", SearchOption.AllDirectories)
                .Where(path => !Path.GetFileName(path).StartsWith("~$", StringComparison.Ordinal))
                .OrderBy(path => path, StringComparer.OrdinalIgnoreCase)
                .ToArray();
            if (workbookPaths.Length == 0) {
                return;
            }

            bool updateBaselines = IsLegacyXlsCorpusBaselineUpdateRequested();
            var missingBaselines = new List<string>();
            foreach (string workbookPath in workbookPaths) {
                LegacyXlsWorkbook workbook = LegacyXlsWorkbook.Load(workbookPath, new LegacyXlsImportOptions {
                    ReportUnsupportedRecords = true
                });
                string actual = NormalizeBaselineText(workbook.CreateImportReport().ToMarkdown());
                string baselinePath = Path.ChangeExtension(workbookPath, ".import-report.md");

                if (updateBaselines) {
                    File.WriteAllText(baselinePath, actual, Encoding.UTF8);
                    continue;
                }

                if (!File.Exists(baselinePath)) {
                    missingBaselines.Add(GetRelativePath(corpusDirectory, baselinePath));
                    continue;
                }

                string expected = NormalizeBaselineText(File.ReadAllText(baselinePath, Encoding.UTF8));
                Assert.Equal(expected, actual);
            }

            Assert.True(
                missingBaselines.Count == 0,
                "Missing legacy XLS corpus baselines. Run with OFFICEIMO_UPDATE_LEGACY_XLS_CORPUS_BASELINES=1 to create: "
                    + string.Join(", ", missingBaselines));
        }

        private static string NormalizeBaselineText(string text) {
            return text.Replace("\r\n", "\n").Replace('\r', '\n').TrimEnd() + "\n";
        }

        private static string GetTestsProjectRoot() {
            var directory = new DirectoryInfo(AppContext.BaseDirectory);
            while (directory != null) {
                if (File.Exists(Path.Combine(directory.FullName, "OfficeIMO.Tests.csproj"))) {
                    return directory.FullName;
                }

                directory = directory.Parent;
            }

            return AppContext.BaseDirectory;
        }

        private static string GetRelativePath(string relativeTo, string path) {
            string root = EnsureTrailingDirectorySeparator(Path.GetFullPath(relativeTo));
            string target = Path.GetFullPath(path);
            Uri rootUri = new Uri(root);
            Uri targetUri = new Uri(target);
            string relative = Uri.UnescapeDataString(rootUri.MakeRelativeUri(targetUri).ToString());
            return relative.Replace('/', Path.DirectorySeparatorChar);
        }

        private static string EnsureTrailingDirectorySeparator(string path) {
            char separator = Path.DirectorySeparatorChar;
            char alternateSeparator = Path.AltDirectorySeparatorChar;
            if (path.Length == 0 || path[path.Length - 1] == separator || path[path.Length - 1] == alternateSeparator) {
                return path;
            }

            return path + separator;
        }

        private static void AssertCorpusFormula(LegacyXlsWorksheet sheet, int row, int column, object expectedValue, string expectedFormulaText) {
            LegacyXlsCell? cell = sheet.Cells.SingleOrDefault(candidate => candidate.Row == row && candidate.Column == column);
            Assert.True(
                cell != null,
                "Expected formula cell was not found. Parsed formula cells: "
                    + string.Join(", ", sheet.Cells
                        .Where(candidate => candidate.IsFormula)
                        .Select(candidate => $"R{candidate.Row}C{candidate.Column}={candidate.FormulaText}")));
            Assert.True(cell.IsFormula);
            Assert.Equal(expectedValue, cell.Value);
            Assert.Equal(expectedFormulaText, cell.FormulaText);
        }

        private static LegacyXlsDataValidation AssertValidation(
            LegacyXlsWorksheet sheet,
            LegacyXlsDataValidationType type,
            string range,
            string formula1,
            string? formula2) {
            LegacyXlsDataValidation validation = Assert.Single(sheet.DataValidations, candidate =>
                candidate.Type == type
                && candidate.Ranges.Contains(range)
                && candidate.Formula1 == formula1
                && candidate.Formula2 == formula2);
            return validation;
        }

        private static LegacyXlsConditionalFormatting AssertConditionalFormatting(
            LegacyXlsWorksheet sheet,
            LegacyXlsConditionalFormattingType type,
            LegacyXlsConditionalFormattingOperator? @operator,
            string range,
            string formula1,
            string? formula2) {
            LegacyXlsConditionalFormatting conditionalFormatting = Assert.Single(sheet.ConditionalFormattings, candidate =>
                candidate.Type == type
                && candidate.Operator == @operator
                && candidate.Ranges.Contains(range)
                && candidate.Formula1 == formula1
                && candidate.Formula2 == formula2);
            return conditionalFormatting;
        }
    }
}
