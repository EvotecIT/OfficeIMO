using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using OfficeIMO.Excel;
using OfficeIMO.Excel.LegacyXls;
using OfficeIMO.Excel.LegacyXls.Biff;
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
        public void LegacyXls_Corpus_ApachePoiTestData_BroadensLicensedCoverage() {
            string corpusDirectory = Path.Combine(
                GetTestsProjectRoot(),
                "Documents",
                "LegacyXlsCorpus",
                "apache-poi-testdata");
            string[] expectedFixtures = {
                "3dFormulas.xls",
                "IntersectionPtg.xls",
                "RangePtg.xls",
                "SimpleWithComments.xls",
                "templateExcelWithAutofilter.xls",
                "UnionPtg.xls",
                "WithExtendedStyles.xls",
                "WithTwoHyperLinks.xls"
            };

            string[] actualFixtures = Directory.GetFiles(corpusDirectory, "*.xls")
                .Select(path => Path.GetFileName(path)!)
                .OrderBy(fileName => fileName, StringComparer.OrdinalIgnoreCase)
                .ToArray();
            Assert.Equal(expectedFixtures.OrderBy(fileName => fileName, StringComparer.OrdinalIgnoreCase), actualFixtures);

            foreach (string fixture in expectedFixtures) {
                using LegacyXlsLoadResult result = LoadApachePoiFixture(fixture);
                Assert.DoesNotContain(result.Diagnostics, diagnostic => diagnostic.Severity == LegacyXlsDiagnosticSeverity.Error);
                Assert.DoesNotContain(result.Diagnostics, diagnostic => diagnostic.Severity == LegacyXlsDiagnosticSeverity.Warning);
                Assert.Equal(0, result.ImportReport.UnsupportedProjectionGapCount);
            }

            using LegacyXlsLoadResult formulas3d = LoadApachePoiFixture("3dFormulas.xls");
            Assert.Equal(5, formulas3d.ImportReport.FormulaTokensByName["PtgRef3d"]);
            Assert.Equal(4, formulas3d.ImportReport.FormulaTokensByName["PtgArea3d"]);
            Assert.Equal(1, formulas3d.ImportReport.ExternalReferenceCount);

            using LegacyXlsLoadResult intersection = LoadApachePoiFixture("IntersectionPtg.xls");
            Assert.Equal(1, intersection.ImportReport.FormulaTokensByName["PtgIsect"]);

            using LegacyXlsLoadResult range = LoadApachePoiFixture("RangePtg.xls");
            Assert.Equal(1, range.ImportReport.FormulaTokensByName["PtgRange"]);
            Assert.Equal(1, range.ImportReport.ExternalReferenceCount);

            using LegacyXlsLoadResult comments = LoadApachePoiFixture("SimpleWithComments.xls");
            Assert.Equal(3, comments.ImportReport.CommentCount);

            using LegacyXlsLoadResult hyperlinks = LoadApachePoiFixture("WithTwoHyperLinks.xls");
            Assert.Equal(2, hyperlinks.ImportReport.HyperlinkCount);

            using LegacyXlsLoadResult styles = LoadApachePoiFixture("WithExtendedStyles.xls");
            Assert.Equal(8, styles.ImportReport.CellStyleRecordCount);
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
            Assert.Equal(8, result.ImportReport.CompoundVbaModulesByContentKind["VbaCompressedContainer"]);
            Assert.Equal(1, result.ImportReport.CompoundVbaModulesByNameAndContentKind["ThisWorkbook|VbaCompressedContainer"]);
            Assert.Equal(1, result.ImportReport.CompoundVbaModulesByNameAndContentKind["Sheet8|VbaCompressedContainer"]);
            Assert.Equal(1, result.ImportReport.CompoundVbaModulesByCodeNameMatchAndName["WorkbookCodeName|ThisWorkbook"]);
            Assert.Equal(1, result.ImportReport.CompoundVbaModulesByCodeNameMatchAndName["WorksheetCodeName|Sheet1"]);
            Assert.Equal(1, result.ImportReport.CompoundVbaProjectsByStructure["Modules:8|DirStreams:1|ProjectStreams:2|Storages:2"]);
            Assert.Equal(1, result.ImportReport.VbaProjectWorkbookStates["BiffMarker:Present|NoMacrosMarker:Present|CompoundProject:Present|Modules:Present"]);
            Assert.Equal(0, result.ImportReport.UnsupportedProjectionGapCount);
            Assert.Empty(result.ImportReport.UnsupportedProjectionGapsByKind);
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
            Assert.Equal(1, result.ImportReport.DrawingPictureStates["PictureObjects:Present|BlipStore:Present|PictureBlipReferences:Present|ReferencedBlips:Resolved"]);
            Assert.Equal(1, result.ImportReport.DrawingShapePropertiesByGroup["Shape"]);
            Assert.False(result.ImportReport.DrawingShapePropertiesByGroup.ContainsKey("Unknown"));
            Assert.Equal(1, result.ImportReport.DrawingShapeComplexPropertiesByText["wzName:Chart 5"]);
            Assert.Equal(1, result.ImportReport.DrawingShapeComplexPropertiesByText["wzName:Picture 4"]);
            Assert.Equal(2, result.ImportReport.ChartGelFrameOfficeArtRecordsByType["OfficeArtFOPT"]);
            Assert.Equal(2, result.ImportReport.ChartGelFrameOfficeArtRecordsByType["EscherRecordType:0xF122"]);
            Assert.Equal(2, result.ImportReport.ChartGelFrameShapePropertyCounts["Properties:30"]);
            Assert.Equal(60, result.ImportReport.ChartGelFrameShapePropertiesByGroup["Fill"]);
            Assert.Equal(2, result.ImportReport.ChartGelFrameShapePropertiesByName["fillColor"]);
            Assert.Equal(2, result.ImportReport.ChartGelFrameShapePropertiesByName["fillBackColor"]);
            Assert.Equal(2, result.ImportReport.ChartGelFrameShapePropertiesByName["FillStyleBooleanProperties"]);
            Assert.Equal(2, result.ImportReport.ChartDataSourceFormulaTexts["'Objects'!$A$2:$A$6"]);
            Assert.Equal(2, result.ImportReport.ChartDataSourceFormulaTexts["'Objects'!$B$1"]);
            Assert.Equal(2, result.ImportReport.ChartDataSourceFormulaTexts["'Objects'!$B$2:$B$6"]);
            Assert.Equal(8, result.ImportReport.ChartLineFormatColors["#000000"]);
            Assert.Equal(8, result.ImportReport.ChartLineFormatColorIndexes["ColorIndex:23"]);
            Assert.Equal(4, result.ImportReport.ChartLineFormatStates["Style:Solid;Weight:Narrow;Automatic:False;AxisVisible:False;AutomaticColor:False"]);
            Assert.Equal(8, result.ImportReport.ChartAreaFormatColors["Foreground:#FFFFFF"]);
            Assert.Equal(4, result.ImportReport.ChartAreaFormatColorIndexes["ForegroundIndex:78"]);
            Assert.Equal(6, result.ImportReport.ChartAreaFormatStates["Pattern:Solid;Automatic:False;InvertNegative:False"]);
            Assert.Equal(2, result.ImportReport.ChartMarkerFormatColors["Foreground:#000000"]);
            Assert.Equal(2, result.ImportReport.ChartMarkerFormatStates["Type:Diamond;Automatic:True;InteriorHidden:False;BorderHidden:False"]);
            List<LegacyXlsChartRecord> gelFrameRecords = result.Workbook.ChartRecords
                .Where(record => record.RecordName == "GelFrame")
                .ToList();
            Assert.Equal(2, gelFrameRecords.Count);
            Assert.All(gelFrameRecords, record => {
                Assert.NotNull(record.GelFrame);
                Assert.Equal(2, record.GelFrame!.OfficeArtRecordCount);
                Assert.Equal(30, record.GelFrame.ShapePropertyCount);
                Assert.Contains(record.GelFrame.ShapeProperties, property => property.PropertyName == "fillColor");
                Assert.Contains(record.GelFrame.ShapeProperties, property => property.PropertyName == "fillBackColor");
            });
            Assert.Equal(1, result.ImportReport.PivotTableRecordsByKind[LegacyXlsPivotTableRecordKind.QueryTableTag]);
            Assert.Equal(1, result.ImportReport.PivotTableQueryTagTargets["PivotTable"]);
            Assert.Equal(1, result.ImportReport.PivotTableQueryTagNames["ObjectsPivot"]);
            LegacyXlsPivotTableRecord queryTag = Assert.Single(result.Workbook.PivotTableRecords, record => record.Kind == LegacyXlsPivotTableRecordKind.QueryTableTag);
            Assert.Equal("ObjectsPivot", queryTag.QueryTableTagName);
            Assert.True(queryTag.QueryTableTagRelatesToPivotTable);
            Assert.False(queryTag.QueryTableTagRefreshEnabled);
            Assert.True(queryTag.QueryTableTagCacheInvalid);
            Assert.False(queryTag.QueryTableTagTensorEx);
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
            Assert.Equal(1, result.ImportReport.WorksheetFeatureStates["DataValidations:0|ConditionalFormatting:0|AutoFilterCriteria:4|AutoFilterDropDowns:4"]);

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
            Assert.Equal(1, result.ImportReport.DataValidationCollectionRecordCount);
            Assert.Equal(1, result.ImportReport.DataValidationCollectionsBySheet["Validation"]);
            Assert.Equal(1, result.ImportReport.DataValidationCollectionsByDeclaredCount["Declared:9"]);
            Assert.Equal(1, result.ImportReport.DataValidationCollectionStates["Declared:9|Parsed:9|Matched"]);
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
            Assert.Equal(1, result.ImportReport.WorksheetFeatureStates["DataValidations:9|ConditionalFormatting:0|AutoFilterCriteria:0|AutoFilterDropDowns:Missing"]);

            LegacyXlsWorksheet validationSheet = Assert.Single(result.Workbook.Worksheets, sheet => sheet.Name == "Validation");
            Assert.Single(result.Workbook.Worksheets, sheet => sheet.Name == "Lookup");
            LegacyXlsDataValidationCollectionRecord collectionRecord = Assert.Single(validationSheet.DataValidationCollections);
            Assert.Equal(9U, collectionRecord.DeclaredValidationCount);
            Assert.Equal((ushort)BiffRecordType.DVal, collectionRecord.RecordType);

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
            Assert.Equal(8, result.ImportReport.ConditionalFormattingExtensionRecordCount);
            Assert.Equal(8, result.ImportReport.ConditionalFormattingExtensionsBySheet["Conditions"]);
            Assert.Equal(8, result.ImportReport.ConditionalFormattingExtensionsByRecordType["0x087B"]);
            Assert.Equal(8, result.ImportReport.ConditionalFormattingExtensionStopIfTrueStates["StopIfTrue"]);
            Assert.Equal(4, result.ImportReport.ConditionalFormattingExtensionStates["Cf12:Missing|UnprojectedFormatting:Present|MatchedRule:Present|Priority:Present|StopIfTrue:StopIfTrue"]);
            Assert.Equal(8, result.ImportReport.ConditionalFormattingExtensionStates.Values.Sum());
            Assert.Equal(1, result.ImportReport.UnsupportedFeaturesByDetail["ConditionalFormatting|XLS-BIFF-FEATURE-CONDITIONAL-FORMATTING-UNSUPPORTED|ConditionalFormatting:Dxf"]);
            Assert.Equal(1, result.ImportReport.WorksheetFeatureStates["DataValidations:0|ConditionalFormatting:8|AutoFilterCriteria:0|AutoFilterDropDowns:Missing"]);

            LegacyXlsWorksheet sheet = Assert.Single(result.Workbook.Worksheets);
            Assert.Equal(8, sheet.ConditionalFormattingExtensions.Count);
            Assert.All(sheet.ConditionalFormattingExtensions, extension => {
                Assert.Equal("Conditions", extension.SheetName);
                Assert.Equal(0x087B, extension.RecordType);
                Assert.False(extension.IsCf12);
                Assert.True(extension.Priority.HasValue);
            });
            Assert.Equal(4, sheet.ConditionalFormattingExtensions.Count(extension => extension.MatchedRule));
            Assert.Equal(4, sheet.ConditionalFormattingExtensions.Count(extension => !extension.MatchedRule));
            Assert.Equal(7, sheet.ConditionalFormattingExtensions.Count(extension => extension.HasUnprojectedFormatting));
            Assert.Single(sheet.ConditionalFormattingExtensions, extension => !extension.HasUnprojectedFormatting);
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
            LegacyXlsChartRecord barRecord = Assert.Single(result.Workbook.ChartRecords, record => record.RecordName == "Bar" && record.SheetName == "RevenueChart");
            Assert.NotNull(barRecord.BarOptions);
            Assert.Equal(0, barRecord.BarOptions!.OverlapPercentage);
            Assert.Equal(150, barRecord.BarOptions.GapWidthPercentage);
            Assert.False(barRecord.BarOptions.IsTransposed);
            Assert.False(barRecord.BarOptions.IsStacked);
            Assert.False(barRecord.BarOptions.IsPercentStacked);
            Assert.False(barRecord.BarOptions.HasShadow);
            Assert.Equal(1, result.ImportReport.ChartSheetPropertyStates["AutoSeries:False;VisibleOnly:False;DoNotSizeWithWindow:False;ManualPlotArea:False;AlwaysAutoPlotArea:False"]);
            Assert.Equal(1, result.ImportReport.ChartBarOverlapPercentages["Overlap:0"]);
            Assert.Equal(1, result.ImportReport.ChartBarGapWidths["Gap:150"]);
            Assert.Equal(1, result.ImportReport.ChartBarStates["Transposed:False;Stacked:False;Percent:False;Shadow:False"]);
            Assert.Equal(1, result.ImportReport.ChartGelFrameOfficeArtRecordsByType["OfficeArtFOPT"]);
            Assert.Equal(1, result.ImportReport.ChartGelFrameOfficeArtRecordsByType["EscherRecordType:0xF122"]);
            Assert.Equal(1, result.ImportReport.ChartGelFrameShapePropertyCounts["Properties:30"]);
            Assert.Equal(30, result.ImportReport.ChartGelFrameShapePropertiesByGroup["Fill"]);
            Assert.Equal(1, result.ImportReport.ChartGelFrameShapePropertiesByName["fillColor"]);
            Assert.Equal(1, result.ImportReport.ChartGelFrameShapePropertiesByName["fillBackColor"]);
            Assert.Equal(1, result.ImportReport.ChartGelFrameShapePropertiesByName["FillStyleBooleanProperties"]);
            LegacyXlsChartRecord gelFrameRecord = Assert.Single(result.Workbook.ChartRecords, record => record.RecordName == "GelFrame");
            Assert.NotNull(gelFrameRecord.GelFrame);
            Assert.Equal(2, gelFrameRecord.GelFrame!.OfficeArtRecordCount);
            Assert.Equal(30, gelFrameRecord.GelFrame.ShapePropertyCount);
            Assert.Contains(gelFrameRecord.GelFrame.ShapeProperties, property => property.PropertyName == "fillColor");
            Assert.Contains(gelFrameRecord.GelFrame.ShapeProperties, property => property.PropertyName == "fillBackColor");
            Assert.Equal(1, result.ImportReport.PivotTableRecordsByKind[LegacyXlsPivotTableRecordKind.QueryTableTag]);
            Assert.Equal(1, result.ImportReport.PivotTableQueryTagTargets["PivotTable"]);
            Assert.Equal(1, result.ImportReport.PivotTableQueryTagNames["SalesPivot"]);
            LegacyXlsPivotTableRecord queryTag = Assert.Single(result.Workbook.PivotTableRecords, record => record.Kind == LegacyXlsPivotTableRecordKind.QueryTableTag);
            Assert.Equal("SalesPivot", queryTag.QueryTableTagName);
            Assert.True(queryTag.QueryTableTagRelatesToPivotTable);
            Assert.False(queryTag.QueryTableTagRefreshEnabled);
            Assert.True(queryTag.QueryTableTagCacheInvalid);
            Assert.False(queryTag.QueryTableTagTensorEx);
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
            Assert.Equal(1, result.ImportReport.ArrayFormulaRecordCount);
            Assert.Equal(1, result.ImportReport.ArrayFormulasBySheetAndRange["Formulas!D1"]);
            Assert.Equal(1, result.ImportReport.ArrayFormulasByProjectionState["FormulaTextProjected"]);
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
        public void LegacyXls_Corpus_FormulaFunctions_ProjectsCommonBuiltInFunctions() {
            string workbookPath = Path.Combine(
                GetTestsProjectRoot(),
                "Documents",
                "LegacyXlsCorpus",
                "excel-com-generated",
                "formula-functions.xls");

            using LegacyXlsLoadResult result = ExcelDocument.LoadLegacyXlsWithReport(workbookPath, new LegacyXlsImportOptions {
                ReportUnsupportedRecords = true
            });

            Assert.DoesNotContain(result.Diagnostics, diagnostic => diagnostic.Severity == LegacyXlsDiagnosticSeverity.Error);
            Assert.DoesNotContain(result.Diagnostics, diagnostic => diagnostic.Code == "XLS-BIFF-FORMULA-TOKENS-UNSUPPORTED");
            Assert.Empty(result.ImportReport.FormulaTokenBlockers);

            foreach (string functionName in new[] {
                "AVERAGE",
                "CONCATENATE",
                "COUNTA",
                "COUNTBLANK",
                "FIND",
                "MAX",
                "MEDIAN",
                "MIN",
                "PRODUCT",
                "REPLACE",
                "SEARCH",
                "SUBSTITUTE",
                "TRIM",
                "VAR"
            }) {
                Assert.Equal(1, result.ImportReport.FormulaFunctionsByName[functionName]);
            }

            LegacyXlsWorksheet sheet = Assert.Single(result.Workbook.Worksheets);
            AssertCorpusFormula(sheet, 1, 2, 20d, "AVERAGE(A1:A3)");
            AssertCorpusFormula(sheet, 2, 2, 10d, "MIN(A1:A3)");
            AssertCorpusFormula(sheet, 3, 2, 30d, "MAX(A1:A3)");
            AssertCorpusFormula(sheet, 4, 2, 6000d, "PRODUCT(A1:A3)");
            AssertCorpusFormula(sheet, 5, 2, 5d, "COUNTA(A1:A5)");
            AssertCorpusFormula(sheet, 6, 2, 20d, "MEDIAN(A1:A3)");
            AssertCorpusFormula(sheet, 7, 2, 100d, "VAR(A1:A3)");
            AssertCorpusFormula(sheet, 8, 2, 0d, "COUNTBLANK(B1:B3)");
            AssertCorpusFormula(sheet, 9, 2, "north-east", "CONCATENATE(A4,\"-\",A5)");
            AssertCorpusFormula(sheet, 10, 2, 2d, "FIND(\"o\",A4)");
            AssertCorpusFormula(sheet, 11, 2, 2d, "SEARCH(\"A\",A5)");
            AssertCorpusFormula(sheet, 12, 2, "padded", "TRIM(\"  padded  \")");
            AssertCorpusFormula(sheet, 13, 2, "north/east", "SUBSTITUTE(\"north-east\",\"-\",\"/\")");
            AssertCorpusFormula(sheet, 14, 2, "nXXth", "REPLACE(\"north\",2,2,\"XX\")");
        }

        [Fact]
        public void LegacyXls_Corpus_FormulaAdvanced_ProjectsAdditionalBuiltInFunctions() {
            string workbookPath = Path.Combine(
                GetTestsProjectRoot(),
                "Documents",
                "LegacyXlsCorpus",
                "excel-com-generated",
                "formula-advanced.xls");

            using LegacyXlsLoadResult result = ExcelDocument.LoadLegacyXlsWithReport(workbookPath, new LegacyXlsImportOptions {
                ReportUnsupportedRecords = true
            });

            Assert.DoesNotContain(result.Diagnostics, diagnostic => diagnostic.Severity == LegacyXlsDiagnosticSeverity.Error);
            Assert.DoesNotContain(result.Diagnostics, diagnostic => diagnostic.Code == "XLS-BIFF-FORMULA-TOKENS-UNSUPPORTED");
            Assert.Equal(2, result.ImportReport.FutureFunctionAliasCount);
            Assert.Equal(1, result.ImportReport.FutureFunctionAliasesByName["_xlfn.AVERAGEIF"]);
            Assert.Equal(1, result.ImportReport.FutureFunctionAliasesByName["_xlfn.IFERROR"]);
            Assert.Equal(1, result.ImportReport.FutureFunctionAliasesByFunction["AVERAGEIF"]);
            Assert.Equal(1, result.ImportReport.FutureFunctionAliasesByFunction["IFERROR"]);
            Assert.Equal(2, result.ImportReport.FutureFunctionAliasesByTokenName["PtgErr"]);

            foreach (string functionName in new[] {
                "DATEVALUE",
                "HLOOKUP",
                "INDEX",
                "ISERROR",
                "LEFT",
                "MATCH",
                "OFFSET",
                "POWER",
                "RIGHT",
                "ROUNDDOWN",
                "ROUNDUP",
                "SUMPRODUCT",
                "VLOOKUP"
            }) {
                Assert.Equal(1, result.ImportReport.FormulaFunctionsByName[functionName]);
            }

            LegacyXlsWorksheet sheet = Assert.Single(result.Workbook.Worksheets);
            AssertCorpusFormula(sheet, 1, 4, 3d, "ROUNDUP(A1/3,2)");
            AssertCorpusFormula(sheet, 2, 4, 3d, "ROUNDDOWN(A1/3,2)");
            AssertCorpusFormula(sheet, 3, 4, 81d, "POWER(A1,2)");
            AssertCorpusFormula(sheet, 4, 4, 61d, "SUMPRODUCT(A1:A3,B1:B3)");
            AssertCorpusFormula(sheet, 5, 4, true, "ISERROR(A1/0)");
            AssertCorpusFormula(sheet, 6, 4, 46197d, "DATEVALUE(\"2026-06-24\")");
            AssertCorpusFormula(sheet, 7, 4, 4d, "_xlfn.AVERAGEIF(A1:A3,\">2\",B1:B3)");
            AssertCorpusFormula(sheet, 8, 4, 0d, "_xlfn.IFERROR(A1/0,0)");
            AssertCorpusFormula(sheet, 9, 4, "nost", "LEFT(\"north\",2)&RIGHT(\"east\",2)");
            AssertCorpusFormula(sheet, 10, 4, "#N/A", "HLOOKUP(4,A1:B3,2,FALSE)");
            AssertCorpusFormula(sheet, 11, 4, 5d, "VLOOKUP(4,A1:B3,2,FALSE)");
            AssertCorpusFormula(sheet, 12, 4, 2d, "MATCH(4,A1:A3,0)");
            AssertCorpusFormula(sheet, 13, 4, 5d, "INDEX(B1:B3,2)");
            AssertCorpusFormula(sheet, 14, 4, 5d, "OFFSET(A1,1,1)");
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

        private static LegacyXlsLoadResult LoadApachePoiFixture(string fileName) {
            string workbookPath = Path.Combine(
                GetTestsProjectRoot(),
                "Documents",
                "LegacyXlsCorpus",
                "apache-poi-testdata",
                fileName);
            return ExcelDocument.LoadLegacyXlsWithReport(workbookPath, new LegacyXlsImportOptions {
                ReportUnsupportedRecords = true
            });
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
