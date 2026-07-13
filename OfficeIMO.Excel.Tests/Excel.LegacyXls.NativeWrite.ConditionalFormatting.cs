using OpenXmlConditionalFormatValues = DocumentFormat.OpenXml.Spreadsheet.ConditionalFormatValues;
using OpenXmlConditionalFormatting = DocumentFormat.OpenXml.Spreadsheet.ConditionalFormatting;
using OpenXmlConditionalFormattingOperatorValues = DocumentFormat.OpenXml.Spreadsheet.ConditionalFormattingOperatorValues;
using OpenXmlConditionalFormattingRule = DocumentFormat.OpenXml.Spreadsheet.ConditionalFormattingRule;
using OpenXmlFormula = DocumentFormat.OpenXml.Spreadsheet.Formula;
using OpenXmlTimePeriodValues = DocumentFormat.OpenXml.Spreadsheet.TimePeriodValues;
using OfficeIMO.Excel;
using OfficeIMO.Excel.LegacyXls;
using OfficeIMO.Excel.LegacyXls.Model;
using Xunit;

namespace OfficeIMO.Tests {
    public partial class Excel {
        [Fact]
        public void LegacyXls_NativeSave_WritesAdditionalWorksheetConditionalFormattingOperators() {
            string openXmlPath = Path.Combine(Path.GetTempPath(), Guid.NewGuid().ToString("N") + ".xlsx");
            string xlsOutputPath = Path.Combine(Path.GetTempPath(), Guid.NewGuid().ToString("N") + ".xls");

            try {
                using (ExcelDocument document = ExcelDocument.Create(openXmlPath)) {
                    ExcelSheet sheet = document.AddWorksheet("MoreConditions");
                    for (int row = 1; row <= 6; row++) {
                        for (int column = 1; column <= 8; column++) {
                            sheet.CellValue(row, column, row * column);
                        }
                    }

                    sheet.AddConditionalRule("A2:A5", OpenXmlConditionalFormattingOperatorValues.NotBetween, "2", "5");
                    sheet.AddConditionalRule("B2:B5", OpenXmlConditionalFormattingOperatorValues.Equal, "8");
                    sheet.AddConditionalRule("C2:C5", OpenXmlConditionalFormattingOperatorValues.NotEqual, "9");
                    sheet.AddConditionalRule("D2:D5", OpenXmlConditionalFormattingOperatorValues.LessThan, "20");
                    sheet.AddConditionalRule("E2:E5", OpenXmlConditionalFormattingOperatorValues.GreaterThanOrEqual, "15");
                    sheet.AddConditionalRule("F2:F3 H2:H3", OpenXmlConditionalFormattingOperatorValues.LessThanOrEqual, "30");

                    document.Save(xlsOutputPath);
                }

                using LegacyXlsLoadResult result = ExcelDocument.LoadLegacyXlsWithReport(xlsOutputPath);
                result.EnsureNoImportErrors();
                Assert.False(result.HasUnsupportedFeatures, FormatUnsupportedFeatures(result.UnsupportedFeatures));

                LegacyXlsWorksheet legacySheet = Assert.Single(result.Workbook.Worksheets);
                Assert.Equal(6, legacySheet.ConditionalFormattings.Count);

                AssertCellIsRule(
                    legacySheet,
                    LegacyXlsConditionalFormattingOperator.NotBetween,
                    new[] { "A2:A5" },
                    "2",
                    "5");
                AssertCellIsRule(
                    legacySheet,
                    LegacyXlsConditionalFormattingOperator.Equal,
                    new[] { "B2:B5" },
                    "8",
                    null);
                AssertCellIsRule(
                    legacySheet,
                    LegacyXlsConditionalFormattingOperator.NotEqual,
                    new[] { "C2:C5" },
                    "9",
                    null);
                AssertCellIsRule(
                    legacySheet,
                    LegacyXlsConditionalFormattingOperator.LessThan,
                    new[] { "D2:D5" },
                    "20",
                    null);
                AssertCellIsRule(
                    legacySheet,
                    LegacyXlsConditionalFormattingOperator.GreaterThanOrEqual,
                    new[] { "E2:E5" },
                    "15",
                    null);
                AssertCellIsRule(
                    legacySheet,
                    LegacyXlsConditionalFormattingOperator.LessThanOrEqual,
                    new[] { "F2:F3", "H2:H3" },
                    "30",
                    null);

                ExcelSheet projectedSheet = result.Document.Sheets[0];
                AssertProjectedConditionalRule(projectedSheet, "A2:A5", nameof(OpenXmlConditionalFormattingOperatorValues.NotBetween), new[] { "2", "5" });
                AssertProjectedConditionalRule(projectedSheet, "B2:B5", nameof(OpenXmlConditionalFormattingOperatorValues.Equal), new[] { "8" });
                AssertProjectedConditionalRule(projectedSheet, "C2:C5", nameof(OpenXmlConditionalFormattingOperatorValues.NotEqual), new[] { "9" });
                AssertProjectedConditionalRule(projectedSheet, "D2:D5", nameof(OpenXmlConditionalFormattingOperatorValues.LessThan), new[] { "20" });
                AssertProjectedConditionalRule(projectedSheet, "E2:E5", nameof(OpenXmlConditionalFormattingOperatorValues.GreaterThanOrEqual), new[] { "15" });
                ExcelConditionalFormattingInfo projectedMultiRange = Assert.Single(projectedSheet.GetConditionalFormattingRules("F2:F3"));
                Assert.Equal("F2:F3 H2:H3", projectedMultiRange.Range);
                Assert.Equal(nameof(OpenXmlConditionalFormattingOperatorValues.LessThanOrEqual), projectedMultiRange.Operator);
                Assert.Equal(new[] { "30" }, projectedMultiRange.Formulas);
            } finally {
                TryDelete(openXmlPath);
                TryDelete(xlsOutputPath);
            }
        }

        [Fact]
        public void LegacyXls_NativeSave_WritesConditionalFormattingStopIfTrueExtensions() {
            string openXmlPath = Path.Combine(Path.GetTempPath(), Guid.NewGuid().ToString("N") + ".xlsx");
            string xlsOutputPath = Path.Combine(Path.GetTempPath(), Guid.NewGuid().ToString("N") + ".xls");

            try {
                using (ExcelDocument document = ExcelDocument.Create(openXmlPath)) {
                    ExcelSheet sheet = document.AddWorksheet("CfStop");
                    sheet.CellValue(1, 1, 1d);
                    sheet.CellValue(2, 1, 5d);
                    sheet.CellValue(3, 1, 10d);

                    sheet.AddConditionalRule("A1:A3", OpenXmlConditionalFormattingOperatorValues.GreaterThan, "3", null, stopIfTrue: true);

                    document.Save(xlsOutputPath);
                }

                using LegacyXlsLoadResult result = ExcelDocument.LoadLegacyXlsWithReport(xlsOutputPath);
                result.EnsureNoImportErrors();
                Assert.False(result.HasUnsupportedFeatures, FormatUnsupportedFeatures(result.UnsupportedFeatures));

                LegacyXlsWorksheet legacySheet = Assert.Single(result.Workbook.Worksheets);
                LegacyXlsConditionalFormatting rule = Assert.Single(legacySheet.ConditionalFormattings);
                Assert.Equal(LegacyXlsConditionalFormattingType.CellIs, rule.Type);
                Assert.Equal(LegacyXlsConditionalFormattingOperator.GreaterThan, rule.Operator);
                Assert.Equal("3", rule.Formula1);
                Assert.True(rule.StopIfTrue);
                Assert.Equal(1, rule.Priority);

                LegacyXlsConditionalFormattingExtensionRecord extension = Assert.Single(legacySheet.ConditionalFormattingExtensions);
                Assert.True(extension.MatchedRule);
                Assert.Equal((ushort?)1, extension.HeaderId);
                Assert.Equal((ushort?)0, extension.RuleIndex);
                Assert.Equal(rule.Priority, extension.Priority);
                Assert.True(extension.StopIfTrue);
                Assert.False(extension.HasUnprojectedFormatting);

                ExcelConditionalFormattingInfo projectedRule = Assert.Single(result.Document.Sheets[0].GetConditionalFormattingRules("A1:A3"));
                Assert.Equal(nameof(OpenXmlConditionalFormattingOperatorValues.GreaterThan), projectedRule.Operator);
                Assert.True(projectedRule.StopIfTrue);
                Assert.Equal(rule.Priority, projectedRule.Priority);
            } finally {
                TryDelete(openXmlPath);
                TryDelete(xlsOutputPath);
            }
        }

        [Fact]
        public void LegacyXls_NativeSave_TreatsSingleCellConditionalFormattingReferencesAsOneCellRanges() {
            string openXmlPath = Path.Combine(Path.GetTempPath(), Guid.NewGuid().ToString("N") + ".xlsx");
            string xlsOutputPath = Path.Combine(Path.GetTempPath(), Guid.NewGuid().ToString("N") + ".xls");

            try {
                using (ExcelDocument document = ExcelDocument.Create(openXmlPath)) {
                    ExcelSheet sheet = document.AddWorksheet("SingleCf");
                    sheet.CellValue(1, 1, 3d);
                    sheet.AddConditionalRule("A1", OpenXmlConditionalFormattingOperatorValues.GreaterThan, "1");

                    document.Save(xlsOutputPath);
                }

                using LegacyXlsLoadResult result = ExcelDocument.LoadLegacyXlsWithReport(xlsOutputPath);
                result.EnsureNoImportErrors();
                Assert.False(result.HasUnsupportedFeatures, FormatUnsupportedFeatures(result.UnsupportedFeatures));

                LegacyXlsWorksheet legacySheet = Assert.Single(result.Workbook.Worksheets);
                LegacyXlsConditionalFormatting rule = Assert.Single(legacySheet.ConditionalFormattings);
                Assert.Equal("A1", Assert.Single(rule.Ranges));
            } finally {
                TryDelete(openXmlPath);
                TryDelete(xlsOutputPath);
            }
        }

        [Fact]
        public void LegacyXls_NativeSave_WritesFormulaBackedConditionalFormattingRuleTypes() {
            string openXmlPath = Path.Combine(Path.GetTempPath(), Guid.NewGuid().ToString("N") + ".xlsx");
            string xlsOutputPath = Path.Combine(Path.GetTempPath(), Guid.NewGuid().ToString("N") + ".xls");

            try {
                using (ExcelDocument document = ExcelDocument.Create(openXmlPath)) {
                    ExcelSheet sheet = document.AddWorksheet("CfExpressions");
                    sheet.CellValue(1, 1, "Ready");
                    sheet.CellValue(2, 1, "Blocked");
                    sheet.CellValue(3, 1, "Ready");
                    sheet.CellValue(1, 2, string.Empty);
                    sheet.CellValue(2, 2, "Filled");
                    sheet.CellValue(1, 3, "#N/A");
                    sheet.CellValue(2, 3, "Ok");

                    sheet.AddConditionalTextRule("A1:A3", OpenXmlConditionalFormatValues.ContainsText, "Ready");
                    sheet.AddConditionalBlanksRule("B1:B3");
                    sheet.AddConditionalErrorsRule("C1:C3", containsErrors: false);

                    document.Save(xlsOutputPath);
                }

                using LegacyXlsLoadResult result = ExcelDocument.LoadLegacyXlsWithReport(xlsOutputPath);
                result.EnsureNoImportErrors();
                Assert.False(result.HasUnsupportedFeatures, FormatUnsupportedFeatures(result.UnsupportedFeatures));

                LegacyXlsWorksheet legacySheet = Assert.Single(result.Workbook.Worksheets);
                Assert.Equal(3, legacySheet.ConditionalFormattings.Count);
                Assert.All(legacySheet.ConditionalFormattings, rule => Assert.Equal(LegacyXlsConditionalFormattingType.Formula, rule.Type));
                Assert.Contains(legacySheet.ConditionalFormattings, rule => rule.Ranges.Single() == "A1:A3" && rule.Formula1 == "NOT(ISERROR(SEARCH(\"Ready\",A1)))");
                Assert.Contains(legacySheet.ConditionalFormattings, rule => rule.Ranges.Single() == "B1:B3" && rule.Formula1 == "LEN(TRIM(B1))=0");
                Assert.Contains(legacySheet.ConditionalFormattings, rule => rule.Ranges.Single() == "C1:C3" && rule.Formula1 == "NOT(ISERROR(C1))");

                ExcelSheet projectedSheet = result.Document.Sheets[0];
                AssertProjectedExpressionRule(projectedSheet, "A1:A3", "NOT(ISERROR(SEARCH(\"Ready\",A1)))");
                AssertProjectedExpressionRule(projectedSheet, "B1:B3", "LEN(TRIM(B1))=0");
                AssertProjectedExpressionRule(projectedSheet, "C1:C3", "NOT(ISERROR(C1))");
            } finally {
                TryDelete(openXmlPath);
                TryDelete(xlsOutputPath);
            }
        }

        [Fact]
        public void LegacyXls_NativeSave_WritesExternalWorkbookFormulaConditionalFormatting() {
            string openXmlPath = Path.Combine(Path.GetTempPath(), Guid.NewGuid().ToString("N") + ".xlsx");
            string xlsOutputPath = Path.Combine(Path.GetTempPath(), Guid.NewGuid().ToString("N") + ".xls");

            try {
                using (ExcelDocument document = ExcelDocument.Create(openXmlPath)) {
                    ExcelSheet sheet = document.AddWorksheet("ExternalCf");
                    sheet.CellValue(1, 1, 5d);

                    sheet.WorksheetPart.Worksheet!.Append(new OpenXmlConditionalFormatting(
                        new OpenXmlConditionalFormattingRule(
                            new OpenXmlFormula("COUNTIF('[Other.xlsx]Data'!$A$1:$A$3,\">0\")>0")) {
                            Type = OpenXmlConditionalFormatValues.Expression,
                            Priority = 1
                        }) {
                        SequenceOfReferences = new DocumentFormat.OpenXml.ListValue<DocumentFormat.OpenXml.StringValue> { InnerText = "A1:A1" }
                    });
                    sheet.WorksheetPart.Worksheet.Save();

                    document.Save(xlsOutputPath);
                }

                using LegacyXlsLoadResult result = ExcelDocument.LoadLegacyXlsWithReport(xlsOutputPath);
                result.EnsureNoImportErrors();
                LegacyXlsExternalReference externalReference = Assert.Single(
                    result.Workbook.ExternalReferences,
                    reference => reference.Kind == LegacyXlsExternalReferenceKind.ExternalWorkbook);
                Assert.Equal("Other.xlsx", externalReference.Target);
                Assert.Equal(new[] { "Data" }, externalReference.SheetNames);

                LegacyXlsWorksheet legacySheet = Assert.Single(result.Workbook.Worksheets);
                LegacyXlsConditionalFormatting rule = Assert.Single(legacySheet.ConditionalFormattings);
                Assert.Equal(LegacyXlsConditionalFormattingType.Formula, rule.Type);
                Assert.Equal("COUNTIF('[Other.xlsx]Data'!$A$1:$A$3,\">0\")>0", rule.Formula1);
                Assert.Equal("A1", Assert.Single(rule.Ranges));
            } finally {
                TryDelete(openXmlPath);
                TryDelete(xlsOutputPath);
            }
        }

        [Fact]
        public void LegacyXls_NativeSave_WritesExternalDefinedNameConditionalFormatting() {
            string openXmlPath = Path.Combine(Path.GetTempPath(), Guid.NewGuid().ToString("N") + ".xlsx");
            string xlsOutputPath = Path.Combine(Path.GetTempPath(), Guid.NewGuid().ToString("N") + ".xls");

            try {
                using (ExcelDocument document = ExcelDocument.Create(openXmlPath)) {
                    ExcelSheet sheet = document.AddWorksheet("ExternalCfName");
                    sheet.CellValue(1, 1, 5d);

                    sheet.WorksheetPart.Worksheet!.Append(new OpenXmlConditionalFormatting(
                        new OpenXmlConditionalFormattingRule(
                            new OpenXmlFormula("COUNTIF([Other.xlsx]HasPositiveValues,\">0\")>0")) {
                            Type = OpenXmlConditionalFormatValues.Expression,
                            Priority = 1
                        }) {
                        SequenceOfReferences = new DocumentFormat.OpenXml.ListValue<DocumentFormat.OpenXml.StringValue> { InnerText = "A1:A1" }
                    });
                    sheet.WorksheetPart.Worksheet.Save();

                    document.Save(xlsOutputPath);
                }

                using LegacyXlsLoadResult result = ExcelDocument.LoadLegacyXlsWithReport(xlsOutputPath);
                result.EnsureNoImportErrors();
                LegacyXlsExternalReference externalReference = Assert.Single(
                    result.Workbook.ExternalReferences,
                    reference => reference.Kind == LegacyXlsExternalReferenceKind.ExternalWorkbook);
                Assert.Equal("Other.xlsx", externalReference.Target);
                LegacyXlsExternalName externalName = Assert.Single(externalReference.ExternalNames);
                Assert.Equal("HasPositiveValues", externalName.Name);

                LegacyXlsWorksheet legacySheet = Assert.Single(result.Workbook.Worksheets);
                LegacyXlsConditionalFormatting rule = Assert.Single(legacySheet.ConditionalFormattings);
                Assert.Equal(LegacyXlsConditionalFormattingType.Formula, rule.Type);
                Assert.Equal("COUNTIF('Other.xlsx'!HasPositiveValues,\">0\")>0", rule.Formula1);
                Assert.Equal("A1", Assert.Single(rule.Ranges));
            } finally {
                TryDelete(openXmlPath);
                TryDelete(xlsOutputPath);
            }
        }

        [Fact]
        public void LegacyXls_NativeSave_WritesSheetScopedExternalDefinedNameConditionalFormatting() {
            string openXmlPath = Path.Combine(Path.GetTempPath(), Guid.NewGuid().ToString("N") + ".xlsx");
            string xlsOutputPath = Path.Combine(Path.GetTempPath(), Guid.NewGuid().ToString("N") + ".xls");

            try {
                using (ExcelDocument document = ExcelDocument.Create(openXmlPath)) {
                    ExcelSheet sheet = document.AddWorksheet("ExternalCfName");
                    sheet.CellValue(1, 1, 5d);

                    sheet.WorksheetPart.Worksheet!.Append(new OpenXmlConditionalFormatting(
                        new OpenXmlConditionalFormattingRule(
                            new OpenXmlFormula("COUNTIF('[Other.xlsx]Feb'!HasPositiveValues,\">0\")>0")) {
                            Type = OpenXmlConditionalFormatValues.Expression,
                            Priority = 1
                        }) {
                        SequenceOfReferences = new DocumentFormat.OpenXml.ListValue<DocumentFormat.OpenXml.StringValue> { InnerText = "A1:A1" }
                    });
                    sheet.WorksheetPart.Worksheet.Save();

                    document.Save(xlsOutputPath);
                }

                using LegacyXlsLoadResult result = ExcelDocument.LoadLegacyXlsWithReport(xlsOutputPath);
                result.EnsureNoImportErrors();
                LegacyXlsExternalReference externalReference = Assert.Single(
                    result.Workbook.ExternalReferences,
                    reference => reference.Kind == LegacyXlsExternalReferenceKind.ExternalWorkbook);
                Assert.Equal("Other.xlsx", externalReference.Target);
                Assert.Equal(new[] { "Feb" }, externalReference.SheetNames);
                LegacyXlsExternalName externalName = Assert.Single(externalReference.ExternalNames);
                Assert.Equal("HasPositiveValues", externalName.Name);
                Assert.Equal(0, externalName.LocalSheetIndex);

                LegacyXlsWorksheet legacySheet = Assert.Single(result.Workbook.Worksheets);
                LegacyXlsConditionalFormatting rule = Assert.Single(legacySheet.ConditionalFormattings);
                Assert.Equal(LegacyXlsConditionalFormattingType.Formula, rule.Type);
                Assert.Equal("COUNTIF('[Other.xlsx]Feb'!HasPositiveValues,\">0\")>0", rule.Formula1);
                Assert.Equal("A1", Assert.Single(rule.Ranges));
            } finally {
                TryDelete(openXmlPath);
                TryDelete(xlsOutputPath);
            }
        }

        [Fact]
        public void LegacyXls_NativeSave_BlocksOversizedConditionalFormattingFormulaPayloadsBeforeWriting() {
            AssertNativeXlsSaveNotSupported("conditional formatting formula token payload lengths outside BIFF8 limits", (document, sheet) => {
                sheet.CellValue(1, 1, "Conditioned");
                string longLiteral = "\"" + new string('A', 255) + "\"";
                string formula = string.Join("&", Enumerable.Repeat(longLiteral, 260));

                sheet.AddConditionalFormulaRule("A1:A1", "A1<>\"\"");
                OpenXmlConditionalFormattingRule rule = sheet.WorksheetPart.Worksheet!
                    .Elements<OpenXmlConditionalFormatting>()
                    .Last()
                    .Elements<OpenXmlConditionalFormattingRule>()
                    .Single();
                rule.Priority = null;
                rule.GetFirstChild<OpenXmlFormula>()!.Text = formula;
                sheet.WorksheetPart.Worksheet.Save();
            });
        }

        [Fact]
        public void LegacyXls_NativeSave_WritesAggregateConditionalFormattingRuleTypesAsExpressions() {
            string openXmlPath = Path.Combine(Path.GetTempPath(), Guid.NewGuid().ToString("N") + ".xlsx");
            string xlsOutputPath = Path.Combine(Path.GetTempPath(), Guid.NewGuid().ToString("N") + ".xls");

            try {
                using (ExcelDocument document = ExcelDocument.Create(openXmlPath)) {
                    ExcelSheet sheet = document.AddWorksheet("CfAggregates");
                    for (int row = 1; row <= 5; row++) {
                        sheet.CellValue(row, 1, row == 5 ? 2 : row);
                        sheet.CellValue(row, 2, row);
                        sheet.CellValue(row, 3, row * 10);
                        sheet.CellValue(row, 4, row * 10);
                        sheet.CellValue(row, 5, row * 10);
                        sheet.CellValue(row, 6, row * 10);
                        sheet.CellValue(row, 7, row * 10);
                        sheet.CellValue(row, 8, row * 10);
                    }

                    sheet.AddConditionalDuplicateValuesRule("A1:A5");
                    sheet.AddConditionalUniqueValuesRule("B1:B5");
                    sheet.AddConditionalAboveAverageRule("C1:C5");
                    sheet.AddConditionalAboveAverageRule("D1:D5", aboveAverage: false, equalAverage: true, standardDeviation: 1);
                    sheet.AddConditionalTopBottomRule("E1:E5", rank: 2);
                    sheet.AddConditionalTopBottomRule("F1:F5", rank: 2, bottom: true);
                    sheet.AddConditionalTopBottomRule("G1:G5", rank: 40, percent: true);
                    sheet.AddConditionalTopBottomRule("H1:H5", rank: 40, bottom: true, percent: true);

                    document.Save(xlsOutputPath);
                }

                using LegacyXlsLoadResult result = ExcelDocument.LoadLegacyXlsWithReport(xlsOutputPath);
                result.EnsureNoImportErrors();
                Assert.False(result.HasUnsupportedFeatures, FormatUnsupportedFeatures(result.UnsupportedFeatures));

                LegacyXlsWorksheet legacySheet = Assert.Single(result.Workbook.Worksheets);
                Assert.Equal(8, legacySheet.ConditionalFormattings.Count);
                Assert.All(legacySheet.ConditionalFormattings, rule => Assert.Equal(LegacyXlsConditionalFormattingType.Formula, rule.Type));
                Assert.Contains(legacySheet.ConditionalFormattings, rule => rule.Ranges.Single() == "A1:A5" && rule.Formula1 == "COUNTIF($A$1:$A$5,A1)>1");
                Assert.Contains(legacySheet.ConditionalFormattings, rule => rule.Ranges.Single() == "B1:B5" && rule.Formula1 == "COUNTIF($B$1:$B$5,B1)=1");
                Assert.Contains(legacySheet.ConditionalFormattings, rule => rule.Ranges.Single() == "C1:C5" && rule.Formula1 == "C1>AVERAGE($C$1:$C$5)");
                Assert.Contains(legacySheet.ConditionalFormattings, rule => rule.Ranges.Single() == "D1:D5" && rule.Formula1 == "D1<=AVERAGE($D$1:$D$5)-1*STDEV($D$1:$D$5)");
                Assert.Contains(legacySheet.ConditionalFormattings, rule => rule.Ranges.Single() == "E1:E5" && rule.Formula1 == "E1>=LARGE($E$1:$E$5,2)");
                Assert.Contains(legacySheet.ConditionalFormattings, rule => rule.Ranges.Single() == "F1:F5" && rule.Formula1 == "F1<=SMALL($F$1:$F$5,2)");
                Assert.Contains(legacySheet.ConditionalFormattings, rule => rule.Ranges.Single() == "G1:G5" && rule.Formula1 == "G1>=LARGE($G$1:$G$5,ROUNDUP(COUNT($G$1:$G$5)*40/100,0))");
                Assert.Contains(legacySheet.ConditionalFormattings, rule => rule.Ranges.Single() == "H1:H5" && rule.Formula1 == "H1<=SMALL($H$1:$H$5,ROUNDUP(COUNT($H$1:$H$5)*40/100,0))");

                ExcelSheet projectedSheet = result.Document.Sheets[0];
                AssertProjectedExpressionRule(projectedSheet, "A1:A5", "COUNTIF($A$1:$A$5,A1)>1");
                AssertProjectedExpressionRule(projectedSheet, "B1:B5", "COUNTIF($B$1:$B$5,B1)=1");
                AssertProjectedExpressionRule(projectedSheet, "C1:C5", "C1>AVERAGE($C$1:$C$5)");
                AssertProjectedExpressionRule(projectedSheet, "D1:D5", "D1<=AVERAGE($D$1:$D$5)-1*STDEV($D$1:$D$5)");
                AssertProjectedExpressionRule(projectedSheet, "E1:E5", "E1>=LARGE($E$1:$E$5,2)");
                AssertProjectedExpressionRule(projectedSheet, "F1:F5", "F1<=SMALL($F$1:$F$5,2)");
                AssertProjectedExpressionRule(projectedSheet, "G1:G5", "G1>=LARGE($G$1:$G$5,ROUNDUP(COUNT($G$1:$G$5)*40/100,0))");
                AssertProjectedExpressionRule(projectedSheet, "H1:H5", "H1<=SMALL($H$1:$H$5,ROUNDUP(COUNT($H$1:$H$5)*40/100,0))");
            } finally {
                TryDelete(openXmlPath);
                TryDelete(xlsOutputPath);
            }
        }

        [Fact]
        public void LegacyXls_NativeSave_WritesConditionalFormattingFormulasWithNamesAndSheetRanges() {
            string openXmlPath = Path.Combine(Path.GetTempPath(), Guid.NewGuid().ToString("N") + ".xlsx");
            string xlsOutputPath = Path.Combine(Path.GetTempPath(), Guid.NewGuid().ToString("N") + ".xls");

            try {
                using (ExcelDocument document = ExcelDocument.Create(openXmlPath)) {
                    ExcelSheet input = document.AddWorksheet("Input");
                    ExcelSheet firstRegion = document.AddWorksheet("Region 1");
                    ExcelSheet secondRegion = document.AddWorksheet("Region 2");

                    input.CellValue(1, 1, 5d);
                    firstRegion.CellValue(1, 1, 1d);
                    firstRegion.CellValue(2, 1, 2d);
                    secondRegion.CellValue(1, 1, 3d);
                    secondRegion.CellValue(2, 1, 4d);

                    document.SetNamedRange("Threshold", "'Input'!$A$1", save: false);
                    input.AddConditionalFormulaRule("B1:B2", "SUM(Threshold,'Region 1:Region 2'!$A$1:$A$2)>0");

                    document.Save(xlsOutputPath);
                }

                using LegacyXlsLoadResult result = ExcelDocument.LoadLegacyXlsWithReport(xlsOutputPath);
                result.EnsureNoImportErrors();
                Assert.False(result.HasUnsupportedFeatures, FormatUnsupportedFeatures(result.UnsupportedFeatures));

                const string expectedFormula = "SUM(Threshold,'Region 1:Region 2'!$A$1:$A$2)>0";
                LegacyXlsWorksheet legacySheet = result.Workbook.Worksheets[0];
                LegacyXlsConditionalFormatting rule = Assert.Single(legacySheet.ConditionalFormattings);
                Assert.Equal(LegacyXlsConditionalFormattingType.Formula, rule.Type);
                Assert.Equal(new[] { "B1:B2" }, rule.Ranges.ToArray());
                Assert.Equal(expectedFormula, rule.Formula1);

                AssertProjectedExpressionRule(result.Document.Sheets[0], "B1:B2", expectedFormula);
            } finally {
                TryDelete(openXmlPath);
                TryDelete(xlsOutputPath);
            }
        }

        [Theory]
        [MemberData(nameof(LegacyXlsTimePeriodConditionalFormattingCases))]
        public void LegacyXls_NativeSave_WritesSupportedTimePeriodConditionalFormatting(OpenXmlTimePeriodValues timePeriod, string expectedFormula) {
            string openXmlPath = Path.Combine(Path.GetTempPath(), Guid.NewGuid().ToString("N") + ".xlsx");
            string xlsOutputPath = Path.Combine(Path.GetTempPath(), Guid.NewGuid().ToString("N") + ".xls");

            try {
                using (ExcelDocument document = ExcelDocument.Create(openXmlPath)) {
                    ExcelSheet sheet = document.AddWorksheet("CfTime");
                    sheet.CellValue(1, 1, new DateTime(2026, 6, 27));
                    sheet.CellValue(2, 1, new DateTime(2026, 6, 26));
                    sheet.AddConditionalTimePeriodRule("A1:A2", timePeriod);

                    document.Save(xlsOutputPath);
                }

                using LegacyXlsLoadResult result = ExcelDocument.LoadLegacyXlsWithReport(xlsOutputPath);
                result.EnsureNoImportErrors();
                Assert.False(result.HasUnsupportedFeatures, FormatUnsupportedFeatures(result.UnsupportedFeatures));

                LegacyXlsWorksheet legacySheet = Assert.Single(result.Workbook.Worksheets);
                LegacyXlsConditionalFormatting rule = Assert.Single(legacySheet.ConditionalFormattings);
                Assert.Equal(LegacyXlsConditionalFormattingType.Formula, rule.Type);
                Assert.Equal(new[] { "A1:A2" }, rule.Ranges.ToArray());
                Assert.Equal(expectedFormula, rule.Formula1);

                AssertProjectedExpressionRule(result.Document.Sheets[0], "A1:A2", expectedFormula);
            } finally {
                TryDelete(openXmlPath);
                TryDelete(xlsOutputPath);
            }
        }

        public static IEnumerable<object[]> LegacyXlsTimePeriodConditionalFormattingCases {
            get {
                yield return new object[] { OpenXmlTimePeriodValues.Yesterday, "FLOOR(A1,1)=TODAY()-1" };
                yield return new object[] { OpenXmlTimePeriodValues.Today, "FLOOR(A1,1)=TODAY()" };
                yield return new object[] { OpenXmlTimePeriodValues.Tomorrow, "FLOOR(A1,1)=TODAY()+1" };
                yield return new object[] { OpenXmlTimePeriodValues.Last7Days, "AND(TODAY()-FLOOR(A1,1)<=6,FLOOR(A1,1)<=TODAY())" };
                yield return new object[] { OpenXmlTimePeriodValues.LastWeek, "AND(FLOOR(A1,1)>=TODAY()-WEEKDAY(TODAY(),2)-6,FLOOR(A1,1)<=TODAY()-WEEKDAY(TODAY(),2))" };
                yield return new object[] { OpenXmlTimePeriodValues.ThisWeek, "AND(FLOOR(A1,1)>=TODAY()-WEEKDAY(TODAY(),2)+1,FLOOR(A1,1)<=TODAY()-WEEKDAY(TODAY(),2)+7)" };
                yield return new object[] { OpenXmlTimePeriodValues.NextWeek, "AND(FLOOR(A1,1)>=TODAY()-WEEKDAY(TODAY(),2)+8,FLOOR(A1,1)<=TODAY()-WEEKDAY(TODAY(),2)+14)" };
                yield return new object[] { OpenXmlTimePeriodValues.LastMonth, "AND(FLOOR(A1,1)>=DATE(YEAR(TODAY()),MONTH(TODAY())-1,1),FLOOR(A1,1)<DATE(YEAR(TODAY()),MONTH(TODAY()),1))" };
                yield return new object[] { OpenXmlTimePeriodValues.ThisMonth, "AND(FLOOR(A1,1)>=DATE(YEAR(TODAY()),MONTH(TODAY()),1),FLOOR(A1,1)<DATE(YEAR(TODAY()),MONTH(TODAY())+1,1))" };
                yield return new object[] { OpenXmlTimePeriodValues.NextMonth, "AND(FLOOR(A1,1)>=DATE(YEAR(TODAY()),MONTH(TODAY())+1,1),FLOOR(A1,1)<DATE(YEAR(TODAY()),MONTH(TODAY())+2,1))" };
            }
        }

        private static void AssertCellIsRule(
            LegacyXlsWorksheet legacySheet,
            LegacyXlsConditionalFormattingOperator @operator,
            string[] ranges,
            string formula1,
            string? formula2) {
            LegacyXlsConditionalFormatting rule = Assert.Single(legacySheet.ConditionalFormattings, candidate =>
                candidate.Type == LegacyXlsConditionalFormattingType.CellIs
                && candidate.Operator == @operator);
            Assert.Equal(ranges, rule.Ranges.ToArray());
            Assert.Equal(formula1, rule.Formula1);
            Assert.Equal(formula2, rule.Formula2);
        }

        private static void AssertProjectedConditionalRule(
            ExcelSheet sheet,
            string range,
            string @operator,
            string[] formulas) {
            ExcelConditionalFormattingInfo rule = Assert.Single(sheet.GetConditionalFormattingRules(range));
            Assert.Equal("CellIs", rule.Type);
            Assert.Equal(@operator, rule.Operator);
            Assert.Equal(formulas, rule.Formulas);
        }

        private static void AssertProjectedExpressionRule(ExcelSheet sheet, string range, string formula) {
            ExcelConditionalFormattingInfo rule = Assert.Single(sheet.GetConditionalFormattingRules(range));
            Assert.Equal(nameof(OpenXmlConditionalFormatValues.Expression), rule.Type);
            Assert.Equal(new[] { formula }, rule.Formulas);
        }
    }
}
