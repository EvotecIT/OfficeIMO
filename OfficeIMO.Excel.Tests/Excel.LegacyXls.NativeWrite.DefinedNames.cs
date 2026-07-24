using DocumentFormat.OpenXml.Spreadsheet;
using OfficeIMO.Excel;
using OfficeIMO.Excel.LegacyXls;
using OfficeIMO.Excel.LegacyXls.Model;
using Xunit;

namespace OfficeIMO.Tests {
    public partial class Excel {
        [Fact]
        public void LegacyXls_NativeSave_ResolvesSheetLocalDefinedNameFormulasBeforeWorkbookNames() {
            string openXmlPath = Path.Combine(Path.GetTempPath(), Guid.NewGuid().ToString("N") + ".xlsx");
            string xlsOutputPath = Path.Combine(Path.GetTempPath(), Guid.NewGuid().ToString("N") + ".xls");

            try {
                using (ExcelDocument document = ExcelDocument.Create(openXmlPath)) {
                    ExcelSheet north = document.AddWorksheet("North");
                    ExcelSheet south = document.AddWorksheet("South");

                    north.CellValue(1, 1, 99d);
                    north.CellValue(1, 2, 10d);
                    south.CellValue(1, 2, 20d);

                    document.SetNamedRange("LocalRate", "'North'!A1", save: false);
                    document.SetNamedRange("LocalRate", "B1", north, save: false);
                    document.SetNamedRange("LocalRate", "B1", south, save: false);

                    north.CellValue(2, 1, 11d);
                    north.CellFormula(2, 1, "LocalRate+1");
                    south.CellValue(2, 1, 21d);
                    south.CellFormula(2, 1, "LocalRate+1");

                    document.Save(xlsOutputPath);
                }

                using LegacyXlsLoadResult result = ExcelDocument.LoadLegacyXlsWithReport(xlsOutputPath);
                result.EnsureNoImportErrors();
                Assert.False(result.HasUnsupportedFeatures, FormatUnsupportedFeatures(result.UnsupportedFeatures));

                Assert.Equal(3, result.Workbook.DefinedNames.Count);
                Assert.Contains(result.Workbook.DefinedNames, name => name.Name == "LocalRate" && name.LocalSheetIndex == null && name.Reference == "'North'!$A$1");
                Assert.Contains(result.Workbook.DefinedNames, name => name.Name == "LocalRate" && name.LocalSheetIndex == 0 && name.Reference == "'North'!$B$1");
                Assert.Contains(result.Workbook.DefinedNames, name => name.Name == "LocalRate" && name.LocalSheetIndex == 1 && name.Reference == "'South'!$B$1");

                LegacyXlsWorksheet northSheet = result.Workbook.Worksheets[0];
                LegacyXlsWorksheet southSheet = result.Workbook.Worksheets[1];
                AssertNamedFormula(northSheet, 2, 11d);
                AssertNamedFormula(southSheet, 2, 21d);

                Assert.Contains(result.Workbook.FormulaTokenRecords, token =>
                    token.Context == "CellFormula"
                    && token.SheetName == "North"
                    && token.CellReference == "A2"
                    && token.OperandKind == "DefinedName"
                    && token.OperandText == "NameIndex:2");
                Assert.Contains(result.Workbook.FormulaTokenRecords, token =>
                    token.Context == "CellFormula"
                    && token.SheetName == "South"
                    && token.CellReference == "A2"
                    && token.OperandKind == "DefinedName"
                    && token.OperandText == "NameIndex:3");

                Assert.Equal("'North'!$A$1", result.Document.GetNamedRange("LocalRate"));
                Assert.Equal("$B$1", result.Document.GetNamedRange("LocalRate", result.Document.Sheets[0]));
                Assert.Equal("$B$1", result.Document.GetNamedRange("LocalRate", result.Document.Sheets[1]));
            } finally {
                TryDelete(openXmlPath);
                TryDelete(xlsOutputPath);
            }
        }

        [Fact]
        public void LegacyXls_NativeSave_WritesFormulaDefinedNamesInSupportedSubset() {
            string openXmlPath = Path.Combine(Path.GetTempPath(), Guid.NewGuid().ToString("N") + ".xlsx");
            string xlsOutputPath = Path.Combine(Path.GetTempPath(), Guid.NewGuid().ToString("N") + ".xls");

            try {
                using (ExcelDocument document = ExcelDocument.Create(openXmlPath)) {
                    ExcelSheet data = document.AddWorksheet("Data");
                    data.CellValue(1, 1, 41d);

                    document.WorkbookRoot.DefinedNames ??= new DefinedNames();
                    document.WorkbookRoot.DefinedNames.Append(new DefinedName {
                        Name = "FormulaAnswer",
                        Text = "'Data'!$A$1+1"
                    });

                    document.Save(xlsOutputPath);
                }

                using LegacyXlsLoadResult result = ExcelDocument.LoadLegacyXlsWithReport(xlsOutputPath);
                result.EnsureNoImportErrors();
                Assert.False(result.HasUnsupportedFeatures, FormatUnsupportedFeatures(result.UnsupportedFeatures));

                LegacyXlsDefinedName definedName = Assert.Single(result.Workbook.DefinedNames, name => name.Name == "FormulaAnswer");
                Assert.Null(definedName.LocalSheetIndex);
                Assert.Equal("'Data'!$A$1+1", definedName.Reference);
                Assert.Equal("'Data'!$A$1+1", result.Document.GetNamedRange("FormulaAnswer"));
            } finally {
                TryDelete(openXmlPath);
                TryDelete(xlsOutputPath);
            }
        }

        [Fact]
        public void LegacyXls_NativeSave_WritesFormulaDefinedNamesThatReferenceDefinedNames() {
            string openXmlPath = Path.Combine(Path.GetTempPath(), Guid.NewGuid().ToString("N") + ".xlsx");
            string xlsOutputPath = Path.Combine(Path.GetTempPath(), Guid.NewGuid().ToString("N") + ".xls");

            try {
                using (ExcelDocument document = ExcelDocument.Create(openXmlPath)) {
                    ExcelSheet data = document.AddWorksheet("Data");
                    data.CellValue(1, 1, 41d);
                    data.CellValue(2, 1, 43d);
                    data.CellFormula(2, 1, "AdjustedAnswer+1");

                    document.WorkbookRoot.DefinedNames ??= new DefinedNames();
                    document.WorkbookRoot.DefinedNames.Append(new DefinedName {
                        Name = "AdjustedAnswer",
                        Text = "BaseAnswer+1"
                    });
                    document.WorkbookRoot.DefinedNames.Append(new DefinedName {
                        Name = "BaseAnswer",
                        Text = "'Data'!$A$1"
                    });

                    document.Save(xlsOutputPath);
                }

                using LegacyXlsLoadResult result = ExcelDocument.LoadLegacyXlsWithReport(xlsOutputPath);
                result.EnsureNoImportErrors();
                Assert.False(result.HasUnsupportedFeatures, FormatUnsupportedFeatures(result.UnsupportedFeatures));

                Assert.Contains(result.Workbook.DefinedNames, name => name.Name == "BaseAnswer" && name.Reference == "'Data'!$A$1");
                Assert.Contains(result.Workbook.DefinedNames, name => name.Name == "AdjustedAnswer" && name.Reference == "BaseAnswer+1");
                Assert.Equal("BaseAnswer+1", result.Document.GetNamedRange("AdjustedAnswer"));

                LegacyXlsCell formulaCell = Assert.Single(result.Workbook.Worksheets[0].Cells, cell => cell.Row == 2 && cell.Column == 1);
                Assert.True(formulaCell.IsFormula);
                Assert.Equal(43d, Assert.IsType<double>(formulaCell.Value));
                Assert.Equal("AdjustedAnswer+1", formulaCell.FormulaText);

                Assert.Contains(result.Workbook.FormulaTokenRecords, token =>
                    token.Context == "DefinedName"
                    && token.CellReference == "AdjustedAnswer"
                    && token.OperandKind == "DefinedName"
                    && token.OperandText == "NameIndex:2");
                Assert.Contains(result.Workbook.FormulaTokenRecords, token =>
                    token.Context == "CellFormula"
                    && token.CellReference == "A2"
                    && token.OperandKind == "DefinedName"
                    && token.OperandText == "NameIndex:1");
            } finally {
                TryDelete(openXmlPath);
                TryDelete(xlsOutputPath);
            }
        }

        [Fact]
        public void LegacyXls_NativeSave_WritesFormulaDefinedNamesWithSheetRangeReferences() {
            string openXmlPath = Path.Combine(Path.GetTempPath(), Guid.NewGuid().ToString("N") + ".xlsx");
            string xlsOutputPath = Path.Combine(Path.GetTempPath(), Guid.NewGuid().ToString("N") + ".xls");

            try {
                using (ExcelDocument document = ExcelDocument.Create(openXmlPath)) {
                    ExcelSheet summary = document.AddWorksheet("Summary");
                    ExcelSheet first = document.AddWorksheet("Region 1");
                    ExcelSheet second = document.AddWorksheet("Region 2");

                    first.CellValue(1, 1, 10d);
                    first.CellValue(2, 1, 20d);
                    second.CellValue(1, 1, 30d);
                    second.CellValue(2, 1, 40d);
                    summary.CellValue(1, 1, 100d);
                    summary.CellFormula(1, 1, "RegionTotal");

                    document.WorkbookRoot.DefinedNames ??= new DefinedNames();
                    document.WorkbookRoot.DefinedNames.Append(new DefinedName {
                        Name = "RegionTotal",
                        Text = "SUM('Region 1:Region 2'!$A$1:$A$2)"
                    });

                    document.Save(xlsOutputPath);
                }

                using LegacyXlsLoadResult result = ExcelDocument.LoadLegacyXlsWithReport(xlsOutputPath);
                result.EnsureNoImportErrors();
                Assert.False(result.HasUnsupportedFeatures, FormatUnsupportedFeatures(result.UnsupportedFeatures));

                LegacyXlsDefinedName definedName = Assert.Single(result.Workbook.DefinedNames, name => name.Name == "RegionTotal");
                Assert.Equal("SUM('Region 1:Region 2'!$A$1:$A$2)", definedName.Reference);
                Assert.Equal("SUM('Region 1:Region 2'!$A$1:$A$2)", result.Document.GetNamedRange("RegionTotal"));

                LegacyXlsWorksheet summarySheet = result.Workbook.Worksheets[0];
                LegacyXlsCell formulaCell = Assert.Single(summarySheet.Cells, cell => cell.Row == 1 && cell.Column == 1);
                Assert.True(formulaCell.IsFormula);
                Assert.Equal(100d, Assert.IsType<double>(formulaCell.Value));
                Assert.Equal("RegionTotal", formulaCell.FormulaText);

                Assert.Contains(result.Workbook.FormulaTokenRecords, token =>
                    token.Context == "DefinedName"
                    && token.CellReference == "RegionTotal"
                    && token.TokenName == "PtgArea3d");
            } finally {
                TryDelete(openXmlPath);
                TryDelete(xlsOutputPath);
            }
        }

        [Fact]
        public void LegacyXls_NativeSave_WritesDefinedNameFormulasWithExternalWorkbookReferences() {
            string openXmlPath = Path.Combine(Path.GetTempPath(), Guid.NewGuid().ToString("N") + ".xlsx");
            string xlsOutputPath = Path.Combine(Path.GetTempPath(), Guid.NewGuid().ToString("N") + ".xls");

            try {
                using (ExcelDocument document = ExcelDocument.Create(openXmlPath)) {
                    ExcelSheet sheet = document.AddWorksheet("Names");
                    sheet.CellValue(1, 1, "External workbook");
                    document.WorkbookRoot.DefinedNames ??= new DefinedNames();
                    document.WorkbookRoot.DefinedNames.Append(new DefinedName {
                        Name = "ExternalWorkbookCell",
                        Text = "'[Budget.xls]Data'!$A$1"
                    });

                    document.Save(xlsOutputPath);
                }

                using LegacyXlsLoadResult result = ExcelDocument.LoadLegacyXlsWithReport(
                    xlsOutputPath,
                    new LegacyXlsImportOptions { PreserveExternalWorkbookLinks = true });
                result.EnsureNoImportErrors();
                LegacyXlsExternalReference externalReference = Assert.Single(
                    result.Workbook.ExternalReferences,
                    reference => reference.Kind == LegacyXlsExternalReferenceKind.ExternalWorkbook);
                Assert.Equal("Budget.xls", externalReference.Target);
                Assert.Equal(new[] { "Data" }, externalReference.SheetNames);

                LegacyXlsDefinedName definedName = Assert.Single(result.Workbook.DefinedNames, name => name.Name == "ExternalWorkbookCell");
                Assert.Equal("'[Budget.xls]Data'!$A$1", definedName.Reference);
                Assert.Equal("'[Budget.xls]Data'!$A$1", result.Document.GetNamedRange("ExternalWorkbookCell"));
                Assert.Contains(result.Workbook.FormulaTokenRecords, token =>
                    token.Context == "DefinedName"
                    && token.CellReference == "ExternalWorkbookCell"
                    && token.TokenName == "PtgRef3d");
            } finally {
                TryDelete(openXmlPath);
                TryDelete(xlsOutputPath);
            }
        }

        [Fact]
        public void LegacyXls_NativeSave_WritesExternalDefinedNameReferencesInDefinedNames() {
            string openXmlPath = Path.Combine(Path.GetTempPath(), Guid.NewGuid().ToString("N") + ".xlsx");
            string xlsOutputPath = Path.Combine(Path.GetTempPath(), Guid.NewGuid().ToString("N") + ".xls");

            try {
                using (ExcelDocument document = ExcelDocument.Create(openXmlPath)) {
                    ExcelSheet sheet = document.AddWorksheet("Names");
                    sheet.CellValue(1, 1, "External name");
                    document.WorkbookRoot.DefinedNames ??= new DefinedNames();
                    document.WorkbookRoot.DefinedNames.Append(new DefinedName {
                        Name = "ExternalNameFormula",
                        Text = "[Budget.xls]SalesTotal"
                    });

                    document.Save(xlsOutputPath);
                }

                using LegacyXlsLoadResult result = ExcelDocument.LoadLegacyXlsWithReport(
                    xlsOutputPath,
                    new LegacyXlsImportOptions { PreserveExternalWorkbookLinks = true });
                result.EnsureNoImportErrors();
                LegacyXlsExternalReference externalReference = Assert.Single(
                    result.Workbook.ExternalReferences,
                    reference => reference.Kind == LegacyXlsExternalReferenceKind.ExternalWorkbook);
                Assert.Equal("Budget.xls", externalReference.Target);
                LegacyXlsExternalName externalName = Assert.Single(externalReference.ExternalNames);
                Assert.Equal("SalesTotal", externalName.Name);

                LegacyXlsDefinedName definedName = Assert.Single(result.Workbook.DefinedNames, name => name.Name == "ExternalNameFormula");
                Assert.Equal("'Budget.xls'!SalesTotal", definedName.Reference);
                Assert.Equal("'Budget.xls'!SalesTotal", result.Document.GetNamedRange("ExternalNameFormula"));
                Assert.Contains(result.Workbook.FormulaTokenRecords, token =>
                    token.Context == "DefinedName"
                    && token.CellReference == "ExternalNameFormula"
                    && token.TokenName == "PtgNameX");
            } finally {
                TryDelete(openXmlPath);
                TryDelete(xlsOutputPath);
            }
        }

        [Fact]
        public void LegacyXls_NativeSave_WritesSheetScopedExternalDefinedNameFormulaReferences() {
            string openXmlPath = Path.Combine(Path.GetTempPath(), Guid.NewGuid().ToString("N") + ".xlsx");
            string xlsOutputPath = Path.Combine(Path.GetTempPath(), Guid.NewGuid().ToString("N") + ".xls");

            try {
                using (ExcelDocument document = ExcelDocument.Create(openXmlPath)) {
                    ExcelSheet sheet = document.AddWorksheet("ExternalNames");
                    sheet.CellValue(1, 1, 0.25d);
                    sheet.CellFormula(1, 1, "'[Budget.xls]Feb'!TaxRate");

                    document.Save(xlsOutputPath);
                }

                using LegacyXlsLoadResult result = ExcelDocument.LoadLegacyXlsWithReport(xlsOutputPath);
                result.EnsureNoImportErrors();

                LegacyXlsExternalReference externalReference = Assert.Single(
                    result.Workbook.ExternalReferences,
                    reference => reference.Kind == LegacyXlsExternalReferenceKind.ExternalWorkbook);
                Assert.Equal("Budget.xls", externalReference.Target);
                Assert.Equal(new[] { "Feb" }, externalReference.SheetNames);
                LegacyXlsExternalName externalName = Assert.Single(externalReference.ExternalNames);
                Assert.Equal("TaxRate", externalName.Name);
                Assert.Equal(0, externalName.LocalSheetIndex);
                Assert.Equal(LegacyXlsExternalNameBodyKind.ExternalDefinedName, externalName.BodyKind);

                LegacyXlsWorksheet worksheet = Assert.Single(result.Workbook.Worksheets);
                AssertNumericFormula(worksheet, 1, 0.25d, "'[Budget.xls]Feb'!TaxRate");
                Assert.Contains(result.Workbook.FormulaTokenRecords, token =>
                    token.TokenName == "PtgNameX"
                    && token.OperandKind == "ExternalName");
            } finally {
                TryDelete(openXmlPath);
                TryDelete(xlsOutputPath);
            }
        }

        [Fact]
        public void LegacyXls_NativeSave_WritesSheetScopedExternalDefinedNameReferencesInDefinedNames() {
            string openXmlPath = Path.Combine(Path.GetTempPath(), Guid.NewGuid().ToString("N") + ".xlsx");
            string xlsOutputPath = Path.Combine(Path.GetTempPath(), Guid.NewGuid().ToString("N") + ".xls");

            try {
                using (ExcelDocument document = ExcelDocument.Create(openXmlPath)) {
                    ExcelSheet sheet = document.AddWorksheet("Names");
                    sheet.CellValue(1, 1, "Scoped external name");
                    document.WorkbookRoot.DefinedNames ??= new DefinedNames();
                    document.WorkbookRoot.DefinedNames.Append(new DefinedName {
                        Name = "ScopedExternalNameFormula",
                        Text = "'[Budget.xls]Feb'!SalesTotal"
                    });

                    document.Save(xlsOutputPath);
                }

                using LegacyXlsLoadResult result = ExcelDocument.LoadLegacyXlsWithReport(
                    xlsOutputPath,
                    new LegacyXlsImportOptions { PreserveExternalWorkbookLinks = true });
                result.EnsureNoImportErrors();

                LegacyXlsExternalReference externalReference = Assert.Single(
                    result.Workbook.ExternalReferences,
                    reference => reference.Kind == LegacyXlsExternalReferenceKind.ExternalWorkbook);
                Assert.Equal("Budget.xls", externalReference.Target);
                Assert.Equal(new[] { "Feb" }, externalReference.SheetNames);
                LegacyXlsExternalName externalName = Assert.Single(externalReference.ExternalNames);
                Assert.Equal("SalesTotal", externalName.Name);
                Assert.Equal(0, externalName.LocalSheetIndex);

                LegacyXlsDefinedName definedName = Assert.Single(result.Workbook.DefinedNames, name => name.Name == "ScopedExternalNameFormula");
                Assert.Equal("'[Budget.xls]Feb'!SalesTotal", definedName.Reference);
                Assert.Equal("'[Budget.xls]Feb'!SalesTotal", result.Document.GetNamedRange("ScopedExternalNameFormula"));
                Assert.Contains(result.Workbook.FormulaTokenRecords, token =>
                    token.Context == "DefinedName"
                    && token.CellReference == "ScopedExternalNameFormula"
                    && token.TokenName == "PtgNameX");
            } finally {
                TryDelete(openXmlPath);
                TryDelete(xlsOutputPath);
            }
        }

        [Fact]
        public void LegacyXls_NativeSave_BlocksOversizedDefinedNameFormulaPayloadsBeforeWriting() {
            AssertNativeXlsSaveNotSupported("defined-name formula payload lengths outside BIFF8 limits", (document, sheet) => {
                sheet.CellValue(1, 1, "Oversized defined-name formula");
                string longLiteral = "\"" + new string('A', 255) + "\"";
                string formula = string.Join("&", Enumerable.Repeat(longLiteral, 260));

                document.WorkbookRoot.DefinedNames ??= new DefinedNames();
                document.WorkbookRoot.DefinedNames.Append(new DefinedName {
                    Name = "TooLongFormula",
                    Text = formula
                });
            });
        }

        private static void AssertNamedFormula(LegacyXlsWorksheet worksheet, int row, double expectedValue) {
            LegacyXlsCell cell = Assert.Single(worksheet.Cells, item => item.Row == row && item.Column == 1);
            Assert.True(cell.IsFormula);
            Assert.Equal(expectedValue, Assert.IsType<double>(cell.Value));
            Assert.Equal("LocalRate+1", cell.FormulaText);
        }
    }
}
