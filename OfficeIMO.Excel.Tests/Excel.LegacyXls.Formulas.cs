using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using OfficeIMO.Excel;
using OfficeIMO.Excel.LegacyXls;
using OfficeIMO.Excel.LegacyXls.Biff;
using OfficeIMO.Excel.LegacyXls.Diagnostics;
using OfficeIMO.Excel.LegacyXls.Model;
using System.Text;
using Xunit;

namespace OfficeIMO.Tests {
    public partial class Excel {
        [Fact]
        public void LegacyXls_Load_ImportsFormulaStringLiteralTokens() {
            byte[] workbookStream = LegacyXlsTestWorkbookBuilder.CreateFormulaStringLiteralWorkbookStream();
            byte[] compound = LegacyXlsCompoundTestBuilder.CreateWorkbookCompoundFile(workbookStream);

            LegacyXlsWorkbook legacy = LegacyXlsWorkbook.Load(compound, new LegacyXlsImportOptions {
                ReportUnsupportedContent = false
            });

            Assert.DoesNotContain(legacy.Diagnostics, d => d.Severity == LegacyXlsDiagnosticSeverity.Error);
            LegacyXlsWorksheet sheet = Assert.Single(legacy.Worksheets);
            LegacyXlsCell formula = Assert.Single(sheet.Cells, cell => cell.Row == 1 && cell.Column == 2);
            Assert.True(formula.IsFormula);
            Assert.Equal("Hello Target", formula.Value);
            Assert.Equal("\"Hello \"&A1", formula.FormulaText);

            using ExcelDocument document = ExcelDocument.LoadLegacyXls(new MemoryStream(compound), new LegacyXlsImportOptions {
                ReportUnsupportedContent = false
            });

            Assert.True(document.Sheets[0].TryGetCellText(1, 2, out string? cachedText));
            Assert.Equal("Hello Target", cachedText);

            using var output = new MemoryStream();
            document.Save(output);
            using SpreadsheetDocument spreadsheet = SpreadsheetDocument.Open(new MemoryStream(output.ToArray()), false);
            WorksheetPart worksheetPart = spreadsheet.WorkbookPart!.WorksheetParts.Single();
            Cell projectedFormula = worksheetPart.Worksheet.Descendants<Cell>().Single(cell => cell.CellReference!.Value == "B1");
            Assert.Equal("\"Hello \"&A1", projectedFormula.CellFormula!.Text);
        }

        [Fact]
        public void LegacyXls_Load_ImportsFormula3dReferences() {
            byte[] workbookStream = LegacyXlsTestWorkbookBuilder.CreateFormula3dReferenceWorkbookStream();
            byte[] compound = LegacyXlsCompoundTestBuilder.CreateWorkbookCompoundFile(workbookStream);

            LegacyXlsWorkbook legacy = LegacyXlsWorkbook.Load(compound, new LegacyXlsImportOptions {
                ReportUnsupportedContent = false
            });

            Assert.DoesNotContain(legacy.Diagnostics, d => d.Severity == LegacyXlsDiagnosticSeverity.Error);
            Assert.Equal(2, legacy.Worksheets.Count);
            LegacyXlsWorksheet totals = legacy.Worksheets[1];
            LegacyXlsCell referenceFormula = Assert.Single(totals.Cells, cell => cell.Row == 1 && cell.Column == 1);
            Assert.True(referenceFormula.IsFormula);
            Assert.Equal(15d, referenceFormula.Value);
            Assert.Equal("'Input Data'!A1+5", referenceFormula.FormulaText);
            LegacyXlsCell areaFormula = Assert.Single(totals.Cells, cell => cell.Row == 2 && cell.Column == 1);
            Assert.True(areaFormula.IsFormula);
            Assert.Equal(42d, areaFormula.Value);
            Assert.Equal("SUM('Input Data'!A1:A2)", areaFormula.FormulaText);

            using ExcelDocument document = ExcelDocument.LoadLegacyXls(new MemoryStream(compound), new LegacyXlsImportOptions {
                ReportUnsupportedContent = false
            });

            Assert.True(document.Sheets[1].TryGetCellText(1, 1, out string? referenceValue));
            Assert.Equal("15", referenceValue);
            Assert.True(document.Sheets[1].TryGetCellText(2, 1, out string? areaValue));
            Assert.Equal("42", areaValue);

            using var output = new MemoryStream();
            document.Save(output);
            using SpreadsheetDocument spreadsheet = SpreadsheetDocument.Open(new MemoryStream(output.ToArray()), false);
            WorkbookPart workbookPart = spreadsheet.WorkbookPart!;
            Sheet totalsSheet = workbookPart.Workbook.Sheets!.Elements<Sheet>().Single(sheet => sheet.Name == "Totals");
            WorksheetPart worksheetPart = (WorksheetPart)workbookPart.GetPartById(totalsSheet.Id!);
            Cell projectedReferenceFormula = worksheetPart.Worksheet.Descendants<Cell>().Single(cell => cell.CellReference!.Value == "A1");
            Assert.Equal("'Input Data'!A1+5", projectedReferenceFormula.CellFormula!.Text);
            Cell projectedAreaFormula = worksheetPart.Worksheet.Descendants<Cell>().Single(cell => cell.CellReference!.Value == "A2");
            Assert.Equal("SUM('Input Data'!A1:A2)", projectedAreaFormula.CellFormula!.Text);
        }

        [Fact]
        public void LegacyXls_Load_ImportsFormulaExternalWorkbookReferences() {
            byte[] workbookStream = LegacyXlsTestWorkbookBuilder.CreateFormulaExternalWorkbookReferenceWorkbookStream();
            byte[] compound = LegacyXlsCompoundTestBuilder.CreateWorkbookCompoundFile(workbookStream);

            LegacyXlsWorkbook legacy = LegacyXlsWorkbook.Load(compound, new LegacyXlsImportOptions {
                ReportUnsupportedContent = true,
                PreserveExternalWorkbookLinks = true
            });

            Assert.DoesNotContain(legacy.Diagnostics, d => d.Severity == LegacyXlsDiagnosticSeverity.Error);
            Assert.DoesNotContain(legacy.Diagnostics, d => d.Code == "XLS-BIFF-FORMULA-TOKENS-UNSUPPORTED");
            LegacyXlsExternalReference externalReference = Assert.Single(legacy.ExternalReferences);
            Assert.Equal(LegacyXlsExternalReferenceKind.ExternalWorkbook, externalReference.Kind);
            Assert.Equal("C:\\Data\\Budget.xls", externalReference.Target);
            Assert.Equal(new[] { "Jan", "Feb" }, externalReference.SheetNames);

            LegacyXlsWorksheet sheet = Assert.Single(legacy.Worksheets);
            LegacyXlsCell referenceFormula = Assert.Single(sheet.Cells, cell => cell.Row == 1 && cell.Column == 1);
            Assert.True(referenceFormula.IsFormula);
            Assert.Equal(15d, referenceFormula.Value);
            Assert.Equal("'[Budget.xls]Jan'!A1+5", referenceFormula.FormulaText);
            LegacyXlsCell areaFormula = Assert.Single(sheet.Cells, cell => cell.Row == 2 && cell.Column == 1);
            Assert.True(areaFormula.IsFormula);
            Assert.Equal(42d, areaFormula.Value);
            Assert.Equal("SUM('[Budget.xls]Jan'!A1:A2)", areaFormula.FormulaText);

            using ExcelDocument document = ExcelDocument.LoadLegacyXls(new MemoryStream(compound), new LegacyXlsImportOptions {
                ReportUnsupportedContent = true,
                PreserveExternalWorkbookLinks = true
            });

            using var output = new MemoryStream();
            document.Save(output);
            using SpreadsheetDocument spreadsheet = SpreadsheetDocument.Open(new MemoryStream(output.ToArray()), false);
            WorkbookPart workbookPart = spreadsheet.WorkbookPart!;
            Assert.NotNull(workbookPart.Workbook.ExternalReferences);
            DocumentFormat.OpenXml.Spreadsheet.ExternalReference projectedExternalReference =
                Assert.Single(workbookPart.Workbook.ExternalReferences!.Elements<DocumentFormat.OpenXml.Spreadsheet.ExternalReference>());
            ExternalWorkbookPart externalWorkbookPart = Assert.IsType<ExternalWorkbookPart>(workbookPart.GetPartById(projectedExternalReference.Id!));
            ExternalRelationship externalRelationship = Assert.Single(externalWorkbookPart.ExternalRelationships);
            Assert.Equal("http://schemas.openxmlformats.org/officeDocument/2006/relationships/externalLinkPath", externalRelationship.RelationshipType);
            Assert.EndsWith("Budget.xls", externalRelationship.Uri.OriginalString, StringComparison.Ordinal);
            ExternalBook projectedExternalBook = externalWorkbookPart.ExternalLink!.GetFirstChild<ExternalBook>()!;
            Assert.NotNull(projectedExternalBook);
            Assert.Equal(new[] { "Jan", "Feb" }, projectedExternalBook.SheetNames!.Elements<SheetName>().Select(sheetName => sheetName.Val!.Value).ToArray());

            WorksheetPart worksheetPart = workbookPart.WorksheetParts.Single();
            Cell projectedReferenceFormula = worksheetPart.Worksheet.Descendants<Cell>().Single(cell => cell.CellReference!.Value == "A1");
            Assert.Equal("'[Budget.xls]Jan'!A1+5", projectedReferenceFormula.CellFormula!.Text);
            Cell projectedAreaFormula = worksheetPart.Worksheet.Descendants<Cell>().Single(cell => cell.CellReference!.Value == "A2");
            Assert.Equal("SUM('[Budget.xls]Jan'!A1:A2)", projectedAreaFormula.CellFormula!.Text);
        }

        [Fact]
        public void LegacyXls_Load_DropsExternalWorkbookFormulasByDefaultButKeepsCachedValues() {
            byte[] workbookStream = LegacyXlsTestWorkbookBuilder.CreateFormulaExternalWorkbookReferenceWorkbookStream();
            byte[] compound = LegacyXlsCompoundTestBuilder.CreateWorkbookCompoundFile(workbookStream);

            using ExcelDocument document = ExcelDocument.LoadLegacyXls(new MemoryStream(compound));
            using var output = new MemoryStream();
            document.Save(output, new ExcelSaveOptions {
                LossPolicy = ExcelConversionLossPolicy.Allow
            });
            using SpreadsheetDocument spreadsheet = SpreadsheetDocument.Open(new MemoryStream(output.ToArray()), false);

            WorkbookPart workbookPart = spreadsheet.WorkbookPart!;
            Assert.Empty(workbookPart.GetPartsOfType<ExternalWorkbookPart>());
            Cell[] cells = workbookPart.WorksheetParts.Single().Worksheet.Descendants<Cell>()
                .OrderBy(cell => cell.CellReference!.Value, StringComparer.Ordinal)
                .ToArray();
            Assert.Equal(2, cells.Length);
            Assert.All(cells, cell => Assert.Null(cell.CellFormula));
            Assert.Equal("15", cells[0].CellValue!.Text);
            Assert.Equal("42", cells[1].CellValue!.Text);
        }

        [Fact]
        public void LegacyXls_Load_DropsExternalFormulaWhenWorkbookNameContainsApostrophe() {
            byte[] workbookStream = LegacyXlsTestWorkbookBuilder.CreateFormulaExternalWorkbookReferenceWorkbookStream("C:\\Data\\O'Brien.xls");
            byte[] compound = LegacyXlsCompoundTestBuilder.CreateWorkbookCompoundFile(workbookStream);

            using ExcelDocument document = ExcelDocument.LoadLegacyXls(new MemoryStream(compound));
            using var output = new MemoryStream();
            document.Save(output, new ExcelSaveOptions { LossPolicy = ExcelConversionLossPolicy.Allow });
            using SpreadsheetDocument spreadsheet = SpreadsheetDocument.Open(new MemoryStream(output.ToArray()), false);

            Cell[] cells = spreadsheet.WorkbookPart!.WorksheetParts.Single().Worksheet.Descendants<Cell>()
                .OrderBy(cell => cell.CellReference!.Value, StringComparer.Ordinal)
                .ToArray();
            Assert.All(cells, cell => Assert.Null(cell.CellFormula));
            Assert.Empty(spreadsheet.WorkbookPart.GetPartsOfType<ExternalWorkbookPart>());
        }

        [Fact]
        public void LegacyXls_Load_PreservesLocalFormulaWhoseStringLiteralMentionsExternalWorkbook() {
            byte[] workbookStream = LegacyXlsTestWorkbookBuilder.CreateFormulaStringLiteralMatchingExternalWorkbookStream();
            byte[] compound = LegacyXlsCompoundTestBuilder.CreateWorkbookCompoundFile(workbookStream);

            using ExcelDocument document = ExcelDocument.LoadLegacyXls(new MemoryStream(compound));
            using var output = new MemoryStream();
            document.Save(output, new ExcelSaveOptions { LossPolicy = ExcelConversionLossPolicy.Allow });
            using SpreadsheetDocument spreadsheet = SpreadsheetDocument.Open(new MemoryStream(output.ToArray()), false);

            WorksheetPart worksheetPart = spreadsheet.WorkbookPart!.WorksheetParts.Single();
            Cell formulaCell = worksheetPart.Worksheet.Descendants<Cell>().Single(cell => cell.CellReference!.Value == "B1");
            Assert.Equal("\"[Budget.xls]Jan\"&A1", formulaCell.CellFormula!.Text);
            Assert.Empty(spreadsheet.WorkbookPart.GetPartsOfType<ExternalWorkbookPart>());
        }

        [Fact]
        public void LegacyXls_Load_ImportsFormulaExternalDefinedNames() {
            byte[] workbookStream = LegacyXlsTestWorkbookBuilder.CreateFormulaExternalDefinedNameWorkbookStream();
            byte[] compound = LegacyXlsCompoundTestBuilder.CreateWorkbookCompoundFile(workbookStream);

            LegacyXlsWorkbook legacy = LegacyXlsWorkbook.Load(compound, new LegacyXlsImportOptions {
                ReportUnsupportedContent = true,
                PreserveExternalWorkbookLinks = true
            });

            Assert.DoesNotContain(legacy.Diagnostics, d => d.Severity == LegacyXlsDiagnosticSeverity.Error);
            Assert.DoesNotContain(legacy.Diagnostics, d => d.Code == "XLS-BIFF-FORMULA-TOKENS-UNSUPPORTED");
            LegacyXlsExternalReference externalReference = Assert.Single(legacy.ExternalReferences);
            LegacyXlsExternalName externalName = Assert.Single(externalReference.ExternalNames);
            Assert.Equal(LegacyXlsExternalReferenceKind.ExternalWorkbook, externalReference.Kind);
            Assert.Equal("TaxRate", externalName.Name);

            LegacyXlsWorksheet sheet = Assert.Single(legacy.Worksheets);
            LegacyXlsCell formula = Assert.Single(sheet.Cells, cell => cell.Row == 1 && cell.Column == 1);
            Assert.True(formula.IsFormula);
            Assert.Equal(0.25d, formula.Value);
            Assert.Equal("'Budget.xls'!TaxRate", formula.FormulaText);

            using ExcelDocument document = ExcelDocument.LoadLegacyXls(new MemoryStream(compound), new LegacyXlsImportOptions {
                ReportUnsupportedContent = true,
                PreserveExternalWorkbookLinks = true
            });

            using var output = new MemoryStream();
            document.Save(output);
            using SpreadsheetDocument spreadsheet = SpreadsheetDocument.Open(new MemoryStream(output.ToArray()), false);
            WorkbookPart workbookPart = spreadsheet.WorkbookPart!;
            ExternalWorkbookPart externalWorkbookPart = Assert.Single(workbookPart.GetPartsOfType<ExternalWorkbookPart>());
            ExternalBook projectedExternalBook = externalWorkbookPart.ExternalLink!.GetFirstChild<ExternalBook>()!;
            Assert.NotNull(projectedExternalBook);
            ExternalDefinedName projectedExternalName = Assert.Single(projectedExternalBook.ExternalDefinedNames!.Elements<ExternalDefinedName>());
            Assert.Equal("TaxRate", projectedExternalName.Name!.Value);

            WorksheetPart worksheetPart = workbookPart.WorksheetParts.Single();
            Cell projectedFormula = worksheetPart.Worksheet.Descendants<Cell>().Single(cell => cell.CellReference!.Value == "A1");
            Assert.Equal("'Budget.xls'!TaxRate", projectedFormula.CellFormula!.Text);
        }

        [Fact]
        public void LegacyXls_Load_DropsExternalDefinedNameFormulasByDefaultButKeepsCachedValue() {
            byte[] workbookStream = LegacyXlsTestWorkbookBuilder.CreateFormulaExternalDefinedNameWorkbookStream();
            byte[] compound = LegacyXlsCompoundTestBuilder.CreateWorkbookCompoundFile(workbookStream);

            using ExcelDocument document = ExcelDocument.LoadLegacyXls(new MemoryStream(compound));
            using var output = new MemoryStream();
            document.Save(output, new ExcelSaveOptions {
                LossPolicy = ExcelConversionLossPolicy.Allow
            });
            using SpreadsheetDocument spreadsheet = SpreadsheetDocument.Open(new MemoryStream(output.ToArray()), false);

            WorkbookPart workbookPart = spreadsheet.WorkbookPart!;
            Assert.Empty(workbookPart.GetPartsOfType<ExternalWorkbookPart>());
            Cell cell = Assert.Single(workbookPart.WorksheetParts.Single().Worksheet.Descendants<Cell>());
            Assert.Null(cell.CellFormula);
            Assert.Equal("0.25", cell.CellValue!.Text);
        }

        [Fact]
        public void LegacyXls_Load_ImportsFormulaAddInUserDefinedFunctions() {
            byte[] workbookStream = LegacyXlsTestWorkbookBuilder.CreateFormulaAddInUserDefinedFunctionWorkbookStream();
            byte[] compound = LegacyXlsCompoundTestBuilder.CreateWorkbookCompoundFile(workbookStream);

            LegacyXlsWorkbook legacy = LegacyXlsWorkbook.Load(compound, new LegacyXlsImportOptions {
                ReportUnsupportedContent = true
            });

            Assert.DoesNotContain(legacy.Diagnostics, d => d.Severity == LegacyXlsDiagnosticSeverity.Error);
            Assert.DoesNotContain(legacy.Diagnostics, d => d.Code == "XLS-BIFF-FORMULA-TOKENS-UNSUPPORTED");
            LegacyXlsExternalReference addInReference = Assert.Single(legacy.ExternalReferences);
            Assert.Equal(LegacyXlsExternalReferenceKind.AddIn, addInReference.Kind);
            LegacyXlsExternalName addInName = Assert.Single(addInReference.ExternalNames);
            Assert.Equal("MYUDF", addInName.Name);
            Assert.Equal(LegacyXlsExternalNameBodyKind.AddInUdf, addInName.BodyKind);
            Assert.False(addInName.BuiltIn);
            Assert.Equal(0, addInName.CachedClipboardFormat);
            LegacyXlsImportReport report = legacy.CreateImportReport();
            Assert.Equal(1, report.ExternalNamesByBodyKind["AddInUdf"]);
            Assert.Equal(1, report.ExternalNamesByFlagShape["Body:AddInUdf|BuiltIn:Missing|Advise:Missing|Picture:Missing|Ole:Missing|OleLink:Missing|Icon:Missing"]);
            Assert.Equal(1, report.FormulaFunctionsById["Function:0x00FF"]);
            Assert.Equal(1, report.FormulaFunctionsByName["UserDefinedFunction"]);
            Assert.Equal(1, report.FormulaFunctionsByParameterCount["UserDefinedFunction|Args:2"]);
            Assert.Contains(legacy.FormulaTokenRecords, record =>
                record.TokenName == "PtgFuncVar"
                && record.FunctionId == 0x00ff
                && record.FunctionName == "UserDefinedFunction"
                && record.FunctionParameterCount == 2
                && record.OperandKind == "VariableFunction"
                && record.OperandText == "UserDefinedFunction");

            LegacyXlsCell formula = Assert.Single(Assert.Single(legacy.Worksheets).Cells);
            Assert.True(formula.IsFormula);
            Assert.Equal(20d, formula.Value);
            Assert.Equal("MYUDF(10)", formula.FormulaText);
        }

        [Fact]
        public void LegacyXls_Load_ImportsFormula3dReferencesAfterUnsupportedBoundSheets() {
            byte[] workbookStream = LegacyXlsTestWorkbookBuilder.CreateFormula3dReferenceAfterChartSheetWorkbookStream();
            byte[] compound = LegacyXlsCompoundTestBuilder.CreateWorkbookCompoundFile(workbookStream);

            LegacyXlsWorkbook legacy = LegacyXlsWorkbook.Load(compound, new LegacyXlsImportOptions {
                ReportUnsupportedContent = true
            });

            Assert.DoesNotContain(legacy.Diagnostics, d => d.Severity == LegacyXlsDiagnosticSeverity.Error);
            Assert.DoesNotContain(legacy.Diagnostics, d => d.Code == "XLS-BIFF-FORMULA-TOKENS-UNSUPPORTED");
            Assert.Equal(2, legacy.Worksheets.Count);
            LegacyXlsChartSheet chartSheet = Assert.Single(legacy.ChartSheets);
            Assert.Equal("Chart1", chartSheet.Name);
            Assert.Empty(legacy.UnsupportedSheets);

            LegacyXlsCell formula = Assert.Single(legacy.Worksheets[0].Cells, cell => cell.Row == 1 && cell.Column == 1);
            Assert.True(formula.IsFormula);
            Assert.Equal(15d, formula.Value);
            Assert.Equal("'Totals'!A1+5", formula.FormulaText);
            LegacyXlsDefinedName scopedName = Assert.Single(legacy.DefinedNames, name => name.Name == "TotalsLocal");
            Assert.Equal(2, scopedName.LocalSheetIndex);

            using ExcelDocument document = ExcelDocument.LoadLegacyXls(new MemoryStream(compound), new LegacyXlsImportOptions {
                ReportUnsupportedContent = true
            });

            using var output = new MemoryStream();
            document.Save(output);
            using SpreadsheetDocument spreadsheet = SpreadsheetDocument.Open(new MemoryStream(output.ToArray()), false);
            Sheet dataSheet = spreadsheet.WorkbookPart!.Workbook.Sheets!.Elements<Sheet>().Single(sheet => sheet.Name == "Data");
            WorksheetPart worksheetPart = (WorksheetPart)spreadsheet.WorkbookPart.GetPartById(dataSheet.Id!);
            Cell projectedFormula = worksheetPart.Worksheet.Descendants<Cell>().Single(cell => cell.CellReference!.Value == "A1");
            Assert.Equal("'Totals'!A1+5", projectedFormula.CellFormula!.Text);
            DefinedName projectedScopedName = spreadsheet.WorkbookPart.Workbook.DefinedNames!.Elements<DefinedName>().Single(name => name.Name == "TotalsLocal");
            Assert.Equal(2U, projectedScopedName.LocalSheetId!.Value);
            Assert.Equal("'Totals'!$A$1", projectedScopedName.Text);
        }

        [Fact]
        public void LegacyXls_Load_ProjectsFormulaReferencesToUnsupportedSheetsAsRef() {
            byte[] workbookStream = LegacyXlsTestWorkbookBuilder.CreateFormula3dReferenceToUnsupportedSheetWorkbookStream();
            byte[] compound = LegacyXlsCompoundTestBuilder.CreateWorkbookCompoundFile(workbookStream);

            LegacyXlsWorkbook legacy = LegacyXlsWorkbook.Load(compound, new LegacyXlsImportOptions {
                ReportUnsupportedContent = true
            });

            Assert.DoesNotContain(legacy.Diagnostics, d => d.Severity == LegacyXlsDiagnosticSeverity.Error);
            Assert.Empty(legacy.UnsupportedSheets);
            LegacyXlsChartSheet chartSheet = Assert.Single(legacy.ChartSheets);
            Assert.Equal("Chart1", chartSheet.Name);
            LegacyXlsCell formula = Assert.Single(legacy.Worksheets[1].Cells, cell => cell.Row == 1 && cell.Column == 1);
            Assert.True(formula.IsFormula);
            Assert.Equal("#REF!A1+5", formula.FormulaText);
            Assert.DoesNotContain("Chart1", formula.FormulaText, StringComparison.OrdinalIgnoreCase);

            using ExcelDocument document = ExcelDocument.LoadLegacyXls(new MemoryStream(compound), new LegacyXlsImportOptions {
                ReportUnsupportedContent = true
            });

            using var output = new MemoryStream();
            document.Save(output);
            using SpreadsheetDocument spreadsheet = SpreadsheetDocument.Open(new MemoryStream(output.ToArray()), false);
            Sheet formulasSheet = spreadsheet.WorkbookPart!.Workbook.Sheets!.Elements<Sheet>().Single(sheet => sheet.Name == "Formulas");
            WorksheetPart worksheetPart = (WorksheetPart)spreadsheet.WorkbookPart.GetPartById(formulasSheet.Id!);
            Cell projectedFormula = worksheetPart.Worksheet.Descendants<Cell>().Single(cell => cell.CellReference!.Value == "A1");
            Assert.Equal("#REF!A1+5", projectedFormula.CellFormula!.Text);
            Assert.DoesNotContain("Chart1", projectedFormula.CellFormula.Text, StringComparison.OrdinalIgnoreCase);
        }

        [Fact]
        public void LegacyXls_Load_ImportsFormulaFixedFunctionTokens() {
            byte[] workbookStream = LegacyXlsTestWorkbookBuilder.CreateFormulaFixedFunctionWorkbookStream();
            byte[] compound = LegacyXlsCompoundTestBuilder.CreateWorkbookCompoundFile(workbookStream);

            LegacyXlsWorkbook legacy = LegacyXlsWorkbook.Load(compound, new LegacyXlsImportOptions {
                ReportUnsupportedContent = false
            });

            Assert.DoesNotContain(legacy.Diagnostics, d => d.Severity == LegacyXlsDiagnosticSeverity.Error);
            LegacyXlsWorksheet sheet = Assert.Single(legacy.Worksheets);
            LegacyXlsCell formula = Assert.Single(sheet.Cells, cell => cell.Row == 1 && cell.Column == 2);
            Assert.True(formula.IsFormula);
            Assert.Equal(12.35d, formula.Value);
            Assert.Equal("ROUND(A1,2)", formula.FormulaText);
            LegacyXlsCell andFormula = Assert.Single(sheet.Cells, cell => cell.Row == 1 && cell.Column == 3);
            Assert.True(andFormula.IsFormula);
            Assert.Equal(true, andFormula.Value);
            Assert.Equal("AND(TRUE,TRUE)", andFormula.FormulaText);
            LegacyXlsCell orFormula = Assert.Single(sheet.Cells, cell => cell.Row == 1 && cell.Column == 4);
            Assert.True(orFormula.IsFormula);
            Assert.Equal(false, orFormula.Value);
            Assert.Equal("OR(FALSE,FALSE)", orFormula.FormulaText);
            LegacyXlsCell rsqFormula = Assert.Single(sheet.Cells, cell => cell.Row == 3 && cell.Column == 5);
            Assert.True(rsqFormula.IsFormula);
            Assert.Equal(1d, rsqFormula.Value);
            Assert.Equal("RSQ(A2:A3,B2:B3)", rsqFormula.FormulaText);
            LegacyXlsCell naFormula = Assert.Single(sheet.Cells, cell => cell.Row == 3 && cell.Column == 6);
            Assert.True(naFormula.IsFormula);
            Assert.Equal("#N/A", naFormula.Value);
            Assert.Equal("NA()", naFormula.FormulaText);

            using ExcelDocument document = ExcelDocument.LoadLegacyXls(new MemoryStream(compound), new LegacyXlsImportOptions {
                ReportUnsupportedContent = false
            });

            Assert.True(document.Sheets[0].TryGetCellText(1, 2, out string? cachedText));
            Assert.Equal("12.35", cachedText);
            Assert.True(document.Sheets[0].TryGetCellText(1, 3, out string? andCachedText));
            Assert.Equal("1", andCachedText);
            Assert.True(document.Sheets[0].TryGetCellText(1, 4, out string? orCachedText));
            Assert.Equal("0", orCachedText);

            using var output = new MemoryStream();
            document.Save(output);
            using SpreadsheetDocument spreadsheet = SpreadsheetDocument.Open(new MemoryStream(output.ToArray()), false);
            WorksheetPart worksheetPart = spreadsheet.WorkbookPart!.WorksheetParts.Single();
            Cell projectedFormula = worksheetPart.Worksheet.Descendants<Cell>().Single(cell => cell.CellReference!.Value == "B1");
            Assert.Equal("ROUND(A1,2)", projectedFormula.CellFormula!.Text);
            Assert.Equal("12.35", projectedFormula.CellValue!.Text);
            Cell projectedAndFormula = worksheetPart.Worksheet.Descendants<Cell>().Single(cell => cell.CellReference!.Value == "C1");
            Assert.Equal("AND(TRUE,TRUE)", projectedAndFormula.CellFormula!.Text);
            Assert.Equal("1", projectedAndFormula.CellValue!.Text);
            Cell projectedOrFormula = worksheetPart.Worksheet.Descendants<Cell>().Single(cell => cell.CellReference!.Value == "D1");
            Assert.Equal("OR(FALSE,FALSE)", projectedOrFormula.CellFormula!.Text);
            Assert.Equal("0", projectedOrFormula.CellValue!.Text);
            Cell projectedRsqFormula = worksheetPart.Worksheet.Descendants<Cell>().Single(cell => cell.CellReference!.Value == "E3");
            Assert.Equal("RSQ(A2:A3,B2:B3)", projectedRsqFormula.CellFormula!.Text);
            Assert.Equal("1", projectedRsqFormula.CellValue!.Text);
            Cell projectedNaFormula = worksheetPart.Worksheet.Descendants<Cell>().Single(cell => cell.CellReference!.Value == "F3");
            Assert.Equal("NA()", projectedNaFormula.CellFormula!.Text);
            Assert.Equal("#N/A", projectedNaFormula.CellValue!.Text);
        }

        [Fact]
        public void LegacyXls_Load_ImportsFormulaCommonFunctionTokens() {
            byte[] workbookStream = LegacyXlsTestWorkbookBuilder.CreateFormulaCommonFunctionWorkbookStream();
            byte[] compound = LegacyXlsCompoundTestBuilder.CreateWorkbookCompoundFile(workbookStream);

            LegacyXlsWorkbook legacy = LegacyXlsWorkbook.Load(compound, new LegacyXlsImportOptions {
                ReportUnsupportedContent = true
            });

            Assert.DoesNotContain(legacy.Diagnostics, d => d.Severity == LegacyXlsDiagnosticSeverity.Error);
            Assert.DoesNotContain(legacy.Diagnostics, d => d.Code == "XLS-BIFF-FORMULA-TOKENS-UNSUPPORTED");
            LegacyXlsWorksheet sheet = Assert.Single(legacy.Worksheets);
            AssertFormula(sheet, 2, 1, "LEFT(A1,6)", "Office");
            AssertFormula(sheet, 3, 1, "RIGHT(A1,3)", "IMO");
            AssertFormula(sheet, 4, 1, "DATE(2026,6,23)", 46196d);
            AssertFormula(sheet, 5, 1, "YEAR(B1)", 2024d);
            AssertFormula(sheet, 6, 1, "MONTH(B1)", 10d);
            AssertFormula(sheet, 7, 1, "DAY(B1)", 1d);
            AssertFormula(sheet, 8, 1, "COUNTIF(A1:A3,\"*IMO\")", 2d);
            AssertFormula(sheet, 9, 1, "LOWER(A1)", "officeimo");
            AssertFormula(sheet, 10, 1, "UPPER(A1)", "OFFICEIMO");
            AssertFormula(sheet, 11, 1, "PROPER(A1)", "Officeimo");
            AssertFormula(sheet, 12, 1, "SUMIF(D1:D3,\"EU\",E1:E3)", 40d);
            AssertFormula(sheet, 13, 1, "SUBTOTAL(9,E1:E3)", 60d);
            AssertFormula(sheet, 14, 1, "INDEX(E1:E3,2)", 20d);
            AssertFormula(sheet, 15, 1, "MATCH(\"US\",D1:D3,0)", 2d);
            AssertFormula(sheet, 16, 1, "OFFSET(E1,1,0)", 20d);
            AssertFormula(sheet, 17, 1, "TODAY()", 46196d);
            AssertFormula(sheet, 18, 1, "NOW()", 46196.5d);
            AssertFormula(sheet, 19, 1, "ROW(A1)", 1d);
            AssertFormula(sheet, 20, 1, "COLUMN(B1)", 2d);
            AssertFormula(sheet, 21, 1, "TIME(9,30,0)", 0.395833333333333d);
            AssertFormula(sheet, 22, 1, "HOUR(0.5)", 12d);
            AssertFormula(sheet, 23, 1, "MINUTE(0.5)", 0d);
            AssertFormula(sheet, 24, 1, "SECOND(0.5)", 0d);
            AssertFormula(sheet, 25, 1, "RAND()", 0.42d);
            AssertFormula(sheet, 26, 1, "ROWS(E1:E3)", 3d);
            AssertFormula(sheet, 27, 1, "COLUMNS(D1:E1)", 2d);
            AssertFormula(sheet, 28, 1, "LARGE(E1:E3,2)", 20d);
            AssertFormula(sheet, 29, 1, "STDEV(E1:E3)", 10d);
            AssertFormula(sheet, 30, 1, "NPV(10,E1:E3)", 100d);

            LegacyXlsImportReport report = legacy.CreateImportReport();
            Assert.True(report.FormulaTokenRecordCount > 0);
            Assert.Contains("PtgFunc", report.FormulaTokensByName.Keys);
            Assert.Contains("PtgFuncVar", report.FormulaTokensByName.Keys);
            Assert.Contains("Reference", report.FormulaTokensByClass.Keys);
            Assert.Contains("PtgFunc|Reference", report.FormulaTokensByNameAndClass.Keys);
            Assert.Contains("PtgFuncVar|Reference", report.FormulaTokensByNameAndClass.Keys);
            Assert.Contains("PtgFunc|Bytes:2", report.FormulaTokensByOperandByteCount.Keys);
            Assert.Contains("PtgFuncVar|Bytes:3", report.FormulaTokensByOperandByteCount.Keys);
            Assert.Contains("CellReference", report.FormulaTokensByOperandKind.Keys);
            Assert.Contains("AreaReference", report.FormulaTokensByOperandKind.Keys);
            Assert.Contains("FixedFunction", report.FormulaTokensByOperandKind.Keys);
            Assert.Contains("VariableFunction", report.FormulaTokensByOperandKind.Keys);
            Assert.Contains("StringLiteral", report.FormulaTokensByOperandKind.Keys);
            Assert.Contains("PtgFunc|FixedFunction", report.FormulaTokensByNameAndOperandKind.Keys);
            Assert.Contains("PtgFuncVar|VariableFunction", report.FormulaTokensByNameAndOperandKind.Keys);
            Assert.Contains("FixedFunction|COUNTIF", report.FormulaTokensByOperandKindAndText.Keys);
            Assert.Contains("CellReference|A1", report.FormulaTokensByOperandKindAndText.Keys);
            Assert.Contains("AreaReference|A1:A3", report.FormulaTokensByOperandKindAndText.Keys);
            Assert.Contains("PtgFunc|COUNTIF", report.FormulaTokensByNameAndOperandText.Keys);
            Assert.Contains("PtgRef|A1", report.FormulaTokensByNameAndOperandText.Keys);
            Assert.Contains("PtgArea|A1:A3", report.FormulaTokensByNameAndOperandText.Keys);
            Assert.True(report.FormulaTokensByContext["CellFormula"] >= 30);
            Assert.Equal(1, report.FormulaFunctionsByName["COUNTIF"]);
            Assert.Equal(1, report.FormulaFunctionsByName["SUMIF"]);
            Assert.Equal(1, report.FormulaFunctionsByName["SUBTOTAL"]);
            Assert.Equal(1, report.FormulaFunctionsByName["LARGE"]);
            Assert.Equal(1, report.FormulaFunctionsByName["STDEV"]);
            Assert.Equal(1, report.FormulaFunctionsByName["NPV"]);
            Assert.Contains(legacy.FormulaTokenRecords, record =>
                record.Context == "CellFormula"
                && record.SheetName == "CommonFunc"
                && record.CellReference == "A8"
                && record.TokenName == "PtgFunc"
                && record.FunctionName == "COUNTIF"
                && record.TokenClassName == "Reference"
                && record.OperandByteCount == 2
                && record.OperandKind == "FixedFunction"
                && record.OperandText == "COUNTIF"
                && record.FunctionParameterCount == 2);
            Assert.Contains(legacy.FormulaTokenRecords, record =>
                record.Context == "CellFormula"
                && record.SheetName == "CommonFunc"
                && record.OperandKind == "CellReference"
                && record.OperandText == "A1");
            Assert.Contains(legacy.FormulaTokenRecords, record =>
                record.Context == "CellFormula"
                && record.SheetName == "CommonFunc"
                && record.OperandKind == "AreaReference"
                && record.OperandText == "A1:A3");

            using ExcelDocument document = ExcelDocument.LoadLegacyXls(new MemoryStream(compound), new LegacyXlsImportOptions {
                ReportUnsupportedContent = true
            });

            using var output = new MemoryStream();
            document.Save(output);
            using SpreadsheetDocument spreadsheet = SpreadsheetDocument.Open(new MemoryStream(output.ToArray()), false);
            WorksheetPart worksheetPart = spreadsheet.WorkbookPart!.WorksheetParts.Single();
            AssertProjectedFormula(worksheetPart, "A2", "LEFT(A1,6)", "Office");
            AssertProjectedFormula(worksheetPart, "A3", "RIGHT(A1,3)", "IMO");
            AssertProjectedFormula(worksheetPart, "A4", "DATE(2026,6,23)", "46196");
            AssertProjectedFormula(worksheetPart, "A5", "YEAR(B1)", "2024");
            AssertProjectedFormula(worksheetPart, "A6", "MONTH(B1)", "10");
            AssertProjectedFormula(worksheetPart, "A7", "DAY(B1)", "1");
            AssertProjectedFormula(worksheetPart, "A8", "COUNTIF(A1:A3,\"*IMO\")", "2");
            AssertProjectedFormula(worksheetPart, "A9", "LOWER(A1)", "officeimo");
            AssertProjectedFormula(worksheetPart, "A10", "UPPER(A1)", "OFFICEIMO");
            AssertProjectedFormula(worksheetPart, "A11", "PROPER(A1)", "Officeimo");
            AssertProjectedFormula(worksheetPart, "A12", "SUMIF(D1:D3,\"EU\",E1:E3)", "40");
            AssertProjectedFormula(worksheetPart, "A13", "SUBTOTAL(9,E1:E3)", "60");
            AssertProjectedFormula(worksheetPart, "A14", "INDEX(E1:E3,2)", "20");
            AssertProjectedFormula(worksheetPart, "A15", "MATCH(\"US\",D1:D3,0)", "2");
            AssertProjectedFormula(worksheetPart, "A16", "OFFSET(E1,1,0)", "20");
            AssertProjectedFormula(worksheetPart, "A17", "TODAY()", "46196");
            AssertProjectedFormula(worksheetPart, "A18", "NOW()", "46196.5");
            AssertProjectedFormula(worksheetPart, "A19", "ROW(A1)", "1");
            AssertProjectedFormula(worksheetPart, "A20", "COLUMN(B1)", "2");
            AssertProjectedFormula(worksheetPart, "A21", "TIME(9,30,0)", "0.395833333333333");
            AssertProjectedFormula(worksheetPart, "A22", "HOUR(0.5)", "12");
            AssertProjectedFormula(worksheetPart, "A23", "MINUTE(0.5)", "0");
            AssertProjectedFormula(worksheetPart, "A24", "SECOND(0.5)", "0");
            AssertProjectedFormula(worksheetPart, "A25", "RAND()", "0.42");
            AssertProjectedFormula(worksheetPart, "A26", "ROWS(E1:E3)", "3");
            AssertProjectedFormula(worksheetPart, "A27", "COLUMNS(D1:E1)", "2");
            AssertProjectedFormula(worksheetPart, "A28", "LARGE(E1:E3,2)", "20");
            AssertProjectedFormula(worksheetPart, "A29", "STDEV(E1:E3)", "10");
            AssertProjectedFormula(worksheetPart, "A30", "NPV(10,E1:E3)", "100");
        }

        [Fact]
        public void LegacyXls_Load_ImportsFormulaLookupFunctionTokens() {
            byte[] workbookStream = LegacyXlsTestWorkbookBuilder.CreateFormulaLookupFunctionWorkbookStream();
            byte[] compound = LegacyXlsCompoundTestBuilder.CreateWorkbookCompoundFile(workbookStream);

            LegacyXlsWorkbook legacy = LegacyXlsWorkbook.Load(compound, new LegacyXlsImportOptions {
                ReportUnsupportedContent = false
            });

            Assert.DoesNotContain(legacy.Diagnostics, d => d.Severity == LegacyXlsDiagnosticSeverity.Error);
            LegacyXlsWorksheet sheet = Assert.Single(legacy.Worksheets);
            LegacyXlsCell formula = Assert.Single(sheet.Cells, cell => cell.Row == 3 && cell.Column == 4);
            Assert.True(formula.IsFormula);
            Assert.Equal(200d, formula.Value);
            Assert.Equal("VLOOKUP(A1,B1:C2,2,FALSE)", formula.FormulaText);

            using ExcelDocument document = ExcelDocument.LoadLegacyXls(new MemoryStream(compound), new LegacyXlsImportOptions {
                ReportUnsupportedContent = false
            });

            Assert.True(document.Sheets[0].TryGetCellText(3, 4, out string? cachedText));
            Assert.Equal("200", cachedText);

            using var output = new MemoryStream();
            document.Save(output);
            using SpreadsheetDocument spreadsheet = SpreadsheetDocument.Open(new MemoryStream(output.ToArray()), false);
            WorksheetPart worksheetPart = spreadsheet.WorkbookPart!.WorksheetParts.Single();
            Cell projectedFormula = worksheetPart.Worksheet.Descendants<Cell>().Single(cell => cell.CellReference!.Value == "D3");
            Assert.Equal("VLOOKUP(A1,B1:C2,2,FALSE)", projectedFormula.CellFormula!.Text);
            Assert.Equal("200", projectedFormula.CellValue!.Text);
        }

        [Fact]
        public void LegacyXls_Load_ImportsFormulaAttributeSumTokens() {
            byte[] workbookStream = LegacyXlsTestWorkbookBuilder.CreateFormulaAttributeSumWorkbookStream();
            byte[] compound = LegacyXlsCompoundTestBuilder.CreateWorkbookCompoundFile(workbookStream);

            LegacyXlsWorkbook legacy = LegacyXlsWorkbook.Load(compound, new LegacyXlsImportOptions {
                ReportUnsupportedContent = true
            });

            Assert.DoesNotContain(legacy.Diagnostics, d => d.Severity == LegacyXlsDiagnosticSeverity.Error);
            Assert.DoesNotContain(legacy.Diagnostics, d => d.Code == "XLS-BIFF-FORMULA-TOKENS-UNSUPPORTED");
            LegacyXlsWorksheet sheet = Assert.Single(legacy.Worksheets);
            LegacyXlsCell formula = Assert.Single(sheet.Cells, cell => cell.Row == 1 && cell.Column == 3);
            Assert.True(formula.IsFormula);
            Assert.Equal(30d, formula.Value);
            Assert.Equal("SUM(A1:B1)", formula.FormulaText);

            using ExcelDocument document = ExcelDocument.LoadLegacyXls(new MemoryStream(compound), new LegacyXlsImportOptions {
                ReportUnsupportedContent = true
            });

            Assert.True(document.Sheets[0].TryGetCellText(1, 3, out string? cachedText));
            Assert.Equal("30", cachedText);

            using var output = new MemoryStream();
            document.Save(output);
            using SpreadsheetDocument spreadsheet = SpreadsheetDocument.Open(new MemoryStream(output.ToArray()), false);
            WorksheetPart worksheetPart = spreadsheet.WorkbookPart!.WorksheetParts.Single();
            Cell projectedFormula = worksheetPart.Worksheet.Descendants<Cell>().Single(cell => cell.CellReference!.Value == "C1");
            Assert.Equal("SUM(A1:B1)", projectedFormula.CellFormula!.Text);
            Assert.Equal("30", projectedFormula.CellValue!.Text);
        }

        [Fact]
        public void LegacyXls_Load_ImportsFormulaErrorConstantTokens() {
            byte[] workbookStream = LegacyXlsTestWorkbookBuilder.CreateFormulaErrorConstantWorkbookStream();
            byte[] compound = LegacyXlsCompoundTestBuilder.CreateWorkbookCompoundFile(workbookStream);

            LegacyXlsWorkbook legacy = LegacyXlsWorkbook.Load(compound, new LegacyXlsImportOptions {
                ReportUnsupportedContent = true
            });

            Assert.DoesNotContain(legacy.Diagnostics, d => d.Severity == LegacyXlsDiagnosticSeverity.Error);
            Assert.DoesNotContain(legacy.Diagnostics, d => d.Code == "XLS-BIFF-FORMULA-TOKENS-UNSUPPORTED");
            LegacyXlsWorksheet sheet = Assert.Single(legacy.Worksheets);
            LegacyXlsCell formula = Assert.Single(sheet.Cells, cell => cell.Row == 1 && cell.Column == 1);
            Assert.True(formula.IsFormula);
            Assert.Equal("#VALUE!", formula.Value);
            Assert.Equal("#VALUE!", formula.FormulaText);

            using ExcelDocument document = ExcelDocument.LoadLegacyXls(new MemoryStream(compound), new LegacyXlsImportOptions {
                ReportUnsupportedContent = true
            });

            Assert.True(document.Sheets[0].TryGetCellText(1, 1, out string? cachedText));
            Assert.Equal("#VALUE!", cachedText);

            using var output = new MemoryStream();
            document.Save(output);
            using SpreadsheetDocument spreadsheet = SpreadsheetDocument.Open(new MemoryStream(output.ToArray()), false);
            WorksheetPart worksheetPart = spreadsheet.WorkbookPart!.WorksheetParts.Single();
            Cell projectedFormula = worksheetPart.Worksheet.Descendants<Cell>().Single(cell => cell.CellReference!.Value == "A1");
            Assert.Equal("#VALUE!", projectedFormula.CellFormula!.Text);
            Assert.Equal("#VALUE!", projectedFormula.CellValue!.Text);
        }

        [Fact]
        public void LegacyXls_Load_ImportsFormulaMissingArgumentTokens() {
            byte[] workbookStream = LegacyXlsTestWorkbookBuilder.CreateFormulaMissingArgumentWorkbookStream();
            byte[] compound = LegacyXlsCompoundTestBuilder.CreateWorkbookCompoundFile(workbookStream);

            LegacyXlsWorkbook legacy = LegacyXlsWorkbook.Load(compound, new LegacyXlsImportOptions {
                ReportUnsupportedContent = true
            });

            Assert.DoesNotContain(legacy.Diagnostics, d => d.Severity == LegacyXlsDiagnosticSeverity.Error);
            Assert.DoesNotContain(legacy.Diagnostics, d => d.Code == "XLS-BIFF-FORMULA-TOKENS-UNSUPPORTED");
            LegacyXlsWorksheet sheet = Assert.Single(legacy.Worksheets);
            LegacyXlsCell formula = Assert.Single(sheet.Cells, cell => cell.Row == 1 && cell.Column == 2);
            Assert.True(formula.IsFormula);
            Assert.Equal(10d, formula.Value);
            Assert.Equal("SUM(,A1)", formula.FormulaText);

            using ExcelDocument document = ExcelDocument.LoadLegacyXls(new MemoryStream(compound), new LegacyXlsImportOptions {
                ReportUnsupportedContent = true
            });

            Assert.True(document.Sheets[0].TryGetCellText(1, 2, out string? cachedText));
            Assert.Equal("10", cachedText);

            using var output = new MemoryStream();
            document.Save(output);
            using SpreadsheetDocument spreadsheet = SpreadsheetDocument.Open(new MemoryStream(output.ToArray()), false);
            WorksheetPart worksheetPart = spreadsheet.WorkbookPart!.WorksheetParts.Single();
            Cell projectedFormula = worksheetPart.Worksheet.Descendants<Cell>().Single(cell => cell.CellReference!.Value == "B1");
            Assert.Equal("SUM(,A1)", projectedFormula.CellFormula!.Text);
            Assert.Equal("10", projectedFormula.CellValue!.Text);
        }

        [Fact]
        public void LegacyXls_Load_ImportsFormulaDisplaySpaceTokens() {
            byte[] workbookStream = LegacyXlsTestWorkbookBuilder.CreateFormulaDisplaySpaceWorkbookStream();
            byte[] compound = LegacyXlsCompoundTestBuilder.CreateWorkbookCompoundFile(workbookStream);

            LegacyXlsWorkbook legacy = LegacyXlsWorkbook.Load(compound, new LegacyXlsImportOptions {
                ReportUnsupportedContent = true
            });

            Assert.DoesNotContain(legacy.Diagnostics, d => d.Severity == LegacyXlsDiagnosticSeverity.Error);
            Assert.DoesNotContain(legacy.Diagnostics, d => d.Code == "XLS-BIFF-FORMULA-TOKENS-UNSUPPORTED");
            LegacyXlsWorksheet sheet = Assert.Single(legacy.Worksheets);
            LegacyXlsCell spaceFormula = Assert.Single(sheet.Cells, cell => cell.Row == 1 && cell.Column == 2);
            Assert.True(spaceFormula.IsFormula);
            Assert.Equal(15d, spaceFormula.Value);
            Assert.Equal("A1+5", spaceFormula.FormulaText);
            LegacyXlsCell volatileSpaceFormula = Assert.Single(sheet.Cells, cell => cell.Row == 1 && cell.Column == 3);
            Assert.True(volatileSpaceFormula.IsFormula);
            Assert.Equal(20d, volatileSpaceFormula.Value);
            Assert.Equal("A1+10", volatileSpaceFormula.FormulaText);

            using ExcelDocument document = ExcelDocument.LoadLegacyXls(new MemoryStream(compound), new LegacyXlsImportOptions {
                ReportUnsupportedContent = true
            });

            Assert.True(document.Sheets[0].TryGetCellText(1, 2, out string? firstCachedText));
            Assert.Equal("15", firstCachedText);
            Assert.True(document.Sheets[0].TryGetCellText(1, 3, out string? secondCachedText));
            Assert.Equal("20", secondCachedText);

            using var output = new MemoryStream();
            document.Save(output);
            using SpreadsheetDocument spreadsheet = SpreadsheetDocument.Open(new MemoryStream(output.ToArray()), false);
            WorksheetPart worksheetPart = spreadsheet.WorkbookPart!.WorksheetParts.Single();
            Cell projectedSpaceFormula = worksheetPart.Worksheet.Descendants<Cell>().Single(cell => cell.CellReference!.Value == "B1");
            Assert.Equal("A1+5", projectedSpaceFormula.CellFormula!.Text);
            Assert.Equal("15", projectedSpaceFormula.CellValue!.Text);
            Cell projectedVolatileSpaceFormula = worksheetPart.Worksheet.Descendants<Cell>().Single(cell => cell.CellReference!.Value == "C1");
            Assert.Equal("A1+10", projectedVolatileSpaceFormula.CellFormula!.Text);
            Assert.Equal("20", projectedVolatileSpaceFormula.CellValue!.Text);
        }

        [Fact]
        public void LegacyXls_Load_ImportsFormulaIfTokens() {
            byte[] workbookStream = LegacyXlsTestWorkbookBuilder.CreateFormulaIfWorkbookStream();
            byte[] compound = LegacyXlsCompoundTestBuilder.CreateWorkbookCompoundFile(workbookStream);

            LegacyXlsWorkbook legacy = LegacyXlsWorkbook.Load(compound, new LegacyXlsImportOptions {
                ReportUnsupportedContent = true
            });

            Assert.DoesNotContain(legacy.Diagnostics, d => d.Severity == LegacyXlsDiagnosticSeverity.Error);
            Assert.DoesNotContain(legacy.Diagnostics, d => d.Code == "XLS-BIFF-FORMULA-TOKENS-UNSUPPORTED");
            LegacyXlsWorksheet sheet = Assert.Single(legacy.Worksheets);
            LegacyXlsCell formula = Assert.Single(sheet.Cells, cell => cell.Row == 1 && cell.Column == 2);
            Assert.True(formula.IsFormula);
            Assert.Equal(1d, formula.Value);
            Assert.Equal("IF(A1>5,1,0)", formula.FormulaText);

            using ExcelDocument document = ExcelDocument.LoadLegacyXls(new MemoryStream(compound), new LegacyXlsImportOptions {
                ReportUnsupportedContent = true
            });

            Assert.True(document.Sheets[0].TryGetCellText(1, 2, out string? cachedText));
            Assert.Equal("1", cachedText);

            using var output = new MemoryStream();
            document.Save(output);
            using SpreadsheetDocument spreadsheet = SpreadsheetDocument.Open(new MemoryStream(output.ToArray()), false);
            WorksheetPart worksheetPart = spreadsheet.WorkbookPart!.WorksheetParts.Single();
            Cell projectedFormula = worksheetPart.Worksheet.Descendants<Cell>().Single(cell => cell.CellReference!.Value == "B1");
            Assert.Equal("IF(A1>5,1,0)", projectedFormula.CellFormula!.Text);
            Assert.Equal("1", projectedFormula.CellValue!.Text);
        }

        [Fact]
        public void LegacyXls_Load_ImportsFormulaChooseTokens() {
            byte[] workbookStream = LegacyXlsTestWorkbookBuilder.CreateFormulaChooseWorkbookStream();
            byte[] compound = LegacyXlsCompoundTestBuilder.CreateWorkbookCompoundFile(workbookStream);

            LegacyXlsWorkbook legacy = LegacyXlsWorkbook.Load(compound, new LegacyXlsImportOptions {
                ReportUnsupportedContent = true
            });

            Assert.DoesNotContain(legacy.Diagnostics, d => d.Severity == LegacyXlsDiagnosticSeverity.Error);
            Assert.DoesNotContain(legacy.Diagnostics, d => d.Code == "XLS-BIFF-FORMULA-TOKENS-UNSUPPORTED");
            LegacyXlsWorksheet sheet = Assert.Single(legacy.Worksheets);
            LegacyXlsCell formula = Assert.Single(sheet.Cells, cell => cell.Row == 1 && cell.Column == 2);
            Assert.True(formula.IsFormula);
            Assert.Equal(20d, formula.Value);
            Assert.Equal("CHOOSE(A1,10,20,30)", formula.FormulaText);

            using ExcelDocument document = ExcelDocument.LoadLegacyXls(new MemoryStream(compound), new LegacyXlsImportOptions {
                ReportUnsupportedContent = true
            });

            Assert.True(document.Sheets[0].TryGetCellText(1, 2, out string? cachedText));
            Assert.Equal("20", cachedText);

            using var output = new MemoryStream();
            document.Save(output);
            using SpreadsheetDocument spreadsheet = SpreadsheetDocument.Open(new MemoryStream(output.ToArray()), false);
            WorksheetPart worksheetPart = spreadsheet.WorkbookPart!.WorksheetParts.Single();
            Cell projectedFormula = worksheetPart.Worksheet.Descendants<Cell>().Single(cell => cell.CellReference!.Value == "B1");
            Assert.Equal("CHOOSE(A1,10,20,30)", projectedFormula.CellFormula!.Text);
            Assert.Equal("20", projectedFormula.CellValue!.Text);
        }

        [Fact]
        public void LegacyXls_Load_ImportsFormulaReferenceOperatorTokens() {
            byte[] workbookStream = LegacyXlsTestWorkbookBuilder.CreateFormulaReferenceOperatorWorkbookStream();
            byte[] compound = LegacyXlsCompoundTestBuilder.CreateWorkbookCompoundFile(workbookStream);

            LegacyXlsWorkbook legacy = LegacyXlsWorkbook.Load(compound, new LegacyXlsImportOptions {
                ReportUnsupportedContent = true
            });

            Assert.DoesNotContain(legacy.Diagnostics, d => d.Severity == LegacyXlsDiagnosticSeverity.Error);
            Assert.DoesNotContain(legacy.Diagnostics, d => d.Code == "XLS-BIFF-FORMULA-TOKENS-UNSUPPORTED");
            LegacyXlsWorksheet sheet = Assert.Single(legacy.Worksheets);
            LegacyXlsCell rangeFormula = Assert.Single(sheet.Cells, cell => cell.Row == 1 && cell.Column == 4);
            Assert.True(rangeFormula.IsFormula);
            Assert.Equal(30d, rangeFormula.Value);
            Assert.Equal("SUM(A1:B1)", rangeFormula.FormulaText);
            LegacyXlsCell unionFormula = Assert.Single(sheet.Cells, cell => cell.Row == 1 && cell.Column == 5);
            Assert.True(unionFormula.IsFormula);
            Assert.Equal(30d, unionFormula.Value);
            Assert.Equal("SUM((A1,B1))", unionFormula.FormulaText);
            LegacyXlsCell intersectionFormula = Assert.Single(sheet.Cells, cell => cell.Row == 1 && cell.Column == 6);
            Assert.True(intersectionFormula.IsFormula);
            Assert.Equal(20d, intersectionFormula.Value);
            Assert.Equal("SUM((A1:C1 B1))", intersectionFormula.FormulaText);

            using ExcelDocument document = ExcelDocument.LoadLegacyXls(new MemoryStream(compound), new LegacyXlsImportOptions {
                ReportUnsupportedContent = true
            });

            Assert.True(document.Sheets[0].TryGetCellText(1, 4, out string? rangeCachedText));
            Assert.Equal("30", rangeCachedText);
            Assert.True(document.Sheets[0].TryGetCellText(1, 5, out string? unionCachedText));
            Assert.Equal("30", unionCachedText);
            Assert.True(document.Sheets[0].TryGetCellText(1, 6, out string? intersectionCachedText));
            Assert.Equal("20", intersectionCachedText);

            using var output = new MemoryStream();
            document.Save(output);
            using SpreadsheetDocument spreadsheet = SpreadsheetDocument.Open(new MemoryStream(output.ToArray()), false);
            WorksheetPart worksheetPart = spreadsheet.WorkbookPart!.WorksheetParts.Single();
            Cell projectedRangeFormula = worksheetPart.Worksheet.Descendants<Cell>().Single(cell => cell.CellReference!.Value == "D1");
            Assert.Equal("SUM(A1:B1)", projectedRangeFormula.CellFormula!.Text);
            Assert.Equal("30", projectedRangeFormula.CellValue!.Text);
            Cell projectedUnionFormula = worksheetPart.Worksheet.Descendants<Cell>().Single(cell => cell.CellReference!.Value == "E1");
            Assert.Equal("SUM((A1,B1))", projectedUnionFormula.CellFormula!.Text);
            Assert.Equal("30", projectedUnionFormula.CellValue!.Text);
            Cell projectedIntersectionFormula = worksheetPart.Worksheet.Descendants<Cell>().Single(cell => cell.CellReference!.Value == "F1");
            Assert.Equal("SUM((A1:C1 B1))", projectedIntersectionFormula.CellFormula!.Text);
            Assert.Equal("20", projectedIntersectionFormula.CellValue!.Text);
        }

        [Fact]
        public void LegacyXls_Load_ImportsFormulaMemAreaTokens() {
            byte[] workbookStream = LegacyXlsTestWorkbookBuilder.CreateFormulaMemAreaWorkbookStream();
            byte[] compound = LegacyXlsCompoundTestBuilder.CreateWorkbookCompoundFile(workbookStream);

            LegacyXlsWorkbook legacy = LegacyXlsWorkbook.Load(compound, new LegacyXlsImportOptions {
                ReportUnsupportedContent = true
            });

            Assert.DoesNotContain(legacy.Diagnostics, d => d.Severity == LegacyXlsDiagnosticSeverity.Error);
            Assert.DoesNotContain(legacy.Diagnostics, d => d.Code == "XLS-BIFF-FORMULA-TOKENS-UNSUPPORTED");
            LegacyXlsWorksheet sheet = Assert.Single(legacy.Worksheets);
            LegacyXlsCell formula = Assert.Single(sheet.Cells, cell => cell.Row == 3 && cell.Column == 1);
            Assert.True(formula.IsFormula);
            Assert.Equal(30d, formula.Value);
            Assert.Equal("SUM(A1:A2)", formula.FormulaText);

            using ExcelDocument document = ExcelDocument.LoadLegacyXls(new MemoryStream(compound), new LegacyXlsImportOptions {
                ReportUnsupportedContent = true
            });

            Assert.True(document.Sheets[0].TryGetCellText(3, 1, out string? cachedText));
            Assert.Equal("30", cachedText);

            using var output = new MemoryStream();
            document.Save(output);
            using SpreadsheetDocument spreadsheet = SpreadsheetDocument.Open(new MemoryStream(output.ToArray()), false);
            WorksheetPart worksheetPart = spreadsheet.WorkbookPart!.WorksheetParts.Single();
            Cell projectedFormula = worksheetPart.Worksheet.Descendants<Cell>().Single(cell => cell.CellReference!.Value == "A3");
            Assert.Equal("SUM(A1:A2)", projectedFormula.CellFormula!.Text);
            Assert.Equal("30", projectedFormula.CellValue!.Text);
        }

        [Fact]
        public void LegacyXls_Load_ImportsFormulaMemFuncTokens() {
            byte[] workbookStream = LegacyXlsTestWorkbookBuilder.CreateFormulaMemFuncWorkbookStream();
            byte[] compound = LegacyXlsCompoundTestBuilder.CreateWorkbookCompoundFile(workbookStream);

            LegacyXlsWorkbook legacy = LegacyXlsWorkbook.Load(compound, new LegacyXlsImportOptions {
                ReportUnsupportedContent = true
            });

            Assert.DoesNotContain(legacy.Diagnostics, d => d.Severity == LegacyXlsDiagnosticSeverity.Error);
            Assert.DoesNotContain(legacy.Diagnostics, d => d.Code == "XLS-BIFF-FORMULA-TOKENS-UNSUPPORTED");
            LegacyXlsWorksheet sheet = Assert.Single(legacy.Worksheets);
            LegacyXlsCell formula = Assert.Single(sheet.Cells, cell => cell.Row == 1 && cell.Column == 3);
            Assert.True(formula.IsFormula);
            Assert.Equal(30d, formula.Value);
            Assert.Equal("SUM((A1,B1))", formula.FormulaText);

            using ExcelDocument document = ExcelDocument.LoadLegacyXls(new MemoryStream(compound), new LegacyXlsImportOptions {
                ReportUnsupportedContent = true
            });

            Assert.True(document.Sheets[0].TryGetCellText(1, 3, out string? cachedText));
            Assert.Equal("30", cachedText);

            using var output = new MemoryStream();
            document.Save(output);
            using SpreadsheetDocument spreadsheet = SpreadsheetDocument.Open(new MemoryStream(output.ToArray()), false);
            WorksheetPart worksheetPart = spreadsheet.WorkbookPart!.WorksheetParts.Single();
            Cell projectedFormula = worksheetPart.Worksheet.Descendants<Cell>().Single(cell => cell.CellReference!.Value == "C1");
            Assert.Equal("SUM((A1,B1))", projectedFormula.CellFormula!.Text);
            Assert.Equal("30", projectedFormula.CellValue!.Text);
        }

        [Fact]
        public void LegacyXls_Load_ImportsFormulaDefinedNameTokens() {
            byte[] workbookStream = LegacyXlsTestWorkbookBuilder.CreateFormulaDefinedNameWorkbookStream();
            byte[] compound = LegacyXlsCompoundTestBuilder.CreateWorkbookCompoundFile(workbookStream);

            LegacyXlsWorkbook legacy = LegacyXlsWorkbook.Load(compound, new LegacyXlsImportOptions {
                ReportUnsupportedContent = true
            });

            Assert.DoesNotContain(legacy.Diagnostics, d => d.Severity == LegacyXlsDiagnosticSeverity.Error);
            Assert.DoesNotContain(legacy.Diagnostics, d => d.Code == "XLS-BIFF-FORMULA-TOKENS-UNSUPPORTED");
            Assert.Contains(legacy.DefinedNames, name => name.Name == "TaxRate" && name.Reference == "'NameFormula'!$C$1");
            LegacyXlsWorksheet sheet = Assert.Single(legacy.Worksheets);
            LegacyXlsCell formula = Assert.Single(sheet.Cells, cell => cell.Row == 1 && cell.Column == 2);
            Assert.True(formula.IsFormula);
            Assert.Equal(2.5d, formula.Value);
            Assert.Equal("A1*TaxRate", formula.FormulaText);

            using ExcelDocument document = ExcelDocument.LoadLegacyXls(new MemoryStream(compound), new LegacyXlsImportOptions {
                ReportUnsupportedContent = true
            });

            Assert.Equal("'NameFormula'!$C$1", document.GetNamedRange("TaxRate"));
            Assert.True(document.Sheets[0].TryGetCellText(1, 2, out string? cachedText));
            Assert.Equal("2.5", cachedText);

            using var output = new MemoryStream();
            document.Save(output);
            using SpreadsheetDocument spreadsheet = SpreadsheetDocument.Open(new MemoryStream(output.ToArray()), false);
            DefinedName projectedName = spreadsheet.WorkbookPart!.Workbook.DefinedNames!.Elements<DefinedName>().Single(name => name.Name == "TaxRate");
            Assert.Equal("'NameFormula'!$C$1", projectedName.Text);
            WorksheetPart worksheetPart = spreadsheet.WorkbookPart.WorksheetParts.Single();
            Cell projectedFormula = worksheetPart.Worksheet.Descendants<Cell>().Single(cell => cell.CellReference!.Value == "B1");
            Assert.Equal("A1*TaxRate", projectedFormula.CellFormula!.Text);
            Assert.Equal("2.5", projectedFormula.CellValue!.Text);
        }

        [Fact]
        public void LegacyXls_Load_ImportsFormulaInvalidReferenceTokens() {
            byte[] workbookStream = LegacyXlsTestWorkbookBuilder.CreateFormulaInvalidReferenceWorkbookStream();
            byte[] compound = LegacyXlsCompoundTestBuilder.CreateWorkbookCompoundFile(workbookStream);

            LegacyXlsWorkbook legacy = LegacyXlsWorkbook.Load(compound, new LegacyXlsImportOptions {
                ReportUnsupportedContent = true
            });

            Assert.DoesNotContain(legacy.Diagnostics, d => d.Severity == LegacyXlsDiagnosticSeverity.Error);
            Assert.DoesNotContain(legacy.Diagnostics, d => d.Code == "XLS-BIFF-FORMULA-TOKENS-UNSUPPORTED");
            Assert.Equal(2, legacy.Worksheets.Count);
            LegacyXlsWorksheet formulas = legacy.Worksheets[1];
            LegacyXlsCell sameSheetReference = Assert.Single(formulas.Cells, cell => cell.Row == 1 && cell.Column == 1);
            Assert.True(sameSheetReference.IsFormula);
            Assert.Equal("#REF!", sameSheetReference.Value);
            Assert.Equal("SUM(#REF!)", sameSheetReference.FormulaText);
            LegacyXlsCell sameSheetArea = Assert.Single(formulas.Cells, cell => cell.Row == 1 && cell.Column == 2);
            Assert.True(sameSheetArea.IsFormula);
            Assert.Equal("#REF!", sameSheetArea.Value);
            Assert.Equal("SUM(#REF!)", sameSheetArea.FormulaText);
            LegacyXlsCell invalid3dReference = Assert.Single(formulas.Cells, cell => cell.Row == 1 && cell.Column == 3);
            Assert.True(invalid3dReference.IsFormula);
            Assert.Equal("#REF!", invalid3dReference.Value);
            Assert.Equal("SUM('Source Data'!#REF!)", invalid3dReference.FormulaText);
            LegacyXlsCell invalid3dArea = Assert.Single(formulas.Cells, cell => cell.Row == 1 && cell.Column == 4);
            Assert.True(invalid3dArea.IsFormula);
            Assert.Equal("#REF!", invalid3dArea.Value);
            Assert.Equal("SUM('Source Data'!#REF!)", invalid3dArea.FormulaText);

            using ExcelDocument document = ExcelDocument.LoadLegacyXls(new MemoryStream(compound), new LegacyXlsImportOptions {
                ReportUnsupportedContent = true
            });

            Assert.True(document.Sheets[1].TryGetCellText(1, 1, out string? firstCachedText));
            Assert.Equal("#REF!", firstCachedText);
            Assert.True(document.Sheets[1].TryGetCellText(1, 4, out string? fourthCachedText));
            Assert.Equal("#REF!", fourthCachedText);

            using var output = new MemoryStream();
            document.Save(output);
            using SpreadsheetDocument spreadsheet = SpreadsheetDocument.Open(new MemoryStream(output.ToArray()), false);
            WorkbookPart workbookPart = spreadsheet.WorkbookPart!;
            Sheet formulaSheet = workbookPart.Workbook.Sheets!.Elements<Sheet>().Single(sheet => sheet.Name == "FormulaRefs");
            WorksheetPart worksheetPart = (WorksheetPart)workbookPart.GetPartById(formulaSheet.Id!);
            Dictionary<string, Cell> cells = worksheetPart.Worksheet.Descendants<Cell>().ToDictionary(cell => cell.CellReference!.Value!);
            Assert.Equal("SUM(#REF!)", cells["A1"].CellFormula!.Text);
            Assert.Equal("#REF!", cells["A1"].CellValue!.Text);
            Assert.Equal("SUM(#REF!)", cells["B1"].CellFormula!.Text);
            Assert.Equal("#REF!", cells["B1"].CellValue!.Text);
            Assert.Equal("SUM('Source Data'!#REF!)", cells["C1"].CellFormula!.Text);
            Assert.Equal("#REF!", cells["C1"].CellValue!.Text);
            Assert.Equal("SUM('Source Data'!#REF!)", cells["D1"].CellFormula!.Text);
            Assert.Equal("#REF!", cells["D1"].CellValue!.Text);
        }

        [Fact]
        public void LegacyXls_Load_ImportsFormulaRelativeReferenceTokens() {
            byte[] workbookStream = LegacyXlsTestWorkbookBuilder.CreateFormulaRelativeReferenceWorkbookStream();
            byte[] compound = LegacyXlsCompoundTestBuilder.CreateWorkbookCompoundFile(workbookStream);

            LegacyXlsWorkbook legacy = LegacyXlsWorkbook.Load(compound, new LegacyXlsImportOptions {
                ReportUnsupportedContent = true
            });

            Assert.DoesNotContain(legacy.Diagnostics, d => d.Severity == LegacyXlsDiagnosticSeverity.Error);
            Assert.DoesNotContain(legacy.Diagnostics, d => d.Code == "XLS-BIFF-FORMULA-TOKENS-UNSUPPORTED");
            LegacyXlsWorksheet sheet = Assert.Single(legacy.Worksheets);
            LegacyXlsCell singleReferenceFormula = Assert.Single(sheet.Cells, cell => cell.Row == 2 && cell.Column == 3);
            Assert.True(singleReferenceFormula.IsFormula);
            Assert.Equal(15d, singleReferenceFormula.Value);
            Assert.Equal("A1+5", singleReferenceFormula.FormulaText);
            LegacyXlsCell areaFormula = Assert.Single(sheet.Cells, cell => cell.Row == 2 && cell.Column == 4);
            Assert.True(areaFormula.IsFormula);
            Assert.Equal(100d, areaFormula.Value);
            Assert.Equal("SUM(A1:B2)", areaFormula.FormulaText);

            using ExcelDocument document = ExcelDocument.LoadLegacyXls(new MemoryStream(compound), new LegacyXlsImportOptions {
                ReportUnsupportedContent = true
            });

            Assert.True(document.Sheets[0].TryGetCellText(2, 3, out string? firstCachedText));
            Assert.Equal("15", firstCachedText);
            Assert.True(document.Sheets[0].TryGetCellText(2, 4, out string? secondCachedText));
            Assert.Equal("100", secondCachedText);

            using var output = new MemoryStream();
            document.Save(output);
            using SpreadsheetDocument spreadsheet = SpreadsheetDocument.Open(new MemoryStream(output.ToArray()), false);
            WorksheetPart worksheetPart = spreadsheet.WorkbookPart!.WorksheetParts.Single();
            Dictionary<string, Cell> cells = worksheetPart.Worksheet.Descendants<Cell>().ToDictionary(cell => cell.CellReference!.Value!);
            Assert.Equal("A1+5", cells["C2"].CellFormula!.Text);
            Assert.Equal("15", cells["C2"].CellValue!.Text);
            Assert.Equal("SUM(A1:B2)", cells["D2"].CellFormula!.Text);
            Assert.Equal("100", cells["D2"].CellValue!.Text);
        }

        [Fact]
        public void LegacyXls_Load_ImportsSharedFormulaRecords() {
            byte[] workbookStream = LegacyXlsTestWorkbookBuilder.CreateFormulaSharedFormulaWorkbookStream();
            byte[] compound = LegacyXlsCompoundTestBuilder.CreateWorkbookCompoundFile(workbookStream);

            LegacyXlsWorkbook legacy = LegacyXlsWorkbook.Load(compound, new LegacyXlsImportOptions {
                ReportUnsupportedContent = true
            });

            Assert.DoesNotContain(legacy.Diagnostics, d => d.Severity == LegacyXlsDiagnosticSeverity.Error);
            Assert.DoesNotContain(legacy.Diagnostics, d => d.Code == "XLS-BIFF-FORMULA-TOKENS-UNSUPPORTED");
            Assert.DoesNotContain(legacy.UnsupportedFeatures, feature => feature.RecordType == (ushort)BiffRecordType.ShrFmla);
            LegacyXlsWorksheet sheet = Assert.Single(legacy.Worksheets);
            LegacyXlsCell firstFormula = Assert.Single(sheet.Cells, cell => cell.Row == 1 && cell.Column == 2);
            Assert.True(firstFormula.IsFormula);
            Assert.Equal(15d, firstFormula.Value);
            Assert.Equal("A1+5", firstFormula.FormulaText);
            LegacyXlsCell secondFormula = Assert.Single(sheet.Cells, cell => cell.Row == 2 && cell.Column == 2);
            Assert.True(secondFormula.IsFormula);
            Assert.Equal(25d, secondFormula.Value);
            Assert.Equal("A2+5", secondFormula.FormulaText);
            LegacyXlsCell thirdFormula = Assert.Single(sheet.Cells, cell => cell.Row == 3 && cell.Column == 2);
            Assert.True(thirdFormula.IsFormula);
            Assert.Equal(35d, thirdFormula.Value);
            Assert.Equal("A3+5", thirdFormula.FormulaText);

            using ExcelDocument document = ExcelDocument.LoadLegacyXls(new MemoryStream(compound), new LegacyXlsImportOptions {
                ReportUnsupportedContent = true
            });

            Assert.True(document.Sheets[0].TryGetCellText(1, 2, out string? firstCachedText));
            Assert.Equal("15", firstCachedText);
            Assert.True(document.Sheets[0].TryGetCellText(2, 2, out string? secondCachedText));
            Assert.Equal("25", secondCachedText);
            Assert.True(document.Sheets[0].TryGetCellText(3, 2, out string? thirdCachedText));
            Assert.Equal("35", thirdCachedText);

            using var output = new MemoryStream();
            document.Save(output);
            using SpreadsheetDocument spreadsheet = SpreadsheetDocument.Open(new MemoryStream(output.ToArray()), false);
            WorksheetPart worksheetPart = spreadsheet.WorkbookPart!.WorksheetParts.Single();
            Dictionary<string, Cell> cells = worksheetPart.Worksheet.Descendants<Cell>().ToDictionary(cell => cell.CellReference!.Value!);
            Assert.Equal("A1+5", cells["B1"].CellFormula!.Text);
            Assert.Equal("15", cells["B1"].CellValue!.Text);
            Assert.Equal("A2+5", cells["B2"].CellFormula!.Text);
            Assert.Equal("25", cells["B2"].CellValue!.Text);
            Assert.Equal("A3+5", cells["B3"].CellFormula!.Text);
            Assert.Equal("35", cells["B3"].CellValue!.Text);
        }

        [Fact]
        public void LegacyXls_Load_DoesNotReportArrayFormulaAsUnresolvedSharedFormula() {
            byte[] workbookStream = LegacyXlsTestWorkbookBuilder.CreateFormulaArrayFollowUpWorkbookStream();
            byte[] compound = LegacyXlsCompoundTestBuilder.CreateWorkbookCompoundFile(workbookStream);

            LegacyXlsWorkbook legacy = LegacyXlsWorkbook.Load(compound, new LegacyXlsImportOptions {
                ReportUnsupportedContent = true
            });

            Assert.DoesNotContain(legacy.Diagnostics, d => d.Severity == LegacyXlsDiagnosticSeverity.Error);
            Assert.DoesNotContain(legacy.Diagnostics, d => d.Code == "XLS-BIFF-FORMULA-SHARED-UNRESOLVED");
            Assert.DoesNotContain(legacy.UnsupportedFeatures, feature => feature.RecordType == (ushort)BiffRecordType.Array);
            LegacyXlsImportReport report = legacy.CreateImportReport();
            Assert.Equal(3, report.FormulaTokensByContext["ArrayFormula"]);
            Assert.Equal(3, report.FormulaTokensByContextAndSheet["ArrayFormula|ArrayFormula"]);
            Assert.Equal(1, report.ArrayFormulaRecordCount);
            Assert.Equal(1, report.ArrayFormulasBySheet["ArrayFormula"]);
            Assert.Equal(1, report.ArrayFormulasByRange["D1"]);
            Assert.Equal(1, report.ArrayFormulasBySheetAndRange["ArrayFormula!D1"]);
            Assert.Equal(1, report.ArrayFormulasByDeclaredCellCount["Cells:1"]);
            Assert.Equal(1, report.ArrayFormulasByMatchedFormulaCellCount["Matched:1"]);
            Assert.Equal(1, report.ArrayFormulasByAlwaysCalculateState["NormalCalculation"]);
            Assert.Equal(1, report.ArrayFormulasByProjectionState["FormulaTextProjected"]);
            Assert.Equal(1, report.ArrayFormulasByTokenByteCount["TokenBytes:9"]);
            Assert.Equal(1, report.ArrayFormulasByExtraByteCount["ExtraBytes:0"]);
            LegacyXlsWorksheet sheet = Assert.Single(legacy.Worksheets);
            LegacyXlsArrayFormulaRecord arrayFormula = Assert.Single(sheet.ArrayFormulaRecords);
            Assert.Equal("D1", arrayFormula.Range);
            Assert.False(arrayFormula.AlwaysCalculate);
            Assert.Equal(9, arrayFormula.FormulaTokenByteCount);
            Assert.Equal(0, arrayFormula.FormulaExtraByteCount);
            Assert.Equal(1, arrayFormula.MatchedFormulaCellCount);
            Assert.True(arrayFormula.FormulaTextProjected);
            LegacyXlsCell formula = Assert.Single(sheet.Cells, cell => cell.Row == 1 && cell.Column == 4);
            Assert.True(formula.IsFormula);
            Assert.Equal(31d, formula.Value);
            Assert.Equal("C1+1", formula.FormulaText);
            string markdown = report.ToMarkdown();
            Assert.Contains("Array Formulas By Range", markdown, StringComparison.Ordinal);
            Assert.Contains("Array Formulas By Projection State", markdown, StringComparison.Ordinal);

            using ExcelDocument document = ExcelDocument.LoadLegacyXls(new MemoryStream(compound), new LegacyXlsImportOptions {
                ReportUnsupportedContent = true
            });
            using var output = new MemoryStream();
            document.Save(output);
            using SpreadsheetDocument spreadsheet = SpreadsheetDocument.Open(new MemoryStream(output.ToArray()), false);
            Cell projectedFormula = Assert.Single(spreadsheet.WorkbookPart!.WorksheetParts.Single().Worksheet.Descendants<Cell>(),
                cell => cell.CellReference?.Value == "D1");
            Assert.Equal(CellFormulaValues.Array, projectedFormula.CellFormula!.FormulaType!.Value);
            Assert.Equal("D1", projectedFormula.CellFormula.Reference!.Value);
            Assert.Equal("C1+1", projectedFormula.CellFormula.Text);
        }

        [Fact]
        public void LegacyXls_Load_ProjectsMultiCellArrayFormulaRange() {
            byte[] workbookStream = LegacyXlsTestWorkbookBuilder.CreateFormulaMultiCellArrayWorkbookStream();
            byte[] compound = LegacyXlsCompoundTestBuilder.CreateWorkbookCompoundFile(workbookStream);

            using LegacyXlsLoadResult result = ExcelDocument.LoadLegacyXlsWithReport(new MemoryStream(compound), new LegacyXlsImportOptions {
                ReportUnsupportedContent = true
            });

            LegacyXlsArrayFormulaRecord arrayFormula = Assert.Single(Assert.Single(result.Workbook.Worksheets).ArrayFormulaRecords);
            Assert.Equal("D1:E1", arrayFormula.Range);
            Assert.Equal(2, arrayFormula.DeclaredCellCount);

            using var output = new MemoryStream();
            result.Document.Save(output);
            using SpreadsheetDocument spreadsheet = SpreadsheetDocument.Open(new MemoryStream(output.ToArray()), false);
            Cell projectedFormula = Assert.Single(spreadsheet.WorkbookPart!.WorksheetParts.Single().Worksheet.Descendants<Cell>(),
                cell => cell.CellReference?.Value == "D1");
            Assert.Equal(CellFormulaValues.Array, projectedFormula.CellFormula!.FormulaType!.Value);
            Assert.Equal("D1:E1", projectedFormula.CellFormula.Reference!.Value);
            Assert.Equal("C1+1", projectedFormula.CellFormula.Text);
        }

        [Fact]
        public void LegacyXls_Load_ImportsFormulaArrayConstantTokens() {
            byte[] workbookStream = LegacyXlsTestWorkbookBuilder.CreateFormulaArrayConstantWorkbookStream();
            byte[] compound = LegacyXlsCompoundTestBuilder.CreateWorkbookCompoundFile(workbookStream);

            LegacyXlsWorkbook legacy = LegacyXlsWorkbook.Load(compound, new LegacyXlsImportOptions {
                ReportUnsupportedContent = true
            });

            Assert.DoesNotContain(legacy.Diagnostics, d => d.Severity == LegacyXlsDiagnosticSeverity.Error);
            Assert.DoesNotContain(legacy.Diagnostics, d => d.Code == "XLS-BIFF-FORMULA-TOKENS-UNSUPPORTED");
            LegacyXlsWorksheet sheet = Assert.Single(legacy.Worksheets);
            LegacyXlsCell formula = Assert.Single(sheet.Cells, cell => cell.Row == 1 && cell.Column == 1);
            Assert.True(formula.IsFormula);
            Assert.Equal(10d, formula.Value);
            Assert.Equal("SUM({1,2;3,4})", formula.FormulaText);

            using ExcelDocument document = ExcelDocument.LoadLegacyXls(new MemoryStream(compound), new LegacyXlsImportOptions {
                ReportUnsupportedContent = true
            });

            Assert.True(document.Sheets[0].TryGetCellText(1, 1, out string? cachedText));
            Assert.Equal("10", cachedText);

            using var output = new MemoryStream();
            document.Save(output);
            using SpreadsheetDocument spreadsheet = SpreadsheetDocument.Open(new MemoryStream(output.ToArray()), false);
            WorksheetPart worksheetPart = spreadsheet.WorkbookPart!.WorksheetParts.Single();
            Cell projectedFormula = worksheetPart.Worksheet.Descendants<Cell>().Single(cell => cell.CellReference!.Value == "A1");
            Assert.Equal("SUM({1,2;3,4})", projectedFormula.CellFormula!.Text);
            Assert.Equal("10", projectedFormula.CellValue!.Text);
        }

        [Fact]
        public void LegacyXls_Load_ReportsUnsupportedFormulaTokensAndImportsCachedValue() {
            byte[] workbookStream = LegacyXlsTestWorkbookBuilder.CreateUnsupportedFormulaTokenWorkbookStream();
            byte[] compound = LegacyXlsCompoundTestBuilder.CreateWorkbookCompoundFile(workbookStream);

            LegacyXlsWorkbook legacy = LegacyXlsWorkbook.Load(compound, new LegacyXlsImportOptions {
                ReportUnsupportedContent = true
            });

            Assert.DoesNotContain(legacy.Diagnostics, d => d.Severity == LegacyXlsDiagnosticSeverity.Error);
            LegacyXlsImportDiagnostic diagnostic = Assert.Single(legacy.Diagnostics, d =>
                d.Code == "XLS-BIFF-FORMULA-TOKENS-UNSUPPORTED"
                && d.SheetName == "FormulaDiag"
                && d.RecordType == (ushort)BiffRecordType.Formula);
            Assert.Equal("FormulaToken0x01", diagnostic.DetailCode);
            Assert.True(diagnostic.FormulaToken.HasValue);
            Assert.Equal((byte)0x01, diagnostic.FormulaToken.Value);
            Assert.Equal("PtgExp", diagnostic.FormulaTokenName);
            Assert.True(diagnostic.FormulaTokenOffset.HasValue);
            Assert.Equal(0, diagnostic.FormulaTokenOffset.Value);
            Assert.Contains("Unsupported formula token PtgExp (0x01)", diagnostic.Message);
            Assert.Contains("Token PtgExp (0x01)", diagnostic.Message);
            Assert.Contains("parsed-expression offset 0", diagnostic.Message);
            LegacyXlsImportReport report = legacy.CreateImportReport();
            Assert.Equal(1, report.FormulaTokenBlockers["FormulaToken0x01"]);
            Assert.Equal(1, report.FormulaTokenBlockersByToken["Token:0x01"]);
            Assert.Equal(1, report.FormulaTokenBlockersByTokenName["PtgExp"]);
            Assert.Equal(1, report.FormulaTokenBlockersByOffset["Offset:0"]);
            Assert.Equal(1, report.FormulaTokenBlockersBySheet["FormulaDiag"]);
            Assert.Equal(1, report.FormulaTokensByName["PtgExp"]);
            Assert.Equal(1, report.FormulaTokensByContext["CellFormula"]);
            Assert.Equal(1, report.FormulaTokensBySheet["FormulaDiag"]);
            Assert.Equal(1, report.FormulaTokensByContextAndSheet["CellFormula|FormulaDiag"]);
            Assert.Equal(1, report.FormulaTokensByClass["Base"]);
            Assert.Equal(1, report.FormulaTokensByNameAndClass["PtgExp|Base"]);
            Assert.Equal(1, report.FormulaTokensByOperandByteCount["PtgExp|Bytes:0"]);
            Assert.Empty(report.FormulaTokensByOperandKind);
            Assert.Empty(report.FormulaTokensByNameAndOperandKind);
            Assert.Empty(report.FormulaTokensByOperandKindAndText);
            Assert.Empty(report.FormulaTokensByNameAndOperandText);
            Assert.Equal(1, report.FormulaTokensBySequenceIndex["Index:0"]);
            Assert.Contains(legacy.FormulaTokenRecords, record =>
                record.Context == "CellFormula"
                && record.SheetName == "FormulaDiag"
                && record.CellReference == "B1"
                && record.Token == 0x01
                && record.TokenName == "PtgExp"
                && record.TokenOffset == 0
                && record.SequenceIndex == 0
                && record.TokenClassName == "Base"
                && record.OperandByteCount == 0
                && record.OperandKind == null
                && record.OperandText == null);
            string markdown = report.ToMarkdown();
            Assert.Contains("Formula Token Blockers By Token", markdown);
            Assert.Contains("Formula Token Blockers By Token Name", markdown);
            Assert.Contains("PtgExp", markdown);
            Assert.Contains("Formula Token Blockers By Offset", markdown);
            Assert.Contains("Formula Token Blockers By Sheet", markdown);
            Assert.Contains("Formula Tokens By Name", markdown);
            Assert.Contains("Formula Tokens By Context", markdown);
            Assert.Contains("Formula Tokens By Sheet", markdown);
            Assert.Contains("Formula Tokens By Context And Sheet", markdown);
            Assert.Contains("Formula Tokens By Class", markdown);
            Assert.Contains("Formula Tokens By Name And Class", markdown);
            Assert.Contains("Formula Tokens By Operand Byte Count", markdown);
            Assert.Contains("Formula Tokens By Sequence Index", markdown);
            LegacyXlsWorksheet sheet = Assert.Single(legacy.Worksheets);
            LegacyXlsCell formula = Assert.Single(sheet.Cells, cell => cell.Row == 1 && cell.Column == 2);
            Assert.True(formula.IsFormula);
            Assert.Equal(99d, formula.Value);
            Assert.Null(formula.FormulaText);

            using ExcelDocument document = ExcelDocument.LoadLegacyXls(new MemoryStream(compound), new LegacyXlsImportOptions {
                ReportUnsupportedContent = true
            });

            Assert.True(document.Sheets[0].TryGetCellText(1, 2, out string? cachedText));
            Assert.Equal("99", cachedText);
        }

        [Fact]
        public void LegacyXls_Load_ImportsKnownVariableFormulaFunction() {
            byte[] workbookStream = LegacyXlsTestWorkbookBuilder.CreateFormulaKnownVariableFunctionWorkbookStream();
            byte[] compound = LegacyXlsCompoundTestBuilder.CreateWorkbookCompoundFile(workbookStream);

            LegacyXlsWorkbook legacy = LegacyXlsWorkbook.Load(compound, new LegacyXlsImportOptions {
                ReportUnsupportedContent = true
            });

            Assert.DoesNotContain(legacy.Diagnostics, d => d.Severity == LegacyXlsDiagnosticSeverity.Error);
            Assert.DoesNotContain(legacy.Diagnostics, d => d.Code == "XLS-BIFF-FORMULA-TOKENS-UNSUPPORTED");

            LegacyXlsImportReport report = legacy.CreateImportReport();
            Assert.Equal(1, report.FormulaFunctionsById["Function:0x002E"]);
            Assert.Equal(1, report.FormulaFunctionsByName["VAR"]);
            Assert.Equal(1, report.FormulaFunctionsByParameterCount["VAR|Args:1"]);
            Assert.Equal(1, report.FormulaFunctionsByCetabState["BuiltIn"]);
            Assert.Equal(1, report.FormulaTokensByOperandKindAndText["VariableFunction|VAR"]);
            Assert.Equal(1, report.FormulaTokensByNameAndOperandText["PtgFuncVar|VAR"]);
            Assert.Empty(report.FormulaTokenBlockers);
            Assert.Contains(legacy.FormulaTokenRecords, record =>
                record.Context == "CellFormula"
                && record.SheetName == "KnownFuncDiag"
                && record.CellReference == "B1"
                && record.Token == 0x42
                && record.TokenName == "PtgFuncVar"
                && record.TokenOffset == 9
                && record.SequenceIndex == 1
                && record.TokenClassName == "Reference"
                && record.OperandByteCount == 3
                && record.FunctionId == 0x002e
                && record.FunctionName == "VAR"
                && record.FunctionParameterCount == 1
                && record.FunctionIsCetab == false);

            LegacyXlsWorksheet sheet = Assert.Single(legacy.Worksheets);
            LegacyXlsCell formula = Assert.Single(sheet.Cells, cell => cell.Row == 1 && cell.Column == 2);
            Assert.True(formula.IsFormula);
            Assert.Equal(50d, formula.Value);
            Assert.Equal("VAR(A1:A2)", formula.FormulaText);

            using ExcelDocument document = ExcelDocument.LoadLegacyXls(new MemoryStream(compound), new LegacyXlsImportOptions {
                ReportUnsupportedContent = true
            });

            using var output = new MemoryStream();
            document.Save(output);
            using SpreadsheetDocument spreadsheet = SpreadsheetDocument.Open(new MemoryStream(output.ToArray()), false);
            WorksheetPart worksheetPart = spreadsheet.WorkbookPart!.WorksheetParts.Single();
            AssertProjectedFormula(worksheetPart, "B1", "VAR(A1:A2)", "50");
        }

        [Fact]
        public void LegacyXls_Load_ReportsSharedFormulaTokenNamesAndImportsCachedValue() {
            byte[] workbookStream = LegacyXlsTestWorkbookBuilder.CreateUnsupportedSharedFormulaTokenWorkbookStream();
            byte[] compound = LegacyXlsCompoundTestBuilder.CreateWorkbookCompoundFile(workbookStream);

            LegacyXlsWorkbook legacy = LegacyXlsWorkbook.Load(compound, new LegacyXlsImportOptions {
                ReportUnsupportedContent = true
            });

            Assert.DoesNotContain(legacy.Diagnostics, d => d.Severity == LegacyXlsDiagnosticSeverity.Error);
            LegacyXlsImportDiagnostic diagnostic = Assert.Single(legacy.Diagnostics, d =>
                d.Code == "XLS-BIFF-FORMULA-TOKENS-UNSUPPORTED"
                && d.SheetName == "SharedDiag"
                && d.RecordType == (ushort)BiffRecordType.ShrFmla);
            Assert.Equal("FormulaToken0x01", diagnostic.DetailCode);
            Assert.Equal((byte)0x01, diagnostic.FormulaToken);
            Assert.Equal("PtgExp", diagnostic.FormulaTokenName);
            Assert.Equal(0, diagnostic.FormulaTokenOffset);
            Assert.Contains("Unsupported formula token PtgExp (0x01)", diagnostic.Message);
            Assert.Contains("Shared formula at B1", diagnostic.Message);

            LegacyXlsImportReport report = legacy.CreateImportReport();
            Assert.Equal(1, report.FormulaTokenBlockers["FormulaToken0x01"]);
            Assert.Equal(1, report.FormulaTokenBlockersByToken["Token:0x01"]);
            Assert.Equal(1, report.FormulaTokenBlockersByTokenName["PtgExp"]);
            Assert.Equal(1, report.FormulaTokenBlockersByOffset["Offset:0"]);
            Assert.Equal(1, report.FormulaTokenBlockersBySheet["SharedDiag"]);
            Assert.Equal(2, report.FormulaTokensBySheet["SharedDiag"]);
            Assert.Equal(1, report.FormulaTokensByContextAndSheet["SharedFormulaReference|SharedDiag"]);

            LegacyXlsWorksheet sheet = Assert.Single(legacy.Worksheets);
            LegacyXlsCell formula = Assert.Single(sheet.Cells, cell => cell.Row == 1 && cell.Column == 2);
            Assert.True(formula.IsFormula);
            Assert.Equal(99d, formula.Value);
            Assert.Null(formula.FormulaText);
        }

        [Fact]
        public void LegacyXls_Load_ReportsFormulaFunctionStackUnderflowWithTokenLocation() {
            byte[] workbookStream = LegacyXlsTestWorkbookBuilder.CreateFormulaFunctionStackUnderflowWorkbookStream();
            byte[] compound = LegacyXlsCompoundTestBuilder.CreateWorkbookCompoundFile(workbookStream);

            LegacyXlsWorkbook legacy = LegacyXlsWorkbook.Load(compound, new LegacyXlsImportOptions {
                ReportUnsupportedContent = true
            });

            Assert.DoesNotContain(legacy.Diagnostics, d => d.Severity == LegacyXlsDiagnosticSeverity.Error);
            LegacyXlsImportDiagnostic diagnostic = Assert.Single(legacy.Diagnostics, d =>
                d.Code == "XLS-BIFF-FORMULA-TOKENS-UNSUPPORTED"
                && d.SheetName == "FormulaStack"
                && d.RecordType == (ushort)BiffRecordType.Formula);
            Assert.Equal("FormulaFunction0x0004StackUnderflow", diagnostic.DetailCode);
            Assert.Equal((byte)0x42, diagnostic.FormulaToken);
            Assert.Equal("PtgFuncVar", diagnostic.FormulaTokenName);
            Assert.Equal(0, diagnostic.FormulaTokenOffset);
            Assert.Contains("Formula function SUM (0x0004) expected 1 stack operands but only 0 were available", diagnostic.Message);
            Assert.Contains("Token PtgFuncVar (0x42)", diagnostic.Message);
            Assert.Contains("parsed-expression offset 0", diagnostic.Message);

            LegacyXlsImportReport report = legacy.CreateImportReport();
            Assert.Equal(1, report.FormulaTokenBlockers["FormulaFunction0x0004StackUnderflow"]);
            Assert.Equal(1, report.FormulaTokenBlockersByToken["Token:0x42"]);
            Assert.Equal(1, report.FormulaTokenBlockersByTokenName["PtgFuncVar"]);
            Assert.Equal(1, report.FormulaTokenBlockersByOffset["Offset:0"]);
            Assert.Equal(1, report.FormulaTokenBlockersBySheet["FormulaStack"]);
            Assert.Equal(1, report.FormulaTokensBySheet["FormulaStack"]);
            Assert.Equal(1, report.FormulaTokensByContextAndSheet["CellFormula|FormulaStack"]);

            LegacyXlsWorksheet sheet = Assert.Single(legacy.Worksheets);
            LegacyXlsCell formula = Assert.Single(sheet.Cells, cell => cell.Row == 1 && cell.Column == 2);
            Assert.True(formula.IsFormula);
            Assert.Equal(0d, formula.Value);
            Assert.Null(formula.FormulaText);
        }

        private static void AssertFormula(LegacyXlsWorksheet sheet, int row, int column, string expectedFormula, object expectedValue) {
            LegacyXlsCell formula = Assert.Single(sheet.Cells, cell => cell.Row == row && cell.Column == column);
            Assert.True(formula.IsFormula);
            Assert.Equal(expectedValue, formula.Value);
            Assert.Equal(expectedFormula, formula.FormulaText);
        }

        private static void AssertProjectedFormula(WorksheetPart worksheetPart, string cellReference, string expectedFormula, string expectedValue) {
            Cell projectedFormula = worksheetPart.Worksheet.Descendants<Cell>().Single(cell => cell.CellReference!.Value == cellReference);
            Assert.Equal(expectedFormula, projectedFormula.CellFormula!.Text);
            Assert.Equal(expectedValue, GetProjectedCellValue(worksheetPart, projectedFormula));
        }

        private static string GetProjectedCellValue(WorksheetPart worksheetPart, Cell cell) {
            string value = cell.CellValue!.Text;
            if (cell.DataType?.Value != CellValues.SharedString || !int.TryParse(value, out int sharedStringIndex)) {
                return value;
            }

            WorkbookPart workbookPart = worksheetPart.GetParentParts().OfType<WorkbookPart>().Single();
            return workbookPart.SharedStringTablePart!.SharedStringTable!.Elements<SharedStringItem>().ElementAt(sharedStringIndex).InnerText;
        }

        private static partial class LegacyXlsTestWorkbookBuilder {
            internal static byte[] CreateFormulaStringLiteralWorkbookStream() {
                using var stream = new MemoryStream();
                WriteRecord(stream, 0x0809, new byte[] { 0x00, 0x06, 0x05, 0x00, 0xdb, 0x0b, 0xcc, 0x07 });
                long boundSheetPosition = stream.Position;
                WriteRecord(stream, 0x0085, BuildBoundSheetPayload(0, "FormulaText"));
                WriteRecord(stream, 0x000a, Array.Empty<byte>());

                int sheetOffset = checked((int)stream.Position);
                WriteRecord(stream, 0x0809, new byte[] { 0x00, 0x06, 0x10, 0x00, 0xdb, 0x0b, 0xcc, 0x07 });
                WriteRecord(stream, 0x0204, BuildLabelPayload(0, 0, "Target"));
                WriteRecord(stream, 0x0006, BuildFormulaStringResultPayload(
                    0,
                    1,
                    BuildStringLiteralConcatFormulaTokens("Hello ", 0, 0)));
                WriteRecord(stream, 0x0207, BuildFormulaStringPayload("Hello Target"));
                WriteRecord(stream, 0x000a, Array.Empty<byte>());

                byte[] bytes = stream.ToArray();
                Buffer.BlockCopy(BitConverter.GetBytes(sheetOffset), 0, bytes, checked((int)boundSheetPosition + 4), 4);
                return bytes;
            }

            internal static byte[] CreateFormula3dReferenceWorkbookStream() {
                using var stream = new MemoryStream();
                WriteRecord(stream, 0x0809, new byte[] { 0x00, 0x06, 0x05, 0x00, 0xdb, 0x0b, 0xcc, 0x07 });
                long inputSheetBoundPosition = stream.Position;
                WriteRecord(stream, 0x0085, BuildBoundSheetPayload(0, "Input Data"));
                long totalsSheetBoundPosition = stream.Position;
                WriteRecord(stream, 0x0085, BuildBoundSheetPayload(0, "Totals"));
                WriteRecord(stream, 0x0017, BuildExternSheetPayload((0, 0, 0)));
                WriteRecord(stream, 0x000a, Array.Empty<byte>());

                int inputSheetOffset = checked((int)stream.Position);
                WriteRecord(stream, 0x0809, new byte[] { 0x00, 0x06, 0x10, 0x00, 0xdb, 0x0b, 0xcc, 0x07 });
                WriteRecord(stream, 0x0203, BuildNumberPayload(0, 0, 10d));
                WriteRecord(stream, 0x0203, BuildNumberPayload(1, 0, 32d));
                WriteRecord(stream, 0x000a, Array.Empty<byte>());

                int totalsSheetOffset = checked((int)stream.Position);
                WriteRecord(stream, 0x0809, new byte[] { 0x00, 0x06, 0x10, 0x00, 0xdb, 0x0b, 0xcc, 0x07 });
                WriteRecord(stream, 0x0006, BuildFormulaNumberPayload(0, 0, 15d, formulaTokens: Build3dReferenceAdditionFormulaTokens(0, 0, 0, 5)));
                WriteRecord(stream, 0x0006, BuildFormulaNumberPayload(1, 0, 42d, formulaTokens: BuildSum3dAreaFormulaTokens(0, 0, 0, 1, 0)));
                WriteRecord(stream, 0x000a, Array.Empty<byte>());

                byte[] bytes = stream.ToArray();
                Buffer.BlockCopy(BitConverter.GetBytes(inputSheetOffset), 0, bytes, checked((int)inputSheetBoundPosition + 4), 4);
                Buffer.BlockCopy(BitConverter.GetBytes(totalsSheetOffset), 0, bytes, checked((int)totalsSheetBoundPosition + 4), 4);
                return bytes;
            }

            internal static byte[] CreateFormulaExternalWorkbookReferenceWorkbookStream() =>
                CreateFormulaExternalWorkbookReferenceWorkbookStream("C:\\Data\\Budget.xls");

            internal static byte[] CreateFormulaExternalWorkbookReferenceWorkbookStream(string target) {
                using var stream = new MemoryStream();
                WriteRecord(stream, 0x0809, new byte[] { 0x00, 0x06, 0x05, 0x00, 0xdb, 0x0b, 0xcc, 0x07 });
                long boundSheetPosition = stream.Position;
                WriteRecord(stream, 0x0085, BuildBoundSheetPayload(0, "ExternalFormula"));
                WriteRecord(stream, 0x01ae, BuildSupBookExternalWorkbookPayload(target, "Jan", "Feb"));
                WriteRecord(stream, 0x0017, BuildExternSheetPayload((0, 0, 0)));
                WriteRecord(stream, 0x000a, Array.Empty<byte>());

                int sheetOffset = checked((int)stream.Position);
                WriteRecord(stream, 0x0809, new byte[] { 0x00, 0x06, 0x10, 0x00, 0xdb, 0x0b, 0xcc, 0x07 });
                WriteRecord(stream, 0x0006, BuildFormulaNumberPayload(0, 0, 15d, formulaTokens: Build3dReferenceAdditionFormulaTokens(0, 0, 0, 5)));
                WriteRecord(stream, 0x0006, BuildFormulaNumberPayload(1, 0, 42d, formulaTokens: BuildSum3dAreaFormulaTokens(0, 0, 0, 1, 0)));
                WriteRecord(stream, 0x000a, Array.Empty<byte>());

                byte[] bytes = stream.ToArray();
                Buffer.BlockCopy(BitConverter.GetBytes(sheetOffset), 0, bytes, checked((int)boundSheetPosition + 4), 4);
                return bytes;
            }

            internal static byte[] CreateFormulaStringLiteralMatchingExternalWorkbookStream() {
                using var stream = new MemoryStream();
                WriteRecord(stream, 0x0809, new byte[] { 0x00, 0x06, 0x05, 0x00, 0xdb, 0x0b, 0xcc, 0x07 });
                long boundSheetPosition = stream.Position;
                WriteRecord(stream, 0x0085, BuildBoundSheetPayload(0, "StringLiteral"));
                WriteRecord(stream, 0x01ae, BuildSupBookExternalWorkbookPayload("C:\\Data\\Budget.xls", "Jan"));
                WriteRecord(stream, 0x000a, Array.Empty<byte>());

                int sheetOffset = checked((int)stream.Position);
                WriteRecord(stream, 0x0809, new byte[] { 0x00, 0x06, 0x10, 0x00, 0xdb, 0x0b, 0xcc, 0x07 });
                WriteRecord(stream, 0x0204, BuildLabelPayload(0, 0, "Local"));
                WriteRecord(stream, 0x0006, BuildFormulaStringResultPayload(
                    0,
                    1,
                    BuildStringLiteralConcatFormulaTokens("[Budget.xls]Jan", 0, 0)));
                WriteRecord(stream, 0x0207, BuildFormulaStringPayload("[Budget.xls]JanLocal"));
                WriteRecord(stream, 0x000a, Array.Empty<byte>());

                byte[] bytes = stream.ToArray();
                Buffer.BlockCopy(BitConverter.GetBytes(sheetOffset), 0, bytes, checked((int)boundSheetPosition + 4), 4);
                return bytes;
            }

            internal static byte[] CreateFormulaExternalDefinedNameWorkbookStream() {
                using var stream = new MemoryStream();
                WriteRecord(stream, 0x0809, new byte[] { 0x00, 0x06, 0x05, 0x00, 0xdb, 0x0b, 0xcc, 0x07 });
                long boundSheetPosition = stream.Position;
                WriteRecord(stream, 0x0085, BuildBoundSheetPayload(0, "ExternalName"));
                WriteRecord(stream, 0x01ae, BuildSupBookExternalWorkbookPayload("C:\\Data\\Budget.xls", "Data"));
                WriteRecord(stream, 0x0023, BuildExternalNamePayload("TaxRate"));
                WriteRecord(stream, 0x0017, BuildExternSheetPayload((0, 0, 0)));
                WriteRecord(stream, 0x000a, Array.Empty<byte>());

                int sheetOffset = checked((int)stream.Position);
                WriteRecord(stream, 0x0809, new byte[] { 0x00, 0x06, 0x10, 0x00, 0xdb, 0x0b, 0xcc, 0x07 });
                WriteRecord(stream, 0x0006, BuildFormulaNumberPayload(0, 0, 0.25d, formulaTokens: BuildExternalDefinedNameFormulaTokens(0, 1)));
                WriteRecord(stream, 0x000a, Array.Empty<byte>());

                byte[] bytes = stream.ToArray();
                Buffer.BlockCopy(BitConverter.GetBytes(sheetOffset), 0, bytes, checked((int)boundSheetPosition + 4), 4);
                return bytes;
            }

            internal static byte[] CreateFormulaAddInUserDefinedFunctionWorkbookStream() {
                using var stream = new MemoryStream();
                WriteRecord(stream, 0x0809, new byte[] { 0x00, 0x06, 0x05, 0x00, 0xdb, 0x0b, 0xcc, 0x07 });
                long boundSheetPosition = stream.Position;
                WriteRecord(stream, 0x0085, BuildBoundSheetPayload(0, "AddInUdf"));
                WriteRecord(stream, 0x01ae, BuildSupBookAddInPayload());
                WriteRecord(stream, 0x0023, BuildAddInExternalNamePayload("MYUDF"));
                WriteRecord(stream, 0x0017, BuildExternSheetPayload((0, -2, -2)));
                WriteRecord(stream, 0x000a, Array.Empty<byte>());

                int sheetOffset = checked((int)stream.Position);
                WriteRecord(stream, 0x0809, new byte[] { 0x00, 0x06, 0x10, 0x00, 0xdb, 0x0b, 0xcc, 0x07 });
                WriteRecord(stream, 0x0006, BuildFormulaNumberPayload(0, 0, 20d, formulaTokens: BuildAddInUserDefinedFunctionFormulaTokens(0, 1, 10)));
                WriteRecord(stream, 0x000a, Array.Empty<byte>());

                byte[] bytes = stream.ToArray();
                Buffer.BlockCopy(BitConverter.GetBytes(sheetOffset), 0, bytes, checked((int)boundSheetPosition + 4), 4);
                return bytes;
            }

            internal static byte[] CreateFormula3dReferenceAfterChartSheetWorkbookStream() {
                using var stream = new MemoryStream();
                WriteRecord(stream, 0x0809, new byte[] { 0x00, 0x06, 0x05, 0x00, 0xdb, 0x0b, 0xcc, 0x07 });
                long dataBoundPosition = stream.Position;
                WriteRecord(stream, 0x0085, BuildBoundSheetPayload(0, "Data"));
                long chartBoundPosition = stream.Position;
                WriteRecord(stream, 0x0085, BuildBoundSheetPayload(0, "Chart1", sheetType: 0x02));
                long totalsBoundPosition = stream.Position;
                WriteRecord(stream, 0x0085, BuildBoundSheetPayload(0, "Totals"));
                WriteRecord(stream, 0x0017, BuildExternSheetPayload((0, 2, 2)));
                WriteRecord(stream, 0x0018, BuildDefinedNamePayload("TotalsLocal", BuildNameRef3dFormula(0, 0, 0), localSheetIndex: 3, hidden: false, builtIn: false));
                WriteRecord(stream, 0x000a, Array.Empty<byte>());

                int dataSheetOffset = checked((int)stream.Position);
                WriteRecord(stream, 0x0809, new byte[] { 0x00, 0x06, 0x10, 0x00, 0xdb, 0x0b, 0xcc, 0x07 });
                WriteRecord(stream, 0x0006, BuildFormulaNumberPayload(0, 0, 15d, formulaTokens: Build3dReferenceAdditionFormulaTokens(0, 0, 0, 5)));
                WriteRecord(stream, 0x000a, Array.Empty<byte>());

                int chartSheetOffset = checked((int)stream.Position);
                WriteRecord(stream, 0x0809, new byte[] { 0x00, 0x06, 0x20, 0x00, 0xdb, 0x0b, 0xcc, 0x07 });
                WriteRecord(stream, 0x000a, Array.Empty<byte>());

                int totalsSheetOffset = checked((int)stream.Position);
                WriteRecord(stream, 0x0809, new byte[] { 0x00, 0x06, 0x10, 0x00, 0xdb, 0x0b, 0xcc, 0x07 });
                WriteRecord(stream, 0x0203, BuildNumberPayload(0, 0, 10d));
                WriteRecord(stream, 0x000a, Array.Empty<byte>());

                byte[] bytes = stream.ToArray();
                Buffer.BlockCopy(BitConverter.GetBytes(dataSheetOffset), 0, bytes, checked((int)dataBoundPosition + 4), 4);
                Buffer.BlockCopy(BitConverter.GetBytes(chartSheetOffset), 0, bytes, checked((int)chartBoundPosition + 4), 4);
                Buffer.BlockCopy(BitConverter.GetBytes(totalsSheetOffset), 0, bytes, checked((int)totalsBoundPosition + 4), 4);
                return bytes;
            }

            internal static byte[] CreateFormula3dReferenceToUnsupportedSheetWorkbookStream() {
                using var stream = new MemoryStream();
                WriteRecord(stream, 0x0809, new byte[] { 0x00, 0x06, 0x05, 0x00, 0xdb, 0x0b, 0xcc, 0x07 });
                long dataBoundPosition = stream.Position;
                WriteRecord(stream, 0x0085, BuildBoundSheetPayload(0, "Data"));
                long chartBoundPosition = stream.Position;
                WriteRecord(stream, 0x0085, BuildBoundSheetPayload(0, "Chart1", sheetType: 0x02));
                long formulasBoundPosition = stream.Position;
                WriteRecord(stream, 0x0085, BuildBoundSheetPayload(0, "Formulas"));
                WriteRecord(stream, 0x0017, BuildExternSheetPayload((0, 1, 1)));
                WriteRecord(stream, 0x000a, Array.Empty<byte>());

                int dataSheetOffset = checked((int)stream.Position);
                WriteRecord(stream, 0x0809, new byte[] { 0x00, 0x06, 0x10, 0x00, 0xdb, 0x0b, 0xcc, 0x07 });
                WriteRecord(stream, 0x0203, BuildNumberPayload(0, 0, 10d));
                WriteRecord(stream, 0x000a, Array.Empty<byte>());

                int chartSheetOffset = checked((int)stream.Position);
                WriteRecord(stream, 0x0809, new byte[] { 0x00, 0x06, 0x20, 0x00, 0xdb, 0x0b, 0xcc, 0x07 });
                WriteRecord(stream, 0x000a, Array.Empty<byte>());

                int formulasSheetOffset = checked((int)stream.Position);
                WriteRecord(stream, 0x0809, new byte[] { 0x00, 0x06, 0x10, 0x00, 0xdb, 0x0b, 0xcc, 0x07 });
                WriteRecord(stream, 0x0006, BuildFormulaNumberPayload(0, 0, 15d, formulaTokens: Build3dReferenceAdditionFormulaTokens(0, 0, 0, 5)));
                WriteRecord(stream, 0x000a, Array.Empty<byte>());

                byte[] bytes = stream.ToArray();
                Buffer.BlockCopy(BitConverter.GetBytes(dataSheetOffset), 0, bytes, checked((int)dataBoundPosition + 4), 4);
                Buffer.BlockCopy(BitConverter.GetBytes(chartSheetOffset), 0, bytes, checked((int)chartBoundPosition + 4), 4);
                Buffer.BlockCopy(BitConverter.GetBytes(formulasSheetOffset), 0, bytes, checked((int)formulasBoundPosition + 4), 4);
                return bytes;
            }

            internal static byte[] CreateFormulaFixedFunctionWorkbookStream() {
                using var stream = new MemoryStream();
                WriteRecord(stream, 0x0809, new byte[] { 0x00, 0x06, 0x05, 0x00, 0xdb, 0x0b, 0xcc, 0x07 });
                long boundSheetPosition = stream.Position;
                WriteRecord(stream, 0x0085, BuildBoundSheetPayload(0, "FixedFunc"));
                WriteRecord(stream, 0x000a, Array.Empty<byte>());

                int sheetOffset = checked((int)stream.Position);
                WriteRecord(stream, 0x0809, new byte[] { 0x00, 0x06, 0x10, 0x00, 0xdb, 0x0b, 0xcc, 0x07 });
                WriteRecord(stream, 0x0203, BuildNumberPayload(0, 0, 12.345d));
                WriteRecord(stream, 0x0203, BuildNumberPayload(1, 0, 2d));
                WriteRecord(stream, 0x0203, BuildNumberPayload(2, 0, 3d));
                WriteRecord(stream, 0x0203, BuildNumberPayload(1, 1, 20d));
                WriteRecord(stream, 0x0203, BuildNumberPayload(2, 1, 30d));
                WriteRecord(stream, 0x0006, BuildFormulaNumberPayload(0, 1, 12.35d, formulaTokens: BuildRoundFormulaTokens(0, 0, 2)));
                WriteRecord(stream, 0x0006, BuildFormulaBooleanPayload(0, 2, value: true, formulaTokens: BuildLogicalFixedFunctionFormulaTokens(0x0024, true, true)));
                WriteRecord(stream, 0x0006, BuildFormulaBooleanPayload(0, 3, value: false, formulaTokens: BuildLogicalFixedFunctionFormulaTokens(0x0025, false, false)));
                WriteRecord(stream, 0x0006, BuildFormulaNumberPayload(2, 4, 1d, formulaTokens: BuildRsqFormulaTokens()));
                WriteRecord(stream, 0x0006, BuildFormulaErrorResultPayload(2, 5, 0x2a, formulaTokens: BuildFixedNoArgumentFunctionFormulaTokens(0x000a)));
                WriteRecord(stream, 0x000a, Array.Empty<byte>());

                byte[] bytes = stream.ToArray();
                Buffer.BlockCopy(BitConverter.GetBytes(sheetOffset), 0, bytes, checked((int)boundSheetPosition + 4), 4);
                return bytes;
            }

            internal static byte[] CreateFormulaCommonFunctionWorkbookStream() {
                using var stream = new MemoryStream();
                WriteRecord(stream, 0x0809, new byte[] { 0x00, 0x06, 0x05, 0x00, 0xdb, 0x0b, 0xcc, 0x07 });
                long boundSheetPosition = stream.Position;
                WriteRecord(stream, 0x0085, BuildBoundSheetPayload(0, "CommonFunc"));
                WriteRecord(stream, 0x000a, Array.Empty<byte>());

                int sheetOffset = checked((int)stream.Position);
                WriteRecord(stream, 0x0809, new byte[] { 0x00, 0x06, 0x10, 0x00, 0xdb, 0x0b, 0xcc, 0x07 });
                WriteRecord(stream, 0x0204, BuildLabelPayload(0, 0, "OfficeIMO"));
                WriteRecord(stream, 0x0203, BuildNumberPayload(0, 1, 45566d));
                WriteRecord(stream, 0x0204, BuildLabelPayload(0, 3, "EU"));
                WriteRecord(stream, 0x0204, BuildLabelPayload(1, 3, "US"));
                WriteRecord(stream, 0x0204, BuildLabelPayload(2, 3, "EU"));
                WriteRecord(stream, 0x0203, BuildNumberPayload(0, 4, 10d));
                WriteRecord(stream, 0x0203, BuildNumberPayload(1, 4, 20d));
                WriteRecord(stream, 0x0203, BuildNumberPayload(2, 4, 30d));
                WriteRecord(stream, 0x0006, BuildFormulaStringResultPayload(1, 0, BuildLeftRightFormulaTokens(0x0073, 0, 0, 6)));
                WriteRecord(stream, 0x0207, BuildFormulaStringPayload("Office"));
                WriteRecord(stream, 0x0006, BuildFormulaStringResultPayload(2, 0, BuildLeftRightFormulaTokens(0x0074, 0, 0, 3)));
                WriteRecord(stream, 0x0207, BuildFormulaStringPayload("IMO"));
                WriteRecord(stream, 0x0006, BuildFormulaNumberPayload(3, 0, 46196d, formulaTokens: BuildDateFormulaTokens(2026, 6, 23)));
                WriteRecord(stream, 0x0006, BuildFormulaNumberPayload(4, 0, 2024d, formulaTokens: BuildSingleReferenceFixedFunctionFormulaTokens(0x0045, 0, 1)));
                WriteRecord(stream, 0x0006, BuildFormulaNumberPayload(5, 0, 10d, formulaTokens: BuildSingleReferenceFixedFunctionFormulaTokens(0x0044, 0, 1)));
                WriteRecord(stream, 0x0006, BuildFormulaNumberPayload(6, 0, 1d, formulaTokens: BuildSingleReferenceFixedFunctionFormulaTokens(0x0043, 0, 1)));
                WriteRecord(stream, 0x0006, BuildFormulaNumberPayload(7, 0, 2d, formulaTokens: BuildCountIfFormulaTokens()));
                WriteRecord(stream, 0x0006, BuildFormulaStringResultPayload(8, 0, BuildSingleReferenceFixedFunctionFormulaTokens(0x0070, 0, 0)));
                WriteRecord(stream, 0x0207, BuildFormulaStringPayload("officeimo"));
                WriteRecord(stream, 0x0006, BuildFormulaStringResultPayload(9, 0, BuildSingleReferenceFixedFunctionFormulaTokens(0x0071, 0, 0)));
                WriteRecord(stream, 0x0207, BuildFormulaStringPayload("OFFICEIMO"));
                WriteRecord(stream, 0x0006, BuildFormulaStringResultPayload(10, 0, BuildSingleReferenceFixedFunctionFormulaTokens(0x0072, 0, 0)));
                WriteRecord(stream, 0x0207, BuildFormulaStringPayload("Officeimo"));
                WriteRecord(stream, 0x0006, BuildFormulaNumberPayload(11, 0, 40d, formulaTokens: BuildSumIfFormulaTokens()));
                WriteRecord(stream, 0x0006, BuildFormulaNumberPayload(12, 0, 60d, formulaTokens: BuildSubtotalFormulaTokens()));
                WriteRecord(stream, 0x0006, BuildFormulaNumberPayload(13, 0, 20d, formulaTokens: BuildIndexFormulaTokens()));
                WriteRecord(stream, 0x0006, BuildFormulaNumberPayload(14, 0, 2d, formulaTokens: BuildMatchFormulaTokens()));
                WriteRecord(stream, 0x0006, BuildFormulaNumberPayload(15, 0, 20d, formulaTokens: BuildOffsetFormulaTokens()));
                WriteRecord(stream, 0x0006, BuildFormulaNumberPayload(16, 0, 46196d, formulaTokens: BuildVolatileFixedNoArgumentFunctionFormulaTokens(0x00dd)));
                WriteRecord(stream, 0x0006, BuildFormulaNumberPayload(17, 0, 46196.5d, formulaTokens: BuildVolatileFixedNoArgumentFunctionFormulaTokens(0x004a)));
                WriteRecord(stream, 0x0006, BuildFormulaNumberPayload(18, 0, 1d, formulaTokens: BuildSingleReferenceFixedFunctionFormulaTokens(0x0008, 0, 0)));
                WriteRecord(stream, 0x0006, BuildFormulaNumberPayload(19, 0, 2d, formulaTokens: BuildSingleReferenceFixedFunctionFormulaTokens(0x0009, 0, 1)));
                WriteRecord(stream, 0x0006, BuildFormulaNumberPayload(20, 0, 0.395833333333333d, formulaTokens: BuildThreeIntegerFixedFunctionFormulaTokens(0x0042, 9, 30, 0)));
                WriteRecord(stream, 0x0006, BuildFormulaNumberPayload(21, 0, 12d, formulaTokens: BuildSingleNumberFixedFunctionFormulaTokens(0x0047, 0.5d)));
                WriteRecord(stream, 0x0006, BuildFormulaNumberPayload(22, 0, 0d, formulaTokens: BuildSingleNumberFixedFunctionFormulaTokens(0x0048, 0.5d)));
                WriteRecord(stream, 0x0006, BuildFormulaNumberPayload(23, 0, 0d, formulaTokens: BuildSingleNumberFixedFunctionFormulaTokens(0x0049, 0.5d)));
                WriteRecord(stream, 0x0006, BuildFormulaNumberPayload(24, 0, 0.42d, formulaTokens: BuildVolatileFixedNoArgumentFunctionFormulaTokens(0x003f)));
                WriteRecord(stream, 0x0006, BuildFormulaNumberPayload(25, 0, 3d, formulaTokens: BuildSingleAreaFixedFunctionFormulaTokens(0x004c, 0, 4, 2, 4)));
                WriteRecord(stream, 0x0006, BuildFormulaNumberPayload(26, 0, 2d, formulaTokens: BuildSingleAreaFixedFunctionFormulaTokens(0x004d, 0, 3, 0, 4)));
                WriteRecord(stream, 0x0006, BuildFormulaNumberPayload(27, 0, 20d, formulaTokens: BuildLargeFormulaTokens()));
                WriteRecord(stream, 0x0006, BuildFormulaNumberPayload(28, 0, 10d, formulaTokens: BuildStDevFormulaTokens()));
                WriteRecord(stream, 0x0006, BuildFormulaNumberPayload(29, 0, 100d, formulaTokens: BuildNpvFormulaTokens()));
                WriteRecord(stream, 0x000a, Array.Empty<byte>());

                byte[] bytes = stream.ToArray();
                Buffer.BlockCopy(BitConverter.GetBytes(sheetOffset), 0, bytes, checked((int)boundSheetPosition + 4), 4);
                return bytes;
            }

            internal static byte[] CreateFormulaLookupFunctionWorkbookStream() {
                using var stream = new MemoryStream();
                WriteRecord(stream, 0x0809, new byte[] { 0x00, 0x06, 0x05, 0x00, 0xdb, 0x0b, 0xcc, 0x07 });
                long boundSheetPosition = stream.Position;
                WriteRecord(stream, 0x0085, BuildBoundSheetPayload(0, "LookupFunc"));
                WriteRecord(stream, 0x000a, Array.Empty<byte>());

                int sheetOffset = checked((int)stream.Position);
                WriteRecord(stream, 0x0809, new byte[] { 0x00, 0x06, 0x10, 0x00, 0xdb, 0x0b, 0xcc, 0x07 });
                WriteRecord(stream, 0x0203, BuildNumberPayload(0, 0, 2d));
                WriteRecord(stream, 0x0203, BuildNumberPayload(0, 1, 1d));
                WriteRecord(stream, 0x0203, BuildNumberPayload(0, 2, 100d));
                WriteRecord(stream, 0x0203, BuildNumberPayload(1, 1, 2d));
                WriteRecord(stream, 0x0203, BuildNumberPayload(1, 2, 200d));
                WriteRecord(stream, 0x0006, BuildFormulaNumberPayload(2, 3, 200d, formulaTokens: BuildVLookupFormulaTokens()));
                WriteRecord(stream, 0x000a, Array.Empty<byte>());

                byte[] bytes = stream.ToArray();
                Buffer.BlockCopy(BitConverter.GetBytes(sheetOffset), 0, bytes, checked((int)boundSheetPosition + 4), 4);
                return bytes;
            }

            internal static byte[] CreateUnsupportedFormulaTokenWorkbookStream() {
                using var stream = new MemoryStream();
                WriteRecord(stream, 0x0809, new byte[] { 0x00, 0x06, 0x05, 0x00, 0xdb, 0x0b, 0xcc, 0x07 });
                long boundSheetPosition = stream.Position;
                WriteRecord(stream, 0x0085, BuildBoundSheetPayload(0, "FormulaDiag"));
                WriteRecord(stream, 0x000a, Array.Empty<byte>());

                int sheetOffset = checked((int)stream.Position);
                WriteRecord(stream, 0x0809, new byte[] { 0x00, 0x06, 0x10, 0x00, 0xdb, 0x0b, 0xcc, 0x07 });
                WriteRecord(stream, 0x0203, BuildNumberPayload(0, 0, 99d));
                WriteRecord(stream, 0x0006, BuildFormulaNumberPayload(0, 1, 99d, formulaTokens: new byte[] { 0x01 }));
                WriteRecord(stream, 0x000a, Array.Empty<byte>());

                byte[] bytes = stream.ToArray();
                Buffer.BlockCopy(BitConverter.GetBytes(sheetOffset), 0, bytes, checked((int)boundSheetPosition + 4), 4);
                return bytes;
            }

            internal static byte[] CreateFormulaKnownVariableFunctionWorkbookStream() {
                using var stream = new MemoryStream();
                WriteRecord(stream, 0x0809, new byte[] { 0x00, 0x06, 0x05, 0x00, 0xdb, 0x0b, 0xcc, 0x07 });
                long boundSheetPosition = stream.Position;
                WriteRecord(stream, 0x0085, BuildBoundSheetPayload(0, "KnownFuncDiag"));
                WriteRecord(stream, 0x000a, Array.Empty<byte>());

                int sheetOffset = checked((int)stream.Position);
                WriteRecord(stream, 0x0809, new byte[] { 0x00, 0x06, 0x10, 0x00, 0xdb, 0x0b, 0xcc, 0x07 });
                WriteRecord(stream, 0x0203, BuildNumberPayload(0, 0, 10d));
                WriteRecord(stream, 0x0203, BuildNumberPayload(1, 0, 20d));
                WriteRecord(stream, 0x0006, BuildFormulaNumberPayload(0, 1, 50d, formulaTokens: BuildVarFormulaTokens()));
                WriteRecord(stream, 0x000a, Array.Empty<byte>());

                byte[] bytes = stream.ToArray();
                Buffer.BlockCopy(BitConverter.GetBytes(sheetOffset), 0, bytes, checked((int)boundSheetPosition + 4), 4);
                return bytes;
            }

            internal static byte[] CreateUnsupportedSharedFormulaTokenWorkbookStream() {
                using var stream = new MemoryStream();
                WriteRecord(stream, 0x0809, new byte[] { 0x00, 0x06, 0x05, 0x00, 0xdb, 0x0b, 0xcc, 0x07 });
                long boundSheetPosition = stream.Position;
                WriteRecord(stream, 0x0085, BuildBoundSheetPayload(0, "SharedDiag"));
                WriteRecord(stream, 0x000a, Array.Empty<byte>());

                int sheetOffset = checked((int)stream.Position);
                WriteRecord(stream, 0x0809, new byte[] { 0x00, 0x06, 0x10, 0x00, 0xdb, 0x0b, 0xcc, 0x07 });
                WriteRecord(stream, 0x0203, BuildNumberPayload(0, 0, 99d));
                WriteRecord(stream, 0x0006, BuildFormulaNumberPayload(0, 1, 99d, formulaTokens: BuildSharedFormulaReferenceTokens(0, 1)));
                WriteRecord(stream, 0x04bc, BuildSharedFormulaPayload(0, 1, 0, 1, new byte[] { 0x01 }));
                WriteRecord(stream, 0x000a, Array.Empty<byte>());

                byte[] bytes = stream.ToArray();
                Buffer.BlockCopy(BitConverter.GetBytes(sheetOffset), 0, bytes, checked((int)boundSheetPosition + 4), 4);
                return bytes;
            }

            internal static byte[] CreateFormulaFunctionStackUnderflowWorkbookStream() {
                using var stream = new MemoryStream();
                WriteRecord(stream, 0x0809, new byte[] { 0x00, 0x06, 0x05, 0x00, 0xdb, 0x0b, 0xcc, 0x07 });
                long boundSheetPosition = stream.Position;
                WriteRecord(stream, 0x0085, BuildBoundSheetPayload(0, "FormulaStack"));
                WriteRecord(stream, 0x000a, Array.Empty<byte>());

                int sheetOffset = checked((int)stream.Position);
                WriteRecord(stream, 0x0809, new byte[] { 0x00, 0x06, 0x10, 0x00, 0xdb, 0x0b, 0xcc, 0x07 });
                WriteRecord(stream, 0x0006, BuildFormulaNumberPayload(0, 1, 0d, formulaTokens: BuildFunctionStackUnderflowFormulaTokens()));
                WriteRecord(stream, 0x000a, Array.Empty<byte>());

                byte[] bytes = stream.ToArray();
                Buffer.BlockCopy(BitConverter.GetBytes(sheetOffset), 0, bytes, checked((int)boundSheetPosition + 4), 4);
                return bytes;
            }

            internal static byte[] CreateFormulaAttributeSumWorkbookStream() {
                using var stream = new MemoryStream();
                WriteRecord(stream, 0x0809, new byte[] { 0x00, 0x06, 0x05, 0x00, 0xdb, 0x0b, 0xcc, 0x07 });
                long boundSheetPosition = stream.Position;
                WriteRecord(stream, 0x0085, BuildBoundSheetPayload(0, "AttrSum"));
                WriteRecord(stream, 0x000a, Array.Empty<byte>());

                int sheetOffset = checked((int)stream.Position);
                WriteRecord(stream, 0x0809, new byte[] { 0x00, 0x06, 0x10, 0x00, 0xdb, 0x0b, 0xcc, 0x07 });
                WriteRecord(stream, 0x0203, BuildNumberPayload(0, 0, 10d));
                WriteRecord(stream, 0x0203, BuildNumberPayload(0, 1, 20d));
                WriteRecord(stream, 0x0006, BuildFormulaNumberPayload(0, 2, 30d, formulaTokens: BuildAttributeSumFormulaTokens(0, 0, 0, 1)));
                WriteRecord(stream, 0x000a, Array.Empty<byte>());

                byte[] bytes = stream.ToArray();
                Buffer.BlockCopy(BitConverter.GetBytes(sheetOffset), 0, bytes, checked((int)boundSheetPosition + 4), 4);
                return bytes;
            }

            internal static byte[] CreateFormulaDisplaySpaceWorkbookStream() {
                using var stream = new MemoryStream();
                WriteRecord(stream, 0x0809, new byte[] { 0x00, 0x06, 0x05, 0x00, 0xdb, 0x0b, 0xcc, 0x07 });
                long boundSheetPosition = stream.Position;
                WriteRecord(stream, 0x0085, BuildBoundSheetPayload(0, "DisplaySpace"));
                WriteRecord(stream, 0x000a, Array.Empty<byte>());

                int sheetOffset = checked((int)stream.Position);
                WriteRecord(stream, 0x0809, new byte[] { 0x00, 0x06, 0x10, 0x00, 0xdb, 0x0b, 0xcc, 0x07 });
                WriteRecord(stream, 0x0203, BuildNumberPayload(0, 0, 10d));
                WriteRecord(stream, 0x0006, BuildFormulaNumberPayload(0, 1, 15d, formulaTokens: BuildDisplaySpaceAdditionFormulaTokens(0x40, 0, 0, 5)));
                WriteRecord(stream, 0x0006, BuildFormulaNumberPayload(0, 2, 20d, formulaTokens: BuildDisplaySpaceAdditionFormulaTokens(0x41, 0, 0, 10)));
                WriteRecord(stream, 0x000a, Array.Empty<byte>());

                byte[] bytes = stream.ToArray();
                Buffer.BlockCopy(BitConverter.GetBytes(sheetOffset), 0, bytes, checked((int)boundSheetPosition + 4), 4);
                return bytes;
            }

            internal static byte[] CreateFormulaIfWorkbookStream() {
                using var stream = new MemoryStream();
                WriteRecord(stream, 0x0809, new byte[] { 0x00, 0x06, 0x05, 0x00, 0xdb, 0x0b, 0xcc, 0x07 });
                long boundSheetPosition = stream.Position;
                WriteRecord(stream, 0x0085, BuildBoundSheetPayload(0, "IfFormula"));
                WriteRecord(stream, 0x000a, Array.Empty<byte>());

                int sheetOffset = checked((int)stream.Position);
                WriteRecord(stream, 0x0809, new byte[] { 0x00, 0x06, 0x10, 0x00, 0xdb, 0x0b, 0xcc, 0x07 });
                WriteRecord(stream, 0x0203, BuildNumberPayload(0, 0, 10d));
                WriteRecord(stream, 0x0006, BuildFormulaNumberPayload(0, 1, 1d, formulaTokens: BuildIfFormulaTokens(0, 0, 5, 1, 0)));
                WriteRecord(stream, 0x000a, Array.Empty<byte>());

                byte[] bytes = stream.ToArray();
                Buffer.BlockCopy(BitConverter.GetBytes(sheetOffset), 0, bytes, checked((int)boundSheetPosition + 4), 4);
                return bytes;
            }

            internal static byte[] CreateFormulaChooseWorkbookStream() {
                using var stream = new MemoryStream();
                WriteRecord(stream, 0x0809, new byte[] { 0x00, 0x06, 0x05, 0x00, 0xdb, 0x0b, 0xcc, 0x07 });
                long boundSheetPosition = stream.Position;
                WriteRecord(stream, 0x0085, BuildBoundSheetPayload(0, "ChooseFormula"));
                WriteRecord(stream, 0x000a, Array.Empty<byte>());

                int sheetOffset = checked((int)stream.Position);
                WriteRecord(stream, 0x0809, new byte[] { 0x00, 0x06, 0x10, 0x00, 0xdb, 0x0b, 0xcc, 0x07 });
                WriteRecord(stream, 0x0203, BuildNumberPayload(0, 0, 2d));
                WriteRecord(stream, 0x0006, BuildFormulaNumberPayload(0, 1, 20d, formulaTokens: BuildChooseFormulaTokens(0, 0, 10, 20, 30)));
                WriteRecord(stream, 0x000a, Array.Empty<byte>());

                byte[] bytes = stream.ToArray();
                Buffer.BlockCopy(BitConverter.GetBytes(sheetOffset), 0, bytes, checked((int)boundSheetPosition + 4), 4);
                return bytes;
            }

            internal static byte[] CreateFormulaReferenceOperatorWorkbookStream() {
                using var stream = new MemoryStream();
                WriteRecord(stream, 0x0809, new byte[] { 0x00, 0x06, 0x05, 0x00, 0xdb, 0x0b, 0xcc, 0x07 });
                long boundSheetPosition = stream.Position;
                WriteRecord(stream, 0x0085, BuildBoundSheetPayload(0, "RefOps"));
                WriteRecord(stream, 0x000a, Array.Empty<byte>());

                int sheetOffset = checked((int)stream.Position);
                WriteRecord(stream, 0x0809, new byte[] { 0x00, 0x06, 0x10, 0x00, 0xdb, 0x0b, 0xcc, 0x07 });
                WriteRecord(stream, 0x0203, BuildNumberPayload(0, 0, 10d));
                WriteRecord(stream, 0x0203, BuildNumberPayload(0, 1, 20d));
                WriteRecord(stream, 0x0203, BuildNumberPayload(0, 2, 30d));
                WriteRecord(stream, 0x0006, BuildFormulaNumberPayload(0, 3, 30d, formulaTokens: BuildRangeOperatorSumFormulaTokens()));
                WriteRecord(stream, 0x0006, BuildFormulaNumberPayload(0, 4, 30d, formulaTokens: BuildUnionOperatorSumFormulaTokens()));
                WriteRecord(stream, 0x0006, BuildFormulaNumberPayload(0, 5, 20d, formulaTokens: BuildIntersectionOperatorSumFormulaTokens()));
                WriteRecord(stream, 0x000a, Array.Empty<byte>());

                byte[] bytes = stream.ToArray();
                Buffer.BlockCopy(BitConverter.GetBytes(sheetOffset), 0, bytes, checked((int)boundSheetPosition + 4), 4);
                return bytes;
            }

            internal static byte[] CreateFormulaMemAreaWorkbookStream() {
                using var stream = new MemoryStream();
                WriteRecord(stream, 0x0809, new byte[] { 0x00, 0x06, 0x05, 0x00, 0xdb, 0x0b, 0xcc, 0x07 });
                long boundSheetPosition = stream.Position;
                WriteRecord(stream, 0x0085, BuildBoundSheetPayload(0, "MemArea"));
                WriteRecord(stream, 0x000a, Array.Empty<byte>());

                int sheetOffset = checked((int)stream.Position);
                WriteRecord(stream, 0x0809, new byte[] { 0x00, 0x06, 0x10, 0x00, 0xdb, 0x0b, 0xcc, 0x07 });
                WriteRecord(stream, 0x0203, BuildNumberPayload(0, 0, 10d));
                WriteRecord(stream, 0x0203, BuildNumberPayload(1, 0, 20d));
                WriteRecord(stream, 0x0006, BuildFormulaNumberPayload(2, 0, 30d, formulaTokens: BuildMemAreaSumFormulaTokens()));
                WriteRecord(stream, 0x000a, Array.Empty<byte>());

                byte[] bytes = stream.ToArray();
                Buffer.BlockCopy(BitConverter.GetBytes(sheetOffset), 0, bytes, checked((int)boundSheetPosition + 4), 4);
                return bytes;
            }

            internal static byte[] CreateFormulaMemFuncWorkbookStream() {
                using var stream = new MemoryStream();
                WriteRecord(stream, 0x0809, new byte[] { 0x00, 0x06, 0x05, 0x00, 0xdb, 0x0b, 0xcc, 0x07 });
                long boundSheetPosition = stream.Position;
                WriteRecord(stream, 0x0085, BuildBoundSheetPayload(0, "MemFunc"));
                WriteRecord(stream, 0x000a, Array.Empty<byte>());

                int sheetOffset = checked((int)stream.Position);
                WriteRecord(stream, 0x0809, new byte[] { 0x00, 0x06, 0x10, 0x00, 0xdb, 0x0b, 0xcc, 0x07 });
                WriteRecord(stream, 0x0203, BuildNumberPayload(0, 0, 10d));
                WriteRecord(stream, 0x0203, BuildNumberPayload(0, 1, 20d));
                WriteRecord(stream, 0x0006, BuildFormulaNumberPayload(0, 2, 30d, formulaTokens: BuildMemFuncSumFormulaTokens()));
                WriteRecord(stream, 0x000a, Array.Empty<byte>());

                byte[] bytes = stream.ToArray();
                Buffer.BlockCopy(BitConverter.GetBytes(sheetOffset), 0, bytes, checked((int)boundSheetPosition + 4), 4);
                return bytes;
            }

            internal static byte[] CreateFormulaRelativeReferenceWorkbookStream() {
                using var stream = new MemoryStream();
                WriteRecord(stream, 0x0809, new byte[] { 0x00, 0x06, 0x05, 0x00, 0xdb, 0x0b, 0xcc, 0x07 });
                long boundSheetPosition = stream.Position;
                WriteRecord(stream, 0x0085, BuildBoundSheetPayload(0, "RelativeRefs"));
                WriteRecord(stream, 0x000a, Array.Empty<byte>());

                int sheetOffset = checked((int)stream.Position);
                WriteRecord(stream, 0x0809, new byte[] { 0x00, 0x06, 0x10, 0x00, 0xdb, 0x0b, 0xcc, 0x07 });
                WriteRecord(stream, 0x0203, BuildNumberPayload(0, 0, 10d));
                WriteRecord(stream, 0x0203, BuildNumberPayload(0, 1, 20d));
                WriteRecord(stream, 0x0203, BuildNumberPayload(1, 0, 30d));
                WriteRecord(stream, 0x0203, BuildNumberPayload(1, 1, 40d));
                WriteRecord(stream, 0x0006, BuildFormulaNumberPayload(1, 2, 15d, formulaTokens: BuildRelativeReferenceAdditionFormulaTokens(-1, -2, 5)));
                WriteRecord(stream, 0x0006, BuildFormulaNumberPayload(1, 3, 100d, formulaTokens: BuildRelativeAreaSumFormulaTokens(-1, -3, 0, -2)));
                WriteRecord(stream, 0x000a, Array.Empty<byte>());

                byte[] bytes = stream.ToArray();
                Buffer.BlockCopy(BitConverter.GetBytes(sheetOffset), 0, bytes, checked((int)boundSheetPosition + 4), 4);
                return bytes;
            }

            internal static byte[] CreateFormulaSharedFormulaWorkbookStream() {
                using var stream = new MemoryStream();
                WriteRecord(stream, 0x0809, new byte[] { 0x00, 0x06, 0x05, 0x00, 0xdb, 0x0b, 0xcc, 0x07 });
                long boundSheetPosition = stream.Position;
                WriteRecord(stream, 0x0085, BuildBoundSheetPayload(0, "SharedFormula"));
                WriteRecord(stream, 0x000a, Array.Empty<byte>());

                int sheetOffset = checked((int)stream.Position);
                WriteRecord(stream, 0x0809, new byte[] { 0x00, 0x06, 0x10, 0x00, 0xdb, 0x0b, 0xcc, 0x07 });
                WriteRecord(stream, 0x0203, BuildNumberPayload(0, 0, 10d));
                WriteRecord(stream, 0x0203, BuildNumberPayload(1, 0, 20d));
                WriteRecord(stream, 0x0203, BuildNumberPayload(2, 0, 30d));
                WriteRecord(stream, 0x0006, BuildFormulaNumberPayload(0, 1, 15d, formulaTokens: BuildSharedFormulaReferenceTokens(0, 1)));
                WriteRecord(stream, 0x04bc, BuildSharedFormulaPayload(0, 1, 2, 1, BuildRelativeReferenceAdditionFormulaTokens(0, -1, 5)));
                WriteRecord(stream, 0x0006, BuildFormulaNumberPayload(1, 1, 25d, formulaTokens: BuildSharedFormulaReferenceTokens(0, 1)));
                WriteRecord(stream, 0x0006, BuildFormulaNumberPayload(2, 1, 35d, formulaTokens: BuildSharedFormulaReferenceTokens(0, 1)));
                WriteRecord(stream, 0x000a, Array.Empty<byte>());

                byte[] bytes = stream.ToArray();
                Buffer.BlockCopy(BitConverter.GetBytes(sheetOffset), 0, bytes, checked((int)boundSheetPosition + 4), 4);
                return bytes;
            }

            internal static byte[] CreateFormulaArrayFollowUpWorkbookStream() {
                using var stream = new MemoryStream();
                WriteRecord(stream, 0x0809, new byte[] { 0x00, 0x06, 0x05, 0x00, 0xdb, 0x0b, 0xcc, 0x07 });
                long boundSheetPosition = stream.Position;
                WriteRecord(stream, 0x0085, BuildBoundSheetPayload(0, "ArrayFormula"));
                WriteRecord(stream, 0x000a, Array.Empty<byte>());

                int sheetOffset = checked((int)stream.Position);
                WriteRecord(stream, 0x0809, new byte[] { 0x00, 0x06, 0x10, 0x00, 0xdb, 0x0b, 0xcc, 0x07 });
                WriteRecord(stream, 0x0203, BuildNumberPayload(0, 0, 10d));
                WriteRecord(stream, 0x0203, BuildNumberPayload(0, 1, 20d));
                WriteRecord(stream, 0x0203, BuildNumberPayload(0, 2, 30d));
                WriteRecord(stream, 0x0006, BuildFormulaNumberPayload(0, 3, 31d, formulaTokens: BuildSharedFormulaReferenceTokens(0, 3)));
                WriteRecord(stream, 0x0221, BuildArrayFormulaPayload(0, 3, 0, 3, BuildRelativeReferenceAdditionFormulaTokens(0, -1, 1)));
                WriteRecord(stream, 0x000a, Array.Empty<byte>());

                byte[] bytes = stream.ToArray();
                Buffer.BlockCopy(BitConverter.GetBytes(sheetOffset), 0, bytes, checked((int)boundSheetPosition + 4), 4);
                return bytes;
            }

            internal static byte[] CreateFormulaMultiCellArrayWorkbookStream() {
                using var stream = new MemoryStream();
                WriteRecord(stream, 0x0809, new byte[] { 0x00, 0x06, 0x05, 0x00, 0xdb, 0x0b, 0xcc, 0x07 });
                long boundSheetPosition = stream.Position;
                WriteRecord(stream, 0x0085, BuildBoundSheetPayload(0, "ArrayFormula"));
                WriteRecord(stream, 0x000a, Array.Empty<byte>());

                int sheetOffset = checked((int)stream.Position);
                WriteRecord(stream, 0x0809, new byte[] { 0x00, 0x06, 0x10, 0x00, 0xdb, 0x0b, 0xcc, 0x07 });
                WriteRecord(stream, 0x0203, BuildNumberPayload(0, 2, 30d));
                WriteRecord(stream, 0x0006, BuildFormulaNumberPayload(0, 3, 31d, formulaTokens: BuildSharedFormulaReferenceTokens(0, 3)));
                WriteRecord(stream, 0x0221, BuildArrayFormulaPayload(0, 3, 0, 4, BuildRelativeReferenceAdditionFormulaTokens(0, -1, 1)));
                WriteRecord(stream, 0x000a, Array.Empty<byte>());

                byte[] bytes = stream.ToArray();
                Buffer.BlockCopy(BitConverter.GetBytes(sheetOffset), 0, bytes, checked((int)boundSheetPosition + 4), 4);
                return bytes;
            }

            internal static byte[] CreateFormulaArrayConstantWorkbookStream() {
                using var stream = new MemoryStream();
                WriteRecord(stream, 0x0809, new byte[] { 0x00, 0x06, 0x05, 0x00, 0xdb, 0x0b, 0xcc, 0x07 });
                long boundSheetPosition = stream.Position;
                WriteRecord(stream, 0x0085, BuildBoundSheetPayload(0, "ArrayConst"));
                WriteRecord(stream, 0x000a, Array.Empty<byte>());

                int sheetOffset = checked((int)stream.Position);
                WriteRecord(stream, 0x0809, new byte[] { 0x00, 0x06, 0x10, 0x00, 0xdb, 0x0b, 0xcc, 0x07 });
                WriteRecord(stream, 0x0006, BuildFormulaNumberPayloadWithExtra(0, 0, 10d, BuildArrayConstantSumFormulaTokens(), BuildNumericArrayConstantExtra(2, 2, 1d, 2d, 3d, 4d)));
                WriteRecord(stream, 0x000a, Array.Empty<byte>());

                byte[] bytes = stream.ToArray();
                Buffer.BlockCopy(BitConverter.GetBytes(sheetOffset), 0, bytes, checked((int)boundSheetPosition + 4), 4);
                return bytes;
            }

            internal static byte[] CreateFormulaDefinedNameWorkbookStream() {
                using var stream = new MemoryStream();
                WriteRecord(stream, 0x0809, new byte[] { 0x00, 0x06, 0x05, 0x00, 0xdb, 0x0b, 0xcc, 0x07 });
                long boundSheetPosition = stream.Position;
                WriteRecord(stream, 0x0085, BuildBoundSheetPayload(0, "NameFormula"));
                WriteRecord(stream, 0x0017, BuildExternSheetPayload((0, 0, 0)));
                WriteRecord(stream, 0x0018, BuildDefinedNamePayload("TaxRate", BuildNameRef3dFormula(0, 0, 2), localSheetIndex: 0, hidden: false, builtIn: false));
                WriteRecord(stream, 0x000a, Array.Empty<byte>());

                int sheetOffset = checked((int)stream.Position);
                WriteRecord(stream, 0x0809, new byte[] { 0x00, 0x06, 0x10, 0x00, 0xdb, 0x0b, 0xcc, 0x07 });
                WriteRecord(stream, 0x0203, BuildNumberPayload(0, 0, 10d));
                WriteRecord(stream, 0x0006, BuildFormulaNumberPayload(0, 1, 2.5d, formulaTokens: BuildDefinedNameMultiplicationFormulaTokens(0, 0, 1)));
                WriteRecord(stream, 0x0203, BuildNumberPayload(0, 2, 0.25d));
                WriteRecord(stream, 0x000a, Array.Empty<byte>());

                byte[] bytes = stream.ToArray();
                Buffer.BlockCopy(BitConverter.GetBytes(sheetOffset), 0, bytes, checked((int)boundSheetPosition + 4), 4);
                return bytes;
            }

            internal static byte[] CreateFormulaInvalidReferenceWorkbookStream() {
                using var stream = new MemoryStream();
                WriteRecord(stream, 0x0809, new byte[] { 0x00, 0x06, 0x05, 0x00, 0xdb, 0x0b, 0xcc, 0x07 });
                long sourceSheetBoundPosition = stream.Position;
                WriteRecord(stream, 0x0085, BuildBoundSheetPayload(0, "Source Data"));
                long formulaSheetBoundPosition = stream.Position;
                WriteRecord(stream, 0x0085, BuildBoundSheetPayload(0, "FormulaRefs"));
                WriteRecord(stream, 0x0017, BuildExternSheetPayload((0, 0, 0)));
                WriteRecord(stream, 0x000a, Array.Empty<byte>());

                int sourceSheetOffset = checked((int)stream.Position);
                WriteRecord(stream, 0x0809, new byte[] { 0x00, 0x06, 0x10, 0x00, 0xdb, 0x0b, 0xcc, 0x07 });
                WriteRecord(stream, 0x0203, BuildNumberPayload(0, 0, 10d));
                WriteRecord(stream, 0x000a, Array.Empty<byte>());

                int formulaSheetOffset = checked((int)stream.Position);
                WriteRecord(stream, 0x0809, new byte[] { 0x00, 0x06, 0x10, 0x00, 0xdb, 0x0b, 0xcc, 0x07 });
                WriteRecord(stream, 0x0006, BuildFormulaErrorResultPayload(0, 0, 0x17, formulaTokens: BuildInvalidReferenceSumFormulaTokens()));
                WriteRecord(stream, 0x0006, BuildFormulaErrorResultPayload(0, 1, 0x17, formulaTokens: BuildInvalidAreaSumFormulaTokens()));
                WriteRecord(stream, 0x0006, BuildFormulaErrorResultPayload(0, 2, 0x17, formulaTokens: BuildInvalid3dReferenceSumFormulaTokens(0)));
                WriteRecord(stream, 0x0006, BuildFormulaErrorResultPayload(0, 3, 0x17, formulaTokens: BuildInvalid3dAreaSumFormulaTokens(0)));
                WriteRecord(stream, 0x000a, Array.Empty<byte>());

                byte[] bytes = stream.ToArray();
                Buffer.BlockCopy(BitConverter.GetBytes(sourceSheetOffset), 0, bytes, checked((int)sourceSheetBoundPosition + 4), 4);
                Buffer.BlockCopy(BitConverter.GetBytes(formulaSheetOffset), 0, bytes, checked((int)formulaSheetBoundPosition + 4), 4);
                return bytes;
            }

            internal static byte[] CreateFormulaErrorConstantWorkbookStream() {
                using var stream = new MemoryStream();
                WriteRecord(stream, 0x0809, new byte[] { 0x00, 0x06, 0x05, 0x00, 0xdb, 0x0b, 0xcc, 0x07 });
                long boundSheetPosition = stream.Position;
                WriteRecord(stream, 0x0085, BuildBoundSheetPayload(0, "ErrorFormula"));
                WriteRecord(stream, 0x000a, Array.Empty<byte>());

                int sheetOffset = checked((int)stream.Position);
                WriteRecord(stream, 0x0809, new byte[] { 0x00, 0x06, 0x10, 0x00, 0xdb, 0x0b, 0xcc, 0x07 });
                WriteRecord(stream, 0x0006, BuildFormulaErrorResultPayload(0, 0, 0x0f, formulaTokens: BuildErrorConstantFormulaTokens(0x0f)));
                WriteRecord(stream, 0x000a, Array.Empty<byte>());

                byte[] bytes = stream.ToArray();
                Buffer.BlockCopy(BitConverter.GetBytes(sheetOffset), 0, bytes, checked((int)boundSheetPosition + 4), 4);
                return bytes;
            }

            internal static byte[] CreateFormulaMissingArgumentWorkbookStream() {
                using var stream = new MemoryStream();
                WriteRecord(stream, 0x0809, new byte[] { 0x00, 0x06, 0x05, 0x00, 0xdb, 0x0b, 0xcc, 0x07 });
                long boundSheetPosition = stream.Position;
                WriteRecord(stream, 0x0085, BuildBoundSheetPayload(0, "MissingArg"));
                WriteRecord(stream, 0x000a, Array.Empty<byte>());

                int sheetOffset = checked((int)stream.Position);
                WriteRecord(stream, 0x0809, new byte[] { 0x00, 0x06, 0x10, 0x00, 0xdb, 0x0b, 0xcc, 0x07 });
                WriteRecord(stream, 0x0203, BuildNumberPayload(0, 0, 10d));
                WriteRecord(stream, 0x0006, BuildFormulaNumberPayload(0, 1, 10d, formulaTokens: BuildMissingArgumentSumFormulaTokens(0, 0)));
                WriteRecord(stream, 0x000a, Array.Empty<byte>());

                byte[] bytes = stream.ToArray();
                Buffer.BlockCopy(BitConverter.GetBytes(sheetOffset), 0, bytes, checked((int)boundSheetPosition + 4), 4);
                return bytes;
            }

            private static byte[] BuildFormulaStringResultPayload(ushort row, ushort column, byte[] formulaTokens) {
                byte[] payload = BuildFormulaPayload(row, column, 0, formulaTokens);
                payload[6] = 0x00;
                WriteUInt16(payload, 12, 0xffff);
                return payload;
            }

            private static byte[] BuildFormulaErrorResultPayload(ushort row, ushort column, byte errorCode, byte[] formulaTokens) {
                byte[] payload = BuildFormulaPayload(row, column, 0, formulaTokens);
                payload[6] = 0x02;
                payload[8] = errorCode;
                WriteUInt16(payload, 12, 0xffff);
                return payload;
            }

            private static byte[] BuildFormulaBooleanPayload(ushort row, ushort column, bool value, byte[] formulaTokens) {
                byte[] payload = BuildFormulaPayload(row, column, 0, formulaTokens);
                payload[6] = 0x01;
                payload[8] = value ? (byte)1 : (byte)0;
                WriteUInt16(payload, 12, 0xffff);
                return payload;
            }

            private static byte[] BuildFormulaNumberPayloadWithExtra(ushort row, ushort column, double value, byte[] formulaTokens, byte[] formulaExtra) {
                byte[] payload = new byte[checked(22 + formulaTokens.Length + formulaExtra.Length)];
                WriteUInt16(payload, 0, row);
                WriteUInt16(payload, 2, column);
                byte[] numberBytes = BitConverter.GetBytes(value);
                Buffer.BlockCopy(numberBytes, 0, payload, 6, numberBytes.Length);
                WriteUInt16(payload, 20, checked((ushort)formulaTokens.Length));
                Buffer.BlockCopy(formulaTokens, 0, payload, 22, formulaTokens.Length);
                Buffer.BlockCopy(formulaExtra, 0, payload, checked(22 + formulaTokens.Length), formulaExtra.Length);
                return payload;
            }

            private static byte[] BuildStringLiteralConcatFormulaTokens(string text, ushort row, ushort column) {
                using var stream = new MemoryStream();
                WriteStringLiteralFormulaToken(stream, text);
                byte[] reference = BuildCellReferenceFormulaToken(row, column);
                stream.Write(reference, 0, reference.Length);
                stream.WriteByte(0x08);
                return stream.ToArray();
            }

            private static void WriteStringLiteralFormulaToken(Stream stream, string text) {
                byte[] textBytes = Encoding.ASCII.GetBytes(text);
                stream.WriteByte(0x17);
                stream.WriteByte(checked((byte)textBytes.Length));
                stream.WriteByte(0);
                stream.Write(textBytes, 0, textBytes.Length);
            }

            private static byte[] Build3dReferenceAdditionFormulaTokens(ushort externSheetIndex, ushort row, ushort column, ushort value) {
                using var stream = new MemoryStream();
                byte[] reference = Build3dCellReferenceFormulaToken(externSheetIndex, row, column);
                stream.Write(reference, 0, reference.Length);
                stream.WriteByte(0x1e);
                WriteUInt16(stream, value);
                stream.WriteByte(0x03);
                return stream.ToArray();
            }

            private static byte[] BuildSum3dAreaFormulaTokens(ushort externSheetIndex, ushort firstRow, ushort firstColumn, ushort lastRow, ushort lastColumn) {
                using var stream = new MemoryStream();
                byte[] area = Build3dAreaReferenceFormulaToken(externSheetIndex, firstRow, firstColumn, lastRow, lastColumn);
                stream.Write(area, 0, area.Length);
                stream.WriteByte(0x42);
                stream.WriteByte(0x01);
                WriteUInt16(stream, 0x0004);
                return stream.ToArray();
            }

            private static byte[] Build3dCellReferenceFormulaToken(ushort externSheetIndex, ushort zeroBasedRow, ushort zeroBasedColumn) {
                byte[] token = new byte[7];
                token[0] = 0x5a;
                WriteUInt16(token, 1, externSheetIndex);
                WriteUInt16(token, 3, zeroBasedRow);
                WriteUInt16(token, 5, (ushort)(zeroBasedColumn | 0xc000));
                return token;
            }

            private static byte[] Build3dAreaReferenceFormulaToken(ushort externSheetIndex, ushort firstRow, ushort firstColumn, ushort lastRow, ushort lastColumn) {
                byte[] token = new byte[11];
                token[0] = 0x5b;
                WriteUInt16(token, 1, externSheetIndex);
                WriteUInt16(token, 3, firstRow);
                WriteUInt16(token, 5, lastRow);
                WriteUInt16(token, 7, (ushort)(firstColumn | 0xc000));
                WriteUInt16(token, 9, (ushort)(lastColumn | 0xc000));
                return token;
            }

            private static byte[] BuildRoundFormulaTokens(ushort row, ushort column, ushort decimals) {
                using var stream = new MemoryStream();
                byte[] reference = BuildCellReferenceFormulaToken(row, column);
                stream.Write(reference, 0, reference.Length);
                stream.WriteByte(0x1e);
                WriteUInt16(stream, decimals);
                stream.WriteByte(0x41);
                WriteUInt16(stream, 0x001b);
                return stream.ToArray();
            }

            private static byte[] BuildLogicalFixedFunctionFormulaTokens(ushort functionId, bool first, bool second) {
                using var stream = new MemoryStream();
                stream.WriteByte(0x1d);
                stream.WriteByte(first ? (byte)1 : (byte)0);
                stream.WriteByte(0x1d);
                stream.WriteByte(second ? (byte)1 : (byte)0);
                stream.WriteByte(0x41);
                WriteUInt16(stream, functionId);
                return stream.ToArray();
            }

            private static byte[] BuildFixedNoArgumentFunctionFormulaTokens(ushort functionId) {
                using var stream = new MemoryStream();
                stream.WriteByte(0x41);
                WriteUInt16(stream, functionId);
                return stream.ToArray();
            }

            private static byte[] BuildRsqFormulaTokens() {
                using var stream = new MemoryStream();
                byte[] knownYValues = BuildAreaReferenceFormulaToken(1, 0, 2, 0);
                byte[] knownXValues = BuildAreaReferenceFormulaToken(1, 1, 2, 1);
                stream.Write(knownYValues, 0, knownYValues.Length);
                stream.Write(knownXValues, 0, knownXValues.Length);
                stream.WriteByte(0x41);
                WriteUInt16(stream, 0x0139);
                return stream.ToArray();
            }

            private static byte[] BuildVLookupFormulaTokens() {
                using var stream = new MemoryStream();
                byte[] lookupValue = BuildCellReferenceFormulaToken(0, 0);
                byte[] tableArray = BuildAreaReferenceFormulaToken(0, 1, 1, 2);
                stream.Write(lookupValue, 0, lookupValue.Length);
                stream.Write(tableArray, 0, tableArray.Length);
                stream.WriteByte(0x1e);
                WriteUInt16(stream, 2);
                stream.WriteByte(0x1d);
                stream.WriteByte(0);
                stream.WriteByte(0x42);
                stream.WriteByte(0x04);
                WriteUInt16(stream, 0x0066);
                return stream.ToArray();
            }

            private static byte[] BuildLeftRightFormulaTokens(ushort functionId, ushort row, ushort column, ushort characterCount) {
                using var stream = new MemoryStream();
                byte[] reference = BuildCellReferenceFormulaToken(row, column);
                stream.Write(reference, 0, reference.Length);
                stream.WriteByte(0x1e);
                WriteUInt16(stream, characterCount);
                stream.WriteByte(0x42);
                stream.WriteByte(0x02);
                WriteUInt16(stream, functionId);
                return stream.ToArray();
            }

            private static byte[] BuildDateFormulaTokens(ushort year, ushort month, ushort day) {
                using var stream = new MemoryStream();
                stream.WriteByte(0x1e);
                WriteUInt16(stream, year);
                stream.WriteByte(0x1e);
                WriteUInt16(stream, month);
                stream.WriteByte(0x1e);
                WriteUInt16(stream, day);
                stream.WriteByte(0x41);
                WriteUInt16(stream, 0x0041);
                return stream.ToArray();
            }

            private static byte[] BuildThreeIntegerFixedFunctionFormulaTokens(ushort functionId, ushort first, ushort second, ushort third) {
                using var stream = new MemoryStream();
                WriteIntegerFormulaToken(stream, first);
                WriteIntegerFormulaToken(stream, second);
                WriteIntegerFormulaToken(stream, third);
                stream.WriteByte(0x41);
                WriteUInt16(stream, functionId);
                return stream.ToArray();
            }

            private static byte[] BuildSingleReferenceFixedFunctionFormulaTokens(ushort functionId, ushort row, ushort column) {
                using var stream = new MemoryStream();
                byte[] reference = BuildCellReferenceFormulaToken(row, column);
                stream.Write(reference, 0, reference.Length);
                stream.WriteByte(0x41);
                WriteUInt16(stream, functionId);
                return stream.ToArray();
            }

            private static byte[] BuildSingleAreaFixedFunctionFormulaTokens(ushort functionId, ushort firstRow, ushort firstColumn, ushort lastRow, ushort lastColumn) {
                using var stream = new MemoryStream();
                byte[] area = BuildAreaReferenceFormulaToken(firstRow, firstColumn, lastRow, lastColumn);
                stream.Write(area, 0, area.Length);
                stream.WriteByte(0x41);
                WriteUInt16(stream, functionId);
                return stream.ToArray();
            }

            private static byte[] BuildLargeFormulaTokens() {
                using var stream = new MemoryStream();
                byte[] area = BuildAreaReferenceFormulaToken(0, 4, 2, 4);
                stream.Write(area, 0, area.Length);
                WriteIntegerFormulaToken(stream, 2);
                stream.WriteByte(0x41);
                WriteUInt16(stream, 0x0145);
                return stream.ToArray();
            }

            private static byte[] BuildStDevFormulaTokens() {
                using var stream = new MemoryStream();
                byte[] area = BuildAreaReferenceFormulaToken(0, 4, 2, 4);
                stream.Write(area, 0, area.Length);
                WriteVariableFunctionCall(stream, 1, 0x000c);
                return stream.ToArray();
            }

            private static byte[] BuildNpvFormulaTokens() {
                using var stream = new MemoryStream();
                WriteIntegerFormulaToken(stream, 10);
                byte[] area = BuildAreaReferenceFormulaToken(0, 4, 2, 4);
                stream.Write(area, 0, area.Length);
                WriteVariableFunctionCall(stream, 2, 0x000b);
                return stream.ToArray();
            }

            private static byte[] BuildSingleNumberFixedFunctionFormulaTokens(ushort functionId, double value) {
                using var stream = new MemoryStream();
                stream.WriteByte(0x1f);
                byte[] numberBytes = BitConverter.GetBytes(value);
                stream.Write(numberBytes, 0, numberBytes.Length);
                stream.WriteByte(0x41);
                WriteUInt16(stream, functionId);
                return stream.ToArray();
            }

            private static void WriteIntegerFormulaToken(Stream stream, ushort value) {
                stream.WriteByte(0x1e);
                WriteUInt16(stream, value);
            }

            private static byte[] BuildCountIfFormulaTokens() {
                using var stream = new MemoryStream();
                byte[] area = BuildAreaReferenceFormulaToken(0, 0, 2, 0);
                stream.Write(area, 0, area.Length);
                WriteStringLiteralFormulaToken(stream, "*IMO");
                stream.WriteByte(0x41);
                WriteUInt16(stream, 0x015a);
                return stream.ToArray();
            }

            private static byte[] BuildSumIfFormulaTokens() {
                using var stream = new MemoryStream();
                byte[] criteriaRange = BuildAreaReferenceFormulaToken(0, 3, 2, 3);
                byte[] sumRange = BuildAreaReferenceFormulaToken(0, 4, 2, 4);
                stream.Write(criteriaRange, 0, criteriaRange.Length);
                WriteStringLiteralFormulaToken(stream, "EU");
                stream.Write(sumRange, 0, sumRange.Length);
                WriteVariableFunctionCall(stream, 3, 0x0159);
                return stream.ToArray();
            }

            private static byte[] BuildSubtotalFormulaTokens() {
                using var stream = new MemoryStream();
                byte[] area = BuildAreaReferenceFormulaToken(0, 4, 2, 4);
                stream.WriteByte(0x1e);
                WriteUInt16(stream, 9);
                stream.Write(area, 0, area.Length);
                WriteVariableFunctionCall(stream, 2, 0x0158);
                return stream.ToArray();
            }

            private static byte[] BuildIndexFormulaTokens() {
                using var stream = new MemoryStream();
                byte[] area = BuildAreaReferenceFormulaToken(0, 4, 2, 4);
                stream.Write(area, 0, area.Length);
                stream.WriteByte(0x1e);
                WriteUInt16(stream, 2);
                WriteVariableFunctionCall(stream, 2, 0x001d);
                return stream.ToArray();
            }

            private static byte[] BuildMatchFormulaTokens() {
                using var stream = new MemoryStream();
                byte[] area = BuildAreaReferenceFormulaToken(0, 3, 2, 3);
                WriteStringLiteralFormulaToken(stream, "US");
                stream.Write(area, 0, area.Length);
                stream.WriteByte(0x1e);
                WriteUInt16(stream, 0);
                WriteVariableFunctionCall(stream, 3, 0x0040);
                return stream.ToArray();
            }

            private static byte[] BuildOffsetFormulaTokens() {
                using var stream = new MemoryStream();
                WriteVolatileAttribute(stream);
                byte[] reference = BuildCellReferenceFormulaToken(0, 4);
                stream.Write(reference, 0, reference.Length);
                stream.WriteByte(0x1e);
                WriteUInt16(stream, 1);
                stream.WriteByte(0x1e);
                WriteUInt16(stream, 0);
                WriteVariableFunctionCall(stream, 3, 0x004e);
                return stream.ToArray();
            }

            private static byte[] BuildVolatileFixedNoArgumentFunctionFormulaTokens(ushort functionId) {
                using var stream = new MemoryStream();
                WriteVolatileAttribute(stream);
                stream.WriteByte(0x41);
                WriteUInt16(stream, functionId);
                return stream.ToArray();
            }

            private static byte[] BuildErrorConstantFormulaTokens(byte errorCode) {
                return new byte[] { 0x1c, errorCode };
            }

            private static byte[] BuildMissingArgumentSumFormulaTokens(ushort row, ushort column) {
                using var stream = new MemoryStream();
                stream.WriteByte(0x16);
                byte[] reference = BuildCellReferenceFormulaToken(row, column);
                stream.Write(reference, 0, reference.Length);
                stream.WriteByte(0x42);
                stream.WriteByte(0x02);
                WriteUInt16(stream, 0x0004);
                return stream.ToArray();
            }

            private static byte[] BuildDisplaySpaceAdditionFormulaTokens(byte spaceAttribute, ushort row, ushort column, ushort value) {
                using var stream = new MemoryStream();
                stream.WriteByte(0x19);
                stream.WriteByte(spaceAttribute);
                stream.WriteByte(0x00);
                stream.WriteByte(0x01);
                byte[] reference = BuildCellReferenceFormulaToken(row, column);
                stream.Write(reference, 0, reference.Length);
                stream.WriteByte(0x1e);
                WriteUInt16(stream, value);
                stream.WriteByte(0x03);
                return stream.ToArray();
            }

            private static byte[] BuildIfFormulaTokens(ushort row, ushort column, ushort comparisonValue, ushort trueValue, ushort falseValue) {
                using var stream = new MemoryStream();
                byte[] reference = BuildCellReferenceFormulaToken(row, column);
                stream.Write(reference, 0, reference.Length);
                stream.WriteByte(0x1e);
                WriteUInt16(stream, comparisonValue);
                stream.WriteByte(0x0d);
                stream.WriteByte(0x19);
                stream.WriteByte(0x02);
                WriteUInt16(stream, 6);
                stream.WriteByte(0x1e);
                WriteUInt16(stream, trueValue);
                stream.WriteByte(0x19);
                stream.WriteByte(0x08);
                WriteUInt16(stream, 3);
                stream.WriteByte(0x1e);
                WriteUInt16(stream, falseValue);
                stream.WriteByte(0x42);
                stream.WriteByte(0x03);
                WriteUInt16(stream, 0x0001);
                return stream.ToArray();
            }

            private static byte[] BuildChooseFormulaTokens(ushort row, ushort column, ushort firstValue, ushort secondValue, ushort thirdValue) {
                using var stream = new MemoryStream();
                byte[] reference = BuildCellReferenceFormulaToken(row, column);
                stream.Write(reference, 0, reference.Length);
                stream.WriteByte(0x19);
                stream.WriteByte(0x04);
                WriteUInt16(stream, 3);
                WriteUInt16(stream, 8);
                WriteUInt16(stream, 7);
                WriteUInt16(stream, 14);
                WriteUInt16(stream, 21);
                WriteChooseOption(stream, firstValue);
                WriteChooseOption(stream, secondValue);
                WriteChooseOption(stream, thirdValue);
                stream.WriteByte(0x42);
                stream.WriteByte(0x04);
                WriteUInt16(stream, 0x0064);
                return stream.ToArray();
            }

            private static void WriteChooseOption(Stream stream, ushort value) {
                stream.WriteByte(0x1e);
                WriteUInt16(stream, value);
                stream.WriteByte(0x19);
                stream.WriteByte(0x08);
                WriteUInt16(stream, 0);
            }

            private static byte[] BuildAttributeSumFormulaTokens(ushort firstRow, ushort firstColumn, ushort lastRow, ushort lastColumn) {
                using var stream = new MemoryStream();
                byte[] area = BuildAreaReferenceFormulaToken(firstRow, firstColumn, lastRow, lastColumn);
                stream.Write(area, 0, area.Length);
                stream.WriteByte(0x19);
                stream.WriteByte(0x10);
                WriteUInt16(stream, 0);
                return stream.ToArray();
            }

            private static byte[] BuildVarFormulaTokens() {
                using var stream = new MemoryStream();
                byte[] area = BuildAreaReferenceFormulaToken(0, 0, 1, 0);
                stream.Write(area, 0, area.Length);
                WriteVariableFunctionCall(stream, 1, 0x002e);
                return stream.ToArray();
            }

            private static byte[] BuildRangeOperatorSumFormulaTokens() {
                using var stream = new MemoryStream();
                byte[] left = BuildCellReferenceFormulaToken(0, 0);
                byte[] right = BuildCellReferenceFormulaToken(0, 1);
                stream.Write(left, 0, left.Length);
                stream.Write(right, 0, right.Length);
                stream.WriteByte(0x11);
                stream.WriteByte(0x42);
                stream.WriteByte(0x01);
                WriteUInt16(stream, 0x0004);
                return stream.ToArray();
            }

            private static byte[] BuildUnionOperatorSumFormulaTokens() {
                using var stream = new MemoryStream();
                byte[] left = BuildCellReferenceFormulaToken(0, 0);
                byte[] right = BuildCellReferenceFormulaToken(0, 1);
                stream.Write(left, 0, left.Length);
                stream.Write(right, 0, right.Length);
                stream.WriteByte(0x10);
                stream.WriteByte(0x42);
                stream.WriteByte(0x01);
                WriteUInt16(stream, 0x0004);
                return stream.ToArray();
            }

            private static byte[] BuildIntersectionOperatorSumFormulaTokens() {
                using var stream = new MemoryStream();
                byte[] area = BuildAreaReferenceFormulaToken(0, 0, 0, 2);
                byte[] reference = BuildCellReferenceFormulaToken(0, 1);
                stream.Write(area, 0, area.Length);
                stream.Write(reference, 0, reference.Length);
                stream.WriteByte(0x0f);
                stream.WriteByte(0x42);
                stream.WriteByte(0x01);
                WriteUInt16(stream, 0x0004);
                return stream.ToArray();
            }

            private static byte[] BuildMemAreaSumFormulaTokens() {
                using var stream = new MemoryStream();
                byte[] area = BuildAreaReferenceFormulaToken(0, 0, 1, 0);
                stream.WriteByte(0x26);
                WriteUInt32(stream, 0);
                WriteUInt16(stream, checked((ushort)area.Length));
                stream.Write(area, 0, area.Length);
                WriteSumFunctionCall(stream);
                return stream.ToArray();
            }

            private static byte[] BuildMemFuncSumFormulaTokens() {
                using var stream = new MemoryStream();
                byte[] left = BuildCellReferenceFormulaToken(0, 0);
                byte[] right = BuildCellReferenceFormulaToken(0, 1);
                ushort expressionLength = checked((ushort)(left.Length + right.Length + 1));
                stream.WriteByte(0x29);
                WriteUInt16(stream, expressionLength);
                stream.Write(left, 0, left.Length);
                stream.Write(right, 0, right.Length);
                stream.WriteByte(0x10);
                WriteSumFunctionCall(stream);
                return stream.ToArray();
            }

            private static byte[] BuildRelativeReferenceAdditionFormulaTokens(short rowOffset, short columnOffset, ushort value) {
                using var stream = new MemoryStream();
                byte[] reference = BuildRelativeCellReferenceFormulaToken(rowOffset, columnOffset);
                stream.Write(reference, 0, reference.Length);
                stream.WriteByte(0x1e);
                WriteUInt16(stream, value);
                stream.WriteByte(0x03);
                return stream.ToArray();
            }

            private static byte[] BuildSharedFormulaReferenceTokens(ushort anchorRow, ushort anchorColumn) {
                byte[] token = new byte[5];
                token[0] = 0x01;
                WriteUInt16(token, 1, anchorRow);
                WriteUInt16(token, 3, anchorColumn);
                return token;
            }

            private static byte[] BuildSharedFormulaPayload(ushort firstRow, ushort firstColumn, ushort lastRow, ushort lastColumn, byte[] formulaTokens) {
                using var stream = new MemoryStream();
                WriteUInt16(stream, firstRow);
                WriteUInt16(stream, lastRow);
                stream.WriteByte(checked((byte)firstColumn));
                stream.WriteByte(checked((byte)lastColumn));
                stream.WriteByte(0);
                stream.WriteByte(checked((byte)(lastRow - firstRow + 1)));
                WriteUInt16(stream, checked((ushort)formulaTokens.Length));
                stream.Write(formulaTokens, 0, formulaTokens.Length);
                return stream.ToArray();
            }

            private static byte[] BuildArrayFormulaPayload(ushort firstRow, ushort firstColumn, ushort lastRow, ushort lastColumn, byte[] formulaTokens) {
                using var stream = new MemoryStream();
                WriteUInt16(stream, firstRow);
                WriteUInt16(stream, lastRow);
                stream.WriteByte(checked((byte)firstColumn));
                stream.WriteByte(checked((byte)lastColumn));
                WriteUInt16(stream, 0);
                WriteUInt32(stream, 0);
                WriteUInt16(stream, checked((ushort)formulaTokens.Length));
                stream.Write(formulaTokens, 0, formulaTokens.Length);
                return stream.ToArray();
            }

            private static byte[] BuildArrayConstantSumFormulaTokens() {
                using var stream = new MemoryStream();
                stream.WriteByte(0x60);
                for (int i = 0; i < 7; i++) {
                    stream.WriteByte(0);
                }

                WriteSumFunctionCall(stream);
                return stream.ToArray();
            }

            private static byte[] BuildNumericArrayConstantExtra(byte columnCount, ushort rowCount, params double[] values) {
                if (values.Length != columnCount * rowCount) {
                    throw new ArgumentException("Array constant value count must match its dimensions.", nameof(values));
                }

                using var stream = new MemoryStream();
                stream.WriteByte(checked((byte)(columnCount - 1)));
                WriteUInt16(stream, checked((ushort)(rowCount - 1)));
                foreach (double value in values) {
                    stream.WriteByte(0x01);
                    byte[] numberBytes = BitConverter.GetBytes(value);
                    stream.Write(numberBytes, 0, numberBytes.Length);
                }

                return stream.ToArray();
            }

            private static byte[] BuildRelativeAreaSumFormulaTokens(short firstRowOffset, short firstColumnOffset, short lastRowOffset, short lastColumnOffset) {
                using var stream = new MemoryStream();
                byte[] area = BuildRelativeAreaReferenceFormulaToken(firstRowOffset, firstColumnOffset, lastRowOffset, lastColumnOffset);
                stream.Write(area, 0, area.Length);
                WriteSumFunctionCall(stream);
                return stream.ToArray();
            }

            private static byte[] BuildRelativeCellReferenceFormulaToken(short rowOffset, short columnOffset) {
                byte[] token = new byte[5];
                token[0] = 0x4c;
                WriteUInt16(token, 1, unchecked((ushort)rowOffset));
                WriteUInt16(token, 3, EncodeRelativeColumn(columnOffset));
                return token;
            }

            private static byte[] BuildRelativeAreaReferenceFormulaToken(short firstRowOffset, short firstColumnOffset, short lastRowOffset, short lastColumnOffset) {
                byte[] token = new byte[9];
                token[0] = 0x4d;
                WriteUInt16(token, 1, unchecked((ushort)firstRowOffset));
                WriteUInt16(token, 3, unchecked((ushort)lastRowOffset));
                WriteUInt16(token, 5, EncodeRelativeColumn(firstColumnOffset));
                WriteUInt16(token, 7, EncodeRelativeColumn(lastColumnOffset));
                return token;
            }

            private static ushort EncodeRelativeColumn(short columnOffset) {
                return unchecked((ushort)(((ushort)columnOffset & 0x3fff) | 0xc000));
            }

            private static byte[] BuildDefinedNameMultiplicationFormulaTokens(ushort row, ushort column, uint oneBasedNameIndex) {
                using var stream = new MemoryStream();
                byte[] reference = BuildCellReferenceFormulaToken(row, column);
                stream.Write(reference, 0, reference.Length);
                stream.WriteByte(0x43);
                WriteUInt32(stream, oneBasedNameIndex);
                stream.WriteByte(0x05);
                return stream.ToArray();
            }

            private static byte[] BuildExternalDefinedNameFormulaTokens(ushort externSheetIndex, uint oneBasedNameIndex) {
                using var stream = new MemoryStream();
                stream.WriteByte(0x39);
                WriteUInt16(stream, externSheetIndex);
                WriteUInt32(stream, oneBasedNameIndex);
                return stream.ToArray();
            }

            private static byte[] BuildAddInUserDefinedFunctionFormulaTokens(ushort externSheetIndex, uint oneBasedNameIndex, ushort argument) {
                using var stream = new MemoryStream();
                stream.WriteByte(0x39);
                WriteUInt16(stream, externSheetIndex);
                WriteUInt32(stream, oneBasedNameIndex);
                stream.WriteByte(0x1e);
                WriteUInt16(stream, argument);
                stream.WriteByte(0x42);
                stream.WriteByte(0x02);
                WriteUInt16(stream, 0x00ff);
                return stream.ToArray();
            }

            private static byte[] BuildFunctionStackUnderflowFormulaTokens() {
                using var stream = new MemoryStream();
                WriteVariableFunctionCall(stream, 1, 0x0004);
                return stream.ToArray();
            }

            private static byte[] BuildExternalNamePayload(
                string name,
                ushort oneBasedSheetIndex = 0,
                bool builtIn = false,
                bool wantsAdvise = false,
                bool wantsPicture = false,
                bool ole = false,
                bool oleLink = false,
                int cachedClipboardFormat = 0,
                bool icon = false) {
                using var stream = new MemoryStream();
                WriteUInt16(stream, BuildExternalNameFlags(
                    builtIn,
                    wantsAdvise,
                    wantsPicture,
                    ole,
                    oleLink,
                    cachedClipboardFormat,
                    icon));
                WriteUInt16(stream, oneBasedSheetIndex);
                WriteUInt16(stream, 0);
                stream.WriteByte(checked((byte)name.Length));
                WriteCompressedUnicodeStringNoCch(stream, name);
                WriteUInt16(stream, 0);
                return stream.ToArray();
            }

            private static ushort BuildExternalNameFlags(
                bool builtIn,
                bool wantsAdvise,
                bool wantsPicture,
                bool ole,
                bool oleLink,
                int cachedClipboardFormat,
                bool icon) {
                if (cachedClipboardFormat < -512 || cachedClipboardFormat > 511) {
                    throw new ArgumentOutOfRangeException(nameof(cachedClipboardFormat));
                }

                ushort flags = 0;
                if (builtIn) flags |= 0x0001;
                if (wantsAdvise) flags |= 0x0002;
                if (wantsPicture) flags |= 0x0004;
                if (ole) flags |= 0x0008;
                if (oleLink) flags |= 0x0010;
                flags |= checked((ushort)((cachedClipboardFormat & 0x03ff) << 5));
                if (icon) flags |= 0x8000;
                return flags;
            }

            private static byte[] BuildAddInExternalNamePayload(string name) {
                using var stream = new MemoryStream();
                WriteUInt16(stream, 0);
                WriteUInt32(stream, 0);
                stream.WriteByte(checked((byte)name.Length));
                WriteCompressedUnicodeStringNoCch(stream, name);
                WriteUInt16(stream, 0);
                return stream.ToArray();
            }

            private static byte[] BuildSupBookAddInPayload() {
                using var stream = new MemoryStream();
                WriteUInt16(stream, 1);
                WriteUInt16(stream, 0x3a01);
                return stream.ToArray();
            }

            private static byte[] BuildInvalidReferenceSumFormulaTokens() {
                using var stream = new MemoryStream();
                stream.WriteByte(0x4a);
                WriteUInt32(stream, 0);
                WriteSumFunctionCall(stream);
                return stream.ToArray();
            }

            private static byte[] BuildInvalidAreaSumFormulaTokens() {
                using var stream = new MemoryStream();
                stream.WriteByte(0x4b);
                WriteUInt32(stream, 0);
                WriteUInt32(stream, 0);
                WriteSumFunctionCall(stream);
                return stream.ToArray();
            }

            private static byte[] BuildInvalid3dReferenceSumFormulaTokens(ushort externSheetIndex) {
                using var stream = new MemoryStream();
                stream.WriteByte(0x5c);
                WriteUInt16(stream, externSheetIndex);
                WriteUInt32(stream, 0);
                WriteSumFunctionCall(stream);
                return stream.ToArray();
            }

            private static byte[] BuildInvalid3dAreaSumFormulaTokens(ushort externSheetIndex) {
                using var stream = new MemoryStream();
                stream.WriteByte(0x5d);
                WriteUInt16(stream, externSheetIndex);
                WriteUInt32(stream, 0);
                WriteUInt32(stream, 0);
                WriteSumFunctionCall(stream);
                return stream.ToArray();
            }

            private static void WriteSumFunctionCall(Stream stream) {
                WriteVariableFunctionCall(stream, 1, 0x0004);
            }

            private static void WriteVariableFunctionCall(Stream stream, byte argumentCount, ushort functionId) {
                stream.WriteByte(0x42);
                stream.WriteByte(argumentCount);
                WriteUInt16(stream, functionId);
            }

            private static void WriteVolatileAttribute(Stream stream) {
                stream.WriteByte(0x19);
                stream.WriteByte(0x01);
                WriteUInt16(stream, 0);
            }
        }
    }
}
