using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using OfficeIMO.Excel;
using OfficeIMO.Excel.LegacyXls;
using OfficeIMO.Excel.LegacyXls.Biff;
using OfficeIMO.Excel.LegacyXls.Diagnostics;
using OfficeIMO.Excel.LegacyXls.Model;
using System.Globalization;
using Xunit;

namespace OfficeIMO.Tests {
    public partial class Excel {
        [Fact]
        public void LegacyXls_BiffRecordReader_ReadsRecordsAndReportsTruncation() {
            var diagnostics = new List<LegacyXlsImportDiagnostic>();
            byte[] bytes = {
                0x09, 0x08, 0x02, 0x00, 0x01, 0x02,
                0x03, 0x02, 0x04, 0x00, 0x01
            };

            IReadOnlyList<BiffRecord> records = BiffRecordReader.ReadRecords(bytes, diagnostics);

            Assert.Single(records);
            Assert.Equal((ushort)BiffRecordType.Bof, records[0].Type);
            Assert.Equal(new byte[] { 0x01, 0x02 }, records[0].Payload);
            Assert.Contains(diagnostics, item => item.Code == "XLS-BIFF-TRUNCATED-PAYLOAD");
        }

        [Fact]
        public void LegacyXls_Load_ImportsMinimalBiff8Workbook() {
            byte[] workbookStream = LegacyXlsTestWorkbookBuilder.CreateMinimalWorkbookStream();
            byte[] compound = LegacyXlsCompoundTestBuilder.CreateWorkbookCompoundFile(workbookStream);

            LegacyXlsWorkbook legacy = LegacyXlsWorkbook.Load(compound, new LegacyXlsImportOptions {
                ReportUnsupportedRecords = false
            });

            Assert.DoesNotContain(legacy.Diagnostics, d => d.Severity == LegacyXlsDiagnosticSeverity.Error);
            LegacyXlsWorksheet sheet = Assert.Single(legacy.Worksheets);
            Assert.Equal("Sheet1", sheet.Name);
            Assert.Contains(sheet.Cells, cell => cell.Row == 1 && cell.Column == 1 && Equals(cell.Value, "Name"));
            Assert.Contains(sheet.Cells, cell => cell.Row == 2 && cell.Column == 2 && Equals(cell.Value, 42d));
            Assert.Contains(sheet.Cells, cell => cell.Row == 3 && cell.Column == 1 && Equals(cell.Value, true));
        }

        [Fact]
        public void LegacyXls_Load_ImportsWorkbookStreamFromCompoundMiniStream() {
            byte[] workbookStream = LegacyXlsTestWorkbookBuilder.CreateMinimalWorkbookStream();
            Assert.True(workbookStream.Length < 4096);
            byte[] compound = LegacyXlsCompoundTestBuilder.CreateMiniStreamWorkbookCompoundFile(workbookStream);

            LegacyXlsWorkbook legacy = LegacyXlsWorkbook.Load(compound, new LegacyXlsImportOptions {
                ReportUnsupportedRecords = true
            });

            Assert.DoesNotContain(legacy.Diagnostics, d => d.Severity == LegacyXlsDiagnosticSeverity.Error);
            LegacyXlsWorksheet sheet = Assert.Single(legacy.Worksheets);
            Assert.Equal("Sheet1", sheet.Name);
            Assert.Contains(sheet.Cells, cell => cell.Row == 1 && cell.Column == 1 && Equals(cell.Value, "Name"));

            using ExcelDocument document = ExcelDocument.LoadLegacyXls(new MemoryStream(compound), new LegacyXlsImportOptions {
                ReportUnsupportedRecords = true
            });

            Assert.True(document.Sheets[0].TryGetCellText(2, 2, out string? amount));
            Assert.Equal("42", amount);
        }

        [Fact]
        public void LegacyXls_Load_ImportsWorkbookStreamFromCompoundDifatSector() {
            byte[] workbookStream = LegacyXlsTestWorkbookBuilder.CreateMinimalWorkbookStream();
            byte[] compound = LegacyXlsCompoundTestBuilder.CreateDifatWorkbookCompoundFile(workbookStream);

            LegacyXlsWorkbook legacy = LegacyXlsWorkbook.Load(compound, new LegacyXlsImportOptions {
                ReportUnsupportedRecords = false
            });

            Assert.DoesNotContain(legacy.Diagnostics, d => d.Severity == LegacyXlsDiagnosticSeverity.Error);
            LegacyXlsWorksheet sheet = Assert.Single(legacy.Worksheets);
            Assert.Equal("Sheet1", sheet.Name);
            Assert.Contains(sheet.Cells, cell => cell.Row == 2 && cell.Column == 2 && Equals(cell.Value, 42d));

            using ExcelDocument document = ExcelDocument.LoadLegacyXls(new MemoryStream(compound), new LegacyXlsImportOptions {
                ReportUnsupportedRecords = false
            });

            Assert.True(document.Sheets[0].TryGetCellText(1, 1, out string? header));
            Assert.Equal("Name", header);
        }

        [Fact]
        public void LegacyXls_Load_ImportsPhase2ValueRecords() {
            byte[] workbookStream = LegacyXlsTestWorkbookBuilder.CreatePhase2ValueWorkbookStream();
            byte[] compound = LegacyXlsCompoundTestBuilder.CreateWorkbookCompoundFile(workbookStream);

            LegacyXlsWorkbook legacy = LegacyXlsWorkbook.Load(compound, new LegacyXlsImportOptions {
                ReportUnsupportedRecords = false
            });

            Assert.DoesNotContain(legacy.Diagnostics, d => d.Severity == LegacyXlsDiagnosticSeverity.Error);
            Assert.True(legacy.Uses1904DateSystem);
            LegacyXlsWorksheet sheet = Assert.Single(legacy.Worksheets);
            Assert.Contains(sheet.Cells, cell => cell.Row == 1 && cell.Column == 1 && Equals(cell.Value, "Inline"));
            Assert.Contains(sheet.Cells, cell => cell.Row == 2 && cell.Column == 1 && Equals(cell.Value, 7d));
            Assert.Contains(sheet.Cells, cell => cell.Row == 2 && cell.Column == 2 && Equals(cell.Value, -3d));
            Assert.Contains(sheet.Cells, cell => cell.Row == 2 && cell.Column == 3 && Equals(cell.Value, 123.45d));
            Assert.Contains(sheet.Cells, cell => cell.Row == 3 && cell.Column == 1 && Equals(cell.Value, 1d));
            Assert.Contains(sheet.Cells, cell => cell.Row == 3 && cell.Column == 2 && Equals(cell.Value, 2d));
            Assert.Contains(sheet.Cells, cell => cell.Row == 4 && cell.Column == 1 && cell.Kind == LegacyXlsCellValueKind.Blank && cell.StyleIndex == 5);
            Assert.Contains(sheet.Cells, cell => cell.Row == 4 && cell.Column == 2 && cell.Kind == LegacyXlsCellValueKind.Blank && cell.StyleIndex == 6);
            Assert.Contains(sheet.Cells, cell => cell.Row == 5 && cell.Column == 1 && cell.Kind == LegacyXlsCellValueKind.Error && Equals(cell.Value, "#DIV/0!"));
        }

        [Fact]
        public void LegacyXls_LoadLegacyXls_Projects1904DateSerialAsOpenXmlDate() {
            byte[] workbookStream = LegacyXlsTestWorkbookBuilder.CreateDate1904DateFormattedWorkbookStream();
            byte[] compound = LegacyXlsCompoundTestBuilder.CreateWorkbookCompoundFile(workbookStream);

            LegacyXlsWorkbook legacy = LegacyXlsWorkbook.Load(compound, new LegacyXlsImportOptions {
                ReportUnsupportedRecords = true
            });

            Assert.DoesNotContain(legacy.Diagnostics, d => d.Severity == LegacyXlsDiagnosticSeverity.Error);
            Assert.True(legacy.Uses1904DateSystem);
            LegacyXlsCell dateCellModel = Assert.Single(Assert.Single(legacy.Worksheets).Cells);
            Assert.Equal(1d, dateCellModel.Value);
            Assert.True(legacy.GetEffectiveCellFormat(dateCellModel.StyleIndex)!.IsDateLike);

            using ExcelDocument document = ExcelDocument.LoadLegacyXls(new MemoryStream(compound), new LegacyXlsImportOptions {
                ReportUnsupportedRecords = true
            });

            using var output = new MemoryStream();
            document.Save(output);
            using SpreadsheetDocument spreadsheet = SpreadsheetDocument.Open(new MemoryStream(output.ToArray()), false);
            WorkbookPart workbookPart = spreadsheet.WorkbookPart!;
            WorksheetPart worksheetPart = workbookPart.WorksheetParts.Single();
            Cell projectedDateCell = worksheetPart.Worksheet.Descendants<Cell>().Single(cell => cell.CellReference!.Value == "A1");
            Assert.Equal(new DateTime(1904, 1, 2).ToOADate().ToString(CultureInfo.InvariantCulture), projectedDateCell.CellValue!.Text);
            Assert.NotNull(projectedDateCell.StyleIndex);
            CellFormat dateFormat = workbookPart.WorkbookStylesPart!.Stylesheet.CellFormats!.Elements<CellFormat>().ElementAt((int)projectedDateCell.StyleIndex!.Value);
            uint dateNumberFormatId = dateFormat.NumberFormatId!.Value;
            NumberingFormat projectedNumberFormat = Assert.Single(workbookPart.WorkbookStylesPart.Stylesheet.NumberingFormats!.Elements<NumberingFormat>(),
                format => format.NumberFormatId!.Value == dateNumberFormatId);
            Assert.Equal("yyyy-mm-dd", projectedNumberFormat.FormatCode!.Value);
        }

        [Fact]
        public void LegacyXls_Load_ImportsContinuedSharedStringTableEntries() {
            byte[] workbookStream = LegacyXlsTestWorkbookBuilder.CreateContinuedSharedStringWorkbookStream();
            byte[] compound = LegacyXlsCompoundTestBuilder.CreateWorkbookCompoundFile(workbookStream);

            LegacyXlsWorkbook legacy = LegacyXlsWorkbook.Load(compound, new LegacyXlsImportOptions {
                ReportUnsupportedRecords = true
            });

            Assert.DoesNotContain(legacy.Diagnostics, d => d.Severity == LegacyXlsDiagnosticSeverity.Error);
            Assert.DoesNotContain(legacy.Diagnostics, d => d.RecordType == (ushort)BiffRecordType.Continue);
            LegacyXlsWorksheet sheet = Assert.Single(legacy.Worksheets);
            Assert.Contains(sheet.Cells, cell => cell.Row == 1 && cell.Column == 1 && Equals(cell.Value, "First"));
            Assert.Contains(sheet.Cells, cell => cell.Row == 1 && cell.Column == 2 && Equals(cell.Value, "Second"));

            using ExcelDocument document = ExcelDocument.LoadLegacyXls(new MemoryStream(compound), new LegacyXlsImportOptions {
                ReportUnsupportedRecords = true
            });

            Assert.True(document.Sheets[0].TryGetCellText(1, 1, out string? first));
            Assert.Equal("First", first);
            Assert.True(document.Sheets[0].TryGetCellText(1, 2, out string? second));
            Assert.Equal("Second", second);
        }

        [Fact]
        public void LegacyXls_Load_ImportsWorksheetDimensions() {
            byte[] workbookStream = LegacyXlsTestWorkbookBuilder.CreateWorksheetDimensionsWorkbookStream();
            byte[] compound = LegacyXlsCompoundTestBuilder.CreateWorkbookCompoundFile(workbookStream);

            LegacyXlsWorkbook legacy = LegacyXlsWorkbook.Load(compound, new LegacyXlsImportOptions {
                ReportUnsupportedRecords = true
            });

            Assert.DoesNotContain(legacy.Diagnostics, d => d.Severity == LegacyXlsDiagnosticSeverity.Error);
            Assert.Equal(2, legacy.Worksheets.Count);

            LegacyXlsWorksheet bounds = legacy.Worksheets[0];
            Assert.NotNull(bounds.DeclaredUsedRange);
            Assert.False(bounds.DeclaredUsedRange!.IsEmpty);
            Assert.Equal(2, bounds.DeclaredUsedRange.FirstRow);
            Assert.Equal(2, bounds.DeclaredUsedRange.FirstColumn);
            Assert.Equal(5, bounds.DeclaredUsedRange.LastRow);
            Assert.Equal(4, bounds.DeclaredUsedRange.LastColumn);
            Assert.Equal("B2:D5", bounds.DeclaredUsedRange.UsedRangeA1);

            LegacyXlsWorksheet empty = legacy.Worksheets[1];
            Assert.NotNull(empty.DeclaredUsedRange);
            Assert.True(empty.DeclaredUsedRange!.IsEmpty);
            Assert.Equal("A1:A1", empty.DeclaredUsedRange.UsedRangeA1);
            Assert.Empty(empty.Cells);

            using ExcelDocument document = ExcelDocument.LoadLegacyXls(new MemoryStream(compound), new LegacyXlsImportOptions {
                ReportUnsupportedRecords = true
            });

            ExcelWorkbookSnapshot snapshot = document.CreateInspectionSnapshot();
            Assert.Equal("B2:D5", snapshot.Worksheets[0].UsedRangeA1);
            Assert.Equal("A1:A1", snapshot.Worksheets[1].UsedRangeA1);
        }

        [Fact]
        public void LegacyXls_Load_ImportsPhase4FormulaCachedResults() {
            byte[] workbookStream = LegacyXlsTestWorkbookBuilder.CreatePhase4FormulaWorkbookStream();
            byte[] compound = LegacyXlsCompoundTestBuilder.CreateWorkbookCompoundFile(workbookStream);

            LegacyXlsWorkbook legacy = LegacyXlsWorkbook.Load(compound, new LegacyXlsImportOptions {
                ReportUnsupportedRecords = false
            });

            Assert.DoesNotContain(legacy.Diagnostics, d => d.Severity == LegacyXlsDiagnosticSeverity.Error);
            LegacyXlsWorksheet sheet = Assert.Single(legacy.Worksheets);
            Assert.Contains(sheet.Cells, cell => cell.Row == 1 && cell.Column == 1 && cell.IsFormula && Equals(cell.Value, 42d));
            Assert.Contains(sheet.Cells, cell => cell.Row == 1 && cell.Column == 2 && cell.IsFormula && Equals(cell.Value, true));
            Assert.Contains(sheet.Cells, cell => cell.Row == 1 && cell.Column == 3 && cell.IsFormula && Equals(cell.Value, "Formula text"));
            Assert.Contains(sheet.Cells, cell => cell.Row == 1 && cell.Column == 4 && cell.IsFormula && cell.Kind == LegacyXlsCellValueKind.Error && Equals(cell.Value, "#VALUE!"));
            LegacyXlsCell decodedFormula = Assert.Single(sheet.Cells, cell => cell.Row == 2 && cell.Column == 3);
            Assert.True(decodedFormula.IsFormula);
            Assert.Equal(42d, decodedFormula.Value);
            Assert.Equal("A2+B2", decodedFormula.FormulaText);
            LegacyXlsCell decodedSumFormula = Assert.Single(sheet.Cells, cell => cell.Row == 2 && cell.Column == 4);
            Assert.True(decodedSumFormula.IsFormula);
            Assert.Equal(42d, decodedSumFormula.Value);
            Assert.Equal("SUM(A2:B2)", decodedSumFormula.FormulaText);

            using ExcelDocument document = ExcelDocument.LoadLegacyXls(new MemoryStream(compound), new LegacyXlsImportOptions {
                ReportUnsupportedRecords = false
            });

            Assert.True(document.Sheets[0].TryGetCellText(1, 1, out string? number));
            Assert.Equal("42", number);
            Assert.True(document.Sheets[0].TryGetCellText(1, 2, out string? boolean));
            Assert.Equal("1", boolean);
            Assert.True(document.Sheets[0].TryGetCellText(1, 3, out string? text));
            Assert.Equal("Formula text", text);
            Assert.True(document.Sheets[0].TryGetCellText(2, 3, out string? formulaCachedValue));
            Assert.Equal("42", formulaCachedValue);
            Assert.True(document.Sheets[0].TryGetCellText(2, 4, out string? sumFormulaCachedValue));
            Assert.Equal("42", sumFormulaCachedValue);

            using var output = new MemoryStream();
            document.Save(output);
            using SpreadsheetDocument spreadsheet = SpreadsheetDocument.Open(new MemoryStream(output.ToArray()), false);
            WorksheetPart worksheetPart = spreadsheet.WorkbookPart!.WorksheetParts.Single();
            Cell projectedFormula = worksheetPart.Worksheet.Descendants<Cell>().Single(cell => cell.CellReference!.Value == "C2");
            Assert.Equal("A2+B2", projectedFormula.CellFormula!.Text);
            Assert.Equal("42", projectedFormula.CellValue!.Text);
            Cell projectedSumFormula = worksheetPart.Worksheet.Descendants<Cell>().Single(cell => cell.CellReference!.Value == "D2");
            Assert.Equal("SUM(A2:B2)", projectedSumFormula.CellFormula!.Text);
            Assert.Equal("42", projectedSumFormula.CellValue!.Text);
        }

        [Fact]
        public void LegacyXls_Load_ReportsEncryptedWorkbookAsUnsupported() {
            byte[] workbookStream = LegacyXlsTestWorkbookBuilder.CreateEncryptedWorkbookStream();
            byte[] compound = LegacyXlsCompoundTestBuilder.CreateWorkbookCompoundFile(workbookStream);

            LegacyXlsWorkbook legacy = LegacyXlsWorkbook.Load(compound);

            Assert.Contains(legacy.Diagnostics, d => d.Code == "XLS-BIFF-FILEPASS-UNSUPPORTED" && d.Severity == LegacyXlsDiagnosticSeverity.Error);
            Assert.Empty(legacy.Worksheets);
        }

        [Fact]
        public void LegacyXls_Load_ReportsFeatureSpecificUnsupportedRecords() {
            byte[] workbookStream = LegacyXlsTestWorkbookBuilder.CreateUnsupportedFeatureWorkbookStream();
            byte[] compound = LegacyXlsCompoundTestBuilder.CreateWorkbookCompoundFile(workbookStream);

            LegacyXlsWorkbook legacy = LegacyXlsWorkbook.Load(compound, new LegacyXlsImportOptions {
                ReportUnsupportedRecords = true
            });

            Assert.DoesNotContain(legacy.Diagnostics, d => d.Severity == LegacyXlsDiagnosticSeverity.Error);
            Assert.Contains(legacy.Diagnostics, d => d.Code == "XLS-BIFF-FEATURE-HYPERLINK-UNSUPPORTED" && d.SheetName == "Features");
            Assert.Contains(legacy.Diagnostics, d => d.Code == "XLS-BIFF-FEATURE-COMMENT-UNSUPPORTED" && d.SheetName == "Features");
            Assert.Contains(legacy.Diagnostics, d => d.Code == "XLS-BIFF-FEATURE-DRAWING-UNSUPPORTED" && d.SheetName == "Features");
            Assert.Contains(legacy.Diagnostics, d => d.Code == "XLS-BIFF-FEATURE-PIVOT-TABLE-UNSUPPORTED" && d.SheetName == "Features");
            Assert.Contains(legacy.Diagnostics, d => d.Code == "XLS-BIFF-FEATURE-AUTOFILTER-CRITERIA-UNSUPPORTED" && d.SheetName == "Features");
            Assert.Contains(legacy.UnsupportedFeatures, feature => feature.Kind == LegacyXlsUnsupportedFeatureKind.Hyperlink && feature.SheetName == "Features" && feature.RecordType == 0x01b8);
            Assert.Contains(legacy.UnsupportedFeatures, feature => feature.Kind == LegacyXlsUnsupportedFeatureKind.Comment && feature.SheetName == "Features" && feature.RecordType == 0x001c);
            Assert.Contains(legacy.UnsupportedFeatures, feature => feature.Kind == LegacyXlsUnsupportedFeatureKind.DrawingObject && feature.SheetName == "Features" && feature.RecordType == 0x00ec);
            Assert.Contains(legacy.UnsupportedFeatures, feature => feature.Kind == LegacyXlsUnsupportedFeatureKind.PivotTable && feature.SheetName == "Features" && feature.RecordType == 0x00b0);
            Assert.Contains(legacy.UnsupportedFeatures, feature => feature.Kind == LegacyXlsUnsupportedFeatureKind.AutoFilterCriteria && feature.SheetName == "Features" && feature.RecordType == 0x009e);
        }

        [Fact]
        public void LegacyXls_Load_PreservesUnsupportedFeatureMetadataWhenDiagnosticsAreDisabled() {
            byte[] workbookStream = LegacyXlsTestWorkbookBuilder.CreateUnsupportedFeatureWorkbookStream();
            byte[] compound = LegacyXlsCompoundTestBuilder.CreateWorkbookCompoundFile(workbookStream);

            LegacyXlsWorkbook legacy = LegacyXlsWorkbook.Load(compound, new LegacyXlsImportOptions {
                ReportUnsupportedRecords = false
            });

            Assert.DoesNotContain(legacy.Diagnostics, d => d.Code.StartsWith("XLS-BIFF-FEATURE-", StringComparison.Ordinal));
            Assert.Contains(legacy.UnsupportedFeatures, feature => feature.Kind == LegacyXlsUnsupportedFeatureKind.Hyperlink && feature.Code == "XLS-BIFF-FEATURE-HYPERLINK-UNSUPPORTED");
            Assert.Contains(legacy.UnsupportedFeatures, feature => feature.Kind == LegacyXlsUnsupportedFeatureKind.Comment && feature.Code == "XLS-BIFF-FEATURE-COMMENT-UNSUPPORTED");
            Assert.Contains(legacy.UnsupportedFeatures, feature => feature.Kind == LegacyXlsUnsupportedFeatureKind.DrawingObject && feature.Code == "XLS-BIFF-FEATURE-DRAWING-UNSUPPORTED");
            Assert.Contains(legacy.UnsupportedFeatures, feature => feature.Kind == LegacyXlsUnsupportedFeatureKind.PivotTable && feature.Code == "XLS-BIFF-FEATURE-PIVOT-TABLE-UNSUPPORTED");
            Assert.Contains(legacy.UnsupportedFeatures, feature => feature.Kind == LegacyXlsUnsupportedFeatureKind.AutoFilterCriteria && feature.Code == "XLS-BIFF-FEATURE-AUTOFILTER-CRITERIA-UNSUPPORTED");
        }

        [Fact]
        public void LegacyXls_LoadLegacyXls_ProjectsToNormalExcelDocument() {
            byte[] workbookStream = LegacyXlsTestWorkbookBuilder.CreateMinimalWorkbookStream();
            byte[] compound = LegacyXlsCompoundTestBuilder.CreateWorkbookCompoundFile(workbookStream);

            using ExcelDocument document = ExcelDocument.LoadLegacyXls(new MemoryStream(compound), new LegacyXlsImportOptions {
                ReportUnsupportedRecords = false
            });

            Assert.Single(document.Sheets);
            Assert.True(document.Sheets[0].TryGetCellText(1, 1, out string? header));
            Assert.Equal("Name", header);
            Assert.True(document.Sheets[0].TryGetCellText(2, 2, out string? amount));
            Assert.Equal("42", amount);
        }

        [Fact]
        public void LegacyXls_LoadLegacyXlsWithReport_ReturnsProjectedDocumentAndImportReport() {
            byte[] workbookStream = LegacyXlsTestWorkbookBuilder.CreateUnsupportedFeatureWorkbookStream();
            byte[] compound = LegacyXlsCompoundTestBuilder.CreateWorkbookCompoundFile(workbookStream);

            using LegacyXlsLoadResult result = ExcelDocument.LoadLegacyXlsWithReport(new MemoryStream(compound), new LegacyXlsImportOptions {
                ReportUnsupportedRecords = false
            });

            Assert.Single(result.Document.Sheets);
            Assert.True(result.Document.Sheets[0].TryGetCellText(1, 1, out string? value));
            Assert.Equal("Feature", value);
            Assert.Single(result.Workbook.Worksheets);
            Assert.DoesNotContain(result.Diagnostics, d => d.Code.StartsWith("XLS-BIFF-FEATURE-", StringComparison.Ordinal));
            Assert.False(result.HasImportErrors);
            Assert.True(result.HasUnsupportedFeatures);
            Assert.Same(result, result.EnsureNoImportErrors());
            Assert.Contains(result.UnsupportedFeatures, feature => feature.Kind == LegacyXlsUnsupportedFeatureKind.Hyperlink && feature.SheetName == "Features");
            Assert.Contains(result.UnsupportedFeatures, feature => feature.Kind == LegacyXlsUnsupportedFeatureKind.Comment && feature.SheetName == "Features");
            Assert.Contains(result.UnsupportedFeatures, feature => feature.Kind == LegacyXlsUnsupportedFeatureKind.DrawingObject && feature.SheetName == "Features");
            Assert.Contains(result.UnsupportedFeatures, feature => feature.Kind == LegacyXlsUnsupportedFeatureKind.PivotTable && feature.SheetName == "Features");
            Assert.Contains(result.UnsupportedFeatures, feature => feature.Kind == LegacyXlsUnsupportedFeatureKind.AutoFilterCriteria && feature.SheetName == "Features");
            Assert.Equal(1, result.ImportReport.UnsupportedFeaturesByKind[LegacyXlsUnsupportedFeatureKind.DrawingObject]);
            Assert.Equal(1, result.ImportReport.UnsupportedFeaturesByKind[LegacyXlsUnsupportedFeatureKind.AutoFilterCriteria]);
            Assert.Equal(1, result.ImportReport.UnsupportedFeaturesByCode["XLS-BIFF-FEATURE-DRAWING-UNSUPPORTED"]);
            Assert.Equal(1, result.ImportReport.UnsupportedFeaturesByCode["XLS-BIFF-FEATURE-AUTOFILTER-CRITERIA-UNSUPPORTED"]);
            Assert.Same(result.Workbook.Diagnostics, result.Diagnostics);
            Assert.Same(result.Workbook.UnsupportedFeatures, result.UnsupportedFeatures);
            InvalidOperationException exception = Assert.Throws<InvalidOperationException>(() => result.EnsureNoUnsupportedFeatures());
            Assert.Contains("XLS-BIFF-FEATURE-HYPERLINK-UNSUPPORTED", exception.Message, StringComparison.Ordinal);
        }

        [Fact]
        public void LegacyXls_LoadLegacyXls_ImportsAndProjectsPhase3LayoutRecords() {
            byte[] workbookStream = LegacyXlsTestWorkbookBuilder.CreatePhase3LayoutWorkbookStream();
            byte[] compound = LegacyXlsCompoundTestBuilder.CreateWorkbookCompoundFile(workbookStream);

            LegacyXlsWorkbook legacy = LegacyXlsWorkbook.Load(compound, new LegacyXlsImportOptions {
                ReportUnsupportedRecords = false
            });

            Assert.DoesNotContain(legacy.Diagnostics, d => d.Severity == LegacyXlsDiagnosticSeverity.Error);
            Assert.Equal(3, legacy.Worksheets.Count);

            LegacyXlsWorksheet layout = legacy.Worksheets[0];
            LegacyXlsColumnLayout column = Assert.Single(layout.Columns);
            Assert.Equal(2, column.StartColumn);
            Assert.Equal(2, column.EndColumn);
            Assert.Equal(12.5d, column.Width);
            Assert.True(column.Hidden);
            Assert.Equal(2, column.OutlineLevel);
            Assert.True(column.Collapsed);

            LegacyXlsRowLayout row = Assert.Single(layout.Rows);
            Assert.Equal(2, row.Row);
            Assert.Equal(18d, row.Height);
            Assert.True(row.Hidden);
            Assert.True(row.CustomHeight);
            Assert.Equal(1, row.OutlineLevel);
            Assert.True(row.Collapsed);

            LegacyXlsMergedRange mergedRange = Assert.Single(layout.MergedRanges);
            Assert.Equal(1, mergedRange.StartRow);
            Assert.Equal(1, mergedRange.StartColumn);
            Assert.Equal(1, mergedRange.EndRow);
            Assert.Equal(3, mergedRange.EndColumn);
            Assert.NotNull(layout.FreezePane);
            Assert.Equal(2, layout.FreezePane!.TopRows);
            Assert.Equal(1, layout.FreezePane.LeftColumns);
            Assert.False(layout.ShowGridLines);
            Assert.False(layout.ShowRowColumnHeadings);
            Assert.False(layout.ShowZeroValues);
            Assert.True(layout.RightToLeft);
            Assert.Equal(18.5d, layout.DefaultRowHeight);
            Assert.False(layout.DefaultRowsHidden);
            Assert.Equal(11d, layout.DefaultColumnWidth);
            Assert.Equal(1, legacy.Worksheets[1].Visibility);
            Assert.Equal(2, legacy.Worksheets[2].Visibility);

            using ExcelDocument document = ExcelDocument.LoadLegacyXls(new MemoryStream(compound), new LegacyXlsImportOptions {
                ReportUnsupportedRecords = false
            });

            Assert.Equal(3, document.Sheets.Count);
            ExcelSheet projected = document.Sheets[0];
            ExcelColumnSnapshot projectedColumn = Assert.Single(projected.GetColumnDefinitions());
            Assert.Equal(2, projectedColumn.StartIndex);
            Assert.Equal(2, projectedColumn.EndIndex);
            Assert.Equal(12.5d, projectedColumn.Width);
            Assert.True(projectedColumn.Hidden);
            Assert.Equal((byte?)2, projectedColumn.OutlineLevel);
            Assert.True(projectedColumn.Collapsed);

            ExcelRowSnapshot projectedRow = Assert.Single(projected.GetRowDefinitions());
            Assert.Equal(2, projectedRow.Index);
            Assert.Equal(18d, projectedRow.Height);
            Assert.True(projectedRow.Hidden);
            Assert.Equal((byte?)1, projectedRow.OutlineLevel);
            Assert.True(projectedRow.Collapsed);

            ExcelMergedRangeSnapshot projectedMerge = Assert.Single(projected.GetMergedRanges());
            Assert.Equal("A1:C1", projectedMerge.A1Range);
            Assert.True(document.Sheets[1].Hidden);
            Assert.False(document.Sheets[1].VeryHidden);
            Assert.True(document.Sheets[2].Hidden);
            Assert.True(document.Sheets[2].VeryHidden);
            Assert.True(projected.RightToLeft);
            Assert.False(projected.RowColumnHeadingsVisible);
            Assert.False(projected.ZeroValuesVisible);
            Assert.Equal(18.5d, projected.DefaultRowHeight);
            Assert.False(projected.DefaultRowsHidden);
            Assert.Equal(11d, projected.DefaultColumnWidth);

            using var output = new MemoryStream();
            document.Save(output);
            using SpreadsheetDocument spreadsheet = SpreadsheetDocument.Open(new MemoryStream(output.ToArray()), false);
            WorkbookPart workbookPart = spreadsheet.WorkbookPart!;
            Sheet hiddenSheet = workbookPart.Workbook.Sheets!.Elements<Sheet>().First(s => s.Name == "Hidden");
            Sheet veryHiddenSheet = workbookPart.Workbook.Sheets!.Elements<Sheet>().First(s => s.Name == "VeryHidden");
            Assert.Equal(SheetStateValues.Hidden, hiddenSheet.State!.Value);
            Assert.Equal(SheetStateValues.VeryHidden, veryHiddenSheet.State!.Value);
            Sheet openXmlSheet = workbookPart.Workbook.Sheets!.Elements<Sheet>().First(s => s.Name == "Layout");
            WorksheetPart worksheetPart = (WorksheetPart)workbookPart.GetPartById(openXmlSheet.Id!);
            Worksheet worksheet = worksheetPart.Worksheet!;
            Pane pane = worksheet.GetFirstChild<SheetViews>()!.GetFirstChild<SheetView>()!.GetFirstChild<Pane>()!;
            SheetView sheetView = worksheet.GetFirstChild<SheetViews>()!.GetFirstChild<SheetView>()!;
            SheetFormatProperties sheetFormat = worksheet.GetFirstChild<SheetFormatProperties>()!;
            Column openXmlColumn = worksheet.GetFirstChild<Columns>()!.Elements<Column>().Single(c => c.Min!.Value == 2U && c.Max!.Value == 2U);
            Row openXmlRow = worksheet.GetFirstChild<SheetData>()!.Elements<Row>().Single(r => r.RowIndex!.Value == 2U);
            Assert.Equal(2d, pane.VerticalSplit!.Value);
            Assert.Equal(1d, pane.HorizontalSplit!.Value);
            Assert.Equal(PaneStateValues.Frozen, pane.State!.Value);
            Assert.False(sheetView.ShowGridLines!.Value);
            Assert.False(sheetView.ShowRowColHeaders!.Value);
            Assert.False(sheetView.ShowZeros!.Value);
            Assert.True(sheetView.RightToLeft!.Value);
            Assert.Equal(18.5d, sheetFormat.DefaultRowHeight!.Value);
            Assert.Equal(11d, sheetFormat.DefaultColumnWidth!.Value);
            Assert.True(sheetFormat.CustomHeight!.Value);
            Assert.Equal(2, openXmlColumn.OutlineLevel!.Value);
            Assert.True(openXmlColumn.Collapsed!.Value);
            Assert.Equal(1, openXmlRow.OutlineLevel!.Value);
            Assert.True(openXmlRow.Collapsed!.Value);
        }

        [Fact]
        public void LegacyXls_LoadLegacyXls_ImportsAndProjectsPhase3StyleRecords() {
            byte[] workbookStream = LegacyXlsTestWorkbookBuilder.CreatePhase3StyleWorkbookStream();
            byte[] compound = LegacyXlsCompoundTestBuilder.CreateWorkbookCompoundFile(workbookStream);

            LegacyXlsWorkbook legacy = LegacyXlsWorkbook.Load(compound, new LegacyXlsImportOptions {
                ReportUnsupportedRecords = false
            });

            Assert.DoesNotContain(legacy.Diagnostics, d => d.Severity == LegacyXlsDiagnosticSeverity.Error);
            LegacyXlsNumberFormat numberFormat = Assert.Single(legacy.NumberFormats);
            Assert.Equal(164, numberFormat.FormatId);
            Assert.Equal("yyyy-mm-dd", numberFormat.FormatCode);
            Assert.Equal(11, legacy.CellFormats.Count);
            Assert.Equal("0.00", legacy.CellFormats[1].NumberFormatCode);
            Assert.True(legacy.CellFormats[1].IsBuiltInNumberFormat);
            Assert.Equal("yyyy-mm-dd", legacy.CellFormats[2].NumberFormatCode);
            Assert.True(legacy.CellFormats[2].IsDateLike);
            Assert.Equal(6, legacy.Fonts.Count);
            Assert.Equal("Consolas", legacy.Fonts[4].Name);
            Assert.Equal(14d, legacy.Fonts[4].Size);
            Assert.Equal(0x0008, legacy.Fonts[4].ColorIndex);
            Assert.True(legacy.Fonts[4].Bold);
            Assert.True(legacy.Fonts[4].Italic);
            Assert.True(legacy.Fonts[4].Underline);
            Assert.True(legacy.Fonts[4].Strikeout);
            Assert.Equal(LegacyXlsFontEscapement.Superscript, legacy.Fonts[4].Escapement);
            Assert.Equal("Courier New", legacy.Fonts[5].Name);
            Assert.Equal(LegacyXlsFontEscapement.Subscript, legacy.Fonts[5].Escapement);
            Assert.Equal(56, legacy.PaletteColors.Count);
            Assert.Equal("FF123456", legacy.PaletteColors[0]);
            Assert.Equal("FFABCDEF", legacy.PaletteColors[1]);
            Assert.Equal(5, legacy.CellFormats[3].FontIndex);
            Assert.Equal(1, legacy.CellFormats[4].FillPattern);
            Assert.Equal(0x0009, legacy.CellFormats[4].FillForegroundColorIndex);
            Assert.Equal(5, legacy.CellFormats[7].FillPattern);
            Assert.Equal(0x0008, legacy.CellFormats[7].FillForegroundColorIndex);
            Assert.Equal(0x0009, legacy.CellFormats[7].FillBackgroundColorIndex);
            Assert.True(legacy.CellFormats[5].ApplyAlignment);
            Assert.Equal(2, legacy.CellFormats[5].HorizontalAlignment);
            Assert.Equal(1, legacy.CellFormats[5].VerticalAlignment);
            Assert.True(legacy.CellFormats[5].WrapText);
            Assert.Equal(45, legacy.CellFormats[5].TextRotation);
            Assert.Equal(3, legacy.CellFormats[5].Indent);
            Assert.True(legacy.CellFormats[5].ShrinkToFit);
            Assert.Equal(2, legacy.CellFormats[5].ReadingOrder);
            Assert.NotNull(legacy.CellFormats[6].Border);
            Assert.Equal(1, legacy.CellFormats[6].Border!.LeftStyle);
            Assert.Equal(0x000a, legacy.CellFormats[6].Border!.LeftColorIndex);
            Assert.Equal(8, legacy.CellFormats[6].Border!.RightStyle);
            Assert.Equal(3, legacy.CellFormats[6].Border!.TopStyle);
            Assert.Equal(6, legacy.CellFormats[6].Border!.BottomStyle);
            Assert.Equal(4, legacy.CellFormats[6].Border!.DiagonalStyle);
            Assert.True(legacy.CellFormats[6].Border!.DiagonalUp);
            Assert.True(legacy.CellFormats[6].Border!.DiagonalDown);
            Assert.True(legacy.CellFormats[8].ApplyProtection);
            Assert.False(legacy.CellFormats[8].Locked);
            Assert.True(legacy.CellFormats[8].FormulaHidden);
            Assert.True(legacy.CellFormats[9].QuotePrefix);
            LegacyXlsColumnLayout defaultStyledColumn = Assert.Single(legacy.Worksheets[0].Columns);
            Assert.Equal(11, defaultStyledColumn.StartColumn);
            Assert.Equal(11, defaultStyledColumn.EndColumn);
            Assert.Equal(4, defaultStyledColumn.StyleIndex);
            LegacyXlsRowLayout defaultStyledRow = Assert.Single(legacy.Worksheets[0].Rows);
            Assert.Equal(3, defaultStyledRow.Row);
            Assert.Equal((ushort?)5, defaultStyledRow.StyleIndex);

            using ExcelDocument document = ExcelDocument.LoadLegacyXls(new MemoryStream(compound), new LegacyXlsImportOptions {
                ReportUnsupportedRecords = false
            });

            using var output = new MemoryStream();
            document.Save(output);
            using SpreadsheetDocument spreadsheet = SpreadsheetDocument.Open(new MemoryStream(output.ToArray()), false);
            WorkbookPart workbookPart = spreadsheet.WorkbookPart!;
            WorksheetPart worksheetPart = workbookPart.WorksheetParts.Single();
            Dictionary<string, Cell> cells = worksheetPart.Worksheet.Descendants<Cell>()
                .ToDictionary(cell => cell.CellReference!.Value!);

            Cell amountCell = cells["A1"];
            Assert.Equal("12.345", amountCell.CellValue!.Text);
            Assert.NotNull(amountCell.StyleIndex);
            CellFormat amountFormat = workbookPart.WorkbookStylesPart!.Stylesheet.CellFormats!.Elements<CellFormat>().ElementAt((int)amountCell.StyleIndex!.Value);
            Assert.Equal(2U, amountFormat.NumberFormatId!.Value);

            Cell dateCell = cells["B1"];
            Assert.Equal(new DateTime(2024, 1, 2).ToOADate().ToString(CultureInfo.InvariantCulture), dateCell.CellValue!.Text);
            Assert.NotNull(dateCell.StyleIndex);
            CellFormat dateFormat = workbookPart.WorkbookStylesPart!.Stylesheet.CellFormats!.Elements<CellFormat>().ElementAt((int)dateCell.StyleIndex!.Value);
            uint dateNumberFormatId = dateFormat.NumberFormatId!.Value;
            NumberingFormat projectedNumberFormat = Assert.Single(workbookPart.WorkbookStylesPart.Stylesheet.NumberingFormats!.Elements<NumberingFormat>(),
                format => format.NumberFormatId!.Value == dateNumberFormatId);
            Assert.Equal("yyyy-mm-dd", projectedNumberFormat.FormatCode!.Value);

            Cell fontCell = cells["C1"];
            Assert.NotNull(fontCell.StyleIndex);
            CellFormat fontFormat = workbookPart.WorkbookStylesPart.Stylesheet.CellFormats!.Elements<CellFormat>().ElementAt((int)fontCell.StyleIndex!.Value);
            DocumentFormat.OpenXml.Spreadsheet.Font projectedFont = workbookPart.WorkbookStylesPart.Stylesheet.Fonts!.Elements<DocumentFormat.OpenXml.Spreadsheet.Font>().ElementAt((int)fontFormat.FontId!.Value);
            Assert.Equal("Consolas", projectedFont.FontName!.Val!.Value);
            Assert.Equal(14d, projectedFont.FontSize!.Val!.Value);
            Assert.Equal("FF123456", projectedFont.Color!.Rgb!.Value);
            Assert.NotNull(projectedFont.Bold);
            Assert.NotNull(projectedFont.Italic);
            Assert.NotNull(projectedFont.Underline);
            Assert.NotNull(projectedFont.Strike);
            Assert.Equal(VerticalAlignmentRunValues.Superscript, projectedFont.VerticalTextAlignment!.Val!.Value);

            Cell subscriptCell = cells["K1"];
            Assert.NotNull(subscriptCell.StyleIndex);
            CellFormat subscriptFormat = workbookPart.WorkbookStylesPart.Stylesheet.CellFormats!.Elements<CellFormat>().ElementAt((int)subscriptCell.StyleIndex!.Value);
            DocumentFormat.OpenXml.Spreadsheet.Font subscriptFont = workbookPart.WorkbookStylesPart.Stylesheet.Fonts!.Elements<DocumentFormat.OpenXml.Spreadsheet.Font>().ElementAt((int)subscriptFormat.FontId!.Value);
            Assert.Equal("Courier New", subscriptFont.FontName!.Val!.Value);
            Assert.Equal(VerticalAlignmentRunValues.Subscript, subscriptFont.VerticalTextAlignment!.Val!.Value);

            Cell fillCell = cells["D1"];
            Assert.NotNull(fillCell.StyleIndex);
            CellFormat fillFormat = workbookPart.WorkbookStylesPart.Stylesheet.CellFormats!.Elements<CellFormat>().ElementAt((int)fillCell.StyleIndex!.Value);
            Fill projectedFill = workbookPart.WorkbookStylesPart.Stylesheet.Fills!.Elements<Fill>().ElementAt((int)fillFormat.FillId!.Value);
            Assert.Equal(PatternValues.Solid, projectedFill.PatternFill!.PatternType!.Value);
            Assert.Equal("FFABCDEF", projectedFill.PatternFill.ForegroundColor!.Rgb!.Value);

            Cell blankFillCell = cells["F1"];
            Assert.Null(blankFillCell.CellValue);
            Assert.NotNull(blankFillCell.StyleIndex);
            CellFormat blankFillFormat = workbookPart.WorkbookStylesPart.Stylesheet.CellFormats!.Elements<CellFormat>().ElementAt((int)blankFillCell.StyleIndex!.Value);
            Fill projectedBlankFill = workbookPart.WorkbookStylesPart.Stylesheet.Fills!.Elements<Fill>().ElementAt((int)blankFillFormat.FillId!.Value);
            Assert.Equal("FFABCDEF", projectedBlankFill.PatternFill!.ForegroundColor!.Rgb!.Value);

            Cell patternCell = cells["H1"];
            Assert.NotNull(patternCell.StyleIndex);
            CellFormat patternFormat = workbookPart.WorkbookStylesPart.Stylesheet.CellFormats!.Elements<CellFormat>().ElementAt((int)patternCell.StyleIndex!.Value);
            Fill projectedPatternFill = workbookPart.WorkbookStylesPart.Stylesheet.Fills!.Elements<Fill>().ElementAt((int)patternFormat.FillId!.Value);
            Assert.Equal(PatternValues.DarkHorizontal, projectedPatternFill.PatternFill!.PatternType!.Value);
            Assert.Equal("FF123456", projectedPatternFill.PatternFill.ForegroundColor!.Rgb!.Value);
            Assert.Equal("FFABCDEF", projectedPatternFill.PatternFill.BackgroundColor!.Rgb!.Value);

            Cell alignmentCell = cells["E1"];
            Assert.NotNull(alignmentCell.StyleIndex);
            CellFormat alignmentFormat = workbookPart.WorkbookStylesPart.Stylesheet.CellFormats!.Elements<CellFormat>().ElementAt((int)alignmentCell.StyleIndex!.Value);
            Assert.True(alignmentFormat.ApplyAlignment!.Value);
            Assert.Equal(HorizontalAlignmentValues.Center, alignmentFormat.Alignment!.Horizontal!.Value);
            Assert.Equal(VerticalAlignmentValues.Center, alignmentFormat.Alignment.Vertical!.Value);
            Assert.True(alignmentFormat.Alignment.WrapText!.Value);
            Assert.Equal(45U, alignmentFormat.Alignment.TextRotation!.Value);
            Assert.Equal(3U, alignmentFormat.Alignment.Indent!.Value);
            Assert.True(alignmentFormat.Alignment.ShrinkToFit!.Value);
            Assert.Equal(2U, alignmentFormat.Alignment.ReadingOrder!.Value);

            Cell borderCell = cells["G1"];
            Assert.NotNull(borderCell.StyleIndex);
            CellFormat borderFormat = workbookPart.WorkbookStylesPart.Stylesheet.CellFormats!.Elements<CellFormat>().ElementAt((int)borderCell.StyleIndex!.Value);
            Assert.True(borderFormat.ApplyBorder!.Value);
            Border projectedBorder = workbookPart.WorkbookStylesPart.Stylesheet.Borders!.Elements<Border>().ElementAt((int)borderFormat.BorderId!.Value);
            Assert.Equal(BorderStyleValues.Thin, projectedBorder.LeftBorder!.Style!.Value);
            Assert.Equal("FF654321", projectedBorder.LeftBorder.GetFirstChild<DocumentFormat.OpenXml.Spreadsheet.Color>()!.Rgb!.Value);
            Assert.Equal(BorderStyleValues.MediumDashed, projectedBorder.RightBorder!.Style!.Value);
            Assert.Equal("FF123456", projectedBorder.RightBorder.GetFirstChild<DocumentFormat.OpenXml.Spreadsheet.Color>()!.Rgb!.Value);
            Assert.Equal(BorderStyleValues.Dashed, projectedBorder.TopBorder!.Style!.Value);
            Assert.Equal("FFABCDEF", projectedBorder.TopBorder.GetFirstChild<DocumentFormat.OpenXml.Spreadsheet.Color>()!.Rgb!.Value);
            Assert.Equal(BorderStyleValues.Double, projectedBorder.BottomBorder!.Style!.Value);
            Assert.Equal("FF654321", projectedBorder.BottomBorder.GetFirstChild<DocumentFormat.OpenXml.Spreadsheet.Color>()!.Rgb!.Value);
            Assert.Equal(BorderStyleValues.Dotted, projectedBorder.DiagonalBorder!.Style!.Value);
            Assert.Equal("FF123456", projectedBorder.DiagonalBorder.GetFirstChild<DocumentFormat.OpenXml.Spreadsheet.Color>()!.Rgb!.Value);
            Assert.True(projectedBorder.DiagonalUp!.Value);
            Assert.True(projectedBorder.DiagonalDown!.Value);

            Cell protectionCell = cells["I1"];
            Assert.NotNull(protectionCell.StyleIndex);
            CellFormat protectionFormat = workbookPart.WorkbookStylesPart.Stylesheet.CellFormats!.Elements<CellFormat>().ElementAt((int)protectionCell.StyleIndex!.Value);
            Assert.True(protectionFormat.ApplyProtection!.Value);
            Assert.False(protectionFormat.Protection!.Locked!.Value);
            Assert.True(protectionFormat.Protection.Hidden!.Value);

            Cell quotePrefixCell = cells["J1"];
            Assert.NotNull(quotePrefixCell.StyleIndex);
            CellFormat quotePrefixFormat = workbookPart.WorkbookStylesPart.Stylesheet.CellFormats!.Elements<CellFormat>().ElementAt((int)quotePrefixCell.StyleIndex!.Value);
            Assert.True(quotePrefixFormat.QuotePrefix!.Value);

            ExcelColumnSnapshot columnSnapshot = Assert.Single(document.Sheets[0].GetColumnDefinitions());
            Assert.Equal(11, columnSnapshot.StartIndex);
            Assert.Equal(11, columnSnapshot.EndIndex);
            Assert.NotNull(columnSnapshot.StyleIndex);
            Column defaultStyledOpenXmlColumn = worksheetPart.Worksheet.GetFirstChild<Columns>()!.Elements<Column>().Single();
            Assert.Equal(columnSnapshot.StyleIndex!.Value, defaultStyledOpenXmlColumn.Style!.Value);
            CellFormat columnFormat = workbookPart.WorkbookStylesPart.Stylesheet.CellFormats!.Elements<CellFormat>().ElementAt((int)columnSnapshot.StyleIndex.Value);
            Fill columnFill = workbookPart.WorkbookStylesPart.Stylesheet.Fills!.Elements<Fill>().ElementAt((int)columnFormat.FillId!.Value);
            Assert.Equal(PatternValues.Solid, columnFill.PatternFill!.PatternType!.Value);
            Assert.Equal("FFABCDEF", columnFill.PatternFill.ForegroundColor!.Rgb!.Value);

            ExcelRowSnapshot rowSnapshot = Assert.Single(document.Sheets[0].GetRowDefinitions());
            Assert.Equal(3, rowSnapshot.Index);
            Assert.True(rowSnapshot.CustomFormat);
            Assert.NotNull(rowSnapshot.StyleIndex);
            Row defaultStyledOpenXmlRow = worksheetPart.Worksheet.GetFirstChild<SheetData>()!.Elements<Row>().Single(row => row.RowIndex!.Value == 3U);
            Assert.Equal(rowSnapshot.StyleIndex!.Value, defaultStyledOpenXmlRow.StyleIndex!.Value);
            Assert.True(defaultStyledOpenXmlRow.CustomFormat!.Value);
            CellFormat rowFormat = workbookPart.WorkbookStylesPart.Stylesheet.CellFormats!.Elements<CellFormat>().ElementAt((int)rowSnapshot.StyleIndex.Value);
            Assert.True(rowFormat.ApplyAlignment!.Value);
            Assert.Equal(HorizontalAlignmentValues.Center, rowFormat.Alignment!.Horizontal!.Value);
            Assert.Equal(VerticalAlignmentValues.Center, rowFormat.Alignment.Vertical!.Value);
        }

        [Fact]
        public void LegacyXls_Load_ReportsInvalidCompoundFile() {
            LegacyXlsWorkbook legacy = LegacyXlsWorkbook.Load(new byte[] { 1, 2, 3 });

            Assert.Contains(legacy.Diagnostics, item => item.Code == "XLS-COMPOUND-SIGNATURE");
            Assert.Empty(legacy.Worksheets);
        }

        private static partial class LegacyXlsTestWorkbookBuilder {
            internal static byte[] CreateMinimalWorkbookStream() {
                using var stream = new MemoryStream();
                WriteRecord(stream, 0x0809, new byte[] { 0x00, 0x06, 0x05, 0x00, 0xdb, 0x0b, 0xcc, 0x07 });
                long boundSheetPosition = stream.Position;
                WriteRecord(stream, 0x0085, BuildBoundSheetPayload(0, "Sheet1"));
                WriteRecord(stream, 0x00fc, BuildSstPayload("Name"));
                WriteRecord(stream, 0x000a, Array.Empty<byte>());

                int sheetOffset = checked((int)stream.Position);
                WriteRecord(stream, 0x0809, new byte[] { 0x00, 0x06, 0x10, 0x00, 0xdb, 0x0b, 0xcc, 0x07 });
                WriteRecord(stream, 0x00fd, BuildLabelSstPayload(0, 0, 0));
                WriteRecord(stream, 0x0203, BuildNumberPayload(1, 1, 42d));
                WriteRecord(stream, 0x0205, BuildBoolErrPayload(2, 0, true));
                WriteRecord(stream, 0x000a, Array.Empty<byte>());

                byte[] bytes = stream.ToArray();
                byte[] offsetBytes = BitConverter.GetBytes(sheetOffset);
                Buffer.BlockCopy(offsetBytes, 0, bytes, checked((int)boundSheetPosition + 4), offsetBytes.Length);
                return bytes;
            }

            internal static byte[] CreateEncryptedWorkbookStream() {
                using var stream = new MemoryStream();
                WriteRecord(stream, 0x0809, new byte[] { 0x00, 0x06, 0x05, 0x00, 0xdb, 0x0b, 0xcc, 0x07 });
                WriteRecord(stream, 0x002f, new byte[] { 0x00, 0x00 });
                WriteRecord(stream, 0x000a, Array.Empty<byte>());
                return stream.ToArray();
            }

            internal static byte[] CreateUnsupportedFeatureWorkbookStream() {
                using var stream = new MemoryStream();
                WriteRecord(stream, 0x0809, new byte[] { 0x00, 0x06, 0x05, 0x00, 0xdb, 0x0b, 0xcc, 0x07 });
                long boundSheetPosition = stream.Position;
                WriteRecord(stream, 0x0085, BuildBoundSheetPayload(0, "Features"));
                WriteRecord(stream, 0x000a, Array.Empty<byte>());

                int sheetOffset = checked((int)stream.Position);
                WriteRecord(stream, 0x0809, new byte[] { 0x00, 0x06, 0x10, 0x00, 0xdb, 0x0b, 0xcc, 0x07 });
                WriteRecord(stream, 0x0204, BuildLabelPayload(0, 0, "Feature"));
                WriteRecord(stream, 0x01b8, new byte[24]);
                WriteRecord(stream, 0x001c, new byte[12]);
                WriteRecord(stream, 0x00ec, Array.Empty<byte>());
                WriteRecord(stream, 0x00b0, Array.Empty<byte>());
                WriteRecord(stream, 0x009e, Array.Empty<byte>());
                WriteRecord(stream, 0x000a, Array.Empty<byte>());

                byte[] bytes = stream.ToArray();
                Buffer.BlockCopy(BitConverter.GetBytes(sheetOffset), 0, bytes, checked((int)boundSheetPosition + 4), 4);
                return bytes;
            }

            internal static byte[] CreateWorksheetDimensionsWorkbookStream() {
                using var stream = new MemoryStream();
                WriteRecord(stream, 0x0809, new byte[] { 0x00, 0x06, 0x05, 0x00, 0xdb, 0x0b, 0xcc, 0x07 });
                long boundsBoundSheetPosition = stream.Position;
                WriteRecord(stream, 0x0085, BuildBoundSheetPayload(0, "Bounds"));
                long emptyBoundSheetPosition = stream.Position;
                WriteRecord(stream, 0x0085, BuildBoundSheetPayload(0, "Empty"));
                WriteRecord(stream, 0x000a, Array.Empty<byte>());

                int boundsSheetOffset = checked((int)stream.Position);
                WriteRecord(stream, 0x0809, new byte[] { 0x00, 0x06, 0x10, 0x00, 0xdb, 0x0b, 0xcc, 0x07 });
                WriteRecord(stream, 0x0200, BuildDimensionsPayload(firstRow: 1, rowAfterLast: 5, firstColumn: 1, columnAfterLast: 4));
                WriteRecord(stream, 0x0204, BuildLabelPayload(1, 1, "Start"));
                WriteRecord(stream, 0x0203, BuildNumberPayload(4, 3, 99d));
                WriteRecord(stream, 0x000a, Array.Empty<byte>());

                int emptySheetOffset = checked((int)stream.Position);
                WriteRecord(stream, 0x0809, new byte[] { 0x00, 0x06, 0x10, 0x00, 0xdb, 0x0b, 0xcc, 0x07 });
                WriteRecord(stream, 0x0200, BuildDimensionsPayload(firstRow: 0, rowAfterLast: 0, firstColumn: 0, columnAfterLast: 0));
                WriteRecord(stream, 0x000a, Array.Empty<byte>());

                byte[] bytes = stream.ToArray();
                Buffer.BlockCopy(BitConverter.GetBytes(boundsSheetOffset), 0, bytes, checked((int)boundsBoundSheetPosition + 4), 4);
                Buffer.BlockCopy(BitConverter.GetBytes(emptySheetOffset), 0, bytes, checked((int)emptyBoundSheetPosition + 4), 4);
                return bytes;
            }

            internal static byte[] CreatePhase3LayoutWorkbookStream() {
                using var stream = new MemoryStream();
                WriteRecord(stream, 0x0809, new byte[] { 0x00, 0x06, 0x05, 0x00, 0xdb, 0x0b, 0xcc, 0x07 });
                long layoutBoundSheetPosition = stream.Position;
                WriteRecord(stream, 0x0085, BuildBoundSheetPayload(0, "Layout"));
                long hiddenBoundSheetPosition = stream.Position;
                WriteRecord(stream, 0x0085, BuildBoundSheetPayload(0, "Hidden", visibility: 1));
                long veryHiddenBoundSheetPosition = stream.Position;
                WriteRecord(stream, 0x0085, BuildBoundSheetPayload(0, "VeryHidden", visibility: 2));
                WriteRecord(stream, 0x000a, Array.Empty<byte>());

                int layoutSheetOffset = checked((int)stream.Position);
                WriteRecord(stream, 0x0809, new byte[] { 0x00, 0x06, 0x10, 0x00, 0xdb, 0x0b, 0xcc, 0x07 });
                WriteRecord(stream, 0x0055, BuildDefaultColumnWidthPayload(11));
                WriteRecord(stream, 0x0225, BuildDefaultRowHeightPayload(18.5d));
                WriteRecord(stream, 0x0204, BuildLabelPayload(0, 0, "Merged"));
                WriteRecord(stream, 0x007d, BuildColInfoPayload(1, 1, 12.5d, hidden: true, styleIndex: 3, outlineLevel: 2, collapsed: true));
                WriteRecord(stream, 0x0208, BuildRowPayload(1, 18d, hidden: true, customHeight: true, outlineLevel: 1, collapsed: true));
                WriteRecord(stream, 0x023e, BuildWindow2Payload(frozen: true, showGridlines: false, showRowColumnHeadings: false, showZeroValues: false, rightToLeft: true));
                WriteRecord(stream, 0x0041, BuildPanePayload(leftColumns: 1, topRows: 2));
                WriteRecord(stream, 0x00e5, BuildMergeCellsPayload((0, 0, 0, 2)));
                WriteRecord(stream, 0x000a, Array.Empty<byte>());

                int hiddenSheetOffset = checked((int)stream.Position);
                WriteRecord(stream, 0x0809, new byte[] { 0x00, 0x06, 0x10, 0x00, 0xdb, 0x0b, 0xcc, 0x07 });
                WriteRecord(stream, 0x0204, BuildLabelPayload(0, 0, "Secret"));
                WriteRecord(stream, 0x000a, Array.Empty<byte>());

                int veryHiddenSheetOffset = checked((int)stream.Position);
                WriteRecord(stream, 0x0809, new byte[] { 0x00, 0x06, 0x10, 0x00, 0xdb, 0x0b, 0xcc, 0x07 });
                WriteRecord(stream, 0x0204, BuildLabelPayload(0, 0, "VerySecret"));
                WriteRecord(stream, 0x000a, Array.Empty<byte>());

                byte[] bytes = stream.ToArray();
                Buffer.BlockCopy(BitConverter.GetBytes(layoutSheetOffset), 0, bytes, checked((int)layoutBoundSheetPosition + 4), 4);
                Buffer.BlockCopy(BitConverter.GetBytes(hiddenSheetOffset), 0, bytes, checked((int)hiddenBoundSheetPosition + 4), 4);
                Buffer.BlockCopy(BitConverter.GetBytes(veryHiddenSheetOffset), 0, bytes, checked((int)veryHiddenBoundSheetPosition + 4), 4);
                return bytes;
            }

            internal static byte[] CreatePhase3StyleWorkbookStream() {
                using var stream = new MemoryStream();
                WriteRecord(stream, 0x0809, new byte[] { 0x00, 0x06, 0x05, 0x00, 0xdb, 0x0b, 0xcc, 0x07 });
                long boundSheetPosition = stream.Position;
                WriteRecord(stream, 0x0085, BuildBoundSheetPayload(0, "Styles"));
                WriteRecord(stream, 0x0031, BuildFontPayload("Arial", 11d, bold: false, italic: false, underline: false));
                WriteRecord(stream, 0x0031, BuildFontPayload("Arial", 11d, bold: true, italic: false, underline: false));
                WriteRecord(stream, 0x0031, BuildFontPayload("Arial", 11d, bold: false, italic: true, underline: false));
                WriteRecord(stream, 0x0031, BuildFontPayload("Arial", 11d, bold: true, italic: true, underline: false));
                WriteRecord(stream, 0x0031, BuildFontPayload("Consolas", 14d, bold: true, italic: true, underline: true, strikeout: true, colorIndex: 0x0008, escapement: 1));
                WriteRecord(stream, 0x0031, BuildFontPayload("Courier New", 11d, bold: false, italic: false, underline: false, escapement: 2));
                WriteRecord(stream, 0x0092, BuildPalettePayload("FF123456", "FFABCDEF", "FF654321"));
                WriteRecord(stream, 0x041e, BuildFormatPayload(164, "yyyy-mm-dd"));
                WriteRecord(stream, 0x00e0, BuildXfPayload(0));
                WriteRecord(stream, 0x00e0, BuildXfPayload(2));
                WriteRecord(stream, 0x00e0, BuildXfPayload(164));
                WriteRecord(stream, 0x00e0, BuildXfPayload(0, fontIndex: 5));
                WriteRecord(stream, 0x00e0, BuildXfPayload(0, fillPattern: 1, fillForegroundColorIndex: 0x0009));
                WriteRecord(stream, 0x00e0, BuildXfPayload(
                    0,
                    applyAlignment: true,
                    horizontalAlignment: 2,
                    verticalAlignment: 1,
                    wrapText: true,
                    textRotation: 45,
                    indent: 3,
                    shrinkToFit: true,
                    readingOrder: 2));
                WriteRecord(stream, 0x00e0, BuildXfPayload(
                    0,
                    leftBorderStyle: 1,
                    leftBorderColorIndex: 0x000a,
                    rightBorderStyle: 8,
                    rightBorderColorIndex: 0x0008,
                    topBorderStyle: 3,
                    topBorderColorIndex: 0x0009,
                    bottomBorderStyle: 6,
                    bottomBorderColorIndex: 0x000a,
                    diagonalBorderStyle: 4,
                    diagonalBorderColorIndex: 0x0008,
                    diagonalFlags: 3));
                WriteRecord(stream, 0x00e0, BuildXfPayload(0, fillPattern: 5, fillForegroundColorIndex: 0x0008, fillBackgroundColorIndex: 0x0009));
                WriteRecord(stream, 0x00e0, BuildXfPayload(0, locked: false, formulaHidden: true, applyProtection: true));
                WriteRecord(stream, 0x00e0, BuildXfPayload(0, quotePrefix: true));
                WriteRecord(stream, 0x00e0, BuildXfPayload(0, fontIndex: 6));
                WriteRecord(stream, 0x000a, Array.Empty<byte>());

                int sheetOffset = checked((int)stream.Position);
                WriteRecord(stream, 0x0809, new byte[] { 0x00, 0x06, 0x10, 0x00, 0xdb, 0x0b, 0xcc, 0x07 });
                WriteRecord(stream, 0x0203, BuildNumberPayload(0, 0, 12.345d, styleIndex: 1));
                WriteRecord(stream, 0x0203, BuildNumberPayload(0, 1, new DateTime(2024, 1, 2).ToOADate(), styleIndex: 2));
                WriteRecord(stream, 0x0204, BuildLabelPayload(0, 2, "Styled", styleIndex: 3));
                WriteRecord(stream, 0x0204, BuildLabelPayload(0, 3, "Filled", styleIndex: 4));
                WriteRecord(stream, 0x0204, BuildLabelPayload(0, 4, "Aligned", styleIndex: 5));
                WriteRecord(stream, 0x0201, BuildBlankPayload(0, 5, styleIndex: 4));
                WriteRecord(stream, 0x0204, BuildLabelPayload(0, 6, "Bordered", styleIndex: 6));
                WriteRecord(stream, 0x0204, BuildLabelPayload(0, 7, "Pattern", styleIndex: 7));
                WriteRecord(stream, 0x0204, BuildLabelPayload(0, 8, "Protection", styleIndex: 8));
                WriteRecord(stream, 0x0204, BuildLabelPayload(0, 9, "Prefixed", styleIndex: 9));
                WriteRecord(stream, 0x0204, BuildLabelPayload(0, 10, "Subscript", styleIndex: 10));
                WriteRecord(stream, 0x007d, BuildColInfoPayload(10, 10, 9.5d, hidden: false, styleIndex: 4));
                WriteRecord(stream, 0x0208, BuildRowPayload(2, 18d, hidden: false, customHeight: false, styleIndex: 5));
                WriteRecord(stream, 0x000a, Array.Empty<byte>());

                byte[] bytes = stream.ToArray();
                Buffer.BlockCopy(BitConverter.GetBytes(sheetOffset), 0, bytes, checked((int)boundSheetPosition + 4), 4);
                return bytes;
            }

            internal static byte[] CreatePhase4FormulaWorkbookStream() {
                using var stream = new MemoryStream();
                WriteRecord(stream, 0x0809, new byte[] { 0x00, 0x06, 0x05, 0x00, 0xdb, 0x0b, 0xcc, 0x07 });
                long boundSheetPosition = stream.Position;
                WriteRecord(stream, 0x0085, BuildBoundSheetPayload(0, "Formulas"));
                WriteRecord(stream, 0x000a, Array.Empty<byte>());

                int sheetOffset = checked((int)stream.Position);
                WriteRecord(stream, 0x0809, new byte[] { 0x00, 0x06, 0x10, 0x00, 0xdb, 0x0b, 0xcc, 0x07 });
                WriteRecord(stream, 0x0006, BuildFormulaNumberPayload(0, 0, 42d));
                WriteRecord(stream, 0x0006, BuildFormulaSpecialPayload(0, 1, valueType: 0x01, value: 1));
                WriteRecord(stream, 0x0006, BuildFormulaSpecialPayload(0, 2, valueType: 0x00, value: 0));
                WriteRecord(stream, 0x0207, BuildFormulaStringPayload("Formula text"));
                WriteRecord(stream, 0x0006, BuildFormulaSpecialPayload(0, 3, valueType: 0x02, value: 0x0f));
                WriteRecord(stream, 0x0203, BuildNumberPayload(1, 0, 10d));
                WriteRecord(stream, 0x0203, BuildNumberPayload(1, 1, 32d));
                WriteRecord(stream, 0x0006, BuildFormulaNumberPayload(1, 2, 42d, formulaTokens: BuildReferenceAdditionFormulaTokens(1, 0, 1, 1)));
                WriteRecord(stream, 0x0006, BuildFormulaNumberPayload(1, 3, 42d, formulaTokens: BuildSumAreaFormulaTokens(1, 0, 1, 1)));
                WriteRecord(stream, 0x000a, Array.Empty<byte>());

                byte[] bytes = stream.ToArray();
                Buffer.BlockCopy(BitConverter.GetBytes(sheetOffset), 0, bytes, checked((int)boundSheetPosition + 4), 4);
                return bytes;
            }

            internal static byte[] CreatePhase2ValueWorkbookStream() {
                using var stream = new MemoryStream();
                WriteRecord(stream, 0x0809, new byte[] { 0x00, 0x06, 0x05, 0x00, 0xdb, 0x0b, 0xcc, 0x07 });
                long boundSheetPosition = stream.Position;
                WriteRecord(stream, 0x0085, BuildBoundSheetPayload(0, "Values"));
                WriteRecord(stream, 0x0022, new byte[] { 0x01, 0x00 });
                WriteRecord(stream, 0x000a, Array.Empty<byte>());

                int sheetOffset = checked((int)stream.Position);
                WriteRecord(stream, 0x0809, new byte[] { 0x00, 0x06, 0x10, 0x00, 0xdb, 0x0b, 0xcc, 0x07 });
                WriteRecord(stream, 0x0204, BuildLabelPayload(0, 0, "Inline"));
                WriteRecord(stream, 0x027e, BuildRkPayload(1, 0, 0, EncodeRkInteger(7)));
                WriteRecord(stream, 0x00bd, BuildMulRkPayload(1, 1, (0, EncodeRkInteger(-3)), (0, EncodeRkInteger(12345, divideBy100: true))));
                WriteRecord(stream, 0x00bd, BuildMulRkPayload(2, 0, (0, EncodeRkDouble(1d)), (0, EncodeRkDouble(2d))));
                WriteRecord(stream, 0x00be, BuildMulBlankPayload(3, 0, 5, 6));
                WriteRecord(stream, 0x0205, BuildBoolErrPayload(4, 0, 0x07, isError: true));
                WriteRecord(stream, 0x000a, Array.Empty<byte>());

                byte[] bytes = stream.ToArray();
                byte[] offsetBytes = BitConverter.GetBytes(sheetOffset);
                Buffer.BlockCopy(offsetBytes, 0, bytes, checked((int)boundSheetPosition + 4), offsetBytes.Length);
                return bytes;
            }

            internal static byte[] CreateDate1904DateFormattedWorkbookStream() {
                using var stream = new MemoryStream();
                WriteRecord(stream, 0x0809, new byte[] { 0x00, 0x06, 0x05, 0x00, 0xdb, 0x0b, 0xcc, 0x07 });
                long boundSheetPosition = stream.Position;
                WriteRecord(stream, 0x0085, BuildBoundSheetPayload(0, "Dates"));
                WriteRecord(stream, 0x0022, new byte[] { 0x01, 0x00 });
                WriteRecord(stream, 0x041e, BuildFormatPayload(164, "yyyy-mm-dd"));
                WriteRecord(stream, 0x00e0, BuildXfPayload(0));
                WriteRecord(stream, 0x00e0, BuildXfPayload(164));
                WriteRecord(stream, 0x000a, Array.Empty<byte>());

                int sheetOffset = checked((int)stream.Position);
                WriteRecord(stream, 0x0809, new byte[] { 0x00, 0x06, 0x10, 0x00, 0xdb, 0x0b, 0xcc, 0x07 });
                WriteRecord(stream, 0x0203, BuildNumberPayload(0, 0, 1d, styleIndex: 1));
                WriteRecord(stream, 0x000a, Array.Empty<byte>());

                byte[] bytes = stream.ToArray();
                Buffer.BlockCopy(BitConverter.GetBytes(sheetOffset), 0, bytes, checked((int)boundSheetPosition + 4), 4);
                return bytes;
            }

            internal static byte[] CreateContinuedSharedStringWorkbookStream() {
                using var stream = new MemoryStream();
                WriteRecord(stream, 0x0809, new byte[] { 0x00, 0x06, 0x05, 0x00, 0xdb, 0x0b, 0xcc, 0x07 });
                long boundSheetPosition = stream.Position;
                WriteRecord(stream, 0x0085, BuildBoundSheetPayload(0, "Strings"));
                WriteRecord(stream, 0x00fc, BuildSstPayload(new[] { "First" }, totalCount: 2, uniqueCount: 2));
                WriteRecord(stream, 0x003c, BuildSharedStringEntriesPayload("Second"));
                WriteRecord(stream, 0x000a, Array.Empty<byte>());

                int sheetOffset = checked((int)stream.Position);
                WriteRecord(stream, 0x0809, new byte[] { 0x00, 0x06, 0x10, 0x00, 0xdb, 0x0b, 0xcc, 0x07 });
                WriteRecord(stream, 0x00fd, BuildLabelSstPayload(0, 0, 0));
                WriteRecord(stream, 0x00fd, BuildLabelSstPayload(0, 1, 1));
                WriteRecord(stream, 0x000a, Array.Empty<byte>());

                byte[] bytes = stream.ToArray();
                Buffer.BlockCopy(BitConverter.GetBytes(sheetOffset), 0, bytes, checked((int)boundSheetPosition + 4), 4);
                return bytes;
            }

            private static byte[] BuildBoundSheetPayload(int streamOffset, string name, byte visibility = 0, byte sheetType = 0) {
                byte[] nameBytes = Encoding.ASCII.GetBytes(name);
                byte[] payload = new byte[8 + nameBytes.Length];
                Buffer.BlockCopy(BitConverter.GetBytes(streamOffset), 0, payload, 0, 4);
                payload[4] = visibility;
                payload[5] = sheetType;
                payload[6] = (byte)nameBytes.Length;
                payload[7] = 0;
                Buffer.BlockCopy(nameBytes, 0, payload, 8, nameBytes.Length);
                return payload;
            }

            private static byte[] BuildSstPayload(params string[] texts) {
                return BuildSstPayload(texts, totalCount: (uint)texts.Length, uniqueCount: (uint)texts.Length);
            }

            private static byte[] BuildSstPayload(string[] texts, uint totalCount, uint uniqueCount) {
                using var stream = new MemoryStream();
                WriteUInt32(stream, totalCount);
                WriteUInt32(stream, uniqueCount);
                byte[] entries = BuildSharedStringEntriesPayload(texts);
                stream.Write(entries, 0, entries.Length);
                return stream.ToArray();
            }

            private static byte[] BuildSharedStringEntriesPayload(params string[] texts) {
                using var stream = new MemoryStream();
                foreach (string text in texts) {
                    WriteSharedStringEntry(stream, text);
                }

                return stream.ToArray();
            }

            private static void WriteSharedStringEntry(Stream stream, string text) {
                byte[] textBytes = Encoding.ASCII.GetBytes(text);
                WriteUInt16(stream, (ushort)textBytes.Length);
                stream.WriteByte(0);
                stream.Write(textBytes, 0, textBytes.Length);
            }

            private static byte[] BuildFormatPayload(ushort formatId, string formatCode) {
                byte[] formatBytes = Encoding.ASCII.GetBytes(formatCode);
                using var stream = new MemoryStream();
                WriteUInt16(stream, formatId);
                WriteUInt16(stream, checked((ushort)formatCode.Length));
                stream.WriteByte(0);
                stream.Write(formatBytes, 0, formatBytes.Length);
                return stream.ToArray();
            }

            private static byte[] BuildXfPayload(
                ushort numberFormatId,
                ushort fontIndex = 0,
                bool isStyle = false,
                ushort parentStyleIndex = 0,
                bool? applyNumberFormat = null,
                bool? applyFont = null,
                byte fillPattern = 0,
                ushort fillForegroundColorIndex = 0x0040,
                ushort fillBackgroundColorIndex = 0x0041,
                bool applyAlignment = false,
                byte horizontalAlignment = 0,
                byte verticalAlignment = 2,
                bool wrapText = false,
                byte textRotation = 0,
                byte indent = 0,
                bool shrinkToFit = false,
                byte readingOrder = 0,
                bool locked = true,
                bool formulaHidden = false,
                bool applyProtection = false,
                bool quotePrefix = false,
                byte leftBorderStyle = 0,
                ushort leftBorderColorIndex = 0,
                byte rightBorderStyle = 0,
                ushort rightBorderColorIndex = 0,
                byte topBorderStyle = 0,
                ushort topBorderColorIndex = 0,
                byte bottomBorderStyle = 0,
                ushort bottomBorderColorIndex = 0,
                byte diagonalBorderStyle = 0,
                ushort diagonalBorderColorIndex = 0,
                byte diagonalFlags = 0) {
                byte[] payload = new byte[20];
                WriteUInt16(payload, 0, fontIndex);
                WriteUInt16(payload, 2, numberFormatId);
                ushort protection = 0;
                if (locked) {
                    protection |= 0x0001;
                }

                if (formulaHidden) {
                    protection |= 0x0002;
                }

                if (quotePrefix) {
                    protection |= 0x0008;
                }

                if (isStyle) {
                    protection |= 0x0004;
                }

                protection |= (ushort)((parentStyleIndex & 0x0fff) << 4);
                WriteUInt16(payload, 4, protection);
                payload[6] = (byte)((horizontalAlignment & 0x07) | (wrapText ? 0x08 : 0) | ((verticalAlignment & 0x07) << 4));
                payload[7] = textRotation;
                ushort attributes = 0;
                if (applyNumberFormat ?? numberFormatId != 0) {
                    attributes |= 0x0400;
                }

                if (applyFont ?? fontIndex != 0) {
                    attributes |= 0x0800;
                }

                if (applyAlignment) {
                    attributes |= 0x1000;
                }

                if (fillPattern != 0) {
                    attributes |= 0x4000;
                }

                if (leftBorderStyle != 0 || rightBorderStyle != 0 || topBorderStyle != 0 || bottomBorderStyle != 0 || diagonalBorderStyle != 0) {
                    attributes |= 0x2000;
                }

                if (applyProtection) {
                    attributes |= 0x8000;
                }

                ushort extendedAlignment = (ushort)((indent & 0x0f) | (shrinkToFit ? 0x10 : 0) | ((readingOrder & 0x03) << 6));
                WriteUInt16(payload, 8, (ushort)(extendedAlignment | attributes));
                uint sideBorderBits = (uint)(leftBorderStyle & 0x0f)
                    | ((uint)(rightBorderStyle & 0x0f) << 4)
                    | ((uint)(topBorderStyle & 0x0f) << 8)
                    | ((uint)(bottomBorderStyle & 0x0f) << 12)
                    | ((uint)(leftBorderColorIndex & 0x7f) << 16)
                    | ((uint)(rightBorderColorIndex & 0x7f) << 23)
                    | ((uint)(diagonalFlags & 0x03) << 30);
                WriteUInt32(payload, 10, sideBorderBits);
                uint topBottomBorderBits = (uint)(topBorderColorIndex & 0x7f)
                    | ((uint)(bottomBorderColorIndex & 0x7f) << 7)
                    | ((uint)(diagonalBorderColorIndex & 0x7f) << 14)
                    | ((uint)(diagonalBorderStyle & 0x0f) << 21)
                    | ((uint)(fillPattern & 0x3f) << 26);
                ushort fillColors = (ushort)((fillForegroundColorIndex & 0x7f) | ((fillBackgroundColorIndex & 0x7f) << 7));
                WriteUInt32(payload, 14, topBottomBorderBits);
                WriteUInt16(payload, 18, fillColors);
                return payload;
            }

            private static byte[] BuildFormulaNumberPayload(ushort row, ushort column, double value, ushort styleIndex = 0, byte[]? formulaTokens = null) {
                byte[] payload = BuildFormulaPayload(row, column, styleIndex, formulaTokens);
                byte[] numberBytes = BitConverter.GetBytes(value);
                Buffer.BlockCopy(numberBytes, 0, payload, 6, numberBytes.Length);
                return payload;
            }

            private static byte[] BuildFormulaSpecialPayload(ushort row, ushort column, byte valueType, byte value, ushort styleIndex = 0) {
                byte[] payload = BuildFormulaPayload(row, column, styleIndex);
                payload[6] = valueType;
                payload[8] = value;
                WriteUInt16(payload, 12, 0xffff);
                return payload;
            }

            private static byte[] BuildFormulaPayload(ushort row, ushort column, ushort styleIndex, byte[]? formulaTokens = null) {
                byte[] tokens = formulaTokens ?? Array.Empty<byte>();
                byte[] payload = new byte[checked(22 + tokens.Length)];
                WriteUInt16(payload, 0, row);
                WriteUInt16(payload, 2, column);
                WriteUInt16(payload, 4, styleIndex);
                WriteUInt16(payload, 20, checked((ushort)tokens.Length));
                if (tokens.Length > 0) {
                    Buffer.BlockCopy(tokens, 0, payload, 22, tokens.Length);
                }

                return payload;
            }

            private static byte[] BuildReferenceAdditionFormulaTokens(ushort leftRow, ushort leftColumn, ushort rightRow, ushort rightColumn) {
                using var stream = new MemoryStream();
                byte[] left = BuildCellReferenceFormulaToken(leftRow, leftColumn);
                byte[] right = BuildCellReferenceFormulaToken(rightRow, rightColumn);
                stream.Write(left, 0, left.Length);
                stream.Write(right, 0, right.Length);
                stream.WriteByte(0x03);
                return stream.ToArray();
            }

            private static byte[] BuildSumAreaFormulaTokens(ushort firstRow, ushort firstColumn, ushort lastRow, ushort lastColumn) {
                using var stream = new MemoryStream();
                byte[] area = BuildAreaReferenceFormulaToken(firstRow, firstColumn, lastRow, lastColumn);
                stream.Write(area, 0, area.Length);
                stream.WriteByte(0x42);
                stream.WriteByte(0x01);
                WriteUInt16(stream, 0x0004);
                return stream.ToArray();
            }

            private static byte[] BuildCellReferenceFormulaToken(ushort zeroBasedRow, ushort zeroBasedColumn) {
                byte[] token = new byte[5];
                token[0] = 0x44;
                WriteUInt16(token, 1, zeroBasedRow);
                WriteUInt16(token, 3, (ushort)(zeroBasedColumn | 0xc000));
                return token;
            }

            private static byte[] BuildAreaReferenceFormulaToken(ushort firstRow, ushort firstColumn, ushort lastRow, ushort lastColumn) {
                byte[] token = new byte[9];
                token[0] = 0x45;
                WriteUInt16(token, 1, firstRow);
                WriteUInt16(token, 3, lastRow);
                WriteUInt16(token, 5, (ushort)(firstColumn | 0xc000));
                WriteUInt16(token, 7, (ushort)(lastColumn | 0xc000));
                return token;
            }

            private static byte[] BuildFormulaStringPayload(string text) {
                byte[] textBytes = Encoding.ASCII.GetBytes(text);
                using var stream = new MemoryStream();
                WriteUInt16(stream, checked((ushort)text.Length));
                stream.WriteByte(0);
                stream.Write(textBytes, 0, textBytes.Length);
                return stream.ToArray();
            }

            private static byte[] BuildLabelSstPayload(ushort row, ushort column, uint sharedStringIndex) {
                using var stream = new MemoryStream();
                WriteUInt16(stream, row);
                WriteUInt16(stream, column);
                WriteUInt16(stream, 0);
                WriteUInt32(stream, sharedStringIndex);
                return stream.ToArray();
            }

            private static byte[] BuildDimensionsPayload(uint firstRow, uint rowAfterLast, ushort firstColumn, ushort columnAfterLast) {
                using var stream = new MemoryStream();
                WriteUInt32(stream, firstRow);
                WriteUInt32(stream, rowAfterLast);
                WriteUInt16(stream, firstColumn);
                WriteUInt16(stream, columnAfterLast);
                WriteUInt16(stream, 0);
                return stream.ToArray();
            }

            private static byte[] BuildBlankPayload(ushort row, ushort column, ushort styleIndex) {
                using var stream = new MemoryStream();
                WriteUInt16(stream, row);
                WriteUInt16(stream, column);
                WriteUInt16(stream, styleIndex);
                return stream.ToArray();
            }

            private static byte[] BuildColInfoPayload(ushort firstColumn, ushort lastColumn, double width, bool hidden, ushort styleIndex, byte outlineLevel = 0, bool collapsed = false) {
                using var stream = new MemoryStream();
                WriteUInt16(stream, firstColumn);
                WriteUInt16(stream, lastColumn);
                WriteUInt16(stream, checked((ushort)Math.Round(width * 256d)));
                WriteUInt16(stream, styleIndex);
                ushort options = hidden ? (ushort)0x0001 : (ushort)0x0000;
                options |= (ushort)((outlineLevel & 0x07) << 8);
                if (collapsed) {
                    options |= 0x1000;
                }

                WriteUInt16(stream, options);
                WriteUInt16(stream, 0);
                return stream.ToArray();
            }

            private static byte[] BuildLabelPayload(ushort row, ushort column, string text, ushort styleIndex = 0) {
                byte[] textBytes = Encoding.ASCII.GetBytes(text);
                using var stream = new MemoryStream();
                WriteUInt16(stream, row);
                WriteUInt16(stream, column);
                WriteUInt16(stream, styleIndex);
                WriteUInt16(stream, (ushort)textBytes.Length);
                stream.WriteByte(0);
                stream.Write(textBytes, 0, textBytes.Length);
                return stream.ToArray();
            }

            private static byte[] BuildFontPayload(string name, double size, bool bold, bool italic, bool underline, bool strikeout = false, ushort colorIndex = 0x7FFF, ushort escapement = 0) {
                byte[] nameBytes = Encoding.ASCII.GetBytes(name);
                using var stream = new MemoryStream();
                WriteUInt16(stream, checked((ushort)Math.Round(size * 20d)));
                ushort options = 0;
                if (italic) {
                    options |= 0x0002;
                }

                if (strikeout) {
                    options |= 0x0008;
                }

                WriteUInt16(stream, options);
                WriteUInt16(stream, colorIndex);
                WriteUInt16(stream, bold ? (ushort)700 : (ushort)400);
                WriteUInt16(stream, escapement);
                stream.WriteByte(underline ? (byte)1 : (byte)0);
                stream.WriteByte(2);
                stream.WriteByte(0);
                stream.WriteByte(0);
                stream.WriteByte(checked((byte)name.Length));
                stream.WriteByte(0);
                stream.Write(nameBytes, 0, nameBytes.Length);
                return stream.ToArray();
            }

            private static byte[] BuildPalettePayload(params string[] argbColors) {
                using var stream = new MemoryStream();
                WriteUInt16(stream, 56);
                for (int i = 0; i < 56; i++) {
                    string rgb = i < argbColors.Length ? NormalizePaletteColor(argbColors[i]) : "000000";
                    stream.WriteByte(byte.Parse(rgb.Substring(0, 2), NumberStyles.HexNumber, CultureInfo.InvariantCulture));
                    stream.WriteByte(byte.Parse(rgb.Substring(2, 2), NumberStyles.HexNumber, CultureInfo.InvariantCulture));
                    stream.WriteByte(byte.Parse(rgb.Substring(4, 2), NumberStyles.HexNumber, CultureInfo.InvariantCulture));
                    stream.WriteByte(0);
                }

                return stream.ToArray();
            }

            private static string NormalizePaletteColor(string argbColor) {
                string color = argbColor.TrimStart('#');
                return color.Length == 8 ? color.Substring(2) : color;
            }

            private static byte[] BuildMergeCellsPayload(params (ushort FirstRow, ushort FirstColumn, ushort LastRow, ushort LastColumn)[] ranges) {
                using var stream = new MemoryStream();
                WriteUInt16(stream, checked((ushort)ranges.Length));
                foreach ((ushort firstRow, ushort firstColumn, ushort lastRow, ushort lastColumn) in ranges) {
                    WriteUInt16(stream, firstRow);
                    WriteUInt16(stream, lastRow);
                    WriteUInt16(stream, firstColumn);
                    WriteUInt16(stream, lastColumn);
                }

                return stream.ToArray();
            }

            private static byte[] BuildPanePayload(ushort leftColumns, ushort topRows) {
                using var stream = new MemoryStream();
                WriteUInt16(stream, leftColumns);
                WriteUInt16(stream, topRows);
                WriteUInt16(stream, topRows);
                WriteUInt16(stream, leftColumns);
                stream.WriteByte(0);
                stream.WriteByte(0);
                return stream.ToArray();
            }

            private static byte[] BuildDefaultColumnWidthPayload(ushort width) {
                using var stream = new MemoryStream();
                WriteUInt16(stream, width);
                return stream.ToArray();
            }

            private static byte[] BuildDefaultRowHeightPayload(double height, bool hidden = false) {
                using var stream = new MemoryStream();
                ushort options = hidden ? (ushort)0x0002 : (ushort)0;
                WriteUInt16(stream, options);
                WriteUInt16(stream, checked((ushort)Math.Round(height * 20d)));
                return stream.ToArray();
            }

            private static byte[] BuildRowPayload(ushort row, double height, bool hidden, bool customHeight, byte outlineLevel = 0, bool collapsed = false, ushort? styleIndex = null) {
                using var stream = new MemoryStream();
                WriteUInt16(stream, row);
                WriteUInt16(stream, 0);
                WriteUInt16(stream, 1);
                WriteUInt16(stream, checked((ushort)Math.Round(height * 20d)));
                WriteUInt16(stream, 0);
                WriteUInt16(stream, 0);
                ushort options = (ushort)(outlineLevel & 0x07);
                if (collapsed) {
                    options |= 0x0010;
                }

                if (hidden) {
                    options |= 0x0020;
                }

                if (customHeight) {
                    options |= 0x0040;
                }

                if (styleIndex.HasValue) {
                    options |= 0x0080;
                }

                WriteUInt16(stream, options);
                WriteUInt16(stream, styleIndex.HasValue ? (ushort)(styleIndex.Value & 0x0fff) : (ushort)0x0100);
                return stream.ToArray();
            }

            private static byte[] BuildWindow2Payload(bool frozen, bool showGridlines = true, bool showRowColumnHeadings = true, bool showZeroValues = true, bool rightToLeft = false) {
                using var stream = new MemoryStream();
                ushort options = 0;
                if (showGridlines) {
                    options |= 0x0002;
                }

                if (showRowColumnHeadings) {
                    options |= 0x0004;
                }

                if (frozen) {
                    options |= 0x0008;
                }

                if (showZeroValues) {
                    options |= 0x0010;
                }

                if (rightToLeft) {
                    options |= 0x0040;
                }

                WriteUInt16(stream, options);
                WriteUInt16(stream, 0);
                WriteUInt16(stream, 0);
                WriteUInt16(stream, 64);
                WriteUInt16(stream, 0);
                WriteUInt16(stream, 0);
                WriteUInt16(stream, 0);
                WriteUInt16(stream, 0);
                WriteUInt16(stream, 0);
                return stream.ToArray();
            }

            private static byte[] BuildNumberPayload(ushort row, ushort column, double value, ushort styleIndex = 0) {
                using var stream = new MemoryStream();
                WriteUInt16(stream, row);
                WriteUInt16(stream, column);
                WriteUInt16(stream, styleIndex);
                byte[] number = BitConverter.GetBytes(value);
                stream.Write(number, 0, number.Length);
                return stream.ToArray();
            }

            private static byte[] BuildBoolErrPayload(ushort row, ushort column, bool value) {
                return BuildBoolErrPayload(row, column, value ? (byte)1 : (byte)0, isError: false);
            }

            private static byte[] BuildBoolErrPayload(ushort row, ushort column, byte value, bool isError) {
                using var stream = new MemoryStream();
                WriteUInt16(stream, row);
                WriteUInt16(stream, column);
                WriteUInt16(stream, 0);
                stream.WriteByte(value);
                stream.WriteByte(isError ? (byte)1 : (byte)0);
                return stream.ToArray();
            }

            private static byte[] BuildRkPayload(ushort row, ushort column, ushort styleIndex, uint rkValue) {
                using var stream = new MemoryStream();
                WriteUInt16(stream, row);
                WriteUInt16(stream, column);
                WriteUInt16(stream, styleIndex);
                WriteUInt32(stream, rkValue);
                return stream.ToArray();
            }

            private static byte[] BuildMulBlankPayload(ushort row, ushort firstColumn, params ushort[] styleIndexes) {
                using var stream = new MemoryStream();
                WriteUInt16(stream, row);
                WriteUInt16(stream, firstColumn);
                foreach (ushort styleIndex in styleIndexes) {
                    WriteUInt16(stream, styleIndex);
                }

                WriteUInt16(stream, checked((ushort)(firstColumn + styleIndexes.Length - 1)));
                return stream.ToArray();
            }

            private static byte[] BuildMulRkPayload(ushort row, ushort firstColumn, params (ushort StyleIndex, uint RkValue)[] cells) {
                using var stream = new MemoryStream();
                WriteUInt16(stream, row);
                WriteUInt16(stream, firstColumn);
                foreach ((ushort styleIndex, uint rkValue) in cells) {
                    WriteUInt16(stream, styleIndex);
                    WriteUInt32(stream, rkValue);
                }

                WriteUInt16(stream, checked((ushort)(firstColumn + cells.Length - 1)));
                return stream.ToArray();
            }

            private static uint EncodeRkDouble(double value) {
                ulong bits = BitConverter.ToUInt64(BitConverter.GetBytes(value), 0);
                return (uint)(bits >> 32) & 0xfffffffc;
            }

            private static uint EncodeRkInteger(int value, bool divideBy100 = false) {
                uint encoded = unchecked((uint)(value << 2)) | 0x02;
                return divideBy100 ? encoded | 0x01 : encoded;
            }

            private static void WriteRecord(Stream stream, ushort type, byte[] payload) {
                WriteUInt16(stream, type);
                WriteUInt16(stream, (ushort)payload.Length);
                stream.Write(payload, 0, payload.Length);
            }
        }

        private static void WriteUInt16(Stream stream, ushort value) {
            stream.WriteByte((byte)(value & 0xff));
            stream.WriteByte((byte)((value >> 8) & 0xff));
        }

        private static void WriteUInt32(Stream stream, uint value) {
            stream.WriteByte((byte)(value & 0xff));
            stream.WriteByte((byte)((value >> 8) & 0xff));
            stream.WriteByte((byte)((value >> 16) & 0xff));
            stream.WriteByte((byte)((value >> 24) & 0xff));
        }

        private static void WriteUInt16(byte[] buffer, int offset, ushort value) {
            buffer[offset] = (byte)(value & 0xff);
            buffer[offset + 1] = (byte)((value >> 8) & 0xff);
        }

        private static void WriteUInt32(byte[] buffer, int offset, uint value) {
            buffer[offset] = (byte)(value & 0xff);
            buffer[offset + 1] = (byte)((value >> 8) & 0xff);
            buffer[offset + 2] = (byte)((value >> 16) & 0xff);
            buffer[offset + 3] = (byte)((value >> 24) & 0xff);
        }

        private static void WriteUInt64(byte[] buffer, int offset, ulong value) {
            WriteUInt32(buffer, offset, (uint)(value & 0xffffffff));
            WriteUInt32(buffer, offset + 4, (uint)(value >> 32));
        }
    }
}
