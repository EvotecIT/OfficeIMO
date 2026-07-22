using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using OfficeIMO.Excel;
using OfficeIMO.Excel.LegacyXls;
using OfficeIMO.Excel.LegacyXls.Biff;
using OfficeIMO.Excel.LegacyXls.Compound;
using OfficeIMO.Excel.LegacyXls.Diagnostics;
using OfficeIMO.Excel.LegacyXls.Model;
using System.Globalization;
using Xunit;

namespace OfficeIMO.Tests {
    public partial class Excel {
        [Fact]
        public void LegacyXls_Load_RejectsNonPositiveDecodedImageBudget() {
            byte[] workbookStream = LegacyXlsTestWorkbookBuilder.CreateMinimalWorkbookStream();
            byte[] compound = LegacyXlsCompoundTestBuilder.CreateWorkbookCompoundFile(workbookStream);

            Assert.Throws<ArgumentOutOfRangeException>(() =>
                LegacyXlsWorkbook.Load(compound,
                    new LegacyXlsImportOptions { MaxDecodedImageBytes = 0 }));
        }

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
                ReportUnsupportedContent = false
            });
            LegacyXlsImportReport report = legacy.CreateImportReport();

            Assert.DoesNotContain(legacy.Diagnostics, d => d.Severity == LegacyXlsDiagnosticSeverity.Error);
            LegacyXlsWorksheet sheet = Assert.Single(legacy.Worksheets);
            Assert.Equal("Sheet1", sheet.Name);
            Assert.Contains(sheet.Cells, cell => cell.Row == 1 && cell.Column == 1 && Equals(cell.Value, "Name"));
            Assert.Contains(sheet.Cells, cell => cell.Row == 2 && cell.Column == 2 && Equals(cell.Value, 42d));
            Assert.Contains(sheet.Cells, cell => cell.Row == 3 && cell.Column == 1 && Equals(cell.Value, true));
            Assert.Equal(1, report.FileFormatStates["WorkbookFormat:SupportedBiff8"]);
            Assert.Equal(1, report.FileFormatStates["Encryption:Missing"]);
            Assert.Equal(1, report.FileFormatStates["UnsupportedBiffVersion:Missing"]);
            Assert.Equal(1, report.FileFormatStates["MalformedBof:Missing"]);
        }

        [Fact]
        public void LegacyXls_Load_ImportsWorkbookStreamFromCompoundMiniStream() {
            byte[] workbookStream = LegacyXlsTestWorkbookBuilder.CreateMinimalWorkbookStream();
            Assert.True(workbookStream.Length < 4096);
            byte[] compound = LegacyXlsCompoundTestBuilder.CreateMiniStreamWorkbookCompoundFile(workbookStream);

            LegacyXlsWorkbook legacy = LegacyXlsWorkbook.Load(compound, new LegacyXlsImportOptions {
                ReportUnsupportedContent = true
            });

            Assert.DoesNotContain(legacy.Diagnostics, d => d.Severity == LegacyXlsDiagnosticSeverity.Error);
            LegacyXlsWorksheet sheet = Assert.Single(legacy.Worksheets);
            Assert.Equal("Sheet1", sheet.Name);
            Assert.Contains(sheet.Cells, cell => cell.Row == 1 && cell.Column == 1 && Equals(cell.Value, "Name"));

            using ExcelDocument document = ExcelDocument.LoadLegacyXls(new MemoryStream(compound), new LegacyXlsImportOptions {
                ReportUnsupportedContent = true
            });

            Assert.True(document.Sheets[0].TryGetCellText(2, 2, out string? amount));
            Assert.Equal("42", amount);
        }

        [Fact]
        public void LegacyXls_Load_ImportsWorkbookStreamFromCompoundDifatSector() {
            byte[] workbookStream = LegacyXlsTestWorkbookBuilder.CreateMinimalWorkbookStream();
            byte[] compound = LegacyXlsCompoundTestBuilder.CreateDifatWorkbookCompoundFile(workbookStream);

            LegacyXlsWorkbook legacy = LegacyXlsWorkbook.Load(compound, new LegacyXlsImportOptions {
                ReportUnsupportedContent = false
            });

            Assert.DoesNotContain(legacy.Diagnostics, d => d.Severity == LegacyXlsDiagnosticSeverity.Error);
            LegacyXlsWorksheet sheet = Assert.Single(legacy.Worksheets);
            Assert.Equal("Sheet1", sheet.Name);
            Assert.Contains(sheet.Cells, cell => cell.Row == 2 && cell.Column == 2 && Equals(cell.Value, 42d));

            using ExcelDocument document = ExcelDocument.LoadLegacyXls(new MemoryStream(compound), new LegacyXlsImportOptions {
                ReportUnsupportedContent = false
            });

            Assert.True(document.Sheets[0].TryGetCellText(1, 1, out string? header));
            Assert.Equal("Name", header);
        }

        [Fact]
        public void LegacyXls_Load_ReportsCorruptCompoundSectorChainsAsDiagnostics() {
            byte[] compound = LegacyXlsCompoundTestBuilder.CreateCompoundHeaderWithInvalidSectorChain();

            LegacyXlsWorkbook legacy = LegacyXlsWorkbook.Load(compound, new LegacyXlsImportOptions());

            LegacyXlsImportDiagnostic diagnostic = Assert.Single(
                legacy.Diagnostics,
                item => item.Code == "XLS-COMPOUND-CORRUPT");
            Assert.Equal(LegacyXlsDiagnosticSeverity.Error, diagnostic.Severity);
            Assert.Contains("could not be read", diagnostic.Message);
        }

        [Fact]
        public void LegacyXls_Load_ImportsPhase2ValueRecords() {
            byte[] workbookStream = LegacyXlsTestWorkbookBuilder.CreatePhase2ValueWorkbookStream();
            byte[] compound = LegacyXlsCompoundTestBuilder.CreateWorkbookCompoundFile(workbookStream);

            LegacyXlsWorkbook legacy = LegacyXlsWorkbook.Load(compound, new LegacyXlsImportOptions {
                ReportUnsupportedContent = false
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
                ReportUnsupportedContent = true
            });

            Assert.DoesNotContain(legacy.Diagnostics, d => d.Severity == LegacyXlsDiagnosticSeverity.Error);
            Assert.True(legacy.Uses1904DateSystem);
            LegacyXlsCell dateCellModel = Assert.Single(Assert.Single(legacy.Worksheets).Cells);
            Assert.Equal(1d, dateCellModel.Value);
            Assert.True(legacy.GetEffectiveCellFormat(dateCellModel.StyleIndex)!.IsDateLike);

            using ExcelDocument document = ExcelDocument.LoadLegacyXls(new MemoryStream(compound), new LegacyXlsImportOptions {
                ReportUnsupportedContent = true
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
                ReportUnsupportedContent = true
            });

            Assert.DoesNotContain(legacy.Diagnostics, d => d.Severity == LegacyXlsDiagnosticSeverity.Error);
            Assert.DoesNotContain(legacy.Diagnostics, d => d.RecordType == (ushort)BiffRecordType.Continue);
            LegacyXlsWorksheet sheet = Assert.Single(legacy.Worksheets);
            Assert.Contains(sheet.Cells, cell => cell.Row == 1 && cell.Column == 1 && Equals(cell.Value, "First"));
            Assert.Contains(sheet.Cells, cell => cell.Row == 1 && cell.Column == 2 && Equals(cell.Value, "Second"));

            using ExcelDocument document = ExcelDocument.LoadLegacyXls(new MemoryStream(compound), new LegacyXlsImportOptions {
                ReportUnsupportedContent = true
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
                ReportUnsupportedContent = true
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
                ReportUnsupportedContent = true
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
                ReportUnsupportedContent = false
            });

            Assert.DoesNotContain(legacy.Diagnostics, d => d.Severity == LegacyXlsDiagnosticSeverity.Error);
            LegacyXlsWorksheet sheet = Assert.Single(legacy.Worksheets);
            Assert.Contains(sheet.Cells, cell => cell.Row == 1 && cell.Column == 1 && cell.IsFormula && Equals(cell.Value, 42d));
            Assert.Contains(sheet.Cells, cell => cell.Row == 1 && cell.Column == 2 && cell.IsFormula && Equals(cell.Value, true));
            Assert.Contains(sheet.Cells, cell => cell.Row == 1 && cell.Column == 3 && cell.IsFormula && Equals(cell.Value, "Formula text"));
            Assert.Contains(sheet.Cells, cell => cell.Row == 1 && cell.Column == 4 && cell.IsFormula && cell.Kind == LegacyXlsCellValueKind.Error && Equals(cell.Value, "#VALUE!"));
            Assert.Contains(sheet.Cells, cell => cell.Row == 1 && cell.Column == 5 && cell.IsFormula && Equals(cell.Value, "Continued formula text"));
            LegacyXlsCell decodedFormula = Assert.Single(sheet.Cells, cell => cell.Row == 2 && cell.Column == 3);
            Assert.True(decodedFormula.IsFormula);
            Assert.Equal(42d, decodedFormula.Value);
            Assert.Equal("A2+B2", decodedFormula.FormulaText);
            LegacyXlsCell decodedSumFormula = Assert.Single(sheet.Cells, cell => cell.Row == 2 && cell.Column == 4);
            Assert.True(decodedSumFormula.IsFormula);
            Assert.Equal(42d, decodedSumFormula.Value);
            Assert.Equal("SUM(A2:B2)", decodedSumFormula.FormulaText);

            using ExcelDocument document = ExcelDocument.LoadLegacyXls(new MemoryStream(compound), new LegacyXlsImportOptions {
                ReportUnsupportedContent = false
            });

            Assert.True(document.Sheets[0].TryGetCellText(1, 1, out string? number));
            Assert.Equal("42", number);
            Assert.True(document.Sheets[0].TryGetCellText(1, 2, out string? boolean));
            Assert.Equal("1", boolean);
            Assert.True(document.Sheets[0].TryGetCellText(1, 3, out string? text));
            Assert.Equal("Formula text", text);
            Assert.True(document.Sheets[0].TryGetCellText(1, 5, out string? continuedText));
            Assert.Equal("Continued formula text", continuedText);
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
            Cell projectedContinuedStringFormula = worksheetPart.Worksheet.Descendants<Cell>().Single(cell => cell.CellReference!.Value == "E1");
            Assert.Equal(CellValues.SharedString, projectedContinuedStringFormula.DataType!.Value);
            int sharedStringIndex = int.Parse(projectedContinuedStringFormula.CellValue!.Text, CultureInfo.InvariantCulture);
            Assert.Equal("Continued formula text", spreadsheet.WorkbookPart.SharedStringTablePart!.SharedStringTable.Elements<SharedStringItem>().ElementAt(sharedStringIndex).InnerText);
        }

        [Fact]
        public void LegacyXls_Load_ReportsEncryptedWorkbookAsUnsupported() {
            byte[] workbookStream = LegacyXlsTestWorkbookBuilder.CreateEncryptedWorkbookStream();
            byte[] compound = LegacyXlsCompoundTestBuilder.CreateWorkbookCompoundFile(workbookStream);

            LegacyXlsWorkbook legacy = LegacyXlsWorkbook.Load(compound);
            LegacyXlsImportReport report = legacy.CreateImportReport();

            Assert.Contains(legacy.Diagnostics, d =>
                d.Code == "XLS-BIFF-FILEPASS-UNSUPPORTED"
                && d.Severity == LegacyXlsDiagnosticSeverity.Error
                && d.DetailCode == "Encryption:FilePass:XorObfuscation"
                && d.Message.Contains("XorObfuscation"));
            Assert.Equal(1, report.ErrorCount);
            Assert.Equal(1, report.UnsupportedFeaturesByKind[LegacyXlsUnsupportedFeatureKind.EncryptedWorkbook]);
            Assert.Equal(1, report.UnsupportedFeaturesByDetail["EncryptedWorkbook|XLS-BIFF-FILEPASS-UNSUPPORTED|Encryption:FilePass:XorObfuscation"]);
            Assert.Equal(1, report.PreservedFeatureRecordCount);
            Assert.Equal(1, report.PreservedFeatureRecordsByKind[LegacyXlsUnsupportedFeatureKind.EncryptedWorkbook]);
            Assert.Equal(0, report.UnsupportedProjectionGapCount);
            Assert.Empty(report.UnsupportedProjectionGapsByKind);
            Assert.Empty(report.UnsupportedProjectionGapsByDetail);
            Assert.Equal(1, report.FileFormatStates["WorkbookFormat:Encrypted"]);
            Assert.Equal(1, report.FileFormatStates["Encryption:Present"]);
            Assert.Equal(1, report.FileFormatStates["UnsupportedBiffVersion:Missing"]);
            Assert.Equal(1, report.FileFormatBlockers["EncryptedWorkbook|Encryption:FilePass:XorObfuscation"]);
            Assert.Equal(1, report.FileFormatBlockersByRecordType["EncryptedWorkbook|0x002F"]);
            Assert.Equal(1, report.FileFormatBlockersByRecordName["EncryptedWorkbook|Record0x002F"]);
            Assert.Equal(1, report.FileFormatBlockersByLocation["XLS-BIFF-FILEPASS-UNSUPPORTED|(workbook)"]);
            Assert.Equal(1, report.EncryptedWorkbooksByMethod["XorObfuscation"]);
            string markdown = report.ToMarkdown();
            Assert.Contains("File Format States", markdown);
            Assert.Contains("File Format Blockers", markdown);
            Assert.Contains("File Format Blockers By Record Type", markdown);
            Assert.Contains("File Format Blockers By Record Name", markdown);
            Assert.Contains("File Format Blockers By Location", markdown);
            Assert.Contains("Encrypted Workbooks By Method", markdown);
            Assert.Empty(legacy.Worksheets);
        }

        [Fact]
        public void LegacyXls_XorPasswordVerifier_MatchesKnownExcelVector() {
            Assert.Equal(0x9A0A, BiffXorObfuscation.CreatePasswordVerifier("VelvetSweatshop"));
        }

        [Fact]
        public void LegacyXls_XorPasswordHelpers_TruncateLongPasswordsBeforeDerivingKeys() {
            string maxLengthPassword = "123456789012345";
            string longPassword = maxLengthPassword + "67890";

            Assert.Equal(BiffXorObfuscation.CreateXorKey(maxLengthPassword), BiffXorObfuscation.CreateXorKey(longPassword));
            Assert.Equal(BiffXorObfuscation.CreatePasswordVerifier(maxLengthPassword), BiffXorObfuscation.CreatePasswordVerifier(longPassword));
            Assert.Equal(
                BiffXorObfuscation.ObfuscateWorkbookStream(Array.Empty<byte>(), maxLengthPassword),
                BiffXorObfuscation.ObfuscateWorkbookStream(Array.Empty<byte>(), longPassword));
        }

        [Fact]
        public void LegacyXls_Load_ImportsXorObfuscatedWorkbookWithPassword() {
            byte[] workbookStream = LegacyXlsTestWorkbookBuilder.CreateXorObfuscatedWorkbookStream("openpass");
            byte[] compound = LegacyXlsCompoundTestBuilder.CreateWorkbookCompoundFile(workbookStream);

            using LegacyXlsLoadResult result = ExcelDocument.LoadLegacyXlsWithReport(new MemoryStream(compound), new LegacyXlsImportOptions {
                ReportUnsupportedContent = true,
                Password = "openpass"
            });

            Assert.True(result.HasDocument);
            Assert.False(result.HasImportErrors);
            Assert.Equal(1, result.ImportReport.WorksheetCount);
            Assert.Equal(0, result.ImportReport.UnsupportedFeatureCount);
            Assert.Empty(result.Workbook.UnsupportedFeatures);
            Assert.DoesNotContain(result.Workbook.Diagnostics, diagnostic => diagnostic.Code.StartsWith("XLS-BIFF-FILEPASS", StringComparison.Ordinal));

            LegacyXlsWorksheet sheet = Assert.Single(result.Workbook.Worksheets);
            Assert.Equal("XorSheet", sheet.Name);
            Assert.Contains(sheet.Cells, cell => cell.Row == 1 && cell.Column == 1 && Equals(cell.Value, "XOR secret"));

            using var output = new MemoryStream();
            result.Document.Save(output);
            using SpreadsheetDocument spreadsheet = SpreadsheetDocument.Open(new MemoryStream(output.ToArray()), false);
            WorksheetPart worksheetPart = spreadsheet.WorkbookPart!.WorksheetParts.Single();
            Cell projectedCell = worksheetPart.Worksheet.Descendants<Cell>().Single(cell => cell.CellReference!.Value == "A1");
            Assert.Equal(CellValues.SharedString, projectedCell.DataType!.Value);
            int sharedStringIndex = int.Parse(projectedCell.CellValue!.Text, CultureInfo.InvariantCulture);
            Assert.Equal("XOR secret", spreadsheet.WorkbookPart.SharedStringTablePart!.SharedStringTable.Elements<SharedStringItem>().ElementAt(sharedStringIndex).InnerText);
        }

        [Fact]
        public void LegacyXls_Load_RejectsXorObfuscatedWorkbookWithWrongPassword() {
            byte[] workbookStream = LegacyXlsTestWorkbookBuilder.CreateXorObfuscatedWorkbookStream("openpass");
            byte[] compound = LegacyXlsCompoundTestBuilder.CreateWorkbookCompoundFile(workbookStream);

            LegacyXlsWorkbook legacy = LegacyXlsWorkbook.Load(compound, new LegacyXlsImportOptions {
                ReportUnsupportedContent = true,
                Password = "wrongpass"
            });
            LegacyXlsImportReport report = legacy.CreateImportReport();

            Assert.Empty(legacy.Worksheets);
            Assert.Empty(legacy.UnsupportedFeatures);
            Assert.Contains(legacy.Diagnostics, diagnostic =>
                diagnostic.Severity == LegacyXlsDiagnosticSeverity.Error
                && diagnostic.Code == "XLS-BIFF-FILEPASS-PASSWORD-INVALID"
                && diagnostic.DetailCode == "Encryption:FilePass:XorObfuscation");
            Assert.True(report.HasImportErrors);
            Assert.Equal(0, report.UnsupportedFeatureCount);
            Assert.Equal(1, report.ErrorCount);
        }

        [Fact]
        public void LegacyXls_Load_ImportsRc4EncryptedWorkbookWithPassword() {
            byte[] workbookStream = LegacyXlsTestWorkbookBuilder.CreateRc4EncryptedWorkbookStream("openpass");
            byte[] compound = LegacyXlsCompoundTestBuilder.CreateWorkbookCompoundFile(workbookStream);

            using LegacyXlsLoadResult result = ExcelDocument.LoadLegacyXlsWithReport(new MemoryStream(compound), new LegacyXlsImportOptions {
                ReportUnsupportedContent = true,
                Password = "openpass"
            });

            Assert.True(result.HasDocument);
            Assert.False(result.HasImportErrors);
            Assert.Equal(1, result.ImportReport.WorksheetCount);
            Assert.Equal(0, result.ImportReport.UnsupportedFeatureCount);
            Assert.Empty(result.Workbook.UnsupportedFeatures);
            Assert.DoesNotContain(result.Workbook.Diagnostics, diagnostic => diagnostic.Code.StartsWith("XLS-BIFF-FILEPASS", StringComparison.Ordinal));

            LegacyXlsWorksheet sheet = Assert.Single(result.Workbook.Worksheets);
            Assert.Equal("Rc4Sheet", sheet.Name);
            Assert.Contains(sheet.Cells, cell => cell.Row == 1 && cell.Column == 1 && Equals(cell.Value, "RC4 secret"));

            using var output = new MemoryStream();
            result.Document.Save(output);
            using SpreadsheetDocument spreadsheet = SpreadsheetDocument.Open(new MemoryStream(output.ToArray()), false);
            WorksheetPart worksheetPart = spreadsheet.WorkbookPart!.WorksheetParts.Single();
            Cell projectedCell = worksheetPart.Worksheet.Descendants<Cell>().Single(cell => cell.CellReference!.Value == "A1");
            Assert.Equal(CellValues.SharedString, projectedCell.DataType!.Value);
            int sharedStringIndex = int.Parse(projectedCell.CellValue!.Text, CultureInfo.InvariantCulture);
            Assert.Equal("RC4 secret", spreadsheet.WorkbookPart.SharedStringTablePart!.SharedStringTable.Elements<SharedStringItem>().ElementAt(sharedStringIndex).InnerText);
        }

        [Fact]
        public void LegacyXls_Load_RejectsRc4EncryptedWorkbookWithWrongPassword() {
            byte[] workbookStream = LegacyXlsTestWorkbookBuilder.CreateRc4EncryptedWorkbookStream("openpass");
            byte[] compound = LegacyXlsCompoundTestBuilder.CreateWorkbookCompoundFile(workbookStream);

            LegacyXlsWorkbook legacy = LegacyXlsWorkbook.Load(compound, new LegacyXlsImportOptions {
                ReportUnsupportedContent = true,
                Password = "wrongpass"
            });
            LegacyXlsImportReport report = legacy.CreateImportReport();

            Assert.Empty(legacy.Worksheets);
            Assert.Empty(legacy.UnsupportedFeatures);
            Assert.Contains(legacy.Diagnostics, diagnostic =>
                diagnostic.Severity == LegacyXlsDiagnosticSeverity.Error
                && diagnostic.Code == "XLS-BIFF-FILEPASS-PASSWORD-INVALID"
                && diagnostic.DetailCode == "Encryption:FilePass:Rc4");
            Assert.True(report.HasImportErrors);
            Assert.Equal(0, report.UnsupportedFeatureCount);
            Assert.Equal(1, report.ErrorCount);
        }

        [Fact]
        public void LegacyXls_LoadLegacyXlsWithReport_ReturnsReportWhenLegacyWorkbookHasNoSheets() {
            byte[] workbookStream = LegacyXlsTestWorkbookBuilder.CreateEncryptedWorkbookStream();
            byte[] compound = LegacyXlsCompoundTestBuilder.CreateWorkbookCompoundFile(workbookStream);

            using LegacyXlsLoadResult result = ExcelDocument.LoadLegacyXlsWithReport(new MemoryStream(compound), new LegacyXlsImportOptions {
                ReportUnsupportedContent = true
            });

            Assert.False(result.HasDocument);
            Assert.NotNull(result.ProjectionException);
            Assert.True(result.HasImportErrors);
            Assert.Contains(result.Diagnostics, diagnostic => diagnostic.Code == "XLS-BIFF-FILEPASS-UNSUPPORTED");
            Assert.Equal(0, result.ImportReport.WorksheetCount);

            InvalidOperationException documentException = Assert.Throws<InvalidOperationException>(() => result.Document);
            Assert.Contains("No OfficeIMO Excel document", documentException.Message, StringComparison.Ordinal);
            Assert.Same(result.ProjectionException, documentException.InnerException);
        }

        [Fact]
        public void LegacyXls_Load_ReportsRc4EncryptedWorkbookMethod() {
            byte[] workbookStream = LegacyXlsTestWorkbookBuilder.CreateMalformedRc4EncryptedWorkbookStream();
            byte[] compound = LegacyXlsCompoundTestBuilder.CreateWorkbookCompoundFile(workbookStream);

            LegacyXlsWorkbook legacy = LegacyXlsWorkbook.Load(compound);
            LegacyXlsImportReport report = legacy.CreateImportReport();

            Assert.Contains(legacy.Diagnostics, d =>
                d.Code == "XLS-BIFF-FILEPASS-UNSUPPORTED"
                && d.Severity == LegacyXlsDiagnosticSeverity.Error
                && d.DetailCode == "Encryption:FilePass:Rc4"
                && d.Message.Contains("Rc4"));
            Assert.Equal(1, report.FileFormatBlockers["EncryptedWorkbook|Encryption:FilePass:Rc4"]);
            Assert.Equal(1, report.EncryptedWorkbooksByMethod["Rc4"]);
            Assert.Empty(legacy.Worksheets);
        }

        [Fact]
        public void LegacyXls_Load_StopsParsingAfterEncryptedWorkbookMarker() {
            byte[] workbookStream = LegacyXlsTestWorkbookBuilder.CreateEncryptedWorkbookWithUnreadablePayloadStream();
            byte[] compound = LegacyXlsCompoundTestBuilder.CreateWorkbookCompoundFile(workbookStream);

            LegacyXlsWorkbook legacy = LegacyXlsWorkbook.Load(compound, new LegacyXlsImportOptions {
                ReportUnsupportedContent = true
            });
            LegacyXlsImportReport report = legacy.CreateImportReport();

            LegacyXlsUnsupportedFeature feature = Assert.Single(legacy.UnsupportedFeatures);
            Assert.Equal(LegacyXlsUnsupportedFeatureKind.EncryptedWorkbook, feature.Kind);
            Assert.Equal("XLS-BIFF-FILEPASS-UNSUPPORTED", feature.Code);
            Assert.Single(legacy.Diagnostics);
            Assert.Equal(1, report.ErrorCount);
            Assert.Equal(1, report.FileFormatBlockers["EncryptedWorkbook|Encryption:FilePass:XorObfuscation"]);
            Assert.Equal(1, report.FileFormatBlockersByRecordType["EncryptedWorkbook|0x002F"]);
            Assert.Equal(1, report.FileFormatBlockersByRecordName["EncryptedWorkbook|Record0x002F"]);
            Assert.Equal(1, report.FileFormatBlockersByLocation["XLS-BIFF-FILEPASS-UNSUPPORTED|(workbook)"]);
            Assert.Equal(1, report.EncryptedWorkbooksByMethod["XorObfuscation"]);
            Assert.Empty(legacy.Worksheets);
        }

        [Fact]
        public void LegacyXls_Load_ReportsFeatureSpecificUnsupportedRecords() {
            byte[] workbookStream = LegacyXlsTestWorkbookBuilder.CreateUnsupportedFeatureWorkbookStream();
            byte[] compound = LegacyXlsCompoundTestBuilder.CreateWorkbookCompoundFile(workbookStream);

            LegacyXlsWorkbook legacy = LegacyXlsWorkbook.Load(compound, new LegacyXlsImportOptions {
                ReportUnsupportedContent = true
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
        public void LegacyXls_Load_ClassifiesKnownPreserveOnlyExtensionRecords() {
            byte[] workbookStream = LegacyXlsTestWorkbookBuilder.CreateKnownPreserveOnlyExtensionWorkbookStream();
            byte[] compound = LegacyXlsCompoundTestBuilder.CreateWorkbookCompoundFile(workbookStream);

            LegacyXlsWorkbook legacy = LegacyXlsWorkbook.Load(compound, new LegacyXlsImportOptions {
                ReportUnsupportedContent = true
            });

            Assert.DoesNotContain(legacy.Diagnostics, d => d.Severity == LegacyXlsDiagnosticSeverity.Error);
            Assert.DoesNotContain(legacy.UnsupportedFeatures, feature => feature.Kind == LegacyXlsUnsupportedFeatureKind.UnsupportedRecord);
            Assert.DoesNotContain(legacy.UnsupportedFeatures, feature => feature.Kind == LegacyXlsUnsupportedFeatureKind.WorkbookMetadata);
            Assert.Equal(10, legacy.FutureMetadataRecords.Count);
            Assert.Contains(legacy.FutureMetadataRecords, record => record.Kind == LegacyXlsWorkbookMetadataKind.RecalculationIdentifier && record.RecordType == 0x01c0);
            Assert.Contains(legacy.FutureMetadataRecords, record => record.Kind == LegacyXlsWorkbookMetadataKind.ExtendedEncryption && record.RecordType == 0x01c1);
            Assert.Contains(legacy.FutureMetadataRecords, record => record.Kind == LegacyXlsWorkbookMetadataKind.PageLayoutView && record.RecordType == 0x088b);
            Assert.Contains(legacy.FutureMetadataRecords, record => record.Kind == LegacyXlsWorkbookMetadataKind.Compatibility12 && record.RecordType == 0x088c);
            Assert.Contains(legacy.FutureMetadataRecords, record => record.Kind == LegacyXlsWorkbookMetadataKind.TypeLibraryGuid && record.HasMatchingFutureRecordHeader);
            Assert.Contains(legacy.FutureMetadataRecords, record => record.Kind == LegacyXlsWorkbookMetadataKind.HeaderFooter && record.HasMatchingFutureRecordHeader);
            Assert.Contains(legacy.UnsupportedFeatures, feature => feature.Kind == LegacyXlsUnsupportedFeatureKind.ExternalReference && feature.DetailCode == "ExternalReference:DConRef");
            Assert.DoesNotContain(legacy.UnsupportedFeatures, feature => feature.Kind == LegacyXlsUnsupportedFeatureKind.ExternalReference && feature.DetailCode == "ExternalReference:DbQueryExt");
            Assert.DoesNotContain(legacy.UnsupportedFeatures, feature => feature.Kind == LegacyXlsUnsupportedFeatureKind.PivotTable && feature.DetailCode == "PivotTable:Sxvs");
            LegacyXlsPivotTableRecord cacheSourceRecord = Assert.Single(legacy.PivotTableRecords, record => record.RecordName == "Sxvs");
            Assert.Equal(LegacyXlsPivotTableRecordKind.CacheSource, cacheSourceRecord.Kind);
            Assert.True(cacheSourceRecord.HasSupportedPivotTableMetadata);
            Assert.Contains(legacy.UnsupportedFeatures, feature => feature.Kind == LegacyXlsUnsupportedFeatureKind.FeatureExtension && feature.DetailCode == "FeatureExtension:Feat");
            Assert.Contains(legacy.UnsupportedFeatures, feature => feature.Kind == LegacyXlsUnsupportedFeatureKind.ConditionalFormatting && feature.DetailCode == "ConditionalFormatting:Dxf");
            Assert.DoesNotContain(legacy.UnsupportedFeatures, feature => feature.Kind == LegacyXlsUnsupportedFeatureKind.TableStyle && feature.DetailCode == "TableStyle:TableStyles");
            LegacyXlsThemeRecord themeRecord = Assert.Single(legacy.ThemeRecords);
            Assert.Equal(0x0896, themeRecord.RecordType);
            Assert.Equal(124226U, themeRecord.ThemeVersion);
            Assert.Equal("Default", themeRecord.ThemeVersionName);
            Assert.False(themeRecord.HasThemeBytes);
            Assert.True(themeRecord.IsDefaultThemeMarker);
            Assert.DoesNotContain(legacy.UnsupportedFeatures, feature => feature.Kind == LegacyXlsUnsupportedFeatureKind.Theme && feature.DetailCode == "Theme:Theme" && feature.RecordType == 0x0896);
            Assert.Contains(legacy.UnsupportedFeatures, feature => feature.Kind == LegacyXlsUnsupportedFeatureKind.DrawingObject && feature.DetailCode == "Drawing:ShapePropsStream");
            Assert.DoesNotContain(legacy.UnsupportedFeatures, feature => feature.Kind == LegacyXlsUnsupportedFeatureKind.Chart && feature.DetailCode == "Chart:CrtLayout12");
            Assert.DoesNotContain(legacy.UnsupportedFeatures, feature => feature.Kind == LegacyXlsUnsupportedFeatureKind.PhoneticGuide);
            LegacyXlsChartRecord layoutRecord = Assert.Single(legacy.ChartRecords, record => record.RecordName == "CrtLayout12");
            Assert.Equal(LegacyXlsChartRecordKind.Layout, layoutRecord.Kind);
            Assert.True(layoutRecord.HasSupportedChartMetadata);
            Assert.Contains(legacy.Worksheets, sheet => sheet.PhoneticSettings != null);
            Assert.Contains(legacy.Worksheets.SelectMany(sheet => sheet.MetadataRecords), record => record.Kind == LegacyXlsWorksheetMetadataKind.PhoneticSettings && record.RecordType == (ushort)BiffRecordType.PhoneticInfo);
            Assert.Equal(legacy.UnsupportedFeatures.Count, legacy.PreservedFeatureRecords.Count);
            LegacyXlsExternalQueryConnection queryConnection = Assert.Single(legacy.ExternalQueryConnections);
            Assert.Equal((ushort)BiffRecordType.DbQueryExt, queryConnection.RecordType);
            Assert.Equal(LegacyXlsExternalQueryConnectionSourceType.OleDb, queryConnection.SourceTypeKind);
            Assert.Equal("OleDb", queryConnection.SourceTypeName);
            Assert.True(queryConnection.MaintainConnection);
            Assert.True(queryConnection.NewQuery);
            Assert.True(queryConnection.SourceIsXml);
            Assert.True(queryConnection.HasTextWizardQuery);
            Assert.True(queryConnection.HasTableNames);
            Assert.Equal(1, queryConnection.ParameterFlagCount);
            Assert.Equal(2, queryConnection.ParameterFlagByteCount);
            Assert.True(queryConnection.HasCompleteParameterFlags);
            Assert.Equal(2, queryConnection.FutureByteCount);
            Assert.Equal(15, queryConnection.RefreshIntervalMinutes);
            Assert.Equal(1, queryConnection.OleDbConnectionCount);

            LegacyXlsImportReport report = legacy.CreateImportReport();
            Assert.Equal(1, report.UnsupportedFeaturesByDetail["ConditionalFormatting|XLS-BIFF-FEATURE-CONDITIONAL-FORMATTING-UNSUPPORTED|ConditionalFormatting:Dxf"]);
            Assert.DoesNotContain("TableStyle|XLS-BIFF-FEATURE-TABLE-STYLE-UNSUPPORTED|TableStyle:TableStyles", report.UnsupportedFeaturesByDetail.Keys);
            Assert.Equal(10, report.WorkbookFutureMetadataRecordCount);
            Assert.Equal(1, report.WorkbookFutureMetadataRecordsByKind["PageLayoutView"]);
            Assert.Equal(1, report.WorkbookFutureMetadataRecordsByKind["Compatibility12"]);
            Assert.Equal(1, report.WorkbookFutureMetadataRecordsByKind["ExtendedEncryption"]);
            Assert.Equal(1, report.WorkbookFutureMetadataRecordsByRecordName["PLV"]);
            Assert.Equal(1, report.WorkbookFutureMetadataRecordsByRecordName["Compat12"]);
            Assert.Equal(5, report.WorkbookFutureMetadataRecordsByHeaderState["MatchingFutureHeader"]);
            Assert.Equal(5, report.WorkbookFutureMetadataRecordsByHeaderState["RawPayload"]);
            Assert.Equal(1, report.ExternalQueryConnectionCount);
            Assert.Equal(1, report.ExternalQueryConnectionsBySourceType["OleDb"]);
            Assert.Equal(1, report.ExternalQueryConnectionsByConnectionFlag["MaintainConnection"]);
            Assert.Equal(1, report.ExternalQueryConnectionsByConnectionFlag["NewQuery"]);
            Assert.Equal(1, report.ExternalQueryConnectionsByConnectionFlag["SourceIsXml"]);
            Assert.Equal(1, report.ExternalQueryConnectionsByQueryOption["TextWizardQuery"]);
            Assert.Equal(1, report.ExternalQueryConnectionsByQueryOption["TableNames"]);
            Assert.Equal(1, report.ExternalQueryConnectionsByParameterFlagCount["Parameters:1"]);
            Assert.Equal(1, report.ExternalQueryConnectionsByParameterFlagByteCount["Bytes:2"]);
            Assert.Equal(1, report.ExternalQueryConnectionsByParameterFlagState["Complete"]);
            Assert.Equal(1, report.ExternalQueryConnectionsByFutureByteCount["Bytes:2"]);
            Assert.Equal(1, report.ExternalQueryConnectionsByRefreshInterval["Minutes:15"]);
            Assert.Equal(1, report.ExternalQueryConnectionsByOleDbConnectionCount["OleDbConnections:1"]);
            Assert.Equal(1, report.ExternalQueryConnectionsByHtmlFormat["HtmlFormat:0x0003"]);
            Assert.Equal(1, report.ExternalQueryConnectionsByVersionTriplet["Edit:3;Refreshed:2;RefreshableMin:1"]);
            Assert.Equal(1, report.ExternalQueryConnectionsBySourceSpecificFlags["Flags:0x010A"]);
            Assert.DoesNotContain("ExternalReference|XLS-BIFF-FEATURE-EXTERNAL-REFERENCE-UNSUPPORTED|ExternalReference:DbQueryExt", report.UnsupportedFeaturesByDetail.Keys);
            Assert.Equal(1, report.ThemeRecordsByVersion["Default"]);
            Assert.Equal(1, report.ThemeRecordsByRawVersion["Version:124226"]);
            Assert.Equal(1, report.ThemeRecordsByContentState["NoEmbeddedThemeBytes"]);
            string markdown = report.ToMarkdown();
            Assert.Contains("External query connections: 1", markdown, StringComparison.Ordinal);
            Assert.Contains("External Query Connections By Source Type", markdown, StringComparison.Ordinal);
            Assert.Contains("Preserved Feature Records By Kind", markdown, StringComparison.Ordinal);
        }

        [Fact]
        public void LegacyXls_LoadLegacyXls_ProjectsExternalQueryConnectionMetadata() {
            byte[] workbookStream = LegacyXlsTestWorkbookBuilder.CreateKnownPreserveOnlyExtensionWorkbookStream();
            byte[] compound = LegacyXlsCompoundTestBuilder.CreateWorkbookCompoundFile(workbookStream);

            using LegacyXlsLoadResult result = ExcelDocument.LoadLegacyXlsWithReport(new MemoryStream(compound), new LegacyXlsImportOptions {
                ReportUnsupportedContent = true
            });

            Assert.DoesNotContain(result.UnsupportedFeatures, feature =>
                feature.Kind == LegacyXlsUnsupportedFeatureKind.ExternalReference
                && feature.DetailCode == "ExternalReference:DbQueryExt");

            using var output = new MemoryStream();
            result.Document.Save(output, new ExcelSaveOptions {
                LossPolicy = ExcelConversionLossPolicy.Allow
            });
            using SpreadsheetDocument spreadsheet = SpreadsheetDocument.Open(new MemoryStream(output.ToArray()), false);
            OpenXmlPart connectionPart = Assert.Single(
                spreadsheet.WorkbookPart!.Parts.Select(part => part.OpenXmlPart),
                part => part.ContentType.IndexOf("connections", StringComparison.OrdinalIgnoreCase) >= 0);

            using Stream connectionStream = connectionPart.GetStream(FileMode.Open, FileAccess.Read);
            using var reader = new StreamReader(connectionStream);
            string connectionXml = reader.ReadToEnd();
            Assert.Contains("<connections", connectionXml, StringComparison.OrdinalIgnoreCase);
            Assert.Contains("count=\"1\"", connectionXml, StringComparison.OrdinalIgnoreCase);
            Assert.Contains("name=\"LegacyXlsQuery1\"", connectionXml, StringComparison.Ordinal);
            Assert.Contains("type=\"5\"", connectionXml, StringComparison.Ordinal);
            Assert.Contains("refreshedVersion=\"2\"", connectionXml, StringComparison.Ordinal);
            Assert.Contains("minRefreshableVersion=\"1\"", connectionXml, StringComparison.Ordinal);
            Assert.Contains("interval=\"15\"", connectionXml, StringComparison.Ordinal);
            Assert.Contains("Legacy XLS DBQueryExt metadata", connectionXml, StringComparison.Ordinal);
            Assert.Contains("Source=OleDb", connectionXml, StringComparison.Ordinal);
            Assert.Contains("Flags=MaintainConnection,NewQuery,SourceIsXml", connectionXml, StringComparison.Ordinal);
            Assert.Contains("QueryOptions=TextWizardQuery,TableNames", connectionXml, StringComparison.Ordinal);
        }

        [Fact]
        public void LegacyXls_Load_ClassifiesObservedCorpusPreserveOnlyRecords() {
            byte[] workbookStream = LegacyXlsTestWorkbookBuilder.CreateObservedCorpusPreserveOnlyRecordWorkbookStream();
            byte[] compound = LegacyXlsCompoundTestBuilder.CreateWorkbookCompoundFile(workbookStream);

            LegacyXlsWorkbook legacy = LegacyXlsWorkbook.Load(compound, new LegacyXlsImportOptions {
                ReportUnsupportedContent = true
            });

            Assert.DoesNotContain(legacy.Diagnostics, d => d.Severity == LegacyXlsDiagnosticSeverity.Error);
            Assert.DoesNotContain(legacy.UnsupportedFeatures, feature => feature.Kind == LegacyXlsUnsupportedFeatureKind.UnsupportedRecord);
            Assert.Contains(legacy.UnsupportedFeatures, feature => feature.Kind == LegacyXlsUnsupportedFeatureKind.Formula && feature.DetailCode == "Formula:Array");
            Assert.DoesNotContain(legacy.UnsupportedFeatures, feature => feature.Kind == LegacyXlsUnsupportedFeatureKind.WorksheetProtection);
            Assert.DoesNotContain(legacy.UnsupportedFeatures, feature => feature.Kind == LegacyXlsUnsupportedFeatureKind.DrawingObject && feature.DetailCode == "Drawing:HFPicture");
            Assert.Equal(legacy.UnsupportedFeatures.Count, legacy.PreservedFeatureRecords.Count);
            LegacyXlsWorksheet sheet = Assert.Single(legacy.Worksheets);
            Assert.NotNull(sheet.Protection);
            Assert.True(sheet.Protection!.ProtectObjects);
            Assert.True(sheet.Protection.ProtectScenarios);

            LegacyXlsImportReport report = legacy.CreateImportReport();
            Assert.Equal(1, report.UnsupportedFeaturesByKind[LegacyXlsUnsupportedFeatureKind.Formula]);
            Assert.DoesNotContain(LegacyXlsUnsupportedFeatureKind.WorksheetProtection, report.UnsupportedFeaturesByKind.Keys);
            Assert.DoesNotContain(LegacyXlsUnsupportedFeatureKind.DrawingObject, report.UnsupportedFeaturesByKind.Keys);
            Assert.Equal(1, report.PreservedFeatureRecordsByKind[LegacyXlsUnsupportedFeatureKind.Formula]);
            Assert.DoesNotContain(LegacyXlsUnsupportedFeatureKind.WorksheetProtection, report.PreservedFeatureRecordsByKind.Keys);
            Assert.Equal(1, report.DrawingRecordsByKind[LegacyXlsDrawingRecordKind.HeaderFooterPicture]);
            Assert.Equal(1, report.DrawingRecordsByName["HFPicture"]);
            Assert.Equal(1, report.DrawingHeaderFooterPictureHeaderStates["Complete"]);
            Assert.Equal(1, report.DrawingHeaderFooterPictureDrawingKinds["Drawing"]);
            Assert.Equal(1, report.DrawingHeaderFooterPictureContinuationStates["First"]);
            Assert.Equal(1, report.WorksheetProtectionObjectStates["Protected"]);
            Assert.Equal(1, report.WorksheetProtectionScenarioStates["Protected"]);
            Assert.Equal(1, report.UnsupportedFeaturesByDetail["Formula|XLS-BIFF-FEATURE-FORMULA-UNSUPPORTED|Formula:Array"]);
            Assert.DoesNotContain("DrawingObject|XLS-BIFF-FEATURE-DRAWING-UNSUPPORTED|Drawing:HFPicture", report.UnsupportedFeaturesByDetail.Keys);
            LegacyXlsDrawingRecord headerFooterPicture = Assert.Single(legacy.DrawingRecords, record => record.Kind == LegacyXlsDrawingRecordKind.HeaderFooterPicture);
            Assert.True(headerFooterPicture.HasHeaderFooterPicture);
            Assert.True(headerFooterPicture.HasSupportedHeaderFooterPictureMetadata);
            Assert.True(headerFooterPicture.OfficeArtPayloadFullyTraversed);
            Assert.Equal(LegacyXlsDrawingEscherRecordType.OfficeArtDgContainer, headerFooterPicture.EscherRecordTypeKind);
            Assert.Equal("Drawing", headerFooterPicture.HeaderFooterPicture?.DrawingKindName);
            Assert.Equal("First", headerFooterPicture.HeaderFooterPicture?.ContinuationState);
            Assert.Equal(new[] {
                "OfficeArtDgContainer",
                "OfficeArtFDG",
                "OfficeArtSpContainer",
                "OfficeArtFSP",
                "OfficeArtFOPT",
                "OfficeArtFClientAnchor",
                "OfficeArtChildAnchor"
            }, headerFooterPicture.OfficeArtRecords.Select(record => record.RecordTypeName).ToArray());
        }

        [Fact]
        public void LegacyXls_LoadLegacyXls_ProjectsHeaderFooterPictureBlipToOpenXmlHeaderImage() {
            byte[] workbookStream = LegacyXlsTestWorkbookBuilder.CreateHeaderFooterPictureImageWorkbookStream();
            byte[] compound = LegacyXlsCompoundTestBuilder.CreateWorkbookCompoundFile(workbookStream);

            using LegacyXlsLoadResult result = ExcelDocument.LoadLegacyXlsWithReport(new MemoryStream(compound), new LegacyXlsImportOptions {
                ReportUnsupportedContent = true
            });

            LegacyXlsDrawingBlipStoreEntry blip = Assert.Single(result.Workbook.DrawingRecords.SelectMany(record => record.BlipStoreEntries));
            Assert.True(blip.HasImportableImagePayload);
            Assert.Equal("image/png", blip.EmbeddedBlipContentType);
            Assert.DoesNotContain(result.UnsupportedFeatures, feature =>
                feature.Kind == LegacyXlsUnsupportedFeatureKind.DrawingObject
                && feature.DetailCode == "Drawing:HFPicture");

            ExcelSheet.HeaderFooterSnapshot headerFooter = result.Document.Sheets.Single().GetHeaderFooter();
            Assert.True(headerFooter.HeaderHasPicturePlaceholder);
            Assert.NotNull(headerFooter.HeaderCenterImage);
            Assert.Equal(HeaderFooterPosition.Center, headerFooter.HeaderCenterImage!.Position);
            Assert.Equal("image/png", headerFooter.HeaderCenterImage.ContentType);
            Assert.Equal(blip.EmbeddedBlipPayloadBytes, headerFooter.HeaderCenterImage.Bytes);

            using var output = new MemoryStream();
            result.Document.Save(output);
            using SpreadsheetDocument spreadsheet = SpreadsheetDocument.Open(new MemoryStream(output.ToArray()), false);
            WorksheetPart worksheetPart = spreadsheet.WorkbookPart!.WorksheetParts.Single();
            VmlDrawingPart vmlPart = Assert.Single(worksheetPart.VmlDrawingParts);
            ImagePart imagePart = Assert.Single(vmlPart.ImageParts);
            Assert.Equal("image/png", imagePart.ContentType);
            using var imageStream = new MemoryStream();
            imagePart.GetStream().CopyTo(imageStream);
            Assert.Equal(blip.EmbeddedBlipPayloadBytes, imageStream.ToArray());
        }

        [Fact]
        public void LegacyXls_Load_KeepsEmptyHeaderFooterPictureUnsupported() {
            byte[] workbookStream = LegacyXlsTestWorkbookBuilder.CreateEmptyHeaderFooterPictureWorkbookStream();
            byte[] compound = LegacyXlsCompoundTestBuilder.CreateWorkbookCompoundFile(workbookStream);

            LegacyXlsWorkbook legacy = LegacyXlsWorkbook.Load(compound, new LegacyXlsImportOptions {
                ReportUnsupportedContent = true
            });

            Assert.Contains(legacy.UnsupportedFeatures, feature => feature.Kind == LegacyXlsUnsupportedFeatureKind.DrawingObject && feature.DetailCode == "Drawing:HFPicture");
            LegacyXlsDrawingRecord headerFooterPicture = Assert.Single(legacy.DrawingRecords, record => record.Kind == LegacyXlsDrawingRecordKind.HeaderFooterPicture);
            Assert.True(headerFooterPicture.HasHeaderFooterPicture);
            Assert.False(headerFooterPicture.HasSupportedHeaderFooterPictureMetadata);
            Assert.Empty(headerFooterPicture.OfficeArtRecords);

            LegacyXlsImportReport report = legacy.CreateImportReport();
            Assert.Equal(1, report.UnsupportedFeaturesByKind[LegacyXlsUnsupportedFeatureKind.DrawingObject]);
            Assert.Equal(1, report.DrawingHeaderFooterPictureHeaderStates["MismatchedFutureRecordHeader"]);
            Assert.Equal(1, report.DrawingHeaderFooterPictureDrawingByteCounts["DrawingBytes:0"]);
        }

        [Fact]
        public void LegacyXls_Load_PreservesUnsupportedFeatureMetadataWhenDiagnosticsAreDisabled() {
            byte[] workbookStream = LegacyXlsTestWorkbookBuilder.CreateUnsupportedFeatureWorkbookStream();
            byte[] compound = LegacyXlsCompoundTestBuilder.CreateWorkbookCompoundFile(workbookStream);

            LegacyXlsWorkbook legacy = LegacyXlsWorkbook.Load(compound, new LegacyXlsImportOptions {
                ReportUnsupportedContent = false
            });

            Assert.DoesNotContain(legacy.Diagnostics, d => d.Code.StartsWith("XLS-BIFF-FEATURE-", StringComparison.Ordinal));
            Assert.Contains(legacy.UnsupportedFeatures, feature => feature.Kind == LegacyXlsUnsupportedFeatureKind.Hyperlink && feature.Code == "XLS-BIFF-FEATURE-HYPERLINK-UNSUPPORTED");
            Assert.Contains(legacy.UnsupportedFeatures, feature => feature.Kind == LegacyXlsUnsupportedFeatureKind.Comment && feature.Code == "XLS-BIFF-FEATURE-COMMENT-UNSUPPORTED");
            Assert.Contains(legacy.UnsupportedFeatures, feature => feature.Kind == LegacyXlsUnsupportedFeatureKind.DrawingObject && feature.Code == "XLS-BIFF-FEATURE-DRAWING-UNSUPPORTED");
            Assert.Contains(legacy.UnsupportedFeatures, feature => feature.Kind == LegacyXlsUnsupportedFeatureKind.PivotTable && feature.Code == "XLS-BIFF-FEATURE-PIVOT-TABLE-UNSUPPORTED");
            Assert.Contains(legacy.UnsupportedFeatures, feature => feature.Kind == LegacyXlsUnsupportedFeatureKind.AutoFilterCriteria && feature.Code == "XLS-BIFF-FEATURE-AUTOFILTER-CRITERIA-UNSUPPORTED");
            Assert.Contains(legacy.PreservedFeatureRecords, record => record.Kind == LegacyXlsUnsupportedFeatureKind.Hyperlink && record.Code == "XLS-BIFF-FEATURE-HYPERLINK-UNSUPPORTED");
            Assert.Contains(legacy.PreservedFeatureRecords, record => record.Kind == LegacyXlsUnsupportedFeatureKind.Comment && record.Code == "XLS-BIFF-FEATURE-COMMENT-UNSUPPORTED");
            Assert.Contains(legacy.PreservedFeatureRecords, record => record.Kind == LegacyXlsUnsupportedFeatureKind.DrawingObject && record.Code == "XLS-BIFF-FEATURE-DRAWING-UNSUPPORTED");
            Assert.Contains(legacy.PreservedFeatureRecords, record => record.Kind == LegacyXlsUnsupportedFeatureKind.PivotTable && record.Code == "XLS-BIFF-FEATURE-PIVOT-TABLE-UNSUPPORTED");
            Assert.Contains(legacy.PreservedFeatureRecords, record => record.Kind == LegacyXlsUnsupportedFeatureKind.AutoFilterCriteria && record.Code == "XLS-BIFF-FEATURE-AUTOFILTER-CRITERIA-UNSUPPORTED");
        }

        [Fact]
        public void LegacyXls_Load_DecodesChartSheetPrintSizeMetadata() {
            byte[] workbookStream = LegacyXlsTestWorkbookBuilder.CreateUnsupportedChartSheetWorkbookStream();
            byte[] compound = LegacyXlsCompoundTestBuilder.CreateWorkbookCompoundFile(workbookStream);

            LegacyXlsWorkbook legacy = LegacyXlsWorkbook.Load(compound, new LegacyXlsImportOptions {
                ReportUnsupportedContent = true
            });

            Assert.Empty(legacy.UnsupportedSheets);
            LegacyXlsChartSheet chartSheet = Assert.Single(legacy.ChartSheets);
            Assert.Equal("Chart1", chartSheet.Name);
            Assert.Equal((ushort)2, chartSheet.ChartPrintSize);
            Assert.Equal(LegacyXlsChartPrintSize.FitPage, chartSheet.ChartPrintSizeKind);
            Assert.Equal("FitPage", chartSheet.ChartPrintSizeName);
            Assert.Equal(1, chartSheet.ChartTextObjectCount);
            Assert.Equal(0, chartSheet.ChartRecordCount);
            Assert.Empty(chartSheet.ChartRecordsByKind);
            Assert.Empty(chartSheet.ChartRecordsByChartType);
            Assert.Equal(2, chartSheet.MetadataRecords.Count);
            LegacyXlsChartSheetMetadataRecord printSizeMetadata = Assert.Single(chartSheet.MetadataRecords, metadata => metadata.Kind == LegacyXlsChartSheetMetadataKind.ChartPrintSize);
            Assert.Equal((ushort)BiffRecordType.PrintSize, printSizeMetadata.RecordType);
            LegacyXlsChartSheetMetadataRecord textObjectMetadata = Assert.Single(chartSheet.MetadataRecords, metadata => metadata.Kind == LegacyXlsChartSheetMetadataKind.ChartTextObject);
            Assert.Equal((ushort)BiffRecordType.Txo, textObjectMetadata.RecordType);
            Assert.DoesNotContain(legacy.UnsupportedFeatures, feature => feature.RecordType == (ushort)BiffRecordType.PrintSize);
            Assert.DoesNotContain(legacy.UnsupportedFeatures, feature => feature.RecordType == (ushort)BiffRecordType.Txo);
            Assert.DoesNotContain(legacy.UnsupportedFeatures, feature => feature.Kind == LegacyXlsUnsupportedFeatureKind.ChartSheet);

            LegacyXlsImportReport report = new LegacyXlsImportReport(legacy);
            Assert.Equal(1, report.ChartSheetCount);
            Assert.Equal(2, report.ChartSheetMetadataRecordCount);
            Assert.Equal(1, report.ChartSheetMetadataRecordsByKind["ChartPrintSize"]);
            Assert.Equal(1, report.ChartSheetMetadataRecordsByKind["ChartTextObject"]);
            Assert.Equal(1, report.ChartSheetPrintSizes["PrintSize:2"]);
            Assert.Equal(1, report.ChartSheetPrintSizeKinds["FitPage"]);
            Assert.Equal(1, report.ChartSheetTextObjectCounts["TextObjects:1"]);
            Assert.Equal(1, report.ChartSheetStates["PrintSize:Present|TextObjects:Present|ChartRecords:Missing|ChartTypes:Missing"]);
            Assert.Equal(0, report.UnsupportedSheetMetadataRecordCount);
            Assert.Empty(report.UnsupportedChartSheetPrintSizes);
            Assert.Empty(report.UnsupportedChartSheetChartRecordCounts);
            Assert.Empty(report.UnsupportedChartSheetChartRecordKinds);
            Assert.Empty(report.UnsupportedChartSheetChartTypes);
            string markdown = report.ToMarkdown();
            Assert.Contains("Chart sheet metadata records: 2", markdown);
            Assert.Contains("Chart Sheet Print Sizes", markdown);
            Assert.Contains("Chart Sheet Print Size Kinds", markdown);
            Assert.Contains("Chart Sheet Text Object Counts", markdown);
            Assert.Contains("Chart Sheet States", markdown);
        }

        [Fact]
        public void LegacyXls_Load_DecodesWorksheetTextObjectContinuePayload() {
            byte[] workbookStream = LegacyXlsTestWorkbookBuilder.CreateWorksheetTextObjectWorkbookStream("Legacy box");
            byte[] compound = LegacyXlsCompoundTestBuilder.CreateWorkbookCompoundFile(workbookStream);

            LegacyXlsWorkbook legacy = LegacyXlsWorkbook.Load(compound, new LegacyXlsImportOptions {
                ReportUnsupportedContent = true
            });

            LegacyXlsDrawingRecord textObjectRecord = Assert.Single(
                legacy.DrawingRecords,
                record => record.Kind == LegacyXlsDrawingRecordKind.TextObject);
            Assert.NotNull(textObjectRecord.TextObject);
            Assert.Equal("Legacy box", textObjectRecord.TextObject!.Text);
            Assert.True(textObjectRecord.TextObject.HasDecodedText);
            LegacyXlsCommentFormattingRun formattingRun = Assert.Single(textObjectRecord.TextObject.FormattingRuns);
            Assert.Equal((ushort)0, formattingRun.StartCharacter);
            Assert.Equal((ushort)0, formattingRun.FontIndex);
            Assert.DoesNotContain(legacy.UnsupportedFeatures, feature => feature.RecordType == (ushort)BiffRecordType.Txo);
            Assert.DoesNotContain(legacy.Diagnostics, diagnostic =>
                diagnostic.Code == "XLS-BIFF-FEATURE-DRAWING-UNSUPPORTED"
                && diagnostic.RecordType == (ushort)BiffRecordType.Txo);

            LegacyXlsImportReport report = legacy.CreateImportReport();
            Assert.Equal(1, report.DrawingTextObjectFlags["TextInContinueRecords:Present"]);
            Assert.Equal(1, report.DrawingTextObjectFlags["FormattingRunsInContinueRecords:Present"]);
            Assert.Equal(1, report.DrawingTextObjectFlags["DecodedText:Present"]);
            Assert.Equal(1, report.DrawingTextObjectFlags["DecodedFormattingRuns:1"]);
        }

        [Fact]
        public void LegacyXls_Load_SupportsOfficeArtClientTextboxDrawingMetadata() {
            byte[] workbookStream = LegacyXlsTestWorkbookBuilder.CreateWorksheetClientTextboxDrawingWorkbookStream();
            byte[] compound = LegacyXlsCompoundTestBuilder.CreateWorkbookCompoundFile(workbookStream);

            LegacyXlsWorkbook legacy = LegacyXlsWorkbook.Load(compound, new LegacyXlsImportOptions {
                ReportUnsupportedContent = true
            });

            Assert.DoesNotContain(legacy.UnsupportedFeatures, feature => feature.Kind == LegacyXlsUnsupportedFeatureKind.DrawingObject);
            LegacyXlsDrawingRecord drawing = Assert.Single(legacy.DrawingRecords);
            Assert.Equal(LegacyXlsDrawingRecordKind.Drawing, drawing.Kind);
            Assert.Equal(LegacyXlsDrawingEscherRecordType.OfficeArtFClientTextbox, drawing.EscherRecordTypeKind);
            Assert.True(drawing.HasSupportedOfficeArtClientTextboxMetadata);
            Assert.True(drawing.HasSupportedDrawingMetadata);
            Assert.True(drawing.OfficeArtPayloadFullyTraversed);
            LegacyXlsDrawingOfficeArtRecord officeArtRecord = Assert.Single(drawing.OfficeArtRecords);
            Assert.Equal(LegacyXlsDrawingEscherRecordType.OfficeArtFClientTextbox, officeArtRecord.RecordTypeKind);
            Assert.False(officeArtRecord.IsContainer);

            LegacyXlsImportReport report = legacy.CreateImportReport();
            Assert.DoesNotContain(LegacyXlsUnsupportedFeatureKind.DrawingObject, report.UnsupportedFeaturesByKind.Keys);
            Assert.Equal(1, report.DrawingRecordsByName["MsoDrawing"]);
            Assert.Equal(1, report.DrawingRecordsByEscherRecordTypeName["OfficeArtFClientTextbox"]);
            Assert.Equal(1, report.DrawingOfficeArtRecordsByTypeName["OfficeArtFClientTextbox"]);
        }

        [Fact]
        public void LegacyXls_Load_SupportsContinuedMsoDrawingOfficeArtMetadata() {
            byte[] workbookStream = LegacyXlsTestWorkbookBuilder.CreateWorksheetContinuedDrawingWorkbookStream();
            byte[] compound = LegacyXlsCompoundTestBuilder.CreateWorkbookCompoundFile(workbookStream);

            LegacyXlsWorkbook legacy = LegacyXlsWorkbook.Load(compound, new LegacyXlsImportOptions {
                ReportUnsupportedContent = true
            });

            Assert.DoesNotContain(legacy.UnsupportedFeatures, feature => feature.Kind == LegacyXlsUnsupportedFeatureKind.DrawingObject);
            LegacyXlsDrawingRecord drawing = Assert.Single(legacy.DrawingRecords, record => record.RecordName == "MsoDrawing");
            Assert.True(drawing.HasSupportedOfficeArtMetadata);
            Assert.True(drawing.HasSupportedDrawingMetadata);
            Assert.True(drawing.OfficeArtPayloadFullyTraversed);
            Assert.Equal(LegacyXlsDrawingEscherRecordType.OfficeArtDgContainer, drawing.EscherRecordTypeKind);
            Assert.Equal(new[] {
                "OfficeArtDgContainer",
                "OfficeArtFDG",
                "OfficeArtSpContainer",
                "OfficeArtFSP",
                "OfficeArtFOPT",
                "OfficeArtFClientAnchor",
                "OfficeArtChildAnchor"
            }, drawing.OfficeArtRecords.Select(record => record.RecordTypeName).ToArray());

            LegacyXlsImportReport report = legacy.CreateImportReport();
            Assert.DoesNotContain(LegacyXlsUnsupportedFeatureKind.DrawingObject, report.UnsupportedFeaturesByKind.Keys);
            Assert.Equal(1, report.DrawingRecordsByName["MsoDrawing"]);
            Assert.Equal(1, report.DrawingRecordsByEscherRecordTypeName["OfficeArtDgContainer"]);
            Assert.Equal(1, report.DrawingShapeEntriesByType["PictureFrame"]);
            Assert.Equal(1, report.DrawingAnchorEntriesByRange["R2C1:R4C3"]);
        }

        [Fact]
        public void LegacyXls_LoadLegacyXls_ProjectsWorksheetPictureBlipToOpenXmlImage() {
            byte[] workbookStream = LegacyXlsTestWorkbookBuilder.CreatePhase5PreserveOnlyFeatureDetailsWorkbookStream();
            byte[] compound = LegacyXlsCompoundTestBuilder.CreateWorkbookCompoundFile(workbookStream);

            using LegacyXlsLoadResult result = ExcelDocument.LoadLegacyXlsWithReport(new MemoryStream(compound), new LegacyXlsImportOptions {
                ReportUnsupportedContent = true
            });

            LegacyXlsDrawingBlipStoreEntry blip = Assert.Single(result.Workbook.DrawingRecords.SelectMany(record => record.BlipStoreEntries));
            Assert.True(blip.HasImportableImagePayload);
            Assert.Equal("image/png", blip.EmbeddedBlipContentType);
            Assert.Equal(new byte[] { 0x89, 0x50, 0x4e, 0x47 }, blip.EmbeddedBlipPayloadBytes.Take(4).ToArray());

            using var output = new MemoryStream();
            result.Document.Save(output, new ExcelSaveOptions {
                LossPolicy = ExcelConversionLossPolicy.Allow
            });
            using SpreadsheetDocument spreadsheet = SpreadsheetDocument.Open(new MemoryStream(output.ToArray()), false);
            WorksheetPart worksheetPart = Assert.Single(spreadsheet.WorkbookPart!.WorksheetParts, part => part.DrawingsPart != null);
            DrawingsPart drawingsPart = worksheetPart.DrawingsPart!;
            ImagePart imagePart = Assert.Single(drawingsPart.ImageParts);
            Assert.Equal("image/png", imagePart.ContentType);
            using var imageStream = new MemoryStream();
            imagePart.GetStream().CopyTo(imageStream);
            Assert.Equal(blip.EmbeddedBlipPayloadBytes, imageStream.ToArray());
            Assert.Single(drawingsPart.WorksheetDrawing!.Elements<DocumentFormat.OpenXml.Drawing.Spreadsheet.TwoCellAnchor>());
        }

        [Fact]
        public void LegacyXls_Load_SupportsSplitMsoDrawingContainerMetadata() {
            byte[] workbookStream = LegacyXlsTestWorkbookBuilder.CreateWorksheetSplitDrawingContainerWorkbookStream();
            byte[] compound = LegacyXlsCompoundTestBuilder.CreateWorkbookCompoundFile(workbookStream);

            LegacyXlsWorkbook legacy = LegacyXlsWorkbook.Load(compound, new LegacyXlsImportOptions {
                ReportUnsupportedContent = true
            });

            Assert.DoesNotContain(legacy.UnsupportedFeatures, feature => feature.Kind == LegacyXlsUnsupportedFeatureKind.DrawingObject);
            LegacyXlsDrawingRecord drawing = Assert.Single(legacy.DrawingRecords, record => record.RecordName == "MsoDrawing");
            Assert.True(drawing.HasSupportedPartialOfficeArtContainerMetadata);
            Assert.True(drawing.HasSupportedDrawingMetadata);
            Assert.False(drawing.OfficeArtPayloadFullyTraversed);
            Assert.Equal(LegacyXlsDrawingEscherRecordType.OfficeArtDgContainer, drawing.EscherRecordTypeKind);
            Assert.Equal(new[] {
                "OfficeArtDgContainer",
                "OfficeArtFDG"
            }, drawing.OfficeArtRecords.Select(record => record.RecordTypeName).ToArray());
            LegacyXlsDrawingGroupInfo drawingGroupInfo = Assert.Single(drawing.DrawingGroupInfos);
            Assert.Equal(1, drawingGroupInfo.DrawingId);

            LegacyXlsImportReport report = legacy.CreateImportReport();
            Assert.DoesNotContain(LegacyXlsUnsupportedFeatureKind.DrawingObject, report.UnsupportedFeaturesByKind.Keys);
            Assert.Equal(1, report.DrawingRecordsByName["MsoDrawing"]);
            Assert.Equal(1, report.DrawingRecordsByEscherRecordTypeName["OfficeArtDgContainer"]);
            Assert.Equal(1, report.DrawingOfficeArtRecordsByTypeName["OfficeArtFDG"]);
        }

        [Fact]
        public void LegacyXls_Load_SupportsSplitMsoDrawingShapeContainerMetadata() {
            byte[] workbookStream = LegacyXlsTestWorkbookBuilder.CreateWorksheetSplitShapeContainerWorkbookStream();
            byte[] compound = LegacyXlsCompoundTestBuilder.CreateWorkbookCompoundFile(workbookStream);

            LegacyXlsWorkbook legacy = LegacyXlsWorkbook.Load(compound, new LegacyXlsImportOptions {
                ReportUnsupportedContent = true
            });

            Assert.DoesNotContain(legacy.UnsupportedFeatures, feature => feature.Kind == LegacyXlsUnsupportedFeatureKind.DrawingObject);
            LegacyXlsDrawingRecord drawing = Assert.Single(legacy.DrawingRecords, record => record.RecordName == "MsoDrawing");
            Assert.True(drawing.HasSupportedPartialOfficeArtContainerMetadata);
            Assert.True(drawing.HasSupportedDrawingMetadata);
            Assert.False(drawing.OfficeArtPayloadFullyTraversed);
            Assert.Equal(LegacyXlsDrawingEscherRecordType.OfficeArtSpContainer, drawing.EscherRecordTypeKind);
            Assert.Equal(new[] {
                "OfficeArtSpContainer",
                "OfficeArtFSP",
                "OfficeArtFOPT",
                "OfficeArtFClientAnchor",
                "OfficeArtFClientData"
            }, drawing.OfficeArtRecords.Select(record => record.RecordTypeName).ToArray());
            LegacyXlsDrawingShape shape = Assert.Single(drawing.ShapeEntries);
            Assert.Equal("Rectangle", shape.ShapeTypeName);
            Assert.Single(drawing.AnchorEntries);

            LegacyXlsImportReport report = legacy.CreateImportReport();
            Assert.DoesNotContain(LegacyXlsUnsupportedFeatureKind.DrawingObject, report.UnsupportedFeaturesByKind.Keys);
            Assert.Equal(1, report.DrawingRecordsByName["MsoDrawing"]);
            Assert.Equal(1, report.DrawingRecordsByEscherRecordTypeName["OfficeArtSpContainer"]);
            Assert.Equal(1, report.DrawingShapeEntriesByType["Rectangle"]);
        }

        [Fact]
        public void LegacyXls_Load_SupportsDecodedWorksheetObjectMetadata() {
            byte[] workbookStream = LegacyXlsTestWorkbookBuilder.CreateWorksheetObjectWorkbookStream();
            byte[] compound = LegacyXlsCompoundTestBuilder.CreateWorkbookCompoundFile(workbookStream);

            LegacyXlsWorkbook legacy = LegacyXlsWorkbook.Load(compound, new LegacyXlsImportOptions {
                ReportUnsupportedContent = true
            });

            LegacyXlsDrawingRecord objectRecord = Assert.Single(legacy.DrawingRecords, record => record.Kind == LegacyXlsDrawingRecordKind.Object);
            Assert.True(objectRecord.HasSupportedObjectMetadata);
            Assert.Equal(LegacyXlsDrawingObjectType.Picture, objectRecord.ObjectTypeKind);
            Assert.Equal((ushort)1, objectRecord.ObjectId);
            Assert.Equal((ushort)0x4011, objectRecord.ObjectFlags);
            Assert.Equal(new[] { "FtCmo", "FtEnd" }, objectRecord.ObjectSubRecords.Select(subRecord => subRecord.SubRecordName).ToArray());
            Assert.DoesNotContain(legacy.UnsupportedFeatures, feature => feature.RecordType == (ushort)BiffRecordType.Obj);
            Assert.DoesNotContain(legacy.Diagnostics, diagnostic =>
                diagnostic.Code == "XLS-BIFF-FEATURE-DRAWING-UNSUPPORTED"
                && diagnostic.RecordType == (ushort)BiffRecordType.Obj);

            LegacyXlsImportReport report = legacy.CreateImportReport();
            Assert.Equal(1, report.DrawingRecordsByObjectTypeName["Picture"]);
            Assert.Equal(1, report.DrawingObjectSubRecordsByName["FtCmo"]);
            Assert.Equal(1, report.DrawingObjectSubRecordsByName["FtEnd"]);
            Assert.DoesNotContain("DrawingObject|XLS-BIFF-FEATURE-DRAWING-UNSUPPORTED|Drawing:Obj", report.UnsupportedFeaturesByDetail.Keys);
        }

        [Fact]
        public void LegacyXls_Load_SupportsContinuedListObjectMetadata() {
            byte[] workbookStream = LegacyXlsTestWorkbookBuilder.CreateWorksheetListObjectWorkbookStream();
            byte[] compound = LegacyXlsCompoundTestBuilder.CreateWorkbookCompoundFile(workbookStream);

            LegacyXlsWorkbook legacy = LegacyXlsWorkbook.Load(compound, new LegacyXlsImportOptions {
                ReportUnsupportedContent = true
            });

            LegacyXlsDrawingRecord objectRecord = Assert.Single(legacy.DrawingRecords, record => record.Kind == LegacyXlsDrawingRecordKind.Object);
            Assert.True(objectRecord.HasSupportedObjectMetadata);
            Assert.Equal(LegacyXlsDrawingObjectType.DropdownList, objectRecord.ObjectTypeKind);
            Assert.Equal(new[] { "FtCmo", "FtSbs", "FtLbsData" }, objectRecord.ObjectSubRecords.Select(subRecord => subRecord.SubRecordName).ToArray());

            LegacyXlsDrawingObjectSubRecord listData = Assert.Single(objectRecord.ObjectSubRecords, subRecord => subRecord.SubRecordName == "FtLbsData");
            Assert.False(listData.IsComplete);
            Assert.True(listData.RequiresContinuation);
            Assert.True(listData.HasSupportedPayload);
            Assert.Equal("RequiresContinuation", listData.CompletionState);
            Assert.DoesNotContain(legacy.UnsupportedFeatures, feature => feature.RecordType == (ushort)BiffRecordType.Obj);
            Assert.DoesNotContain(legacy.Diagnostics, diagnostic =>
                diagnostic.Code == "XLS-BIFF-FEATURE-DRAWING-UNSUPPORTED"
                && diagnostic.RecordType == (ushort)BiffRecordType.Obj);

            LegacyXlsImportReport report = legacy.CreateImportReport();
            Assert.Equal(1, report.DrawingRecordsByObjectTypeName["DropdownList"]);
            Assert.Equal(1, report.DrawingObjectSubRecordsByName["FtLbsData"]);
            Assert.Equal(1, report.DrawingObjectSubRecordsByCompleteness["RequiresContinuation"]);
            Assert.DoesNotContain("DrawingObject|XLS-BIFF-FEATURE-DRAWING-UNSUPPORTED|Drawing:Obj", report.UnsupportedFeaturesByDetail.Keys);
        }

        [Fact]
        public void LegacyXls_Load_PreservesReservedDsfWorkbookMetadata() {
            byte[] workbookStream = LegacyXlsTestWorkbookBuilder.CreateReservedDsfWorkbookStream();
            byte[] compound = LegacyXlsCompoundTestBuilder.CreateWorkbookCompoundFile(workbookStream);

            LegacyXlsWorkbook legacy = LegacyXlsWorkbook.Load(compound, new LegacyXlsImportOptions {
                ReportUnsupportedContent = true
            });

            LegacyXlsWorkbookMetadataRecord metadata = Assert.Single(
                legacy.MetadataRecords,
                record => record.Kind == LegacyXlsWorkbookMetadataKind.ReservedDsf);
            Assert.Equal((ushort)BiffRecordType.Dsf, metadata.RecordType);
            Assert.DoesNotContain(legacy.UnsupportedFeatures, feature => feature.RecordType == (ushort)BiffRecordType.Dsf);
            Assert.DoesNotContain(legacy.Diagnostics, diagnostic => diagnostic.RecordType == (ushort)BiffRecordType.Dsf);
        }

        [Fact]
        public void LegacyXls_Load_PreservesVbaProjectMarkerWorkbookMetadata() {
            byte[] workbookStream = LegacyXlsTestWorkbookBuilder.CreateVbaProjectMarkerWorkbookStream();
            byte[] compound = LegacyXlsCompoundTestBuilder.CreateWorkbookCompoundFile(workbookStream);

            LegacyXlsWorkbook legacy = LegacyXlsWorkbook.Load(compound, new LegacyXlsImportOptions {
                ReportUnsupportedContent = true
            });

            Assert.True(legacy.HasVbaProjectMarker);
            LegacyXlsWorkbookMetadataRecord metadata = Assert.Single(
                legacy.MetadataRecords,
                record => record.Kind == LegacyXlsWorkbookMetadataKind.VbaProjectMarker);
            Assert.Equal((ushort)BiffRecordType.ObProj, metadata.RecordType);
            Assert.DoesNotContain(legacy.UnsupportedFeatures, feature => feature.RecordType == (ushort)BiffRecordType.ObProj);
            Assert.DoesNotContain(legacy.Diagnostics, diagnostic => diagnostic.RecordType == (ushort)BiffRecordType.ObProj);

            LegacyXlsImportReport report = new LegacyXlsImportReport(legacy);
            Assert.Equal(1, report.WorkbookMetadataRecordsByKind[LegacyXlsWorkbookMetadataKind.VbaProjectMarker]);
            Assert.Equal(1, report.VbaProjectWorkbookStates["BiffMarker:Present|NoMacrosMarker:Missing|CompoundProject:Missing|Modules:Missing"]);
        }

        [Fact]
        public void LegacyXls_Load_ImportsBuiltInFunctionGroupCountWorkbookMetadata() {
            byte[] workbookStream = LegacyXlsTestWorkbookBuilder.CreateBuiltInFunctionGroupCountWorkbookStream(0x0010);
            byte[] compound = LegacyXlsCompoundTestBuilder.CreateWorkbookCompoundFile(workbookStream);

            LegacyXlsWorkbook legacy = LegacyXlsWorkbook.Load(compound, new LegacyXlsImportOptions {
                ReportUnsupportedContent = true
            });

            Assert.Equal((ushort)0x0010, legacy.BuiltInFunctionGroupCount.GetValueOrDefault());
            LegacyXlsWorkbookMetadataRecord metadata = Assert.Single(
                legacy.MetadataRecords,
                record => record.Kind == LegacyXlsWorkbookMetadataKind.BuiltInFunctionGroupCount);
            Assert.Equal((ushort)BiffRecordType.BuiltInFnGroupCount, metadata.RecordType);
            Assert.DoesNotContain(legacy.UnsupportedFeatures, feature => feature.RecordType == (ushort)BiffRecordType.BuiltInFnGroupCount);
            Assert.DoesNotContain(legacy.Diagnostics, diagnostic => diagnostic.RecordType == (ushort)BiffRecordType.BuiltInFnGroupCount);

            LegacyXlsImportReport report = new LegacyXlsImportReport(legacy);
            Assert.Equal(1, report.WorkbookMetadataRecordsByKind[LegacyXlsWorkbookMetadataKind.BuiltInFunctionGroupCount]);
            Assert.Equal(1, report.WorkbookBuiltInFunctionGroupCounts["Count:16"]);
        }

        [Fact]
        public void LegacyXls_Load_AcceptsExcelBuiltInFunctionGroupCountSeventeenWorkbookMetadata() {
            byte[] workbookStream = LegacyXlsTestWorkbookBuilder.CreateBuiltInFunctionGroupCountWorkbookStream(0x0011);
            byte[] compound = LegacyXlsCompoundTestBuilder.CreateWorkbookCompoundFile(workbookStream);

            LegacyXlsWorkbook legacy = LegacyXlsWorkbook.Load(compound, new LegacyXlsImportOptions {
                ReportUnsupportedContent = true
            });

            Assert.Equal((ushort)0x0011, legacy.BuiltInFunctionGroupCount.GetValueOrDefault());
            Assert.DoesNotContain(legacy.UnsupportedFeatures, feature => feature.RecordType == (ushort)BiffRecordType.BuiltInFnGroupCount);
            Assert.DoesNotContain(legacy.Diagnostics, diagnostic => diagnostic.RecordType == (ushort)BiffRecordType.BuiltInFnGroupCount);

            LegacyXlsImportReport report = new LegacyXlsImportReport(legacy);
            Assert.Equal(1, report.WorkbookBuiltInFunctionGroupCounts["Count:17"]);
            Assert.Contains("Workbook Built-In Function Group Counts", report.ToMarkdown());
        }

        [Fact]
        public void LegacyXls_LoadLegacyXls_ProjectsToNormalExcelDocument() {
            byte[] workbookStream = LegacyXlsTestWorkbookBuilder.CreateMinimalWorkbookStream();
            byte[] compound = LegacyXlsCompoundTestBuilder.CreateWorkbookCompoundFile(workbookStream);

            using ExcelDocument document = ExcelDocument.LoadLegacyXls(new MemoryStream(compound), new LegacyXlsImportOptions {
                ReportUnsupportedContent = false
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
                ReportUnsupportedContent = false
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
                ReportUnsupportedContent = false
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
            Assert.True(layout.FrozenWithoutSplit);
            Assert.False(layout.ShowGridLines);
            Assert.True(layout.ShowFormulas);
            Assert.False(layout.ShowRowColumnHeadings);
            Assert.False(layout.ShowZeroValues);
            Assert.True(layout.RightToLeft);
            Assert.False(layout.DefaultGridColor);
            Assert.Equal((ushort)22, layout.GridLineColorIndex);
            Assert.False(layout.ShowOutlineSymbols);
            Assert.True(layout.TabSelected);
            Assert.True(layout.PageBreakPreview);
            Assert.Equal(120U, layout.ZoomScale);
            Assert.Equal(90U, layout.ZoomScaleNormal);
            Assert.Equal(4, layout.FirstVisibleRow);
            Assert.Equal(2, layout.FirstVisibleColumn);
            Assert.Equal(18.5d, layout.DefaultRowHeight);
            Assert.False(layout.DefaultRowsHidden);
            Assert.Equal(11d, layout.DefaultColumnWidth);
            Assert.Equal(LegacyXlsSheetVisibility.Visible, layout.VisibilityKind);
            Assert.Equal("Visible", layout.VisibilityName);
            Assert.Equal(1, legacy.Worksheets[1].Visibility);
            Assert.Equal(LegacyXlsSheetVisibility.Hidden, legacy.Worksheets[1].VisibilityKind);
            Assert.Equal("Hidden", legacy.Worksheets[1].VisibilityName);
            Assert.Equal(2, legacy.Worksheets[2].Visibility);
            Assert.Equal(LegacyXlsSheetVisibility.VeryHidden, legacy.Worksheets[2].VisibilityKind);
            Assert.Equal("VeryHidden", legacy.Worksheets[2].VisibilityName);
            LegacyXlsImportReport report = legacy.CreateImportReport();
            Assert.Equal(1, report.WorksheetsByVisibility["Visible"]);
            Assert.Equal(1, report.WorksheetsByVisibility["Hidden"]);
            Assert.Equal(1, report.WorksheetsByVisibility["VeryHidden"]);

            using ExcelDocument document = ExcelDocument.LoadLegacyXls(new MemoryStream(compound), new LegacyXlsImportOptions {
                ReportUnsupportedContent = false
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
            ExcelWorksheetViewInfo projectedView = projected.GetViewInfo();
            Assert.Equal(120U, projectedView.ZoomScale);
            Assert.Equal(90U, projectedView.ZoomScaleNormal);

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
            Assert.Equal(PaneStateValues.FrozenSplit, pane.State!.Value);
            Assert.False(sheetView.ShowGridLines!.Value);
            Assert.True(sheetView.ShowFormulas!.Value);
            Assert.False(sheetView.ShowRowColHeaders!.Value);
            Assert.False(sheetView.ShowZeros!.Value);
            Assert.True(sheetView.RightToLeft!.Value);
            Assert.False(sheetView.DefaultGridColor!.Value);
            Assert.Equal(22U, sheetView.ColorId!.Value);
            Assert.False(sheetView.ShowOutlineSymbols!.Value);
            Assert.True(sheetView.TabSelected!.Value);
            Assert.Equal(SheetViewValues.PageBreakPreview, sheetView.View!.Value);
            Assert.Equal(120U, sheetView.ZoomScale!.Value);
            Assert.Equal(90U, sheetView.ZoomScaleNormal!.Value);
            Assert.Equal("C5", sheetView.TopLeftCell!.Value);
            Assert.Equal(18.5d, sheetFormat.DefaultRowHeight!.Value);
            Assert.Equal(11d, sheetFormat.DefaultColumnWidth!.Value);
            Assert.True(sheetFormat.CustomHeight!.Value);
            Assert.Equal(2, openXmlColumn.OutlineLevel!.Value);
            Assert.True(openXmlColumn.Collapsed!.Value);
            Assert.Equal(1, openXmlRow.OutlineLevel!.Value);
            Assert.True(openXmlRow.Collapsed!.Value);
        }

        [Fact]
        public void LegacyXls_LoadLegacyXls_ImportsAndProjectsNonFrozenSplitPaneRecords() {
            byte[] workbookStream = LegacyXlsTestWorkbookBuilder.CreateSplitPaneWorkbookStream();
            byte[] compound = LegacyXlsCompoundTestBuilder.CreateWorkbookCompoundFile(workbookStream);

            using LegacyXlsLoadResult result = ExcelDocument.LoadLegacyXlsWithReport(new MemoryStream(compound), new LegacyXlsImportOptions {
                ReportUnsupportedContent = true
            });

            Assert.Empty(result.UnsupportedFeatures.Where(item =>
                item.Kind == LegacyXlsUnsupportedFeatureKind.WorksheetView
                && item.RecordType == (ushort)BiffRecordType.Pane));
            Assert.DoesNotContain(result.Diagnostics, diagnostic =>
                diagnostic.Code == "XLS-BIFF-FEATURE-WORKSHEET-VIEW-UNSUPPORTED"
                && diagnostic.RecordType == (ushort)BiffRecordType.Pane);

            LegacyXlsWorksheet legacySheet = Assert.Single(result.Workbook.Worksheets);
            Assert.NotNull(legacySheet.SplitPane);
            Assert.Equal((ushort)1200, legacySheet.SplitPane!.HorizontalSplit);
            Assert.Equal((ushort)900, legacySheet.SplitPane.VerticalSplit);
            Assert.Equal((ushort)900, legacySheet.SplitPane.TopRow);
            Assert.Equal((ushort)1200, legacySheet.SplitPane.LeftColumn);
            Assert.Equal((byte)0, legacySheet.SplitPane.ActivePane);

            Pane projectedPane = result.Document.Sheets.Single().WorksheetPart.Worksheet
                .GetFirstChild<SheetViews>()!
                .GetFirstChild<SheetView>()!
                .GetFirstChild<Pane>()!;
            Assert.Equal(PaneStateValues.Split, projectedPane.State!.Value);
            Assert.Equal(1200D, projectedPane.HorizontalSplit!.Value);
            Assert.Equal(900D, projectedPane.VerticalSplit!.Value);
            Assert.Equal(A1.CellReference(901, 1201), projectedPane.TopLeftCell!.Value);
            Assert.Equal(PaneValues.BottomRight, projectedPane.ActivePane!.Value);
        }

        [Fact]
        public void LegacyXls_LoadLegacyXls_ImportsAndProjectsPhase3StyleRecords() {
            byte[] workbookStream = LegacyXlsTestWorkbookBuilder.CreatePhase3StyleWorkbookStream();
            byte[] compound = LegacyXlsCompoundTestBuilder.CreateWorkbookCompoundFile(workbookStream);

            LegacyXlsWorkbook legacy = LegacyXlsWorkbook.Load(compound, new LegacyXlsImportOptions {
                ReportUnsupportedContent = false
            });

            Assert.DoesNotContain(legacy.Diagnostics, d => d.Severity == LegacyXlsDiagnosticSeverity.Error);
            LegacyXlsNumberFormat numberFormat = Assert.Single(legacy.NumberFormats);
            Assert.Equal(164, numberFormat.FormatId);
            Assert.Equal("yyyy-mm-dd", numberFormat.FormatCode);
            Assert.Equal(13, legacy.CellFormats.Count);
            Assert.Equal(2, legacy.CellStyles.Count);
            Assert.Equal(3, legacy.CellStyleExtensions.Count);
            LegacyXlsCellStyleExtension xfCrcExtension = Assert.Single(legacy.CellStyleExtensions, extension => extension.RecordType == 0x087c);
            Assert.Equal("XFCRC", xfCrcExtension.RecordName);
            Assert.Equal((ushort)13, xfCrcExtension.XfRecordCount.GetValueOrDefault());
            Assert.Equal(0x12345678U, xfCrcExtension.Checksum.GetValueOrDefault());
            LegacyXlsCellStyleExtension xfExtension = Assert.Single(legacy.CellStyleExtensions, extension => extension.RecordType == 0x087d);
            Assert.Equal("XfExt", xfExtension.RecordName);
            Assert.True(xfExtension.HasFormatIndex);
            Assert.Equal((ushort)4, xfExtension.FormatIndex);
            Assert.True(xfExtension.HasExtensionCount);
            Assert.Equal((ushort)5, xfExtension.ExtensionCount);
            Assert.Equal(5, xfExtension.Properties.Count);
            Assert.Equal((ushort)0x000e, xfExtension.Properties[0].PropertyType);
            Assert.Equal("FontScheme", xfExtension.Properties[0].PropertyTypeName);
            Assert.Equal(2, xfExtension.Properties[0].DataByteCount);
            Assert.Equal((ushort)1, xfExtension.Properties[0].NumericValue);
            Assert.Equal("Major", xfExtension.Properties[0].NumericValueName);
            Assert.Equal((ushort)0x000e, xfExtension.Properties[1].PropertyType);
            Assert.Equal("FontScheme", xfExtension.Properties[1].PropertyTypeName);
            Assert.Equal(1, xfExtension.Properties[1].DataByteCount);
            Assert.Equal((ushort)2, xfExtension.Properties[1].NumericValue);
            Assert.Equal("Minor", xfExtension.Properties[1].NumericValueName);
            Assert.Equal((ushort)0x000f, xfExtension.Properties[2].PropertyType);
            Assert.Equal("Indentation", xfExtension.Properties[2].PropertyTypeName);
            Assert.Equal(2, xfExtension.Properties[2].DataByteCount);
            Assert.Equal((ushort)3, xfExtension.Properties[2].NumericValue);
            Assert.Equal("Indent:3", xfExtension.Properties[2].NumericValueName);
            Assert.Equal((ushort)0x000d, xfExtension.Properties[3].PropertyType);
            Assert.Equal("TextColor", xfExtension.Properties[3].PropertyTypeName);
            Assert.Equal(16, xfExtension.Properties[3].DataByteCount);
            Assert.Equal((ushort)0x0003, xfExtension.Properties[3].ColorType);
            Assert.Equal("Theme", xfExtension.Properties[3].ColorTypeName);
            Assert.Equal((short)0, xfExtension.Properties[3].ColorTintShade);
            Assert.Equal(0x00000000U, xfExtension.Properties[3].ColorValue);
            Assert.Equal("0x00000000", xfExtension.Properties[3].ColorValueHex);
            Assert.Equal((ushort)0x0004, xfExtension.Properties[4].PropertyType);
            Assert.Equal("FillForegroundColor", xfExtension.Properties[4].PropertyTypeName);
            Assert.Equal(16, xfExtension.Properties[4].DataByteCount);
            Assert.Equal((ushort)0x0003, xfExtension.Properties[4].ColorType);
            Assert.Equal("Theme", xfExtension.Properties[4].ColorTypeName);
            Assert.Equal((short)13106, xfExtension.Properties[4].ColorTintShade);
            Assert.Equal(0x00000004U, xfExtension.Properties[4].ColorValue);
            Assert.Equal("0x00000004", xfExtension.Properties[4].ColorValueHex);
            LegacyXlsCellStyleExtension styleExtension = Assert.Single(legacy.CellStyleExtensions, extension => extension.RecordType == 0x0892);
            Assert.Equal("StyleExt", styleExtension.RecordName);
            Assert.False(styleExtension.HasFormatIndex);
            Assert.False(styleExtension.HasExtensionCount);
            Assert.False(styleExtension.IsBuiltInStyle.GetValueOrDefault());
            Assert.False(styleExtension.IsHidden.GetValueOrDefault());
            Assert.False(styleExtension.IsCustom.GetValueOrDefault());
            Assert.Equal((byte)0, styleExtension.StyleCategory.GetValueOrDefault());
            Assert.Equal("Custom", styleExtension.StyleCategoryName);
            Assert.Equal((ushort)0xffff, styleExtension.BuiltInData.GetValueOrDefault());
            Assert.Equal("OfficeIMO Accent", styleExtension.StyleName);
            Assert.Equal(6, styleExtension.Properties.Count);
            Assert.Equal((ushort)0x0000, styleExtension.Properties[0].PropertyType);
            Assert.Equal("FillPattern", styleExtension.Properties[0].PropertyTypeName);
            Assert.Equal((ushort)1, styleExtension.Properties[0].NumericValue);
            Assert.Equal("Solid", styleExtension.Properties[0].NumericValueName);
            Assert.Equal((ushort)0x0001, styleExtension.Properties[1].PropertyType);
            Assert.Equal("FillForegroundColor", styleExtension.Properties[1].PropertyTypeName);
            Assert.Equal((ushort)0x0003, styleExtension.Properties[1].ColorType);
            Assert.Equal("Theme", styleExtension.Properties[1].ColorTypeName);
            Assert.Equal((short)13106, styleExtension.Properties[1].ColorTintShade);
            Assert.Equal(0x00000004U, styleExtension.Properties[1].ColorValue);
            Assert.Equal((ushort)0x0002, styleExtension.Properties[2].PropertyType);
            Assert.Equal("FillBackgroundColor", styleExtension.Properties[2].PropertyTypeName);
            Assert.Equal((ushort)0x0002, styleExtension.Properties[2].ColorType);
            Assert.Equal("Rgb", styleExtension.Properties[2].ColorTypeName);
            Assert.Equal("0xFF00AA00", styleExtension.Properties[2].ColorValueHex);
            Assert.Equal((ushort)0x0005, styleExtension.Properties[3].PropertyType);
            Assert.Equal("TextColor", styleExtension.Properties[3].PropertyTypeName);
            Assert.Equal((ushort)0x0003, styleExtension.Properties[3].ColorType);
            Assert.Equal("Theme", styleExtension.Properties[3].ColorTypeName);
            Assert.Equal(0x00000000U, styleExtension.Properties[3].ColorValue);
            Assert.Equal((ushort)0x0006, styleExtension.Properties[4].PropertyType);
            Assert.Equal("TopBorder", styleExtension.Properties[4].PropertyTypeName);
            Assert.Equal((ushort)0x0002, styleExtension.Properties[4].ColorType);
            Assert.Equal("Rgb", styleExtension.Properties[4].ColorTypeName);
            Assert.Equal("0xFF0066FF", styleExtension.Properties[4].ColorValueHex);
            Assert.Equal((ushort)1, styleExtension.Properties[4].BorderStyle);
            Assert.Equal("Thin", styleExtension.Properties[4].BorderStyleName);
            Assert.Equal((ushort)0x0025, styleExtension.Properties[5].PropertyType);
            Assert.Equal("FontScheme", styleExtension.Properties[5].PropertyTypeName);
            Assert.Equal((ushort)2, styleExtension.Properties[5].NumericValue);
            Assert.Equal("Minor", styleExtension.Properties[5].NumericValueName);
            LegacyXlsCellStyle builtInStyle = legacy.CellStyles[0];
            Assert.True(builtInStyle.IsBuiltIn);
            Assert.Equal(0, builtInStyle.StyleFormatIndex);
            Assert.Equal((byte)0, builtInStyle.BuiltInStyleId.GetValueOrDefault());
            Assert.Equal((byte)0, builtInStyle.OutlineLevel.GetValueOrDefault());
            Assert.Null(builtInStyle.Name);
            LegacyXlsCellStyle customStyle = legacy.CellStyles[1];
            Assert.False(customStyle.IsBuiltIn);
            Assert.Equal(4, customStyle.StyleFormatIndex);
            Assert.Equal("OfficeIMO Accent", customStyle.Name);
            Assert.Null(customStyle.BuiltInStyleId);
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
            Assert.Equal((byte)2, legacy.Fonts[4].Family);
            Assert.Equal((byte)238, legacy.Fonts[4].CharacterSet);
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
            Assert.True(legacy.CellFormats[6].ApplyBorder);
            Assert.Equal(1, legacy.CellFormats[6].Border!.LeftStyle);
            Assert.Equal(0x000a, legacy.CellFormats[6].Border!.LeftColorIndex);
            Assert.Equal(8, legacy.CellFormats[6].Border!.RightStyle);
            Assert.Equal(3, legacy.CellFormats[6].Border!.TopStyle);
            Assert.Equal(6, legacy.CellFormats[6].Border!.BottomStyle);
            Assert.Equal(4, legacy.CellFormats[6].Border!.DiagonalStyle);
            Assert.True(legacy.CellFormats[6].Border!.DiagonalUp);
            Assert.True(legacy.CellFormats[6].Border!.DiagonalDown);
            Assert.True(legacy.CellFormats[12].ApplyBorder);
            Assert.Null(legacy.CellFormats[12].Border);
            Assert.True(legacy.CellFormats[8].ApplyProtection);
            Assert.False(legacy.CellFormats[8].Locked);
            Assert.True(legacy.CellFormats[8].FormulaHidden);
            Assert.True(legacy.CellFormats[9].QuotePrefix);
            Assert.Equal(2, legacy.Worksheets[0].Columns.Count);
            LegacyXlsColumnLayout defaultStyledColumn = Assert.Single(legacy.Worksheets[0].Columns, column => column.StartColumn == 11);
            Assert.Equal(11, defaultStyledColumn.StartColumn);
            Assert.Equal(11, defaultStyledColumn.EndColumn);
            Assert.Equal(4, defaultStyledColumn.StyleIndex);
            LegacyXlsColumnLayout fontStyledColumn = Assert.Single(legacy.Worksheets[0].Columns, column => column.StartColumn == 12);
            Assert.Equal(12, fontStyledColumn.EndColumn);
            Assert.Equal(10, fontStyledColumn.StyleIndex);
            Assert.Equal(2, legacy.Worksheets[0].Rows.Count);
            LegacyXlsRowLayout defaultStyledRow = Assert.Single(legacy.Worksheets[0].Rows, row => row.Row == 3);
            Assert.Equal(3, defaultStyledRow.Row);
            Assert.Equal((ushort?)5, defaultStyledRow.StyleIndex);
            LegacyXlsRowLayout fontStyledRow = Assert.Single(legacy.Worksheets[0].Rows, row => row.Row == 4);
            Assert.Equal((ushort?)10, fontStyledRow.StyleIndex);
            using LegacyXlsLoadResult reportResult = ExcelDocument.LoadLegacyXlsWithReport(new MemoryStream(compound), new LegacyXlsImportOptions {
                ReportUnsupportedContent = true
            });

            Assert.Equal(2, reportResult.ImportReport.CellStyleRecordCount);
            Assert.Equal(3, reportResult.ImportReport.CellStyleExtensionRecordCount);
            Assert.Equal(1, reportResult.ImportReport.CellStylesByKind["BuiltIn"]);
            Assert.Equal(1, reportResult.ImportReport.CellStylesByKind["Custom"]);
            Assert.Equal(1, reportResult.ImportReport.CellStyleExtensionsByRecordName["XFCRC"]);
            Assert.Equal(1, reportResult.ImportReport.CellStyleExtensionsByRecordName["XfExt"]);
            Assert.Equal(1, reportResult.ImportReport.CellStyleExtensionsByRecordName["StyleExt"]);
            Assert.Equal(1, reportResult.ImportReport.CellStyleExtensionsByFormatIndex["FormatIndex:4"]);
            Assert.Equal(1, reportResult.ImportReport.CellStyleExtensionsByExtensionCount["Extensions:5"]);
            Assert.Equal(1, reportResult.ImportReport.CellStyleExtensionsByStyleCategory["Custom"]);
            Assert.Equal(1, reportResult.ImportReport.CellStyleExtensionsByStyleFlags["BuiltIn:False;Hidden:False;Custom:False"]);
            Assert.Equal(1, reportResult.ImportReport.CellStyleExtensionsByStyleName["OfficeIMO Accent"]);
            Assert.Equal(1, reportResult.ImportReport.CellStyleExtensionsByXfRecordCount["XFs:13"]);
            Assert.Equal(1, reportResult.ImportReport.CellStyleExtensionsByChecksum["Checksum:0x12345678"]);
            Assert.Equal(2, reportResult.ImportReport.CellStyleExtensionPropertiesByType["0x000E"]);
            Assert.Equal(1, reportResult.ImportReport.CellStyleExtensionPropertiesByType["0x000F"]);
            Assert.Equal(1, reportResult.ImportReport.CellStyleExtensionPropertiesByType["0x000D"]);
            Assert.Equal(1, reportResult.ImportReport.CellStyleExtensionPropertiesByType["0x0004"]);
            Assert.Equal(1, reportResult.ImportReport.CellStyleExtensionPropertiesByType["0x0000"]);
            Assert.Equal(1, reportResult.ImportReport.CellStyleExtensionPropertiesByType["0x0001"]);
            Assert.Equal(1, reportResult.ImportReport.CellStyleExtensionPropertiesByType["0x0002"]);
            Assert.Equal(1, reportResult.ImportReport.CellStyleExtensionPropertiesByType["0x0005"]);
            Assert.Equal(1, reportResult.ImportReport.CellStyleExtensionPropertiesByType["0x0006"]);
            Assert.Equal(1, reportResult.ImportReport.CellStyleExtensionPropertiesByType["0x0025"]);
            Assert.Equal(3, reportResult.ImportReport.CellStyleExtensionPropertiesByName["FontScheme"]);
            Assert.Equal(1, reportResult.ImportReport.CellStyleExtensionPropertiesByName["Indentation"]);
            Assert.Equal(2, reportResult.ImportReport.CellStyleExtensionPropertiesByName["TextColor"]);
            Assert.Equal(2, reportResult.ImportReport.CellStyleExtensionPropertiesByName["FillForegroundColor"]);
            Assert.Equal(1, reportResult.ImportReport.CellStyleExtensionPropertiesByName["FillPattern"]);
            Assert.Equal(1, reportResult.ImportReport.CellStyleExtensionPropertiesByName["FillBackgroundColor"]);
            Assert.Equal(1, reportResult.ImportReport.CellStyleExtensionPropertiesByName["TopBorder"]);
            Assert.Equal(2, reportResult.ImportReport.CellStyleExtensionPropertiesByDataByteCount["Bytes:2"]);
            Assert.Equal(3, reportResult.ImportReport.CellStyleExtensionPropertiesByDataByteCount["Bytes:1"]);
            Assert.Equal(2, reportResult.ImportReport.CellStyleExtensionPropertiesByDataByteCount["Bytes:16"]);
            Assert.Equal(3, reportResult.ImportReport.CellStyleExtensionPropertiesByDataByteCount["Bytes:8"]);
            Assert.Equal(1, reportResult.ImportReport.CellStyleExtensionPropertiesByDataByteCount["Bytes:10"]);
            Assert.Equal(1, reportResult.ImportReport.CellStyleExtensionPropertiesByNumericValue["FontScheme:1"]);
            Assert.Equal(2, reportResult.ImportReport.CellStyleExtensionPropertiesByNumericValue["FontScheme:2"]);
            Assert.Equal(1, reportResult.ImportReport.CellStyleExtensionPropertiesByNumericValue["Indentation:3"]);
            Assert.Equal(1, reportResult.ImportReport.CellStyleExtensionPropertiesByNumericValue["FillPattern:1"]);
            Assert.Equal(1, reportResult.ImportReport.CellStyleExtensionPropertiesByNumericValueName["FontScheme:Major"]);
            Assert.Equal(2, reportResult.ImportReport.CellStyleExtensionPropertiesByNumericValueName["FontScheme:Minor"]);
            Assert.Equal(1, reportResult.ImportReport.CellStyleExtensionPropertiesByNumericValueName["Indentation:Indent:3"]);
            Assert.Equal(2, reportResult.ImportReport.CellStyleExtensionPropertiesByColorType["TextColor:Theme"]);
            Assert.Equal(2, reportResult.ImportReport.CellStyleExtensionPropertiesByColorType["FillForegroundColor:Theme"]);
            Assert.Equal(1, reportResult.ImportReport.CellStyleExtensionPropertiesByColorType["FillBackgroundColor:Rgb"]);
            Assert.Equal(1, reportResult.ImportReport.CellStyleExtensionPropertiesByColorType["TopBorder:Rgb"]);
            Assert.Equal(2, reportResult.ImportReport.CellStyleExtensionPropertiesByColorTintShade["TextColor:TintShade:0"]);
            Assert.Equal(2, reportResult.ImportReport.CellStyleExtensionPropertiesByColorTintShade["FillForegroundColor:TintShade:13106"]);
            Assert.Equal(2, reportResult.ImportReport.CellStyleExtensionPropertiesByColorValue["TextColor:0x00000000"]);
            Assert.Equal(2, reportResult.ImportReport.CellStyleExtensionPropertiesByColorValue["FillForegroundColor:0x00000004"]);
            Assert.Equal(1, reportResult.ImportReport.CellStyleExtensionPropertiesByColorValue["FillBackgroundColor:0xFF00AA00"]);
            Assert.Equal(1, reportResult.ImportReport.CellStyleExtensionPropertiesByColorValue["TopBorder:0xFF0066FF"]);
            Assert.DoesNotContain(reportResult.Workbook.UnsupportedFeatures, feature => feature.RecordType == 0x0293);
            Assert.DoesNotContain(reportResult.Workbook.UnsupportedFeatures, feature => feature.Kind == LegacyXlsUnsupportedFeatureKind.StyleExtension && feature.RecordType == 0x087c);
            Assert.DoesNotContain(reportResult.Workbook.UnsupportedFeatures, feature => feature.Kind == LegacyXlsUnsupportedFeatureKind.StyleExtension && feature.RecordType == 0x087d);
            Assert.DoesNotContain(reportResult.Workbook.UnsupportedFeatures, feature => feature.Kind == LegacyXlsUnsupportedFeatureKind.StyleExtension && feature.RecordType == 0x0892);
            Assert.DoesNotContain(reportResult.Workbook.PreservedFeatureRecords, record => record.Kind == LegacyXlsUnsupportedFeatureKind.StyleExtension && record.RecordType == 0x087c);
            Assert.DoesNotContain(reportResult.Workbook.PreservedFeatureRecords, record => record.Kind == LegacyXlsUnsupportedFeatureKind.StyleExtension && record.RecordType == 0x087d);
            Assert.DoesNotContain(reportResult.Workbook.PreservedFeatureRecords, record => record.Kind == LegacyXlsUnsupportedFeatureKind.StyleExtension && record.RecordType == 0x0892);
            Assert.DoesNotContain(reportResult.Workbook.Diagnostics, diagnostic => diagnostic.Code == "XLS-BIFF-FEATURE-STYLE-EXTENSION-UNSUPPORTED" && diagnostic.DetailCode == "StyleExtension:XFCRC");
            Assert.DoesNotContain(reportResult.Workbook.Diagnostics, diagnostic => diagnostic.Code == "XLS-BIFF-FEATURE-STYLE-EXTENSION-UNSUPPORTED" && diagnostic.DetailCode == "StyleExtension:XfExt");
            Assert.DoesNotContain(reportResult.Workbook.Diagnostics, diagnostic => diagnostic.Code == "XLS-BIFF-FEATURE-STYLE-EXTENSION-UNSUPPORTED" && diagnostic.DetailCode == "StyleExtension:StyleExt");
            Assert.Contains("Cell style records: 2", reportResult.ImportReport.ToMarkdown());
            Assert.Contains("Cell style extension records: 3", reportResult.ImportReport.ToMarkdown());
            Assert.Contains("Cell Style Extensions By Record Name", reportResult.ImportReport.ToMarkdown());
            Assert.Contains("Cell Style Extensions By Format Index", reportResult.ImportReport.ToMarkdown());
            Assert.Contains("Cell Style Extensions By Style Category", reportResult.ImportReport.ToMarkdown());
            Assert.Contains("Cell Style Extensions By XF Record Count", reportResult.ImportReport.ToMarkdown());
            Assert.Contains("Cell Style Extension Properties By Name", reportResult.ImportReport.ToMarkdown());
            Assert.Contains("Cell Style Extension Properties By Color Type", reportResult.ImportReport.ToMarkdown());

            using ExcelDocument document = ExcelDocument.LoadLegacyXls(new MemoryStream(compound), new LegacyXlsImportOptions {
                ReportUnsupportedContent = false
            });

            using var output = new MemoryStream();
            document.Save(output);
            using SpreadsheetDocument spreadsheet = SpreadsheetDocument.Open(new MemoryStream(output.ToArray()), false);
            WorkbookPart workbookPart = spreadsheet.WorkbookPart!;
            WorksheetPart worksheetPart = workbookPart.WorksheetParts.Single();
            Dictionary<string, Cell> cells = worksheetPart.Worksheet.Descendants<Cell>()
                .ToDictionary(cell => cell.CellReference!.Value!);
            Stylesheet savedStyles = workbookPart.WorkbookStylesPart!.Stylesheet!;
            CellStyle projectedCellStyle = Assert.Single(savedStyles.CellStyles!.Elements<CellStyle>(),
                style => string.Equals(style.Name?.Value, "OfficeIMO Accent", StringComparison.Ordinal));
            Assert.NotNull(projectedCellStyle.FormatId);
            Assert.NotEqual(0U, projectedCellStyle.FormatId!.Value);
            CellFormat projectedCellStyleFormat = savedStyles.CellStyleFormats!.Elements<CellFormat>().ElementAt((int)projectedCellStyle.FormatId.Value);
            Assert.NotEqual(0U, projectedCellStyleFormat.FillId!.Value);
            Fill projectedCellStyleFill = savedStyles.Fills!.Elements<Fill>().ElementAt((int)projectedCellStyleFormat.FillId.Value);
            Assert.Equal(PatternValues.Solid, projectedCellStyleFill.PatternFill!.PatternType!.Value);
            Assert.Equal(4U, projectedCellStyleFill.PatternFill.ForegroundColor!.Theme!.Value);
            Assert.Equal("FF00AA00", projectedCellStyleFill.PatternFill.BackgroundColor!.Rgb!.Value);
            DocumentFormat.OpenXml.Spreadsheet.Font projectedCellStyleFont = savedStyles.Fonts!.Elements<DocumentFormat.OpenXml.Spreadsheet.Font>().ElementAt((int)projectedCellStyleFormat.FontId!.Value);
            Assert.Equal(1U, projectedCellStyleFont.Color!.Theme!.Value);
            Assert.Equal(FontSchemeValues.Minor, projectedCellStyleFont.FontScheme!.Val!.Value);
            Border projectedCellStyleBorder = savedStyles.Borders!.Elements<Border>().ElementAt((int)projectedCellStyleFormat.BorderId!.Value);
            Assert.Equal(BorderStyleValues.Thin, projectedCellStyleBorder.TopBorder!.Style!.Value);
            Assert.Equal("FFFF6600", projectedCellStyleBorder.TopBorder.Color!.Rgb!.Value);

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
            Assert.True(projectedFont.FontFamilyNumbering!.Val!.Value == 2U);
            OpenXmlElement projectedCharset = Assert.Single(projectedFont.ChildElements, child => child.LocalName == "charset");
            Assert.Equal("238", projectedCharset.GetAttribute("val", string.Empty).Value);

            Cell subscriptCell = cells["K1"];
            Assert.NotNull(subscriptCell.StyleIndex);
            CellFormat subscriptFormat = workbookPart.WorkbookStylesPart.Stylesheet.CellFormats!.Elements<CellFormat>().ElementAt((int)subscriptCell.StyleIndex!.Value);
            DocumentFormat.OpenXml.Spreadsheet.Font subscriptFont = workbookPart.WorkbookStylesPart.Stylesheet.Fonts!.Elements<DocumentFormat.OpenXml.Spreadsheet.Font>().ElementAt((int)subscriptFormat.FontId!.Value);
            Assert.Equal("Courier New", subscriptFont.FontName!.Val!.Value);
            Assert.Equal(VerticalAlignmentRunValues.Subscript, subscriptFont.VerticalTextAlignment!.Val!.Value);

            Cell clearedBorderCell = cells["L1"];
            Assert.NotNull(clearedBorderCell.StyleIndex);
            CellFormat clearedBorderFormat = workbookPart.WorkbookStylesPart.Stylesheet.CellFormats!.Elements<CellFormat>().ElementAt((int)clearedBorderCell.StyleIndex!.Value);
            Border clearedBorder = workbookPart.WorkbookStylesPart.Stylesheet.Borders!.Elements<Border>().ElementAt((int)(clearedBorderFormat.BorderId?.Value ?? 0U));
            Assert.True(clearedBorder.LeftBorder == null || clearedBorder.LeftBorder.Style == null);

            Cell fillCell = cells["D1"];
            Assert.NotNull(fillCell.StyleIndex);
            CellFormat fillFormat = workbookPart.WorkbookStylesPart.Stylesheet.CellFormats!.Elements<CellFormat>().ElementAt((int)fillCell.StyleIndex!.Value);
            Fill projectedFill = workbookPart.WorkbookStylesPart.Stylesheet.Fills!.Elements<Fill>().ElementAt((int)fillFormat.FillId!.Value);
            Assert.Equal(PatternValues.Solid, projectedFill.PatternFill!.PatternType!.Value);
            Assert.Equal(4U, projectedFill.PatternFill.ForegroundColor!.Theme!.Value);
            Assert.NotNull(projectedFill.PatternFill.ForegroundColor.Tint);
            Assert.InRange(projectedFill.PatternFill.ForegroundColor.Tint!.Value, 0.3999D, 0.4001D);
            DocumentFormat.OpenXml.Spreadsheet.Font fillFont = workbookPart.WorkbookStylesPart.Stylesheet.Fonts!.Elements<DocumentFormat.OpenXml.Spreadsheet.Font>().ElementAt((int)fillFormat.FontId!.Value);
            Assert.Equal(1U, fillFont.Color!.Theme!.Value);
            Assert.Null(fillFont.Color.Rgb);
            Assert.Null(fillFont.Color.Tint);
            Assert.Equal(FontSchemeValues.Minor, fillFont.FontScheme!.Val!.Value);
            Assert.True(fillFormat.ApplyAlignment!.Value);
            Assert.Equal(3U, fillFormat.Alignment!.Indent!.Value);

            Cell blankFillCell = cells["F1"];
            Assert.Null(blankFillCell.CellValue);
            Assert.NotNull(blankFillCell.StyleIndex);
            CellFormat blankFillFormat = workbookPart.WorkbookStylesPart.Stylesheet.CellFormats!.Elements<CellFormat>().ElementAt((int)blankFillCell.StyleIndex!.Value);
            Fill projectedBlankFill = workbookPart.WorkbookStylesPart.Stylesheet.Fills!.Elements<Fill>().ElementAt((int)blankFillFormat.FillId!.Value);
            Assert.Equal(4U, projectedBlankFill.PatternFill!.ForegroundColor!.Theme!.Value);
            Assert.NotNull(projectedBlankFill.PatternFill.ForegroundColor.Tint);
            Assert.InRange(projectedBlankFill.PatternFill.ForegroundColor.Tint!.Value, 0.3999D, 0.4001D);

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

            List<ExcelColumnSnapshot> columnSnapshots = document.Sheets[0].GetColumnDefinitions().ToList();
            Assert.Equal(2, columnSnapshots.Count);
            ExcelColumnSnapshot columnSnapshot = Assert.Single(columnSnapshots, column => column.StartIndex == 11);
            Assert.Equal(11, columnSnapshot.StartIndex);
            Assert.Equal(11, columnSnapshot.EndIndex);
            Assert.NotNull(columnSnapshot.StyleIndex);
            Column defaultStyledOpenXmlColumn = worksheetPart.Worksheet.GetFirstChild<Columns>()!.Elements<Column>().Single(column => column.Min!.Value == 11U);
            Assert.Equal(columnSnapshot.StyleIndex!.Value, defaultStyledOpenXmlColumn.Style!.Value);
            CellFormat columnFormat = workbookPart.WorkbookStylesPart.Stylesheet.CellFormats!.Elements<CellFormat>().ElementAt((int)columnSnapshot.StyleIndex.Value);
            Fill columnFill = workbookPart.WorkbookStylesPart.Stylesheet.Fills!.Elements<Fill>().ElementAt((int)columnFormat.FillId!.Value);
            Assert.Equal(PatternValues.Solid, columnFill.PatternFill!.PatternType!.Value);
            Assert.Equal(4U, columnFill.PatternFill.ForegroundColor!.Theme!.Value);
            Assert.NotNull(columnFill.PatternFill.ForegroundColor.Tint);
            Assert.InRange(columnFill.PatternFill.ForegroundColor.Tint!.Value, 0.3999D, 0.4001D);

            ExcelColumnSnapshot fontColumnSnapshot = Assert.Single(columnSnapshots, column => column.StartIndex == 12);
            CellFormat fontColumnFormat = workbookPart.WorkbookStylesPart.Stylesheet.CellFormats!.Elements<CellFormat>().ElementAt((int)fontColumnSnapshot.StyleIndex!.Value);
            DocumentFormat.OpenXml.Spreadsheet.Font fontColumnFont = workbookPart.WorkbookStylesPart.Stylesheet.Fonts!.Elements<DocumentFormat.OpenXml.Spreadsheet.Font>().ElementAt((int)fontColumnFormat.FontId!.Value);
            Assert.Equal(VerticalAlignmentRunValues.Subscript, fontColumnFont.VerticalTextAlignment!.Val!.Value);

            List<ExcelRowSnapshot> rowSnapshots = document.Sheets[0].GetRowDefinitions().ToList();
            Assert.Equal(2, rowSnapshots.Count);
            ExcelRowSnapshot rowSnapshot = Assert.Single(rowSnapshots, row => row.Index == 3);
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
            ExcelRowSnapshot fontRowSnapshot = Assert.Single(rowSnapshots, row => row.Index == 4);
            CellFormat fontRowFormat = workbookPart.WorkbookStylesPart.Stylesheet.CellFormats!.Elements<CellFormat>().ElementAt((int)fontRowSnapshot.StyleIndex!.Value);
            DocumentFormat.OpenXml.Spreadsheet.Font fontRowFont = workbookPart.WorkbookStylesPart.Stylesheet.Fonts!.Elements<DocumentFormat.OpenXml.Spreadsheet.Font>().ElementAt((int)fontRowFormat.FontId!.Value);
            Assert.Equal(VerticalAlignmentRunValues.Subscript, fontRowFont.VerticalTextAlignment!.Val!.Value);
        }

        [Fact]
        public void LegacyXls_LoadLegacyXls_ProjectsEmbeddedThemePackage() {
            byte[] workbookStream = LegacyXlsTestWorkbookBuilder.CreateEmbeddedThemeWorkbookStream("OfficeIMO Legacy Theme");
            byte[] compound = LegacyXlsCompoundTestBuilder.CreateWorkbookCompoundFile(workbookStream);

            using LegacyXlsLoadResult result = ExcelDocument.LoadLegacyXlsWithReport(new MemoryStream(compound), new LegacyXlsImportOptions {
                ReportUnsupportedContent = true
            });

            LegacyXlsThemeRecord themeRecord = Assert.Single(result.Workbook.ThemeRecords);
            Assert.True(themeRecord.HasThemeBytes);
            Assert.DoesNotContain(result.Workbook.UnsupportedFeatures, feature =>
                feature.Kind == LegacyXlsUnsupportedFeatureKind.Theme
                && feature.RecordType == (ushort)BiffRecordType.Theme);

            ExcelWorkbookThemeInfo projectedTheme = result.Document.GetWorkbookTheme(includeXml: true);
            Assert.True(projectedTheme.HasTheme);
            Assert.Equal("OfficeIMO Legacy Theme", projectedTheme.Name);
            Assert.Contains("OfficeIMO Legacy Theme", projectedTheme.Xml, StringComparison.Ordinal);
        }

        [Fact]
        public void LegacyXls_LoadLegacyXls_ProjectsDefaultThemeMarkerWithoutUnsupportedFeature() {
            byte[] workbookStream = LegacyXlsTestWorkbookBuilder.CreateDefaultThemeWorkbookStream();
            byte[] compound = LegacyXlsCompoundTestBuilder.CreateWorkbookCompoundFile(workbookStream);

            using LegacyXlsLoadResult result = ExcelDocument.LoadLegacyXlsWithReport(new MemoryStream(compound), new LegacyXlsImportOptions {
                ReportUnsupportedContent = true
            });

            LegacyXlsThemeRecord themeRecord = Assert.Single(result.Workbook.ThemeRecords);
            Assert.True(themeRecord.IsDefaultThemeMarker);
            Assert.DoesNotContain(result.Workbook.UnsupportedFeatures, feature =>
                feature.Kind == LegacyXlsUnsupportedFeatureKind.Theme
                && feature.RecordType == (ushort)BiffRecordType.Theme);
            Assert.DoesNotContain(LegacyXlsUnsupportedFeatureKind.Theme, result.ImportReport.UnsupportedFeaturesByKind.Keys);

            ExcelWorkbookThemeInfo projectedTheme = result.Document.GetWorkbookTheme(includeXml: true);
            Assert.True(projectedTheme.HasTheme);
            Assert.False(string.IsNullOrWhiteSpace(projectedTheme.Xml));
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

            internal static byte[] CreateWorksheetTextObjectWorkbookStream(string text) {
                using var stream = new MemoryStream();
                WriteRecord(stream, 0x0809, new byte[] { 0x00, 0x06, 0x05, 0x00, 0xdb, 0x0b, 0xcc, 0x07 });
                long boundSheetPosition = stream.Position;
                WriteRecord(stream, 0x0085, BuildBoundSheetPayload(0, "TextBox"));
                WriteRecord(stream, 0x000a, Array.Empty<byte>());

                int sheetOffset = checked((int)stream.Position);
                WriteRecord(stream, 0x0809, new byte[] { 0x00, 0x06, 0x10, 0x00, 0xdb, 0x0b, 0xcc, 0x07 });
                WriteRecord(stream, 0x01b6, BuildWorksheetTxoPayload(text.Length, formattingRunBytes: 16));
                WriteRecord(stream, 0x003c, BuildCompressedTextContinuePayload(text));
                WriteRecord(stream, 0x003c, BuildTxoFormattingRunContinuePayload(startCharacter: 0, fontIndex: 0));
                WriteRecord(stream, 0x000a, Array.Empty<byte>());

                byte[] bytes = stream.ToArray();
                byte[] offsetBytes = BitConverter.GetBytes(sheetOffset);
                Buffer.BlockCopy(offsetBytes, 0, bytes, checked((int)boundSheetPosition + 4), offsetBytes.Length);
                return bytes;
            }

            internal static byte[] CreateWorksheetClientTextboxDrawingWorkbookStream() {
                using var stream = new MemoryStream();
                WriteRecord(stream, 0x0809, new byte[] { 0x00, 0x06, 0x05, 0x00, 0xdb, 0x0b, 0xcc, 0x07 });
                long boundSheetPosition = stream.Position;
                WriteRecord(stream, 0x0085, BuildBoundSheetPayload(0, "TextBox"));
                WriteRecord(stream, 0x000a, Array.Empty<byte>());

                int sheetOffset = checked((int)stream.Position);
                WriteRecord(stream, 0x0809, new byte[] { 0x00, 0x06, 0x10, 0x00, 0xdb, 0x0b, 0xcc, 0x07 });
                WriteRecord(stream, 0x00ec, BuildOfficeArtClientTextboxPayload());
                WriteRecord(stream, 0x000a, Array.Empty<byte>());

                byte[] bytes = stream.ToArray();
                byte[] offsetBytes = BitConverter.GetBytes(sheetOffset);
                Buffer.BlockCopy(offsetBytes, 0, bytes, checked((int)boundSheetPosition + 4), offsetBytes.Length);
                return bytes;
            }

            internal static byte[] CreateWorksheetContinuedDrawingWorkbookStream() {
                using var stream = new MemoryStream();
                WriteRecord(stream, 0x0809, new byte[] { 0x00, 0x06, 0x05, 0x00, 0xdb, 0x0b, 0xcc, 0x07 });
                long boundSheetPosition = stream.Position;
                WriteRecord(stream, 0x0085, BuildBoundSheetPayload(0, "Drawing"));
                WriteRecord(stream, 0x000a, Array.Empty<byte>());

                byte[] drawingPayload = BuildDrawingWithPictureShapePayload();
                const int firstPayloadLength = 16;
                int sheetOffset = checked((int)stream.Position);
                WriteRecord(stream, 0x0809, new byte[] { 0x00, 0x06, 0x10, 0x00, 0xdb, 0x0b, 0xcc, 0x07 });
                WriteRecord(stream, 0x00ec, drawingPayload.Take(firstPayloadLength).ToArray());
                WriteRecord(stream, 0x003c, drawingPayload.Skip(firstPayloadLength).ToArray());
                WriteRecord(stream, 0x000a, Array.Empty<byte>());

                byte[] bytes = stream.ToArray();
                byte[] offsetBytes = BitConverter.GetBytes(sheetOffset);
                Buffer.BlockCopy(offsetBytes, 0, bytes, checked((int)boundSheetPosition + 4), offsetBytes.Length);
                return bytes;
            }

            internal static byte[] CreateWorksheetSplitDrawingContainerWorkbookStream() {
                using var stream = new MemoryStream();
                WriteRecord(stream, 0x0809, new byte[] { 0x00, 0x06, 0x05, 0x00, 0xdb, 0x0b, 0xcc, 0x07 });
                long boundSheetPosition = stream.Position;
                WriteRecord(stream, 0x0085, BuildBoundSheetPayload(0, "Drawing"));
                WriteRecord(stream, 0x000a, Array.Empty<byte>());

                int sheetOffset = checked((int)stream.Position);
                WriteRecord(stream, 0x0809, new byte[] { 0x00, 0x06, 0x10, 0x00, 0xdb, 0x0b, 0xcc, 0x07 });
                WriteRecord(stream, 0x00ec, BuildSplitDrawingContainerPayload());
                WriteRecord(stream, 0x005d, BuildObjectPayload(0x0005, 1));
                WriteRecord(stream, 0x000a, Array.Empty<byte>());

                byte[] bytes = stream.ToArray();
                byte[] offsetBytes = BitConverter.GetBytes(sheetOffset);
                Buffer.BlockCopy(offsetBytes, 0, bytes, checked((int)boundSheetPosition + 4), offsetBytes.Length);
                return bytes;
            }

            internal static byte[] CreateWorksheetSplitShapeContainerWorkbookStream() {
                using var stream = new MemoryStream();
                WriteRecord(stream, 0x0809, new byte[] { 0x00, 0x06, 0x05, 0x00, 0xdb, 0x0b, 0xcc, 0x07 });
                long boundSheetPosition = stream.Position;
                WriteRecord(stream, 0x0085, BuildBoundSheetPayload(0, "Shape"));
                WriteRecord(stream, 0x000a, Array.Empty<byte>());

                int sheetOffset = checked((int)stream.Position);
                WriteRecord(stream, 0x0809, new byte[] { 0x00, 0x06, 0x10, 0x00, 0xdb, 0x0b, 0xcc, 0x07 });
                WriteRecord(stream, 0x00ec, BuildSplitShapeContainerPayload());
                WriteRecord(stream, 0x000a, Array.Empty<byte>());

                byte[] bytes = stream.ToArray();
                byte[] offsetBytes = BitConverter.GetBytes(sheetOffset);
                Buffer.BlockCopy(offsetBytes, 0, bytes, checked((int)boundSheetPosition + 4), offsetBytes.Length);
                return bytes;
            }

            internal static byte[] CreateWorksheetObjectWorkbookStream() {
                using var stream = new MemoryStream();
                WriteRecord(stream, 0x0809, new byte[] { 0x00, 0x06, 0x05, 0x00, 0xdb, 0x0b, 0xcc, 0x07 });
                long boundSheetPosition = stream.Position;
                WriteRecord(stream, 0x0085, BuildBoundSheetPayload(0, "Object"));
                WriteRecord(stream, 0x000a, Array.Empty<byte>());

                int sheetOffset = checked((int)stream.Position);
                WriteRecord(stream, 0x0809, new byte[] { 0x00, 0x06, 0x10, 0x00, 0xdb, 0x0b, 0xcc, 0x07 });
                WriteRecord(stream, 0x005d, BuildWorksheetObjectPayload());
                WriteRecord(stream, 0x000a, Array.Empty<byte>());

                byte[] bytes = stream.ToArray();
                byte[] offsetBytes = BitConverter.GetBytes(sheetOffset);
                Buffer.BlockCopy(offsetBytes, 0, bytes, checked((int)boundSheetPosition + 4), offsetBytes.Length);
                return bytes;
            }

            internal static byte[] CreateWorksheetListObjectWorkbookStream() {
                using var stream = new MemoryStream();
                WriteRecord(stream, 0x0809, new byte[] { 0x00, 0x06, 0x05, 0x00, 0xdb, 0x0b, 0xcc, 0x07 });
                long boundSheetPosition = stream.Position;
                WriteRecord(stream, 0x0085, BuildBoundSheetPayload(0, "ListObj"));
                WriteRecord(stream, 0x000a, Array.Empty<byte>());

                int sheetOffset = checked((int)stream.Position);
                WriteRecord(stream, 0x0809, new byte[] { 0x00, 0x06, 0x10, 0x00, 0xdb, 0x0b, 0xcc, 0x07 });
                WriteRecord(stream, 0x005d, BuildWorksheetListObjectPayload());
                WriteRecord(stream, 0x003c, new byte[] { 0x00, 0x00, 0x00, 0x00 });
                WriteRecord(stream, 0x000a, Array.Empty<byte>());

                byte[] bytes = stream.ToArray();
                byte[] offsetBytes = BitConverter.GetBytes(sheetOffset);
                Buffer.BlockCopy(offsetBytes, 0, bytes, checked((int)boundSheetPosition + 4), offsetBytes.Length);
                return bytes;
            }

            internal static byte[] CreateEmbeddedThemeWorkbookStream(string themeName) {
                using var stream = new MemoryStream();
                WriteRecord(stream, 0x0809, new byte[] { 0x00, 0x06, 0x05, 0x00, 0xdb, 0x0b, 0xcc, 0x07 });
                long boundSheetPosition = stream.Position;
                WriteRecord(stream, 0x0085, BuildBoundSheetPayload(0, "Theme"));
                WriteRecord(stream, (ushort)BiffRecordType.Theme, BuildThemePayload(202300, BuildThemePackage(themeName)));
                WriteRecord(stream, 0x000a, Array.Empty<byte>());

                int sheetOffset = checked((int)stream.Position);
                WriteRecord(stream, 0x0809, new byte[] { 0x00, 0x06, 0x10, 0x00, 0xdb, 0x0b, 0xcc, 0x07 });
                WriteRecord(stream, 0x0204, BuildLabelPayload(0, 0, "Theme"));
                WriteRecord(stream, 0x000a, Array.Empty<byte>());

                byte[] bytes = stream.ToArray();
                Buffer.BlockCopy(BitConverter.GetBytes(sheetOffset), 0, bytes, checked((int)boundSheetPosition + 4), 4);
                return bytes;
            }

            internal static byte[] CreateDefaultThemeWorkbookStream() {
                using var stream = new MemoryStream();
                WriteRecord(stream, 0x0809, new byte[] { 0x00, 0x06, 0x05, 0x00, 0xdb, 0x0b, 0xcc, 0x07 });
                long boundSheetPosition = stream.Position;
                WriteRecord(stream, 0x0085, BuildBoundSheetPayload(0, "Theme"));
                WriteRecord(stream, (ushort)BiffRecordType.Theme, BuildThemePayload(124226));
                WriteRecord(stream, 0x000a, Array.Empty<byte>());

                int sheetOffset = checked((int)stream.Position);
                WriteRecord(stream, 0x0809, new byte[] { 0x00, 0x06, 0x10, 0x00, 0xdb, 0x0b, 0xcc, 0x07 });
                WriteRecord(stream, 0x0204, BuildLabelPayload(0, 0, "Theme"));
                WriteRecord(stream, 0x000a, Array.Empty<byte>());

                byte[] bytes = stream.ToArray();
                Buffer.BlockCopy(BitConverter.GetBytes(sheetOffset), 0, bytes, checked((int)boundSheetPosition + 4), 4);
                return bytes;
            }

            internal static byte[] CreateReservedDsfWorkbookStream() {
                using var stream = new MemoryStream();
                WriteRecord(stream, 0x0809, new byte[] { 0x00, 0x06, 0x05, 0x00, 0xdb, 0x0b, 0xcc, 0x07 });
                long boundSheetPosition = stream.Position;
                WriteRecord(stream, 0x0085, BuildBoundSheetPayload(0, "Sheet1"));
                WriteRecord(stream, 0x0161, new byte[] { 0x00, 0x00 });
                WriteRecord(stream, 0x000a, Array.Empty<byte>());

                int sheetOffset = checked((int)stream.Position);
                WriteRecord(stream, 0x0809, new byte[] { 0x00, 0x06, 0x10, 0x00, 0xdb, 0x0b, 0xcc, 0x07 });
                WriteRecord(stream, 0x0204, BuildLabelPayload(0, 0, "DSF"));
                WriteRecord(stream, 0x000a, Array.Empty<byte>());

                byte[] bytes = stream.ToArray();
                Buffer.BlockCopy(BitConverter.GetBytes(sheetOffset), 0, bytes, checked((int)boundSheetPosition + 4), 4);
                return bytes;
            }

            internal static byte[] CreateVbaProjectMarkerWorkbookStream() {
                using var stream = new MemoryStream();
                WriteRecord(stream, 0x0809, new byte[] { 0x00, 0x06, 0x05, 0x00, 0xdb, 0x0b, 0xcc, 0x07 });
                long boundSheetPosition = stream.Position;
                WriteRecord(stream, 0x0085, BuildBoundSheetPayload(0, "Sheet1"));
                WriteRecord(stream, 0x00d3, Array.Empty<byte>());
                WriteRecord(stream, 0x000a, Array.Empty<byte>());

                int sheetOffset = checked((int)stream.Position);
                WriteRecord(stream, 0x0809, new byte[] { 0x00, 0x06, 0x10, 0x00, 0xdb, 0x0b, 0xcc, 0x07 });
                WriteRecord(stream, 0x0204, BuildLabelPayload(0, 0, "VBA marker"));
                WriteRecord(stream, 0x000a, Array.Empty<byte>());

                byte[] bytes = stream.ToArray();
                Buffer.BlockCopy(BitConverter.GetBytes(sheetOffset), 0, bytes, checked((int)boundSheetPosition + 4), 4);
                return bytes;
            }

            internal static byte[] CreateBuiltInFunctionGroupCountWorkbookStream(ushort functionGroupCount) {
                using var stream = new MemoryStream();
                WriteRecord(stream, 0x0809, new byte[] { 0x00, 0x06, 0x05, 0x00, 0xdb, 0x0b, 0xcc, 0x07 });
                long boundSheetPosition = stream.Position;
                WriteRecord(stream, 0x0085, BuildBoundSheetPayload(0, "Sheet1"));
                WriteRecord(stream, 0x009c, BuildUInt16Payload(functionGroupCount));
                WriteRecord(stream, 0x000a, Array.Empty<byte>());

                int sheetOffset = checked((int)stream.Position);
                WriteRecord(stream, 0x0809, new byte[] { 0x00, 0x06, 0x10, 0x00, 0xdb, 0x0b, 0xcc, 0x07 });
                WriteRecord(stream, 0x0204, BuildLabelPayload(0, 0, "Function groups"));
                WriteRecord(stream, 0x000a, Array.Empty<byte>());

                byte[] bytes = stream.ToArray();
                Buffer.BlockCopy(BitConverter.GetBytes(sheetOffset), 0, bytes, checked((int)boundSheetPosition + 4), 4);
                return bytes;
            }

            internal static byte[] CreateEncryptedWorkbookStream() {
                using var stream = new MemoryStream();
                WriteRecord(stream, 0x0809, new byte[] { 0x00, 0x06, 0x05, 0x00, 0xdb, 0x0b, 0xcc, 0x07 });
                WriteRecord(stream, 0x002f, new byte[] { 0x00, 0x00 });
                WriteRecord(stream, 0x000a, Array.Empty<byte>());
                return stream.ToArray();
            }

            internal static byte[] CreateMalformedRc4EncryptedWorkbookStream() {
                using var stream = new MemoryStream();
                WriteRecord(stream, 0x0809, new byte[] { 0x00, 0x06, 0x05, 0x00, 0xdb, 0x0b, 0xcc, 0x07 });
                WriteRecord(stream, 0x002f, new byte[] { 0x01, 0x00 });
                WriteRecord(stream, 0x000a, Array.Empty<byte>());
                return stream.ToArray();
            }

            internal static byte[] CreateRc4EncryptedWorkbookStream(string password) {
                byte[] salt = {
                    0x10, 0x22, 0x34, 0x46, 0x58, 0x6A, 0x7C, 0x8E,
                    0x90, 0xA2, 0xB4, 0xC6, 0xD8, 0xEA, 0xFC, 0x0E
                };
                byte[] verifier = {
                    0xA5, 0x5A, 0x19, 0x91, 0x28, 0x82, 0x37, 0x73,
                    0x46, 0x64, 0x55, 0x99, 0xAB, 0xBA, 0xCD, 0xDC
                };

                using var stream = new MemoryStream();
                WriteRecord(stream, 0x0809, new byte[] { 0x00, 0x06, 0x05, 0x00, 0xdb, 0x0b, 0xcc, 0x07 });
                WriteRecord(stream, 0x002f, BiffRc4Encryption.CreateFilePassPayload(password, salt, verifier));
                long boundSheetPosition = stream.Position;
                WriteRecord(stream, 0x0085, BuildBoundSheetPayload(0, "Rc4Sheet"));
                WriteRecord(stream, 0x000a, Array.Empty<byte>());

                int sheetOffset = checked((int)stream.Position);
                WriteRecord(stream, 0x0809, new byte[] { 0x00, 0x06, 0x10, 0x00, 0xdb, 0x0b, 0xcc, 0x07 });
                WriteRecord(stream, 0x0204, BuildLabelPayload(0, 0, "RC4 secret"));
                WriteRecord(stream, 0x000a, Array.Empty<byte>());

                byte[] bytes = stream.ToArray();
                Buffer.BlockCopy(BitConverter.GetBytes(sheetOffset), 0, bytes, checked((int)boundSheetPosition + 4), 4);
                return BiffRc4Encryption.EncryptWorkbookStream(bytes, password, salt);
            }

            internal static byte[] CreateXorObfuscatedWorkbookStream(string password) {
                using var stream = new MemoryStream();
                WriteRecord(stream, 0x0809, new byte[] { 0x00, 0x06, 0x05, 0x00, 0xdb, 0x0b, 0xcc, 0x07 });
                WriteRecord(stream, 0x002f, BuildXorFilePassPayload(password));
                long boundSheetPosition = stream.Position;
                WriteRecord(stream, 0x0085, BuildBoundSheetPayload(0, "XorSheet"));
                WriteRecord(stream, 0x000a, Array.Empty<byte>());

                int sheetOffset = checked((int)stream.Position);
                WriteRecord(stream, 0x0809, new byte[] { 0x00, 0x06, 0x10, 0x00, 0xdb, 0x0b, 0xcc, 0x07 });
                WriteRecord(stream, 0x0204, BuildLabelPayload(0, 0, "XOR secret"));
                WriteRecord(stream, 0x000a, Array.Empty<byte>());

                byte[] bytes = stream.ToArray();
                Buffer.BlockCopy(BitConverter.GetBytes(sheetOffset), 0, bytes, checked((int)boundSheetPosition + 4), 4);
                return BiffXorObfuscation.ObfuscateWorkbookStream(bytes, password);
            }

            internal static byte[] CreateEncryptedWorkbookWithUnreadablePayloadStream() {
                using var stream = new MemoryStream();
                WriteRecord(stream, 0x0809, new byte[] { 0x00, 0x06, 0x05, 0x00, 0xdb, 0x0b, 0xcc, 0x07 });
                WriteRecord(stream, 0x002f, new byte[] { 0x00, 0x00 });
                WriteRecord(stream, 0x00c1, new byte[8]);
                WriteRecord(stream, 0x01c0, new byte[4]);
                WriteRecord(stream, 0x087d, new byte[12]);
                WriteRecord(stream, 0x000a, Array.Empty<byte>());
                return stream.ToArray();
            }

            private static byte[] BuildXorFilePassPayload(string password) {
                byte[] payload = new byte[6];
                Buffer.BlockCopy(BitConverter.GetBytes((ushort)0x0000), 0, payload, 0, 2);
                Buffer.BlockCopy(BitConverter.GetBytes(BiffXorObfuscation.CreateXorKey(password)), 0, payload, 2, 2);
                Buffer.BlockCopy(BitConverter.GetBytes(BiffXorObfuscation.CreatePasswordVerifier(password)), 0, payload, 4, 2);
                return payload;
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

            internal static byte[] CreateKnownPreserveOnlyExtensionWorkbookStream() {
                using var stream = new MemoryStream();
                WriteRecord(stream, 0x0809, new byte[] { 0x00, 0x06, 0x05, 0x00, 0xdb, 0x0b, 0xcc, 0x07 });
                long boundSheetPosition = stream.Position;
                WriteRecord(stream, 0x0085, BuildBoundSheetPayload(0, "Extensions"));
                WriteRecord(stream, 0x0051, new byte[12]);
                WriteRecord(stream, (ushort)BiffRecordType.DbQueryExt, BuildExternalQueryConnectionPayload());
                WriteRecord(stream, 0x00e3, new byte[8]);
                WriteRecord(stream, 0x01c0, new byte[4]);
                WriteRecord(stream, 0x01c1, new byte[4]);
                WriteRecord(stream, 0x0810, new byte[12]);
                WriteRecord(stream, 0x0867, new byte[16]);
                WriteRecord(stream, (ushort)BiffRecordType.Plv, BuildFutureRecordPayload((ushort)BiffRecordType.Plv, 16));
                WriteRecord(stream, (ushort)BiffRecordType.Compat12, BuildFutureRecordPayload((ushort)BiffRecordType.Compat12, 16));
                WriteRecord(stream, (ushort)BiffRecordType.Dxf, Array.Empty<byte>());
                WriteRecord(stream, (ushort)BiffRecordType.TableStyles, BuildTableStylesPayload(145, "TableStyleMedium2", "PivotStyleLight16"));
                WriteRecord(stream, (ushort)BiffRecordType.TableStyle, BuildTableStylePayload("OfficeIMO Custom", appliesToTables: true, appliesToPivotTables: true, declaredElementCount: 1));
                WriteRecord(stream, (ushort)BiffRecordType.TableStyleElement, BuildTableStyleElementPayload(0x00000001, 0, 3));
                WriteRecord(stream, 0x0896, BuildThemePayload(124226));
                WriteRecord(stream, 0x0897, BuildFutureRecordPayload(0x0897, 16));
                WriteRecord(stream, 0x0899, BuildFutureRecordPayload(0x0899, 12));
                WriteRecord(stream, 0x089a, new byte[4]);
                WriteRecord(stream, 0x089b, new byte[12]);
                WriteRecord(stream, 0x089c, BuildFutureRecordPayload(0x089c, 12));
                WriteRecord(stream, 0x000a, Array.Empty<byte>());

                int sheetOffset = checked((int)stream.Position);
                WriteRecord(stream, 0x0809, new byte[] { 0x00, 0x06, 0x10, 0x00, 0xdb, 0x0b, 0xcc, 0x07 });
                WriteRecord(stream, 0x0204, BuildLabelPayload(0, 0, "Extension"));
                WriteRecord(stream, 0x00ed, new byte[8]);
                WriteRecord(stream, 0x089d, BuildFutureRecordPayload(0x089d, 12));
                WriteRecord(stream, 0x089e, BuildFutureRecordPayload(0x089e, 12));
                WriteRecord(stream, 0x089f, BuildFutureRecordPayload(0x089f, 12));
                WriteRecord(stream, 0x08a3, new byte[12]);
                WriteRecord(stream, 0x08a4, new byte[12]);
                WriteRecord(stream, 0x08a5, new byte[12]);
                WriteRecord(stream, 0x08a7, new byte[12]);
                WriteRecord(stream, 0x000a, Array.Empty<byte>());

                byte[] bytes = stream.ToArray();
                Buffer.BlockCopy(BitConverter.GetBytes(sheetOffset), 0, bytes, checked((int)boundSheetPosition + 4), 4);
                return bytes;
            }

            internal static byte[] CreatePhase5TableStyleMetadataWorkbookStream() {
                using var stream = new MemoryStream();
                WriteRecord(stream, 0x0809, new byte[] { 0x00, 0x06, 0x05, 0x00, 0xdb, 0x0b, 0xcc, 0x07 });
                long boundSheetPosition = stream.Position;
                WriteRecord(stream, 0x0085, BuildBoundSheetPayload(0, "TableStyles"));
                WriteRecord(stream, (ushort)BiffRecordType.TableStyles, BuildTableStylesPayload(145, "TableStyleMedium2", "PivotStyleLight16"));
                WriteRecord(stream, (ushort)BiffRecordType.TableStyle, BuildTableStylePayload("OfficeIMO Custom", appliesToTables: true, appliesToPivotTables: true, declaredElementCount: 2));
                WriteRecord(stream, (ushort)BiffRecordType.TableStyleElement, BuildTableStyleElementPayload(0x00000001, 0, 3));
                WriteRecord(stream, (ushort)BiffRecordType.TableStyleElement, BuildTableStyleElementPayload(0x00000005, 2, 4));
                WriteRecord(stream, 0x000a, Array.Empty<byte>());

                int sheetOffset = checked((int)stream.Position);
                WriteRecord(stream, 0x0809, new byte[] { 0x00, 0x06, 0x10, 0x00, 0xdb, 0x0b, 0xcc, 0x07 });
                WriteRecord(stream, 0x0204, BuildLabelPayload(0, 0, "Styled table metadata"));
                WriteRecord(stream, 0x000a, Array.Empty<byte>());

                byte[] bytes = stream.ToArray();
                Buffer.BlockCopy(BitConverter.GetBytes(sheetOffset), 0, bytes, checked((int)boundSheetPosition + 4), 4);
                return bytes;
            }

            internal static byte[] CreatePhase5CustomTableStyleOnlyWorkbookStream() {
                using var stream = new MemoryStream();
                WriteRecord(stream, 0x0809, new byte[] { 0x00, 0x06, 0x05, 0x00, 0xdb, 0x0b, 0xcc, 0x07 });
                long boundSheetPosition = stream.Position;
                WriteRecord(stream, 0x0085, BuildBoundSheetPayload(0, "TableStyleOnly"));
                WriteRecord(stream, (ushort)BiffRecordType.TableStyle, BuildTableStylePayload("OfficeIMO Custom Only", appliesToTables: true, appliesToPivotTables: false, declaredElementCount: 1));
                WriteRecord(stream, (ushort)BiffRecordType.TableStyleElement, BuildTableStyleElementPayload(0x00000001, 0, 3));
                WriteRecord(stream, 0x000a, Array.Empty<byte>());

                int sheetOffset = checked((int)stream.Position);
                WriteRecord(stream, 0x0809, new byte[] { 0x00, 0x06, 0x10, 0x00, 0xdb, 0x0b, 0xcc, 0x07 });
                WriteRecord(stream, 0x0204, BuildLabelPayload(0, 0, "Custom style metadata"));
                WriteRecord(stream, 0x000a, Array.Empty<byte>());

                byte[] bytes = stream.ToArray();
                Buffer.BlockCopy(BitConverter.GetBytes(sheetOffset), 0, bytes, checked((int)boundSheetPosition + 4), 4);
                return bytes;
            }

            private static byte[] BuildTableStylesPayload(uint totalStyleCount, string defaultTableStyleName, string defaultPivotStyleName) {
                using var stream = new MemoryStream();
                WriteFutureRecordHeader(stream, (ushort)BiffRecordType.TableStyles);
                WriteUInt32(stream, totalStyleCount);
                WriteUInt16(stream, checked((ushort)defaultTableStyleName.Length));
                WriteUInt16(stream, checked((ushort)defaultPivotStyleName.Length));
                WriteUnicodeCharacters(stream, defaultTableStyleName);
                WriteUnicodeCharacters(stream, defaultPivotStyleName);
                return stream.ToArray();
            }

            private static byte[] BuildTableStylePayload(string name, bool appliesToTables, bool appliesToPivotTables, uint declaredElementCount) {
                using var stream = new MemoryStream();
                WriteFutureRecordHeader(stream, (ushort)BiffRecordType.TableStyle);
                ushort flags = 0;
                if (appliesToPivotTables) {
                    flags |= 0x0002;
                }

                if (appliesToTables) {
                    flags |= 0x0004;
                }

                WriteUInt16(stream, flags);
                WriteUInt32(stream, declaredElementCount);
                WriteUInt16(stream, checked((ushort)name.Length));
                WriteUnicodeCharacters(stream, name);
                return stream.ToArray();
            }

            private static byte[] BuildTableStyleElementPayload(uint elementType, uint stripeSize, uint differentialFormatIndex) {
                using var stream = new MemoryStream();
                WriteFutureRecordHeader(stream, (ushort)BiffRecordType.TableStyleElement);
                WriteUInt32(stream, elementType);
                WriteUInt32(stream, stripeSize);
                WriteUInt32(stream, differentialFormatIndex);
                return stream.ToArray();
            }

            private static void WriteFutureRecordHeader(Stream stream, ushort recordType) {
                WriteUInt16(stream, recordType);
                WriteUInt16(stream, 0);
                WriteUInt32(stream, 0);
                WriteUInt32(stream, 0);
            }

            private static void WriteUnicodeCharacters(Stream stream, string value) {
                byte[] bytes = Encoding.Unicode.GetBytes(value);
                stream.Write(bytes, 0, bytes.Length);
            }

            internal static byte[] CreateObservedCorpusPreserveOnlyRecordWorkbookStream() {
                using var stream = new MemoryStream();
                WriteRecord(stream, 0x0809, new byte[] { 0x00, 0x06, 0x05, 0x00, 0xdb, 0x0b, 0xcc, 0x07 });
                long boundSheetPosition = stream.Position;
                WriteRecord(stream, 0x0085, BuildBoundSheetPayload(0, "CorpusGaps"));
                WriteRecord(stream, 0x000a, Array.Empty<byte>());

                int sheetOffset = checked((int)stream.Position);
                WriteRecord(stream, 0x0809, new byte[] { 0x00, 0x06, 0x10, 0x00, 0xdb, 0x0b, 0xcc, 0x07 });
                WriteRecord(stream, 0x0204, BuildLabelPayload(0, 0, "Preserve"));
                WriteRecord(stream, 0x0221, new byte[16]);
                WriteRecord(stream, 0x0063, new byte[] { 0x01, 0x00 });
                WriteRecord(stream, 0x00dd, new byte[] { 0x01, 0x00 });
                WriteRecord(stream, 0x0866, BuildHeaderFooterPicturePayload());
                WriteRecord(stream, 0x000a, Array.Empty<byte>());

                byte[] bytes = stream.ToArray();
                Buffer.BlockCopy(BitConverter.GetBytes(sheetOffset), 0, bytes, checked((int)boundSheetPosition + 4), 4);
                return bytes;
            }

            internal static byte[] CreateHeaderFooterPictureImageWorkbookStream() {
                using var stream = new MemoryStream();
                WriteRecord(stream, 0x0809, new byte[] { 0x00, 0x06, 0x05, 0x00, 0xdb, 0x0b, 0xcc, 0x07 });
                long boundSheetPosition = stream.Position;
                WriteRecord(stream, 0x0085, BuildBoundSheetPayload(0, "HFPicture"));
                WriteRecord(stream, 0x00eb, BuildDrawingGroupWithPngBlipStorePayload());
                WriteRecord(stream, 0x000a, Array.Empty<byte>());

                int sheetOffset = checked((int)stream.Position);
                WriteRecord(stream, 0x0809, new byte[] { 0x00, 0x06, 0x10, 0x00, 0xdb, 0x0b, 0xcc, 0x07 });
                WriteRecord(stream, 0x0204, BuildLabelPayload(0, 0, "Imported"));
                WriteRecord(stream, 0x0014, BuildUnicodeStringPayload("&C&G"));
                WriteRecord(stream, 0x0866, BuildHeaderFooterPicturePayload());
                WriteRecord(stream, 0x000a, Array.Empty<byte>());

                byte[] bytes = stream.ToArray();
                Buffer.BlockCopy(BitConverter.GetBytes(sheetOffset), 0, bytes, checked((int)boundSheetPosition + 4), 4);
                return bytes;
            }

            internal static byte[] CreateEmptyHeaderFooterPictureWorkbookStream() {
                using var stream = new MemoryStream();
                WriteRecord(stream, 0x0809, new byte[] { 0x00, 0x06, 0x05, 0x00, 0xdb, 0x0b, 0xcc, 0x07 });
                long boundSheetPosition = stream.Position;
                WriteRecord(stream, 0x0085, BuildBoundSheetPayload(0, "HFPicture"));
                WriteRecord(stream, 0x000a, Array.Empty<byte>());

                int sheetOffset = checked((int)stream.Position);
                WriteRecord(stream, 0x0809, new byte[] { 0x00, 0x06, 0x10, 0x00, 0xdb, 0x0b, 0xcc, 0x07 });
                WriteRecord(stream, 0x0204, BuildLabelPayload(0, 0, "Preserve"));
                WriteRecord(stream, 0x0866, new byte[14]);
                WriteRecord(stream, 0x000a, Array.Empty<byte>());

                byte[] bytes = stream.ToArray();
                Buffer.BlockCopy(BitConverter.GetBytes(sheetOffset), 0, bytes, checked((int)boundSheetPosition + 4), 4);
                return bytes;
            }

            internal static byte[] CreateUnsupportedChartSheetWorkbookStream() {
                using var stream = new MemoryStream();
                WriteRecord(stream, 0x0809, new byte[] { 0x00, 0x06, 0x05, 0x00, 0xdb, 0x0b, 0xcc, 0x07 });
                long chartBoundSheetPosition = stream.Position;
                WriteRecord(stream, 0x0085, BuildBoundSheetPayload(0, "Chart1", sheetType: 2));
                WriteRecord(stream, 0x000a, Array.Empty<byte>());

                int chartSheetOffset = checked((int)stream.Position);
                WriteRecord(stream, 0x0809, new byte[] { 0x00, 0x06, 0x20, 0x00, 0xdb, 0x0b, 0xcc, 0x07 });
                WriteRecord(stream, 0x0033, new byte[] { 0x02, 0x00 });
                WriteRecord(stream, 0x01b6, new byte[18]);
                WriteRecord(stream, 0x000a, Array.Empty<byte>());

                byte[] bytes = stream.ToArray();
                Buffer.BlockCopy(BitConverter.GetBytes(chartSheetOffset), 0, bytes, checked((int)chartBoundSheetPosition + 4), 4);
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

            internal static byte[] CreateSplitPaneWorkbookStream() {
                using var stream = new MemoryStream();
                WriteRecord(stream, 0x0809, new byte[] { 0x00, 0x06, 0x05, 0x00, 0xdb, 0x0b, 0xcc, 0x07 });
                long splitBoundSheetPosition = stream.Position;
                WriteRecord(stream, 0x0085, BuildBoundSheetPayload(0, "SplitPane"));
                WriteRecord(stream, 0x000a, Array.Empty<byte>());

                int splitSheetOffset = checked((int)stream.Position);
                WriteRecord(stream, 0x0809, new byte[] { 0x00, 0x06, 0x10, 0x00, 0xdb, 0x0b, 0xcc, 0x07 });
                WriteRecord(stream, 0x0204, BuildLabelPayload(0, 0, "Split"));
                WriteRecord(stream, 0x023e, BuildWindow2Payload(frozen: false));
                WriteRecord(stream, 0x0041, BuildPanePayload(leftColumns: 1200, topRows: 900));
                WriteRecord(stream, 0x000a, Array.Empty<byte>());

                byte[] bytes = stream.ToArray();
                Buffer.BlockCopy(BitConverter.GetBytes(splitSheetOffset), 0, bytes, checked((int)splitBoundSheetPosition + 4), 4);
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
                WriteRecord(stream, 0x023e, BuildWindow2Payload(frozen: true, frozenWithoutSplit: true, showFormulas: true, showGridlines: false, showRowColumnHeadings: false, showZeroValues: false, defaultGridColor: false, gridLineColorIndex: 22, rightToLeft: true, showOutlineSymbols: false, tabSelected: true, pageBreakPreview: true, firstVisibleRow: 4, firstVisibleColumn: 2, pageBreakPreviewZoom: 120, normalZoom: 90));
                WriteRecord(stream, 0x0041, BuildPanePayload(leftColumns: 1, topRows: 2));
                WriteRecord(stream, 0x001d, BuildSelectionPayload(pane: 3));
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
                WriteRecord(stream, 0x0031, BuildFontPayload("Consolas", 14d, bold: true, italic: true, underline: true, strikeout: true, colorIndex: 0x0008, escapement: 1, family: 2, characterSet: 238));
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
                WriteRecord(stream, 0x00e0, BuildXfPayload(0, isStyle: true, leftBorderStyle: 1, leftBorderColorIndex: 0x000a));
                WriteRecord(stream, 0x00e0, BuildXfPayload(0, parentStyleIndex: 11, applyBorder: true));
                WriteRecord(stream, 0x0293, BuildStylePayload(0, builtInStyleId: 0));
                WriteRecord(stream, 0x0293, BuildStylePayload(4, name: "OfficeIMO Accent"));
                WriteRecord(stream, 0x087c, BuildXfCrcPayload(xfRecordCount: 13, checksum: 0x12345678U));
                WriteRecord(stream, 0x087d, BuildXfExtPayload(
                    formatIndex: 4,
                    BuildXfExtProperty(0x000e, 0x0001),
                    BuildXfExtByteProperty(0x000e, 0x02),
                    BuildXfExtProperty(0x000f, 0x0003),
                    BuildXfExtFullColorProperty(0x000d, colorType: 0x0003, tintShade: 0, colorValue: 0x00000000),
                    BuildXfExtFullColorProperty(0x0004, colorType: 0x0003, tintShade: 13106, colorValue: 0x00000004)));
                WriteRecord(stream, 0x0892, BuildStyleExtPayload(
                    "OfficeIMO Accent",
                    BuildStyleXfPropByteProperty(0x0000, 0x01),
                    BuildStyleXfPropColorProperty(0x0001, colorType: 0x0003, indexedOrThemeValue: 0x04, tintShade: 13106, colorValue: 0),
                    BuildStyleXfPropColorProperty(0x0002, colorType: 0x0002, indexedOrThemeValue: 0xff, tintShade: 0, colorValue: 0xff00aa00),
                    BuildStyleXfPropColorProperty(0x0005, colorType: 0x0003, indexedOrThemeValue: 0x00, tintShade: 0, colorValue: 0),
                    BuildStyleXfPropBorderProperty(0x0006, colorType: 0x0002, indexedOrThemeValue: 0xff, tintShade: 0, colorValue: 0xff0066ff, borderStyle: 1),
                    BuildStyleXfPropByteProperty(0x0025, 0x02)));
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
                WriteRecord(stream, 0x0204, BuildLabelPayload(0, 11, "Clear border", styleIndex: 12));
                WriteRecord(stream, 0x007d, BuildColInfoPayload(10, 10, 9.5d, hidden: false, styleIndex: 4));
                WriteRecord(stream, 0x007d, BuildColInfoPayload(11, 11, 9.5d, hidden: false, styleIndex: 10));
                WriteRecord(stream, 0x0208, BuildRowPayload(2, 18d, hidden: false, customHeight: false, styleIndex: 5));
                WriteRecord(stream, 0x0208, BuildRowPayload(3, 18d, hidden: false, customHeight: false, styleIndex: 10));
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
                WriteRecord(stream, 0x0006, BuildFormulaSpecialPayload(0, 4, valueType: 0x00, value: 0));
                (byte[] stringPayload, byte[] continuePayload) = BuildContinuedFormulaStringPayload("Continued formula text", firstCharacterCount: 10);
                WriteRecord(stream, 0x0207, stringPayload);
                WriteRecord(stream, 0x003c, continuePayload);
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
                bool? applyFill = null,
                bool? applyBorder = null,
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

                if (applyFill ?? fillPattern != 0) {
                    attributes |= 0x4000;
                }

                if (applyBorder ?? (leftBorderStyle != 0 || rightBorderStyle != 0 || topBorderStyle != 0 || bottomBorderStyle != 0 || diagonalBorderStyle != 0)) {
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

            private static byte[] BuildStylePayload(ushort styleFormatIndex, byte? builtInStyleId = null, byte outlineLevel = 0, string? name = null) {
                using var stream = new MemoryStream();
                ushort flags = (ushort)(styleFormatIndex & 0x0fff);
                if (builtInStyleId.HasValue) {
                    flags |= 0x8000;
                }

                WriteUInt16(stream, flags);
                if (builtInStyleId.HasValue) {
                    stream.WriteByte(builtInStyleId.Value);
                    stream.WriteByte(outlineLevel);
                    return stream.ToArray();
                }

                string styleName = name ?? string.Empty;
                byte[] nameBytes = Encoding.ASCII.GetBytes(styleName);
                WriteUInt16(stream, checked((ushort)styleName.Length));
                stream.WriteByte(0);
                stream.Write(nameBytes, 0, nameBytes.Length);
                return stream.ToArray();
            }

            private static byte[] BuildWorksheetTxoPayload(int textLength, ushort formattingRunBytes) {
                byte[] payload = new byte[16];
                WriteUInt16(payload, 0, 0);
                WriteUInt16(payload, 2, 0);
                WriteUInt16(payload, 10, checked((ushort)textLength));
                WriteUInt16(payload, 12, formattingRunBytes);
                WriteUInt16(payload, 14, 0);
                return payload;
            }

            private static byte[] BuildCompressedTextContinuePayload(string text) {
                if (text == null) {
                    throw new ArgumentNullException(nameof(text));
                }

                byte[] payload = new byte[checked(text.Length + 1)];
                payload[0] = 0;
                for (int i = 0; i < text.Length; i++) {
                    payload[i + 1] = checked((byte)text[i]);
                }

                return payload;
            }

            private static byte[] BuildTxoFormattingRunContinuePayload(ushort startCharacter, ushort fontIndex) {
                byte[] payload = new byte[16];
                WriteUInt16(payload, 0, startCharacter);
                WriteUInt16(payload, 2, fontIndex);
                WriteUInt16(payload, 8, 0xffff);
                return payload;
            }

            private static byte[] BuildWorksheetObjectPayload() {
                using var stream = new MemoryStream();
                WriteUInt16(stream, 0x0015);
                WriteUInt16(stream, 18);
                WriteUInt16(stream, (ushort)LegacyXlsDrawingObjectType.Picture);
                WriteUInt16(stream, 1);
                WriteUInt16(stream, 0x4011);
                stream.Write(new byte[12], 0, 12);
                WriteUInt16(stream, 0x0000);
                WriteUInt16(stream, 0);
                return stream.ToArray();
            }

            private static byte[] BuildWorksheetListObjectPayload() {
                using var stream = new MemoryStream();
                WriteUInt16(stream, 0x0015);
                WriteUInt16(stream, 18);
                WriteUInt16(stream, (ushort)LegacyXlsDrawingObjectType.DropdownList);
                WriteUInt16(stream, 1);
                WriteUInt16(stream, 0x2101);
                stream.Write(new byte[12], 0, 12);

                WriteUInt16(stream, 0x000c);
                WriteUInt16(stream, 20);
                stream.Write(new byte[20], 0, 20);

                WriteUInt16(stream, 0x0013);
                WriteUInt16(stream, 8174);
                stream.Write(new byte[20], 0, 20);
                return stream.ToArray();
            }

            private static byte[] BuildXfCrcPayload(ushort xfRecordCount, uint checksum) {
                using var stream = new MemoryStream();
                WriteUInt16(stream, 0x087c);
                WriteUInt16(stream, 0);
                WriteUInt32(stream, 0);
                WriteUInt32(stream, 0);
                WriteUInt16(stream, 0);
                WriteUInt16(stream, xfRecordCount);
                WriteUInt32(stream, checksum);
                return stream.ToArray();
            }

            private static byte[] BuildXfExtPayload(ushort formatIndex, params byte[][] properties) {
                using var stream = new MemoryStream();
                WriteUInt16(stream, 0x087d);
                WriteUInt16(stream, 0);
                WriteUInt32(stream, 0);
                WriteUInt32(stream, 0);
                WriteUInt16(stream, 0);
                WriteUInt16(stream, formatIndex);
                WriteUInt16(stream, 0);
                WriteUInt16(stream, checked((ushort)properties.Length));
                foreach (byte[] property in properties) {
                    stream.Write(property, 0, property.Length);
                }
                return stream.ToArray();
            }

            private static byte[] BuildXfExtProperty(ushort propertyType, ushort value) {
                using var stream = new MemoryStream();
                WriteUInt16(stream, propertyType);
                WriteUInt16(stream, 6);
                WriteUInt16(stream, value);
                return stream.ToArray();
            }

            private static byte[] BuildXfExtByteProperty(ushort propertyType, byte value) {
                using var stream = new MemoryStream();
                WriteUInt16(stream, propertyType);
                WriteUInt16(stream, 5);
                stream.WriteByte(value);
                return stream.ToArray();
            }

            private static byte[] BuildXfExtFullColorProperty(ushort propertyType, ushort colorType, short tintShade, uint colorValue) {
                using var stream = new MemoryStream();
                WriteUInt16(stream, propertyType);
                WriteUInt16(stream, 20);
                WriteUInt16(stream, colorType);
                WriteUInt16(stream, unchecked((ushort)tintShade));
                WriteUInt32(stream, colorValue);
                stream.Write(new byte[8], 0, 8);
                return stream.ToArray();
            }

            private static byte[] BuildStyleExtPayload(string name, params byte[][] properties) {
                using var stream = new MemoryStream();
                WriteUInt16(stream, 0x0892);
                WriteUInt16(stream, 0);
                WriteUInt32(stream, 0);
                WriteUInt32(stream, 0);
                stream.WriteByte(0);
                stream.WriteByte(0);
                WriteUInt16(stream, 0xffff);
                byte[] nameBytes = Encoding.Unicode.GetBytes(name);
                WriteUInt16(stream, checked((ushort)name.Length));
                stream.Write(nameBytes, 0, nameBytes.Length);
                if (properties.Length > 0) {
                    WriteUInt16(stream, 0);
                    WriteUInt16(stream, checked((ushort)properties.Length));
                    foreach (byte[] property in properties) {
                        stream.Write(property, 0, property.Length);
                    }
                }

                return stream.ToArray();
            }

            private static byte[] BuildStyleXfPropByteProperty(ushort propertyType, byte value) {
                using var stream = new MemoryStream();
                WriteUInt16(stream, propertyType);
                WriteUInt16(stream, 5);
                stream.WriteByte(value);
                return stream.ToArray();
            }

            private static byte[] BuildStyleXfPropColorProperty(ushort propertyType, byte colorType, byte indexedOrThemeValue, short tintShade, uint colorValue) {
                using var stream = new MemoryStream();
                WriteUInt16(stream, propertyType);
                WriteUInt16(stream, 12);
                stream.WriteByte(checked((byte)((colorType << 1) | 0x01)));
                stream.WriteByte(indexedOrThemeValue);
                WriteUInt16(stream, unchecked((ushort)tintShade));
                WriteUInt32(stream, colorValue);
                return stream.ToArray();
            }

            private static byte[] BuildStyleXfPropBorderProperty(ushort propertyType, byte colorType, byte indexedOrThemeValue, short tintShade, uint colorValue, ushort borderStyle) {
                using var stream = new MemoryStream();
                WriteUInt16(stream, propertyType);
                WriteUInt16(stream, 14);
                stream.WriteByte(checked((byte)((colorType << 1) | 0x01)));
                stream.WriteByte(indexedOrThemeValue);
                WriteUInt16(stream, unchecked((ushort)tintShade));
                WriteUInt32(stream, colorValue);
                WriteUInt16(stream, borderStyle);
                return stream.ToArray();
            }

            private static byte[] BuildThemePayload(uint themeVersion, byte[]? themeBytes = null) {
                using var stream = new MemoryStream();
                WriteUInt16(stream, 0x0896);
                WriteUInt16(stream, 0);
                WriteUInt32(stream, 0);
                WriteUInt32(stream, 0);
                WriteUInt32(stream, themeVersion);
                byte[] payload = themeBytes ?? Array.Empty<byte>();
                stream.Write(payload, 0, payload.Length);
                return stream.ToArray();
            }

            private static byte[] BuildThemePackage(string themeName) {
                using var workbookStream = new MemoryStream();
                using var document = OfficeIMO.Excel.ExcelDocument.Create(workbookStream);
                document.ResetWorkbookTheme(themeName);
                string themeXml = document.GetWorkbookThemeXml()!;

                using var packageStream = new MemoryStream();
                using (var archive = new System.IO.Compression.ZipArchive(packageStream, System.IO.Compression.ZipArchiveMode.Create, leaveOpen: true)) {
                    System.IO.Compression.ZipArchiveEntry entry = archive.CreateEntry("theme/theme/theme1.xml");
                    using Stream entryStream = entry.Open();
                    using var writer = new StreamWriter(entryStream, Encoding.UTF8);
                    writer.Write(themeXml);
                }

                return packageStream.ToArray();
            }

            private static byte[] BuildFutureRecordPayload(ushort recordType, int payloadLength) {
                byte[] payload = new byte[payloadLength];
                if (payloadLength >= 2) {
                    WriteUInt16(payload, 0, recordType);
                }

                return payload;
            }

            private static byte[] BuildExternalQueryConnectionPayload() {
                using var stream = new MemoryStream();
                WriteUInt16(stream, 0x0803);
                WriteUInt16(stream, 0);
                WriteUInt16(stream, 0x0005);
                stream.WriteByte(0x83);
                stream.WriteByte(0);
                WriteUInt16(stream, 0x010a);
                WriteUInt16(stream, 0x0003);
                stream.WriteByte(3);
                stream.WriteByte(2);
                stream.WriteByte(1);
                stream.WriteByte(0);
                WriteUInt16(stream, 0);
                WriteUInt16(stream, 1);
                WriteUInt16(stream, 2);
                WriteUInt16(stream, 15);
                WriteUInt16(stream, 3);
                WriteUInt16(stream, 1);
                WriteUInt16(stream, 0x000a);
                stream.WriteByte(0xaa);
                stream.WriteByte(0xbb);
                return stream.ToArray();
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

            private static (byte[] StringPayload, byte[] ContinuePayload) BuildContinuedFormulaStringPayload(string text, int firstCharacterCount) {
                byte[] textBytes = Encoding.ASCII.GetBytes(text);
                if (firstCharacterCount <= 0 || firstCharacterCount >= textBytes.Length) {
                    throw new ArgumentOutOfRangeException(nameof(firstCharacterCount));
                }

                using var first = new MemoryStream();
                WriteUInt16(first, checked((ushort)text.Length));
                first.WriteByte(0);
                first.Write(textBytes, 0, firstCharacterCount);

                using var continued = new MemoryStream();
                continued.WriteByte(0);
                continued.Write(textBytes, firstCharacterCount, textBytes.Length - firstCharacterCount);
                return (first.ToArray(), continued.ToArray());
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

            private static byte[] BuildFontPayload(string name, double size, bool bold, bool italic, bool underline, bool strikeout = false, ushort colorIndex = 0x7FFF, ushort escapement = 0, byte family = 2, byte characterSet = 0) {
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
                stream.WriteByte(family);
                stream.WriteByte(characterSet);
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

            private static byte[] BuildWindow2Payload(bool frozen, bool frozenWithoutSplit = false, bool showFormulas = false, bool showGridlines = true, bool showRowColumnHeadings = true, bool showZeroValues = true, bool defaultGridColor = true, ushort gridLineColorIndex = 64, bool rightToLeft = false, bool showOutlineSymbols = true, bool tabSelected = false, bool pageBreakPreview = false, ushort firstVisibleRow = 0, ushort firstVisibleColumn = 0, ushort pageBreakPreviewZoom = 0, ushort normalZoom = 0) {
                using var stream = new MemoryStream();
                ushort options = 0;
                if (showFormulas) {
                    options |= 0x0001;
                }

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

                if (defaultGridColor) {
                    options |= 0x0020;
                }

                if (rightToLeft) {
                    options |= 0x0040;
                }

                if (showOutlineSymbols) {
                    options |= 0x0080;
                }

                if (frozen && frozenWithoutSplit) {
                    options |= 0x0100;
                }

                if (tabSelected) {
                    options |= 0x0200;
                }

                if (pageBreakPreview) {
                    options |= 0x0800;
                }

                WriteUInt16(stream, options);
                WriteUInt16(stream, firstVisibleRow);
                WriteUInt16(stream, firstVisibleColumn);
                WriteUInt16(stream, gridLineColorIndex);
                WriteUInt16(stream, 0);
                WriteUInt16(stream, pageBreakPreviewZoom);
                WriteUInt16(stream, normalZoom);
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

        private static void WriteInt32(Stream stream, int value) {
            WriteUInt32(stream, unchecked((uint)value));
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
