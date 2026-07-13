using System.Globalization;
using System.Runtime.InteropServices;
#if NET8_0_OR_GREATER
using System.Runtime.Versioning;
#endif
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using OfficeIMO.Excel;
using Xunit;

namespace OfficeIMO.Tests {
    public partial class Excel {
        [Fact]
        public void Test_AutoFitVisualMatrix_MatchesExcelReferenceWhenEnabled() {
            if (!AutoFitVisualValidationEnabled()) {
                return;
            }

#if NET8_0_OR_GREATER
            if (!OperatingSystem.IsWindows()) {
                return;
            }
#else
            if (!RuntimeInformation.IsOSPlatform(OSPlatform.Windows)) {
                return;
            }
#endif

            if (Type.GetTypeFromProgID("Excel.Application") == null) {
                return;
            }

            string outputDirectory = Path.Combine(_directoryWithFiles, "AutoFitVisual");
            Directory.CreateDirectory(outputDirectory);

            string officeImoPath = Path.Combine(outputDirectory, "OfficeIMO.AutoFit.Visual.xlsx");
            string excelReferencePath = Path.Combine(outputDirectory, "Excel.AutoFit.Reference.xlsx");
            string reportPath = Path.Combine(outputDirectory, "AutoFit.Visual.Report.md");
            string csvPath = Path.Combine(outputDirectory, "AutoFit.Visual.Metrics.csv");

            CreateAutoFitVisualWorkbook(officeImoPath, applyOfficeImoAutoFit: true);
            CreateAutoFitVisualWorkbook(excelReferencePath, applyOfficeImoAutoFit: false);

            var excel = new ExcelAutoFitAutomation();
            try {
                var officeImoMetrics = excel.CaptureMetrics(officeImoPath, applyExcelAutoFit: false, exportPdf: true);
                var referenceMetrics = excel.CaptureMetrics(excelReferencePath, applyExcelAutoFit: true, exportPdf: true);

                var failures = CompareAutoFitMetrics(officeImoMetrics, referenceMetrics);
                WriteAutoFitVisualReport(reportPath, csvPath, officeImoMetrics, referenceMetrics, failures);

                Assert.True(
                    failures.Count == 0,
                    "AutoFit visual validation found differences. See " + reportPath + Environment.NewLine + string.Join(Environment.NewLine, failures));
            } finally {
                excel.Dispose();
            }
        }

        private static bool AutoFitVisualValidationEnabled() {
            string? value = Environment.GetEnvironmentVariable("OFFICEIMO_EXCEL_AUTOFIT_VISUAL");
            return string.Equals(value, "1", StringComparison.Ordinal)
                || string.Equals(value, "true", StringComparison.OrdinalIgnoreCase);
        }

        private static void CreateAutoFitVisualWorkbook(string filePath, bool applyOfficeImoAutoFit) {
            using var document = ExcelDocument.Create(filePath);
            var sheet = document.AddWorksheet("AutoFit Visual");

            sheet.CellValue(1, 1, "Case");
            sheet.CellValue(1, 2, "Auto column text");
            sheet.CellValue(1, 3, "Fixed wrapped text");
            sheet.CellBold(1, 1, true);
            sheet.CellBold(1, 2, true);
            sheet.CellBold(1, 3, true);

            var cases = new[] {
                new AutoFitVisualCase("Plain short", "Plain text", "Plain wrapped text", "Calibri", 11.0),
                new AutoFitVisualCase("Plain long", "The quick brown fox jumps over the lazy dog", "The quick brown fox jumps over the lazy dog several times to force wrapping.", "Calibri", 11.0),
                new AutoFitVisualCase("Numeric", "1234567890.12345", "1234567890.12345 9876543210.54321", "Calibri", 11.0),
                new AutoFitVisualCase("Punctuation", "!@#$%^&*()[]{} /\\ <> += - _", "!@#$%^&*()[]{} /\\ <> += - _ with wrapping", "Calibri", 11.0),
                new AutoFitVisualCase("Aptos 8", "Small Aptos text 12345", "Small Aptos text with two wrapped visual lines 12345", "Aptos", 8.0),
                new AutoFitVisualCase("Aptos 14", "Medium Aptos text 12345", "Medium Aptos text with two wrapped visual lines 12345", "Aptos", 14.0),
                new AutoFitVisualCase("Aptos 20", "Large Aptos text 12345", "Large Aptos text with explicit\nline break", "Aptos", 20.0),
                new AutoFitVisualCase("Arial bold", "Bold Arial width check", "Bold Arial wrapped height check with enough words", "Arial", 12.0, bold: true),
                new AutoFitVisualCase("Arial italic", "Italic Arial width check", "Italic Arial wrapped height check with enough words", "Arial", 12.0, italic: true),
                new AutoFitVisualCase("Arial bold italic", "Bold italic Arial width check", "Bold italic Arial wrapped height check with enough words", "Arial", 12.0, bold: true, italic: true),
                new AutoFitVisualCase("Arial underline", "Underlined Arial text", "Underlined Arial wrapped text with descenders: gyjpq", "Arial", 12.0, underline: true),
                new AutoFitVisualCase("Times", "Times New Roman text 12345", "Times New Roman wrapped paragraph with descenders gyjpq", "Times New Roman", 12.0),
                new AutoFitVisualCase("Consolas", "Consolas monospace 12345", "Consolas monospace wrapped text 12345 67890", "Consolas", 11.0),
                new AutoFitVisualCase("Courier", "Courier New monospace 12345", "Courier New monospace wrapped text 12345 67890", "Courier New", 11.0),
                new AutoFitVisualCase("CJK", "OfficeIMO 日本語 測試 한글 123", "OfficeIMO 日本語 測試 한글 wrapping with Latin words", "Calibri", 11.0),
                new AutoFitVisualCase("Tabs", "Text\twith\ttabs", "Text\twith\ttabs and wrapping words after tabs", "Calibri", 11.0),
                new AutoFitVisualCase("Multiline", "Top\nBottom", "Top\nMiddle\nBottom", "Calibri", 11.0),
                new AutoFitVisualCase("Large underline", "Large underlined descenders gyjpq", "Large underlined wrapped descenders gyjpq gyjpq", "Arial", 20.0, underline: true),
                new AutoFitVisualCase("Currency format", 1234.5D, "Currency formatted wrapped text", "Calibri", 11.0, numberFormat: "$#,##0.00"),
                new AutoFitVisualCase("Percent format", 1.0D, "Percent formatted wrapped text", "Calibri", 11.0, numberFormat: "0.00%"),
                new AutoFitVisualCase("Date format", new DateTime(2026, 5, 7), "Date formatted wrapped text", "Calibri", 11.0, numberFormat: "yyyy-mm-dd")
            };

            for (int i = 0; i < cases.Length; i++) {
                int row = i + 2;
                AutoFitVisualCase current = cases[i];
                sheet.CellValue(row, 1, current.Name);
                if (!string.IsNullOrEmpty(current.NumberFormat)) {
                    sheet.Cell(row, 2, current.ColumnValue, numberFormat: current.NumberFormat);
                } else {
                    sheet.CellValue(row, 2, current.ColumnValue);
                }
                sheet.CellValue(row, 3, current.WrappedText);
            }

            SpreadsheetDocument spreadsheet = document._spreadSheetDocument;
            for (int i = 0; i < cases.Length; i++) {
                int row = i + 2;
                AutoFitVisualCase current = cases[i];
                uint style = AddAutoFitVisualStyle(
                    spreadsheet,
                    current.FontName,
                    current.FontSize,
                    current.Bold,
                    current.Italic,
                    current.Underline,
                    wrapText: false,
                    numberFormat: current.NumberFormat);
                uint wrapStyle = AddAutoFitVisualStyle(
                    spreadsheet,
                    current.FontName,
                    current.FontSize,
                    current.Bold,
                    current.Italic,
                    current.Underline,
                    wrapText: true);

                SetCellStyle(spreadsheet, "B" + row.ToString(CultureInfo.InvariantCulture), style);
                SetCellStyle(spreadsheet, "C" + row.ToString(CultureInfo.InvariantCulture), wrapStyle);
            }

            sheet.SetColumnWidth(3, 18);
            if (applyOfficeImoAutoFit) {
                sheet.AutoFitColumnsFor(new[] { 1, 2 });
                sheet.AutoFitRows();
            }

            document.Save();
        }

        private static uint AddAutoFitVisualStyle(
            SpreadsheetDocument document,
            string fontName,
            double fontSize,
            bool bold,
            bool italic,
            bool underline,
            bool wrapText,
            string? numberFormat = null) {
            var stylesPart = document.WorkbookPart!.WorkbookStylesPart ?? document.WorkbookPart!.AddNewPart<WorkbookStylesPart>();
            if (stylesPart.Stylesheet == null) {
                stylesPart.Stylesheet = new Stylesheet(
                    new Fonts(new DocumentFormat.OpenXml.Spreadsheet.Font()),
                    new Fills(new Fill()),
                    new Borders(new Border()),
                    new CellFormats(new CellFormat()));
                stylesPart.Stylesheet.Fonts!.Count = 1;
                stylesPart.Stylesheet.Fills!.Count = 1;
                stylesPart.Stylesheet.Borders!.Count = 1;
                stylesPart.Stylesheet.CellFormats!.Count = 1;
            }

            var stylesheet = stylesPart.Stylesheet!;
            var font = new DocumentFormat.OpenXml.Spreadsheet.Font(
                new FontName { Val = fontName },
                new FontSize { Val = fontSize });
            if (bold) font.Append(new Bold());
            if (italic) font.Append(new Italic());
            if (underline) font.Append(new Underline());

            stylesheet.Fonts!.Append(font);
            stylesheet.Fonts.Count = (uint)stylesheet.Fonts.ChildElements.Count;

            uint? numberFormatId = null;
            if (!string.IsNullOrWhiteSpace(numberFormat)) {
                stylesheet.NumberingFormats ??= new NumberingFormats();
                var existing = stylesheet.NumberingFormats.Elements<NumberingFormat>()
                    .FirstOrDefault(format => string.Equals(format.FormatCode?.Value, numberFormat, StringComparison.Ordinal));
                if (existing?.NumberFormatId?.Value is uint existingId) {
                    numberFormatId = existingId;
                } else {
                    uint nextId = stylesheet.NumberingFormats.Elements<NumberingFormat>().Any()
                        ? Math.Max(164U, stylesheet.NumberingFormats.Elements<NumberingFormat>().Max(format => format.NumberFormatId?.Value ?? 0U) + 1U)
                        : 164U;
                    stylesheet.NumberingFormats.Append(new NumberingFormat { NumberFormatId = nextId, FormatCode = numberFormat });
                    stylesheet.NumberingFormats.Count = (uint)stylesheet.NumberingFormats.ChildElements.Count;
                    numberFormatId = nextId;
                }
            }

            var format = new CellFormat {
                FontId = stylesheet.Fonts.Count - 1,
                ApplyFont = true
            };
            if (numberFormatId.HasValue) {
                format.NumberFormatId = numberFormatId.Value;
                format.ApplyNumberFormat = true;
            }

            if (wrapText) {
                format.Alignment = new Alignment { WrapText = true };
                format.ApplyAlignment = true;
            }

            stylesheet.CellFormats!.Append(format);
            stylesheet.CellFormats.Count = (uint)stylesheet.CellFormats.ChildElements.Count;
            stylesPart.Stylesheet!.Save();
            return stylesheet.CellFormats.Count - 1;
        }

        private static void SetCellStyle(SpreadsheetDocument document, string cellReference, uint styleIndex) {
            var worksheetPart = document.WorkbookPart!.WorksheetParts.First();
            var cell = worksheetPart.Worksheet.Descendants<Cell>().First(c => c.CellReference == cellReference);
            cell.StyleIndex = styleIndex;
            worksheetPart.Worksheet.Save();
        }

        private static List<string> CompareAutoFitMetrics(AutoFitMetrics officeImo, AutoFitMetrics reference) {
            var failures = new List<string>();

            foreach (var expected in reference.Columns) {
                if (!officeImo.Columns.TryGetValue(expected.Key, out double actual)) {
                    failures.Add($"Missing OfficeIMO column metric {expected.Key}.");
                    continue;
                }

                if (actual + 1.0 < expected.Value) {
                    failures.Add($"Column {expected.Key} is narrower than Excel reference. OfficeIMO={actual:0.##}, Excel={expected.Value:0.##}.");
                }

                if (actual > Math.Max(expected.Value * 2.5, expected.Value + 12.0)) {
                    failures.Add($"Column {expected.Key} is much wider than Excel reference. OfficeIMO={actual:0.##}, Excel={expected.Value:0.##}.");
                }
            }

            foreach (var expected in reference.Rows) {
                if (!officeImo.Rows.TryGetValue(expected.Key, out double actual)) {
                    failures.Add($"Missing OfficeIMO row metric {expected.Key}.");
                    continue;
                }

                if (actual + 1.5 < expected.Value) {
                    failures.Add($"Row {expected.Key} is shorter than Excel reference. OfficeIMO={actual:0.##}, Excel={expected.Value:0.##}.");
                }

                if (actual > Math.Max(expected.Value * 2.5, expected.Value + 30.0)) {
                    failures.Add($"Row {expected.Key} is much taller than Excel reference. OfficeIMO={actual:0.##}, Excel={expected.Value:0.##}.");
                }
            }

            return failures;
        }

        private static void WriteAutoFitVisualReport(
            string reportPath,
            string csvPath,
            AutoFitMetrics officeImo,
            AutoFitMetrics reference,
            IReadOnlyList<string> failures) {
            var csv = new List<string> {
                "Kind,Index,OfficeIMO,ExcelReference,Delta"
            };

            foreach (int column in reference.Columns.Keys.OrderBy(static key => key)) {
                double actual = officeImo.Columns.TryGetValue(column, out double value) ? value : 0;
                double expected = reference.Columns[column];
                csv.Add(string.Format(CultureInfo.InvariantCulture, "Column,{0},{1:0.####},{2:0.####},{3:0.####}", column, actual, expected, actual - expected));
            }

            foreach (int row in reference.Rows.Keys.OrderBy(static key => key)) {
                double actual = officeImo.Rows.TryGetValue(row, out double value) ? value : 0;
                double expected = reference.Rows[row];
                csv.Add(string.Format(CultureInfo.InvariantCulture, "Row,{0},{1:0.####},{2:0.####},{3:0.####}", row, actual, expected, actual - expected));
            }

            File.WriteAllLines(csvPath, csv);

            var report = new List<string> {
                "# AutoFit Visual Validation",
                string.Empty,
                "Artifacts:",
                "- OfficeIMO workbook: `" + officeImo.WorkbookPath + "`",
                "- OfficeIMO PDF: `" + officeImo.PdfPath + "`",
                "- Excel reference workbook: `" + reference.WorkbookPath + "`",
                "- Excel reference PDF: `" + reference.PdfPath + "`",
                "- Metrics CSV: `" + csvPath + "`",
                string.Empty,
                "Result: " + (failures.Count == 0 ? "PASS" : "FAIL"),
                string.Empty
            };

            if (failures.Count > 0) {
                report.Add("Failures:");
                foreach (string failure in failures) {
                    report.Add("- " + failure);
                }
            }

            File.WriteAllLines(reportPath, report);
        }

        private readonly struct AutoFitVisualCase {
            internal AutoFitVisualCase(
                string name,
                object columnValue,
                string wrappedText,
                string fontName,
                double fontSize,
                bool bold = false,
                bool italic = false,
                bool underline = false,
                string? numberFormat = null) {
                Name = name;
                ColumnValue = columnValue;
                WrappedText = wrappedText;
                FontName = fontName;
                FontSize = fontSize;
                Bold = bold;
                Italic = italic;
                Underline = underline;
                NumberFormat = numberFormat;
            }

            internal string Name { get; }
            internal object ColumnValue { get; }
            internal string WrappedText { get; }
            internal string FontName { get; }
            internal double FontSize { get; }
            internal bool Bold { get; }
            internal bool Italic { get; }
            internal bool Underline { get; }
            internal string? NumberFormat { get; }
        }

        private sealed class AutoFitMetrics {
            internal AutoFitMetrics(string workbookPath, string pdfPath, Dictionary<int, double> columns, Dictionary<int, double> rows) {
                WorkbookPath = workbookPath;
                PdfPath = pdfPath;
                Columns = columns;
                Rows = rows;
            }

            internal string WorkbookPath { get; }
            internal string PdfPath { get; }
            internal Dictionary<int, double> Columns { get; }
            internal Dictionary<int, double> Rows { get; }
        }

#if NET8_0_OR_GREATER
        [SupportedOSPlatform("windows")]
#endif
        private sealed class ExcelAutoFitAutomation : IDisposable {
            private readonly object _application;
            private readonly dynamic _excel;

            internal ExcelAutoFitAutomation() {
                var type = Type.GetTypeFromProgID("Excel.Application") ?? throw new InvalidOperationException("Excel COM is not registered.");
                _application = Activator.CreateInstance(type) ?? throw new InvalidOperationException("Excel COM could not be created.");
                _excel = _application;
                _excel.Visible = false;
                _excel.DisplayAlerts = false;
            }

            internal AutoFitMetrics CaptureMetrics(string workbookPath, bool applyExcelAutoFit, bool exportPdf) {
                object? workbook = null;
                try {
                    workbook = _excel.Workbooks.Open(workbookPath);
                    dynamic dynamicWorkbook = workbook;
                    dynamic sheet = dynamicWorkbook.Worksheets[1];

                    if (applyExcelAutoFit) {
                        sheet.Columns["A:B"].AutoFit();
                        sheet.Rows.AutoFit();
                        dynamicWorkbook.Save();
                    }

                    string pdfPath = Path.ChangeExtension(workbookPath, ".pdf");
                    if (exportPdf) {
                        dynamicWorkbook.ExportAsFixedFormat(0, pdfPath);
                    }

                    var columns = new Dictionary<int, double>();
                    for (int column = 1; column <= 3; column++) {
                        columns[column] = Convert.ToDouble(sheet.Cells[1, column].EntireColumn.ColumnWidth, CultureInfo.InvariantCulture);
                    }

                    int rowCount = Convert.ToInt32(sheet.UsedRange.Rows.Count, CultureInfo.InvariantCulture);
                    var rows = new Dictionary<int, double>();
                    for (int row = 1; row <= rowCount; row++) {
                        rows[row] = Convert.ToDouble(sheet.Cells[row, 1].EntireRow.RowHeight, CultureInfo.InvariantCulture);
                    }

                    dynamicWorkbook.Close(false);
                    ReleaseComObject(workbook);
                    workbook = null;

                    return new AutoFitMetrics(workbookPath, pdfPath, columns, rows);
                } finally {
                    if (workbook != null) {
                        try {
                            dynamic dynamicWorkbook = workbook;
                            dynamicWorkbook.Close(false);
                        } catch {
                            // Best effort cleanup for desktop Excel automation.
                        }

                        ReleaseComObject(workbook);
                    }
                }
            }

            public void Dispose() {
                try {
                    _excel.Quit();
                } catch {
                    // Best effort cleanup for desktop Excel automation.
                }

                ReleaseComObject(_application);
                GC.Collect();
                GC.WaitForPendingFinalizers();
            }

            private static void ReleaseComObject(object value) {
                try {
                    if (Marshal.IsComObject(value)) {
                        Marshal.FinalReleaseComObject(value);
                    }
                } catch {
                    // Best effort cleanup for desktop Excel automation.
                }
            }
        }
    }
}
