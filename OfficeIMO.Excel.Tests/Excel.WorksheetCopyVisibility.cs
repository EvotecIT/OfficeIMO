using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using DocumentFormat.OpenXml.Validation;
using OfficeIMO.Excel;
using Xunit;

namespace OfficeIMO.Tests {
    public partial class Excel {
        [Theory]
        [InlineData(ExcelWorksheetCopyMode.Package)]
        [InlineData(ExcelWorksheetCopyMode.Values)]
        public void Test_ExcelWorksheetCopy_CopyWorksheetFromKeepsSourceVisibility(ExcelWorksheetCopyMode copyMode) {
            using ExcelDocument source = LoadVisibilitySource();
            using var target = ExcelDocument.Create();
            target.AddSheet("Summary").Cell(1, 1).SetValue("summary");
            foreach (string name in new[] { "Lookup", "Report", "Secret" }) {
                target.CopyWorksheetFrom(source, name, name, ExcelSheetNameValidationMode.Strict, new ExcelWorksheetCopyOptions { CopyMode = copyMode });
            }

            Assert.True(target["Lookup"].Hidden);
            Assert.False(target["Lookup"].VeryHidden);
            Assert.True(target["Secret"].VeryHidden);
            Assert.False(target["Report"].Hidden);

            AssertSavedVisibility(
                SaveToBytes(target),
                "Summary",
                ("Summary", null),
                ("Lookup", "hidden"),
                ("Report", null),
                ("Secret", "veryHidden"));
        }

        [Theory]
        [InlineData(ExcelWorksheetCopyMode.Package)]
        [InlineData(ExcelWorksheetCopyMode.Values)]
        public void Test_ExcelWorkbookMerge_KeepsSourceVisibilityAndActivatesVisibleSheet(ExcelWorksheetCopyMode copyMode) {
            using ExcelDocument source = LoadVisibilitySource();
            using var target = ExcelDocument.Create();
            target.MergeWorkbookFrom(source, new ExcelWorkbookMergeOptions {
                TargetSheetNames = new Dictionary<string, string> {
                    ["Lookup"] = "b - Lookup",
                    ["Report"] = "b - Report",
                    ["Secret"] = "b - Secret"
                },
                SheetNameValidationMode = ExcelSheetNameValidationMode.Strict,
                CopyMode = copyMode
            });

            // The empty target's active tab defaulted to the first imported sheet, which is hidden.
            Assert.Equal("b - Report", target.ActiveWorksheet?.Name);

            target.MergeWorkbookFrom(source, new ExcelWorkbookMergeOptions {
                SheetNamePrefix = "c - ",
                CopyMode = copyMode
            });

            byte[] saved = SaveToBytes(target);
            AssertSavedVisibility(
                saved,
                "b - Report",
                ("b - Lookup", "hidden"),
                ("b - Report", null),
                ("b - Secret", "veryHidden"),
                ("c - Lookup", "hidden"),
                ("c - Report", null),
                ("c - Secret", "veryHidden"));
            Assert.Equal(new[] { "b - Report" }, GetSavedTabSelectedSheetNames(saved));
        }

        [Fact]
        public void Test_ExcelWorkbookMerge_KeepsSourceVisibilityInWorkbookWithoutStyles() {
            // A workbook with no sheets, styles or theme takes the first source worksheet's part graph as it is.
            using ExcelDocument source = LoadVisibilitySource();
            using ExcelDocument target = LoadWorkbookWithoutSheets();
            target.MergeWorkbookFrom(source, new ExcelWorkbookMergeOptions { CopyMode = ExcelWorksheetCopyMode.Package });

            AssertSavedVisibility(
                SaveToBytes(target),
                "Report",
                ("Lookup", "hidden"),
                ("Report", null),
                ("Secret", "veryHidden"));
        }

        [Fact]
        public void Test_ExcelWorksheetCopy_CopyWorksheetKeepsVisibilityWithinWorkbook() {
            using ExcelDocument workbook = LoadVisibilitySource(readOnly: false);
            workbook.CopyWorksheet("Lookup", "Lookup copy");
            workbook.CopyWorksheet("Secret", "Secret copy");
            workbook.CopyWorksheetFrom(workbook, "Report", "Report copy");

            AssertSavedVisibility(
                SaveToBytes(workbook),
                "Report",
                ("Lookup", "hidden"),
                ("Report", null),
                ("Secret", "veryHidden"),
                ("Lookup copy", "hidden"),
                ("Secret copy", "veryHidden"),
                ("Report copy", null));
        }

        [Fact]
        public void Test_ExcelWorksheetCopy_HiddenCopiesNeverLeaveOnlyHiddenSheets() {
            using ExcelDocument source = LoadVisibilitySource();

            using (var target = ExcelDocument.Create()) {
                target.CopyWorksheetFrom(source, "Lookup", "Lookup", ExcelSheetNameValidationMode.Strict, new ExcelWorksheetCopyOptions { CopyMode = ExcelWorksheetCopyMode.Package });
                Assert.False(target["Lookup"].Hidden);

                target.CopyWorksheetFrom(source, "Secret", "Secret", ExcelSheetNameValidationMode.Strict, new ExcelWorksheetCopyOptions { CopyMode = ExcelWorksheetCopyMode.Package });

                AssertSavedVisibility(
                    SaveToBytes(target),
                    "Lookup",
                    ("Lookup", null),
                    ("Secret", "veryHidden"));
            }

            using (var target = ExcelDocument.Create()) {
                target.MergeWorkbookFrom(source, new ExcelWorkbookMergeOptions {
                    SheetNames = new[] { "Secret", "Lookup" },
                    CopyMode = ExcelWorksheetCopyMode.Package
                });

                AssertSavedVisibility(
                    SaveToBytes(target),
                    "Secret",
                    ("Secret", null),
                    ("Lookup", "hidden"));
            }
        }

        [Fact]
        public void Test_ExcelWorksheetCopy_HiddenCopyOpensOnTheOnlyVisibleChartSheet() {
            using ExcelDocument source = LoadVisibilitySource();
            var stream = new MemoryStream();
            using (SpreadsheetDocument package = SpreadsheetDocument.Create(stream, SpreadsheetDocumentType.Workbook)) {
                WorkbookPart workbookPart = package.AddWorkbookPart();
                WorksheetPart worksheetPart = workbookPart.AddNewPart<WorksheetPart>();
                worksheetPart.Worksheet = new Worksheet(new SheetData());
                ChartsheetPart chartsheetPart = workbookPart.AddNewPart<ChartsheetPart>();
                chartsheetPart.Chartsheet = new Chartsheet();
                // No workbook view, so the hidden worksheet in the first tab is the active tab.
                workbookPart.Workbook = new Workbook(new Sheets(
                    new Sheet { Id = workbookPart.GetIdOfPart(worksheetPart), SheetId = 1U, Name = "Data", State = SheetStateValues.Hidden },
                    new Sheet { Id = workbookPart.GetIdOfPart(chartsheetPart), SheetId = 2U, Name = "Chart" }));
            }

            stream.Position = 0;
            using (ExcelDocument target = ExcelDocument.Load(stream)) {
                target.CopyWorksheetFrom(source, "Lookup", "Lookup", ExcelSheetNameValidationMode.Strict, new ExcelWorksheetCopyOptions { CopyMode = ExcelWorksheetCopyMode.Package });
                Assert.True(target["Lookup"].Hidden);

                using var saved = new MemoryStream(SaveToBytes(target));
                using SpreadsheetDocument spreadsheet = SpreadsheetDocument.Open(saved, false);
                Workbook root = spreadsheet.WorkbookPart!.Workbook;
                List<Sheet> sheets = root.Sheets!.Elements<Sheet>().ToList();
                Assert.Equal(new[] { ("Data", "hidden"), ("Chart", (string?)null), ("Lookup", "hidden") }, sheets.Select(_ => (_.Name!.Value!, _.State?.InnerText)).ToArray());
                Assert.Equal(1U, root.BookViews?.GetFirstChild<WorkbookView>()?.ActiveTab?.Value);
            }
        }

        /// <summary>Loads a workbook whose active "Report" sits between a hidden "Lookup" and a very hidden "Secret".</summary>
        private static ExcelDocument LoadVisibilitySource(bool readOnly = true) {
            byte[] bytes;
            using (var book = ExcelDocument.Create()) {
                book.AddSheet("Lookup").Cell(1, 1).SetValue(1);
                book.AddSheet("Report").Cell(1, 1).SetValue(2);
                book.AddSheet("Secret").Cell(1, 1).SetValue(3);
                book.SetActiveWorksheet("Report");
                book["Lookup"].SetHidden(true);
                book["Secret"].SetVeryHidden(true);
                bytes = SaveToBytes(book);
            }

            if (readOnly) {
                return ExcelDocument.Load(new MemoryStream(bytes), new OfficeIMO.Excel.ExcelLoadOptions { AccessMode = OfficeIMO.DocumentAccessMode.ReadOnly });
            }

            var stream = new MemoryStream();
            stream.Write(bytes, 0, bytes.Length);
            stream.Position = 0;
            return ExcelDocument.Load(stream);
        }

        private static ExcelDocument LoadWorkbookWithoutSheets() {
            var stream = new MemoryStream();
            using (SpreadsheetDocument package = SpreadsheetDocument.Create(stream, SpreadsheetDocumentType.Workbook)) {
                package.AddWorkbookPart().Workbook = new Workbook(new Sheets());
            }

            stream.Position = 0;
            return ExcelDocument.Load(stream);
        }

        private static byte[] SaveToBytes(ExcelDocument document) {
            using var stream = new MemoryStream();
            document.Save(stream);
            return stream.ToArray();
        }

        /// <summary>
        /// Checks the saved workbook is valid, has the expected sheets and workbook states in tab order, and opens on
        /// <paramref name="activeSheet"/>.
        /// </summary>
        private static void AssertSavedVisibility(byte[] workbook, string activeSheet, params (string Name, string? State)[] expected) {
            using var stream = new MemoryStream(workbook);
            using SpreadsheetDocument spreadsheet = SpreadsheetDocument.Open(stream, false);
            Assert.Empty(new OpenXmlValidator().Validate(spreadsheet));

            Workbook root = spreadsheet.WorkbookPart!.Workbook;
            List<Sheet> sheets = root.Sheets!.Elements<Sheet>().ToList();
            Assert.Equal(expected, sheets.Select(_ => (_.Name!.Value!, _.State?.InnerText)).ToArray());

            uint activeTab = root.BookViews?.GetFirstChild<WorkbookView>()?.ActiveTab?.Value ?? 0U;
            Assert.Equal(activeSheet, sheets[(int)activeTab].Name!.Value);
        }

        private static string[] GetSavedTabSelectedSheetNames(byte[] workbook) {
            using var stream = new MemoryStream(workbook);
            using SpreadsheetDocument spreadsheet = SpreadsheetDocument.Open(stream, false);
            return spreadsheet.WorkbookPart!.Workbook.Sheets!.Elements<Sheet>()
                .Where(_ => GetWorksheetPartByNameForOperations(spreadsheet, _.Name!.Value!).Worksheet
                    .Descendants<SheetView>()
                    .Any(_ => _.TabSelected?.Value == true))
                .Select(_ => _.Name!.Value!)
                .ToArray();
        }
    }
}
