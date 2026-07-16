using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using OfficeIMO.Excel;
using Xunit;

namespace OfficeIMO.Tests {
    public partial class Excel {
        [Fact]
        public void GetRichText_ReadsSharedStringRuns() {
            string filePath = Path.Combine(Path.GetTempPath(), Guid.NewGuid().ToString("N") + ".xlsx");

            try {
                using (ExcelDocument document = ExcelDocument.Create(filePath)) {
                    ExcelSheet sheet = document.AddWorksheet("Rich Text");
                    sheet.CellValue(1, 1, "placeholder");
                    document.Save();
                }

                using (SpreadsheetDocument spreadsheet = SpreadsheetDocument.Open(filePath, true)) {
                    WorksheetPart worksheetPart = spreadsheet.WorkbookPart!.WorksheetParts.Single();
                    Cell cell = worksheetPart.Worksheet.Descendants<Cell>().Single(item => item.CellReference?.Value == "A1");
                    int sharedStringIndex = int.Parse(cell.CellValue!.InnerText, System.Globalization.CultureInfo.InvariantCulture);
                    SharedStringTable sharedStringTable = spreadsheet.WorkbookPart.SharedStringTablePart!.SharedStringTable!;
                    SharedStringItem originalItem = sharedStringTable.Elements<SharedStringItem>().ElementAt(sharedStringIndex);

                    sharedStringTable.ReplaceChild(
                        new SharedStringItem(
                            new Run(
                                new RunProperties(new Bold(), new RunFont { Val = "Arial" }),
                                new Text("Shared ")),
                            new Run(
                                new RunProperties(new Italic(), new Underline()),
                                new Text("rich text"))),
                        originalItem);
                    sharedStringTable.Save();
                }

                using ExcelDocument loaded = ExcelDocument.Load(filePath);
                IReadOnlyList<ExcelRichTextRun> runs = loaded.Sheets[0].CellAt(1, 1).GetRichText();

                Assert.Equal(2, runs.Count);
                Assert.Equal("Shared ", runs[0].Text);
                Assert.True(runs[0].Bold);
                Assert.Equal("Arial", runs[0].FontName);
                Assert.Equal("rich text", runs[1].Text);
                Assert.True(runs[1].Italic);
                Assert.True(runs[1].Underline);
            } finally {
                TryDelete(filePath);
            }
        }
    }
}
