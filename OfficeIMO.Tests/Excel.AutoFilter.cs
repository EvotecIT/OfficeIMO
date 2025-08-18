using System.Collections.Generic;
using System.IO;
using System.Linq;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using OfficeIMO.Excel;
using Xunit;

namespace OfficeIMO.Tests {
    /// <summary>
    /// Tests for adding and persisting auto filters in Excel sheets.
    /// </summary>
    public partial class Excel {
        [Fact]
        public void Test_AddAutoFilterPersists() {
            string filePath = Path.Combine(_directoryWithFiles, "AutoFilter.xlsx");
            using (ExcelDocument document = ExcelDocument.Create(filePath)) {
                ExcelSheet sheet = document.AddWorkSheet("Data");
                sheet.SetCellValue(1, 1, "Name");
                sheet.SetCellValue(1, 2, "Value");
                sheet.SetCellValue(2, 1, "A");
                sheet.SetCellValue(2, 2, 10d);
                sheet.SetCellValue(3, 1, "B");
                sheet.SetCellValue(3, 2, 20d);
                Dictionary<uint, IEnumerable<string>> criteria = new Dictionary<uint, IEnumerable<string>> {
                    { 0, new[] { "A" } }
                };
                sheet.AddAutoFilter("A1:B3", criteria);
                document.Save();
            }

            using (ExcelDocument document = ExcelDocument.Load(filePath)) {
                document.Save();
            }

            using (SpreadsheetDocument spreadsheet = SpreadsheetDocument.Open(filePath, false)) {
                WorksheetPart wsPart = spreadsheet.WorkbookPart.WorksheetParts.First();
                AutoFilter autoFilter = wsPart.Worksheet.Elements<AutoFilter>().FirstOrDefault();
                Assert.NotNull(autoFilter);
                Assert.Equal("A1:B3", autoFilter.Reference.Value);
                FilterColumn filterColumn = autoFilter.Elements<FilterColumn>().FirstOrDefault();
                Assert.NotNull(filterColumn);
                Filters filters = filterColumn.GetFirstChild<Filters>();
                Assert.NotNull(filters);
                Filter filter = filters.Elements<Filter>().First();
                Assert.Equal("A", filter.Val.Value);
            }
        }
    }
}
