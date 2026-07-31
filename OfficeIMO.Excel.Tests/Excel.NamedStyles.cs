using System.Linq;
using DocumentFormat.OpenXml.Spreadsheet;
using OfficeIMO.Excel;
using Xunit;

namespace OfficeIMO.Tests {
    public partial class Excel {
        [Fact]
        public void Test_NamedStyles_DefineApplyListAndRemoveCatalogEntry() {
            using var document = ExcelDocument.Create();
            ExcelSheet sheet = document.AddWorksheet("Data");
            sheet.CellValue(1, 1, "Source");
            sheet.CellAt(1, 1).SetBold().SetFillColor("D9EAF7").SetNumberFormat("0.00");

            ExcelNamedStyleInfo defined = sheet.DefineNamedStyle("Input Value", 1, 1);
            sheet.ApplyNamedStyle("Input Value", "B2:C3");

            Assert.Equal("Input Value", defined.Name);
            Assert.Contains(document.GetNamedStyles(), style => style.Name == "Input Value");

            Stylesheet stylesheet = document.OpenXmlDocument.WorkbookPart!.WorkbookStylesPart!.Stylesheet!;
            uint expectedFormat = stylesheet.CellStyles!.Elements<CellStyle>()
                .Single(style => style.Name?.Value == "Input Value").FormatId!.Value;
            Worksheet worksheet = document.OpenXmlDocument.WorkbookPart!.WorksheetParts.Single().Worksheet!;
            foreach (Cell cell in worksheet.Descendants<Cell>().Where(cell =>
                cell.CellReference?.Value is "B2" or "C2" or "B3" or "C3")) {
                CellFormat format = stylesheet.CellFormats!.Elements<CellFormat>().ElementAt((int)cell.StyleIndex!.Value);
                Assert.Equal(expectedFormat, format.FormatId!.Value);
            }

            Assert.True(document.RemoveNamedStyle("Input Value"));
            Assert.DoesNotContain(document.GetNamedStyles(), style => style.Name == "Input Value");
        }

        [Fact]
        public void Test_NamedStyles_EnforcesApplicationBudgetBeforeMutation() {
            using var document = ExcelDocument.Create();
            ExcelSheet sheet = document.AddWorksheet("Data");
            sheet.DefineNamedStyle("Input", 1, 1);

            Assert.Throws<InvalidOperationException>(() =>
                sheet.ApplyNamedStyle("Input", "A1:C3", maximumCells: 8));
            Assert.Empty(document.OpenXmlDocument.WorkbookPart!.WorksheetParts.Single().Worksheet!.Descendants<Cell>());
        }
    }
}
