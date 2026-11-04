using System;
using System.IO;
using System.Linq;
using DocumentFormat.OpenXml.Spreadsheet;
using OfficeIMO.Excel;
using Xunit;

namespace OfficeIMO.Tests {
    public partial class Excel {
        [Fact]
        public void Test_FormulaSearch_ExpandsThreeDimensionalSheetQualifiers() {
            using var document = ExcelDocument.Create(new MemoryStream());
            document.AddSheet("First Sheet");
            document.AddSheet("Middle Sheet");
            document.AddSheet("Last Sheet");
            ExcelSheet summary = document.AddSheet("Summary");
            summary.Cell(1, 1).SetFormula("SUM('First Sheet':'Last Sheet'!B2)");

            ExcelFormulaCellInfo match = Assert.Single(document.SearchFormulas(
                new ExcelFormulaSearchOptions { Reference = "'Middle Sheet'!B2" }));

            Assert.Equal("Summary", match.SheetName);
            Assert.Empty(document.SearchFormulas(
                new ExcelFormulaSearchOptions { Reference = "Summary!B2" }));
        }

        [Fact]
        public void Test_NamedStyleRedefinition_LeavesAmbiguousSharedConsumersUnchanged() {
            using var document = ExcelDocument.Create(new MemoryStream());
            ExcelSheet first = document.AddSheet("First");
            ExcelSheet second = document.AddSheet("Second");
            first.Cell(1, 1).SetValue("Replacement");
            first.Cell(1, 1).SetFillColor("C6EFCE");
            Stylesheet stylesheet = document.WorkbookPartRoot.WorkbookStylesPart!.Stylesheet!;
            CellStyle normal = stylesheet.CellStyles!.Elements<CellStyle>()
                .Single(style => style.Name?.Value == "Normal");
            uint normalFormatId = normal.FormatId!.Value;
            string normalFormatXml = stylesheet.CellStyleFormats!.Elements<CellFormat>()
                .ElementAt((int)normalFormatId).OuterXml;
            stylesheet.CellStyles.Append(new CellStyle { Name = "Alias", FormatId = normalFormatId });
            stylesheet.CellStyles.Count = (uint)stylesheet.CellStyles.Count();
            first.ApplyNamedStyle("Alias", "B2:B2");
            second.ApplyNamedStyle("Alias", "C3:C3");
            uint oldAppliedIndex = first.WorksheetPart.Worksheet.Descendants<Cell>()
                .Single(cell => cell.CellReference?.Value == "B2").StyleIndex!.Value;

            Assert.Throws<InvalidOperationException>(() => first.DefineNamedStyle("Alias", 1, 1));

            foreach (Cell applied in new[] {
                first.WorksheetPart.Worksheet.Descendants<Cell>().Single(cell => cell.CellReference?.Value == "B2"),
                second.WorksheetPart.Worksheet.Descendants<Cell>().Single(cell => cell.CellReference?.Value == "C3")
            }) {
                Assert.Equal(oldAppliedIndex, applied.StyleIndex!.Value);
                CellFormat appliedFormat = stylesheet.CellFormats!.Elements<CellFormat>()
                    .ElementAt((int)applied.StyleIndex.Value);
                Assert.Equal(normalFormatId, appliedFormat.FormatId!.Value);
            }
            Assert.Equal(normalFormatId, normal.FormatId!.Value);
            Assert.Equal(normalFormatXml, stylesheet.CellStyleFormats.Elements<CellFormat>()
                .ElementAt((int)normalFormatId).OuterXml);
            Assert.Empty(document.ValidateOpenXml());
        }

        [Fact]
        public void Test_TableSchema_PreservesImplicitShownTotalsRow() {
            using var document = ExcelDocument.Create(new MemoryStream());
            ExcelSheet sheet = document.AddSheet("Data");
            sheet.Cell(1, 1).SetValue("A");
            sheet.Cell(1, 2).SetValue("B");
            sheet.Cell(2, 1).SetValue(1);
            sheet.Cell(2, 2).SetValue(2);
            sheet.Cell(3, 1).SetValue("Total");
            sheet.AddTable("A1:B3", true, "DataTable", OfficeIMO.Excel.ExcelTableStyle.TableStyleMedium2);
            Table table = sheet.WorksheetPart.TableDefinitionParts.Single().Table!;
            table.TotalsRowShown = true;
            table.TotalsRowCount = null;

            Assert.Throws<InvalidOperationException>(() =>
                sheet.SetTableSchema("DataTable", new[] { "A", "B" }, "A1:B1"));

            sheet.SetTableSchema("DataTable", new[] { "A", "B" }, "A1:B2");

            Assert.Equal("A1:B2", table.Reference!.Value);
            Assert.Equal("A1:B1", table.GetFirstChild<AutoFilter>()!.Reference!.Value);
            Assert.Empty(document.ValidateOpenXml());
        }
    }
}
