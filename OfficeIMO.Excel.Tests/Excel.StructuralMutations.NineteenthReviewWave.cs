using System;
using System.IO;
using System.Linq;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using OfficeIMO.Excel;
using Xunit;
using Rich = DocumentFormat.OpenXml.Office2019.Excel.RichData;

namespace OfficeIMO.Tests {
    public partial class Excel {
        private static readonly byte[] TinyGif = Convert.FromBase64String(
            "R0lGODlhAQABAJAAAAAAAP///ywAAAAAAQABAAACAkwBADs=");

        [Fact]
        public void Test_AutoFilterCriteria_PreserveEquivalentImportedRangeState() {
            using var document = ExcelDocument.Create();
            ExcelSheet sheet = document.AddWorksheet("Data");
            sheet.AutoFilterBlanks("A1:C10", 0);
            AutoFilter autoFilter = Assert.Single(sheet.WorksheetPart.Worksheet.Elements<AutoFilter>());
            autoFilter.Reference = "$A$1:$C$10";

            sheet.AutoFilterBlanks("A1:C10", 1);

            ExcelAutoFilterInfo state = Assert.Single(sheet.GetAutoFilters());
            Assert.Equal("$A$1:$C$10", state.Range);
            Assert.Equal(new uint[] { 0U, 1U }, state.Columns.Select(column => column.ColumnOffset).ToArray());
        }

        [Fact]
        public void Test_TableSchema_NormalizesDisplayAliasWhenStableNameMatchesTarget() {
            using var document = ExcelDocument.Create();
            ExcelSheet sheet = document.AddWorksheet("Data");
            sheet.CellValue(1, 1, "Amount");
            sheet.CellValue(2, 1, 10);
            sheet.AddTable("A1:A2", true, "Internal", OfficeIMO.Excel.ExcelTableStyle.TableStyleMedium2);
            Table table = Assert.Single(sheet.WorksheetPart.TableDefinitionParts).Table!;
            table.DisplayName = "Sales";
            table.Save();
            sheet.CellFormula(1, 3, "SUM(Sales[Amount])");

            Assert.Equal("Internal", sheet.RenameTable("Sales", "Internal"));

            Assert.Equal("Internal", table.Name!.Value);
            Assert.Equal("Internal", table.DisplayName!.Value);
            Assert.Equal("SUM(Internal[Amount])", sheet.GetFormulaText(1, 3));
        }

        [Fact]
        public void Test_InCellImage_ReplacementReusesExclusiveAssetsAndPreservesSharedCopies() {
            using var document = ExcelDocument.Create();
            ExcelSheet sheet = document.AddWorksheet("Images");
            sheet.SetInCellImage(1, 1, TinyPng, altText: "First");

            sheet.SetInCellImage(1, 1, TinyPng, altText: "Second");
            AssertRichImageGraphCounts(document, expected: 1);
            Assert.Equal("Second", Assert.Single(sheet.GetInCellImages()).AltText);

            sheet.SetInCellImage(1, 1, TinyGif, "image/gif", "Third");
            AssertRichImageGraphCounts(document, expected: 1);
            ExcelInCellImage replacement = Assert.Single(sheet.GetInCellImages());
            Assert.Equal("image/gif", replacement.ContentType);
            Assert.Equal(TinyGif, replacement.Bytes);

            sheet.Range("A1").CopyTo("B1");
            sheet.SetInCellImage(1, 1, TinyPng, altText: "Independent");

            AssertRichImageGraphCounts(document, expected: 2);
            ExcelInCellImage[] images = sheet.GetInCellImages().OrderBy(image => image.CellReference).ToArray();
            Assert.Equal(new[] { "A1", "B1" }, images.Select(image => image.CellReference).ToArray());
            Assert.Equal("Independent", images[0].AltText);
            Assert.Equal("Third", images[1].AltText);
            Assert.Equal(TinyGif, images[1].Bytes);

            sheet.Range("B1").CopyTo("C1");
            long aggregateBudget = TinyPng.LongLength + (2L * TinyGif.LongLength);
            ExcelInCellImage[] shared = sheet.GetInCellImages(aggregateBudget).ToArray();
            Assert.Equal(3, shared.Length);
            shared[1].Bytes[0] = 0;
            Assert.Equal(TinyGif[0], shared[2].Bytes[0]);
            Assert.Throws<InvalidOperationException>(() => sheet.GetInCellImages(aggregateBudget - 1L));
        }

        [Fact]
        public void Test_StructuralColumns_HandleImplicitRowsAndHeaderlessTables() {
            using var document = ExcelDocument.Create();
            ExcelSheet implicitRows = document.AddWorksheet("Implicit");
            implicitRows.CellValue(2, 2, "B2");
            implicitRows.CellValue(2, 3, "C2");
            Row implicitRow = Assert.Single(implicitRows.WorksheetPart.Worksheet
                .GetFirstChild<SheetData>()!.Elements<Row>());
            implicitRow.RowIndex = null;

            implicitRows.DeleteColumns(1);

            Assert.Null(implicitRow.RowIndex);
            Assert.Equal(new[] { "A2", "B2" }, implicitRow.Elements<Cell>()
                .Select(cell => cell.CellReference!.Value).ToArray());

            ExcelSheet headerless = document.AddWorksheet("Headerless");
            headerless.CellValue(1, 1, "Alpha");
            headerless.CellValue(1, 2, "Beta");
            headerless.CellValue(2, 1, "Gamma");
            headerless.CellValue(2, 2, "Delta");
            headerless.AddTable("A1:B2", false, "Data", OfficeIMO.Excel.ExcelTableStyle.TableStyleMedium2);

            headerless.InsertColumns(2);

            Cell[] firstRow = headerless.WorksheetPart.Worksheet.GetFirstChild<SheetData>()!
                .Elements<Row>().First().Elements<Cell>().ToArray();
            Assert.Equal(new[] { "A1", "C1" }, firstRow.Select(cell => cell.CellReference!.Value).ToArray());
            Assert.True(headerless.TryGetCellValueSnapshot(1, 1, out ExcelCellValueSnapshot? alpha));
            Assert.True(headerless.TryGetCellValueSnapshot(1, 3, out ExcelCellValueSnapshot? beta));
            Assert.Equal("Alpha", alpha!.Text);
            Assert.Equal("Beta", beta!.Text);
            Table resized = Assert.Single(headerless.WorksheetPart.TableDefinitionParts).Table!;
            Assert.Equal(0U, resized.HeaderRowCount!.Value);
            Assert.Equal(3, resized.TableColumns!.Elements<TableColumn>().Count());
        }

        private static void AssertRichImageGraphCounts(ExcelDocument document, int expected) {
            WorkbookPart workbookPart = document.WorkbookPartRoot;
            Assert.Equal(expected, workbookPart.RdRichValueParts.Single().RichValueData!
                .Elements<Rich.RichValue>().Count());
            Assert.Equal(expected, workbookPart.CellMetadataPart!.Metadata!
                .GetFirstChild<ValueMetadata>()!.Elements<MetadataBlock>().Count());
            ExtendedPart relationships = workbookPart.Parts.Select(pair => pair.OpenXmlPart)
                .OfType<ExtendedPart>()
                .Single(part => part.RelationshipType.EndsWith("/richValueRel", StringComparison.Ordinal));
            Assert.Equal(expected, relationships.Parts.Count());
        }
    }
}
