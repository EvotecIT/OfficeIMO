using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using OfficeIMO.Excel;
using OfficeIMO.Excel.Html;
using OfficeIMO.Html;
using System.Threading;
using Xunit;
using Rich = DocumentFormat.OpenXml.Office2019.Excel.RichData;

namespace OfficeIMO.Tests {
    public partial class Excel {
        [Theory]
        [InlineData("8", false, "#SPILL!")]
        [InlineData("13", false, "#CALC!")]
        [InlineData("42", false, "#VALUE!")]
        [InlineData("8", true, "#VALUE!")]
        public void RichValueAliasesPreserveErrorsAndMalformedFallback(string code, bool duplicateKey, string expected) {
            using var document = ExcelDocument.Create();
            var sheet = document.AddWorksheet("Data");
            document.AddWorksheet("Other");
            sheet.SetArrayFormula("A1:A1", "SEQUENCE(0)");
            document.Calculate();
            var workbook = document.WorkbookPartRoot;
            Metadata metadata = workbook.CellMetadataPart!.Metadata!;
            ValueMetadata blocks = metadata.GetFirstChild<ValueMetadata>()!;
            MetadataBlock template = blocks.Elements<MetadataBlock>().First();
            for (int index = 1; index < 128; index++) blocks.Append(template.CloneNode(true));
            blocks.Count = 128;
            var values = workbook.RdRichValueParts.Single().RichValueData!;
            Rich.RichValue value = values.Elements<Rich.RichValue>().First();
            var structures = workbook.GetPartsOfType<RdRichValueStructurePart>().Single().RichValueStructures!;
            Rich.RichValueStructure structure = structures.Elements<Rich.RichValueStructure>().First();
            value.RemoveAllChildren();
            structure.RemoveAllChildren();
            for (int index = 0; index < 64; index++) {
                structure.Append(new Rich.Key {
                    N = index == 63 || (duplicateKey && index == 0) ? "errorType" : "key" + index,
                    T = Rich.RichValueValueType.I
                });
                value.Append(new Rich.Value(index == 63 ? code : "0"));
            }
            Cell cell = sheet.WorksheetPart.Worksheet.Descendants<Cell>().Single();
            cell.ValueMetaIndex = 128;
            var lookup = RichValueErrorLookup.FromRoots(metadata, values, structures);
            for (uint index = 1; index <= 128; index++) Assert.Equal(expected, lookup.Resolve(index, "#VALUE!"));
            Assert.Equal(expected, sheet.CellAt(1, 1).GetValue().Value);
            byte[] bytes = document.ToBytes();
            foreach (string? selection in new string?[] { null, "Data" }) {
                using var reader = ExcelDocument.OpenDataReader(new MemoryStream(bytes),
                    new ExcelReadOptions { SheetName = selection, A1Range = "A1:A1", HasHeaderRow = false });
                Assert.True(reader.Read());
                Assert.Equal(expected, reader.GetValue(0));
            }
            if (!duplicateKey) {
                value.Elements<Rich.Value>().Last().Text = "13";
                Assert.Equal("#CALC!", RichValueErrorLookup.FromRoots(metadata, values, structures).Resolve(128, "#VALUE!"));
            }
        }

        [Fact]
        public void ImportedRowLayoutPreservesCellsPrecisionAndHiddenState() {
            using var document = ExcelDocument.Create();
            var sheet = document.AddWorksheet("Data");
            sheet.CellValue(3, 1, "Keep");
            sheet.CellValue(1002, 1, "Tail");
            sheet.SetRowHeight(3, 25);
            sheet.SetRowHidden(1002, true);
            var heights = Enumerable.Range(1, 1000).Reverse().ToDictionary(index => index, _ => 20.123456789d);
            heights[3] = 0;
            sheet.SetImportedRowLayout(heights, new[] { 2, 1001 }, CancellationToken.None);
            using var reopened = ExcelDocument.Load(new MemoryStream(document.ToBytes()));
            var rows = reopened["Data"].WorksheetPart.Worksheet.GetFirstChild<SheetData>()!.Elements<Row>().ToArray();
            Assert.Equal(Enumerable.Range(1, 1002).Select(index => (uint)index), rows.Select(row => row.RowIndex!.Value));
            Assert.Equal(20.123456789d, rows[0].Height!.Value);
            Assert.Null(rows[2].Height);
            Assert.Null(rows[2].CustomHeight);
            Assert.True(rows[1].Hidden!.Value);
            Assert.True(rows[1000].Hidden!.Value);
            Assert.True(rows[1001].Hidden!.Value);
            Assert.Equal("Keep", reopened["Data"].CellAt(3, 1).GetValue().Value);
            Assert.Equal("Tail", reopened["Data"].CellAt(1002, 1).GetValue().Value);
            Assert.Empty(reopened.ValidateOpenXml());
            using var canceled = new CancellationTokenSource();
            canceled.Cancel();
            Assert.Throws<OperationCanceledException>(() => sheet.SetImportedRowLayout(
                new Dictionary<int, double> { [2000] = 30 }, Array.Empty<int>(), canceled.Token));
            Assert.DoesNotContain(sheet.WorksheetPart.Worksheet.Descendants<Row>(), row => row.RowIndex?.Value == 2000);
        }

        [Theory]
        [InlineData("WWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWWW")]
        [InlineData("👩‍💻👩‍💻👩‍💻👩‍💻👩‍💻👩‍💻👩‍💻👩‍💻👩‍💻👩‍💻👩‍💻👩‍💻")]
        [InlineData("ẂẂẂẂẂẂẂẂẂẂẂẂẂẂẂẂẂẂẂẂẂẂẂẂẂẂẂẂẂẂẂẂ")]
        public void WrappedRowsIgnoreZeroWidthPrefixWithoutLosingGraphemes(string visible) {
            using var document = ExcelDocument.Create();
            var sheet = document.AddWorksheet("Data");
            sheet.SetColumnWidth(1, 5);
            sheet.CellValue(1, 1, visible);
            sheet.CellValue(2, 1, new string('\u200B', 30000) + visible);
            sheet.CellWrapText(1, 1);
            sheet.CellWrapText(2, 1);
            sheet.AutoFitRows();
            var rows = sheet.WorksheetPart.Worksheet.Descendants<Row>().ToArray();
            Assert.True(rows[0].Height!.Value > 15);
            Assert.Equal(rows[0].Height!.Value, rows[1].Height!.Value);
            Assert.Equal(new string('\u200B', 30000) + visible, sheet.CellAt(2, 1).GetValue().Value);
        }

        [Fact]
        public void GenericHtmlTableAndCaptionRetainLongZeroWidthText() {
            string text = new string('\u200B', 30000) + new string('W', 100);
            using var document = HtmlConversionDocument.Parse(
                "<table><caption>" + text + "</caption><tr><td>" + text + "</td><td>Definition</td></tr></table>")
                .ToExcelDocumentResult(new HtmlToExcelOptions { Mode = HtmlImportMode.Generic }).RequireValue();
            var sheet = Assert.Single(document.Sheets);
            Assert.Equal(text, sheet.CellAt(1, 1).GetValue().Value);
            Assert.Equal(text, sheet.CellAt(3, 1).GetValue().Value);
            Assert.All(sheet.WorksheetPart.Worksheet.Descendants<Row>()
                .Where(row => row.RowIndex?.Value == 1 || row.RowIndex?.Value == 3),
                row => Assert.True(row.Height?.Value > 15));
        }
    }
}
