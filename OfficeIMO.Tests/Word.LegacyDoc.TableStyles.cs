using DocumentFormat.OpenXml.Wordprocessing;
using OfficeIMO.Word;
using Xunit;

namespace OfficeIMO.Tests {
    public partial class Word {
        [Fact]
        public void LegacyDoc_SaveDocPath_WritesNativeDocConditionalTableStyleCornerPrecedenceAndReloadsThroughLegacyReader() {
            string docPath = Path.Combine(Path.GetTempPath(), Guid.NewGuid().ToString("N") + ".doc");
            const string styleId = "NativeDocCornerPrecedenceTable";

            try {
                using (WordDocument document = WordDocument.Create()) {
                    var style = new Style { Type = StyleValues.Table, StyleId = styleId, CustomStyle = true };
                    style.Append(new StyleName { Val = "Native DOC Corner Precedence Table" });
                    style.Append(new BasedOn { Val = "TableNormal" });
                    style.Append(new StyleTableProperties(new TableStyleRowBandSize { Val = 1 }));
                    style.Append(new TableStyleProperties(
                        new TableStyleConditionalFormattingTableCellProperties(
                            new Shading { Val = ShadingPatternValues.Clear, Fill = "ff0000" })) {
                        Type = TableStyleOverrideValues.FirstRow
                    });
                    style.Append(new TableStyleProperties(
                        new TableStyleConditionalFormattingTableCellProperties(
                            new Shading { Val = ShadingPatternValues.Clear, Fill = "00ff00" })) {
                        Type = TableStyleOverrideValues.NorthWestCell
                    });
                    style.Append(new TableStyleProperties(
                        new TableStyleConditionalFormattingTableCellProperties(
                            new Shading { Val = ShadingPatternValues.Clear, Fill = "ffff00" })) {
                        Type = TableStyleOverrideValues.Band1Horizontal
                    });
                    document._wordprocessingDocument!.MainDocumentPart!.StyleDefinitionsPart!.Styles!.Append(style);

                    WordTable table = document.AddTable(3, 2, WordTableStyle.TableNormal);
                    table._tableProperties!.TableStyle = new TableStyle { Val = styleId };
                    table._tableProperties.TableLook = new TableLook {
                        FirstRow = true,
                        FirstColumn = true,
                        LastRow = false,
                        LastColumn = false,
                        NoHorizontalBand = false,
                        NoVerticalBand = true
                    };
                    table.Rows[0].Cells[0].AddParagraph("A1", removeExistingParagraphs: true);
                    table.Rows[0].Cells[1].AddParagraph("B1", removeExistingParagraphs: true);
                    table.Rows[1].Cells[0].AddParagraph("A2", removeExistingParagraphs: true);
                    table.Rows[1].Cells[1].AddParagraph("B2", removeExistingParagraphs: true);
                    table.Rows[2].Cells[0].AddParagraph("A3", removeExistingParagraphs: true);
                    table.Rows[2].Cells[1].AddParagraph("B3", removeExistingParagraphs: true);

                    document.Save(docPath);
                }

                using WordDocument reloaded = WordDocument.Load(docPath);

                Assert.True(reloaded.WasLoadedFromLegacyDoc);
                WordTable reloadedTable = Assert.Single(reloaded.Tables);
                Assert.Equal("A1", reloadedTable.Rows[0].Cells[0].Paragraphs[0].Text);
                Assert.Equal("00ff00", reloadedTable.Rows[0].Cells[0].ShadingFillColorHex);
                Assert.Equal("ff0000", reloadedTable.Rows[0].Cells[1].ShadingFillColorHex);
                Assert.Equal("ffff00", reloadedTable.Rows[1].Cells[0].ShadingFillColorHex);
                Assert.Equal("ffff00", reloadedTable.Rows[1].Cells[1].ShadingFillColorHex);
                Assert.Equal(string.Empty, reloadedTable.Rows[2].Cells[0].ShadingFillColorHex);
                Assert.Equal(string.Empty, reloadedTable.Rows[2].Cells[1].ShadingFillColorHex);
            } finally {
                DeleteIfExists(docPath);
            }
        }
    }
}
