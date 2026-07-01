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

        [Fact]
        public void LegacyDoc_SaveDocPath_WritesNativeDocConditionalTableStyleWholeTableAndReloadsThroughLegacyReader() {
            string docPath = Path.Combine(Path.GetTempPath(), Guid.NewGuid().ToString("N") + ".doc");
            const string styleId = "NativeDocWholeTableConditional";

            try {
                using (WordDocument document = WordDocument.Create()) {
                    var style = new Style { Type = StyleValues.Table, StyleId = styleId, CustomStyle = true };
                    style.Append(new StyleName { Val = "Native DOC Whole Table Conditional" });
                    style.Append(new BasedOn { Val = "TableNormal" });
                    style.Append(new TableStyleProperties(
                        new TableStyleConditionalFormattingTableCellProperties(
                            new Shading { Val = ShadingPatternValues.Clear, Fill = "ffff00" })) {
                        Type = TableStyleOverrideValues.WholeTable
                    });
                    style.Append(new TableStyleProperties(
                        new TableStyleConditionalFormattingTableCellProperties(
                            new Shading { Val = ShadingPatternValues.Clear, Fill = "ff0000" })) {
                        Type = TableStyleOverrideValues.FirstRow
                    });
                    document._wordprocessingDocument!.MainDocumentPart!.StyleDefinitionsPart!.Styles!.Append(style);

                    WordTable table = document.AddTable(2, 2, WordTableStyle.TableNormal);
                    table._tableProperties!.TableStyle = new TableStyle { Val = styleId };
                    table._tableProperties.TableLook = new TableLook {
                        FirstRow = true,
                        FirstColumn = false,
                        LastRow = false,
                        LastColumn = false,
                        NoHorizontalBand = true,
                        NoVerticalBand = true
                    };
                    table.Rows[0].Cells[0].AddParagraph("A1", removeExistingParagraphs: true);
                    table.Rows[0].Cells[1].AddParagraph("B1", removeExistingParagraphs: true);
                    table.Rows[1].Cells[0].AddParagraph("A2", removeExistingParagraphs: true);
                    table.Rows[1].Cells[1].AddParagraph("B2", removeExistingParagraphs: true);

                    document.Save(docPath);
                }

                using WordDocument reloaded = WordDocument.Load(docPath);

                Assert.True(reloaded.WasLoadedFromLegacyDoc);
                WordTable reloadedTable = Assert.Single(reloaded.Tables);
                Assert.Equal("ff0000", reloadedTable.Rows[0].Cells[0].ShadingFillColorHex);
                Assert.Equal("ff0000", reloadedTable.Rows[0].Cells[1].ShadingFillColorHex);
                Assert.Equal("ffff00", reloadedTable.Rows[1].Cells[0].ShadingFillColorHex);
                Assert.Equal("ffff00", reloadedTable.Rows[1].Cells[1].ShadingFillColorHex);
            } finally {
                DeleteIfExists(docPath);
            }
        }

        [Fact]
        public void LegacyDoc_SaveDocPath_WritesNativeDocCustomTableStyleBasedOnCustomLayoutStyleAndReloadsThroughLegacyReader() {
            string docPath = Path.Combine(Path.GetTempPath(), Guid.NewGuid().ToString("N") + ".doc");
            const string baseStyleId = "NativeDocBaseLayoutTable";
            const string childStyleId = "NativeDocInheritedLayoutTable";

            try {
                using (WordDocument document = WordDocument.Create()) {
                    var baseStyle = new Style { Type = StyleValues.Table, StyleId = baseStyleId, CustomStyle = true };
                    baseStyle.Append(new StyleName { Val = "Native DOC Base Layout Table" });
                    baseStyle.Append(new BasedOn { Val = "TableNormal" });
                    baseStyle.Append(new StyleTableProperties(
                        new TableJustification { Val = TableRowAlignmentValues.Right },
                        new TableIndentation { Width = 720, Type = TableWidthUnitValues.Dxa },
                        new TableWidth { Width = "3750", Type = TableWidthUnitValues.Pct },
                        new TableLayout { Type = TableLayoutValues.Fixed },
                        new TableCellMarginDefault(
                            new TopMargin { Width = "120", Type = TableWidthUnitValues.Dxa },
                            new TableCellLeftMargin { Width = 180, Type = TableWidthValues.Dxa },
                            new BottomMargin { Width = "160", Type = TableWidthUnitValues.Dxa },
                            new TableCellRightMargin { Width = 300, Type = TableWidthValues.Dxa }),
                        new TableCellSpacing { Width = "240", Type = TableWidthUnitValues.Dxa }));

                    var childStyle = new Style { Type = StyleValues.Table, StyleId = childStyleId, CustomStyle = true };
                    childStyle.Append(new StyleName { Val = "Native DOC Inherited Layout Table" });
                    childStyle.Append(new BasedOn { Val = baseStyleId });
                    childStyle.Append(new StyleTableProperties(
                        new TableWidth { Width = "2160", Type = TableWidthUnitValues.Dxa },
                        new TableCellMarginDefault(
                            new TableCellLeftMargin { Width = 360, Type = TableWidthValues.Dxa })));

                    Styles styles = document._wordprocessingDocument!.MainDocumentPart!.StyleDefinitionsPart!.Styles!;
                    styles.Append(baseStyle);
                    styles.Append(childStyle);

                    WordTable table = document.AddTable(1, 2, WordTableStyle.TableNormal);
                    table._tableProperties!.TableStyle = new TableStyle { Val = childStyleId };
                    table.Rows[0].Cells[0].WidthType = TableWidthUnitValues.Dxa;
                    table.Rows[0].Cells[0].Width = 1440;
                    table.Rows[0].Cells[0].AddParagraph("Inherited layout", removeExistingParagraphs: true);
                    table.Rows[0].Cells[1].WidthType = TableWidthUnitValues.Dxa;
                    table.Rows[0].Cells[1].Width = 1440;
                    table.Rows[0].Cells[1].AddParagraph("Merged margins", removeExistingParagraphs: true);

                    document.Save(docPath);
                }

                using WordDocument reloaded = WordDocument.Load(docPath);

                Assert.True(reloaded.WasLoadedFromLegacyDoc);
                WordTable reloadedTable = Assert.Single(reloaded.Tables);
                Assert.Equal(TableRowAlignmentValues.Right, reloadedTable.Alignment);
                Assert.Equal((short)720, reloadedTable.StyleDetails!.TableIndentationWidth);
                Assert.Equal(TableWidthUnitValues.Dxa, reloadedTable.WidthType);
                Assert.Equal(2160, reloadedTable.Width);
                Assert.Equal(TableLayoutValues.Fixed, reloadedTable.LayoutType);
                Assert.Equal((short)240, reloadedTable.StyleDetails.CellSpacing);

                WordTableRow row = Assert.Single(reloadedTable.Rows);
                Assert.Equal("Inherited layout", row.Cells[0].Paragraphs[0].Text);
                Assert.Equal((short)120, row.Cells[0].MarginTopWidth);
                Assert.Equal((short)360, row.Cells[0].MarginLeftWidth);
                Assert.Equal((short)160, row.Cells[0].MarginBottomWidth);
                Assert.Equal((short)300, row.Cells[0].MarginRightWidth);
                Assert.Equal("Merged margins", row.Cells[1].Paragraphs[0].Text);
                Assert.Equal((short)360, row.Cells[1].MarginLeftWidth);
            } finally {
                DeleteIfExists(docPath);
            }
        }

        [Fact]
        public void LegacyDoc_SaveDocPath_WritesNativeDocCustomTableStyleInheritedBandSizeAndReloadsThroughLegacyReader() {
            string docPath = Path.Combine(Path.GetTempPath(), Guid.NewGuid().ToString("N") + ".doc");
            const string baseStyleId = "NativeDocBaseBandSizeTable";
            const string childStyleId = "NativeDocInheritedBandSizeTable";

            try {
                using (WordDocument document = WordDocument.Create()) {
                    var baseStyle = new Style { Type = StyleValues.Table, StyleId = baseStyleId, CustomStyle = true };
                    baseStyle.Append(new StyleName { Val = "Native DOC Base Band Size Table" });
                    baseStyle.Append(new BasedOn { Val = "TableNormal" });
                    baseStyle.Append(new StyleTableProperties(new TableStyleRowBandSize { Val = 2 }));

                    var childStyle = new Style { Type = StyleValues.Table, StyleId = childStyleId, CustomStyle = true };
                    childStyle.Append(new StyleName { Val = "Native DOC Inherited Band Size Table" });
                    childStyle.Append(new BasedOn { Val = baseStyleId });
                    childStyle.Append(new TableStyleProperties(
                        new TableStyleConditionalFormattingTableCellProperties(
                            new Shading { Val = ShadingPatternValues.Clear, Fill = "ffff00" })) {
                        Type = TableStyleOverrideValues.Band1Horizontal
                    });

                    Styles styles = document._wordprocessingDocument!.MainDocumentPart!.StyleDefinitionsPart!.Styles!;
                    styles.Append(baseStyle);
                    styles.Append(childStyle);

                    WordTable table = document.AddTable(4, 1, WordTableStyle.TableNormal);
                    table._tableProperties!.TableStyle = new TableStyle { Val = childStyleId };
                    table._tableProperties.TableLook = new TableLook {
                        FirstRow = false,
                        FirstColumn = false,
                        LastRow = false,
                        LastColumn = false,
                        NoHorizontalBand = false,
                        NoVerticalBand = true
                    };

                    for (int row = 0; row < 4; row++) {
                        table.Rows[row].Cells[0].AddParagraph($"R{row + 1}", removeExistingParagraphs: true);
                    }

                    document.Save(docPath);
                }

                using WordDocument reloaded = WordDocument.Load(docPath);

                Assert.True(reloaded.WasLoadedFromLegacyDoc);
                WordTable reloadedTable = Assert.Single(reloaded.Tables);
                Assert.Equal("ffff00", reloadedTable.Rows[0].Cells[0].ShadingFillColorHex);
                Assert.Equal("ffff00", reloadedTable.Rows[1].Cells[0].ShadingFillColorHex);
                Assert.Equal(string.Empty, reloadedTable.Rows[2].Cells[0].ShadingFillColorHex);
                Assert.Equal(string.Empty, reloadedTable.Rows[3].Cells[0].ShadingFillColorHex);
            } finally {
                DeleteIfExists(docPath);
            }
        }

        [Fact]
        public void LegacyDoc_SaveDocPath_WritesNativeDocCustomTableStyleInheritedParagraphAndRunFormattingAndReloadsThroughLegacyReader() {
            string docPath = Path.Combine(Path.GetTempPath(), Guid.NewGuid().ToString("N") + ".doc");
            const string baseStyleId = "NativeDocBaseTextFormattingTable";
            const string childStyleId = "NativeDocInheritedTextFormattingTable";

            try {
                using (WordDocument document = WordDocument.Create()) {
                    var baseStyle = new Style { Type = StyleValues.Table, StyleId = baseStyleId, CustomStyle = true };
                    baseStyle.Append(new StyleName { Val = "Native DOC Base Text Formatting Table" });
                    baseStyle.Append(new BasedOn { Val = "TableNormal" });
                    baseStyle.Append(new StyleParagraphProperties(
                        new Justification { Val = JustificationValues.Center },
                        new Indentation { Left = "360" }));
                    baseStyle.Append(new StyleRunProperties(
                        new Bold(),
                        new Color { Val = "ff0000" },
                        new FontSize { Val = "28" }));

                    var childStyle = new Style { Type = StyleValues.Table, StyleId = childStyleId, CustomStyle = true };
                    childStyle.Append(new StyleName { Val = "Native DOC Inherited Text Formatting Table" });
                    childStyle.Append(new BasedOn { Val = baseStyleId });
                    childStyle.Append(new StyleParagraphProperties(
                        new SpacingBetweenLines { After = "240" }));
                    childStyle.Append(new StyleRunProperties(
                        new Italic(),
                        new Color { Val = "0000ff" }));

                    Styles styles = document._wordprocessingDocument!.MainDocumentPart!.StyleDefinitionsPart!.Styles!;
                    styles.Append(baseStyle);
                    styles.Append(childStyle);

                    WordTable table = document.AddTable(1, 1, WordTableStyle.TableNormal);
                    table._tableProperties!.TableStyle = new TableStyle { Val = childStyleId };
                    table.Rows[0].Cells[0].AddParagraph("Inherited text", removeExistingParagraphs: true);

                    document.Save(docPath);
                }

                using WordDocument reloaded = WordDocument.Load(docPath);

                Assert.True(reloaded.WasLoadedFromLegacyDoc);
                WordTable reloadedTable = Assert.Single(reloaded.Tables);
                WordParagraph paragraph = Assert.Single(Assert.Single(reloadedTable.Rows).Cells[0].Paragraphs);
                Assert.Equal("Inherited text", paragraph.Text);
                Assert.Equal(JustificationValues.Center, paragraph.ParagraphAlignment);
                Assert.Equal(360, paragraph.IndentationBefore);
                Assert.Equal(240, paragraph.LineSpacingAfter);
                Assert.True(paragraph.Bold);
                Assert.True(paragraph.Italic);
                Assert.Equal("0000ff", paragraph.ColorHex);
                Assert.Equal(14, paragraph.FontSize);
            } finally {
                DeleteIfExists(docPath);
            }
        }
    }
}
