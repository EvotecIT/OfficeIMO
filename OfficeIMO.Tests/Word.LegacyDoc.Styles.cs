using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using OfficeIMO.Word;
using System;
using System.IO;
using System.Linq;
using Xunit;

namespace OfficeIMO.Tests {
    public partial class Word {
        [Fact]
        public void LegacyDoc_SaveDocPath_WritesNativeDocBuiltInStyleLayoutFlagsAndReloadsThroughLegacyReader() {
            string docPath = Path.Combine(Path.GetTempPath(), Guid.NewGuid().ToString("N") + ".doc");
            string headingStyleId = WordParagraphStyles.Heading1.ToStringStyle();

            try {
                using (WordDocument document = WordDocument.Create()) {
                    Styles styles = document._wordprocessingDocument!.MainDocumentPart!.StyleDefinitionsPart!.Styles!;
                    Style headingStyle = styles.Elements<Style>().FirstOrDefault(style => style.StyleId == headingStyleId)
                        ?? new Style { Type = StyleValues.Paragraph, StyleId = headingStyleId };
                    if (headingStyle.Parent == null) {
                        styles.Append(headingStyle);
                    }

                    headingStyle.StyleParagraphProperties = new StyleParagraphProperties(
                        new SuppressLineNumbers(),
                        new SuppressAutoHyphens(),
                        new ContextualSpacing(),
                        new MirrorIndents(),
                        new BiDi(),
                        new Kinsoku(),
                        new WordWrap(),
                        new OverflowPunctuation(),
                        new TopLinePunctuation(),
                        new AutoSpaceDE(),
                        new AutoSpaceDN());

                    document.AddParagraph("Styled built-in layout flags").SetStyle(WordParagraphStyles.Heading1);

                    document.Save(docPath);
                }

                using WordDocument reloaded = WordDocument.Load(docPath);

                Assert.True(reloaded.WasLoadedFromLegacyDoc);
                WordParagraph paragraph = Assert.Single(reloaded.Paragraphs);
                Assert.Equal("Styled built-in layout flags", paragraph.Text);
                Assert.Equal(headingStyleId, paragraph.StyleId);

                Styles stylesAfterReload = reloaded._wordprocessingDocument!.MainDocumentPart!.StyleDefinitionsPart!.Styles!;
                Style headingStyleAfterReload = Assert.Single(stylesAfterReload.Elements<Style>(), style => style.StyleId == headingStyleId);
                StyleParagraphProperties paragraphProperties = Assert.IsType<StyleParagraphProperties>(headingStyleAfterReload.StyleParagraphProperties);
                Assert.NotNull(paragraphProperties.GetFirstChild<SuppressLineNumbers>());
                Assert.NotNull(paragraphProperties.GetFirstChild<SuppressAutoHyphens>());
                Assert.NotNull(paragraphProperties.GetFirstChild<ContextualSpacing>());
                Assert.NotNull(paragraphProperties.GetFirstChild<MirrorIndents>());
                Assert.NotNull(paragraphProperties.GetFirstChild<BiDi>());
                Assert.NotNull(paragraphProperties.GetFirstChild<Kinsoku>());
                Assert.NotNull(paragraphProperties.GetFirstChild<WordWrap>());
                Assert.NotNull(paragraphProperties.GetFirstChild<OverflowPunctuation>());
                Assert.NotNull(paragraphProperties.GetFirstChild<TopLinePunctuation>());
                Assert.NotNull(paragraphProperties.GetFirstChild<AutoSpaceDE>());
                Assert.NotNull(paragraphProperties.GetFirstChild<AutoSpaceDN>());
            } finally {
                DeleteIfExists(docPath);
            }
        }

        [Fact]
        public void LegacyDoc_SaveDocPath_WritesNativeDocBuiltInStyleCapsAndVerticalPositionAndReloadsThroughLegacyReader() {
            string docPath = Path.Combine(Path.GetTempPath(), Guid.NewGuid().ToString("N") + ".doc");
            string capsStyleId = WordParagraphStyles.Heading1.ToStringStyle();
            string smallCapsStyleId = WordParagraphStyles.Heading2.ToStringStyle();
            string superscriptStyleId = WordParagraphStyles.Heading3.ToStringStyle();
            string subscriptStyleId = WordParagraphStyles.Heading4.ToStringStyle();

            try {
                using (WordDocument document = WordDocument.Create()) {
                    Styles styles = document._wordprocessingDocument!.MainDocumentPart!.StyleDefinitionsPart!.Styles!;
                    EnsureParagraphStyle(styles, capsStyleId).StyleRunProperties = new StyleRunProperties(new Caps());
                    EnsureParagraphStyle(styles, smallCapsStyleId).StyleRunProperties = new StyleRunProperties(new SmallCaps());
                    EnsureParagraphStyle(styles, superscriptStyleId).StyleRunProperties = new StyleRunProperties(new VerticalTextAlignment { Val = VerticalPositionValues.Superscript });
                    EnsureParagraphStyle(styles, subscriptStyleId).StyleRunProperties = new StyleRunProperties(new VerticalTextAlignment { Val = VerticalPositionValues.Subscript });

                    document.AddParagraph("Built-in caps paragraph").SetStyle(WordParagraphStyles.Heading1);
                    document.AddParagraph("Built-in small caps paragraph").SetStyle(WordParagraphStyles.Heading2);
                    document.AddParagraph("Built-in superscript paragraph").SetStyle(WordParagraphStyles.Heading3);
                    document.AddParagraph("Built-in subscript paragraph").SetStyle(WordParagraphStyles.Heading4);

                    document.Save(docPath);
                }

                byte[] tableStream = ReadCompoundStream(File.ReadAllBytes(docPath), "1Table");
                Assert.True(
                    ContainsBytePattern(tableStream, 0x3B, 0x08, 0x01),
                    "Expected the native DOC stylesheet stream to contain sprmCFCaps for built-in paragraph style all-caps.");
                Assert.True(
                    ContainsBytePattern(tableStream, 0x3A, 0x08, 0x01),
                    "Expected the native DOC stylesheet stream to contain sprmCFSmallCaps for built-in paragraph style small-caps.");
                Assert.True(
                    ContainsBytePattern(tableStream, 0x48, 0x2A, 0x01),
                    "Expected the native DOC stylesheet stream to contain sprmCIss superscript for built-in paragraph style vertical position.");
                Assert.True(
                    ContainsBytePattern(tableStream, 0x48, 0x2A, 0x02),
                    "Expected the native DOC stylesheet stream to contain sprmCIss subscript for built-in paragraph style vertical position.");

                using WordDocument reloaded = WordDocument.Load(docPath);

                Assert.True(reloaded.WasLoadedFromLegacyDoc);
                Assert.Equal(
                    new[] {
                        "Built-in caps paragraph",
                        "Built-in small caps paragraph",
                        "Built-in superscript paragraph",
                        "Built-in subscript paragraph"
                    },
                    reloaded.Paragraphs.Select(paragraph => paragraph.Text).ToArray());

                Styles stylesAfterReload = reloaded._wordprocessingDocument!.MainDocumentPart!.StyleDefinitionsPart!.Styles!;
                Assert.NotNull(AssertBuiltInStyleRunProperties(stylesAfterReload, capsStyleId).GetFirstChild<Caps>());
                Assert.NotNull(AssertBuiltInStyleRunProperties(stylesAfterReload, smallCapsStyleId).GetFirstChild<SmallCaps>());
                VerticalTextAlignment superscriptPosition = Assert.IsType<VerticalTextAlignment>(AssertBuiltInStyleRunProperties(stylesAfterReload, superscriptStyleId).GetFirstChild<VerticalTextAlignment>());
                Assert.Equal(VerticalPositionValues.Superscript, superscriptPosition.Val!.Value);
                VerticalTextAlignment subscriptPosition = Assert.IsType<VerticalTextAlignment>(AssertBuiltInStyleRunProperties(stylesAfterReload, subscriptStyleId).GetFirstChild<VerticalTextAlignment>());
                Assert.Equal(VerticalPositionValues.Subscript, subscriptPosition.Val!.Value);
            } finally {
                DeleteIfExists(docPath);
            }
        }

        [Fact]
        public void LegacyDoc_SaveDocPath_WritesNativeDocBuiltInStyleShadingAndBordersAndReloadsThroughLegacyReader() {
            string docPath = Path.Combine(Path.GetTempPath(), Guid.NewGuid().ToString("N") + ".doc");
            string headingStyleId = WordParagraphStyles.Heading1.ToStringStyle();

            try {
                using (WordDocument document = WordDocument.Create()) {
                    Styles styles = document._wordprocessingDocument!.MainDocumentPart!.StyleDefinitionsPart!.Styles!;
                    Style headingStyle = styles.Elements<Style>().FirstOrDefault(style => style.StyleId == headingStyleId)
                        ?? new Style { Type = StyleValues.Paragraph, StyleId = headingStyleId };
                    if (headingStyle.Parent == null) {
                        styles.Append(headingStyle);
                    }

                    headingStyle.StyleParagraphProperties = new StyleParagraphProperties(
                        new Shading { Val = ShadingPatternValues.Clear, Color = "auto", Fill = "FF0000" },
                        new ParagraphBorders(
                            new TopBorder { Val = BorderValues.Single, Color = "FF0000", Size = 8, Space = 2 },
                            new LeftBorder { Val = BorderValues.Double, Color = "0000FF", Size = 12, Space = 1 },
                            new BottomBorder { Val = BorderValues.Dotted, Color = "00FF00", Size = 4, Space = 0 },
                            new RightBorder { Val = BorderValues.Dashed, Color = "FFFF00", Size = 6, Space = 3 }));

                    document.AddParagraph("Styled built-in borders and shading").SetStyle(WordParagraphStyles.Heading1);

                    document.Save(docPath);
                }

                using WordDocument reloaded = WordDocument.Load(docPath);

                Assert.True(reloaded.WasLoadedFromLegacyDoc);
                WordParagraph paragraph = Assert.Single(reloaded.Paragraphs);
                Assert.Equal("Styled built-in borders and shading", paragraph.Text);
                Assert.Equal(headingStyleId, paragraph.StyleId);

                Styles stylesAfterReload = reloaded._wordprocessingDocument!.MainDocumentPart!.StyleDefinitionsPart!.Styles!;
                Style headingStyleAfterReload = Assert.Single(stylesAfterReload.Elements<Style>(), style => style.StyleId == headingStyleId);
                StyleParagraphProperties paragraphProperties = Assert.IsType<StyleParagraphProperties>(headingStyleAfterReload.StyleParagraphProperties);
                Shading shading = Assert.IsType<Shading>(paragraphProperties.GetFirstChild<Shading>());
                Assert.Equal(ShadingPatternValues.Clear, shading.Val!.Value);
                Assert.Equal("ff0000", shading.Fill!.Value);

                ParagraphBorders borders = Assert.IsType<ParagraphBorders>(paragraphProperties.GetFirstChild<ParagraphBorders>());
                Assert.Equal(BorderValues.Single, borders.TopBorder!.Val!.Value);
                Assert.Equal("ff0000", borders.TopBorder.Color!.Value);
                Assert.Equal(8U, borders.TopBorder.Size!.Value);
                Assert.Equal(2U, borders.TopBorder.Space!.Value);
                Assert.Equal(BorderValues.Double, borders.LeftBorder!.Val!.Value);
                Assert.Equal("0000ff", borders.LeftBorder.Color!.Value);
                Assert.Equal(12U, borders.LeftBorder.Size!.Value);
                Assert.Equal(1U, borders.LeftBorder.Space!.Value);
                Assert.Equal(BorderValues.Dotted, borders.BottomBorder!.Val!.Value);
                Assert.Equal("00ff00", borders.BottomBorder.Color!.Value);
                Assert.Equal(4U, borders.BottomBorder.Size!.Value);
                Assert.Equal(BorderValues.Dashed, borders.RightBorder!.Val!.Value);
                Assert.Equal("ffff00", borders.RightBorder.Color!.Value);
                Assert.Equal(6U, borders.RightBorder.Size!.Value);
                Assert.Equal(3U, borders.RightBorder.Space!.Value);
            } finally {
                DeleteIfExists(docPath);
            }
        }

        [Fact]
        public void LegacyDoc_SaveDocPath_WritesNativeDocBuiltInStyleTabStopsAndReloadsThroughLegacyReader() {
            string docPath = Path.Combine(Path.GetTempPath(), Guid.NewGuid().ToString("N") + ".doc");
            string headingStyleId = WordParagraphStyles.Heading1.ToStringStyle();

            try {
                using (WordDocument document = WordDocument.Create()) {
                    Styles styles = document._wordprocessingDocument!.MainDocumentPart!.StyleDefinitionsPart!.Styles!;
                    Style headingStyle = styles.Elements<Style>().FirstOrDefault(style => style.StyleId == headingStyleId)
                        ?? new Style { Type = StyleValues.Paragraph, StyleId = headingStyleId };
                    if (headingStyle.Parent == null) {
                        styles.Append(headingStyle);
                    }

                    headingStyle.StyleParagraphProperties = new StyleParagraphProperties(
                        new Tabs(
                            new TabStop { Val = TabStopValues.Left, Position = 1440 },
                            new TabStop { Val = TabStopValues.Decimal, Leader = TabStopLeaderCharValues.Dot, Position = 2880 },
                            new TabStop { Val = TabStopValues.Right, Leader = TabStopLeaderCharValues.Underscore, Position = 4320 }));

                    document.AddParagraph("Styled built-in tabs").SetStyle(WordParagraphStyles.Heading1);

                    document.Save(docPath);
                }

                using WordDocument reloaded = WordDocument.Load(docPath);

                Assert.True(reloaded.WasLoadedFromLegacyDoc);
                WordParagraph paragraph = Assert.Single(reloaded.Paragraphs);
                Assert.Equal("Styled built-in tabs", paragraph.Text);
                Assert.Equal(headingStyleId, paragraph.StyleId);

                Styles stylesAfterReload = reloaded._wordprocessingDocument!.MainDocumentPart!.StyleDefinitionsPart!.Styles!;
                Style headingStyleAfterReload = Assert.Single(stylesAfterReload.Elements<Style>(), style => style.StyleId == headingStyleId);
                StyleParagraphProperties paragraphProperties = Assert.IsType<StyleParagraphProperties>(headingStyleAfterReload.StyleParagraphProperties);
                Tabs tabs = Assert.IsType<Tabs>(paragraphProperties.GetFirstChild<Tabs>());
                TabStop[] tabStops = tabs.Elements<TabStop>().ToArray();
                Assert.Equal(3, tabStops.Length);
                Assert.Equal(TabStopValues.Left, tabStops[0].Val!.Value);
                Assert.Equal(1440, tabStops[0].Position!.Value);
                Assert.Equal(TabStopValues.Decimal, tabStops[1].Val!.Value);
                Assert.Equal(TabStopLeaderCharValues.Dot, tabStops[1].Leader!.Value);
                Assert.Equal(2880, tabStops[1].Position!.Value);
                Assert.Equal(TabStopValues.Right, tabStops[2].Val!.Value);
                Assert.Equal(TabStopLeaderCharValues.Underscore, tabStops[2].Leader!.Value);
                Assert.Equal(4320, tabStops[2].Position!.Value);
            } finally {
                DeleteIfExists(docPath);
            }
        }

        private static Style EnsureParagraphStyle(Styles styles, string styleId) {
            Style style = styles.Elements<Style>().FirstOrDefault(item => item.StyleId == styleId)
                ?? new Style { Type = StyleValues.Paragraph, StyleId = styleId };
            if (style.Parent == null) {
                styles.Append(style);
            }

            return style;
        }

        private static StyleRunProperties AssertBuiltInStyleRunProperties(Styles styles, string styleId) {
            Style style = Assert.Single(styles.Elements<Style>(), item => item.StyleId == styleId);
            return Assert.IsType<StyleRunProperties>(style.StyleRunProperties);
        }
    }
}
