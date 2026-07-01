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
    }
}
