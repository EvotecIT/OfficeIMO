using DocumentFormat.OpenXml.Wordprocessing;
using OfficeIMO.Word;
using Xunit;

namespace OfficeIMO.Tests {
    public partial class Word {
        [Fact]
        public void LegacyDoc_SaveDocPath_WritesNativeDocSectionPageBordersAndReloadsThroughLegacyReader() {
            string docPath = Path.Combine(Path.GetTempPath(), Guid.NewGuid().ToString("N") + ".doc");

            try {
                using (WordDocument document = WordDocument.Create()) {
                    document.AddParagraph("Section page borders");
                    document.Sections[0]._sectionProperties.Append(new PageBorders(
                        new TopBorder {
                            Val = BorderValues.Single,
                            Size = 8U,
                            Space = 12U,
                            Color = "FF0000"
                        },
                        new LeftBorder {
                            Val = BorderValues.Double,
                            Size = 12U,
                            Space = 10U,
                            Color = "0000FF"
                        },
                        new BottomBorder {
                            Val = BorderValues.Dotted,
                            Size = 4U,
                            Space = 8U,
                            Color = "00FF00"
                        },
                        new RightBorder {
                            Val = BorderValues.Dashed,
                            Size = 6U,
                            Space = 6U,
                            Color = "FFFF00"
                        }));

                    document.Save(docPath);
                }

                byte[] wordDocumentStream = ReadCompoundStream(File.ReadAllBytes(docPath), "WordDocument");
                Assert.True(
                    ContainsBytePattern(wordDocumentStream, 0x2B, 0x70, 0x08, 0x01, 0x06, 0x0C),
                    "Expected the native DOC section property block to contain sprmSBrcTop80 for the top page border.");
                Assert.True(
                    ContainsBytePattern(wordDocumentStream, 0x2C, 0x70, 0x0C, 0x03, 0x02, 0x0A),
                    "Expected the native DOC section property block to contain sprmSBrcLeft80 for the left page border.");
                Assert.True(
                    ContainsBytePattern(wordDocumentStream, 0x2D, 0x70, 0x04, 0x06, 0x04, 0x08),
                    "Expected the native DOC section property block to contain sprmSBrcBottom80 for the bottom page border.");
                Assert.True(
                    ContainsBytePattern(wordDocumentStream, 0x2E, 0x70, 0x06, 0x07, 0x07, 0x06),
                    "Expected the native DOC section property block to contain sprmSBrcRight80 for the right page border.");

                using WordDocument reloaded = WordDocument.Load(docPath);

                Assert.True(reloaded.WasLoadedFromLegacyDoc);
                Assert.Equal("Section page borders", Assert.Single(reloaded.Paragraphs).Text);

                PageBorders pageBorders = Assert.IsType<PageBorders>(reloaded.Sections[0]._sectionProperties.GetFirstChild<PageBorders>());
                AssertPageBorder(pageBorders.TopBorder, BorderValues.Single, "ff0000", 8U, 12U);
                AssertPageBorder(pageBorders.LeftBorder, BorderValues.Double, "0000ff", 12U, 10U);
                AssertPageBorder(pageBorders.BottomBorder, BorderValues.Dotted, "00ff00", 4U, 8U);
                AssertPageBorder(pageBorders.RightBorder, BorderValues.Dashed, "ffff00", 6U, 6U);
            } finally {
                DeleteIfExists(docPath);
            }
        }

        [Fact]
        public void LegacyDoc_SaveDocPath_WritesNativeDocSectionPageBorderPositioningAndReloadsThroughLegacyReader() {
            string docPath = Path.Combine(Path.GetTempPath(), Guid.NewGuid().ToString("N") + ".doc");

            try {
                using (WordDocument document = WordDocument.Create()) {
                    document.AddParagraph("Section page border positioning");
                    document.Sections[0]._sectionProperties.Append(new PageBorders(
                        new TopBorder {
                            Val = BorderValues.Single,
                            Size = 4U,
                            Space = 24U,
                            Color = "000000"
                        }) {
                        Display = PageBorderDisplayValues.NotFirstPage,
                        OffsetFrom = PageBorderOffsetValues.Page,
                        ZOrder = PageBorderZOrderValues.Back
                    });

                    document.Save(docPath);
                }

                byte[] wordDocumentStream = ReadCompoundStream(File.ReadAllBytes(docPath), "WordDocument");
                Assert.True(
                    ContainsBytePattern(wordDocumentStream, 0x2F, 0x52, 0x2A, 0x00),
                    "Expected the native DOC section property block to contain sprmSPgbProp for page-border positioning.");

                using WordDocument reloaded = WordDocument.Load(docPath);

                Assert.True(reloaded.WasLoadedFromLegacyDoc);
                Assert.Equal("Section page border positioning", Assert.Single(reloaded.Paragraphs).Text);

                PageBorders pageBorders = Assert.IsType<PageBorders>(reloaded.Sections[0]._sectionProperties.GetFirstChild<PageBorders>());
                Assert.Equal(PageBorderDisplayValues.NotFirstPage, pageBorders.Display!.Value);
                Assert.Equal(PageBorderOffsetValues.Page, pageBorders.OffsetFrom!.Value);
                Assert.Equal(PageBorderZOrderValues.Back, pageBorders.ZOrder!.Value);
                AssertPageBorder(pageBorders.TopBorder, BorderValues.Single, "000000", 4U, 24U);
            } finally {
                DeleteIfExists(docPath);
            }
        }

        private static void AssertPageBorder(BorderType? border, BorderValues expectedStyle, string expectedColor, uint expectedSize, uint expectedSpace) {
            Assert.NotNull(border);
            Assert.Equal(expectedStyle, border!.Val!.Value);
            Assert.Equal(expectedColor, border.Color!.Value);
            Assert.Equal(expectedSize, border.Size!.Value);
            Assert.Equal(expectedSpace, border.Space!.Value);
        }
    }
}
