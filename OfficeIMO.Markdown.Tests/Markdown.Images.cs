using System;
using System.Collections.Generic;
using System.IO;
using Color = OfficeIMO.OfficeColor;
using OfficeIMO.Word;
using OfficeIMO.Word.Markdown;
using OfficeIMO.Markdown.Pdf;
using OfficeIMO.Markdown.Html;
using Xunit;

namespace OfficeIMO.Tests {
    public partial class Markdown {
        [Fact]
        public void MarkdownReader_PreservesWindowsImagePathsUnlessFileUrlsAreDisallowed() {
            const string windowsPath = @"C:\Reports\Images\OfficeIMO.png";
            string markdown = $"![Local]({windowsPath})";

            var document = global::OfficeIMO.Markdown.MarkdownDocument.Parse(markdown);
            var image = Assert.IsType<global::OfficeIMO.Markdown.ImageBlock>(Assert.Single(document.Blocks));

            Assert.Equal(windowsPath, image.Path);

            var blocked = global::OfficeIMO.Markdown.MarkdownDocument.Parse(markdown, new global::OfficeIMO.Markdown.MarkdownLoadOptions {
                DisallowFileUrls = true
            });

            Assert.DoesNotContain(blocked.Blocks, block => block is global::OfficeIMO.Markdown.ImageBlock);
        }

        [Fact]
        public void MarkdownToWord_ParsesImageHints() {
            string imagePath = Path.GetFullPath(Path.Combine(AppContext.BaseDirectory, "..", "..", "..", "..", "Assets", "OfficeIMO.png"));
            const string remoteUrl = "https://images.example.test/logo.png";
            string md = $"![Local]({imagePath}){{width=40 height=30}}\n" +
                        $"![Remote]({remoteUrl}){{width=50 height=20}}";
            var warnings = new List<string>();
            using var doc = OfficeIMO.Markdown.MarkdownDocument.Parse(md).ToWord(new WordMarkdownImportOptions {
                AllowLocalImages = true,
                RemoteImageResolver = uri => uri.AbsoluteUri == remoteUrl ? File.ReadAllBytes(imagePath) : null,
                OnWarning = warnings.Add
            });

            Assert.Equal(2, doc.Images.Count);
            Assert.Empty(warnings);

            string dataUri = "data:image/png;base64," + Convert.ToBase64String(File.ReadAllBytes(imagePath));
            var sizedDocuments = new[] {
                OfficeIMO.Markdown.MarkdownDocument.Parse($"![Badge]({dataUri}){{width=96 height=48}}"),
                OfficeIMO.Markdown.MarkdownDocument.Parse($"![Badge]({dataUri}){{width=96px height=48px}}"),
                OfficeIMO.Markdown.MarkdownDocument.Parse($"![Badge]({dataUri}){{width=1in height=36pt}}"),
                OfficeIMO.Markdown.MarkdownDocument.Create().Image(dataUri, "Badge",
                    width: OfficeIMO.OfficeLength.Inches(1), height: OfficeIMO.OfficeLength.Inches(0.5)),
                OfficeIMO.Html.HtmlConversionDocument.Parse($"<figure><img src=\"{dataUri}\" width=\"96\" height=\"48\" /></figure>").ToMarkdown(),
                OfficeIMO.Html.HtmlConversionDocument.Parse($"<figure><img src=\"{dataUri}\" style=\"width:1in;height:36pt\" /></figure>").ToMarkdown()
            };
            foreach (var sized in sizedDocuments) {
                using (sized) {
                    var image = Assert.IsType<OfficeIMO.Markdown.ImageBlock>(Assert.Single(sized.Blocks));
                    Assert.Equal(OfficeIMO.OfficeLength.Inches(1), image.Width);
                    Assert.Equal(OfficeIMO.OfficeLength.Inches(0.5), image.Height);
                    Assert.Contains("{width=96 height=48}", sized.ToMarkdown(), StringComparison.Ordinal);
                    string html = sized.ToHtml(new OfficeIMO.Markdown.MarkdownHtmlExportOptions {
                        Kind = OfficeIMO.Markdown.MarkdownHtmlKind.Fragment
                    });
                    Assert.Contains("width=\"96\"", html, StringComparison.Ordinal);
                    Assert.Contains("height=\"48\"", html, StringComparison.Ordinal);
                    var native = OfficeIMO.Markdown.MarkdownNativeDocument.Parse(sized.ToMarkdown());
                    var nativeImage = Assert.IsType<OfficeIMO.Markdown.MarkdownNativeImageBlock>(Assert.Single(native.Blocks));
                    Assert.Equal(OfficeIMO.OfficeLength.Inches(1), nativeImage.Width);
                    Assert.Equal(OfficeIMO.OfficeLength.Inches(0.5), nativeImage.Height);
                    Assert.Contains("{width=96 height=48}", Assert.Single(native.ToSnapshot().Blocks).Markdown!, StringComparison.Ordinal);

                    using var word = sized.ToWord();
                    using var wordBytes = new MemoryStream(word.ToBytes());
                    using var reopenedWord = WordDocument.Load(wordBytes);
                    var wordImage = Assert.Single(reopenedWord.Images);
                    Assert.Equal(72d, wordImage.Width!.Value.ToPoints(), 4);
                    Assert.Equal(36d, wordImage.Height!.Value.ToPoints(), 4);

                    using var pdf = sized.ToPdf();
                    using var pdfBytes = new MemoryStream(pdf.ToBytes());
                    using var reopenedPdf = OfficeIMO.Pdf.PdfDocument.Load(pdfBytes);
                    var placement = Assert.Single(reopenedPdf.Images.Placements());
                    Assert.Equal(72d, placement.Width, 4);
                    Assert.Equal(36d, placement.Height, 4);
                }
            }
            using var fractional = OfficeIMO.Markdown.MarkdownDocument.Create().Image(dataUri, "Fractional",
                width: OfficeIMO.OfficeLength.Emus(919163), height: OfficeIMO.OfficeLength.Emus(459581));
            foreach (var rawHtmlHandling in new[] {
                OfficeIMO.Markdown.MarkdownRawHtmlHandling.Allow,
                OfficeIMO.Markdown.MarkdownRawHtmlHandling.Strip
            }) {
                string fractionalHtml = fractional.ToHtml(new OfficeIMO.Markdown.MarkdownHtmlExportOptions {
                    Kind = OfficeIMO.Markdown.MarkdownHtmlKind.Fragment,
                    RawHtmlHandling = rawHtmlHandling
                });
                Assert.Contains("style=\"", fractionalHtml, StringComparison.Ordinal);
                using var imported = OfficeIMO.Html.HtmlConversionDocument.Parse("<figure>" + fractionalHtml + "</figure>").ToMarkdown();
                var importedImage = Assert.IsType<OfficeIMO.Markdown.ImageBlock>(Assert.Single(imported.Blocks));
                Assert.Equal(OfficeIMO.OfficeLength.Emus(919163), importedImage.Width);
                Assert.Equal(OfficeIMO.OfficeLength.Emus(459581), importedImage.Height);
            }
        }

        [Fact]
        public void MarkdownToWord_Result_Reports_Unresolved_Remote_Images() {
            global::OfficeIMO.Markdown.MarkdownDocument markdown = global::OfficeIMO.Markdown.MarkdownDocument.Parse(
                "![Remote](https://images.example.test/logo.png)");

            MarkdownToWordResult result = markdown.ConvertToWordWithEvidence();
            using WordDocument document = result.Value;

            Assert.True(result.Succeeded);
            Assert.True(result.HasLoss);
            Assert.Single(document.HyperLinks);
            Assert.Contains(result.Report.Diagnostics, diagnostic =>
                diagnostic.Code == "markdown-to-word-warning" &&
                diagnostic.Message.Contains("RemoteImageResolver", StringComparison.Ordinal));
            Assert.Throws<WordMarkdownConversionException>(() => result.RequireNoLoss());
        }

        [Fact]
        public void MarkdownToWord_UsesNaturalSizeWhenNoHints() {
            string imagePath = Path.GetFullPath(Path.Combine(AppContext.BaseDirectory, "..", "..", "..", "..", "Assets", "OfficeIMO.png"));
            string md = $"![Local]({imagePath})";
            var doc = OfficeIMO.Markdown.MarkdownDocument.Parse(md).ToWord(new WordMarkdownImportOptions {
                AllowLocalImages = true
            });

            var image = OfficeIMO.Drawing.OfficeImageReader.Identify(imagePath);

            Assert.Single(doc.Images);
            Assert.Equal("Local", doc.Images[0].Description);
            Assert.Equal(image.Width, (doc.Images[0].Width?.ToPixels()));
            Assert.Equal(image.Height, (doc.Images[0].Height?.ToPixels()));

        }

        [Fact]
        public void MarkdownToWord_RestoresDataUriImagesExportedByWordToMarkdown() {
            string imagePath = Path.GetFullPath(Path.Combine(AppContext.BaseDirectory, "..", "..", "..", "..", "Assets", "OfficeIMO.png"));
            using var source = WordDocument.Create();
            source.AddParagraph("Before ");
            source.AddParagraph().AddImage(imagePath,options:new OfficeIMO.Word.WordImageOptions { Description = "Logo" });
            source.AddParagraph("After");

            string markdown = ((source.ToMarkdown(new WordMarkdownExportOptions()).ToMarkdown()));
            using var restored = OfficeIMO.Markdown.MarkdownDocument.Parse(markdown).ToWord();

            Assert.Contains("data:image/png;base64", markdown, StringComparison.Ordinal);
            Assert.Single(restored.Images);
            Assert.Equal("Logo", restored.Images[0].Description);
        }

        [Fact]
        public void MarkdownToWord_Rejects_Oversized_DataUri_Before_Decoding() {
            var warnings = new List<string>();
            const string markdown = "![Too large](data:image/png;base64,AAAA)";

            using var doc = OfficeIMO.Markdown.MarkdownDocument.Parse(markdown).ToWord(new WordMarkdownImportOptions {
                MaxDataUriImageBytes = 1,
                OnWarning = warnings.Add
            });

            Assert.Empty(doc.Images);
            Assert.Contains(warnings, warning => warning.Contains("exceeding the configured limit", StringComparison.OrdinalIgnoreCase));
        }

        [Fact]
        public void MarkdownToWord_Invalid_DataUri_Image_Falls_Back_With_Warning() {
            var warnings = new List<string>();
            const string markdown = "![Bad image](data:image/png;base64,AAAA)";

            using var doc = OfficeIMO.Markdown.MarkdownDocument.Parse(markdown).ToWord(new WordMarkdownImportOptions {
                OnWarning = warnings.Add
            });

            Assert.Empty(doc.Images);
            Assert.Contains(doc.Paragraphs, paragraph => paragraph.Text.Contains("Bad image", StringComparison.Ordinal));
            Assert.Contains(warnings, warning => warning.Contains("could not be inserted", StringComparison.OrdinalIgnoreCase));
        }

        [Fact]
        public void MarkdownToWord_RendersInlineLocalImagesInsideParagraphs() {
            string imagePath = Path.GetFullPath(Path.Combine(AppContext.BaseDirectory, "..", "..", "..", "..", "Assets", "OfficeIMO.png"));
            string md = $"Before ![Logo]({imagePath}) after";

            using var doc = OfficeIMO.Markdown.MarkdownDocument.Parse(md).ToWord(new WordMarkdownImportOptions {
                AllowLocalImages = true
            });

            string roundTrip = ((doc.ToMarkdown(new WordMarkdownExportOptions()).ToMarkdown()));

            Assert.Single(doc.Images);
            Assert.Contains("Before", roundTrip, StringComparison.Ordinal);
            Assert.Contains("![Logo](data:image/png;base64", roundTrip, StringComparison.Ordinal);
            Assert.Contains("after", roundTrip, StringComparison.Ordinal);
        }

        [Fact]
        public void MarkdownToWord_FitsImageToPageContentWidthWhenEnabled() {
            string imagePath = Path.GetFullPath(Path.Combine(AppContext.BaseDirectory, "..", "..", "..", "..", "Assets", "OfficeIMO.png"));
            string md = $"![Local]({imagePath}){{width=1200 height=300}}";
            var doc = OfficeIMO.Markdown.MarkdownDocument.Parse(md).ToWord(new WordMarkdownImportOptions {
                AllowLocalImages = true,
                FitImagesToPageContentWidth = true,
                ImageLayout = {
                    AllowUpscale = true
                },
                DefaultPageSize = WordPageSize.Letter
            });

            Assert.Single(doc.Images);
            Assert.InRange((doc.Images[0].Width?.ToPixels()) ?? 0, 623, 625);
            Assert.InRange((doc.Images[0].Height?.ToPixels()) ?? 0, 155, 157);
        }

        [Fact]
        public void MarkdownToWord_FitsInlineListImagesToIndentedContextWidth() {
            string imagePath = Path.GetFullPath(Path.Combine(AppContext.BaseDirectory, "..", "..", "..", "..", "Assets", "OfficeIMO.png"));
            string md = $"- ![Local]({imagePath}){{width=1200 height=300}}";
            var doc = OfficeIMO.Markdown.MarkdownDocument.Parse(md).ToWord(new WordMarkdownImportOptions {
                AllowLocalImages = true,
                FitImagesToContextWidth = true,
                ImageLayout = {
                    AllowUpscale = true
                },
                DefaultPageSize = WordPageSize.Letter
            });

            Assert.Single(doc.Images);
            Assert.InRange((doc.Images[0].Width?.ToPixels()) ?? 0, 488, 490);
        }

        [Fact]
        public void MarkdownToWord_AppliesConfiguredImageMaxWidth() {
            string imagePath = Path.GetFullPath(Path.Combine(AppContext.BaseDirectory, "..", "..", "..", "..", "Assets", "OfficeIMO.png"));
            string md = $"![Local]({imagePath}){{width=1200 height=300}}";
            var doc = OfficeIMO.Markdown.MarkdownDocument.Parse(md).ToWord(new WordMarkdownImportOptions {
                AllowLocalImages = true,
                MaxImageWidth = global::OfficeIMO.OfficeLength.Inches(5)
            });

            Assert.Single(doc.Images);
            Assert.InRange((doc.Images[0].Width?.ToPixels()) ?? 0, 479.5, 480.5);
            Assert.InRange((doc.Images[0].Height?.ToPixels()) ?? 0, 119.5, 120.5);
        }

        [Fact]
        public void MarkdownToWord_AppliesConfiguredImageMaxWidthPercentOfContent() {
            string imagePath = Path.GetFullPath(Path.Combine(AppContext.BaseDirectory, "..", "..", "..", "..", "Assets", "OfficeIMO.png"));
            string md = $"![Local]({imagePath}){{width=1200 height=300}}";
            var doc = OfficeIMO.Markdown.MarkdownDocument.Parse(md).ToWord(new WordMarkdownImportOptions {
                AllowLocalImages = true,
                DefaultPageSize = WordPageSize.Letter,
                MaxImageWidthPercentOfContent = 50,
                ImageLayout = {
                    AllowUpscale = true
                }
            });

            Assert.Single(doc.Images);
            Assert.InRange((doc.Images[0].Width?.ToPixels()) ?? 0, 311, 313);
            Assert.InRange((doc.Images[0].Height?.ToPixels()) ?? 0, 77, 79);
        }

        [Fact]
        public void MarkdownToWord_AppliesConfiguredImageMaxHeight() {
            string imagePath = Path.GetFullPath(Path.Combine(AppContext.BaseDirectory, "..", "..", "..", "..", "Assets", "OfficeIMO.png"));
            string md = $"![Local]({imagePath}){{width=1200 height=300}}";
            var doc = OfficeIMO.Markdown.MarkdownDocument.Parse(md).ToWord(new WordMarkdownImportOptions {
                AllowLocalImages = true,
                MaxImageHeight = global::OfficeIMO.OfficeLength.Points(75)
            });

            Assert.Single(doc.Images);
            Assert.InRange((doc.Images[0].Height?.ToPixels()) ?? 0, 99.5, 100.5);
            Assert.InRange((doc.Images[0].Width?.ToPixels()) ?? 0, 399.5, 400.5);
        }

        [Fact]
        public void MarkdownToWord_EmitsImageLayoutDiagnosticsWhenScaled() {
            string imagePath = Path.GetFullPath(Path.Combine(AppContext.BaseDirectory, "..", "..", "..", "..", "Assets", "OfficeIMO.png"));
            string md = $"![Local]({imagePath}){{width=1200 height=300}}";
            var diagnostics = new List<WordMarkdownImageLayoutDiagnostic>();

            using var doc = OfficeIMO.Markdown.MarkdownDocument.Parse(md).ToWord(new WordMarkdownImportOptions {
                AllowLocalImages = true,
                MaxImageWidth = global::OfficeIMO.OfficeLength.Inches(5),
                OnImageLayoutDiagnostic = diagnostics.Add
            });

            Assert.Single(doc.Images);
            var diagnostic = Assert.Single(diagnostics);
            Assert.True(diagnostic.ScaledByLayout);
            Assert.Equal("block-local", diagnostic.Context);
            Assert.Equal(1200, diagnostic.RequestedWidth?.ToPixels());
            Assert.InRange(diagnostic.FinalWidth?.ToPixels() ?? 0, 479.5, 480.5);

            Assert.Equal(imagePath, diagnostic.Source);
            Assert.Equal(300, diagnostic.RequestedHeight?.ToPixels());
            Assert.Equal(489, diagnostic.NaturalWidth?.ToPixels());
            Assert.Equal(436, diagnostic.NaturalHeight?.ToPixels());
            Assert.Equal(480, diagnostic.EffectiveMaxWidth?.ToPixels());
            Assert.Equal(436, diagnostic.EffectiveMaxHeight?.ToPixels());
            Assert.Equal(120, diagnostic.FinalHeight?.ToPixels());

            using var savedBytes = new MemoryStream(doc.ToBytes());
            using var reopened = WordDocument.Load(savedBytes);
            var savedImage = Assert.Single(reopened.Images);
            Assert.Equal(360D, savedImage.Width!.Value.ToPoints(), 4);
            Assert.Equal(90D, savedImage.Height!.Value.ToPoints(), 4);
        }

        [Fact]
        public void MarkdownToWord_FitImagesToPageContentWidth_ForcesPageMode() {
            var options = new WordMarkdownImportOptions();
            options.FitImagesToContextWidth = true;

            options.FitImagesToPageContentWidth = true;

            Assert.True(options.FitImagesToPageContentWidth);
            Assert.False(options.FitImagesToContextWidth);
            Assert.Equal(WordMarkdownImageFitMode.PageContentWidth, options.ImageLayout.FitMode);
        }

        [Fact]
        public void WordToMarkdown_WritesImageDescription() {
            using var doc = WordDocument.Create();
            string imagePath = Path.GetFullPath(Path.Combine(AppContext.BaseDirectory, "..", "..", "..", "..", "Assets", "OfficeIMO.png"));
            doc.AddParagraph().AddImage(imagePath);
            doc.Images[0].Description = "Sample";

            string markdown = ((doc.ToMarkdown(new WordMarkdownExportOptions()).ToMarkdown()));

            Assert.Contains("![Sample]", markdown);
        }

    }
}
