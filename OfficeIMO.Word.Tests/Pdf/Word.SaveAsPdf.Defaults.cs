using DocumentFormat.OpenXml.Wordprocessing;
using OfficeIMO.Drawing;
using OfficeIMO.Word.Pdf;
using OfficeIMO.Word;
using System.Globalization;
using System.IO;
using System.Text;
using System.Text.RegularExpressions;
using Xunit;

namespace OfficeIMO.Tests {
    public partial class Word {
        [Fact]
        public void SaveAsPdf_DefaultsDoNotEmbedDocumentNamedHostFontsOrAllowArbitraryResourceReads() {
            var options = new WordPdfSaveOptions();

            Assert.True(options.ResourcePolicy.AllowSystemFontEmbedding);
            Assert.False(options.ResourcePolicy.AllowDocumentFontEmbedding);
            Assert.False(options.ResourcePolicy.AllowLocalFileAccess);
            Assert.False(options.ResourcePolicy.AllowRemoteResourceResolution);
        }

        [Fact]
        public void SaveAsPdf_ShapingOnlyProfileDetectsLaterDirectFontConfiguration() {
            var options = new WordPdfSaveOptions()
                .UseRenderingProfile(new OfficeRenderingProfile("shaping-only"));

            Assert.False(options.HasExplicitPdfFontConfiguration);

            options.PdfOptions!.DefaultFontSize = 13;

            Assert.True(options.HasExplicitPdfFontConfiguration);
            Assert.True(options.CloneForConversion().HasExplicitPdfFontConfiguration);
        }

        [Fact]
        public void SaveAsPdf_InvalidRenderingProfileDoesNotCreatePdfOptions() {
            var options = new WordPdfSaveOptions();

            Assert.Throws<ArgumentNullException>(
                () => options.UseRenderingProfile(null!));
            Assert.Null(options.PdfOptions);

            Assert.Throws<ArgumentOutOfRangeException>(
                () => options.UseRenderingProfile(
                    new OfficeRenderingProfile("invalid-mode"),
                    (OfficeRenderingProfileApplyMode)999));
            Assert.Null(options.PdfOptions);
        }

        [Fact]
        public void SaveAsPdf_Uses_DefaultPageSettings() {
            string docPath = Path.Combine(_directoryWithFiles, "PdfDefaultSettings.docx");
            string pdfPath = Path.Combine(_directoryWithFiles, "PdfDefaultSettings.pdf");

            using (WordDocument document = WordDocument.Create(docPath)) {
                document.AddParagraph("Hello World");
                document.Save();
                document.SaveAsPdf(pdfPath, new WordPdfSaveOptions {
                    DefaultOrientation = OfficePageOrientation.Landscape,
                    DefaultPageSize = WordPageSize.A4
                });
            }

            Assert.True(File.Exists(pdfPath));

            string pdfContent = PdfOperatorSearchText.From(File.ReadAllBytes(pdfPath));
            Match mediaBox = Regex.Match(pdfContent, @"/MediaBox\s*\[\s*0\s+0\s+(?<w>[0-9\.]+)\s+(?<h>[0-9\.]+)\s*\]");
            Assert.True(mediaBox.Success, "MediaBox not found");
            double width = double.Parse(mediaBox.Groups["w"].Value, CultureInfo.InvariantCulture);
            double height = double.Parse(mediaBox.Groups["h"].Value, CultureInfo.InvariantCulture);
            Assert.True(width > height);
        }
    }
}
