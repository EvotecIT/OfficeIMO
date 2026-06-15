using System.Globalization;
using DocumentFormat.OpenXml.Wordprocessing;
using OfficeIMO.Word;
using Xunit;

namespace OfficeIMO.Tests {
    public partial class Word {
        [Fact]
        public void NormalizeColor_IsCultureInvariant() {
            var current = CultureInfo.CurrentCulture;
            var currentUi = CultureInfo.CurrentUICulture;
            try {
                CultureInfo.CurrentCulture = new CultureInfo("tr-TR");
                CultureInfo.CurrentUICulture = new CultureInfo("tr-TR");

                var normalized = Helpers.NormalizeColor("AAEEDD");
                Assert.Equal("aaeedd", normalized);
            } finally {
                CultureInfo.CurrentCulture = current;
                CultureInfo.CurrentUICulture = currentUi;
            }
        }

        [Fact]
        public void WatermarkParsing_IsCultureInvariant() {
            var current = CultureInfo.CurrentCulture;
            var currentUi = CultureInfo.CurrentUICulture;
            try {
                CultureInfo.CurrentCulture = new CultureInfo("tr-TR");
                CultureInfo.CurrentUICulture = new CultureInfo("tr-TR");

                using var document = WordDocument.Create();
                document.AddParagraph("Section 0");
                document.AddHeadersAndFooters();
                var header = RequireSectionHeader(document, 0, HeaderFooterValues.Default);
                var watermark = header.AddWatermark(WordWatermarkStyle.Text, "Watermark");

                Assert.Equal(90, watermark.Rotation);

                watermark.Width = 633.42;
                watermark.Height = 158.34;
                watermark.HorizontalOffset = 12.5;
                watermark.VerticalOffset = 24.75;
                watermark.FontSize = 72.25;

                Assert.Equal(633.42, watermark.Width);
                Assert.Equal(158.34, watermark.Height);
                Assert.Equal(12.5, watermark.HorizontalOffset);
                Assert.Equal(24.75, watermark.VerticalOffset);
                Assert.Equal(72.25, watermark.FontSize);

                watermark.Rotation = 135;
                Assert.Equal(135, watermark.Rotation);
            } finally {
                CultureInfo.CurrentCulture = current;
                CultureInfo.CurrentUICulture = currentUi;
            }
        }
    }
}

