using System.Globalization;
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

                var watermark = document.Sections[0].Header.Default.AddWatermark(WordWatermarkStyle.Text, "Watermark");

                Assert.Equal(90, watermark.Rotation);

                watermark.Rotation = 135;
                Assert.Equal(135, watermark.Rotation);
            } finally {
                CultureInfo.CurrentCulture = current;
                CultureInfo.CurrentUICulture = currentUi;
            }
        }
    }
}

