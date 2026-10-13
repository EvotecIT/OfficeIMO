using System;
using System.IO;
using System.Linq;
using OfficeIMO.Pdf;
using Xunit;

namespace OfficeIMO.Tests.Pdf;

public sealed class PdfCachedCffFontsTests {
    [Theory]
    [InlineData(true)]
    [InlineData(false)]
    public void NumericOpenTypeCacheFileUsesTheExistingCffFontProgram(bool directoryDiscovery) {
        string root = Path.Combine(Path.GetTempPath(), "OfficeIMO.CachedCff", Guid.NewGuid().ToString("N"));
        try {
            Directory.CreateDirectory(root);
            string path = Path.Combine(root, "123456789.otf");
            File.Copy(PdfComplianceTestFonts.FindBundledOpenTypeCffFont()!, path);
            if (directoryDiscovery) {
                Assert.Contains(path, PdfEmbeddedFontFamily.EnumerateTrueTypeFontFiles(root));
                return;
            }
            Assert.True(PdfEmbeddedFontFamily.TryFromSystemFontFiles("Source Serif 4", new[] { path }, out PdfEmbeddedFontFamily? family));
            Assert.NotNull(family);
            PdfOpenTypeCffFontProgram program = PdfFontProgramCache.GetOpenTypeCff(family!.Regular, null);
            Assert.True(program.MeasureTextWidth("Office fonts 0123456789", 11D) > 0D);
            byte[] pdf = PdfDocument.Create(new PdfOptions().UseFontFamily(family))
                .Paragraph(paragraph => paragraph.Text("Office fonts 0123456789")).ToBytes();
            using var reopened = UglyToad.PdfPig.PdfDocument.Open(pdf);
            Assert.Equal("Office fonts 0123456789", reopened.GetPage(1).Text);
        } finally {
            if (Directory.Exists(root)) Directory.Delete(root, recursive: true);
        }
    }
}
