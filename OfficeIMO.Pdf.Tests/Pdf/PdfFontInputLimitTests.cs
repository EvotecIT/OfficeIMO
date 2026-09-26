using OfficeIMO.Pdf;
using Xunit;

namespace OfficeIMO.Pdf.Tests.Pdf;

public class PdfFontInputLimitTests {
    [Fact]
    public void OversizedFontFileIsRejectedBeforeAuthoringOptionsChange() {
        string path = CreateOversizedFile();
        try {
            var options = new PdfOptions();
            Assert.Throws<InvalidDataException>(() => options.EmbedStandardFont(PdfStandardFont.Helvetica, path));
            Assert.Empty(options.EmbeddedFonts);
            Assert.Throws<InvalidDataException>(() => PdfEmbeddedFontFamily.FromFiles("Oversized", path));
            Assert.Throws<InvalidDataException>(() => options.UseFontFamily("Oversized", path));
            Assert.Empty(options.EmbeddedFonts);
            var formOptions = new PdfFormFillerOptions();
            Assert.Throws<InvalidDataException>(() => formOptions.UseAppearanceFontFile("Oversized", path));
            Assert.Null(formOptions.AppearanceFontFamily);
        } finally {
            File.Delete(path);
        }
    }

    [Fact]
    public void OversizedOptionalFontFaceIsRejectedBeforeRegistration() {
        string oversizedPath = CreateOversizedFile();
        string regularPath = Path.Combine(Path.GetTempPath(), "officeimo-font-small-" + Guid.NewGuid().ToString("N") + ".ttf");
        try {
            File.WriteAllBytes(regularPath, new byte[] { 1 });
            var options = new PdfOptions();
            Assert.Throws<InvalidDataException>(() => PdfEmbeddedFontFamily.FromFiles("Oversized Bold", regularPath, boldPath: oversizedPath));
            Assert.Throws<InvalidDataException>(() => options.UseFontFamily("Oversized Bold", regularPath, boldPath: oversizedPath));
            Assert.Empty(options.EmbeddedFonts);
        } finally {
            File.Delete(regularPath);
            File.Delete(oversizedPath);
        }
    }

    [Fact]
    public void OversizedFontBytesAreRejectedBeforeCopyOrRegistration() {
        byte[] oversized = new byte[128 * 1024 * 1024 + 1];
        var options = new PdfOptions();
        Assert.Throws<InvalidDataException>(() => new PdfEmbeddedFont(PdfStandardFont.Helvetica, oversized));
        Assert.Throws<InvalidDataException>(() => new PdfEmbeddedFontFamily("Oversized", oversized));
        Assert.Throws<InvalidDataException>(() => new PdfEmbeddedFontFamily("Oversized Bold", new byte[] { 1 }, bold: oversized));
        Assert.Throws<InvalidDataException>(() => options.EmbedStandardFont(PdfStandardFont.Helvetica, oversized));
        Assert.Empty(options.EmbeddedFonts);
    }

    private static string CreateOversizedFile() {
        string path = Path.Combine(Path.GetTempPath(), "officeimo-font-limit-" + Guid.NewGuid().ToString("N") + ".ttf");
        using (var stream = new FileStream(path, FileMode.CreateNew, FileAccess.Write, FileShare.None)) {
            stream.SetLength(PdfEmbeddedFontFamily.MaxSystemFontFileBytes + 1);
        }
        return path;
    }
}
