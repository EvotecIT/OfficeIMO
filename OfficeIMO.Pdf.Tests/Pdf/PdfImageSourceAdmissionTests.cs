using OfficeIMO.Pdf;
using Xunit;

namespace OfficeIMO.Tests;

public sealed class PdfImageSourceAdmissionTests {
    [Theory]
    [InlineData(".png", "Unsupported OCR review source")]
    [InlineData(".svg", "<svg xmlns='http://www.w3.org/2000/svg' width='10' height='10'><rect width='10' height='10'/></svg>")]
    public void FileSnapshotRejectsNonRasterContentBeforeReturningAnOcrSource(string extension, string content) {
        string path = Path.Combine(Path.GetTempPath(), "unsupported-image-" + Guid.NewGuid().ToString("N") + extension);
        try {
            File.WriteAllText(path, content);
            Assert.Throws<NotSupportedException>(() => PdfImageDocumentSource.FromFile(path));
        } finally { File.Delete(path); }
    }
}
