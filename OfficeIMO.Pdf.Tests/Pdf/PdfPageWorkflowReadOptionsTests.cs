using OfficeIMO.Pdf;
using Xunit;

namespace OfficeIMO.Tests.Pdf;

public class PdfPageWorkflowReadOptionsTests {
    [Fact]
    public void CropAndTranslateUsesThePasswordCapturedWhenLoadingTheSource() {
        byte[] source = PdfDocument.Create().Paragraph(p => p.Text("Protected crop content")).ToBytes();
        byte[] encrypted = PdfDocument.Load(source).Security.Encrypt(new PdfStandardEncryptionOptions("reader") {
            OwnerPassword = "owner", Algorithm = PdfStandardEncryptionAlgorithm.Aes128
        }).Pdf;
        var options = new PdfLoadOptions {
            Password = "owner",
            Limits = new PdfReadLimits { MaxJavaScripts = PdfReadLimits.DefaultMaxJavaScripts + 7 }
        };
        PdfDocument document = PdfDocument.Load(encrypted, options);

        PdfDocument cropped = document.Pages.CropAndTranslate(20, 20, 400, 700);

        Assert.Equal(380, cropped.Inspect().Pages[0].MediaBox!.Width);
        Assert.Equal(680, cropped.Inspect().Pages[0].MediaBox!.Height);
        Assert.Contains("Protected crop content", cropped.Read().Text, StringComparison.Ordinal);
        Assert.Equal(PdfReadLimits.DefaultMaxJavaScripts + 7, cropped.ReadOptions.Limits.MaxJavaScripts);
        Assert.Equal(encrypted, document.ToBytes());
    }
}
