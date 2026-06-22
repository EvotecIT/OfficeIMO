using OfficeIMO.Pdf;
using Xunit;

namespace OfficeIMO.Tests.Pdf;

public class PdfEncryptedWriteTests {
    [Fact]
    public void GeneratedEncryptedPdfRequiresValidPassword() {
        byte[] pdf = PdfDocument.Create(new PdfOptions().SetEncryption("open", "owner"))
            .Paragraph(paragraph => paragraph.Text("Generated Secret PDF Text"))
            .ToBytes();

        PdfDocumentPreflight preflight = PdfInspector.Preflight(pdf);

        Assert.True(PdfInspector.Probe(pdf).HasEncryption);
        Assert.False(preflight.CanRead);
        Assert.Contains(preflight.ReadBlockers, blocker => blocker.Kind == PdfReadBlockerKind.Encryption);
        Assert.Throws<PdfPasswordRequiredException>(() => PdfReadDocument.Load(pdf));
        Assert.Throws<PdfInvalidPasswordException>(() => PdfReadDocument.Load(pdf, new PdfReadOptions { Password = "wrong" }));
    }

    [Fact]
    public void GeneratedEncryptedPdfReadsTextWithUserPassword() {
        byte[] pdf = PdfDocument.Create(new PdfOptions().SetEncryption("open", "owner"))
            .Paragraph(paragraph => paragraph.Text("Generated Secret PDF Text"))
            .ToBytes();

        var readOptions = new PdfReadOptions { Password = "open" };
        PdfDocumentPreflight preflight = PdfInspector.Preflight(pdf, readOptions);
        string text = PdfTextExtractor.ExtractAllText(pdf, (PdfTextLayoutOptions?)null, readOptions);
        PdfDocument opened = PdfDocument.Open(pdf, readOptions);
        string fluentText = opened.Read.Text();
        PdfOperationResult<string> tryText = opened.Read.TryText();

        Assert.True(preflight.CanRead);
        Assert.False(preflight.CanRewrite);
        Assert.Equal("Standard", preflight.Probe.Security.EncryptionFilter);
        Assert.Equal(3, preflight.Probe.Security.EncryptionRevision);
        Assert.Equal(128, preflight.Probe.Security.EncryptionLengthBits);
        Assert.Contains("Generated Secret PDF Text", text, StringComparison.Ordinal);
        Assert.Contains("Generated Secret PDF Text", fluentText, StringComparison.Ordinal);
        Assert.True(tryText.Succeeded, string.Join(" ", tryText.Diagnostics));
        Assert.Contains("Generated Secret PDF Text", tryText.Value, StringComparison.Ordinal);
    }

    [Fact]
    public void GeneratedEncryptedPdfCanBeConfiguredThroughDocumentFluentApi() {
        byte[] pdf = PdfDocument.Create()
            .Encryption("open", "owner")
            .Paragraph(paragraph => paragraph.Text("Fluent Encryption Secret"))
            .ToBytes();

        string text = PdfTextExtractor.ExtractAllText(pdf, (PdfTextLayoutOptions?)null, new PdfReadOptions { Password = "open" });

        Assert.Contains("Fluent Encryption Secret", text, StringComparison.Ordinal);
    }
}
