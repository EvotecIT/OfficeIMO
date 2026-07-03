using System.Text.RegularExpressions;
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
    public void GeneratedEncryptedFormPdfBlocksFormMutationEvenWithPassword() {
        byte[] pdf = PdfDocument.Create(new PdfOptions().SetEncryption("open", "owner"))
            .TextField("Name", width: 180, height: 24, value: "Ada")
            .ToBytes();

        PdfDocumentPreflight preflight = PdfInspector.Preflight(pdf, new PdfReadOptions { Password = "open" });

        Assert.True(preflight.CanRead);
        Assert.False(preflight.CanFillSimpleFormFields);
        Assert.False(preflight.CanFlattenSimpleFormFields);
        Assert.False(preflight.Can(PdfPreflightCapability.FillSimpleFormFields));
        Assert.Contains(
            "Encrypted PDF files are not supported for form filling or flattening by OfficeIMO.Pdf yet.",
            preflight.GetCapabilityDiagnostics(PdfPreflightCapability.FillSimpleFormFields));
        Assert.Throws<NotSupportedException>(() => PdfFormFiller.FillFields(pdf, new Dictionary<string, string> {
            ["Name"] = "Grace"
        }));
    }

    [Fact]
    public void GeneratedEncryptedPdfReadsTextWithOwnerPasswordWhenUserPasswordUsesUtf8Fallback() {
        byte[] pdf = PdfDocument.Create(new PdfOptions().SetEncryption("open \ud83d\udd12", "owner"))
            .Paragraph(paragraph => paragraph.Text("Generated UTF8 Password Secret"))
            .ToBytes();

        string text = PdfTextExtractor.ExtractAllText(pdf, (PdfTextLayoutOptions?)null, new PdfReadOptions { Password = "owner" });

        Assert.Contains("Generated UTF8 Password Secret", text, StringComparison.Ordinal);
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

    [Fact]
    public void GeneratedEncryptedPdfReadsTextWithOwnerPasswordForWinAnsiUserPassword() {
        byte[] pdf = PdfDocument.Create(new PdfOptions().SetEncryption("€", "owner"))
            .Paragraph(paragraph => paragraph.Text("Owner Password Secret"))
            .ToBytes();

        string text = PdfTextExtractor.ExtractAllText(pdf, (PdfTextLayoutOptions?)null, new PdfReadOptions { Password = "owner" });

        Assert.Contains("Owner Password Secret", text, StringComparison.Ordinal);
    }

    [Fact]
    public void GeneratedEncryptedPdfReadsWhenEncryptReferenceHasNonZeroGeneration() {
        byte[] pdf = PdfDocument.Create(new PdfOptions().SetEncryption("open", "owner"))
            .Paragraph(paragraph => paragraph.Text("Generation Encryption Secret"))
            .ToBytes();
        string text = PdfEncoding.Latin1GetString(pdf);
        Match match = Regex.Match(text, @"/Encrypt\s+(\d+)\s+0\s+R");
        Assert.True(match.Success);
        string objectNumber = match.Groups[1].Value;
        text = text.Replace("/Encrypt " + objectNumber + " 0 R", "/Encrypt " + objectNumber + " 2 R")
            .Replace("\n" + objectNumber + " 0 obj", "\n" + objectNumber + " 2 obj");

        string extracted = PdfTextExtractor.ExtractAllText(PdfEncoding.Latin1GetBytes(text), (PdfTextLayoutOptions?)null, new PdfReadOptions { Password = "open" });

        Assert.Contains("Generation Encryption Secret", extracted, StringComparison.Ordinal);
    }
}
