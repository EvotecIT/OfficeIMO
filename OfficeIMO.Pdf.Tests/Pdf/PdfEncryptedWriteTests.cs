using System.Text.RegularExpressions;
using OfficeIMO.Pdf;
using Xunit;

namespace OfficeIMO.Tests.Pdf;

public class PdfEncryptedWriteTests {
    [Fact]
    public void GeneratedEncryptionDefaultsToAes256Revision6() {
        byte[] pdf = PdfDocument.Create(new PdfOptions().SetEncryption("open", "owner"))
            .Paragraph(paragraph => paragraph.Text("AES default source"))
            .ToBytes();
        string raw = PdfEncoding.Latin1GetString(pdf);

        Assert.StartsWith("%PDF-2.0", raw, StringComparison.Ordinal);
        Assert.Contains("/V 5 /R 6 /Length 256", raw, StringComparison.Ordinal);
        Assert.Contains("/CFM /AESV3", raw, StringComparison.Ordinal);
        Assert.DoesNotContain("AES default source", raw, StringComparison.Ordinal);
        Assert.Contains(
            "AES default source",
            PdfTextExtractor.ExtractAllText(pdf, (PdfTextLayoutOptions?)null, new PdfReadOptions { Password = "open" }),
            StringComparison.Ordinal);
    }

    [Fact]
    public void GeneratedEncryptionSupportsExplicitAes128InteroperabilityMode() {
        var encryption = new PdfStandardEncryptionOptions("open") {
            OwnerPassword = "owner",
            Algorithm = PdfStandardEncryptionAlgorithm.Aes128
        };

        byte[] pdf = PdfDocument.Create(new PdfOptions().SetEncryption(encryption))
            .Paragraph(paragraph => paragraph.Text("AES-128 interoperability source"))
            .ToBytes();
        string raw = PdfEncoding.Latin1GetString(pdf);

        Assert.StartsWith("%PDF-1.6", raw, StringComparison.Ordinal);
        Assert.Contains("/V 4 /R 4 /Length 128", raw, StringComparison.Ordinal);
        Assert.Contains("/CFM /AESV2", raw, StringComparison.Ordinal);
        Assert.Contains(
            "AES-128 interoperability source",
            PdfTextExtractor.ExtractAllText(pdf, (PdfTextLayoutOptions?)null, new PdfReadOptions { Password = "owner" }),
            StringComparison.Ordinal);
    }

    [Fact]
    public void GeneratedEncryptionUsesRc4OnlyWhenExplicitlyRequested() {
        var encryption = new PdfStandardEncryptionOptions("open") {
            OwnerPassword = "owner",
            Algorithm = PdfStandardEncryptionAlgorithm.LegacyRc4
        };

        byte[] pdf = PdfDocument.Create(new PdfOptions().SetEncryption(encryption))
            .Paragraph(paragraph => paragraph.Text("Legacy RC4 source"))
            .ToBytes();
        string raw = PdfEncoding.Latin1GetString(pdf);

        Assert.StartsWith("%PDF-1.4", raw, StringComparison.Ordinal);
        Assert.Contains("/V 2 /R 3 /Length 128", raw, StringComparison.Ordinal);
        Assert.DoesNotContain("/AESV2", raw, StringComparison.Ordinal);
        Assert.Contains(
            "Legacy RC4 source",
            PdfTextExtractor.ExtractAllText(pdf, (PdfTextLayoutOptions?)null, new PdfReadOptions { Password = "open" }),
            StringComparison.Ordinal);
    }

    [Fact]
    public void GeneratedEncryptionMapsTypedPermissionsAndMetadataPolicy() {
        var encryption = new PdfStandardEncryptionOptions("open") {
            AllowedPermissions = PdfStandardPermissions.Print | PdfStandardPermissions.FillForms,
            EncryptMetadata = false
        };
        byte[] pdf = PdfDocument.Create(new PdfOptions {
                IncludeXmpMetadata = true
            }.SetEncryption(encryption))
            .Meta(title: "Visible XMP policy")
            .Paragraph(paragraph => paragraph.Text("Permission source"))
            .ToBytes();

        PdfDocumentInfo info = PdfInspector.Inspect(pdf, new PdfReadOptions { Password = "open" });

        Assert.Equal(PdfStandardPermissions.Print | PdfStandardPermissions.FillForms, info.Security.AllowedStandardPermissions);
        Assert.True(info.Security.AllowsPrinting);
        Assert.True(info.Security.AllowsFormFilling);
        Assert.False(info.Security.AllowsCopying);
        Assert.False(info.Security.EncryptMetadata);
        Assert.Equal("Visible XMP policy", info.Metadata.Title);
        Assert.NotNull(info.XmpMetadata);
    }

    [Fact]
    public void GeneratedEncryptedPdfRequiresValidPassword() {
        byte[] pdf = PdfDocument.Create(new PdfOptions().SetEncryption("open", "owner"))
            .Paragraph(paragraph => paragraph.Text("Generated Secret PDF Text"))
            .ToBytes();

        PdfDocumentPreflight preflight = PdfInspector.Preflight(pdf);

        Assert.True(PdfInspector.Probe(pdf).HasEncryption);
        Assert.False(preflight.CanRead);
        Assert.Contains(preflight.ReadBlockers, blocker => blocker.Kind == PdfReadBlockerKind.Encryption);
        Assert.Throws<PdfPasswordRequiredException>(() => PdfReadDocument.Open(pdf));
        Assert.Throws<PdfInvalidPasswordException>(() => PdfReadDocument.Open(pdf, new PdfReadOptions { Password = "wrong" }));
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
        Assert.Equal(6, preflight.Probe.Security.EncryptionRevision);
        Assert.Equal(256, preflight.Probe.Security.EncryptionLengthBits);
        Assert.Contains("Generated Secret PDF Text", text, StringComparison.Ordinal);
        Assert.Contains("Generated Secret PDF Text", fluentText, StringComparison.Ordinal);
        Assert.True(tryText.Succeeded, string.Join(" ", tryText.Diagnostics));
        Assert.Contains("Generated Secret PDF Text", tryText.Value, StringComparison.Ordinal);
    }

    [Fact]
    public void GeneratedEncryptedFormPdfSupportsFormMutationWithPassword() {
        byte[] pdf = PdfDocument.Create(new PdfOptions().SetEncryption("open", "owner"))
            .TextField("Name", width: 180, height: 24, value: "Ada")
            .ToBytes();

        var readOptions = new PdfReadOptions { Password = "open" };
        PdfDocumentPreflight preflight = PdfInspector.Preflight(pdf, readOptions);
        PdfDocument filled = PdfDocument.Open(pdf, readOptions).Forms.Fill(new Dictionary<string, string> {
            ["Name"] = "Grace"
        });

        Assert.True(preflight.CanRead);
        Assert.True(preflight.CanFillSimpleFormFields);
        Assert.True(preflight.CanFlattenSimpleFormFields);
        Assert.True(preflight.Can(PdfPreflightCapability.FillSimpleFormFields));
        Assert.False(PdfInspector.Probe(filled.ToBytes()).HasEncryption);
        Assert.Equal("Grace", Assert.Single(filled.Inspect().FormFields).Value);
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
    public void Aes256PasswordsUseUnicodeCompatibilityNormalization() {
        const string composed = "café";
        const string decomposed = "cafe\u0301";
        byte[] pdf = PdfDocument.Create(new PdfOptions().SetEncryption(composed, "owner"))
            .Paragraph(paragraph => paragraph.Text("Normalized Unicode password"))
            .ToBytes();

        string text = PdfTextExtractor.ExtractAllText(
            pdf,
            (PdfTextLayoutOptions?)null,
            new PdfReadOptions { Password = decomposed });

        Assert.Contains("Normalized Unicode password", text, StringComparison.Ordinal);
    }

    [Fact]
    public void Aes256PasswordsApplyTheStandard127ByteLimit() {
        string sharedPrefix = new('a', 127);
        byte[] pdf = PdfDocument.Create(new PdfOptions().SetEncryption(sharedPrefix + "first", "owner"))
            .Paragraph(paragraph => paragraph.Text("Truncated Unicode password"))
            .ToBytes();

        string text = PdfTextExtractor.ExtractAllText(
            pdf,
            (PdfTextLayoutOptions?)null,
            new PdfReadOptions { Password = sharedPrefix + "second" });

        Assert.Contains("Truncated Unicode password", text, StringComparison.Ordinal);
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
