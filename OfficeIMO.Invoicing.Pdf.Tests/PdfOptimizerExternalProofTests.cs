#if NET8_0_OR_GREATER
using System.Security.Cryptography;
using System.Security.Cryptography.X509Certificates;
using System.Text.Json;
using System.Text.Json.Serialization;
using OfficeIMO.Invoicing.Tests;
using OfficeIMO.Pdf;
using OfficeIMO.Security;
using OfficeIMO.Tests;
using OfficeIMO.Tests.Pdf;
using Xunit;

namespace OfficeIMO.Invoicing.Pdf.Tests;

/// <summary>Exports actual optimizer inputs, outputs and settings for independent compliance checks.</summary>
public sealed class PdfOptimizerExternalProofTests {
    private static readonly JsonSerializerOptions JsonOptions = new() {
        WriteIndented = true,
        PropertyNamingPolicy = JsonNamingPolicy.CamelCase,
        Converters = { new JsonStringEnumConverter() }
    };

    [Fact]
    public void ExportEveryProfileAndCustomRewriteForExactArtifactValidation() {
        string? configured = Environment.GetEnvironmentVariable("OFFICEIMO_PDF_COMPLIANCE_PROOF_OUTPUT");
        string directory = string.IsNullOrWhiteSpace(configured)
            ? Path.Combine(Path.GetTempPath(), "OfficeIMO-optimizer-proof-" + Guid.NewGuid().ToString("N"))
            : Path.Combine(configured, "optimizer");
        Directory.CreateDirectory(directory);
        var rows = new List<object>();
        var refusals = new List<object>();
        try {
            foreach (string source in new[] { "pdfa2b", "invoice-pdfa3b-ua1", "pdfa4", "pdfua2", "pdfa1b" }) {
                byte[] before = CreateSource(source, out byte[]? xml);
                string input = source + "-before.pdf";
                File.WriteAllBytes(Path.Combine(directory, input), before);
                if (xml is not null) File.WriteAllBytes(Path.Combine(directory, source + "-before.xml"), xml);
                foreach (var (setting, options) in Settings()) {
                    string output = source + "-" + setting + "-after.pdf";
                    if (source == "pdfa1b" && (options.UseObjectStreams || options.XrefFormat == PdfOptimizationXrefFormat.XrefStream)) {
                        refusals.Add(AssertRefused(directory, input, before, setting, options, "pdfa1-streams"));
                        continue;
                    }
                    PdfOptimizationActionResult result = PdfDocument.Load(before).Optimization.Apply(options);
                    byte[] after = result.Bytes;
                    // Persist returned bytes directly: another save must not repair or alter the candidate.
                    File.WriteAllBytes(Path.Combine(directory, output), after);
                    if (xml is not null) {
                        File.WriteAllBytes(Path.Combine(directory, Path.ChangeExtension(output, ".xml")), xml);
                        Assert.Equal(xml, Assert.Single(PdfDocument.Load(after).Attachments.Extract()).Bytes);
                    }
                    Assert.Equal(before.LongLength, result.OriginalLengthBytes);
                    Assert.Equal(after.LongLength, result.OptimizedLengthBytes);
                    Assert.Equal(options.Profile, result.RequestedProfile);
                    if (!options.KeepOriginalWhenNotSmaller) Assert.False(result.ReturnedOriginal);
                    if (result.ReturnedOriginal) Assert.Equal(before, after);
                    rows.Add(new {
                        source, setting, options = Snapshot(options), input, output,
                        result.OriginalLengthBytes, result.OptimizedLengthBytes, result.CandidateLengthBytes,
                        sizeDeltaBytes = after.LongLength - before.LongLength,
                        result.ReturnedOriginal, result.CandidateXrefFormat, result.CandidateUsesObjectStreams,
                        result.CandidateLinearized, result.Actions, result.SkippedActions
                    });
                }
            }
            byte[] signed = CreateSignedSource();
            const string signedInput = "signed-before.pdf";
            File.WriteAllBytes(Path.Combine(directory, signedInput), signed);
            foreach (var (setting, options) in Settings())
                refusals.Add(AssertRefused(directory, signedInput, signed, setting, options, "signed"));
            File.WriteAllText(Path.Combine(directory, "measurements.json"), JsonSerializer.Serialize(new {
                schemaVersion = 1, profiles = Enum.GetNames<PdfOptimizationProfile>(), rows, refusals
            }, JsonOptions) + "\n");
        } finally {
            if (string.IsNullOrWhiteSpace(configured)) Directory.Delete(directory, recursive: true);
        }
    }

    private static IEnumerable<(string Name, PdfOptimizationOptions Options)> Settings() {
        foreach (PdfOptimizationProfile profile in Enum.GetValues<PdfOptimizationProfile>())
            yield return (profile.ToString(), PdfOptimizationOptions.Create(profile));
        yield return ("CustomRewrite", new PdfOptimizationOptions {
            Profile = PdfOptimizationProfile.Custom, KeepOriginalWhenNotSmaller = false
        });
        yield return ("CustomXrefStream", new PdfOptimizationOptions {
            Profile = PdfOptimizationProfile.Custom, KeepOriginalWhenNotSmaller = false,
            XrefFormat = PdfOptimizationXrefFormat.XrefStream
        });
        yield return ("CustomObjectStreams", new PdfOptimizationOptions {
            Profile = PdfOptimizationProfile.Custom, KeepOriginalWhenNotSmaller = false, UseObjectStreams = true
        });
    }

    private static object AssertRefused(string directory, string input, byte[] before, string setting,
        PdfOptimizationOptions options, string reason) {
        string name = Path.GetFileNameWithoutExtension(input) + "-" + setting;
        string absent = Path.Combine(directory, name + "-must-not-exist.pdf");
        string existing = Path.Combine(directory, name + "-sentinel.bin");
        Assert.False(File.Exists(absent), "Refusal output must be absent before invoking the optimizer.");
        byte[] sentinel = "Existing destination must survive optimizer refusal"u8.ToArray();
        File.WriteAllBytes(existing, sentinel);
        Exception? missingError = Record.Exception(() => PdfDocument.Load(before).Optimization.Apply(options).ToDocument().Save(absent));
        Exception? existingError = Record.Exception(() => PdfDocument.Load(before).Optimization.Apply(options).ToDocument().Save(existing));
        Assert.IsAssignableFrom<NotSupportedException>(missingError);
        Assert.IsAssignableFrom<NotSupportedException>(existingError);
        Assert.False(File.Exists(absent));
        Assert.Equal(sentinel, File.ReadAllBytes(existing));
        Assert.Equal(before, File.ReadAllBytes(Path.Combine(directory, input)));
        return new {
            input, setting, options = Snapshot(options), reason, outputAbsent = Path.GetFileName(absent),
            existingOutput = Path.GetFileName(existing), existingOutputUnchanged = true, sourceUnchanged = true,
            exceptionType = missingError!.GetType().Name, diagnostic = missingError.Message
        };
    }

    private static Dictionary<string, object?> Snapshot(PdfOptimizationOptions options) =>
        typeof(PdfOptimizationOptions).GetProperties()
            .Where(property => property.Name != nameof(PdfOptimizationOptions.CancellationToken))
            .ToDictionary(property => JsonNamingPolicy.CamelCase.ConvertName(property.Name), property => property.GetValue(options));

    private static byte[] CreateSource(string source, out byte[]? xml) {
        xml = null;
        if (source == "pdfa1b") return File.ReadAllBytes(Path.Combine(RepositoryTestPaths.Find(),
            "OfficeIMO.Pdf.Tests/Pdf/Fixtures/Interoperability/openpreserve-pdfa1b-text.pdf"));
        byte[] font = File.ReadAllBytes(PdfComplianceTestFonts.FindBundledOpenTypeCffFont()!);
        var options = new PdfOptions { IncludeStandardFontToUnicodeMaps = true };
        options.EmbedStandardFont(PdfStandardFont.Helvetica, font, "OfficeIMO Source Serif")
            .EmbedStandardFont(PdfStandardFont.HelveticaBold, font, "OfficeIMO Source Serif");
        switch (source) {
            case "pdfa2b": options.UsePdfA(PdfComplianceProfile.PdfA2B); break;
            case "invoice-pdfa3b-ua1":
                options.UsePdfA(PdfComplianceProfile.PdfA3B).UsePdfUa();
                var contract = new InvoiceXmlOptions(InvoiceSpecificationRelease.FacturX_1_09_2_Zugferd_2_5_2,
                    InvoiceSyntax.Cii, InvoiceProfile.En16931);
                PdfInvoiceDocument invoice = PdfInvoiceDocument.Create(InvoiceFixture.Create(), contract);
                xml = invoice.ToXmlBytes();
                return invoice.ToPdfBytes(options);
            case "pdfa4": options.UsePdfA(PdfComplianceProfile.PdfA4); break;
            case "pdfua2": options.UsePdfUa(PdfComplianceProfile.PdfUa2); break;
            default: throw new ArgumentOutOfRangeException(nameof(source));
        }
        PdfDocument document = PdfDocument.Create(options).Meta(title: "Optimizer compliance proof", author: "OfficeIMO");
        document.Content.H1("Optimizer compliance proof")
            .Paragraph(p => p.Text("Embedded Unicode text must survive every optimizer setting."))
            .Table(new[] { new[] { "Item", "Amount" }, new[] { "Service", "119.00 EUR" } });
        return document.ToBytes();
    }

    private static byte[] CreateSignedSource() {
        using RSA key = RSA.Create(2048);
        var request = new CertificateRequest("CN=OfficeIMO optimizer proof", key, HashAlgorithmName.SHA256,
            RSASignaturePadding.Pkcs1);
        using X509Certificate2 certificate = request.CreateSelfSigned(DateTimeOffset.UtcNow.AddMinutes(-5),
            DateTimeOffset.UtcNow.AddDays(1));
        using var signer = new PdfCmsExternalSigner(OfficeSecurityProvider.Default, certificate);
        PdfDocument document = PdfDocument.Create();
        document.Content.Paragraph(p => p.Text("Signed optimizer input must be refused"));
        return document.Security.SignExternal(signer, new PdfExternalSignatureOptions {
                FieldName = "OptimizerProof", ReservedSignatureContentsBytes = 8192
            }).ToDocument().ToBytes();
    }
}
#endif
