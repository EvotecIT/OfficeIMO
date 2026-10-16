#if NET8_0_OR_GREATER
using System;
using System.IO;
using System.Linq;
using System.Security.Cryptography;
using System.Security.Cryptography.X509Certificates;
using System.Threading;
using System.Threading.Tasks;
using OfficeIMO.Pdf;
using OfficeIMO.Security;
using Org.BouncyCastle.Asn1;
using Org.BouncyCastle.Asn1.Cms;
using Org.BouncyCastle.Cms;
using Xunit;

namespace OfficeIMO.Tests.Pdf;

public sealed class PdfPadesSigningTests {
    [Theory]
    [InlineData(false, PdfExternalSignatureSubFilter.CadesDetached)]
    [InlineData(true, PdfExternalSignatureSubFilter.CadesDetached)]
    [InlineData(false, PdfExternalSignatureSubFilter.DetachedCms)]
    public void ExternalSignerHonorsPreparedSubfilterWithoutMutatingSharedOptions(bool ecdsa, PdfExternalSignatureSubFilter subfilter) {
        using X509Certificate2 certificate = CreateCertificate(ecdsa);
        var options = new CmsSigningOptions { IncludeSigningTime = true };
        using var signer = new PdfCmsExternalSigner(OfficeSecurityProvider.Default, certificate, signingOptions: options);
        PdfExternalSignaturePreparation preparation = Source().Security.PrepareExternalSignature(new PdfExternalSignatureOptions { SubFilter = subfilter });
        byte[] encoded = signer.Sign(new PdfExternalSignatureRequest(preparation));
        var cms = new CmsSignedData(new CmsProcessableByteArray(preparation.SignedContent), encoded);
        var attributes = Assert.Single(cms.GetSignerInfos().GetSigners()).SignedAttributes;
        if (subfilter == PdfExternalSignatureSubFilter.CadesDetached) {
            Assert.Null(attributes[CmsAttributes.SigningTime]);
            Assert.NotNull(attributes[new DerObjectIdentifier("1.2.840.113549.1.9.16.2.47")]);
        } else {
            Assert.NotNull(attributes[CmsAttributes.SigningTime]);
        }
        Assert.True(options.IncludeSigningTime);
        Assert.False(options.UsePadesBaseline);
        var report = preparation.Complete(encoded).ValidateSignatures(Provider());
        Assert.True(report.MathematicalSignaturesVerified);
        Assert.True(report.DigestVerified);
    }

    [Fact]
    public void CmsSignerRejectsDocumentTimestampRequest() {
        using X509Certificate2 certificate = CreateCertificate(false);
        using var signer = new PdfCmsExternalSigner(OfficeSecurityProvider.Default, certificate);
        var preparation = Source().Security.PrepareExternalSignature(new PdfExternalSignatureOptions { Profile = PdfSignatureProfile.DocumentTimestamp });
        Assert.Throws<NotSupportedException>(() => signer.Sign(new PdfExternalSignatureRequest(preparation)));
    }

    [Fact]
    public void VisibleTextAppearanceRequiresCallerSuppliedEmbeddedFont() {
        var exception = Assert.Throws<ArgumentException>(() => Source().Security.PrepareExternalSignature(
            new PdfExternalSignatureOptions { VisibleAppearance = new PdfVisibleSignatureAppearanceOptions { Text = "Approved" } }));
        Assert.Contains("VisibleAppearance.Font", exception.Message, StringComparison.Ordinal);
    }

    [Fact]
    public void VisibleUnicodeAppearanceEmbedsFontAndRetainsValidSignature() {
        using X509Certificate2 certificate = CreateCertificate(true);
        byte[] fontData = File.ReadAllBytes(Path.Combine(AppContext.BaseDirectory, "Typography", "Carlito-Regular.ttf"));
        var appearance = new PdfVisibleSignatureAppearanceOptions {
            Text = "Approved: Żółć €", Font = new PdfEmbeddedFont(PdfStandardFont.Helvetica, fontData), Width = 240
        };
        using var signer = new PdfCmsExternalSigner(OfficeSecurityProvider.Default, certificate);
        byte[] signed = Source().Security.SignExternal(signer, new PdfExternalSignatureOptions {
            SubFilter = PdfExternalSignatureSubFilter.CadesDetached, VisibleAppearance = appearance
        }).ToDocument().ToBytes();
        var objects = PdfSyntax.ParseObjects(signed).Map;
        var form = Assert.Single(objects.Values.Select(value => value.Value).OfType<PdfStream>(), stream =>
            stream.Dictionary.Items.TryGetValue("Subtype", out PdfObject? type) && type is PdfName { Name: "Form" });
        var resources = Assert.IsType<PdfDictionary>(form.Dictionary.Items["Resources"]);
        var fonts = Assert.IsType<PdfDictionary>(resources.Items["Font"]);
        var reference = Assert.IsType<PdfReference>(fonts.Items["F1"]);
        var font = Assert.IsType<PdfDictionary>(objects[reference.ObjectNumber].Value);
        Assert.Equal("Type0", Assert.IsType<PdfName>(font.Items["Subtype"]).Name);
        Assert.Contains(objects.Values.Select(value => value.Value).OfType<PdfDictionary>(), value => value.Items.ContainsKey("FontFile2"));
        Assert.True(PdfSignatureValidator.Validate(signed, Provider()).MathematicalSignaturesVerified);
    }

    [Fact]
    public async Task FetchHookMergesSuppliedEvidenceAndPreservesPriorDssRevision() {
        using X509Certificate2 certificate = CreateCertificate(false);
        using X509Certificate2 extra = CreateCertificate(true);
        using var signer = new PdfCmsExternalSigner(OfficeSecurityProvider.Default, certificate);
        byte[] signed = Source().Security.SignExternal(signer,
            new PdfExternalSignatureOptions { SubFilter = PdfExternalSignatureSubFilter.CadesDetached }).ToDocument().ToBytes();
        var provider = Provider();
        int signature = Assert.Single(PdfSignatureValidator.Validate(signed, provider).Signatures).Signature.ObjectNumber;
        var first = PdfLongTermValidationEnricher.Enrich(signed,
            new PdfLongTermValidationEvidence(signature, certificates: new[] { certificate.RawData }), provider);
        // A well-formed unsuccessful OCSP response is still caller-owned evidence, not a trust claim.
        byte[] ocsp = { 0x30, 0x03, 0x0A, 0x01, 0x06 };
        var supplied = new PdfLongTermValidationEvidence(signature, certificates: new[] { certificate.RawData },
            ocspResponses: new[] { ocsp });
        var second = await PdfLongTermValidationEnricher.EnrichAsync(first.Pdf, supplied, provider,
            (target, token) => {
                Assert.Equal(signature, target.Signature.ObjectNumber);
                return Task.FromResult(new PdfLongTermValidationEvidence(signature, certificates: new[] { extra.RawData }));
            });
        Assert.True(second.IsVerifiedAppendOnlyEnrichment);
        Assert.True(second.Pdf.AsSpan(0, first.Pdf.Length).SequenceEqual(first.Pdf));
        Assert.Equal(2, second.Evidence.Certificates.Count);
        var enrichedObjects = PdfSyntax.ParseObjects(second.Pdf).Map;
        Assert.Equal(ocsp, Assert.IsType<PdfStream>(enrichedObjects[Assert.Single(second.OcspObjectNumbers)].Value).Data);
        Assert.Contains(first.CertificateObjectNumbers[0], second.ValidationAfter.Security.DocumentSecurityStore.CertificateObjectNumbers);
        await Assert.ThrowsAsync<ArgumentException>(() => PdfLongTermValidationEnricher.EnrichAsync(signed, supplied, provider,
            (_, _) => Task.FromResult(new PdfLongTermValidationEvidence(signature + 1, certificates: new[] { extra.RawData }))));
        using var cancellation = new CancellationTokenSource();
        cancellation.Cancel();
        await Assert.ThrowsAnyAsync<OperationCanceledException>(() => PdfLongTermValidationEnricher.EnrichAsync(signed, supplied, provider,
            cancellationToken: cancellation.Token));
    }

    private static PdfDocument Source() => PdfDocument.Create().Paragraph(paragraph => paragraph.Text("PAdES signature test"));

    private static PdfCmsSignatureCryptographyProvider Provider() {
        var options = new CmsVerificationOptions();
        options.CertificateValidation.ChainEvaluator = (_, _) => true;
        return new PdfCmsSignatureCryptographyProvider(OfficeSecurityProvider.Default, options);
    }

    private static X509Certificate2 CreateCertificate(bool ecdsa) {
        using RSA? rsa = ecdsa ? null : RSA.Create(2048);
        using ECDsa? ec = ecdsa ? ECDsa.Create(ECCurve.NamedCurves.nistP256) : null;
        var request = ec is null
            ? new CertificateRequest("CN=OfficeIMO signature test", rsa!, HashAlgorithmName.SHA256, RSASignaturePadding.Pkcs1)
            : new CertificateRequest("CN=OfficeIMO EC signature test", ec, HashAlgorithmName.SHA256);
        request.CertificateExtensions.Add(new X509KeyUsageExtension(X509KeyUsageFlags.DigitalSignature, true));
        return request.CreateSelfSigned(DateTimeOffset.UtcNow.AddMinutes(-5), DateTimeOffset.UtcNow.AddDays(1));
    }
}
#endif
