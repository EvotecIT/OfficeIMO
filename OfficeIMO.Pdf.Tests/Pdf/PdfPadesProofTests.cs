#if NET8_0_OR_GREATER
using System;
using System.Collections.Generic;
using System.IO;
using System.Security.Cryptography;
using System.Security.Cryptography.X509Certificates;
using System.Threading.Tasks;
using OfficeIMO.Pdf;
using OfficeIMO.Security;
using Org.BouncyCastle.Asn1;
using Org.BouncyCastle.Cms;
using Org.BouncyCastle.Security;
using Org.BouncyCastle.Tsp;
using Org.BouncyCastle.Utilities.Collections;
using Xunit;

namespace OfficeIMO.Tests.Pdf;

/// <summary>Exports PAdES samples with in-memory software keys and public test roots only.</summary>
[Collection("PDF standards proof output")]
public sealed class PdfPadesProofTests {
    [Fact]
    public void GenerateBaselineAndTimestampedSignaturesForIndependentValidation() {
        using RSA signerRootKey = RSA.Create(2048);
        using RSA tsaRootKey = RSA.Create(2048);
        using RSA signerKey = RSA.Create(2048);
        using RSA tsaKey = RSA.Create(2048);
        using X509Certificate2 signerRoot = CreateRoot("OfficeIMO proof signer root", signerRootKey);
        using X509Certificate2 tsaRoot = CreateRoot("OfficeIMO proof TSA root", tsaRootKey);
        using X509Certificate2 signerCertificate = CreateLeaf("OfficeIMO proof signer", signerKey, signerRoot, tsa: false);
        using X509Certificate2 tsaCertificate = CreateLeaf("OfficeIMO proof TSA", tsaKey, tsaRoot, tsa: true);
        var trust = new CertificateValidationOptions {
            ChainEvaluator = (certificate, _) => certificate.Thumbprint == tsaCertificate.Thumbprint
        };
        trust.ExtraCertificates.Add(tsaRoot);
        var timestamps = new Rfc3161TimestampClient(
            (request, _) => Task.FromResult(CreateTimestampResponse(request, tsaCertificate, tsaRoot)), trust);
        byte[] baseline = Sign(signerCertificate, signerRoot, new CmsSigningOptions());
        byte[] timestamped = Sign(signerCertificate, signerRoot,
            new CmsSigningOptions { SignatureTimestampProvider = timestamps.GetTimestamp });
        foreach (byte[] pdf in new[] { baseline, timestamped }) {
            PdfSignatureValidationReport report = PdfSignatureValidator.Validate(pdf);
            Assert.True(report.IsStructurallyValid);
            Assert.Single(report.Signatures);
        }

        string? output = Environment.GetEnvironmentVariable("OFFICEIMO_PDF_COMPLIANCE_PROOF_OUTPUT");
        if (string.IsNullOrWhiteSpace(output)) return;
        output = Path.Combine(output, "standards");
        Directory.CreateDirectory(output);
        File.WriteAllBytes(Path.Combine(output, "pades-bb.pdf"), baseline);
        File.WriteAllBytes(Path.Combine(output, "pades-bt.pdf"), timestamped);
        File.WriteAllText(Path.Combine(output, "signer-root.pem"), signerRoot.ExportCertificatePem());
        File.WriteAllText(Path.Combine(output, "tsa-root.pem"), tsaRoot.ExportCertificatePem());
    }

    private static byte[] Sign(X509Certificate2 certificate, X509Certificate2 root, CmsSigningOptions options) {
        using var signer = new PdfCmsExternalSigner(OfficeSecurityProvider.Default, certificate,
            signingOptions: options, certificateChain: new[] { root });
        return PdfDocument.Create().Meta(title: "OfficeIMO PAdES proof")
            .Paragraph(p => p.Text("Independent signature and timestamp proof"))
            .Security.SignExternal(signer, new PdfExternalSignatureOptions {
                SubFilter = PdfExternalSignatureSubFilter.CadesDetached,
                FieldName = "ProofSignature",
                Name = "OfficeIMO proof signer",
                ReservedSignatureContentsBytes = 16384
            }).ToDocument().ToBytes();
    }

    private static X509Certificate2 CreateRoot(string name, RSA key) {
        var request = new CertificateRequest("CN=" + name, key, HashAlgorithmName.SHA256, RSASignaturePadding.Pkcs1);
        request.CertificateExtensions.Add(new X509BasicConstraintsExtension(true, false, 0, true));
        request.CertificateExtensions.Add(new X509KeyUsageExtension(X509KeyUsageFlags.KeyCertSign | X509KeyUsageFlags.CrlSign, true));
        request.CertificateExtensions.Add(new X509SubjectKeyIdentifierExtension(request.PublicKey, false));
        return request.CreateSelfSigned(DateTimeOffset.UtcNow.AddMinutes(-10), DateTimeOffset.UtcNow.AddDays(2));
    }

    private static X509Certificate2 CreateLeaf(string name, RSA key, X509Certificate2 issuer, bool tsa) {
        var request = new CertificateRequest("CN=" + name, key, HashAlgorithmName.SHA256, RSASignaturePadding.Pkcs1);
        request.CertificateExtensions.Add(new X509BasicConstraintsExtension(false, false, 0, true));
        request.CertificateExtensions.Add(new X509KeyUsageExtension(tsa ? X509KeyUsageFlags.DigitalSignature :
            X509KeyUsageFlags.DigitalSignature | X509KeyUsageFlags.NonRepudiation, true));
        request.CertificateExtensions.Add(new X509SubjectKeyIdentifierExtension(request.PublicKey, false));
        if (tsa) request.CertificateExtensions.Add(new X509EnhancedKeyUsageExtension(
            new OidCollection { new Oid("1.3.6.1.5.5.7.3.8") }, true));
        using X509Certificate2 issued = request.Create(issuer, DateTimeOffset.UtcNow.AddMinutes(-5),
            DateTimeOffset.UtcNow.AddDays(1), RandomNumberGenerator.GetBytes(16));
        return issued.CopyWithPrivateKey(key);
    }

    private static byte[] CreateTimestampResponse(byte[] request, X509Certificate2 certificate, X509Certificate2 root) {
        using RSA key = certificate.GetRSAPrivateKey()!;
        var converted = DotNetUtilities.FromX509Certificate(certificate);
        var factory = new Org.BouncyCastle.Crypto.Operators.Asn1SignatureFactory("SHA256WITHRSA",
            DotNetUtilities.GetRsaKeyPair(key).Private);
        var signer = new SignerInfoGeneratorBuilder().Build(factory, converted);
        var tokens = new TimeStampTokenGenerator(signer, Org.BouncyCastle.Crypto.Operators.Asn1DigestFactory.Get("SHA256"),
            new DerObjectIdentifier("1.3.6.1.4.1.59069.1.1"), true);
        tokens.SetCertificates(new TimestampCertificates(converted, DotNetUtilities.FromX509Certificate(root)));
        var responses = new TimeStampResponseGenerator(tokens, TspAlgorithms.Allowed);
        return responses.Generate(new TimeStampRequest(request), Org.BouncyCastle.Math.BigInteger.One, DateTime.UtcNow).GetEncoded();
    }

    private sealed class TimestampCertificates(params Org.BouncyCastle.X509.X509Certificate[] certificates)
        : IStore<Org.BouncyCastle.X509.X509Certificate> {
        public IEnumerable<Org.BouncyCastle.X509.X509Certificate> EnumerateMatches(ISelector<Org.BouncyCastle.X509.X509Certificate> selector) {
            foreach (var certificate in certificates)
                if (selector is null || selector.Match(certificate)) yield return certificate;
        }
    }
}
#endif
