using System.Xml;

namespace OfficeIMO.Security.Tests;

public sealed class XmlDigitalSignaturePolicyTests {
    [Fact]
    public void VerificationDoesNotLetCallerExpandTheSignatureMethodSet() {
        using X509Certificate2 certificate = CreateCertificate();
        byte[] signedXml = CreateSignature(certificate);
        byte[] unsupported = Encoding.UTF8.GetBytes(
            Encoding.UTF8.GetString(signedXml)
                .Replace(XmlDigitalSignatureAlgorithms.RsaSha256, "urn:officeimo:unsupported-signature", StringComparison.Ordinal));
        var request = new XmlDigitalSignatureVerificationRequest(unsupported, new[] { certificate }) {
            AllowedSignatureMethods = new[] { "urn:officeimo:unsupported-signature" }
        };

        XmlDigitalSignatureVerificationResult result = OfficeSecurityProvider.Default.VerifyXmlSignature(request);

        Assert.Equal(SecurityValidationStatus.Indeterminate, result.Status);
        Assert.Contains(result.Findings, finding => finding.Code == "UnsupportedSignatureMethod");
    }

    [Fact]
    public void VerificationDoesNotLetCallerExpandTheDigestMethodSet() {
        using X509Certificate2 certificate = CreateCertificate();
        byte[] signedXml = CreateSignature(certificate);
        byte[] unsupported = Encoding.UTF8.GetBytes(
            Encoding.UTF8.GetString(signedXml)
                .Replace(XmlDigitalSignatureAlgorithms.Sha256, "urn:officeimo:unsupported-digest", StringComparison.Ordinal));
        var request = new XmlDigitalSignatureVerificationRequest(unsupported, new[] { certificate }) {
            AllowedDigestMethods = new[] { "urn:officeimo:unsupported-digest" }
        };

        XmlDigitalSignatureVerificationResult result = OfficeSecurityProvider.Default.VerifyXmlSignature(request);

        Assert.Equal(SecurityValidationStatus.Indeterminate, result.Status);
        Assert.Contains(result.Findings, finding => finding.Code == "UnsupportedDigestMethod");
    }

    [Fact]
    public void VerificationDoesNotLetCallerExpandTheTransformSet() {
        using X509Certificate2 certificate = CreateCertificate();
        var document = new XmlDocument { PreserveWhitespace = true };
        document.LoadXml(Encoding.UTF8.GetString(CreateSignature(certificate)));
        XmlNamespaceManager namespaces = new(document.NameTable);
        namespaces.AddNamespace("ds", XmlDigitalSignatureAlgorithms.Namespace);
        XmlElement digestMethod = (XmlElement?)document.SelectSingleNode(
            "/ds:Signature/ds:SignedInfo/ds:Reference/ds:DigestMethod",
            namespaces) ?? throw new InvalidOperationException("DigestMethod was not generated.");
        XmlElement transforms = document.CreateElement("ds", "Transforms", XmlDigitalSignatureAlgorithms.Namespace);
        XmlElement transform = document.CreateElement("ds", "Transform", XmlDigitalSignatureAlgorithms.Namespace);
        transform.SetAttribute("Algorithm", "urn:officeimo:unsupported-transform");
        transforms.AppendChild(transform);
        digestMethod.ParentNode!.InsertBefore(transforms, digestMethod);
        var request = new XmlDigitalSignatureVerificationRequest(
            Encoding.UTF8.GetBytes(document.OuterXml),
            new[] { certificate }) {
            AllowedReferenceTransforms = new[] { "urn:officeimo:unsupported-transform" }
        };

        XmlDigitalSignatureVerificationResult result = OfficeSecurityProvider.Default.VerifyXmlSignature(request);

        Assert.Equal(SecurityValidationStatus.Indeterminate, result.Status);
        Assert.Contains(result.Findings, finding => finding.Code == "UnsupportedSignedInfoTransform");
    }

    [Fact]
    public void CreationRejectsAlgorithmsOutsideTheProviderSet() {
        using X509Certificate2 certificate = CreateCertificate();
        var request = CreateRequest(certificate, signatureMethod: "urn:officeimo:unsupported-signature");

        Assert.Throws<NotSupportedException>(() => OfficeSecurityProvider.Default.CreateXmlSignature(request));
    }

    private static byte[] CreateSignature(X509Certificate2 certificate) =>
        OfficeSecurityProvider.Default.CreateXmlSignature(CreateRequest(certificate));

    private static XmlDigitalSignatureCreationRequest CreateRequest(
        X509Certificate2 certificate,
        string signatureMethod = XmlDigitalSignatureAlgorithms.RsaSha256) =>
        new(
            Encoding.UTF8.GetBytes("<Root><Payload>OfficeIMO XML policy</Payload></Root>"),
            certificate,
            "OfficeIMOPolicySignature",
            "OfficeIMOPolicyObject",
            "urn:officeimo:policy-object",
            XmlDigitalSignatureAlgorithms.CanonicalXml,
            signatureMethod,
            XmlDigitalSignatureAlgorithms.Sha256);

    private static X509Certificate2 CreateCertificate() {
        using RSA rsa = RSA.Create(2048);
        var request = new CertificateRequest(
            "CN=OfficeIMO XML Policy",
            rsa,
            HashAlgorithmName.SHA256,
            RSASignaturePadding.Pkcs1);
        request.CertificateExtensions.Add(new X509KeyUsageExtension(
            X509KeyUsageFlags.DigitalSignature,
            critical: true));
        return request.CreateSelfSigned(
            DateTimeOffset.UtcNow.AddMinutes(-1),
            DateTimeOffset.UtcNow.AddDays(1));
    }
}
