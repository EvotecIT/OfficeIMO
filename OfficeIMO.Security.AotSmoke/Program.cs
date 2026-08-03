using System.Security.Cryptography;
using System.Security.Cryptography.X509Certificates;
using System.Text;
using OfficeIMO.Security;

using RSA rsa = RSA.Create(2048);
var request = new CertificateRequest(
    "CN=OfficeIMO Security NativeAOT",
    rsa,
    HashAlgorithmName.SHA256,
    RSASignaturePadding.Pkcs1);
request.CertificateExtensions.Add(new X509KeyUsageExtension(
    X509KeyUsageFlags.DigitalSignature,
    critical: true));
using X509Certificate2 certificate = request.CreateSelfSigned(
    DateTimeOffset.UtcNow.AddMinutes(-1),
    DateTimeOffset.UtcNow.AddDays(1));

IOfficeSecurityProvider provider = OfficeSecurityProvider.Default;
byte[] content = Encoding.UTF8.GetBytes("OfficeIMO Security NativeAOT CMS marker");
byte[] cms = provider.SignCmsDetached(content, certificate);
CmsVerificationResult cmsResult = provider.VerifyCmsDetached(
    cms,
    content,
    new CmsVerificationOptions { ValidateTimestamps = false });
if (!cmsResult.IsCryptographicallyValid) {
    throw new InvalidOperationException("The NativeAOT CMS detached-signature round trip failed.");
}

var xmlRequest = new XmlDigitalSignatureCreationRequest(
    Encoding.UTF8.GetBytes("<Root><Payload>OfficeIMO Security NativeAOT XML marker</Payload></Root>"),
    certificate,
    "OfficeIMOSecurityAotSignature",
    "OfficeIMOSecurityAotObject",
    "urn:officeimo:aot-object",
    XmlDigitalSignatureAlgorithms.CanonicalXml,
    XmlDigitalSignatureAlgorithms.RsaSha256,
    "http://www.w3.org/2001/04/xmlenc#sha256");
byte[] signedXml = provider.CreateXmlSignature(xmlRequest);
XmlDigitalSignatureVerificationResult xmlResult = provider.VerifyXmlSignature(
    new XmlDigitalSignatureVerificationRequest(signedXml, new[] { certificate }));
if (xmlResult.Status != SecurityValidationStatus.Valid) {
    throw new InvalidOperationException("The NativeAOT XML signature round trip failed.");
}

Console.WriteLine("PASS | explicit OfficeIMO.Security CMS and XML provider contracts passed from NativeAOT.");
