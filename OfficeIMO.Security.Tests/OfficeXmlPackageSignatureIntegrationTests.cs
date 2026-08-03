using System.IO.Compression;
using System.IO;
using System.Text;
using System.Xml.Linq;
using OfficeIMO.Epub;
using OfficeIMO.OpenDocument;

namespace OfficeIMO.Security.Tests;

public sealed class OfficeXmlPackageSignatureIntegrationTests {
    [Theory]
    [InlineData(OfficeXmlPackageSignatureFormat.OpenDocument)]
    [InlineData(OfficeXmlPackageSignatureFormat.Epub)]
    public void XmlPackageSignatureRoundTripsAndDetectsTampering(OfficeXmlPackageSignatureFormat format) {
        string extension = format == OfficeXmlPackageSignatureFormat.OpenDocument ? "odt" : "epub";
        string path = Path.Combine(Path.GetTempPath(), $"OfficeIMO-{Guid.NewGuid():N}.{extension}");
        try {
            if (format == OfficeXmlPackageSignatureFormat.OpenDocument) {
                OdtDocument.Create().Save(path);
            } else {
                CreateMinimalEpub(path);
            }
            using X509Certificate2 certificate = CreateCertificate();
            OfficeXmlPackageSignatureOptions options = TrustlessOptions();
            OfficeXmlPackageSigningResult signed = format == OfficeXmlPackageSignatureFormat.OpenDocument
                ? OdfDocument.SignPackage(path, OfficeSecurityProvider.Default, certificate, options)
                : EpubDocument.SignPackage(path, OfficeSecurityProvider.Default, certificate, options);

            Assert.True(signed.Succeeded, string.Join(" ", signed.Findings));
            OfficeXmlPackageSignatureValidationReport valid = Validate(path, format, options);
            Assert.True(valid.IsValidUnderPolicy, string.Join(" ", valid.Findings));

            TamperEntry(path, format == OfficeXmlPackageSignatureFormat.OpenDocument ? "content.xml" : "EPUB/chapter.xhtml");
            OfficeXmlPackageSignatureValidationReport tampered = Validate(path, format, options);
            Assert.False(tampered.IsValidUnderPolicy);
            Assert.Contains(tampered.Signatures.SelectMany(item => item.Entries),
                entry => entry.Status == OfficePackageSignatureValidationState.Failed);
        } finally {
            if (File.Exists(path)) File.Delete(path);
        }
    }

    [Theory]
    [InlineData(OfficeXmlPackageSignatureFormat.OpenDocument)]
    [InlineData(OfficeXmlPackageSignatureFormat.Epub)]
    public void XmlPackageValidationRejectsManifestOutsideAuthenticatedObject(
        OfficeXmlPackageSignatureFormat format) {
        string extension = format == OfficeXmlPackageSignatureFormat.OpenDocument ? "odt" : "epub";
        string path = Path.Combine(Path.GetTempPath(), $"OfficeIMO-{Guid.NewGuid():N}.{extension}");
        try {
            if (format == OfficeXmlPackageSignatureFormat.OpenDocument) OdtDocument.Create().Save(path);
            else CreateMinimalEpub(path);
            using X509Certificate2 certificate = CreateCertificate();
            OfficeXmlPackageSignatureOptions options = TrustlessOptions();
            OfficeXmlPackageSigningResult signed = format == OfficeXmlPackageSignatureFormat.OpenDocument
                ? OdfDocument.SignPackage(path, OfficeSecurityProvider.Default, certificate, options)
                : EpubDocument.SignPackage(path, OfficeSecurityProvider.Default, certificate, options);
            Assert.True(signed.Succeeded);

            AddUnsignedManifest(path, format);
            OfficeXmlPackageSignatureValidationReport validation = Validate(path, format, options);

            Assert.False(validation.IsValidUnderPolicy);
            Assert.Contains(validation.Findings,
                finding => finding.Contains("not authenticated", StringComparison.OrdinalIgnoreCase));
        } finally {
            if (File.Exists(path)) File.Delete(path);
        }
    }

    private static OfficeXmlPackageSignatureValidationReport Validate(
        string path, OfficeXmlPackageSignatureFormat format, OfficeXmlPackageSignatureOptions options) =>
        format == OfficeXmlPackageSignatureFormat.OpenDocument
            ? OdfDocument.ValidatePackageSignatures(path, OfficeSecurityProvider.Default, options)
            : EpubDocument.ValidatePackageSignatures(path, OfficeSecurityProvider.Default, options);

    private static OfficeXmlPackageSignatureOptions TrustlessOptions() => new() { ValidateCertificateTrust = false };

    private static void CreateMinimalEpub(string path) {
        using ZipArchive archive = ZipFile.Open(path, ZipArchiveMode.Create);
        WriteEntry(archive, "mimetype", "application/epub+zip", CompressionLevel.NoCompression);
        WriteEntry(archive, "META-INF/container.xml",
            "<?xml version=\"1.0\"?><container xmlns=\"urn:oasis:names:tc:opendocument:xmlns:container\" version=\"1.0\"><rootfiles><rootfile full-path=\"EPUB/package.opf\" media-type=\"application/oebps-package+xml\"/></rootfiles></container>");
        WriteEntry(archive, "EPUB/package.opf",
            "<?xml version=\"1.0\"?><package xmlns=\"http://www.idpf.org/2007/opf\" version=\"3.0\" unique-identifier=\"id\"><metadata xmlns:dc=\"http://purl.org/dc/elements/1.1/\"><dc:identifier id=\"id\">urn:uuid:test</dc:identifier><dc:title>Test</dc:title><dc:language>en</dc:language></metadata><manifest><item id=\"chapter\" href=\"chapter.xhtml\" media-type=\"application/xhtml+xml\"/></manifest><spine><itemref idref=\"chapter\"/></spine></package>");
        WriteEntry(archive, "EPUB/chapter.xhtml", "<html xmlns=\"http://www.w3.org/1999/xhtml\"><body>signed content</body></html>");
    }

    private static void WriteEntry(ZipArchive archive, string name, string text, CompressionLevel compression = CompressionLevel.Optimal) {
        ZipArchiveEntry entry = archive.CreateEntry(name, compression);
        using var writer = new StreamWriter(entry.Open(), new UTF8Encoding(false));
        writer.Write(text);
    }

    private static void TamperEntry(string path, string entryName) {
        using ZipArchive archive = ZipFile.Open(path, ZipArchiveMode.Update);
        ZipArchiveEntry entry = archive.GetEntry(entryName) ?? throw new InvalidOperationException("Test entry missing.");
        using Stream output = entry.Open();
        output.Position = output.Length;
        output.WriteByte(0x20);
    }

    private static void AddUnsignedManifest(string path, OfficeXmlPackageSignatureFormat format) {
        string carrierPath = format == OfficeXmlPackageSignatureFormat.OpenDocument
            ? "META-INF/documentsignatures.xml"
            : "META-INF/signatures.xml";
        using ZipArchive archive = ZipFile.Open(path, ZipArchiveMode.Update);
        ZipArchiveEntry entry = archive.GetEntry(carrierPath)
            ?? throw new InvalidOperationException("Test signature carrier missing.");
        XDocument document;
        using (Stream input = entry.Open()) document = XDocument.Load(input, LoadOptions.PreserveWhitespace);
        XNamespace ds = XmlDigitalSignatureAlgorithms.Namespace;
        XNamespace manifestNamespace = "urn:officeimo:security:package-manifest:1";
        XElement signature = document.Root!.Elements(ds + "Signature").Single();
        XElement manifest = signature.Descendants(manifestNamespace + "PackageManifest").Single();
        XElement signedObject = manifest.Ancestors(ds + "Object").Single();
        signedObject.AddBeforeSelf(new XElement(ds + "Object",
            new XAttribute("Id", "UnsignedPackageManifest"), new XElement(manifest)));
        entry.Delete();
        ZipArchiveEntry replacement = archive.CreateEntry(carrierPath, CompressionLevel.Optimal);
        using Stream output = replacement.Open();
        document.Save(output, SaveOptions.DisableFormatting);
    }

    private static X509Certificate2 CreateCertificate() {
        using RSA rsa = RSA.Create(2048);
        var request = new CertificateRequest("CN=OfficeIMO XML Package Test", rsa, HashAlgorithmName.SHA256, RSASignaturePadding.Pkcs1);
        request.CertificateExtensions.Add(new X509KeyUsageExtension(X509KeyUsageFlags.DigitalSignature, true));
        return request.CreateSelfSigned(DateTimeOffset.UtcNow.AddMinutes(-1), DateTimeOffset.UtcNow.AddDays(1));
    }
}
