using System.IO.Compression;
using System.IO;
using System.Text;
using System.Xml.Linq;
using OfficeIMO.Excel;
using OfficeIMO.PowerPoint;
using OfficeIMO.Visio;
using OfficeIMO.Word;

namespace OfficeIMO.Security.Tests;

public sealed class OfficePackageSignatureIntegrationTests {
    [Theory]
    [InlineData("docx")]
    [InlineData("xlsx")]
    [InlineData("pptx")]
    [InlineData("vsdx")]
    public void SharedOpcEngineSignsAndValidatesEachHost(string extension) {
        string path = Path.Combine(Path.GetTempPath(), $"OfficeIMO-{Guid.NewGuid():N}.{extension}");
        try {
            CreatePackage(path, extension);
            using X509Certificate2 certificate = CreateCertificate();
            OfficePackageSigningResult signed = Sign(path, extension, certificate);

            Assert.True(signed.Succeeded, string.Join(" ", signed.Details));
            OfficePackageSignatureInfo inspection = Inspect(path, extension);
            Assert.True(inspection.HasDigitalSignatureOriginPart);
            Assert.Single(inspection.SignatureParts);

            OfficePackageSignatureValidationReport validation = Validate(path, extension);
            Assert.True(validation.IsCryptographicallyValid, string.Join(" ", validation.Findings));
            Assert.True(validation.IsValidUnderPolicy, string.Join(" ", validation.Findings));

            TamperFirstContentPart(path);
            OfficePackageSignatureValidationReport tampered = Validate(path, extension);
            Assert.False(tampered.IsCryptographicallyValid);
            Assert.Contains(tampered.Signatures.SelectMany(item => item.SignaturePart.SignedReferences),
                reference => reference.DigestVerificationStatus == OfficePackageSignatureValidationState.Failed);
        } finally {
            if (File.Exists(path)) File.Delete(path);
        }
    }

    [Fact]
    public void OpcValidationRejectsUnsignedManifestAndOrphanSignatureParts() {
        string path = Path.Combine(Path.GetTempPath(), $"OfficeIMO-{Guid.NewGuid():N}.docx");
        try {
            CreatePackage(path, "docx");
            using X509Certificate2 certificate = CreateCertificate();
            Assert.True(Sign(path, "docx", certificate).Succeeded);

            AddUnsignedManifest(path);
            OfficePackageSignatureValidationReport wrapped = Validate(path, "docx");
            Assert.False(wrapped.IsCryptographicallyValid);
            Assert.Contains(wrapped.SignatureInfo.SignatureParts,
                part => part.ParseError?.Contains("not authenticated", StringComparison.OrdinalIgnoreCase) == true);

            RemoveSignatureOriginRelationship(path);
            OfficePackageSignatureValidationReport orphan = Validate(path, "docx");
            Assert.False(orphan.IsCryptographicallyValid);
            Assert.Contains(orphan.SignatureInfo.SignatureParts, part => !part.IsReachableFromOrigin);
        } finally {
            if (File.Exists(path)) File.Delete(path);
        }
    }

    [Fact]
    public void OpcValidationBoundsAggregateDigestWorkBeforeReadingParts() {
        string path = Path.Combine(Path.GetTempPath(), $"OfficeIMO-{Guid.NewGuid():N}.docx");
        try {
            CreatePackage(path, "docx");
            using X509Certificate2 certificate = CreateCertificate();
            Assert.True(Sign(path, "docx", certificate).Succeeded);
            var options = new OfficePackageSignatureValidationOptions { ValidateCertificateTrust = false };
            options.Inspection.MaxTotalDigestBytes = 1;

            OfficePackageSignatureValidationReport validation =
                WordDocument.ValidatePackageSignatures(path, OfficeSecurityProvider.Default, options);

            Assert.False(validation.IsCryptographicallyValid);
            Assert.Contains(validation.SignatureInfo.SignatureParts,
                part => part.ParseError?.Contains("aggregate limit", StringComparison.OrdinalIgnoreCase) == true);
        } finally {
            if (File.Exists(path)) File.Delete(path);
        }
    }

    [Fact]
    public void OpcValidationFailsClosedWhenSignatureDiscoveryIsTruncated() {
        string path = Path.Combine(Path.GetTempPath(), $"OfficeIMO-{Guid.NewGuid():N}.docx");
        try {
            CreatePackage(path, "docx");
            using X509Certificate2 certificate = CreateCertificate();
            Assert.True(Sign(path, "docx", certificate).Succeeded);
            CloneSignaturePart(path);
            var options = new OfficePackageSignatureValidationOptions { ValidateCertificateTrust = false };
            options.Inspection.MaxSignatureParts = 1;

            OfficePackageSignatureValidationReport validation =
                WordDocument.ValidatePackageSignatures(path, OfficeSecurityProvider.Default, options);

            Assert.False(validation.SignatureInfo.SignatureDiscoveryComplete);
            Assert.False(validation.IsCryptographicallyValid);
            Assert.False(validation.IsValidUnderPolicy);
        } finally {
            if (File.Exists(path)) File.Delete(path);
        }
    }

    [Fact]
    public void OpcRevocationRequirementRejectsAnInconclusiveResult() {
        var reference = new OfficePackageSignatureReferenceInfo("/word/document.xml", "sha256", "digest",
            "/word/document.xml", true, Array.Empty<string>(), OfficePackageSignatureValidationState.Passed, null);
        var part = new OfficePackageSignaturePartInfo("/_xmlsignatures/sig1.xml", 1, true, "rsa-sha256",
            new[] { reference }, Array.Empty<OfficePackageSignatureTimestampInfo>(), Array.Empty<string>(),
            Array.Empty<byte[]>(), null);
        var validated = new OfficePackageSignaturePartValidationResult(part,
            OfficePackageSignatureValidationState.Passed, OfficePackageSignatureValidationState.Passed,
            OfficePackageSignatureValidationState.NotChecked, true, true, Array.Empty<SecurityFinding>());

        Assert.False(validated.IsValidUnderPolicy);
    }

    [Fact]
    public void OpcApplicationSignatureMetadataRequiresExactRootNamespaceAndElement() {
        string path = Path.Combine(Path.GetTempPath(), $"OfficeIMO-{Guid.NewGuid():N}.docx");
        try {
            CreatePackage(path, "docx");
            WriteApplicationProperties(path,
                "<ep:Properties xmlns:ep=\"http://schemas.openxmlformats.org/officeDocument/2006/extended-properties\"><x:DigSig xmlns:x=\"urn:not-office\" /></ep:Properties>");
            Assert.False(WordDocument.InspectPackageSignatures(path).HasApplicationSignatureMetadata);

            WriteApplicationProperties(path,
                "<ep:Properties xmlns:ep=\"http://schemas.openxmlformats.org/officeDocument/2006/extended-properties\"><ep:DigSig /></ep:Properties>");
            Assert.True(WordDocument.InspectPackageSignatures(path).HasApplicationSignatureMetadata);
        } finally {
            if (File.Exists(path)) File.Delete(path);
        }
    }

    [Fact]
    public void WordProjectionPreservesPackageTimestampValueAndFormat() {
        string path = Path.Combine(Path.GetTempPath(), $"OfficeIMO-{Guid.NewGuid():N}.docx");
        try {
            CreatePackage(path, "docx");
            using X509Certificate2 certificate = CreateCertificate();
            Assert.True(Sign(path, "docx", certificate).Succeeded);
            AddPackageSignatureTime(path, "2026-08-03T12:34:56Z", "YYYY-MM-DDThh:mm:ssTZD");

            using WordDocument document = WordDocument.Load(path,
                new WordLoadOptions { AccessMode = OfficeIMO.DocumentAccessMode.ReadOnly });
            WordSignatureTimestampInfo timestamp = Assert.Single(
                Assert.Single(document.InspectSignatures().SignatureParts).Timestamps,
                item => item.Value == "2026-08-03T12:34:56Z");
            Assert.Equal("SignatureTime", timestamp.Kind);
            Assert.Equal("2026-08-03T12:34:56Z", timestamp.Value);
            Assert.Equal("YYYY-MM-DDThh:mm:ssTZD", timestamp.Format);
        } finally {
            if (File.Exists(path)) File.Delete(path);
        }
    }

    [Theory]
    [InlineData("wrong-root")]
    [InlineData("wrong-namespace")]
    [InlineData("nested")]
    [InlineData("malformed-target")]
    public void OpcValidationRejectsNonOpcOriginRelationshipShapes(string shape) {
        string path = Path.Combine(Path.GetTempPath(), $"OfficeIMO-{Guid.NewGuid():N}.docx");
        try {
            CreatePackage(path, "docx");
            using X509Certificate2 certificate = CreateCertificate();
            Assert.True(Sign(path, "docx", certificate).Succeeded);

            MutateSignatureOriginRelationship(path, shape);
            OfficePackageSignatureValidationReport validation = Validate(path, "docx");

            Assert.False(validation.IsCryptographicallyValid);
            Assert.Equal(shape == "malformed-target" ? 1 : 0, validation.SignatureInfo.OriginRelationshipCount);
            if (shape == "malformed-target") Assert.True(validation.SignatureInfo.HasSignatures);
            Assert.Contains(validation.SignatureInfo.SignatureParts, part => !part.IsReachableFromOrigin);
        } finally {
            if (File.Exists(path)) File.Delete(path);
        }
    }

    private static void CreatePackage(string path, string extension) {
        switch (extension) {
            case "docx":
                using (WordDocument document = WordDocument.Create(path)) {
                    document.AddParagraph("signed content");
                    document.Save();
                }
                break;
            case "xlsx":
                using (ExcelDocument document = ExcelDocument.Create(path)) {
                    document.AddWorksheet("Data").CellValue(1, 1, "signed content");
                    document.Save();
                }
                break;
            case "pptx":
                using (PowerPointPresentation presentation = PowerPointPresentation.Create(path)) presentation.Save();
                break;
            case "vsdx":
                VisioDocument.Create(path).Save();
                break;
            default:
                throw new ArgumentOutOfRangeException(nameof(extension));
        }
    }

    private static OfficePackageSigningResult Sign(string path, string extension, X509Certificate2 certificate) => extension switch {
        "docx" => WordDocument.SignPackageSignature(path, OfficeSecurityProvider.Default, certificate),
        "xlsx" => ExcelDocument.SignPackageSignature(path, OfficeSecurityProvider.Default, certificate),
        "pptx" => PowerPointPresentation.SignPackageSignature(path, OfficeSecurityProvider.Default, certificate),
        "vsdx" => VisioDocument.SignPackageSignature(path, OfficeSecurityProvider.Default, certificate),
        _ => throw new ArgumentOutOfRangeException(nameof(extension))
    };

    private static OfficePackageSignatureInfo Inspect(string path, string extension) => extension switch {
        "docx" => WordDocument.InspectPackageSignatures(path),
        "xlsx" => ExcelDocument.InspectPackageSignatures(path),
        "pptx" => PowerPointPresentation.InspectPackageSignatures(path),
        "vsdx" => VisioDocument.InspectPackageSignatures(path),
        _ => throw new ArgumentOutOfRangeException(nameof(extension))
    };

    private static OfficePackageSignatureValidationReport Validate(string path, string extension) {
        var options = new OfficePackageSignatureValidationOptions { ValidateCertificateTrust = false };
        return extension switch {
            "docx" => WordDocument.ValidatePackageSignatures(path, OfficeSecurityProvider.Default, options),
            "xlsx" => ExcelDocument.ValidatePackageSignatures(path, OfficeSecurityProvider.Default, options),
            "pptx" => PowerPointPresentation.ValidatePackageSignatures(path, OfficeSecurityProvider.Default, options),
            "vsdx" => VisioDocument.ValidatePackageSignatures(path, OfficeSecurityProvider.Default, options),
            _ => throw new ArgumentOutOfRangeException(nameof(extension))
        };
    }

    private static void TamperFirstContentPart(string path) {
        using ZipArchive archive = ZipFile.Open(path, ZipArchiveMode.Update);
        ZipArchiveEntry entry = archive.Entries.First(item =>
            !item.FullName.StartsWith("_xmlsignatures/", StringComparison.OrdinalIgnoreCase) &&
            !string.Equals(item.FullName, "[Content_Types].xml", StringComparison.OrdinalIgnoreCase) &&
            !item.FullName.EndsWith(".rels", StringComparison.OrdinalIgnoreCase));
        using Stream output = entry.Open();
        output.Position = output.Length;
        output.WriteByte(0x20);
    }

    private static void AddUnsignedManifest(string path) {
        using ZipArchive archive = ZipFile.Open(path, ZipArchiveMode.Update);
        ZipArchiveEntry entry = archive.GetEntry("_xmlsignatures/sig1.xml")
            ?? throw new InvalidOperationException("Test signature entry missing.");
        XDocument document;
        using (Stream input = entry.Open()) document = XDocument.Load(input, LoadOptions.PreserveWhitespace);
        XNamespace ds = XmlDigitalSignatureAlgorithms.Namespace;
        XElement manifest = document.Descendants(ds + "Manifest").Single();
        XElement signedObject = manifest.Ancestors(ds + "Object").Single();
        signedObject.AddBeforeSelf(new XElement(ds + "Object",
            new XAttribute("Id", "UnsignedPackageObject"), new XElement(manifest)));
        entry.Delete();
        ZipArchiveEntry replacement = archive.CreateEntry("_xmlsignatures/sig1.xml", CompressionLevel.Optimal);
        using Stream output = replacement.Open();
        document.Save(output, SaveOptions.DisableFormatting);
    }

    private static void RemoveSignatureOriginRelationship(string path) {
        using ZipArchive archive = ZipFile.Open(path, ZipArchiveMode.Update);
        ZipArchiveEntry entry = archive.GetEntry("_rels/.rels")
            ?? throw new InvalidOperationException("Root relationships entry missing.");
        XDocument document;
        using (Stream input = entry.Open()) document = XDocument.Load(input, LoadOptions.PreserveWhitespace);
        document.Descendants().Where(element => element.Name.LocalName == "Relationship" &&
            ((string?)element.Attribute("Type"))?.EndsWith("/digital-signature/origin", StringComparison.Ordinal) == true)
            .Remove();
        entry.Delete();
        ZipArchiveEntry replacement = archive.CreateEntry("_rels/.rels", CompressionLevel.Optimal);
        using Stream output = replacement.Open();
        document.Save(output, SaveOptions.DisableFormatting);
    }

    private static void MutateSignatureOriginRelationship(string path, string shape) {
        using ZipArchive archive = ZipFile.Open(path, ZipArchiveMode.Update);
        ZipArchiveEntry entry = archive.GetEntry("_rels/.rels")
            ?? throw new InvalidOperationException("Root relationships entry missing.");
        XDocument document;
        using (Stream input = entry.Open()) document = XDocument.Load(input, LoadOptions.PreserveWhitespace);
        XNamespace relationshipsNamespace = "http://schemas.openxmlformats.org/package/2006/relationships";
        XElement root = document.Root ?? throw new InvalidOperationException("Root relationships element missing.");
        XElement origin = root.Elements(relationshipsNamespace + "Relationship").Single(element =>
            ((string?)element.Attribute("Type"))?.EndsWith("/digital-signature/origin", StringComparison.Ordinal) == true);
        switch (shape) {
            case "wrong-root":
                root.Name = relationshipsNamespace + "NotRelationships";
                break;
            case "wrong-namespace":
                XNamespace wrongNamespace = "urn:officeimo:wrong-relationships";
                origin.Name = wrongNamespace + "Relationship";
                break;
            case "nested":
                origin.Remove();
                root.Add(new XElement(relationshipsNamespace + "Wrapper", origin));
                break;
            case "malformed-target":
                origin.SetAttributeValue("Target", "http://[");
                break;
            default:
                throw new ArgumentOutOfRangeException(nameof(shape));
        }
        entry.Delete();
        ZipArchiveEntry replacement = archive.CreateEntry("_rels/.rels", CompressionLevel.Optimal);
        using Stream output = replacement.Open();
        document.Save(output, SaveOptions.DisableFormatting);
    }

    private static void CloneSignaturePart(string path) {
        using ZipArchive archive = ZipFile.Open(path, ZipArchiveMode.Update);
        ZipArchiveEntry source = archive.GetEntry("_xmlsignatures/sig1.xml")
            ?? throw new InvalidOperationException("Test signature entry missing.");
        byte[] bytes;
        using (Stream input = source.Open()) {
            using var memory = new MemoryStream();
            input.CopyTo(memory);
            bytes = memory.ToArray();
        }
        ZipArchiveEntry clone = archive.CreateEntry("_xmlsignatures/sig2.xml", CompressionLevel.Optimal);
        using (Stream output = clone.Open()) output.Write(bytes, 0, bytes.Length);

        MutateXmlEntry(archive, "_xmlsignatures/_rels/origin.sigs.rels", document => {
            XNamespace relationships = "http://schemas.openxmlformats.org/package/2006/relationships";
            document.Root!.Add(new XElement(relationships + "Relationship",
                new XAttribute("Id", "rIdOfficeImoClone"),
                new XAttribute("Type", "http://schemas.openxmlformats.org/package/2006/relationships/digital-signature/signature"),
                new XAttribute("Target", "sig2.xml")));
        });
        MutateXmlEntry(archive, "[Content_Types].xml", document => {
            XNamespace types = "http://schemas.openxmlformats.org/package/2006/content-types";
            document.Root!.Add(new XElement(types + "Override",
                new XAttribute("PartName", "/_xmlsignatures/sig2.xml"),
                new XAttribute("ContentType", "application/vnd.openxmlformats-package.digital-signature-xmlsignature+xml")));
        });
    }

    private static void WriteApplicationProperties(string path, string xml) {
        using ZipArchive archive = ZipFile.Open(path, ZipArchiveMode.Update);
        archive.GetEntry("docProps/app.xml")?.Delete();
        ZipArchiveEntry entry = archive.CreateEntry("docProps/app.xml", CompressionLevel.Optimal);
        using var writer = new StreamWriter(entry.Open(), new UTF8Encoding(false));
        writer.Write(xml);
    }

    private static void AddPackageSignatureTime(string path, string value, string format) {
        using ZipArchive archive = ZipFile.Open(path, ZipArchiveMode.Update);
        MutateXmlEntry(archive, "_xmlsignatures/sig1.xml", document => {
            XNamespace ds = XmlDigitalSignatureAlgorithms.Namespace;
            XNamespace package = "http://schemas.openxmlformats.org/package/2006/digital-signature";
            document.Root!.Add(new XElement(ds + "Object",
                new XElement(package + "SignatureTime",
                    new XElement(package + "Format", format),
                    new XElement(package + "Value", value))));
        });
    }

    private static void MutateXmlEntry(ZipArchive archive, string path, Action<XDocument> mutation) {
        ZipArchiveEntry entry = archive.GetEntry(path)
            ?? throw new InvalidOperationException("Test XML entry missing: " + path);
        XDocument document;
        using (Stream input = entry.Open()) document = XDocument.Load(input, LoadOptions.PreserveWhitespace);
        mutation(document);
        entry.Delete();
        ZipArchiveEntry replacement = archive.CreateEntry(path, CompressionLevel.Optimal);
        using Stream output = replacement.Open();
        document.Save(output, SaveOptions.DisableFormatting);
    }

    private static X509Certificate2 CreateCertificate() {
        using RSA rsa = RSA.Create(2048);
        var request = new CertificateRequest("CN=OfficeIMO OPC Test", rsa, HashAlgorithmName.SHA256, RSASignaturePadding.Pkcs1);
        request.CertificateExtensions.Add(new X509KeyUsageExtension(X509KeyUsageFlags.DigitalSignature, true));
        return request.CreateSelfSigned(DateTimeOffset.UtcNow.AddMinutes(-1), DateTimeOffset.UtcNow.AddDays(1));
    }
}
