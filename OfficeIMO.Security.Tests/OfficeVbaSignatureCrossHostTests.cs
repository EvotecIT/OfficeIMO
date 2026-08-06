using System.IO;
using System.IO.Compression;
using System.Security.Cryptography;
using System.Security.Cryptography.X509Certificates;
using System.Text;
using System.Xml.Linq;
using OfficeIMO.Core.Internal;
using OfficeIMO.Excel;
using OfficeIMO.PowerPoint;
using OfficeIMO.Word;

namespace OfficeIMO.Security.Tests;

public sealed class OfficeVbaSignatureCrossHostTests {
    [Theory]
    [InlineData("docm", "word")]
    [InlineData("xlsm", "xl")]
    [InlineData("xlsb", "xl")]
    [InlineData("pptm", "ppt")]
    [InlineData("ppam", "ppt")]
    public void SharedInspectorReportsLegacyAgileAndV3Profiles(string extension, string hostRoot) {
        string path = Path.Combine(Path.GetTempPath(), $"OfficeIMO-{Guid.NewGuid():N}.{extension}");
        try {
            CreateMacroPackage(path, hostRoot);

            OfficeVbaSignatureInfo info = extension switch {
                "docm" => WordDocument.InspectVbaSignatures(path),
                "xlsm" or "xlsb" => ExcelDocument.InspectVbaSignatures(path),
                "pptm" or "ppam" => PowerPointPresentation.InspectVbaSignatures(path),
                _ => throw new ArgumentOutOfRangeException(nameof(extension))
            };

            Assert.True(info.IsMacroEnabledFormat);
            Assert.True(info.HasMacroProject);
            Assert.Collection(info.Signatures.OrderBy(item => item.Profile),
                legacy => Assert.Equal(OfficeVbaSignatureProfile.Legacy, legacy.Profile),
                agile => Assert.Equal(OfficeVbaSignatureProfile.Agile, agile.Profile),
                v3 => Assert.Equal(OfficeVbaSignatureProfile.V3, v3.Profile));
            Assert.All(info.Signatures, signature => Assert.True(signature.CmsParsed));
        } finally {
            if (File.Exists(path)) File.Delete(path);
        }
    }

    [Theory]
    [InlineData("docm", "word")]
    [InlineData("xlsm", "xl")]
    [InlineData("xlsb", "xl")]
    [InlineData("pptm", "ppt")]
    [InlineData("ppam", "ppt")]
    public void ManagedSigningCreatesAndBindsAllProfilesAcrossHosts(string extension, string hostRoot) {
        string path = Path.Combine(Path.GetTempPath(), $"OfficeIMO-managed-vba-{Guid.NewGuid():N}.{extension}");
        try {
            CreateUnsignedMacroPackage(path, hostRoot, CreateVbaProject("Test"));
            using X509Certificate2 certificate = CreateSigningCertificate();
            var options = new OfficeVbaSigningOptions();
            options.CmsVerification.CertificateValidation.ChainEvaluator = (_, _) => true;
            options.CmsVerification.CertificateValidation.DisableCertificateDownloads = false;

            OfficeVbaSigningResult signing = extension switch {
                "docm" => WordDocument.TrySignVbaProject(path, OfficeSecurityProvider.Default, certificate, options),
                "xlsm" or "xlsb" => ExcelDocument.TrySignVbaProject(path, OfficeSecurityProvider.Default, certificate, options),
                "pptm" or "ppam" => PowerPointPresentation.TrySignVbaProject(path, OfficeSecurityProvider.Default, certificate, options),
                _ => throw new ArgumentOutOfRangeException(nameof(extension))
            };

            Assert.True(signing.Succeeded, string.Join(" | ", signing.Findings.Select(finding => finding.Code + ": " + finding.Message)));
            Assert.NotNull(signing.Validation);
            Assert.True(signing.Validation!.IsValidUnderPolicy,
                string.Join(" | ", signing.Validation.Findings.Select(finding => finding.Code + ": " + finding.Message)));
            Assert.Equal(3, signing.Validation.SignatureInfo.Signatures.Count);
            Assert.All(signing.Validation.SignatureInfo.Signatures, signature => {
                Assert.True(signature.CmsParsed);
                Assert.Equal(OfficePackageSignatureValidationState.Passed, signature.CryptographicStatus);
            });

            ReplaceVbaProject(path, hostRoot, CreateVbaProject("Tampered"));
            OfficeVbaSignatureValidationResult tampered = extension switch {
                "docm" => WordDocument.ValidateVbaSignatures(path, OfficeSecurityProvider.Default, options),
                "xlsm" or "xlsb" => ExcelDocument.ValidateVbaSignatures(path, OfficeSecurityProvider.Default, options),
                "pptm" or "ppam" => PowerPointPresentation.ValidateVbaSignatures(path, OfficeSecurityProvider.Default, options),
                _ => throw new ArgumentOutOfRangeException(nameof(extension))
            };
            Assert.False(tampered.IsValidUnderPolicy);
            Assert.Equal(OfficePackageSignatureValidationState.Failed, tampered.ContentBindingStatus);
        } finally {
            if (File.Exists(path)) File.Delete(path);
        }
    }

    [Fact]
    public void ManagedCanonicalizerHasFrozenProfileDigestsForProducerFixture() {
        byte[] project = CreateVbaProject("Test");
        Assert.True(OfficeVbaProjectCanonicalizer.TryCreate(project, 4 * 1024 * 1024,
            out OfficeVbaProjectCanonicalizer.Result? canonical, out string detail), detail);
        Assert.NotNull(canonical);
        Assert.Equal("32D8B654F9A2391A1140B0FD1FF6EE3D", Hex(canonical!.ComputeLegacyHash()));
        Assert.Equal("07B03AD2ED4A8266556638E4FB8ECD4A4743A8F96277711C34978281F9288CC7",
            Hex(canonical.ComputeAgileHash()));
        Assert.Equal("EBACF918E4BDCF6460858BDCE208B91B09ADB764CFA69E5E1F895EFBAD6DF7DB",
            Hex(canonical.ComputeV3Hash()));
    }

    [Fact]
    public void ManagedCanonicalizerProcessesManySourceLinesWithoutRetainingASplitArray() {
        string sourceLines = string.Join("\r\n", Enumerable.Repeat("' bounded source line", 20_000));
        byte[] project = CreateVbaProject("ManyLines()\r\n" + sourceLines + "\r\nSub Tail");

        bool created = OfficeVbaProjectCanonicalizer.TryCreate(
            project,
            1024 * 1024,
            out OfficeVbaProjectCanonicalizer.Result? canonical,
            out string detail);

        Assert.True(created, detail);
        Assert.NotNull(canonical);
        Assert.NotEmpty(canonical!.ComputeV3Hash());
    }

    [Fact]
    public void ManagedCanonicalizerPreservesDesignerStorageTraversalOrder() {
        byte[] project = CreateVbaProject("Test", "UserForm1", 0, new[] {
            new OfficeCompoundStream("UserForm1/AA", new byte[] { 0xAA }),
            new OfficeCompoundStream("UserForm1/Z", new byte[] { 0x5A })
        });

        Assert.True(OfficeVbaProjectCanonicalizer.TryCreate(project, 4 * 1024 * 1024,
            out OfficeVbaProjectCanonicalizer.Result? canonical, out string detail), detail);
        Assert.NotNull(canonical);
        Assert.Equal(2046, canonical!.FormsNormalizedData.Length);
        Assert.Equal(0x5A, canonical.FormsNormalizedData[0]);
        Assert.Equal(0xAA, canonical.FormsNormalizedData[1023]);
    }

    [Fact]
    public void ManagedCanonicalizerFailsClosedWhenDesignerTranscriptExceedsLimit() {
        byte[] project = CreateVbaProject("Test", "UserForm1", 0, new[] {
            new OfficeCompoundStream("UserForm1/Data", new byte[2000])
        });

        bool created = OfficeVbaProjectCanonicalizer.TryCreate(project, 1500,
            out _, out string detail);

        Assert.False(created);
        Assert.Contains("forms-normalized transcript exceeds", detail, StringComparison.OrdinalIgnoreCase);
    }

    [Fact]
    public void ManagedCanonicalizerRejectsModuleOffsetAboveInt32WithoutThrowing() {
        byte[] project = CreateVbaProject("Test", "Module1", uint.MaxValue, null);

        bool created = OfficeVbaProjectCanonicalizer.TryCreate(project, 4 * 1024 * 1024,
            out _, out string detail);

        Assert.False(created);
        Assert.Contains("invalid module record", detail, StringComparison.OrdinalIgnoreCase);
    }

    [Fact]
    public void ManagedSignaturePartContainsBoundedCertificateStoreAndReservedTerminators() {
        byte[] cms = { 0x30, 0x01, 0x00 };
        byte[] certificate = { 0x30, 0x02, 0x01, 0x00 };

        byte[] encoded = OfficeVbaSignatureEncoding.CreateDigSigInfoSerialized(cms, certificate);

        Assert.Equal((uint)cms.Length, ReadU32(encoded, 0));
        Assert.Equal(44U, ReadU32(encoded, 4));
        Assert.Equal((uint)(32 + certificate.Length), ReadU32(encoded, 8));
        Assert.Equal((uint)(44 + cms.Length), ReadU32(encoded, 12));
        int storeStart = 36 + cms.Length;
        Assert.Equal(0x54524543U, ReadU32(encoded, storeStart + 4));
        Assert.Equal(0x20U, ReadU32(encoded, storeStart + 8));
        Assert.Equal((uint)certificate.Length, ReadU32(encoded, storeStart + 16));
        Assert.Equal(new byte[4], encoded[^4..]);
    }

    [Fact]
    public void V2SignatureDataAcceptsBoundedCompiledAndSourceHashes() {
        byte[] algorithm = Encoding.ASCII.GetBytes(OfficeVbaSignatureEncoding.Sha256Oid + "\0");
        byte[] compiled = Enumerable.Range(0, 32).Select(index => (byte)index).ToArray();
        byte[] source = Enumerable.Range(32, 32).Select(index => (byte)index).ToArray();
        using var output = new MemoryStream();
        using (var writer = new BinaryWriter(output, Encoding.ASCII, leaveOpen: true)) {
            writer.Write(algorithm.Length);
            writer.Write(compiled.Length);
            writer.Write(source.Length);
            writer.Write(24);
            writer.Write(24 + algorithm.Length);
            writer.Write(24 + algorithm.Length + compiled.Length);
            writer.Write(algorithm);
            writer.Write(compiled);
            writer.Write(source);
        }

        bool decoded = OfficeVbaSignatureEncoding.TryExtractV2SourceHash(output.ToArray(),
            OfficeVbaSignatureEncoding.Sha256Oid, out byte[] decodedCompiled,
            out byte[] decodedSource, out string detail);

        Assert.True(decoded, detail);
        Assert.Equal(compiled, decodedCompiled);
        Assert.Equal(source, decodedSource);
    }

    [Fact]
    public void VbaPolicyRejectsStructuralFailuresEvenWhenCryptoPasses() {
        var structuralFailure = new OfficeVbaSignatureFinding("MultipleMacroProjects",
            OfficePackageSignatureValidationState.Failed, "The package contains multiple VBA projects.");
        var signature = new OfficeVbaSignaturePartInfo(OfficeVbaSignatureProfile.V3,
            "/word/vbaProjectSignatureV3.bin", "relationship", "content-type", 1, true,
            OfficePackageSignatureValidationState.Passed, OfficePackageSignatureValidationState.Passed,
            OfficePackageSignatureValidationState.NotChecked, OfficePackageSignatureValidationState.NotPresent,
            "CN=Signer", "thumbprint", OfficeVbaSignatureEncoding.Sha256Oid, new byte[] { 1 },
            Array.Empty<OfficeVbaSignatureFinding>());
        var info = new OfficeVbaSignatureInfo("test.docm", true, true, "/word/vbaProject.bin", 1, "hash",
            new[] { signature }, new[] { structuralFailure });
        var result = new OfficeVbaSignatureValidationResult(info, true,
            OfficePackageSignatureValidationState.Passed, false, Array.Empty<OfficeVbaSignatureFinding>());

        Assert.False(result.IsValidUnderPolicy);
    }

    [VbaOfficeSipInteropFact]
    public void ManagedProfilesMatchFrozenCorpusAndMicrosoftOfficeSipAcrossHosts() {
        string corpus = Environment.GetEnvironmentVariable("OFFICEIMO_VBA_INTEROP_CORPUS")!;
        var documents = new[] {
            new { Name = "Word.docm", Host = "word", Legacy = "C38D392A26CD18BF40E49E55FEA99B93", Agile = "30A396A11684C27E28088E8F53702210235ACFFE242FCBF0C376349A2F32C852", V3 = "8DF63546428827D036376B9AE2C780214988BCCF60A833557175B6FACD15F26B" },
            new { Name = "Excel.xlsm", Host = "xl", Legacy = "71B063DBD601529AF06B5FE30892F4E2", Agile = "0E43ED223FE00053DC7F73C8D11C7209D9EEFD4569C52D7CB7D58DBA86C20572", V3 = "BC7CF7EE7EA94218DEA92AF42F3B8E8D0E76DF46C099886ACC3682CA8E6AA29C" },
            new { Name = "ExcelBinary.xlsb", Host = "xl", Legacy = "CD24111C88662907A4BBB0B0F4A25E96", Agile = "E6090009BD909D817CB0731155A239B8A14218F835A3323AD94BD1020D02B372", V3 = "D0E4DF0572B6BF6BB2521CCB47C614627AF64C2B9B8EF2E0E6F68312512D6B34" },
            new { Name = "PowerPoint.pptm", Host = "ppt", Legacy = "9493F0B77991C6026CFBE09E209F2272", Agile = "EDA11433317FB717D1808EBB37ECE8B14DB13AB6351C6E09638853A0B7401A37", V3 = "E634FD2EC6B02D4BE53991B23A7976D7F255F25561838EE08E2806C4F741BE90" }
        };
        using X509Certificate2 certificate = CreateSigningCertificate();
        foreach (var document in documents) {
            string name = document.Name;
            string host = document.Host;
            string source = Path.Combine(corpus, name);
            byte[] project = ReadZipEntry(source, host + "/vbaProject.bin");
            Assert.True(OfficeVbaProjectCanonicalizer.TryCreate(project, 16 * 1024 * 1024,
                out OfficeVbaProjectCanonicalizer.Result? canonical, out string detail), name + ": " + detail);
            Assert.NotNull(canonical);
            Assert.Equal(document.Legacy, Hex(canonical!.ComputeLegacyHash()));
            Assert.Equal(document.Agile, Hex(canonical.ComputeAgileHash()));
            Assert.Equal(document.V3, Hex(canonical.ComputeV3Hash()));
            string signedPath = Path.Combine(Path.GetTempPath(), $"OfficeIMO-managed-sip-{Guid.NewGuid():N}{Path.GetExtension(name)}");
            try {
                File.Copy(source, signedPath, true);
                var options = new OfficeVbaSigningOptions { ValidateWithWindowsSipWhenAvailable = false };
                options.CmsVerification.CertificateValidation.ChainEvaluator = (_, _) => true;
                options.CmsVerification.CertificateValidation.DisableCertificateDownloads = false;

                OfficeVbaSigningResult signing = SignByExtension(signedPath, certificate, options);
                Assert.True(signing.Succeeded, name + ": " + string.Join(" | ", signing.Findings.Select(
                    finding => finding.Code + ": " + finding.Message)));
                Assert.Equal(3, signing.Validation!.SignatureInfo.Signatures.Count);
                options.ValidateWithWindowsSipWhenAvailable = true;

                foreach (OfficeVbaSignatureProfile profile in Enum.GetValues<OfficeVbaSignatureProfile>()) {
                    string profilePath = Path.Combine(Path.GetTempPath(),
                        $"OfficeIMO-managed-sip-{profile}-{Guid.NewGuid():N}{Path.GetExtension(name)}");
                    try {
                        File.Copy(signedPath, profilePath, true);
                        RemoveHigherProfiles(profilePath, host, profile);
                        OfficeVbaSignatureValidationResult validation = ValidateByExtension(profilePath, options);
                        Assert.True(validation.IsValidUnderPolicy, name + " / " + profile + ": " +
                            string.Join(" | ", validation.Findings.Select(finding => finding.Code + ": " + finding.Message)));
                        Assert.Contains(validation.Findings, finding => finding.Code ==
                            (profile == OfficeVbaSignatureProfile.Agile
                                ? "VbaWindowsSipDifferentialUnavailable"
                                : "VbaWindowsSipDifferentialValid"));
                    } finally {
                        if (File.Exists(profilePath)) File.Delete(profilePath);
                    }
                }
            } finally {
                if (File.Exists(signedPath)) File.Delete(signedPath);
            }
        }
    }

    private static OfficeVbaSigningResult SignByExtension(string path, X509Certificate2 certificate,
        OfficeVbaSigningOptions options) => Path.GetExtension(path).ToLowerInvariant() switch {
        ".docm" => WordDocument.TrySignVbaProject(path, OfficeSecurityProvider.Default, certificate, options),
        ".xlsm" or ".xlsb" => ExcelDocument.TrySignVbaProject(path, OfficeSecurityProvider.Default, certificate, options),
        ".pptm" => PowerPointPresentation.TrySignVbaProject(path, OfficeSecurityProvider.Default, certificate, options),
        _ => throw new ArgumentOutOfRangeException(nameof(path))
    };

    private static OfficeVbaSignatureValidationResult ValidateByExtension(string path,
        OfficeVbaSigningOptions options) => Path.GetExtension(path).ToLowerInvariant() switch {
        ".docm" => WordDocument.ValidateVbaSignatures(path, OfficeSecurityProvider.Default, options),
        ".xlsm" or ".xlsb" => ExcelDocument.ValidateVbaSignatures(path, OfficeSecurityProvider.Default, options),
        ".pptm" => PowerPointPresentation.ValidateVbaSignatures(path, OfficeSecurityProvider.Default, options),
        _ => throw new ArgumentOutOfRangeException(nameof(path))
    };

    private static void RemoveHigherProfiles(string path, string host,
        OfficeVbaSignatureProfile selectedProfile) {
        using ZipArchive archive = ZipFile.Open(path, ZipArchiveMode.Update);
        string directory = host + "/";
        var removed = new List<string>();
        if (selectedProfile < OfficeVbaSignatureProfile.V3) removed.Add(directory + "vbaProjectSignatureV3.bin");
        if (selectedProfile < OfficeVbaSignatureProfile.Agile) removed.Add(directory + "vbaProjectSignatureAgile.bin");
        foreach (string part in removed) archive.GetEntry(part)?.Delete();

        XNamespace relationships = "http://schemas.openxmlformats.org/package/2006/relationships";
        string relationshipPath = directory + "_rels/vbaProject.bin.rels";
        MutateXmlEntry(archive, relationshipPath, document =>
            document.Root!.Elements(relationships + "Relationship").Where(element =>
                removed.Any(part => part.EndsWith('/' + (string?)element.Attribute("Target"),
                    StringComparison.OrdinalIgnoreCase))).Remove());
        XNamespace types = "http://schemas.openxmlformats.org/package/2006/content-types";
        MutateXmlEntry(archive, "[Content_Types].xml", document =>
            document.Root!.Elements(types + "Override").Where(element =>
                removed.Contains(((string?)element.Attribute("PartName") ?? string.Empty).TrimStart('/'),
                    StringComparer.OrdinalIgnoreCase)).Remove());
    }

    private static void MutateXmlEntry(ZipArchive archive, string path, Action<XDocument> mutation) {
        ZipArchiveEntry entry = archive.GetEntry(path)
            ?? throw new InvalidOperationException("Corpus XML entry missing: " + path);
        XDocument document;
        using (Stream input = entry.Open()) document = XDocument.Load(input, LoadOptions.PreserveWhitespace);
        mutation(document);
        entry.Delete();
        ZipArchiveEntry replacement = archive.CreateEntry(path, CompressionLevel.Optimal);
        using Stream output = replacement.Open();
        document.Save(output, SaveOptions.DisableFormatting);
    }

    private static byte[] ReadZipEntry(string path, string entryPath) {
        using ZipArchive archive = ZipFile.OpenRead(path);
        ZipArchiveEntry entry = archive.GetEntry(entryPath)
            ?? throw new InvalidOperationException("Corpus entry missing: " + entryPath);
        using Stream input = entry.Open();
        using var output = new MemoryStream();
        input.CopyTo(output);
        return output.ToArray();
    }

    private static void CreateMacroPackage(string path, string hostRoot) {
        string vbaPath = hostRoot + "/vbaProject.bin";
        string signaturePrefix = hostRoot + "/vbaProjectSignature";
        using ZipArchive archive = ZipFile.Open(path, ZipArchiveMode.Create);
        WriteText(archive, "[Content_Types].xml", ContentTypes(hostRoot));
        WriteBytes(archive, vbaPath, new byte[] { 0xD0, 0xCF, 0x11, 0xE0 });
        WriteText(archive, hostRoot + "/_rels/vbaProject.bin.rels", Relationships());
        WriteBytes(archive, signaturePrefix + ".bin", DigSigInfo());
        WriteBytes(archive, signaturePrefix + "Agile.bin", DigSigInfo());
        WriteBytes(archive, signaturePrefix + "V3.bin", DigSigInfo());
    }

    private static void CreateUnsignedMacroPackage(string path, string hostRoot, byte[] vbaProject) {
        using ZipArchive archive = ZipFile.Open(path, ZipArchiveMode.Create);
        WriteText(archive, "[Content_Types].xml",
            "<?xml version=\"1.0\"?><Types xmlns=\"http://schemas.openxmlformats.org/package/2006/content-types\">" +
            $"<Override PartName=\"/{hostRoot}/vbaProject.bin\" ContentType=\"application/vnd.ms-office.vbaProject\"/>" +
            "</Types>");
        WriteBytes(archive, hostRoot + "/vbaProject.bin", vbaProject);
        WriteText(archive, hostRoot + "/_rels/vbaProject.bin.rels",
            "<?xml version=\"1.0\"?><Relationships xmlns=\"http://schemas.openxmlformats.org/package/2006/relationships\"/>");
    }

    private static byte[] CreateVbaProject(string procedureSuffix) =>
        CreateVbaProject(procedureSuffix, "Module1", 0, null);

    private static byte[] CreateVbaProject(string procedureSuffix, string moduleName,
        uint textOffset, IEnumerable<OfficeCompoundStream>? designerStreams) {
        byte[] directory = CreateDirectoryStream(moduleName, textOffset);
        byte[] source = Encoding.ASCII.GetBytes(
            "Attribute VB_Name = \"" + moduleName + "\"\r\nOption Explicit\r\nSub " +
            procedureSuffix + "()\r\nEnd Sub\r\n");
        bool designer = designerStreams != null;
        var streams = new List<OfficeCompoundStream> {
            new OfficeCompoundStream("PROJECT", Encoding.ASCII.GetBytes(
                "ID=\"{00000000-0000-0000-0000-000000000000}\"\r\n" +
                (designer ? "BaseClass=" : "Module=") + moduleName + "\r\nName=\"VBAProject\"\r\n" +
                "[Host Extender Info]\r\n&H00000001={00000000-0000-0000-0000-000000000000};VBE;&H00000000\r\n[Workspace]\r\n")),
            new OfficeCompoundStream("VBA/dir", CompressLiteral(directory)),
            new OfficeCompoundStream("VBA/_VBA_PROJECT", new byte[] { 0x61, 0xCC, 0x01, 0x00, 0x00, 0x00, 0x00 }),
            new OfficeCompoundStream("VBA/" + moduleName, CompressLiteral(source))
        };
        if (designerStreams != null) streams.AddRange(designerStreams);
        return OfficeCompoundFileWriter.Write(streams);
    }

    private static byte[] CreateDirectoryStream(string moduleName = "Module1", uint textOffset = 0) {
        var bytes = new List<byte>();
        Sized(bytes, 0x0001, new byte[4]);
        Sized(bytes, 0x0002, new byte[] { 0x09, 0x04, 0x00, 0x00 });
        Sized(bytes, 0x0003, new byte[] { 0xE4, 0x04 });
        Sized(bytes, 0x0004, Encoding.ASCII.GetBytes("VBAProject"));
        Sized(bytes, 0x0005, Array.Empty<byte>());
        Sized(bytes, 0x0040, Array.Empty<byte>());
        Sized(bytes, 0x0006, Array.Empty<byte>());
        Sized(bytes, 0x003D, Array.Empty<byte>());
        Sized(bytes, 0x0007, new byte[4]);
        Sized(bytes, 0x0008, new byte[4]);
        U16(bytes, 0x0009); U32(bytes, 4); U32(bytes, 1); U16(bytes, 0);
        Sized(bytes, 0x000C, Array.Empty<byte>());
        Sized(bytes, 0x003C, Array.Empty<byte>());
        Sized(bytes, 0x000F, new byte[] { 1, 0 });
        Sized(bytes, 0x0013, new byte[] { 0, 0 });
        Sized(bytes, 0x0019, Encoding.ASCII.GetBytes(moduleName));
        Sized(bytes, 0x0047, Encoding.Unicode.GetBytes(moduleName));
        Sized(bytes, 0x001A, Encoding.ASCII.GetBytes(moduleName));
        Sized(bytes, 0x0032, Encoding.Unicode.GetBytes(moduleName));
        Sized(bytes, 0x001C, Array.Empty<byte>());
        Sized(bytes, 0x0048, Array.Empty<byte>());
        Sized(bytes, 0x0031, new[] {
            (byte)textOffset, (byte)(textOffset >> 8), (byte)(textOffset >> 16), (byte)(textOffset >> 24)
        });
        Sized(bytes, 0x001E, new byte[4]);
        Sized(bytes, 0x002C, new byte[2]);
        Fixed(bytes, 0x0021);
        Fixed(bytes, 0x002B);
        Fixed(bytes, 0x0010);
        return bytes.ToArray();
    }

    private static byte[] CompressLiteral(byte[] uncompressed) {
        const int maximumLiteralBytesPerChunk = 3_640;
        var result = new List<byte> { 0x01 };
        for (int chunkOffset = 0; chunkOffset < uncompressed.Length; chunkOffset += maximumLiteralBytesPerChunk) {
            int chunkInputLength = Math.Min(maximumLiteralBytesPerChunk, uncompressed.Length - chunkOffset);
            var payload = new List<byte>(chunkInputLength + ((chunkInputLength + 7) / 8));
            for (int offset = 0; offset < chunkInputLength; offset += 8) {
                payload.Add(0);
                int count = Math.Min(8, chunkInputLength - offset);
                for (int index = 0; index < count; index++) payload.Add(uncompressed[chunkOffset + offset + index]);
            }
            int chunkSize = payload.Count + 2;
            ushort header = checked((ushort)(0xB000 | (chunkSize - 3)));
            result.Add((byte)header);
            result.Add((byte)(header >> 8));
            result.AddRange(payload);
        }
        return result.ToArray();
    }

    private static X509Certificate2 CreateSigningCertificate() {
        using RSA rsa = RSA.Create(2048);
        var request = new CertificateRequest("CN=OfficeIMO Managed VBA Test", rsa,
            HashAlgorithmName.SHA256, RSASignaturePadding.Pkcs1);
        request.CertificateExtensions.Add(new X509KeyUsageExtension(X509KeyUsageFlags.DigitalSignature, true));
        request.CertificateExtensions.Add(new X509EnhancedKeyUsageExtension(
            new OidCollection { new Oid("1.3.6.1.5.5.7.3.3") }, false));
        using X509Certificate2 created = request.CreateSelfSigned(
            DateTimeOffset.UtcNow.AddMinutes(-5), DateTimeOffset.UtcNow.AddDays(1));
        byte[] pfx = created.Export(X509ContentType.Pfx);
        const X509KeyStorageFlags flags = X509KeyStorageFlags.Exportable |
                                          X509KeyStorageFlags.EphemeralKeySet;
#if NET9_0_OR_GREATER
        return X509CertificateLoader.LoadPkcs12(pfx, null, flags);
#else
        return new X509Certificate2(pfx, (string?)null, flags);
#endif
    }

    private static void ReplaceVbaProject(string path, string hostRoot, byte[] replacement) {
        using ZipArchive archive = ZipFile.Open(path, ZipArchiveMode.Update);
        ZipArchiveEntry entry = archive.Entries.Single(item =>
            string.Equals(item.FullName, hostRoot + "/vbaProject.bin", StringComparison.OrdinalIgnoreCase));
        entry.Delete();
        WriteBytes(archive, hostRoot + "/vbaProject.bin", replacement);
    }

    private static void Sized(List<byte> bytes, ushort id, byte[] data) {
        U16(bytes, id); U32(bytes, checked((uint)data.Length)); bytes.AddRange(data);
    }

    private static void Fixed(List<byte> bytes, ushort id) { U16(bytes, id); U32(bytes, 0); }
    private static void U16(List<byte> bytes, ushort value) { bytes.Add((byte)value); bytes.Add((byte)(value >> 8)); }
    private static void U32(List<byte> bytes, uint value) {
        bytes.Add((byte)value); bytes.Add((byte)(value >> 8)); bytes.Add((byte)(value >> 16)); bytes.Add((byte)(value >> 24));
    }

    private static uint ReadU32(byte[] bytes, int offset) =>
        (uint)(bytes[offset] | bytes[offset + 1] << 8 | bytes[offset + 2] << 16 | bytes[offset + 3] << 24);

    private static string Hex(byte[] bytes) => BitConverter.ToString(bytes).Replace("-", string.Empty);

    private static string ContentTypes(string hostRoot) =>
        "<?xml version=\"1.0\"?><Types xmlns=\"http://schemas.openxmlformats.org/package/2006/content-types\">" +
        $"<Override PartName=\"/{hostRoot}/vbaProject.bin\" ContentType=\"application/vnd.ms-office.vbaProject\"/>" +
        $"<Override PartName=\"/{hostRoot}/vbaProjectSignature.bin\" ContentType=\"application/vnd.ms-office.vbaProjectSignature\"/>" +
        $"<Override PartName=\"/{hostRoot}/vbaProjectSignatureAgile.bin\" ContentType=\"application/vnd.ms-office.vbaProjectSignatureAgile\"/>" +
        $"<Override PartName=\"/{hostRoot}/vbaProjectSignatureV3.bin\" ContentType=\"application/vnd.ms-office.vbaProjectSignatureV3\"/>" +
        "</Types>";

    private static string Relationships() =>
        "<?xml version=\"1.0\"?><Relationships xmlns=\"http://schemas.openxmlformats.org/package/2006/relationships\">" +
        "<Relationship Id=\"rId1\" Type=\"http://schemas.microsoft.com/office/2006/relationships/vbaProjectSignature\" Target=\"vbaProjectSignature.bin\"/>" +
        "<Relationship Id=\"rId2\" Type=\"http://schemas.microsoft.com/office/2014/relationships/vbaProjectSignatureAgile\" Target=\"vbaProjectSignatureAgile.bin\"/>" +
        "<Relationship Id=\"rId3\" Type=\"http://schemas.microsoft.com/office/2020/07/relationships/vbaProjectSignatureV3\" Target=\"vbaProjectSignatureV3.bin\"/>" +
        "</Relationships>";

    private static byte[] DigSigInfo() {
        var bytes = new byte[37];
        bytes[0] = 1;
        bytes[4] = 44;
        bytes[36] = 0x30;
        return bytes;
    }

    private static void WriteText(ZipArchive archive, string path, string text) =>
        WriteBytes(archive, path, Encoding.UTF8.GetBytes(text));

    private static void WriteBytes(ZipArchive archive, string path, byte[] bytes) {
        ZipArchiveEntry entry = archive.CreateEntry(path);
        using Stream output = entry.Open();
        output.Write(bytes, 0, bytes.Length);
    }
}
