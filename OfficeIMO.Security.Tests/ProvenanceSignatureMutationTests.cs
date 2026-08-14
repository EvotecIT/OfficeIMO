using System.IO;
using System.IO.Compression;
using System.Xml.Linq;
using OfficeIMO;
using OfficeIMO.Provenance;
using OfficeIMO.Security;
using OfficeIMO.Visio;
using OfficeIMO.Word;
using Xunit;

namespace OfficeIMO.Security.Tests;

public sealed class ProvenanceSignatureMutationTests {
    [Fact]
    public void WordRemovalCanExplicitlyStripARealInvalidatedPackageSignature() {
        string path = Path.Combine(Path.GetTempPath(), $"OfficeIMO-Provenance-{Guid.NewGuid():N}.docx");
        try {
            using (WordDocument document = WordDocument.Create(path)) {
                document.AddParagraph("signed provenance fixture");
                document.Save();
            }
            AddZipEntry(path, "word/media/provenance.png", CreatePngWithManifest());
            using X509Certificate2 certificate = CreateCertificate();
            OfficePackageSigningResult signed = WordDocument.SignPackageSignature(path, OfficeSecurityProvider.Default, certificate);
            Assert.True(signed.Succeeded, string.Join(" ", signed.Details));
            Assert.True(WordDocument.InspectPackageSignatures(path).HasSignatures);

            var options = new OfficeProvenanceRemovalOptions {
                SignatureMutationPolicy = OfficeSignatureMutationPolicy.RemoveInvalidatedSignatures
            };
            OfficeProvenanceRemovalResult result = WordDocument.RemoveProvenance(File.ReadAllBytes(path), "document.docx", options);

            Assert.True(result.WasChanged);
            Assert.True(result.WereInvalidatedSignaturesRemoved);
            Assert.Empty(result.After.Evidence);
            Assert.False(OfficePackageSignatureService.Inspect(result.ToArray()).HasSignatures);
        } finally {
            if (File.Exists(path)) File.Delete(path);
        }
    }

    [Fact]
    public void VisioRemovalCanExplicitlyStripARealInvalidatedPackageSignature() {
        string path = Path.Combine(Path.GetTempPath(), $"OfficeIMO-Provenance-{Guid.NewGuid():N}.vsdx");
        try {
            VisioDocument.Create(path).Save();
            AddZipEntry(path, "visio/media/provenance.png", CreatePngWithManifest());
            using X509Certificate2 certificate = CreateCertificate();
            OfficePackageSigningResult signed = VisioDocument.SignPackageSignature(path, OfficeSecurityProvider.Default, certificate);
            Assert.True(signed.Succeeded, string.Join(" ", signed.Details));

            var options = new OfficeProvenanceRemovalOptions {
                SignatureMutationPolicy = OfficeSignatureMutationPolicy.RemoveInvalidatedSignatures
            };
            OfficeProvenanceRemovalResult result = VisioDocument.RemoveProvenance(File.ReadAllBytes(path), "drawing.vsdx", options);

            Assert.True(result.WasChanged);
            Assert.True(result.WereInvalidatedSignaturesRemoved);
            Assert.Empty(result.After.Evidence);
            Assert.False(OfficePackageSignatureService.Inspect(result.ToArray()).HasSignatures);
        } finally {
            if (File.Exists(path)) File.Delete(path);
        }
    }

    private static void AddZipEntry(string path, string entryName, byte[] data) {
        using ZipArchive archive = ZipFile.Open(path, ZipArchiveMode.Update);
        ZipArchiveEntry contentTypes = archive.GetEntry("[Content_Types].xml")
            ?? throw new InvalidOperationException("Content types part is missing.");
        XDocument types;
        using (Stream input = contentTypes.Open()) types = XDocument.Load(input, LoadOptions.PreserveWhitespace);
        XNamespace ns = "http://schemas.openxmlformats.org/package/2006/content-types";
        if (!types.Root!.Elements(ns + "Default").Any(item =>
                string.Equals((string?)item.Attribute("Extension"), "png", StringComparison.OrdinalIgnoreCase))) {
            types.Root.Add(new XElement(ns + "Default",
                new XAttribute("Extension", "png"),
                new XAttribute("ContentType", "image/png")));
            contentTypes.Delete();
            contentTypes = archive.CreateEntry("[Content_Types].xml", CompressionLevel.Optimal);
            using Stream contentTypesOutput = contentTypes.Open();
            types.Save(contentTypesOutput, SaveOptions.DisableFormatting);
        }
        ZipArchiveEntry entry = archive.CreateEntry(entryName, CompressionLevel.Optimal);
        using Stream output = entry.Open();
        output.Write(data, 0, data.Length);
    }

    private static X509Certificate2 CreateCertificate() {
        using RSA rsa = RSA.Create(2048);
        var request = new CertificateRequest("CN=OfficeIMO Provenance Test", rsa, HashAlgorithmName.SHA256, RSASignaturePadding.Pkcs1);
        request.CertificateExtensions.Add(new X509KeyUsageExtension(X509KeyUsageFlags.DigitalSignature, true));
        return request.CreateSelfSigned(DateTimeOffset.UtcNow.AddMinutes(-1), DateTimeOffset.UtcNow.AddDays(1));
    }

    private static byte[] CreatePngWithManifest() {
        byte[] storeDescription = CreateBox("jumd", Join(C2paUuid("c2pa"), new byte[] { 0x03 }, Encoding.ASCII.GetBytes("c2pa\0")));
        byte[] manifestDescription = CreateBox("jumd", Join(C2paUuid("c2ma"), new byte[] { 0x03 }, Encoding.ASCII.GetBytes("m\0")));
        byte[] assertionStoreDescription = CreateBox("jumd", Join(C2paUuid("c2as"), new byte[] { 0x03 }, Encoding.ASCII.GetBytes("c2pa.assertions\0")));
        byte[] assertionDescription = CreateBox("jumd", Join(C2paUuid("c2ac"), new byte[] { 0x03 }, Encoding.ASCII.GetBytes("c2pa.test\0")));
        byte[] assertionStore = CreateBox("jumb", Join(assertionStoreDescription,
            CreateBox("jumb", Join(assertionDescription, CreateBox("cbor", new byte[] { 0xA0 }))))) ;
        byte[] claimDescription = CreateBox("jumd", Join(C2paUuid("c2cl"), new byte[] { 0x03 }, Encoding.ASCII.GetBytes("c2pa.claim\0")));
        byte[] claim = CreateBox("jumb", Join(claimDescription, CreateBox("cbor", new byte[] { 0xA0 })));
        byte[] signatureDescription = CreateBox("jumd", Join(C2paUuid("c2cs"), new byte[] { 0x03 }, Encoding.ASCII.GetBytes("c2pa.signature\0")));
        byte[] signature = CreateBox("jumb", Join(signatureDescription, CreateBox("cbor", new byte[] { 0xA0 })));
        byte[] manifest = CreateBox("jumb", Join(storeDescription, CreateBox("jumb", Join(manifestDescription, assertionStore, claim, signature))));
        byte[] header = { 0x89, 0x50, 0x4E, 0x47, 0x0D, 0x0A, 0x1A, 0x0A };
        byte[] imageHeader = { 0, 0, 0, 1, 0, 0, 0, 1, 8, 2, 0, 0, 0 };
        return Join(
            header,
            CreatePngChunk("IHDR", imageHeader),
            CreatePngChunk("caBX", manifest),
            CreatePngChunk("IDAT", Array.Empty<byte>()),
            CreatePngChunk("IEND", Array.Empty<byte>()));
    }

    private static byte[] C2paUuid(string code) => Join(
        Encoding.ASCII.GetBytes(code),
        new byte[] { 0x00, 0x11, 0x00, 0x10, 0x80, 0x00, 0x00, 0xAA, 0x00, 0x38, 0x9B, 0x71 });

    private static byte[] CreateBox(string type, byte[] payload) {
        byte[] box = new byte[payload.Length + 8];
        WriteBigEndian(box, 0, box.Length);
        Encoding.ASCII.GetBytes(type).CopyTo(box, 4);
        payload.CopyTo(box, 8);
        return box;
    }

    private static byte[] CreatePngChunk(string type, byte[] payload) {
        byte[] chunk = new byte[payload.Length + 12];
        WriteBigEndian(chunk, 0, payload.Length);
        Encoding.ASCII.GetBytes(type).CopyTo(chunk, 4);
        payload.CopyTo(chunk, 8);
        WriteBigEndian(chunk, chunk.Length - 4, unchecked((int)ComputePngCrc(chunk, 4, payload.Length + 4)));
        return chunk;
    }

    private static uint ComputePngCrc(byte[] data, int offset, int count) {
        uint crc = 0xFFFFFFFF;
        for (int index = offset; index < offset + count; index++) {
            crc ^= data[index];
            for (int bit = 0; bit < 8; bit++) crc = (crc & 1) != 0 ? 0xEDB88320U ^ (crc >> 1) : crc >> 1;
        }
        return crc ^ 0xFFFFFFFF;
    }

    private static void WriteBigEndian(byte[] data, int offset, int value) {
        data[offset] = (byte)(value >> 24);
        data[offset + 1] = (byte)(value >> 16);
        data[offset + 2] = (byte)(value >> 8);
        data[offset + 3] = (byte)value;
    }

    private static byte[] Join(params byte[][] arrays) {
        byte[] output = new byte[arrays.Sum(item => item.Length)];
        int offset = 0;
        foreach (byte[] item in arrays) {
            Buffer.BlockCopy(item, 0, output, offset, item.Length);
            offset += item.Length;
        }
        return output;
    }
}
