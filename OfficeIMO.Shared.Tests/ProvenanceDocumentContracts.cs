using System.IO.Compression;
using OfficeIMO;
using OfficeIMO.Epub;
using OfficeIMO.Excel;
using OfficeIMO.Html;
using OfficeIMO.Markdown;
using OfficeIMO.OpenDocument;
using OfficeIMO.PowerPoint;
using OfficeIMO.Provenance;
using OfficeIMO.Word;
using Xunit;

namespace OfficeIMO.Shared.Tests;

public sealed class ProvenanceDocumentContracts {
    [Fact]
    public void HtmlRemovesNativeCarriersAndEmbeddedImageProvenanceOffline() {
        byte[] manifest = CreateManifestStore();
        byte[] image = CreatePngWithManifest(manifest);
        string html = "<!doctype html><html><head>" +
            "<link rel=\"stylesheet c2pa-manifest\" href=\"https://example.test/claim.c2pa\">" +
            "</head><body>" +
            $"<img src=\"data:image/png;base64,{Convert.ToBase64String(image)}\" alt=\"kept\">" +
            "</body></html>";

        OfficeProvenanceReport report = HtmlProvenance.Inspect(html);
        OfficeProvenanceRemovalResult result = HtmlProvenance.Remove(html);
        string output = Encoding.UTF8.GetString(result.ToArray());

        Assert.Equal(2, report.Evidence.Count);
        Assert.Contains(report.Evidence, item => item.Location.StartsWith("HTML/img[src]", StringComparison.Ordinal));
        Assert.Empty(result.After.Evidence);
        Assert.DoesNotContain("c2pa-manifest", output, StringComparison.OrdinalIgnoreCase);
        Assert.StartsWith("<!DOCTYPE html>", output, StringComparison.OrdinalIgnoreCase);
        Assert.Contains("rel=\"stylesheet\"", output, StringComparison.OrdinalIgnoreCase);
        Assert.Contains("href=\"https://example.test/claim.c2pa\"", output, StringComparison.OrdinalIgnoreCase);
        Assert.Contains("alt=\"kept\"", output, StringComparison.Ordinal);
    }

    [Fact]
    public void HtmlRemovesASingleEmbeddedManifestAssociation() {
        string html = $"<!doctype html><html><head><script type=\"application/c2pa\">{Convert.ToBase64String(CreateManifestStore())}</script></head><body>kept</body></html>";

        OfficeProvenanceRemovalResult result = HtmlProvenance.Remove(html);
        string output = Encoding.UTF8.GetString(result.ToArray());

        Assert.True(result.WasChanged);
        Assert.Empty(result.After.Evidence);
        Assert.DoesNotContain("application/c2pa", output, StringComparison.OrdinalIgnoreCase);
        Assert.Contains("kept", output, StringComparison.Ordinal);
    }

    [Fact]
    public void HtmlPreservesMultipleManifestAssociationsByDefault() {
        string manifest = Convert.ToBase64String(CreateManifestStore());
        string html = $"<html><head><script type=\"application/c2pa\">{manifest}</script><link rel=\"c2pa-manifest\" href=\"claim.c2pa\"></head><body></body></html>";

        OfficeProvenanceRemovalResult result = HtmlProvenance.Remove(html);

        Assert.False(result.WasChanged);
        Assert.Equal(2, result.Before.Evidence.Count);
        Assert.All(result.Before.Evidence, item => Assert.False(item.IsStructurallyValid));
        Assert.Contains(result.Before.Diagnostics, item => item.Contains("manifest.html.multipleManifests", StringComparison.Ordinal));
    }

    [Fact]
    public void HtmlPreservesMalformedNativeCarrierByDefault() {
        const string html = "<html><head><script type=\"application/c2pa\">not-base64</script></head><body>ok</body></html>";

        OfficeProvenanceRemovalResult result = HtmlProvenance.Remove(html);

        Assert.False(result.WasChanged);
        Assert.Single(result.Before.Evidence);
        Assert.False(result.Before.Evidence[0].IsStructurallyValid);
        Assert.Equal(html, Encoding.UTF8.GetString(result.ToArray()));
    }

    [Fact]
    public void HtmlUsesOnlyHeadAssociationsAndAcceptsSafeRelativeReferences() {
        string html = "<html><head><link rel=\"c2pa-manifest\" href=\"claims/active.c2pa\"></head>" +
            "<body><script type=\"application/c2pa\">not-a-head-carrier</script></body></html>";

        OfficeProvenanceReport report = HtmlProvenance.Inspect(html);
        OfficeProvenanceRemovalResult result = HtmlProvenance.Remove(html);
        string output = Encoding.UTF8.GetString(result.ToArray());

        OfficeProvenanceEvidence evidence = Assert.Single(report.Evidence);
        Assert.True(evidence.IsStructurallyValid);
        Assert.Equal("claims/active.c2pa", evidence.Value);
        Assert.DoesNotContain("rel=\"c2pa-manifest\"", output, StringComparison.OrdinalIgnoreCase);
        Assert.Contains("not-a-head-carrier", output, StringComparison.Ordinal);
    }

    [Fact]
    public void HtmlFileRemovalPreservesDetectedLegacyEncoding() {
        Encoding.RegisterProvider(CodePagesEncodingProvider.Instance);
        Encoding windows1252 = Encoding.GetEncoding(1252);
        string html = "<!doctype html><html><head><meta charset=\"windows-1252\"><link rel=\"c2pa-manifest\" href=\"claim.c2pa\"></head><body>café</body></html>";
        string inputPath = Path.Combine(Path.GetTempPath(), $"OfficeIMO-Provenance-{Guid.NewGuid():N}.html");
        string outputPath = Path.Combine(Path.GetTempPath(), $"OfficeIMO-Provenance-{Guid.NewGuid():N}.html");
        try {
            File.WriteAllBytes(inputPath, windows1252.GetBytes(html));

            OfficeProvenanceReport report = HtmlProvenance.InspectFile(inputPath);
            OfficeProvenanceRemovalResult result = HtmlProvenance.RemoveFile(inputPath, outputPath);
            string output = windows1252.GetString(File.ReadAllBytes(outputPath));

            Assert.Single(report.Evidence);
            Assert.True(result.WasChanged);
            Assert.Contains("café", output, StringComparison.Ordinal);
            Assert.DoesNotContain("c2pa-manifest", output, StringComparison.OrdinalIgnoreCase);
        } finally {
            if (File.Exists(inputPath)) File.Delete(inputPath);
            if (File.Exists(outputPath)) File.Delete(outputPath);
        }
    }

    [Fact]
    public void HtmlFileRemovalEscapesCharactersOutsideTheLegacyEncoding() {
        Encoding.RegisterProvider(CodePagesEncodingProvider.Instance);
        Encoding windows1252 = Encoding.GetEncoding(1252);
        string html = "<!doctype html><html><head><meta charset=\"windows-1252\"><link rel=\"c2pa-manifest\" href=\"claim.c2pa\"></head><body>&#x2603;</body></html>";
        string inputPath = Path.Combine(Path.GetTempPath(), $"OfficeIMO-Provenance-{Guid.NewGuid():N}.html");
        string outputPath = Path.Combine(Path.GetTempPath(), $"OfficeIMO-Provenance-{Guid.NewGuid():N}.html");
        try {
            File.WriteAllBytes(inputPath, windows1252.GetBytes(html));

            HtmlProvenance.RemoveFile(inputPath, outputPath);
            string output = windows1252.GetString(File.ReadAllBytes(outputPath));

            Assert.Contains("&#x2603;", output, StringComparison.OrdinalIgnoreCase);
            Assert.DoesNotContain("?", output, StringComparison.Ordinal);
        } finally {
            if (File.Exists(inputPath)) File.Delete(inputPath);
            if (File.Exists(outputPath)) File.Delete(outputPath);
        }
    }

    [Fact]
    public void HtmlFileInspectionUsesTheBoundedSourceEncodingSize() {
        Encoding.RegisterProvider(CodePagesEncodingProvider.Instance);
        Encoding windows1252 = Encoding.GetEncoding(1252);
        string html = "<html><head><meta charset=\"windows-1252\"></head><body>café</body></html>";
        byte[] data = windows1252.GetBytes(html);
        string inputPath = Path.Combine(Path.GetTempPath(), $"OfficeIMO-Provenance-{Guid.NewGuid():N}.html");
        try {
            File.WriteAllBytes(inputPath, data);

            OfficeProvenanceReport report = HtmlProvenance.InspectFile(inputPath, new OfficeProvenanceOptions {
                MaxAssetBytes = data.Length,
                MaxManifestBytes = data.Length
            });

            Assert.Equal(OfficeProvenanceAssetFormat.Html, report.Format);
        } finally {
            if (File.Exists(inputPath)) File.Delete(inputPath);
        }
    }

    [Fact]
    public void HtmlMalformedEmbeddedDataUriIsDiagnosticInsteadOfAnException() {
        const string html = "<html><head></head><body><img src=\"data:image/png,%ZZ\"></body></html>";

        OfficeProvenanceReport report = HtmlProvenance.Inspect(html);
        OfficeProvenanceRemovalResult result = HtmlProvenance.Remove(html);

        Assert.Contains(report.Diagnostics, item => item.Contains("could not be decoded", StringComparison.Ordinal));
        Assert.False(result.WasChanged);
    }

    [Fact]
    public void HtmlSanitizesEmbeddedImagesInResponsiveSourceSets() {
        byte[] image = CreatePngWithManifest(CreateManifestStore());
        string dataUri = "data:image/png;base64," + Convert.ToBase64String(image);
        string html = $"<html><head></head><body><picture><source srcset=\"{dataUri} 1x, image.png 2x\"></picture></body></html>";

        OfficeProvenanceRemovalResult result = HtmlProvenance.Remove(html);
        string output = Encoding.UTF8.GetString(result.ToArray());

        Assert.Contains(result.Before.Evidence, item => item.Location.Contains("[srcset]", StringComparison.Ordinal));
        Assert.Empty(result.After.Evidence);
        Assert.Contains("image.png 2x", output, StringComparison.Ordinal);
    }

    [Fact]
    public void HtmlRemovalSkipsEmbeddedAssetsWhenDisabled() {
        string html = "<html><head><link rel=\"c2pa-manifest\" href=\"claim.c2pa\"></head>" +
            "<body><img src=\"data:image/png;base64," + new string('A', 512) + "\"></body></html>";
        var options = new OfficeProvenanceRemovalOptions { ProcessEmbeddedAssets = false };
        options.Limits.MaxAssetBytes = Encoding.UTF8.GetByteCount(html) + 32;
        options.Limits.MaxManifestBytes = 32;

        OfficeProvenanceRemovalResult result = HtmlProvenance.Remove(html, options);

        Assert.True(result.WasChanged);
        Assert.Empty(result.After.Evidence);
    }

    [Fact]
    public void MarkdownUsesTheSharedStructuredTextContract() {
        string markdown = "# Before\n\n-----BEGIN C2PA MANIFEST-----\n" +
            "data:application/c2pa;base64," + Convert.ToBase64String(CreateManifestStore()) + "\n" +
            "-----END C2PA MANIFEST-----\n\nAfter\n";

        OfficeProvenanceRemovalResult result = MarkdownProvenance.Remove(markdown);

        Assert.True(result.WasChanged);
        Assert.Empty(result.After.Evidence);
        Assert.Equal("# Before\n\n\nAfter\n", Encoding.UTF8.GetString(result.ToArray()));
    }

    [Theory]
    [InlineData("docx")]
    [InlineData("xlsx")]
    [InlineData("pptx")]
    public void OpenXmlOwnerApisSanitizeEmbeddedImages(string extension) {
        string path = Path.Combine(Path.GetTempPath(), $"OfficeIMO-Provenance-{Guid.NewGuid():N}.{extension}");
        try {
            CreateOpenXmlPackage(path, extension);
            AddZipEntry(path, "media/provenance.png", CreatePngWithManifest(CreateManifestStore()));
            byte[] package = File.ReadAllBytes(path);

            OfficeProvenanceReport report = InspectOpenXml(path, extension);
            OfficeProvenanceRemovalResult result = RemoveOpenXml(package, extension);

            Assert.Single(report.Evidence);
            Assert.True(result.WasChanged);
            Assert.Empty(result.After.Evidence);
            Assert.Empty(OfficeProvenanceInspector.Inspect(ReadZipEntry(result.ToArray(), "media/provenance.png"), "image.png").Evidence);
        } finally {
            if (File.Exists(path)) File.Delete(path);
        }
    }

    [Fact]
    public void GenericZipPreservePolicyDoesNotParseMalformedOpcSignatureMetadata() {
        using var output = new MemoryStream();
        using (var archive = new ZipArchive(output, ZipArchiveMode.Create, leaveOpen: true)) {
            WriteEntry(archive, "[Content_Types].xml", "<Types", CompressionLevel.Optimal);
            WriteEntry(archive, "media/provenance.png", CreatePngWithManifest(CreateManifestStore()), CompressionLevel.Optimal);
        }

        OfficeProvenanceRemovalResult result = OfficeProvenanceRemover.Remove(output.ToArray(), "package.zip", new OfficeProvenanceRemovalOptions {
            SignatureMutationPolicy = OfficeSignatureMutationPolicy.PreserveSignatureMarkup
        });

        Assert.True(result.WasChanged);
        Assert.Empty(result.After.Evidence);
    }

    [Fact]
    public void OpenXmlOwnerFailsClosedWhenOrphanSignatureEvidenceCannotBeRemoved() {
        string path = Path.Combine(Path.GetTempPath(), $"OfficeIMO-Provenance-{Guid.NewGuid():N}.docx");
        try {
            CreateOpenXmlPackage(path, "docx");
            AddZipEntry(path, "_xmlsignatures/orphan.xml", Encoding.UTF8.GetBytes("<signature/>"));
            AddZipEntry(path, "media/provenance.png", CreatePngWithManifest(CreateManifestStore()));
            byte[] package = File.ReadAllBytes(path);

            InvalidOperationException exception = Assert.Throws<InvalidOperationException>(() =>
                WordDocument.RemoveProvenance(package, options: new OfficeProvenanceRemovalOptions {
                    SignatureMutationPolicy = OfficeSignatureMutationPolicy.RemoveInvalidatedSignatures
                }));

            Assert.Contains("could not remove", exception.Message, StringComparison.Ordinal);
        } finally {
            if (File.Exists(path)) File.Delete(path);
        }
    }

    [Theory]
    [InlineData("odt", "META-INF/customsignatures.xml")]
    [InlineData("epub", "META-INF/signatures.xml")]
    public void ZipDocumentOwnersRemoveInvalidatedNativeSignatures(string extension, string signaturePath) {
        byte[] package = CreateZipPackage(extension, signaturePath, CreatePngWithManifest(CreateManifestStore()));
        var options = new OfficeProvenanceRemovalOptions {
            SignatureMutationPolicy = OfficeSignatureMutationPolicy.RemoveInvalidatedSignatures
        };

        OfficeProvenanceRemovalResult result = extension == "odt"
            ? OdfDocument.RemoveProvenance(package, "document.odt", options)
            : EpubDocument.RemoveProvenance(package, "publication.epub", options);

        using var archive = new ZipArchive(new MemoryStream(result.ToArray()), ZipArchiveMode.Read);
        Assert.True(result.WereInvalidatedSignaturesRemoved);
        Assert.DoesNotContain(archive.Entries, entry => entry.FullName.Equals(signaturePath, StringComparison.OrdinalIgnoreCase));
        Assert.Empty(result.After.Evidence);
        Assert.Equal("mimetype", archive.Entries[0].FullName);
        Assert.Equal(CompressionMethodStored, ReadCompressionMethod(result.ToArray(), archive.Entries[0].FullName));
    }

    private const ushort CompressionMethodStored = 0;

    private static void CreateOpenXmlPackage(string path, string extension) {
        switch (extension) {
            case "docx":
                using (WordDocument document = WordDocument.Create(path)) {
                    document.AddParagraph("provenance fixture");
                    document.Save();
                }
                break;
            case "xlsx":
                using (ExcelDocument document = ExcelDocument.Create(path)) {
                    document.AddWorksheet("Data").CellValue(1, 1, "provenance fixture");
                    document.Save();
                }
                break;
            case "pptx":
                using (PowerPointPresentation presentation = PowerPointPresentation.Create(path)) presentation.Save();
                break;
            default:
                throw new ArgumentOutOfRangeException(nameof(extension));
        }
    }

    private static OfficeProvenanceReport InspectOpenXml(string path, string extension) => extension switch {
        "docx" => WordDocument.InspectProvenance(path),
        "xlsx" => ExcelDocument.InspectProvenance(path),
        "pptx" => PowerPointPresentation.InspectProvenance(path),
        _ => throw new ArgumentOutOfRangeException(nameof(extension))
    };

    private static OfficeProvenanceRemovalResult RemoveOpenXml(byte[] package, string extension) => extension switch {
        "docx" => WordDocument.RemoveProvenance(package),
        "xlsx" => ExcelDocument.RemoveProvenance(package),
        "pptx" => PowerPointPresentation.RemoveProvenance(package),
        _ => throw new ArgumentOutOfRangeException(nameof(extension))
    };

    private static void AddZipEntry(string path, string entryName, byte[] data) {
        using ZipArchive archive = ZipFile.Open(path, ZipArchiveMode.Update);
        ZipArchiveEntry entry = archive.CreateEntry(entryName, CompressionLevel.Optimal);
        using Stream output = entry.Open();
        output.Write(data, 0, data.Length);
    }

    private static byte[] CreateZipPackage(string extension, string signaturePath, byte[] image) {
        using var output = new MemoryStream();
        using (var archive = new ZipArchive(output, ZipArchiveMode.Create, leaveOpen: true)) {
            WriteEntry(archive, "mimetype", extension == "odt" ? "application/vnd.oasis.opendocument.text" : "application/epub+zip", CompressionLevel.NoCompression);
            WriteEntry(archive, signaturePath, "<signatures/>", CompressionLevel.Optimal);
            WriteEntry(archive, signaturePath, "<signatures duplicate=\"true\"/>", CompressionLevel.Optimal);
            WriteEntry(archive, "media/provenance.png", image, CompressionLevel.Optimal);
        }
        return output.ToArray();
    }

    private static void WriteEntry(ZipArchive archive, string name, string content, CompressionLevel level) =>
        WriteEntry(archive, name, Encoding.UTF8.GetBytes(content), level);

    private static void WriteEntry(ZipArchive archive, string name, byte[] content, CompressionLevel level) {
        ZipArchiveEntry entry = archive.CreateEntry(name, level);
        using Stream stream = entry.Open();
        stream.Write(content, 0, content.Length);
    }

    private static byte[] ReadZipEntry(byte[] package, string name) {
        using var archive = new ZipArchive(new MemoryStream(package), ZipArchiveMode.Read);
        using Stream input = (archive.GetEntry(name) ?? throw new InvalidOperationException("Missing ZIP entry: " + name)).Open();
        using var output = new MemoryStream();
        input.CopyTo(output);
        return output.ToArray();
    }

    private static ushort ReadCompressionMethod(byte[] package, string entryName) {
        byte[] name = Encoding.UTF8.GetBytes(entryName);
        int offset = 0;
        while (offset <= package.Length - 30) {
            if (BitConverter.ToUInt32(package, offset) != 0x04034B50) break;
            ushort method = BitConverter.ToUInt16(package, offset + 8);
            ushort nameLength = BitConverter.ToUInt16(package, offset + 26);
            ushort extraLength = BitConverter.ToUInt16(package, offset + 28);
            string currentName = Encoding.UTF8.GetString(package, offset + 30, nameLength);
            if (currentName == entryName) return method;
            uint compressedLength = BitConverter.ToUInt32(package, offset + 18);
            offset += 30 + nameLength + extraLength + checked((int)compressedLength);
        }
        throw new InvalidDataException("ZIP local header was not found: " + entryName);
    }

    private static byte[] CreatePngWithManifest(byte[] manifest) {
        byte[] header = { 0x89, 0x50, 0x4E, 0x47, 0x0D, 0x0A, 0x1A, 0x0A };
        return Join(header, CreatePngChunk("caBX", manifest), CreatePngChunk("IEND", Array.Empty<byte>()));
    }

    private static byte[] CreateManifestStore() {
        byte[] data = new byte[38];
        WriteBigEndian(data, 0, data.Length);
        Encoding.ASCII.GetBytes("jumb").CopyTo(data, 4);
        WriteBigEndian(data, 8, 30);
        Encoding.ASCII.GetBytes("jumd").CopyTo(data, 12);
        new byte[] { 0x63, 0x32, 0x70, 0x61, 0x00, 0x11, 0x00, 0x10, 0x80, 0x00, 0x00, 0xAA, 0x00, 0x38, 0x9B, 0x71 }.CopyTo(data, 16);
        data[32] = 0x02;
        Encoding.ASCII.GetBytes("c2pa").CopyTo(data, 33);
        return data;
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
