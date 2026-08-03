using System.IO;
using System.IO.Compression;
using System.Text;
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
