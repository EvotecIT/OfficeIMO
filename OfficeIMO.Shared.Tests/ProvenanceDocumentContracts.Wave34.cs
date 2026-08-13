using System.IO.Compression;
using System.Net;
using System.Text;
using OfficeIMO.Excel;
using OfficeIMO.Html;
using OfficeIMO.Provenance;
using Xunit;

namespace OfficeIMO.Shared.Tests;

public sealed partial class ProvenanceDocumentContracts {
    [Fact]
    public void ExcelXlsbDetectionHonorsConfiguredXmlNodeLimits() {
        using var output = new MemoryStream();
        using (var archive = new ZipArchive(output, ZipArchiveMode.Create, leaveOpen: true)) {
            WriteWave34Entry(archive, "[Content_Types].xml",
                "<Types xmlns=\"http://schemas.openxmlformats.org/package/2006/content-types\">" +
                string.Concat(Enumerable.Repeat("<Default Extension=\"x\" ContentType=\"application/octet-stream\"/>", 6)) +
                "<Default Extension=\"bin\" ContentType=\"application/vnd.ms-excel.sheet.binary.macroEnabled.main\"/>" +
                "</Types>");
            WriteWave34Entry(archive, "_rels/.rels",
                "<Relationships xmlns=\"http://schemas.openxmlformats.org/package/2006/relationships\">" +
                "<Relationship Id=\"rId1\" Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument\" Target=\"xl/workbook.bin\"/>" +
                "</Relationships>");
            WriteWave34Entry(archive, "xl/workbook.bin", new byte[] { 0x83, 0x01, 0x00 });
        }
        var options = new OfficeProvenanceRemovalOptions();
        options.Limits.MaxContainerEntries = 4;

        Assert.Throws<InvalidDataException>(() => ExcelDocument.RemoveProvenance(output.ToArray(), "workbook.xlsb", options));
    }

    [Fact]
    public void HtmlRejectsSrcdocNestingBeyondTheInspectionBoundary() {
        string nested = "<p>deep</p>";
        for (int index = 0; index < 9; index++) {
            nested = $"<iframe srcdoc=\"{WebUtility.HtmlEncode(nested)}\"></iframe>";
        }

        Assert.Throws<InvalidDataException>(() => HtmlProvenance.Inspect(nested));
        Assert.Throws<InvalidDataException>(() => HtmlProvenance.Remove(nested));
    }

    private static void WriteWave34Entry(ZipArchive archive, string name, string content) =>
        WriteWave34Entry(archive, name, Encoding.UTF8.GetBytes(content));

    private static void WriteWave34Entry(ZipArchive archive, string name, byte[] content) {
        ZipArchiveEntry entry = archive.CreateEntry(name, CompressionLevel.Optimal);
        using Stream stream = entry.Open();
        stream.Write(content, 0, content.Length);
    }
}
