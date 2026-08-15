using System.IO.Compression;
using OfficeIMO.Excel;
using OfficeIMO.Html;
using OfficeIMO.Provenance;
using Xunit;

namespace OfficeIMO.Shared.Tests;

public sealed partial class ProvenanceDocumentContracts {
    [Theory]
    [InlineData("")]
    [InlineData("normal")]
    [InlineData("none")]
    public void NonGeneratedPseudoElementsDoNotOwnImageCarriers(string content) {
        string dataUri = "data:image/png;base64," +
            Convert.ToBase64String(CreatePngWithManifest(CreateManifestStore()));
        string contentDeclaration = content.Length == 0 ? string.Empty : "content:" + content + ";";
        string html = "<style>.box::before{" + contentDeclaration + "background-image:url('" +
            dataUri + "')}</style><div class='box'></div>";

        OfficeProvenanceRemovalResult result = HtmlProvenance.Remove(html);

        Assert.Empty(result.Before.Evidence);
        Assert.False(result.WasChanged);
    }

    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public void DataUriStylesheetsParticipateInImageProvenanceRemoval(bool base64) {
        string image = "data:image/png;base64," +
            Convert.ToBase64String(CreatePngWithManifest(CreateManifestStore()));
        string css = ".box{background-image:url('" + image + "')}";
        string stylesheet = base64
            ? "data:text/css;base64," + Convert.ToBase64String(System.Text.Encoding.UTF8.GetBytes(css))
            : "data:text/css," + Uri.EscapeDataString(css);
        string html = "<link rel='stylesheet' href='" + stylesheet + "'><div class='box'></div>";

        OfficeProvenanceRemovalResult result = HtmlProvenance.Remove(html);

        Assert.True(result.WasChanged);
        Assert.Single(result.Before.Evidence);
        Assert.Empty(result.After.Evidence);
        Assert.Contains("data:text/css;charset=utf-8;base64,", System.Text.Encoding.UTF8.GetString(result.ToArray()), StringComparison.Ordinal);
    }

    [Fact]
    public void ExcelXlsbOwnershipRejectsDuplicateNonWorkbookParts() {
        byte[] original = CreateWave33XlsbProvenancePackage(signed: false);
        byte[] duplicated;
        using (var input = new ZipArchive(new MemoryStream(original), ZipArchiveMode.Read))
        using (var output = new MemoryStream()) {
            using (var archive = new ZipArchive(output, ZipArchiveMode.Create, leaveOpen: true)) {
                foreach (ZipArchiveEntry source in input.Entries) {
                    ZipArchiveEntry target = archive.CreateEntry(source.FullName, CompressionLevel.Optimal);
                    using Stream sourceStream = source.Open();
                    using Stream targetStream = target.Open();
                    sourceStream.CopyTo(targetStream);
                }
                WriteWave33Entry(archive, "DOCPROPS/APP.XML", "<Properties/>");
            }
            duplicated = output.ToArray();
        }

        string path = Path.Combine(Path.GetTempPath(), Guid.NewGuid().ToString("N") + ".xlsb");
        try {
            File.WriteAllBytes(path, duplicated);
            Assert.Throws<InvalidDataException>(() => ExcelDocument.InspectProvenance(path));
            Assert.Throws<InvalidDataException>(() => ExcelDocument.RemoveProvenance(duplicated, "workbook.xlsb"));
        } finally {
            if (File.Exists(path)) File.Delete(path);
        }
    }
}
