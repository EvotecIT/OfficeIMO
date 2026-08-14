using System.IO.Compression;
using OfficeIMO.Excel;
using OfficeIMO.Html;
using OfficeIMO.Markdown;
using OfficeIMO.Provenance;
using Xunit;

namespace OfficeIMO.Shared.Tests;

public sealed partial class ProvenanceDocumentContracts {
    [Fact]
    public void HtmlDirectImageUrlsApplyUrlTabAndNewlinePreprocessing() {
        string base64 = Convert.ToBase64String(CreatePngWithManifest(CreateManifestStore()));
        string html = "<img src=\"da&#10;ta:image/png;base64," + base64 + "\">";

        OfficeProvenanceRemovalResult result = HtmlProvenance.Remove(html);

        Assert.True(result.WasChanged);
        Assert.Empty(result.After.Evidence);
    }

    [Fact]
    public void ForeignNamespaceImageLikeElementsRemainInert() {
        string dataUri = "data:image/png;base64," +
            Convert.ToBase64String(CreatePngWithManifest(CreateManifestStore()));
        string html = "<math><input type=\"image\" src=\"" + dataUri + "\"></math>";

        OfficeProvenanceRemovalResult result = HtmlProvenance.Remove(html);

        Assert.Empty(result.Before.Evidence);
        Assert.False(result.WasChanged);
    }

    [Fact]
    public void LegacyBeforePseudoElementImagesAreActive() {
        string dataUri = "data:image/png;base64," +
            Convert.ToBase64String(CreatePngWithManifest(CreateManifestStore()));
        string html = "<style>.badge:before{content:'';background-image:url('" + dataUri +
            "')}</style><div class=\"badge\"></div>";

        OfficeProvenanceRemovalResult result = HtmlProvenance.Remove(html);

        Assert.True(result.WasChanged);
        Assert.Empty(result.After.Evidence);
    }

    [Fact]
    public void EscapedSelectorCommasDoNotSplitTheSelectorList() {
        string dataUri = "data:image/png;base64," +
            Convert.ToBase64String(CreatePngWithManifest(CreateManifestStore()));
        string html = "<style>.foo\\,bar{background-image:url('" + dataUri +
            "')}</style><div class=\"foo,bar\"></div>";

        OfficeProvenanceRemovalResult result = HtmlProvenance.Remove(html);

        Assert.True(result.WasChanged);
        Assert.Empty(result.After.Evidence);
    }

    [Fact]
    public void MarkdownStringApisRejectUnpairedSurrogates() {
        string markdown = "before\uD800after";

        Assert.Throws<InvalidDataException>(() => MarkdownProvenance.Inspect(markdown));
        Assert.Throws<InvalidDataException>(() => MarkdownProvenance.Remove(markdown));
    }

    [Fact]
    public void ExcelXlsbOwnershipRejectsDuplicateWorkbookParts() {
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
                WriteWave33Entry(archive, "XL/WORKBOOK.BIN", new byte[] { 0x83, 0x01, 0x00 });
            }
            duplicated = output.ToArray();
        }

        Assert.ThrowsAny<Exception>(() => ExcelDocument.RemoveProvenance(duplicated, "workbook.xlsb"));
    }
}
