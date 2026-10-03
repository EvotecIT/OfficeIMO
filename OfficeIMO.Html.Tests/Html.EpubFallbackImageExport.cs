using System.Text;
using System.Threading.Tasks;
using System.Xml.Linq;
using OfficeIMO.Drawing;
using OfficeIMO.Epub;
using OfficeIMO.Epub.Image;
using OfficeIMO.Shared.Tests;
using Xunit;

namespace OfficeIMO.Html.Tests;

public sealed class EpubFallbackImageExportContracts {
    [Theory]
    [InlineData(false, false)]
    [InlineData(false, true)]
    [InlineData(true, false)]
    [InlineData(true, true)]
    public async Task FailedFallback_IsAnOmissionAndStrictImageExportRejectsIt(bool cycle, bool asynchronous) {
        byte[] package = EpubIntegrityFixtures.Package(new[] {
            ("good", "good.xhtml", "application/xhtml+xml", ""), ("foreign", "foreign.bin", "application/vnd.example.foreign", "")
        }, "<itemref idref='good'/><itemref idref='foreign'/>", new[] {
            ("good.xhtml", EpubIntegrityFixtures.Xhtml("<p>Visible chapter</p>")), ("foreign.bin", "opaque")
        });
        package = EditPackage(package, root => root.Descendants(Opf + "item").Single(item => (string?)item.Attribute("id") == "foreign")
            .SetAttributeValue("fallback", cycle ? "foreign" : "missing"));
        EpubDocument read = EpubDocument.Load(new MemoryStream(package), new EpubReadOptions { IncludeRawHtml = true });
        Assert.False(read.ReadSummary.IsComplete);
        var options = new EpubImageExportOptions { Policy = new OfficeImageExportPolicy { RequireNoOmissions = true } };
        int accepted = 0;
        OfficeImageExportPolicyException error = asynchronous ?
            await Assert.ThrowsAsync<OfficeImageExportPolicyException>(() => read.ExportImagesAsync(OfficeImageExportFormat.Png,
                (_, _) => { accepted++; return Task.CompletedTask; }, options)) :
            Assert.Throws<OfficeImageExportPolicyException>(() => read.ExportImages(OfficeImageExportFormat.Png, _ => accepted++, options));
        Assert.Equal(0, accepted);
        Assert.Contains(error.Diagnostics, diagnostic => diagnostic.LossKind == OfficeConversionLossKind.Omission &&
            diagnostic.Code == (cycle ? "EPUB_IMAGE_EPUB_SPINE_FALLBACK_CYCLE" : "EPUB_IMAGE_EPUB_SPINE_FALLBACK_MISSING"));
    }

    private static readonly XNamespace Opf = "http://www.idpf.org/2007/opf";
    private static byte[] EditPackage(byte[] package, Action<XElement> edit) {
        using var archive = new System.IO.Compression.ZipArchive(new MemoryStream(package));
        using Stream stream = archive.GetEntry("EPUB/package.opf")!.Open();
        XDocument xml = XDocument.Load(stream); edit(xml.Root!);
        return EpubIntegrityFixtures.ReplaceEntry(package, "EPUB/package.opf", Encoding.UTF8.GetBytes(xml.ToString()));
    }
}
