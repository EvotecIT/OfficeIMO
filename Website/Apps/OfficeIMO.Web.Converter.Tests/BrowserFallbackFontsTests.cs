using System.IO.Compression;
using System.Text;
using OfficeIMO.Web.Converter.Services;
using Xunit;

namespace OfficeIMO.Web.Converter.Tests;

public sealed class BrowserFallbackFontsTests {
    [Fact]
    public void SplitFontAssembliesStillMakeThePinnedPack() {
        Assert.Equal(BrowserPortablePdfProfile.ExpectedFontPackFingerprint, BrowserPortablePdfProfile.ComputeFullFingerprint());
        Assert.Equal(BrowserPortablePdfProfile.ExpectedFontPackFingerprint, BrowserPortablePdfProfile.FontPackFingerprint);
        Assert.True(BrowserFallbackFonts.IsAvailable);
    }

    [Theory]
    [InlineData("Quarterly report: revenue up 12% — “strong” year, café, naïve, €5", false)]
    [InlineData("Привет, Ελλάδα", false)]
    [InlineData("&amp; &lt;b&gt; &nbsp; &copy; &#169;", false)]
    [InlineData("会議の議事録", true)]
    [InlineData("مرحبا بالعالم", true)]
    [InlineData("Done ✓", true)]
    [InlineData("&#x3042;", true)]
    [InlineData("&#12354;", true)]
    [InlineData("&rarr; next", true)]
    [InlineData("<span style=\"font-family: Symbol\">a</span>", true)]
    [InlineData("<span style=\"font-family: symbol\">a</span>", true)]
    [InlineData("<span style=\"font-family: zapfdingbats\">a</span>", true)]
    [InlineData("<w:rFonts w:ascii=\"Symbol\" w:hAnsi=\"Symbol\"/>", true)]
    [InlineData("a symbolic gesture", false)]
    public void TextNeedsFallbackOnlyWhenCarlitoCannotDrawIt(string text, bool expected) {
        Assert.Equal(expected, BrowserFallbackFonts.Needed(text));
    }

    [Fact]
    public void OfficePackagesIgnoreThemeFontListsButSeeBodyText() {
        const string theme = "<a:theme><a:font script=\"Jpan\" typeface=\"ＭＳ ゴシック\"/><a:font script=\"Arab\" typeface=\"Times New Roman\"/></a:theme>";
        byte[] latin = Package(("word/document.xml", "<w:document><w:t>Hello world</w:t></w:document>"), ("word/theme/theme1.xml", theme),
            ("word/fontTable.xml", "<w:font w:name=\"ＭＳ 明朝\"/>"), ("docProps/core.xml", "<dc:title>議事録</dc:title>"));
        byte[] japanese = Package(("word/document.xml", "<w:document><w:t>議事録</w:t></w:document>"), ("word/theme/theme1.xml", theme));
        byte[] symbolBullets = Package(("word/document.xml", "<w:document><w:t>Item</w:t></w:document>"),
            ("word/numbering.xml", "<w:lvl><w:lvlText w:val=\"\"/><w:rFonts w:ascii=\"Symbol\"/></w:lvl>"));

        Assert.False(BrowserFallbackFonts.Needed(latin, ".docx"));
        Assert.True(BrowserFallbackFonts.Needed(japanese, ".docx"));
        Assert.True(BrowserFallbackFonts.Needed(symbolBullets, ".docx"));
        Assert.True(BrowserFallbackFonts.Needed([1, 2, 3], ".docx"));
        Assert.True(BrowserFallbackFonts.Needed([1, 2, 3], ".pdf"));
    }

    [Fact]
    public void SampleWordDocumentDoesNotDownloadFallbackFonts() {
        byte[] sample = File.ReadAllBytes(Path.Combine(AppContext.BaseDirectory, "samples", "business-summary.docx"));
        Assert.False(BrowserFallbackFonts.Needed(sample, ".docx"));
    }

    [Theory]
    [InlineData(".docx", "word/document.xml")]
    [InlineData(".xlsx", "xl/worksheets/sheet1.xml")]
    [InlineData(".pptx", "ppt/slides/slide1.xml")]
    public void FalseZipLengthsCannotTriggerAnUnboundedFontScan(string extension, string part) {
        byte[] package = Package((part, "<text>" + new string('a', 1024 * 1024) + "</text>"));
        // Keep compressed ranges intact, but understate the local and central expanded sizes.
        for (int index = 0; index <= package.Length - 46; index++) {
            if (package[index] != 0x50 || package[index + 1] != 0x4B) continue;
            int sizeOffset = package[index + 2] == 3 && package[index + 3] == 4 ? index + 22
                : package[index + 2] == 1 && package[index + 3] == 2 ? index + 24 : -1;
            if (sizeOffset < 0) continue;
            package[sizeOffset] = 1;
            package[sizeOffset + 1] = package[sizeOffset + 2] = package[sizeOffset + 3] = 0;
        }
        Assert.True(BrowserFallbackFonts.Needed(package, extension));
    }

    private static byte[] Package(params (string Name, string Xml)[] parts) {
        using var buffer = new MemoryStream();
        using (var archive = new ZipArchive(buffer, ZipArchiveMode.Create, leaveOpen: true)) {
            foreach ((string name, string xml) in parts) {
                using Stream stream = archive.CreateEntry(name).Open();
                byte[] bytes = Encoding.UTF8.GetBytes(xml);
                stream.Write(bytes, 0, bytes.Length);
            }
        }
        return buffer.ToArray();
    }
}
