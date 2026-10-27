using System;
using System.IO;
using System.Linq;
using System.Collections.Generic;
using System.Threading.Tasks;
using OfficeIMO.Drawing;
using OfficeIMO.Word;
using OfficeIMO.Word.Markdown;
using Xunit;

namespace OfficeIMO.Markdown.Tests;

public sealed class WordMarkdownOutputAcceptanceTests {
    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public void OmittedRunningContentRemainsVisibleToLossAcceptance(bool include) {
        using var word = WordDocument.Create();
        word.AddHeadersAndFooters();
        word.Header!.Default!.AddParagraph("Header content");
        word.Footer!.Default!.AddParagraph("Footer content");
        word.AddParagraph("Body content");
        var options = new WordMarkdownExportOptions { IncludeHeadersAndFootersAsSemanticBlocks = include };
        using var markdown = word.ToMarkdown(options, out OfficeConversionReport report);
        Assert.Equal(!include, report.Diagnostics.Any(d => d.LossKind == OfficeConversionLossKind.Omission));
        Assert.Equal(include, markdown.ToMarkdown().Contains("Header content", StringComparison.Ordinal));
        if (!include) {
            options.FailOnLoss = true;
            Assert.Throws<OfficeConversionException>(() => word.ToMarkdown(options));
        }
    }

    [Fact]
    public void ExplicitPageBreakOmissionStillProducesALossReport() {
        using var word = WordDocument.Create(); word.AddParagraph("Before"); word.AddPageBreak(); word.AddParagraph("After");
        using var markdown = word.ToMarkdown(new WordMarkdownExportOptions { PageBreakMode = WordMarkdownPageBreakMode.Omit }, out var report);
        Assert.Contains(report.Diagnostics, d => d.LossKind == OfficeConversionLossKind.Omission && d.Location == "PageBreak");
        Assert.Throws<OfficeConversionException>(() => word.ToMarkdown(new WordMarkdownExportOptions { PageBreakMode = WordMarkdownPageBreakMode.Omit, FailOnLoss = true }));
    }

    [Fact]
    public void SilentUnsupportedContentPolicyStillReportsOmission() {
        using var word = WordDocument.Create();
        word.AddShape(WordShapeType.Rectangle, 60, 30);
        var warnings = new List<string>();
        var options = new WordMarkdownExportOptions { UnsupportedContentMode = WordMarkdownUnsupportedContentMode.Omit, OnWarning = warnings.Add };
        using var markdown = word.ToMarkdown(options, out var report);
        Assert.Empty(warnings);
        Assert.Contains(report.Diagnostics, d => d.LossKind == OfficeConversionLossKind.Omission && d.Location == "shape");
        options.FailOnLoss = true;
        Assert.Throws<OfficeConversionException>(() => word.ToMarkdown(options));
    }

    [Fact]
    public void LossRejectionDoesNotCreateFileImageDestinations() {
        string directory = Path.Combine(Path.GetTempPath(), "officeimo-markdown-image-acceptance-" + Guid.NewGuid().ToString("N"));
        Directory.CreateDirectory(directory);
        try {
            using var word = WordDocument.Create(); word.AddHeadersAndFooters();
            word.Header!.Default!.AddParagraph("Omitted running content");
            byte[] png = Convert.FromBase64String("iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAQAAAC1HAwCAAAAC0lEQVR42mNk+A8AAQUBAScY42YAAAAASUVORK5CYII=");
            using (var source = new MemoryStream(png)) word.AddParagraph().AddImage(source, "embedded.png", width: OfficeIMO.OfficeLength.Pixels(32), height: OfficeIMO.OfficeLength.Pixels(32));
            string output = Path.Combine(directory, "report.md"), images = Path.Combine(directory, "images");
            File.WriteAllText(output, "original markdown");
            Assert.Throws<OfficeConversionException>(() => word.SaveAsMarkdown(output, new WordMarkdownExportOptions {
                FailOnLoss = true, ImageExportMode = WordMarkdownImageExportMode.File, ImageDirectory = images
            }));
            Assert.False(Directory.Exists(images)); Assert.Equal("original markdown", File.ReadAllText(output));
        } finally { Directory.Delete(directory, true); }
    }

    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public async Task LossRejectionPreservesMainOutputAndExistingSidecars(bool asynchronous) {
        string directory = Path.Combine(Path.GetTempPath(), "officeimo-markdown-acceptance-" + Guid.NewGuid().ToString("N"));
        Directory.CreateDirectory(directory);
        try {
            using var word = WordDocument.Create();
            var chart = word.AddChart("Chart",width: OfficeIMO.OfficeLength.Pixels(400),height: OfficeIMO.OfficeLength.Pixels(240));
            chart.AddCategories(new List<string> { "Q1", "Q2" });
            chart.AddBar("Series", new List<int> { 10, 20 }, OfficeColor.CornflowerBlue);
            string output = Path.Combine(directory, "report.md"), assets = Path.Combine(directory, "report.assets");
            Directory.CreateDirectory(assets);
            string sidecar = Path.Combine(assets, "01-chart.svg");
            File.WriteAllText(output, "original markdown"); File.WriteAllText(sidecar, "original asset");
            var options = new WordMarkdownExportOptions { VisualFallbackMode = WordMarkdownVisualFallbackMode.SvgFile, FailOnLoss = true };
            if (asynchronous) await Assert.ThrowsAsync<OfficeConversionException>(() => word.SaveAsMarkdownAsync(output, options));
            else Assert.Throws<OfficeConversionException>(() => word.SaveAsMarkdown(output, options));
            Assert.Equal("original markdown", File.ReadAllText(output)); Assert.Equal("original asset", File.ReadAllText(sidecar));
            Assert.Single(Directory.GetFiles(assets));
        } finally { Directory.Delete(directory, true); }
    }
}
