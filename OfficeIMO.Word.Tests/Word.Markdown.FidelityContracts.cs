using OfficeIMO.Word.Markdown;
using Xunit;

namespace OfficeIMO.Tests;

public sealed class WordMarkdownFidelityContracts {
    [Fact]
    public void CommonProjectionPreservesLegacyDiagnosticsWithEmptyText() {
        var report = new WordMarkdownConversionReport(new[] {
            new WordMarkdownConversionDiagnostic(
                "WORD_EMPTY_MESSAGE",
                string.Empty,
                OfficeConversionLossKind.Omission)
        });

        OfficeDiagnostic diagnostic = Assert.Single(report.Diagnostics);
        Assert.Equal(string.Empty, diagnostic.Message);
        Assert.Equal(OfficeConversionLossKind.Omission, diagnostic.LossKind);
        Assert.True(report.HasLoss);
        Assert.Throws<WordMarkdownConversionException>(() => report.ThrowIfLoss());
    }

    [Fact]
    public void WordExportLeavesTheMarkdownDocumentBlocksUnchanged() {
        var markdown = OfficeIMO.Markdown.MarkdownDocument.Create();
        markdown.Add(new OfficeIMO.Markdown.HeadingBlock(2, "Contents"));
        markdown.Add(new OfficeIMO.Markdown.TocBlock {
            IncludeTitle = true,
            Title = "Contents",
            TitleLevel = 2,
            TitleHeadingAlreadyRendered = true
        });

        using (markdown.ToWord()) {
        }

        Assert.Equal(2, markdown.Blocks.Count);
        Assert.IsType<OfficeIMO.Markdown.HeadingBlock>(markdown.Blocks[0]);
    }
}
