using OfficeIMO.Markdown;
using OfficeIMO.PowerPoint;

namespace OfficeIMO.Workflows.Tests;

public sealed class MarkdownDeckInlineLossTests {
    private const string Target = "https://example.test/runbook";

    [Theory]
    [InlineData("# Title\n\nPlain **Important** [Runbook](https://example.test/runbook)")]
    [InlineData("# Title\n\n- Plain **Important** [Runbook](https://example.test/runbook)")]
    [InlineData("# Title\n\n1. Outer\n   - Plain **Important** [Runbook](https://example.test/runbook)")]
    [InlineData("# Title\n\n> Plain **Important** [Runbook](https://example.test/runbook)")]
    [InlineData("# Plain **Important** [Runbook](https://example.test/runbook)")]
    [InlineData("# Title\n\n## Plain **Important** [Runbook](https://example.test/runbook)")]
    public async Task FlattenedInlineSemanticsAreReportedAndStrictSaveWritesNothing(string markdown) {
        using MarkdownDocument source = MarkdownDocument.Parse(markdown);
        using var output = new MemoryStream();
        var options = new MarkdownToPowerPointOptions { BlockquotesAsNotes = false };
        OfficeConversionReport report = await source.SaveAsPowerPointAsync(output, options);
        AssertInlineLoss(report);
        output.Position = 0;
        using (PowerPointDocument deck = PowerPointDocument.Load(output)) {
            string visible = string.Join("\n", deck.Slides.SelectMany(slide => slide.TextBoxes).Select(box => box.Text));
            Assert.Contains("Important", visible, StringComparison.Ordinal);
            Assert.Contains("Runbook", visible, StringComparison.Ordinal);
        }

        options.FailOnLoss = true;
        using var rejected = new MemoryStream();
        OfficeConversionException error = await Assert.ThrowsAsync<OfficeConversionException>(() => source.SaveAsPowerPointAsync(rejected, options));
        AssertInlineLoss(error.Report);
        Assert.Equal(0, rejected.Length);
        Assert.True(rejected.CanWrite);
    }

    [Fact]
    public async Task UniformFormattingUsesNativeBlockStyleWithoutRejectingStrictOutput() {
        using MarkdownDocument source = MarkdownDocument.Parse("# Title\n\n**_Uniform marker_**");
        using var output = new MemoryStream();
        OfficeConversionReport report = await source.SaveAsPowerPointAsync(output, new MarkdownToPowerPointOptions { FailOnLoss = true });
        Assert.False(report.HasLoss);
        output.Position = 0;
        using PowerPointDocument deck = PowerPointDocument.Load(output);
        var runs = deck.Slides.SelectMany(slide => slide.TextBoxes).SelectMany(box => box.Paragraphs)
            .SelectMany(paragraph => paragraph.Runs).Where(run => run.Text.Contains("Uniform marker", StringComparison.Ordinal)).ToArray();
        Assert.NotEmpty(runs);
        Assert.All(runs, run => { Assert.True(run.Bold); Assert.True(run.Italic); });
    }

    [Fact]
    public async Task NotesRetainMarkdownTargetsButReportNoninteractiveFormatting() {
        using MarkdownDocument source = MarkdownDocument.Parse("# Title\n\n> Plain **Important** [Runbook](https://example.test/runbook)");
        var converted = await source.ToPowerPointAsync();
        using PowerPointDocument deck = converted.Value;
        Assert.True(converted.Report.HasLoss);
        Assert.Contains(converted.Report.Diagnostics, finding => finding.Code == OfficeDiagnosticCodes.SimplifiedBlock
            && finding.LossKind == OfficeConversionLossKind.Approximation && finding.Message.Contains(Target, StringComparison.Ordinal));
        Assert.DoesNotContain(converted.Report.Diagnostics, finding => finding.Code == OfficeDiagnosticCodes.ShapeHyperlinkOmitted);
        Assert.Contains(deck.Slides, slide => slide.Notes.Text.Contains(Target, StringComparison.Ordinal));
        using var output = new MemoryStream();
        await Assert.ThrowsAsync<OfficeConversionException>(() => source.SaveAsPowerPointAsync(output, new MarkdownToPowerPointOptions { FailOnLoss = true }));
        Assert.Equal(0, output.Length);
    }

    private static void AssertInlineLoss(OfficeConversionReport report) {
        Assert.True(report.HasLoss);
        Assert.Contains(report.Diagnostics, finding => finding.Code == OfficeDiagnosticCodes.SimplifiedBlock
            && finding.LossKind == OfficeConversionLossKind.Approximation && finding.Message.Contains("bold", StringComparison.OrdinalIgnoreCase));
        Assert.Contains(report.Diagnostics, finding => finding.Code == OfficeDiagnosticCodes.ShapeHyperlinkOmitted
            && finding.LossKind == OfficeConversionLossKind.Omission && finding.Location == Target);
    }
}
