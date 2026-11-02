using OfficeIMO.Rtf;
using OfficeIMO.Rtf.Pdf;
using PdfCore = OfficeIMO.Pdf;
using PdfPigDocument = UglyToad.PdfPig.PdfDocument;
using Xunit;

namespace OfficeIMO.Tests.Rtf;

public sealed class RtfPdfWhitespaceRegressionTests {
    [Theory]
    [InlineData(false, 1)]
    [InlineData(false, 3)]
    [InlineData(true, 1)]
    [InlineData(true, 3)]
    public void DefaultImportRetainsSourceSpaceAdvancesAcrossFontChanges(bool unrelatedPdfOptions, int spaces) {
        RtfDocument document = CreateDocument(spaces);
        PdfCore.PdfCreateOptions? caller = unrelatedPdfOptions
            ? new PdfCore.PdfCreateOptions { CompressContentStreams = false }
            : null;
        var options = new RtfToPdfOptions {
            ResourcePolicy = PdfCore.PdfResourcePolicy.CreatePortableDeterministic(),
            PdfOptions = caller
        };

        PdfCore.PdfDocumentConversionResult result = document.ConvertToPdfWithEvidence(options);
        using PdfPigDocument pdf = PdfPigDocument.Open(result.Value.ToBytes());
        Assert.Equal(1, pdf.NumberOfPages);
        var letters = pdf.GetPage(1).Letters;
        var first = Assert.Single(letters, letter => letter.Value == "A");
        var next = Assert.Single(letters, letter => letter.Value == "B");
        Assert.Equal(12D, first.FontSize, 3);
        Assert.Equal(24D, next.FontSize, 3);
        // Standard Courier advances 0.6 em: both A and its source spaces are 12 pt.
        Assert.Equal((spaces + 1) * 7.2D, next.StartBaseLine.X - first.StartBaseLine.X, 3);
        Assert.Equal(PdfCore.PdfTextWhitespaceMode.Preserve, result.Value.Options.TextWhitespaceMode);
        if (caller != null) {
            Assert.Equal(PdfCore.PdfTextWhitespaceMode.Collapse, caller.TextWhitespaceMode);
            Assert.False(caller.HasExplicitTextWhitespaceConfiguration);
        }
    }

    [Theory]
    [InlineData(PdfCore.PdfTextWhitespaceMode.Collapse, 14.4D)]
    [InlineData(PdfCore.PdfTextWhitespaceMode.Preserve, 28.8D)]
    [InlineData(PdfCore.PdfTextWhitespaceMode.Preformatted, 28.8D)]
    public void ExplicitWhitespaceModeRetainsItsLayoutAndLeavesCallerOptionsUnchanged(
        PdfCore.PdfTextWhitespaceMode mode, double expectedAdvance) {
        // Repeated spaces distinguish Collapse from both preserving modes.
        RtfDocument document = CreateDocument(spaces: 3);
        var caller = new PdfCore.PdfCreateOptions { TextWhitespaceMode = mode };

        PdfCore.PdfDocumentConversionResult result = document.ConvertToPdfWithEvidence(new RtfToPdfOptions {
            ResourcePolicy = PdfCore.PdfResourcePolicy.CreatePortableDeterministic(),
            PdfOptions = caller
        });
        Assert.Equal(mode, result.Value.Options.TextWhitespaceMode);
        Assert.Equal(mode, caller.TextWhitespaceMode);
        Assert.True(caller.HasExplicitTextWhitespaceConfiguration);
        using PdfPigDocument pdf = PdfPigDocument.Open(result.Value.ToBytes());
        Assert.Equal(1, pdf.NumberOfPages);
        var letters = pdf.GetPage(1).Letters;
        var first = Assert.Single(letters, letter => letter.Value == "A");
        var next = Assert.Single(letters, letter => letter.Value == "B");
        Assert.Equal(expectedAdvance, next.StartBaseLine.X - first.StartBaseLine.X, 3);
    }

    [Fact]
    public void DefaultImportKeepsAHyphenatedTokenWholeWhileExplicitCollapseUsesTheEarlierBreak() {
        RtfDocument document = RtfDocument.Create();
        document.Settings.SetDefaultFont(document.AddFont("Courier New"));
        RtfParagraph paragraph = document.AddParagraph();
        paragraph.AddText("PREFIX ").SetFontSize(12);
        paragraph.AddText("END-STYLE-OVERRIDE").SetFontSize(12);
        // 140 pt fits the whole 129.6 pt token on a fresh line. Collapse may
        // instead use the earlier hyphen boundary after the 50.4 pt prefix.
        document.PageSetup.SetPaperSize(3600, 4800);
        document.PageSetup.SetMargins(400, 400, 400, 400);

        var collapsed = ReadLines(document, new PdfCore.PdfCreateOptions { TextWhitespaceMode = PdfCore.PdfTextWhitespaceMode.Collapse });
        Assert.Equal(PdfCore.PdfTextWhitespaceMode.Collapse, collapsed.Mode);
        Assert.Equal(new[] { "PREFIXEND-STYLE-", "OVERRIDE" }, collapsed.Lines);
        var imported = ReadLines(document, options: null);
        // Check actual line contents first so the RED policy reversal proves
        // the observed wrap defect, rather than only an option-property change.
        Assert.Equal(new[] { "PREFIX", "END-STYLE-OVERRIDE" }, imported.Lines);
        Assert.Equal(PdfCore.PdfTextWhitespaceMode.Preserve, imported.Mode);

        static (PdfCore.PdfTextWhitespaceMode Mode, string[] Lines) ReadLines(RtfDocument source, PdfCore.PdfCreateOptions? options) {
            PdfCore.PdfDocumentConversionResult result = source.ConvertToPdfWithEvidence(new RtfToPdfOptions {
                ResourcePolicy = PdfCore.PdfResourcePolicy.CreatePortableDeterministic(),
                PdfOptions = options
            });
            using PdfPigDocument pdf = PdfPigDocument.Open(result.Value.ToBytes());
            Assert.Equal(1, pdf.NumberOfPages);
            var letters = pdf.GetPage(1).Letters.Where(letter => !string.IsNullOrWhiteSpace(letter.Value)).ToArray();
            Assert.Equal("PREFIXEND-STYLE-OVERRIDE", string.Concat(letters.Select(letter => letter.Value)));
            Assert.All(letters, letter => Assert.Equal(12D, letter.FontSize, 3));
            string[] lines = letters.GroupBy(letter => Math.Round(letter.StartBaseLine.Y, 3))
                .OrderByDescending(line => line.Key).Select(line => string.Concat(line.Select(letter => letter.Value))).ToArray();
            return (result.Value.Options.TextWhitespaceMode, lines);
        }
    }

    private static RtfDocument CreateDocument(int spaces) {
        RtfDocument document = RtfDocument.Create();
        document.Settings.SetDefaultFont(document.AddFont("Courier New"));
        RtfParagraph paragraph = document.AddParagraph();
        paragraph.AddText("A" + new string(' ', spaces)).SetFontSize(12);
        paragraph.AddText("B").SetFontSize(24);
        return document;
    }
}
