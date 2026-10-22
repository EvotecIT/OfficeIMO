using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using OfficeIMO.Pdf;
using Letter = UglyToad.PdfPig.Content.Letter;
using PdfPigDocument = UglyToad.PdfPig.PdfDocument;
using Xunit;

namespace OfficeIMO.Tests.Pdf;

public sealed class PdfWordCompressionAdmissionTests {
    [Theory]
    [InlineData("preserve")]
    [InlineData("split-word")]
    [InlineData("collapse")]
    public void NativeSpaceBudgetAdmitsTheWholeWordBeforeClippingOrFormattingSplits(string scenario) {
        PdfTextRun[] runs = scenario == "split-word"
            ? new[] { new PdfTextRun("Alpha beta gam"), new PdfTextRun("ma", color: new PdfColor(1, 0, 0)) }
            : new[] { new PdfTextRun("Alpha beta gamma") };
        var whitespace = scenario == "collapse" ? PdfTextWhitespaceMode.Collapse : PdfTextWhitespaceMode.Preserve;
        List<Letter> control = Assert.Single(Lines(Render(runs, 300D, false, whitespace), "pdf-natural-" + scenario));
        double natural = Span(control);
        double frame = natural - 1D;
        Assert.True(Lines(Render(runs, frame, false, whitespace), "pdf-default-" + scenario).Count > 1);
        var style = new PdfParagraphStyle { UseWordJustificationSpacing = true, TextWhitespaceMode = whitespace };
        var document = PdfDocument.Create(Options(frame)).Paragraph(p => p.Runs(runs), PdfAlign.Justify, style: style);
        style.UseWordJustificationSpacing = false;
        List<Letter> actual = Assert.Single(Lines(document.ToBytes(), "pdf-native-" + scenario));
        Assert.Equal("Alphabetagamma", string.Concat(actual.Select(letter => letter.Value)));
        Assert.Equal(frame, Span(actual), 3);
        for (int index = 0; index < control.Count; index++) Assert.Equal(control[index].Width, actual[index].Width, 3);
    }

    [Fact]
    public void MixedSourceSpaceAdvancesOwnTheirProportionalCompression() {
        var runs = new[] { new PdfTextRun("Alpha"), new PdfTextRun(" ", fontSize: 12D),
            new PdfTextRun("beta"), new PdfTextRun(" ", fontSize: 24D), new PdfTextRun("gamma") };
        List<Letter> natural = Assert.Single(Lines(Render(runs, 300D, false), "mixed-natural"));
        List<Letter> actual = Assert.Single(Lines(Render(runs, Span(natural) - 1.2D, true), "mixed-compressed"));
        double first = Gap(natural, 4, 5) - Gap(actual, 4, 5);
        double second = Gap(natural, 8, 9) - Gap(actual, 8, 9);
        Assert.InRange(first, 0.39D, 0.41D);
        Assert.InRange(second, 0.79D, 0.81D);
        Assert.Equal(1.2D, Span(natural) - Span(actual), 3);
    }

    [Theory]
    [InlineData("nonbreaking")]
    [InlineData("tab")]
    public void NonSpaceSeparatorsDoNotCreateAnOrdinarySpaceBudget(string separator) {
        string text = separator == "tab" ? "Alpha\tbeta\tgamma" : "Alpha\u00a0beta\u00a0gamma";
        var style = new PdfParagraphStyle { UseWordJustificationSpacing = true, TextWhitespaceMode = PdfTextWhitespaceMode.Preserve };
        if (separator == "tab") {
            style.TabStops.Add(new PdfTabStop(40D));
            style.TabStops.Add(new PdfTabStop(80D));
        }
        byte[] wide = PdfDocument.Create(Options(300D)).Paragraph(p => p.Text(text), PdfAlign.Justify, style: style).ToBytes();
        double natural = Span(Assert.Single(Lines(wide, "separator-wide-" + separator)));
        byte[] narrow = PdfDocument.Create(Options(natural - 1D)).Paragraph(p => p.Text(text), PdfAlign.Justify, style: style).ToBytes();
        Assert.True(Lines(narrow, "separator-narrow-" + separator).Count > 1);
    }

    [Fact]
    public void DifferentWidthColumnContinuationRebindsTheNaturalLine() {
        var runs = new[] { new PdfTextRun("Alpha beta gamma") };
        double natural = Span(Assert.Single(Lines(Render(runs, 300D, false), "columns-natural")));
        double narrow = natural - 1D, wide = natural - 0.5D;
        var options = Options(narrow + wide + 20D);
        options.PageHeight = 100D;
        var document = PdfDocument.Create(options);
        document.Content.Columns(content => content.Paragraph(p => p.Text(string.Join("\n", Enumerable.Repeat("Alpha beta gamma", 8))),
                PdfAlign.Justify, style: new PdfParagraphStyle { UseWordJustificationSpacing = true,
                    TextWhitespaceMode = PdfTextWhitespaceMode.Preserve, LineSpacing = PdfLineSpacing.Exactly(18D) }),
            new PdfMultiColumnOptions { BalanceLastPage = false, ColumnDefinitions = new[] {
                new PdfFlowColumn(PdfColumnWidth.Fixed(narrow), 20D), new PdfFlowColumn(PdfColumnWidth.Fixed(wide)) } });
        byte[] bytes = document.ToBytes();
        string path = Save(bytes, "different-width-columns");
        using var pdf = PdfPigDocument.Open(path);
        var lines = pdf.GetPages().SelectMany(page => page.Letters.Where(letter => !string.IsNullOrWhiteSpace(letter.Value))
            .GroupBy(letter => (Column: letter.StartBaseLine.X > 20D + narrow ? 1 : 0, Y: Math.Round(letter.StartBaseLine.Y, 2)))
            .Select(group => (group.Key.Column, Letters: group.OrderBy(letter => letter.StartBaseLine.X).ToList()))).ToList();
        Assert.Equal(8, lines.Count);
        Assert.Contains(lines, line => line.Column == 1);
        foreach (var line in lines) {
            Assert.Equal("Alphabetagamma", string.Concat(line.Letters.Select(letter => letter.Value)));
            Assert.Equal(line.Column == 0 ? narrow : wide, Span(line.Letters), 3);
        }
    }

    private static byte[] Render(PdfTextRun[] runs, double width, bool native, PdfTextWhitespaceMode whitespace = PdfTextWhitespaceMode.Preserve) =>
        PdfDocument.Create(Options(width)).Paragraph(p => p.Runs(runs), PdfAlign.Justify,
            style: new PdfParagraphStyle { UseWordJustificationSpacing = native, TextWhitespaceMode = whitespace }).ToBytes();
    private static PdfOptions Options(double width) => new() { PageWidth = width + 40D, PageHeight = 300D,
        MarginLeft = 20D, MarginRight = 20D, MarginTop = 20D, MarginBottom = 20D,
        DefaultFont = PdfStandardFont.Helvetica, DefaultFontSize = 12D, CompressContentStreams = false };
    private static double Span(IReadOnlyList<Letter> letters) => letters[letters.Count - 1].EndBaseLine.X - letters[0].StartBaseLine.X;
    private static double Gap(IReadOnlyList<Letter> letters, int left, int right) => letters[right].StartBaseLine.X - letters[left].EndBaseLine.X;
    private static List<List<Letter>> Lines(byte[] bytes, string label) {
        using var pdf = PdfPigDocument.Open(Save(bytes, label));
        return pdf.GetPages().SelectMany(page => page.Letters.Where(letter => !string.IsNullOrWhiteSpace(letter.Value))
            .GroupBy(letter => Math.Round(letter.StartBaseLine.Y, 2)).Select(line => line.OrderBy(letter => letter.StartBaseLine.X).ToList())).ToList();
    }
    private static string Save(byte[] bytes, string label) {
        string directory = Path.Combine(Path.GetTempPath(), "officeimo-compression-evidence");
        Directory.CreateDirectory(directory);
        string path = Path.Combine(directory, label + ".pdf");
        File.WriteAllBytes(path, bytes);
        return path;
    }
}
