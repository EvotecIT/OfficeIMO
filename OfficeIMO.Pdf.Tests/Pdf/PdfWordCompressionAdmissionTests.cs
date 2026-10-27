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
    private const string CandidateQualityPrefix = "one two three four five six seven eight nine ten";

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

    [Theory]
    [InlineData("preserve")]
    [InlineData("collapse")]
    [InlineData("split-word")]
    public void WordsCandidateQualityChoosesOppositeEndpointsWithinTheSameCompressionBudget(string scenario) {
        PdfTextWhitespaceMode whitespace = scenario == "collapse" ? PdfTextWhitespaceMode.Collapse : PdfTextWhitespaceMode.Preserve;
        string text = CandidateQualityPrefix + " it tail";
        PdfTextRun[] runs = scenario == "split-word"
            ? new[] { new PdfTextRun(CandidateQualityPrefix + " i"), new PdfTextRun("t", color: new PdfColor(1, 0, 0)), new PdfTextRun(" tail") }
            : new[] { new PdfTextRun(scenario == "collapse" ? text.Replace(" ", "  ") : text) };
        var frames = CandidateQualityFrames(runs, whitespace, scenario);
        string prefix = CandidateQualityPrefix.Replace(" ", "");
        foreach (var frame in new[] { (Width: frames.Narrow, First: prefix, Label: "narrow"),
            (Width: frames.Wide, First: prefix + "it", Label: "wide") }) {
            List<List<Letter>> lines = Lines(Render(runs, frame.Width, true, whitespace), "quality-" + scenario + "-" + frame.Label);
            Assert.Equal(2, lines.Count);
            Assert.Equal(frame.First, string.Concat(lines[0].Select(letter => letter.Value)));
            Letter[] actual = lines.SelectMany(line => line).ToArray();
            Assert.Equal(prefix + "ittail", string.Concat(actual.Select(letter => letter.Value)));
            Assert.Equal(frames.Control.Count, actual.Length);
            Assert.Equal(frame.Width, Span(lines[0]), 3);
            for (int index = 0; index < actual.Length; index++) Assert.Equal(frames.Control[index].Width, actual[index].Width, 3);
            Assert.All(actual, letter => {
                Assert.True(letter.StartBaseLine.X >= 19.999D);
                Assert.True(letter.EndBaseLine.X <= 20D + frame.Width + 0.001D);
            });
            if (scenario == "preserve") {
                List<List<Letter>> generic = Lines(Render(runs, frame.Width, false), "quality-generic-" + frame.Label);
                List<List<Letter>> characters = Lines(Render(runs, frame.Width, true, mode: PdfJustificationMode.Characters), "quality-characters-" + frame.Label);
                Assert.Equal(prefix, string.Concat(generic[0].Select(letter => letter.Value)));
                Assert.Equal(prefix + "it", string.Concat(characters[0].Select(letter => letter.Value)));
                Assert.Equal(prefix + "ittail", string.Concat(generic.SelectMany(line => line).Select(letter => letter.Value)));
                Assert.Equal(prefix + "ittail", string.Concat(characters.SelectMany(line => line).Select(letter => letter.Value)));
            }
        }
    }

    [Fact]
    public void RejectedCompressionCandidateRetainsItsNoteAndFollowingHardBreakAcrossSerialization() {
        var runs = new[] { new PdfTextRun(CandidateQualityPrefix + " it tail") };
        double frame = CandidateQualityFrames(runs, PdfTextWhitespaceMode.Preserve, "note-break").Narrow;
        PdfDocument document = PdfDocument.Create(Options(frame));
        int noteCalls = 0, paragraphCalls = 0;
        PdfFootnoteDefinition note = document.Content.DefineFootnote(body => {
            noteCalls++;
            body.Text("QualityNote", style: new PdfParagraphStyle { FontSize = 8D, LineSpacing = PdfLineSpacing.Exactly(10D) });
        });
        document.Content.Paragraph(paragraph => {
            paragraphCalls++;
            paragraph.Text(CandidateQualityPrefix + " ").Runs(new[] { PdfTextRun.Normal("it").WithFootnoteReference(note) })
                .LineBreak().Text("tail");
        }, PdfAlign.Justify, style: new PdfParagraphStyle {
            UseWordJustificationSpacing = true, JustificationMode = PdfJustificationMode.Words,
            TextWhitespaceMode = PdfTextWhitespaceMode.Preserve, JustifyExplicitLineBreaks = true
        });
        Assert.Equal(1, noteCalls);
        Assert.Equal(1, paragraphCalls);
        for (int attempt = 0; attempt < 2; attempt++) {
            List<List<Letter>> lines = Lines(document.ToBytes(), "quality-note-break-" + attempt);
            string[] text = lines.Select(line => string.Concat(line.Select(letter => letter.Value))).ToArray();
            Assert.Equal(4, text.Length);
            Assert.Equal(new[] { CandidateQualityPrefix.Replace(" ", ""), "it", "tail" }, text.Where(line => line != "QualityNote").ToArray());
            Assert.Single(text, line => line == "QualityNote");
            Assert.Equal(1, noteCalls);
            Assert.Equal(1, paragraphCalls);
        }
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
        RetainEvidence(bytes, "different-width-columns");
        using var pdf = PdfPigDocument.Open(bytes);
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

    private static (double Narrow, double Wide, List<Letter> Control) CandidateQualityFrames(
        PdfTextRun[] runs, PdfTextWhitespaceMode whitespace, string label) {
        List<Letter> control = Assert.Single(Lines(Render(runs, 600D, false, whitespace), "quality-natural-" + label));
        int prefixGlyphs = CandidateQualityPrefix.Replace(" ", "").Length;
        double previousNatural = Span(control.Take(prefixGlyphs).ToList());
        double candidateNatural = Span(control.Take(prefixGlyphs + "it".Length).ToList());
        double ordinarySpace = Gap(control, "one".Length - 1, "one".Length);
        // Fixture widths lie inside the existing quarter-space feasibility allowance.
        // No quality score is computed here: expected legal endpoints are explicit above.
        double capacity = (CandidateQualityPrefix.Count(character => character == ' ') + 1) * ordinarySpace / 4D;
        double narrow = candidateNatural - capacity * 0.75D;
        double wide = candidateNatural - capacity * 0.25D;
        Assert.True(ordinarySpace > 0D && previousNatural < narrow && narrow < wide && wide < candidateNatural);
        Assert.True(candidateNatural - narrow < capacity && candidateNatural - wide < capacity);
        return (narrow, wide, control);
    }

    private static byte[] Render(PdfTextRun[] runs, double width, bool native, PdfTextWhitespaceMode whitespace = PdfTextWhitespaceMode.Preserve,
        PdfJustificationMode mode = PdfJustificationMode.Words) =>
        PdfDocument.Create(Options(width)).Paragraph(p => p.Runs(runs), PdfAlign.Justify,
            style: new PdfParagraphStyle { UseWordJustificationSpacing = native, TextWhitespaceMode = whitespace, JustificationMode = mode }).ToBytes();
    private static PdfCreateOptions Options(double width) => new() { PageWidth = width + 40D, PageHeight = 300D,
        MarginLeft = 20D, MarginRight = 20D, MarginTop = 20D, MarginBottom = 20D,
        DefaultFont = PdfStandardFont.Helvetica, DefaultFontSize = 12D, CompressContentStreams = false };
    private static double Span(IReadOnlyList<Letter> letters) => letters[letters.Count - 1].EndBaseLine.X - letters[0].StartBaseLine.X;
    private static double Gap(IReadOnlyList<Letter> letters, int left, int right) => letters[right].StartBaseLine.X - letters[left].EndBaseLine.X;
    private static List<List<Letter>> Lines(byte[] bytes, string label) {
        RetainEvidence(bytes, label);
        using var pdf = PdfPigDocument.Open(bytes);
        return pdf.GetPages().SelectMany(page => page.Letters.Where(letter => !string.IsNullOrWhiteSpace(letter.Value))
            .GroupBy(letter => Math.Round(letter.StartBaseLine.Y, 2)).Select(line => line.OrderBy(letter => letter.StartBaseLine.X).ToList())).ToList();
    }
    private static void RetainEvidence(byte[] bytes, string label) {
        string? output = Environment.GetEnvironmentVariable("OFFICEIMO_PDF_ENGINE_PROOF_OUTPUT");
        if (string.IsNullOrWhiteSpace(output)) return;
        string directory = output!;
        Directory.CreateDirectory(directory);
        File.WriteAllBytes(Path.Combine(directory, "compression-" + label + "-" + Guid.NewGuid().ToString("N") + ".pdf"), bytes);
    }
}
