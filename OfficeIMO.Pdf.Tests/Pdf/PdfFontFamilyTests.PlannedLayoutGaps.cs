using System.Collections.Generic;
using System.Linq;
using PdfPigDocument = UglyToad.PdfPig.PdfDocument;
using OfficeIMO.Pdf;
using OfficeIMO.TestAssets;
using Xunit;

namespace OfficeIMO.Tests.Pdf;

public partial class PdfFontFamilyTests {
    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public void CallerNamedFallbackExplicitPlanLayoutGapsDoNotInheritTemplateCandidate(bool namedMapping) {
        var fallbacks = new PdfEmbeddedFontFallbackSet(new[] {
            new PdfEmbeddedFontFallbackCandidate("Caller Family",
                ManagedTextShapingTestAssets.CreateFontWithUniformAdvance(1000, 300, ' ', 'A'))
        });
        var options = new PdfCreateOptions { DefaultFontSize = global::OfficeIMO.OfficeLength.Points(10) }
            .RegisterNamedFontFamily(CreateCallerCollisionFamily())
            .RegisterFontFamily(PdfStandardFont.TimesRoman, new PdfEmbeddedFontFamily("Explicit Mapped Font",
                ManagedTextShapingTestAssets.CreateFontWithUniformAdvance(1000, 700, ' ', 'A')))
            .RegisterEmbeddedFontFallbacks(fallbacks);
        PdfTextRun template = fallbacks.PlanTextRuns("A", styleTemplate:
                PdfTextRun.Normal("template", font: PdfStandardFont.Courier))[0]
            .WithCharacterSpacing(1).WithHorizontalTextScaling(50);
        PdfTextFallbackPlan plan = PdfTextDiagnostics.PlanEmbeddedFontFallbackText("A A", new[] {
            new PdfEmbeddedFontFallbackCandidate("Planner Without Spaces",
                ManagedTextShapingTestAssets.CreateFontWithUniformAdvance(1000, 500, 'A'))
        });
        Assert.True(plan.IsFullyCovered);
        Assert.Equal(2, plan.Segments.Count);

        IReadOnlyList<PdfTextRun> runs = namedMapping
            ? plan.ToNamedTextRuns(new[] { "Caller Family" }, template)
            : plan.ToTextRuns(new[] { PdfStandardFont.TimesRoman }, template);
        PdfTextRun MappedLetter() => PdfTextRun.Normal("A",
                font: namedMapping ? null : PdfStandardFont.TimesRoman,
                fontFamily: namedMapping ? "Caller Family" : null)
            .WithCharacterSpacing(1).WithHorizontalTextScaling(50);
        PdfTextRun gap = PdfTextRun.Normal(" ", font: PdfStandardFont.Courier)
            .WithCharacterSpacing(1).WithHorizontalTextScaling(50);
        byte[] Render(IReadOnlyList<PdfTextRun> textRuns) => PdfDocument.Create(options.Clone())
            .Paragraph(paragraph => paragraph.Runs(textRuns)).ToBytes();

        // The explicitly mapped letters and the Courier layout gap form the
        // independent public-API reference. Reusing a planned template must
        // not silently force that gap onto its old 300-unit fallback face.
        byte[] expectedPdf = Render(new[] { MappedLetter(), gap, MappedLetter() });
        byte[] actualPdf = Render(runs);
        Assert.Equal("A A", string.Concat(PdfReadDocument.Open(actualPdf).Pages[0].GetTextSpans()
            .Select(span => span.Text)));
        using var reference = PdfPigDocument.Open(expectedPdf);
        using var actual = PdfPigDocument.Open(actualPdf);
        var expectedLetters = reference.GetPage(1).Letters.Where(letter => letter.Value == "A").ToArray();
        var actualLetters = actual.GetPage(1).Letters.Where(letter => letter.Value == "A").ToArray();
        Assert.Equal(2, expectedLetters.Length);
        Assert.Equal(2, actualLetters.Length);
        for (int index = 0; index < expectedLetters.Length; index++) {
            Assert.Equal(expectedLetters[index].StartBaseLine.X, actualLetters[index].StartBaseLine.X, 3);
            Assert.Equal(expectedLetters[index].EndBaseLine.X, actualLetters[index].EndBaseLine.X, 3);
            Assert.Equal(expectedLetters[index].StartBaseLine.Y, actualLetters[index].StartBaseLine.Y, 3);
        }
    }
}
