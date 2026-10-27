using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using OfficeIMO.Pdf;
using Xunit;

namespace OfficeIMO.Tests.Pdf;

public sealed class PdfActualTextBaselineTests {
    [Fact]
    public void DefaultMarkedArabicExtractsOneLogicalLineWithoutMovingTheFirstPaintedMarks() {
        const string source = "لَا لَأ السَّلَامُ";
        byte[] font = File.ReadAllBytes(Path.Combine(RepositoryTestPaths.Find(),
            "Website", "Apps", "OfficeIMO.Web.Converter", "Assets", "Fonts", "NotoSansArabic-Regular.ttf"));
        var report = new PdfConversionReport();
        var options = new PdfCreateOptions {
            PageWidth = 360D, PageHeight = 160D, MarginLeft = 24D, MarginRight = 24D,
            MarginTop = 24D, MarginBottom = 24D, DefaultFontSize = 20D, Language = "ar"
        };
        options.EmbedStandardFont(PdfStandardFont.Helvetica, font, "Arabic baseline").ReportDiagnosticsTo(report);
        byte[] bytes = PdfDocument.Create(options).Paragraph(paragraph => paragraph.Text(source)).ToBytes();
        PdfReadDocument document = PdfReadDocument.Open(bytes);
        PdfReadPage page = Assert.Single(document.Pages);
        PdfTextSpan[] words = page.GetTextSpans().Where(span => span.HasActualText).ToArray();

        Assert.Equal(3, words.Length);
        Assert.All(words, word => Assert.Equal(0D, word.Advance, 6));
        // The exposed origin still belongs to the initial raised/lowered mark.
        PdfTextSpan space = page.GetTextSpans().First(span => span.Text == " ");
        Assert.True(words[0].Y - space.Y > 2.5D);
        Assert.True(words[2].Y < space.Y);
        Assert.Equal(source, document.ExtractText());
        Assert.Equal(source, page.ExtractTextWithColumns(new PdfTextLayoutOptions { ForceSingleColumn = true }));
        TextLayoutEngine.TextLine line = Assert.Single(TextLayoutEngine.BuildLines(page.GetTextSpans()));
        Assert.Equal(space.Y, line.Y, 6);
        Assert.False(words[0].TryGetGlyphs(out _));
        report.RequireNoLoss();
    }

    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public void ClosedMarkFirstReplacementUsesCompatibleBasePaintEvenUnderRotation(bool transformed) {
        string content = Replacement("6 Ts (m) Tj 0 Ts (a) Tj") + Tail;
        if (transformed) content = "q 0 2 -3 0 400 0 cm " + content + " Q";
        List<PdfTextSpan> spans = Parse(content);
        PdfTextSpan replacement = spans[0];

        TextLayoutEngine.TextLine line = Assert.Single(TextLayoutEngine.BuildLines(spans));

        Assert.Equal("a\u0301 tail", line.Text);
        Assert.Equal(0D, replacement.Advance);
        Assert.Equal(transformed ? 82D : 20D, replacement.X, 6);
        Assert.Equal(transformed ? 40D : 106D, replacement.Y, 6);
        Assert.Equal(transformed ? 100D : 20D, replacement.LogicalLayoutX, 6);
        Assert.Equal(transformed ? 40D : 100D, replacement.LogicalLayoutY, 6);
        // Editing's ordinary coordinate projection continues to see painted marks.
        Assert.Equal(transformed ? -82D : 106D,
            new PdfTextBaselineCoordinates(replacement.RotationDegrees).Normal(replacement), 6);
    }

    [Theory]
    [InlineData(6)]
    [InlineData(-3)]
    public void UniformRaisedReplacementKeepsItsAuthoredBaseline(int rise) {
        List<PdfTextSpan> spans = Parse(Replacement(rise + " Ts (m) Tj (a) Tj") + Tail);
        PdfTextSpan replacement = spans[0];

        Assert.Equal(100D + rise, replacement.Y, 6);
        Assert.Equal(replacement.Y, replacement.LogicalLayoutY, 6);
        Assert.Equal(2, TextLayoutEngine.BuildLines(spans).Count);
    }

    [Theory]
    [InlineData("multiline")]
    [InlineData("mixed-font")]
    [InlineData("mixed-transform")]
    [InlineData("mixed-clip")]
    [InlineData("backward")]
    [InlineData("unmapped")]
    [InlineData("nested")]
    [InlineData("unclosed")]
    [InlineData("no-base")]
    public void AmbiguousMarkFirstPaintKeepsItsConservativeOrigin(string kind) {
        string basePaint = kind switch {
            "multiline" => "0 Ts 1 0 0 1 20 80 Tm (a) Tj",
            "mixed-font" => "0 Ts /F2 10 Tf (a) Tj",
            "mixed-transform" => "0 Ts 2 0 0 1 20 100 Tm (a) Tj",
            "mixed-clip" => "0 Ts q 0 0 200 200 re W n (a) Tj Q",
            "backward" => "0 Ts (a) Tj 1 0 0 1 15 100 Tm (b) Tj",
            "unmapped" => "0 Ts (z) Tj",
            "nested" => "0 Ts /Span << /ActualText (a) >> BDC (a) Tj EMC",
            "no-base" => string.Empty,
            _ => "0 Ts (a) Tj"
        };
        string content = Replacement("6 Ts (m) Tj " + basePaint, closed: kind != "unclosed");
        List<PdfTextSpan> spans = Parse(content + (kind == "unclosed" ? string.Empty : Tail));
        PdfTextSpan replacement = spans[0];

        Assert.Equal(106D, replacement.Y, 6);
        Assert.Equal(replacement.X, replacement.LogicalLayoutX, 6);
        Assert.Equal(replacement.Y, replacement.LogicalLayoutY, 6);
        if (kind != "unclosed") Assert.Equal(2, TextLayoutEngine.BuildLines(spans).Count);
    }

    [Fact]
    public void LayoutAndIdentityClonesKeepTheBaselineSeparateFromPaintGeometry() {
        PdfTextSpan original = Parse(Replacement("6 Ts (m) Tj 0 Ts (a) Tj"))[0];
        foreach (PdfTextSpan clone in new[] {
            original.WithCanRestamp(false), original.WithVisualFontSize(12D), original.WithVisualText(original.Text)
        }) {
            Assert.Equal(106D, clone.Y, 6);
            Assert.Equal(100D, clone.LogicalLayoutY, 6);
        }
        PdfTextSpan translated = original.WithOffset(10D, 20D);
        Assert.Equal(126D, translated.Y, 6);
        Assert.Equal(120D, translated.LogicalLayoutY, 6);
        PdfTextSpan rotatedLayout = original.WithLayoutGeometry(50D, 60D, 90D, 200D);
        Assert.Equal(50D, rotatedLayout.X, 6);
        Assert.Equal(56D, rotatedLayout.LogicalLayoutX, 6);
        Assert.Equal(60D, rotatedLayout.LogicalLayoutY, 6);
    }

    // The parser boundary receives a zero-width mark, ordinary advancing glyphs, and an
    // intentionally unmapped code. ActualText owns logical Unicode independently of glyph codes.
    private static List<PdfTextSpan> Parse(string content) => TextContentParser.Parse(content,
        static (_, bytes) => string.Concat(bytes.Select(value => value == (byte)'m' ? "\u0301"
            : value == (byte)'z' ? string.Empty : ((char)value).ToString())),
        static (_, bytes) => bytes.Sum(value => value == (byte)'m' ? 0D : 500D), pageHeight: 200D);

    private static string Replacement(string paint, bool closed = true) =>
        "BT /F1 10 Tf 1 0 0 1 20 100 Tm /Span << /ActualText <FEFF00610301> >> BDC " +
        paint + (closed ? " EMC" : string.Empty) + " ET ";

    private const string Tail = "BT /F1 10 Tf 0 Ts 1 0 0 1 35 100 Tm (tail) Tj ET";
}
