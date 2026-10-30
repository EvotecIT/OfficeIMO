using System;
using System.Linq;
using OfficeIMO.Pdf;
using Xunit;
using PdfPigDocument = UglyToad.PdfPig.PdfDocument;

namespace OfficeIMO.Tests.Pdf;

public sealed class PdfImportedLineFrameFitTests {
    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public void ImportedFrameFittingRetainsTheLastLineWhenAFloatingTableIsOnThePage(bool keep) {
        var tableStyle = new PdfTableStyle {
            HeaderRowCount = 0, FontSize = 2D, CellPaddingX = 0D, CellPaddingY = 0D, MinRowHeight = 4D,
            ColumnWidthPoints = new System.Collections.Generic.List<double?> { 10D },
            Position = new PdfTablePosition(PdfTableAnchor.Page, PdfTableAnchor.Page,
                horizontalOffset: 150D, verticalOffset: 4D)
        };
        var document = PdfDocument.Create(Options(107D)).Table(new[] { new[] { "T" } }, style: tableStyle)
            .Paragraph(p => p.Text("A\nB\nC\nD"), style: new PdfParagraphStyle {
                FontSize = global::OfficeIMO.OfficeLength.Points(10D), SpacingBefore = global::OfficeIMO.OfficeLength.Points(0D), SpacingAfter = global::OfficeIMO.OfficeLength.Points(0D), KeepTogether = keep,
                LineSpacing = PdfLineSpacing.Multiple(2.8D).WithFontLineBoxBaseline(1.4D)
            });
        using var pdf = PdfPigDocument.Open(document.ToBytes());
        var page = Assert.Single(pdf.GetPages());
        var letters = page.Letters.Where(letter => "ABCD".Contains(letter.Value)).ToArray();
        Assert.Equal("ABCD", string.Concat(letters.Select(letter => letter.Value)));
        for (int index = 1; index < letters.Length; index++)
            Assert.Equal(28D, letters[index - 1].StartBaseLine.Y - letters[index].StartBaseLine.Y, 3);
    }

    [Theory]
    [InlineData("paragraph", false, 1)]
    [InlineData("paragraph", true, 1)]
    [InlineData("paragraph", false, 4)]
    [InlineData("paragraph", true, 4)]
    [InlineData("heading", false, 1)]
    [InlineData("heading", false, 4)]
    [InlineData("list", false, 1)]
    [InlineData("list", true, 1)]
    [InlineData("list", false, 4)]
    [InlineData("list", true, 4)]
    public void ImportedMultipleSpacingFitsTheLastFontBoxAndRetainsLineAdvances(string kind, bool keep, int count) {
        double bodyHeight = count == 1 ? 20D : 107D;
        var spacing = PdfLineSpacing.Multiple(2.8D).WithFontLineBoxBaseline(1.4D);
        var document = PdfDocument.Create(Options(bodyHeight));
        string text = count == 1 ? "A" : "A\nB\nC\nD";
        if (kind == "heading") document.Content.Heading(1, text, style: new PdfHeadingStyle {
            FontSize = 10D, SpacingBefore = 0D, SpacingAfter = 0D, LineSpacing = spacing, KeepWithNext = false
        });
        else if (kind == "list") document.Content.Numbered(new[] { text }, style: new PdfListStyle {
            FontSize = 10D, MarkerFontSize = 10D, SpacingBefore = 0D, SpacingAfter = 0D,
            LineSpacing = spacing, KeepTogether = keep
        });
        else document.Paragraph(p => p.Text(text), style: new PdfParagraphStyle {
            FontSize = global::OfficeIMO.OfficeLength.Points(10D), SpacingBefore = global::OfficeIMO.OfficeLength.Points(0D), SpacingAfter = global::OfficeIMO.OfficeLength.Points(0D), LineSpacing = spacing, KeepTogether = keep
        });
        using var pdf = PdfPigDocument.Open(document.ToBytes());
        var page = Assert.Single(pdf.GetPages());
        var letters = page.Letters.Where(letter => "ABCD".Contains(letter.Value)).ToArray();
        Assert.Equal(count, letters.Length);
        for (int line = 1; line < letters.Length; line++)
            Assert.Equal(28D, letters[line - 1].StartBaseLine.Y - letters[line].StartBaseLine.Y, 3);
    }

    [Theory]
    [InlineData("exact")]
    [InlineData("minimum")]
    [InlineData("ordinary-multiple")]
    public void FixedMinimumAndOrdinarySpacingRetainTheirFullFrameRequirement(string kind) {
        PdfLineSpacing spacing = kind switch {
            "exact" => PdfLineSpacing.Exact(global::OfficeIMO.OfficeLength.Points(28D)).WithFontLineBoxBaseline(1.4D),
            "minimum" => PdfLineSpacing.AtLeast(global::OfficeIMO.OfficeLength.Points(28D)).WithFontLineBoxBaseline(1.4D),
            _ => PdfLineSpacing.Multiple(2.8D)
        };
        using var pdf = PdfPigDocument.Open(PdfDocument.Create(Options(107D)).Paragraph(p => p.Text("A\nB\nC\nD"),
            style: new PdfParagraphStyle { FontSize = global::OfficeIMO.OfficeLength.Points(10D), SpacingBefore = global::OfficeIMO.OfficeLength.Points(0D), SpacingAfter = global::OfficeIMO.OfficeLength.Points(0D), LineSpacing = spacing }).ToBytes());
        Assert.Equal(2, pdf.NumberOfPages);
        Assert.Equal("ABC", string.Concat(pdf.GetPage(1).Letters.Select(letter => letter.Value)));
        Assert.Equal("D", string.Concat(pdf.GetPage(2).Letters.Select(letter => letter.Value)));
    }

    [Fact]
    public void ImportedFrameFittingStillRejectsAnOversizedVisibleFontBox() {
        var document = PdfDocument.Create(Options(20D)).Paragraph(p => p.FontSize(30D).Text("A"),
            style: new PdfParagraphStyle { FontSize = global::OfficeIMO.OfficeLength.Points(10D), SpacingAfter = global::OfficeIMO.OfficeLength.Points(0D),
                LineSpacing = PdfLineSpacing.Multiple(2.8D).WithFontLineBoxBaseline(1.4D) });
        Assert.Throws<ArgumentException>(() => document.ToBytes());
    }

    private static PdfCreateOptions Options(double bodyHeight) => new() {
        PageWidth = global::OfficeIMO.OfficeLength.Points(200D), PageHeight = global::OfficeIMO.OfficeLength.Points(bodyHeight + 40D), DefaultFont = PdfStandardFont.Helvetica,
        DefaultFontSize = global::OfficeIMO.OfficeLength.Points(10D), MarginTop = global::OfficeIMO.OfficeLength.Points(20D), MarginBottom = global::OfficeIMO.OfficeLength.Points(20D), MarginLeft = global::OfficeIMO.OfficeLength.Points(20D), MarginRight = global::OfficeIMO.OfficeLength.Points(20D)
    };
}
