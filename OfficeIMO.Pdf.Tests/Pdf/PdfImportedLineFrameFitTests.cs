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
                FontSize = 10D, SpacingBefore = 0D, SpacingAfter = 0D, KeepTogether = keep,
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
            FontSize = 10D, SpacingBefore = 0D, SpacingAfter = 0D, LineSpacing = spacing, KeepTogether = keep
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
            "exact" => PdfLineSpacing.Exactly(28D).WithFontLineBoxBaseline(1.4D),
            "minimum" => PdfLineSpacing.AtLeast(28D).WithFontLineBoxBaseline(1.4D),
            _ => PdfLineSpacing.Multiple(2.8D)
        };
        using var pdf = PdfPigDocument.Open(PdfDocument.Create(Options(107D)).Paragraph(p => p.Text("A\nB\nC\nD"),
            style: new PdfParagraphStyle { FontSize = 10D, SpacingBefore = 0D, SpacingAfter = 0D, LineSpacing = spacing }).ToBytes());
        Assert.Equal(2, pdf.NumberOfPages);
        Assert.Equal("ABC", string.Concat(pdf.GetPage(1).Letters.Select(letter => letter.Value)));
        Assert.Equal("D", string.Concat(pdf.GetPage(2).Letters.Select(letter => letter.Value)));
    }

    [Fact]
    public void ImportedFrameFittingStillRejectsAnOversizedVisibleFontBox() {
        var document = PdfDocument.Create(Options(20D)).Paragraph(p => p.FontSize(30D).Text("A"),
            style: new PdfParagraphStyle { FontSize = 10D, SpacingAfter = 0D,
                LineSpacing = PdfLineSpacing.Multiple(2.8D).WithFontLineBoxBaseline(1.4D) });
        Assert.Throws<ArgumentException>(() => document.ToBytes());
    }

    [Theory]
    [InlineData("paragraph", 1D, 0D, false)]
    [InlineData("heading", 1D, 0D, false)]
    [InlineData("list", 1D, 0D, false)]
    [InlineData("paragraph", 0D, 0D, false)]
    [InlineData("paragraph", 1D, 1D, false)]
    [InlineData("paragraph", 1D, 0D, true)]
    [InlineData("section-heading", 0D, 0D, false)]
    public void ImportedStrictContainerMovesTheFinalAdvanceBeforeOpeningItsDecoration(
        string kind, double padding, double inset, bool nested) {
        var document = PdfDocument.Create(Options(50D)).Paragraph(p => p.Text("Lead"), style: new PdfParagraphStyle {
            FontSize = 10D, LineSpacing = PdfLineSpacing.Exactly(30D), SpacingBefore = 0D, SpacingAfter = 0D
        });
        var panelStyle = StrictPanel(padding, inset);
        document.Panel(panel => {
            if (nested) panel.Panel(inner => inner.Semantic(PdfSemanticRole.Division,
                content => AddImportedText(content, kind, "A")), panelStyle);
            else AddImportedText(panel, kind, "A");
        }, panelStyle);

        using var pdf = PdfPigDocument.Open(document.ToBytes());
        Assert.Equal(2, pdf.NumberOfPages);
        Assert.Contains("Lead", pdf.GetPage(1).Text, StringComparison.Ordinal);
        Assert.DoesNotContain("A", pdf.GetPage(1).Text, StringComparison.Ordinal);
        Assert.DoesNotContain(pdf.GetPage(1).Paths, path => path.IsStroked);
        var page = pdf.GetPage(2);
        var letter = Assert.Single(page.Letters, letter => letter.Value == "A");
        Assert.True(letter.BoundingBox.Bottom >= 20D);
        var borders = page.Paths.Where(path => path.IsStroked).Select(path => path.GetBoundingRectangle())
            .Where(bounds => bounds.HasValue).Select(bounds => bounds!.Value).ToArray();
        Assert.NotEmpty(borders);
        Assert.All(borders, bounds => Assert.True(bounds.Bottom >= 20D - .001D));
    }

    [Theory]
    [InlineData("paragraph")]
    [InlineData("list")]
    public void ImportedStrictContainerMovesKeptMultilineContentIntact(string kind) {
        var document = PdfDocument.Create(Options(70D)).Panel(panel => {
            panel.Paragraph(p => p.Text("Lead"), style: new PdfParagraphStyle {
                FontSize = 10D, LineSpacing = PdfLineSpacing.Exactly(20D), SpacingBefore = 0D, SpacingAfter = 0D
            });
            AddImportedText(panel, kind, "A\nB", keep: true);
        }, StrictPanel(0D));

        using var pdf = PdfPigDocument.Open(document.ToBytes());
        Assert.Equal(2, pdf.NumberOfPages);
        Assert.Contains("Lead", pdf.GetPage(1).Text, StringComparison.Ordinal);
        Assert.DoesNotContain("A", pdf.GetPage(1).Text, StringComparison.Ordinal);
        Assert.DoesNotContain("B", pdf.GetPage(1).Text, StringComparison.Ordinal);
        var letters = pdf.GetPage(2).Letters.Where(letter => letter.Value is "A" or "B").ToArray();
        Assert.Equal("AB", string.Concat(letters.Select(letter => letter.Value)));
        Assert.Equal(28D, letters[0].StartBaseLine.Y - letters[1].StartBaseLine.Y, 3);
    }

    [Fact]
    public void ImportedStrictContainerRetainsIntermediateFragmentFontBoxCapacity() {
        var document = PdfDocument.Create(Options(107D)).Panel(
            panel => AddImportedText(panel, "paragraph", "A\nB\nC\nD\nE"), StrictPanel(0D));
        using var pdf = PdfPigDocument.Open(document.ToBytes());
        Assert.Equal(2, pdf.NumberOfPages);
        Assert.Equal("ABCD", string.Concat(pdf.GetPage(1).Letters.Select(letter => letter.Value)));
        Assert.Equal("E", string.Concat(pdf.GetPage(2).Letters.Select(letter => letter.Value)));
        var letters = pdf.GetPage(1).Letters.ToArray();
        for (int line = 1; line < letters.Length; line++)
            Assert.Equal(28D, letters[line - 1].StartBaseLine.Y - letters[line].StartBaseLine.Y, 3);
    }

    [Fact]
    public void ImportedStrictContainerDoesNotReserveClosingAdvanceForASectionHeadingBeforeItsBody() {
        var document = PdfDocument.Create(Options(50D)).Paragraph(p => p.Text("Lead"), style: new PdfParagraphStyle {
            FontSize = 10D, LineSpacing = PdfLineSpacing.Exactly(30D), SpacingBefore = 0D, SpacingAfter = 0D
        }).Panel(panel => panel.Section("A", section => AddImportedText(section, "paragraph", "B"),
            new PdfSectionOptions { StartOnNewPage = false, HeadingStyle = ImportedHeadingStyle() }), StrictPanel(1D));
        using var pdf = PdfPigDocument.Open(document.ToBytes());
        Assert.Equal(2, pdf.NumberOfPages);
        Assert.Single(pdf.GetPage(1).Letters, letter => letter.Value == "A");
        Assert.DoesNotContain("B", pdf.GetPage(1).Text, StringComparison.Ordinal);
        Assert.Single(pdf.GetPage(2).Letters, letter => letter.Value == "B");
    }

    [Fact]
    public void ImportedRepeatingContainerRetainsItsFontBoxFittingPolicy() {
        var document = PdfDocument.Create(Options(50D)).Paragraph(p => p.Text("Lead"), style: new PdfParagraphStyle {
            FontSize = 10D, LineSpacing = PdfLineSpacing.Exactly(30D), SpacingBefore = 0D, SpacingAfter = 0D
        }).Panel(panel => AddImportedText(panel, "paragraph", "A"), new PdfPanelStyle {
            PaddingX = 0D, PaddingY = 0D, SpacingBefore = 0D, SpacingAfter = 0D, RepeatFragmentDecoration = true
        });
        using var pdf = PdfPigDocument.Open(document.ToBytes());
        var page = Assert.Single(pdf.GetPages());
        Assert.Contains("Lead", page.Text, StringComparison.Ordinal);
        Assert.Single(page.Letters, letter => letter.Value == "A");
    }

    [Fact]
    public void ImportedStrictContainerReservesItsFinalAdvanceWhenAdmittingHeaderAndFooterStories() {
        var document = PdfDocument.Create(Options(70D));
        document.Header(header => header.Text("HEADER TEXT").Content(
            content => content.Table(new[] { new[] { "FLOWHEADER" } }, style: RunningTableStyle(32D)), 18D));
        document.Footer(footer => footer.Content(
            content => content.Table(new[] { new[] { "FLOWFOOTER" } }, style: RunningTableStyle(20D)), 18D));
        document.Options.AllowRunningContentTextFallback(isHeader: true);
        document.Panel(panel => AddImportedText(panel, "paragraph", "A"), StrictPanel(0D));

        using var pdf = PdfPigDocument.Open(document.ToBytes());
        var page = Assert.Single(pdf.GetPages());
        Assert.Contains("HEADER TEXT", page.Text, StringComparison.Ordinal);
        Assert.DoesNotContain("FLOWHEADER", page.Text, StringComparison.Ordinal);
        Assert.Contains("FLOWFOOTER", page.Text, StringComparison.Ordinal);
        Assert.Single(page.GetWords(), word => word.Text == "A");
        var borders = page.Paths.Where(path => path.IsStroked).Select(path => path.GetBoundingRectangle())
            .Where(bounds => bounds.HasValue).Select(bounds => bounds!.Value).ToArray();
        Assert.NotEmpty(borders);
        Assert.All(borders, bounds => Assert.True(bounds.Bottom >= 38D - .001D));
    }

    private static PdfPanelStyle StrictPanel(double padding, double inset = 0D) => new() {
        PaddingX = 0D, PaddingY = padding, FragmentBottomInset = inset, RepeatFragmentDecoration = false,
        SpacingBefore = 0D, SpacingAfter = 0D, BorderColor = PdfColor.Black
    };

    private static PdfHeadingStyle ImportedHeadingStyle() => new() {
        FontSize = 10D, SpacingBefore = 0D, SpacingAfter = 0D, KeepWithNext = false,
        LineSpacing = PdfLineSpacing.Multiple(2.8D).WithFontLineBoxBaseline(1.4D)
    };

    private static void AddImportedText(PdfContentBuilder content, string kind, string text, bool keep = false) {
        var spacing = PdfLineSpacing.Multiple(2.8D).WithFontLineBoxBaseline(1.4D);
        if (kind == "heading") content.Heading(1, text, style: ImportedHeadingStyle());
        else if (kind == "section-heading") content.Section(text, _ => { }, new PdfSectionOptions {
            StartOnNewPage = false, HeadingStyle = ImportedHeadingStyle()
        });
        else if (kind == "list") content.Numbered(new[] { text }, style: new PdfListStyle {
            FontSize = 10D, MarkerFontSize = 10D, LineSpacing = spacing,
            SpacingBefore = 0D, SpacingAfter = 0D, ItemSpacing = 0D, KeepTogether = keep
        });
        else content.Paragraph(p => p.Text(text), style: new PdfParagraphStyle {
            FontSize = 10D, LineSpacing = spacing, SpacingBefore = 0D, SpacingAfter = 0D,
            KeepTogether = keep, WidowControl = false
        });
    }

    private static PdfTableStyle RunningTableStyle(double height) => new() {
        HeaderRowCount = 0, FontSize = 2D, CellPaddingX = 0D, CellPaddingY = 0D, BorderWidth = 0D,
        MinRowHeight = height, SpacingBefore = 0D, SpacingAfter = 0D
    };

    private static PdfOptions Options(double bodyHeight) => new() {
        PageWidth = 200D, PageHeight = bodyHeight + 40D, DefaultFont = PdfStandardFont.Helvetica,
        DefaultFontSize = 10D, MarginTop = 20D, MarginBottom = 20D, MarginLeft = 20D, MarginRight = 20D
    };
}
