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
            HeaderRowCount = 0, FontSize = Pt(2D), CellPaddingX = Pt(0D), CellPaddingY = Pt(0D), MinRowHeight = Pt(4D),
            ColumnWidths = new System.Collections.Generic.List<global::OfficeIMO.OfficeLength?> { Pt(10D) },
            Position = new PdfTablePosition(PdfTableAnchor.Page, PdfTableAnchor.Page,
                horizontalOffset: Pt(150D), verticalOffset: Pt(4D))
        };
        var document = PdfDocument.Create(Options(107D)).Table(new[] { new[] { "T" } }, style: tableStyle)
            .Paragraph(p => p.Text("A\nB\nC\nD"), style: new PdfParagraphStyle {
                FontSize = Pt(10D), SpacingBefore = Pt(0D), SpacingAfter = Pt(0D), KeepTogether = keep,
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
            FontSize = Pt(10D), SpacingBefore = Pt(0D), SpacingAfter = Pt(0D), LineSpacing = spacing, KeepWithNext = false
        });
        else if (kind == "list") document.Content.Numbered(new[] { text }, style: new PdfListStyle {
            FontSize = Pt(10D), MarkerFontSize = Pt(10D), SpacingBefore = Pt(0D), SpacingAfter = Pt(0D),
            LineSpacing = spacing, KeepTogether = keep
        });
        else document.Paragraph(p => p.Text(text), style: new PdfParagraphStyle {
            FontSize = Pt(10D), SpacingBefore = Pt(0D), SpacingAfter = Pt(0D), LineSpacing = spacing, KeepTogether = keep
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
            "exact" => PdfLineSpacing.Exact(Pt(28D)).WithFontLineBoxBaseline(1.4D),
            "minimum" => PdfLineSpacing.AtLeast(Pt(28D)).WithFontLineBoxBaseline(1.4D),
            _ => PdfLineSpacing.Multiple(2.8D)
        };
        using var pdf = PdfPigDocument.Open(PdfDocument.Create(Options(107D)).Paragraph(p => p.Text("A\nB\nC\nD"),
            style: new PdfParagraphStyle { FontSize = Pt(10D), SpacingBefore = Pt(0D), SpacingAfter = Pt(0D), LineSpacing = spacing }).ToBytes());
        Assert.Equal(2, pdf.NumberOfPages);
        Assert.Equal("ABC", string.Concat(pdf.GetPage(1).Letters.Select(letter => letter.Value)));
        Assert.Equal("D", string.Concat(pdf.GetPage(2).Letters.Select(letter => letter.Value)));
    }

    [Fact]
    public void ImportedFrameFittingStillRejectsAnOversizedVisibleFontBox() {
        var document = PdfDocument.Create(Options(20D)).Paragraph(p => p.FontSize(Pt(30D)).Text("A"),
            style: new PdfParagraphStyle { FontSize = Pt(10D), SpacingAfter = Pt(0D),
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
            FontSize = Pt(10D), LineSpacing = PdfLineSpacing.Exact(Pt(30D)), SpacingBefore = Pt(0D), SpacingAfter = Pt(0D)
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
                FontSize = Pt(10D), LineSpacing = PdfLineSpacing.Exact(Pt(20D)), SpacingBefore = Pt(0D), SpacingAfter = Pt(0D)
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
            FontSize = Pt(10D), LineSpacing = PdfLineSpacing.Exact(Pt(30D)), SpacingBefore = Pt(0D), SpacingAfter = Pt(0D)
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
            FontSize = Pt(10D), LineSpacing = PdfLineSpacing.Exact(Pt(30D)), SpacingBefore = Pt(0D), SpacingAfter = Pt(0D)
        }).Panel(panel => AddImportedText(panel, "paragraph", "A"), new PdfPanelStyle {
            PaddingX = Pt(0D), PaddingY = Pt(0D), SpacingBefore = Pt(0D), SpacingAfter = Pt(0D), RepeatFragmentDecoration = true
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
            content => content.Table(new[] { new[] { "FLOWHEADER" } }, style: RunningTableStyle(32D)), Pt(18D)));
        document.Footer(footer => footer.Content(
            content => content.Table(new[] { new[] { "FLOWFOOTER" } }, style: RunningTableStyle(20D)), Pt(18D)));
        document.Options.AllowRunningContentTextFallback(isHeader: true);
        document.Panel(panel => AddImportedText(panel, "paragraph", "A"), StrictPanel(0D));

        // Dropping the flow header in favour of its text fallback is a reported loss.
        using var pdf = PdfPigDocument.Open(document.ToBytes(options: new PdfSaveOptions { AllowLoss = true }));
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

    [Fact]
    public void StrictContainerMovesItsFirstKeptListBeforeOpeningItsDecoration() {
        var spacing = PdfLineSpacing.Multiple(2.8D).WithFontLineBoxBaseline(1.4D);
        var document = PdfDocument.Create(Options(70D)).Paragraph(p => p.Text("Lead"), style: new PdfParagraphStyle {
            FontSize = Pt(10D), LineSpacing = PdfLineSpacing.Exact(Pt(30D)), SpacingBefore = Pt(0D), SpacingAfter = Pt(0D)
        }).Panel(panel => panel.Numbered(new[] { "A\nB" }, style: new PdfListStyle {
            FontSize = Pt(10D), MarkerFontSize = Pt(10D), LineSpacing = spacing, KeepTogether = true,
            SpacingBefore = Pt(0D), SpacingAfter = Pt(0D), ItemSpacing = Pt(0D)
        }), StrictPanel(1D));

        using var pdf = PdfPigDocument.Open(document.ToBytes());
        Assert.Equal(2, pdf.NumberOfPages);
        Assert.Contains("Lead", pdf.GetPage(1).Text, StringComparison.Ordinal);
        Assert.DoesNotContain(pdf.GetPage(1).Paths, path => path.IsStroked);
        var page = pdf.GetPage(2);
        var letters = page.Letters.Where(letter => letter.Value is "A" or "B").ToArray();
        Assert.Equal("AB", string.Concat(letters.Select(letter => letter.Value)));
        Assert.Equal(28D, letters[0].StartBaseLine.Y - letters[1].StartBaseLine.Y, 3);
        AssertPanelTopBorder(pdf, 2, 90D);
    }

    [Theory]
    [InlineData(0D, 0D)]
    [InlineData(1D, 1D)]
    public void StrictContainerReservesTerminalSplitTableSpacingWithNoUnreservedPadding(double padding, double inset) {
        var document = PdfDocument.Create(Options(60D + padding + inset)).Panel(
            panel => panel.Table(new[] { new[] { "A\nB\nC\nD\nE\nF" } }, style: new PdfTableStyle {
                HeaderRowCount = 0, FontSize = Pt(10D), LineHeight = 2D, CellPaddingX = Pt(0D), CellPaddingY = Pt(0D),
                SpacingBefore = Pt(0D), SpacingAfter = Pt(5D), BorderWidth = Pt(0D),
                RowAllowBreakAcrossPages = new System.Collections.Generic.List<bool?> { true }
            }), StrictPanel(padding, inset));

        using var pdf = PdfPigDocument.Open(document.ToBytes());
        Assert.Equal(3, pdf.NumberOfPages);
        Assert.Equal("ABC", string.Concat(pdf.GetPage(1).Letters.Select(letter => letter.Value)));
        Assert.Equal("DE", string.Concat(pdf.GetPage(2).Letters.Select(letter => letter.Value)));
        Assert.Equal("F", string.Concat(pdf.GetPage(3).Letters.Select(letter => letter.Value)));
        Assert.All(pdf.GetPages().SelectMany(page => page.Paths).Where(path => path.IsStroked), path =>
            Assert.True(path.GetBoundingRectangle() is { } bounds && bounds.Bottom >= 20D - .001D));
    }

    [Fact]
    public void StrictContainerMovesItsFirstWholeTableWithTerminalSpacingBeforeDecoration() {
        const double padding = 1D, inset = 1D;
        var document = PdfDocument.Create(Options(60D + padding + inset)).Paragraph(p => p.Text("Lead"), style: new PdfParagraphStyle {
            FontSize = Pt(10D), LineSpacing = PdfLineSpacing.Exact(Pt(35D)), SpacingBefore = Pt(0D), SpacingAfter = Pt(0D)
        }).Panel(panel => panel.Table(new[] { new[] { "A" } }, style: new PdfTableStyle {
            HeaderRowCount = 0, FontSize = Pt(10D), LineHeight = 2D, CellPaddingX = Pt(0D), CellPaddingY = Pt(0D),
            SpacingBefore = Pt(0D), SpacingAfter = Pt(10D), BorderWidth = Pt(0D),
            RowAllowBreakAcrossPages = new System.Collections.Generic.List<bool?> { false }
        }), StrictPanel(padding, inset));

        using var pdf = PdfPigDocument.Open(document.ToBytes());
        Assert.Equal(2, pdf.NumberOfPages);
        Assert.Contains("Lead", pdf.GetPage(1).Text, StringComparison.Ordinal);
        Assert.DoesNotContain(pdf.GetPage(1).Paths, path => path.IsStroked);
        Assert.Equal("A", pdf.GetPage(2).Text);
        AssertPanelTopBorder(pdf, 2, 80D + padding + inset);
    }

    [Fact]
    public void StrictContainerMovesItsFirstKeptTableBeforeOpeningItsDecoration() {
        var document = PdfDocument.Create(Options(70D)).Paragraph(p => p.Text("Lead"), style: new PdfParagraphStyle {
            FontSize = Pt(10D), LineSpacing = PdfLineSpacing.Exact(Pt(30D)), SpacingBefore = Pt(0D), SpacingAfter = Pt(0D)
        }).Panel(panel => panel.Table(new[] { new[] { "A" }, new[] { "B" } }, style: new PdfTableStyle {
            HeaderRowCount = 0, FontSize = Pt(10D), LineHeight = 2D, CellPaddingX = Pt(0D), CellPaddingY = Pt(0D),
            SpacingBefore = Pt(0D), SpacingAfter = Pt(5D), BorderWidth = Pt(0D), KeepTogether = true
        }), StrictPanel(1D));

        using var pdf = PdfPigDocument.Open(document.ToBytes());
        Assert.Equal(2, pdf.NumberOfPages);
        Assert.Contains("Lead", pdf.GetPage(1).Text, StringComparison.Ordinal);
        Assert.DoesNotContain(pdf.GetPage(1).Paths, path => path.IsStroked);
        Assert.Equal("AB", string.Concat(pdf.GetPage(2).Letters.Select(letter => letter.Value)));
        AssertPanelTopBorder(pdf, 2, 90D);
    }

    [Theory]
    [InlineData("A", 35D, 10D, "", "A")]
    [InlineData("A\nB\nC", 0D, 5D, "AB", "C")]
    public void StrictContainerReservesItsTerminalRowSpacing(
        string text, double leadHeight, double rowSpacingAfter, string firstPageText, string secondPageText) {
        var document = PdfDocument.Create(Options(60D));
        if (leadHeight > 0D) document.Paragraph(p => p.Text("Lead"), style: new PdfParagraphStyle {
            FontSize = Pt(10D), LineSpacing = PdfLineSpacing.Exact(Pt(leadHeight)), SpacingBefore = Pt(0D), SpacingAfter = Pt(0D)
        });
        document.Panel(panel => panel.Row(row => row.Style(new PdfRowStyle {
            Gap = Pt(0D), SpacingBefore = Pt(0D), SpacingAfter = Pt(rowSpacingAfter), KeepTogether = false
        }).PercentColumn(100D, column => column.Paragraph(p => p.Text(text), style: new PdfParagraphStyle {
            FontSize = Pt(10D), LineSpacing = PdfLineSpacing.Exact(Pt(20D)), SpacingBefore = Pt(0D), SpacingAfter = Pt(0D),
            KeepTogether = false, WidowControl = false
        }))), StrictPanel(0D));

        using var pdf = PdfPigDocument.Open(document.ToBytes());
        Assert.Equal(2, pdf.NumberOfPages);
        if (leadHeight > 0D) {
            Assert.Contains("Lead", pdf.GetPage(1).Text, StringComparison.Ordinal);
            Assert.DoesNotContain(pdf.GetPage(1).Paths, path => path.IsStroked);
        } else Assert.Equal(firstPageText, string.Concat(pdf.GetPage(1).Letters.Select(letter => letter.Value)));
        Assert.Equal(secondPageText, string.Concat(pdf.GetPage(2).Letters.Select(letter => letter.Value)));
        Assert.All(pdf.GetPages().SelectMany(page => page.Paths).Where(path => path.IsStroked), path =>
            Assert.True(path.GetBoundingRectangle() is { } bounds && bounds.Bottom >= 20D - .001D));
    }

    [Theory]
    [InlineData(false, false)]
    [InlineData(true, true)]
    public void ImportedStrictContainerStartsOversizedKeptListBeforeDecoration(bool nested, bool splitPages) {
        var document = PdfDocument.Create(Options(70D)).Paragraph(p => p.Text("Lead"), style: new PdfParagraphStyle {
            FontSize = Pt(10D), LineSpacing = PdfLineSpacing.Exact(Pt(30D)), SpacingBefore = Pt(0D), SpacingAfter = Pt(0D)
        });
        void AddList(PdfContentBuilder content) => content.Numbered(new[] { "A\nB\nC" }, style: new PdfListStyle {
            FontSize = Pt(10D), MarkerFontSize = Pt(10D), KeepTogether = true,
            LineSpacing = PdfLineSpacing.Multiple(2.8D).WithFontLineBoxBaseline(1.4D),
            SpacingBefore = Pt(0D), SpacingAfter = Pt(0D), ItemSpacing = Pt(0D),
            KeepTogetherOverflow = splitPages ? PdfKeepTogetherOverflowBehavior.SplitAcrossPages :
                PdfKeepTogetherOverflowBehavior.SplitAcrossFrames
        });
        document.Panel(panel => {
            if (nested) panel.Panel(inner => inner.Semantic(PdfSemanticRole.Division, AddList), StrictPanel(1D));
            else AddList(panel);
        }, StrictPanel(1D));

        using var pdf = PdfPigDocument.Open(document.ToBytes());
        Assert.Equal(3, pdf.NumberOfPages);
        Assert.Contains("Lead", pdf.GetPage(1).Text, StringComparison.Ordinal);
        Assert.DoesNotContain(pdf.GetPage(1).Paths, path => path.IsStroked);
        Assert.Equal("AB", string.Concat(pdf.GetPage(2).Letters.Where(letter => letter.Value is "A" or "B" or "C")
            .Select(letter => letter.Value)));
        Assert.Equal("C", string.Concat(pdf.GetPage(3).Letters.Where(letter => letter.Value is "A" or "B" or "C")
            .Select(letter => letter.Value)));
        AssertPanelTopBorder(pdf, 2, 90D);
    }

    [Theory]
    [InlineData("table")]
    [InlineData("row")]
    [InlineData("column-paragraph")]
    [InlineData("column-table")]
    public void KeptContentIsAdmittedByItsDestinationFrameNotItsSpacingBefore(string kind) {
        var document = PdfDocument.Create(Options(60D)).Paragraph(p => p.Text("Lead"), style: new PdfParagraphStyle {
            FontSize = Pt(10D), LineSpacing = PdfLineSpacing.Exact(Pt(20D)), SpacingBefore = Pt(0D), SpacingAfter = Pt(0D)
        });
        var table = new PdfTableStyle {
            HeaderRowCount = 0, FontSize = Pt(10D), LineHeight = 5D, CellPaddingX = Pt(0D), CellPaddingY = Pt(0D),
            SpacingBefore = Pt(30D), SpacingAfter = Pt(0D), BorderWidth = Pt(0D), KeepTogether = true
        };
        var paragraph = new PdfParagraphStyle {
            FontSize = Pt(10D), LineSpacing = PdfLineSpacing.Exact(Pt(50D)), SpacingBefore = Pt(30D), SpacingAfter = Pt(0D), KeepTogether = true
        };
        void Row(Action<PdfContentBuilder> content) => document.Content.Row(row => row.Style(new PdfRowStyle {
            Gap = Pt(0D), SpacingBefore = Pt(0D), SpacingAfter = Pt(0D)
        }).PercentColumn(100D, column => {
            column.Spacer(Pt(5D));
            content(column);
        }));
        if (kind == "table") document.Content.Table(new[] { new[] { "A" } }, style: table);
        else if (kind == "row") document.Content.Row(row => row.Style(new PdfRowStyle {
            Gap = Pt(0D), SpacingBefore = Pt(30D), SpacingAfter = Pt(0D), KeepTogether = true
        }).PercentColumn(100D, column => column.Paragraph(p => p.Text("A"), style: new PdfParagraphStyle {
            FontSize = Pt(10D), LineSpacing = PdfLineSpacing.Exact(Pt(50D)), SpacingBefore = Pt(0D), SpacingAfter = Pt(0D)
        })));
        else if (kind == "column-paragraph") Row(column => column.Paragraph(p => p.Text("A"), style: paragraph));
        else Row(column => column.Table(new[] { new[] { "A" } }, style: table));

        using var pdf = PdfPigDocument.Open(document.ToBytes());
        Assert.Equal(2, pdf.NumberOfPages);
        Assert.Contains("A", pdf.GetPage(2).Text, StringComparison.Ordinal);
    }

    private static void AssertPanelTopBorder(PdfPigDocument document, int pageNumber, double y) =>
        Assert.Contains(document.GetPage(pageNumber).Paths.Where(path => path.IsStroked)
            .SelectMany(path => path).SelectMany(subpath => subpath.Commands).OfType<UglyToad.PdfPig.Core.PdfSubpath.Line>(),
            line => Math.Abs(line.From.Y - y) < .001D && Math.Abs(line.To.Y - y) < .001D &&
                Math.Abs(Math.Min(line.From.X, line.To.X) - 20D) < .001D &&
                Math.Abs(Math.Max(line.From.X, line.To.X) - 180D) < .001D);

    private static PdfPanelStyle StrictPanel(double padding, double inset = 0D) => new() {
        PaddingX = Pt(0D), PaddingY = Pt(padding), FragmentBottomInset = Pt(inset), RepeatFragmentDecoration = false,
        SpacingBefore = Pt(0D), SpacingAfter = Pt(0D), BorderColor = global::OfficeIMO.OfficeColor.FromRgb(0, 0, 0)
    };

    private static PdfHeadingStyle ImportedHeadingStyle() => new() {
        FontSize = Pt(10D), SpacingBefore = Pt(0D), SpacingAfter = Pt(0D), KeepWithNext = false,
        LineSpacing = PdfLineSpacing.Multiple(2.8D).WithFontLineBoxBaseline(1.4D)
    };

    private static void AddImportedText(PdfContentBuilder content, string kind, string text, bool keep = false) {
        var spacing = PdfLineSpacing.Multiple(2.8D).WithFontLineBoxBaseline(1.4D);
        if (kind == "heading") content.Heading(1, text, style: ImportedHeadingStyle());
        else if (kind == "section-heading") content.Section(text, _ => { }, new PdfSectionOptions {
            StartOnNewPage = false, HeadingStyle = ImportedHeadingStyle()
        });
        else if (kind == "list") content.Numbered(new[] { text }, style: new PdfListStyle {
            FontSize = Pt(10D), MarkerFontSize = Pt(10D), LineSpacing = spacing,
            SpacingBefore = Pt(0D), SpacingAfter = Pt(0D), ItemSpacing = Pt(0D), KeepTogether = keep
        });
        else content.Paragraph(p => p.Text(text), style: new PdfParagraphStyle {
            FontSize = Pt(10D), LineSpacing = spacing, SpacingBefore = Pt(0D), SpacingAfter = Pt(0D),
            KeepTogether = keep, WidowControl = false
        });
    }

    private static PdfTableStyle RunningTableStyle(double height) => new() {
        HeaderRowCount = 0, FontSize = Pt(2D), CellPaddingX = Pt(0D), CellPaddingY = Pt(0D), BorderWidth = Pt(0D),
        MinRowHeight = Pt(height), SpacingBefore = Pt(0D), SpacingAfter = Pt(0D)
    };

    private static PdfCreateOptions Options(double bodyHeight) => new() {
        PageWidth = Pt(200D), PageHeight = Pt(bodyHeight + 40D), DefaultFont = PdfStandardFont.Helvetica,
        DefaultFontSize = Pt(10D), MarginTop = Pt(20D), MarginBottom = Pt(20D), MarginLeft = Pt(20D), MarginRight = Pt(20D)
    };

    private static global::OfficeIMO.OfficeLength Pt(double points) => global::OfficeIMO.OfficeLength.Points(points);
}
