using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using OfficeIMO.Drawing;
using OfficeIMO.Pdf;
using PdfPigDocument = UglyToad.PdfPig.PdfDocument;
using Xunit;

namespace OfficeIMO.Tests.Pdf;

public partial class PdfDocumentVisualQualityTests {
    [Fact]
    public void PdfDocument_DefaultTextStyleAppliesToFollowingContent() {
        byte[] bytes = PdfDocument.Create(new PdfOptions {
                PageWidth = 300,
                PageHeight = 220,
                MarginLeft = 30,
                MarginRight = 30,
                MarginTop = 30,
                MarginBottom = 30
            })
            .DefaultTextStyle(style => style
                .Font(PdfStandardFont.Helvetica)
                .FontSize(16)
                .Color(PdfColor.FromRgb(10, 20, 30)))
            .Paragraph(p => p.Text("FluentTextStyle"))
            .ToBytes();

        using var pdf = PdfPigDocument.Open(new MemoryStream(bytes));
        var page = pdf.GetPage(1);
        double pointSize = page.Letters
            .Where(letter => !string.IsNullOrWhiteSpace(letter.Value))
            .Select(letter => letter.PointSize)
            .First();
        string rawPdf = Encoding.ASCII.GetString(bytes);

        Assert.InRange(pointSize, 15.5, 16.5);
        Assert.Contains("0.039 0.078 0.118 rg", rawPdf, StringComparison.Ordinal);
    }

    [Fact]
    public void PdfDocument_DefaultTextStyleObjectAppliesToFollowingContentAndSnapshotsInput() {
        var style = new PdfTextStyle {
            Font = PdfStandardFont.Helvetica,
            FontSize = 16,
            Color = PdfColor.FromRgb(10, 20, 30)
        };

        byte[] bytes = PdfDocument.Create(new PdfOptions {
                PageWidth = 300,
                PageHeight = 220,
                MarginLeft = 30,
                MarginRight = 30,
                MarginTop = 30,
                MarginBottom = 30
            })
            .DefaultTextStyle(style)
            .Paragraph(p => p.Text("ObjectTextStyle"))
            .ToBytes();

        style.FontSize = 8;
        style.Color = PdfColor.Black;

        using var pdf = PdfPigDocument.Open(new MemoryStream(bytes));
        var page = pdf.GetPage(1);
        double pointSize = page.Letters
            .Where(letter => !string.IsNullOrWhiteSpace(letter.Value))
            .Select(letter => letter.PointSize)
            .First();
        string rawPdf = Encoding.ASCII.GetString(bytes);

        Assert.InRange(pointSize, 15.5, 16.5);
        Assert.Contains("0.039 0.078 0.118 rg", rawPdf, StringComparison.Ordinal);
    }

    [Fact]
    public void PdfDocument_DefaultHeadingStyleAppliesToFollowingHeadingsAndSnapshotsInput() {
        var style = new PdfHeadingStyle {
            FontSize = 12,
            LineHeight = 1,
            SpacingAfter = 24,
            Color = PdfColor.FromRgb(10, 20, 30)
        };

        byte[] bytes = PdfDocument.Create(new PdfOptions {
                PageWidth = 300,
                PageHeight = 220,
                MarginLeft = 30,
                MarginRight = 30,
                MarginTop = 30,
                MarginBottom = 30,
                DefaultFont = PdfStandardFont.Helvetica,
                DefaultFontSize = 10
            })
            .DefaultHeadingStyle(3, style)
            .H3("StyledHeading")
            .Paragraph(p => p.Text("StyledBody"))
            .ToBytes();

        style.FontSize = 30;
        style.SpacingAfter = 0;
        style.Color = PdfColor.Black;

        using var pdf = PdfPigDocument.Open(new MemoryStream(bytes));
        var page = pdf.GetPage(1);
        double pointSize = page.Letters
            .Where(letter => !string.IsNullOrWhiteSpace(letter.Value))
            .Select(letter => letter.PointSize)
            .First();
        var lines = GetVisualTextLines(page, 0, 300);
        var headingLine = lines.Single(line => line.Text.Contains("StyledHeading", StringComparison.Ordinal));
        var bodyLine = lines.Single(line => line.Text.Contains("StyledBody", StringComparison.Ordinal));
        double baselineGap = headingLine.BaselineY - bodyLine.BaselineY;
        string rawPdf = Encoding.ASCII.GetString(bytes);

        Assert.InRange(pointSize, 11.5, 12.5);
        Assert.InRange(baselineGap, 28, 42);
        Assert.Contains("0.039 0.078 0.118 rg", rawPdf, StringComparison.Ordinal);
    }

    [Fact]
    public void HeadingStyle_BoldFalse_UsesNormalFontResource() {
        byte[] bytes = PdfDocument.Create(new PdfOptions {
                PageWidth = 300,
                PageHeight = 220,
                MarginLeft = 30,
                MarginRight = 30,
                MarginTop = 30,
                MarginBottom = 30,
                DefaultFont = PdfStandardFont.Helvetica,
                DefaultFontSize = 10
            })
            .H1("RegularHeading", style: new PdfHeadingStyle {
                FontSize = 14,
                Bold = false
            })
            .ToBytes();

        string rawPdf = Encoding.ASCII.GetString(bytes);
        using var pdf = PdfPigDocument.Open(new MemoryStream(bytes));

        Assert.Contains("RegularHeading", pdf.GetPage(1).Text);
        Assert.Contains("/BaseFont /Helvetica", rawPdf, StringComparison.Ordinal);
        Assert.DoesNotContain("/Helvetica-Bold", rawPdf, StringComparison.Ordinal);
    }

    [Fact]
    public void Heading_UsesConfiguredSpacingBeforeAndAfter() {
        var options = new PdfOptions {
            PageWidth = 320,
            PageHeight = 260,
            MarginLeft = 30,
            MarginRight = 30,
            MarginTop = 30,
            MarginBottom = 30,
            DefaultFont = PdfStandardFont.Helvetica,
            DefaultFontSize = 10
        };
        var defaultStyle = new PdfHeadingStyle {
            FontSize = 12,
            LineHeight = 1,
            SpacingBefore = 0,
            SpacingAfter = 0
        };
        var spacedStyle = new PdfHeadingStyle {
            FontSize = 12,
            LineHeight = 1,
            SpacingBefore = 12,
            SpacingAfter = 18
        };

        byte[] defaultBytes = PdfDocument.Create(options)
            .Paragraph(p => p.Text("BeforeMarker"), style: new PdfParagraphStyle { SpacingAfter = 0 })
            .H2("HeadingMarker", style: defaultStyle)
            .Paragraph(p => p.Text("AfterMarker"))
            .ToBytes();
        byte[] spacedBytes = PdfDocument.Create(options)
            .Paragraph(p => p.Text("BeforeMarker"), style: new PdfParagraphStyle { SpacingAfter = 0 })
            .H2("HeadingMarker", style: spacedStyle)
            .Paragraph(p => p.Text("AfterMarker"))
            .ToBytes();

        using var defaultPdf = PdfPigDocument.Open(new MemoryStream(defaultBytes));
        using var spacedPdf = PdfPigDocument.Open(new MemoryStream(spacedBytes));
        var defaultPage = defaultPdf.GetPage(1);
        var spacedPage = spacedPdf.GetPage(1);

        double defaultHeadingY = FindWordStartY(defaultPage, "HeadingMarker");
        double spacedHeadingY = FindWordStartY(spacedPage, "HeadingMarker");
        double defaultAfterY = FindWordStartY(defaultPage, "AfterMarker");
        double spacedAfterY = FindWordStartY(spacedPage, "AfterMarker");

        Assert.True(defaultHeadingY - spacedHeadingY >= 10, $"Expected heading spacing before to move heading text down. Default y: {defaultHeadingY:0.##}, spaced y: {spacedHeadingY:0.##}.");
        Assert.True(defaultAfterY - spacedAfterY >= 28, $"Expected heading spacing before and after to move following content down. Default y: {defaultAfterY:0.##}, spaced y: {spacedAfterY:0.##}.");
    }

    [Fact]
    public void Heading_SuppressesSpacingBeforeAtPageTop() {
        var options = new PdfOptions {
            PageWidth = 320,
            PageHeight = 220,
            MarginLeft = 30,
            MarginRight = 30,
            MarginTop = 30,
            MarginBottom = 30,
            DefaultFont = PdfStandardFont.Helvetica,
            DefaultFontSize = 10
        };
        var defaultStyle = new PdfHeadingStyle {
            FontSize = 12,
            LineHeight = 1,
            SpacingBefore = 0,
            SpacingAfter = 0
        };
        var spacedStyle = new PdfHeadingStyle {
            FontSize = 12,
            LineHeight = 1,
            SpacingBefore = 28,
            SpacingAfter = 0
        };

        byte[] defaultBytes = PdfDocument.Create(options)
            .H2("TopHeadingMarker", style: defaultStyle)
            .ToBytes();
        byte[] spacedBytes = PdfDocument.Create(options)
            .H2("TopHeadingMarker", style: spacedStyle)
            .ToBytes();

        using var defaultPdf = PdfPigDocument.Open(new MemoryStream(defaultBytes));
        using var spacedPdf = PdfPigDocument.Open(new MemoryStream(spacedBytes));

        double defaultTopY = FindWordStartY(defaultPdf.GetPage(1), "TopHeadingMarker");
        double spacedTopY = FindWordStartY(spacedPdf.GetPage(1), "TopHeadingMarker");

        Assert.InRange(Math.Abs(defaultTopY - spacedTopY), 0, 1.5);
    }

    [Fact]
    public void Heading_CanApplySpacingBeforeAtPageTop() {
        var options = new PdfOptions {
            PageWidth = 320,
            PageHeight = 220,
            MarginLeft = 30,
            MarginRight = 30,
            MarginTop = 30,
            MarginBottom = 30,
            DefaultFont = PdfStandardFont.Helvetica,
            DefaultFontSize = 10
        };
        var defaultStyle = new PdfHeadingStyle {
            FontSize = 12,
            LineHeight = 1,
            SpacingBefore = 0,
            SpacingAfter = 0
        };
        var spacedStyle = new PdfHeadingStyle {
            FontSize = 12,
            LineHeight = 1,
            SpacingBefore = 28,
            SpacingAfter = 0,
            ApplySpacingBeforeAtTop = true
        };

        byte[] defaultBytes = PdfDocument.Create(options)
            .H2("TopHeadingMarker", style: defaultStyle)
            .ToBytes();
        byte[] spacedBytes = PdfDocument.Create(options)
            .H2("TopHeadingMarker", style: spacedStyle)
            .ToBytes();

        using var defaultPdf = PdfPigDocument.Open(new MemoryStream(defaultBytes));
        using var spacedPdf = PdfPigDocument.Open(new MemoryStream(spacedBytes));

        double defaultTopY = FindWordStartY(defaultPdf.GetPage(1), "TopHeadingMarker");
        double spacedTopY = FindWordStartY(spacedPdf.GetPage(1), "TopHeadingMarker");

        Assert.True(defaultTopY - spacedTopY >= 26, $"Expected opt-in top spacing to move heading text down. Default y: {defaultTopY:0.##}, spaced y: {spacedTopY:0.##}.");
    }

    [Fact]
    public void Heading_CanApplySpacingBeforeAfterAutomaticPageBreak() {
        var options = new PdfOptions {
            PageWidth = 320,
            PageHeight = 160,
            MarginLeft = 30,
            MarginRight = 30,
            MarginTop = 30,
            MarginBottom = 30,
            DefaultFont = PdfStandardFont.Helvetica,
            DefaultFontSize = 10
        };
        var defaultStyle = new PdfHeadingStyle {
            FontSize = 12,
            LineHeight = 1,
            SpacingBefore = 0,
            SpacingAfter = 0
        };
        var spacedStyle = new PdfHeadingStyle {
            FontSize = 12,
            LineHeight = 1,
            SpacingBefore = 24,
            SpacingAfter = 0,
            ApplySpacingBeforeAtTop = true
        };

        byte[] defaultBytes = PdfDocument.Create(options)
            .Paragraph(p => p.Text("BeforeMarker"), style: new PdfParagraphStyle { SpacingAfter = 0 })
            .Spacer(80)
            .H2("PagedHeadingMarker", style: defaultStyle)
            .ToBytes();
        byte[] spacedBytes = PdfDocument.Create(options)
            .Paragraph(p => p.Text("BeforeMarker"), style: new PdfParagraphStyle { SpacingAfter = 0 })
            .Spacer(80)
            .H2("PagedHeadingMarker", style: spacedStyle)
            .ToBytes();

        using var defaultPdf = PdfPigDocument.Open(new MemoryStream(defaultBytes));
        using var spacedPdf = PdfPigDocument.Open(new MemoryStream(spacedBytes));

        Assert.Equal(2, defaultPdf.NumberOfPages);
        Assert.Equal(2, spacedPdf.NumberOfPages);

        double defaultTopY = FindWordStartY(defaultPdf.GetPage(2), "PagedHeadingMarker");
        double spacedTopY = FindWordStartY(spacedPdf.GetPage(2), "PagedHeadingMarker");

        Assert.True(defaultTopY - spacedTopY >= 22, $"Expected opt-in top spacing after a page break to move heading text down. Default y: {defaultTopY:0.##}, spaced y: {spacedTopY:0.##}.");
    }

    [Fact]
    public void RowColumnHeading_SuppressesSpacingBeforeAtColumnTop() {
        var options = new PdfOptions {
            PageWidth = 320,
            PageHeight = 220,
            MarginLeft = 30,
            MarginRight = 30,
            MarginTop = 30,
            MarginBottom = 30,
            DefaultFont = PdfStandardFont.Helvetica,
            DefaultFontSize = 10
        };
        var defaultStyle = new PdfHeadingStyle {
            FontSize = 12,
            LineHeight = 1,
            SpacingBefore = 0,
            SpacingAfter = 0
        };
        var spacedStyle = new PdfHeadingStyle {
            FontSize = 12,
            LineHeight = 1,
            SpacingBefore = 28,
            SpacingAfter = 0
        };

        byte[] defaultBytes = PdfDocument.Create(options)
            .Compose(document => document.Page(page => page.Content(content => content.Row(row => row.Column(100, column => column
                .H2("ColumnHeadingMarker", style: defaultStyle))))))
            .ToBytes();
        byte[] spacedBytes = PdfDocument.Create(options)
            .Compose(document => document.Page(page => page.Content(content => content.Row(row => row.Column(100, column => column
                .H2("ColumnHeadingMarker", style: spacedStyle))))))
            .ToBytes();

        using var defaultPdf = PdfPigDocument.Open(new MemoryStream(defaultBytes));
        using var spacedPdf = PdfPigDocument.Open(new MemoryStream(spacedBytes));

        double defaultTopY = FindWordStartY(defaultPdf.GetPage(1), "ColumnHeadingMarker");
        double spacedTopY = FindWordStartY(spacedPdf.GetPage(1), "ColumnHeadingMarker");

        Assert.InRange(Math.Abs(defaultTopY - spacedTopY), 0, 1.5);
    }

    [Theory]
    [InlineData("bullet-list", "ListTopMarker")]
    [InlineData("numbered-list", "ListTopMarker")]
    [InlineData("panel", "PanelTopMarker")]
    [InlineData("horizontal-rule", "AfterFixedMarker")]
    [InlineData("image", "AfterFixedMarker")]
    [InlineData("shape", "AfterFixedMarker")]
    [InlineData("drawing", "AfterFixedMarker")]
    [InlineData("row", "RowTopMarker")]
    public void FlowBlock_SuppressesSpacingBeforeAtPageTop(string blockKind, string marker) {
        byte[] defaultBytes = CreateTopLevelFlowSpacingBeforeProbe(blockKind, 0);
        byte[] spacedBytes = CreateTopLevelFlowSpacingBeforeProbe(blockKind, 28);

        using var defaultPdf = PdfPigDocument.Open(new MemoryStream(defaultBytes));
        using var spacedPdf = PdfPigDocument.Open(new MemoryStream(spacedBytes));

        double defaultTopY = FindWordStartY(defaultPdf.GetPage(1), marker);
        double spacedTopY = FindWordStartY(spacedPdf.GetPage(1), marker);

        Assert.InRange(Math.Abs(defaultTopY - spacedTopY), 0, 1.5);
    }

    [Theory]
    [InlineData("bullet-list", "ColumnListMarker")]
    [InlineData("numbered-list", "ColumnListMarker")]
    [InlineData("panel", "ColumnPanelMarker")]
    [InlineData("horizontal-rule", "ColumnAfterFixedMarker")]
    [InlineData("image", "ColumnAfterFixedMarker")]
    [InlineData("shape", "ColumnAfterFixedMarker")]
    [InlineData("drawing", "ColumnAfterFixedMarker")]
    public void RowColumnFlowBlock_SuppressesSpacingBeforeAtColumnTop(string blockKind, string marker) {
        byte[] defaultBytes = CreateColumnFlowSpacingBeforeProbe(blockKind, 0);
        byte[] spacedBytes = CreateColumnFlowSpacingBeforeProbe(blockKind, 28);

        using var defaultPdf = PdfPigDocument.Open(new MemoryStream(defaultBytes));
        using var spacedPdf = PdfPigDocument.Open(new MemoryStream(spacedBytes));

        double defaultTopY = FindWordStartY(defaultPdf.GetPage(1), marker);
        double spacedTopY = FindWordStartY(spacedPdf.GetPage(1), marker);

        Assert.InRange(Math.Abs(defaultTopY - spacedTopY), 0, 1.5);
    }

    [Fact]
    public void PdfDocument_DefaultPanelStyleAppliesToFollowingPanelsAndSnapshotsInput() {
        var style = new PanelStyle {
            Background = PdfColor.FromRgb(240, 248, 255),
            BorderColor = PdfColor.FromRgb(20, 40, 60),
            BorderWidth = 0.8,
            PaddingX = 14,
            PaddingY = 8,
            MaxWidth = 180,
            Align = PdfAlign.Center,
            SpacingAfter = 16
        };
        var options = new PdfOptions {
            PageWidth = 360,
            PageHeight = 240,
            MarginLeft = 30,
            MarginRight = 30,
            MarginTop = 30,
            MarginBottom = 30,
            DefaultFont = PdfStandardFont.Helvetica,
            DefaultFontSize = 10
        };

        byte[] bytes = PdfDocument.Create(options)
            .DefaultPanelStyle(style)
            .PanelParagraph(p => p.Text("DefaultPanel"))
            .Paragraph(p => p.Text("AfterPanel"), style: new PdfParagraphStyle {
                SpacingAfter = 0
            })
            .ToBytes();

        style.PaddingX = 2;
        style.MaxWidth = 300;
        style.Align = PdfAlign.Right;
        style.Background = PdfColor.Black;

        using var pdf = PdfPigDocument.Open(new MemoryStream(bytes));
        var page = pdf.GetPage(1);
        double panelTextX = FindWordStartX(page, "DefaultPanel");
        double afterPanelY = FindWordStartY(page, "AfterPanel");
        double panelTextY = FindWordStartY(page, "DefaultPanel");
        string rawPdf = Encoding.ASCII.GetString(bytes);

        Assert.InRange(panelTextX, 103, 106);
        Assert.True(panelTextY - afterPanelY >= 28, $"Expected default panel spacing to leave visible rhythm. Panel y: {panelTextY:0.##}, after y: {afterPanelY:0.##}.");
        Assert.Contains("0.941 0.973 1 rg", rawPdf, StringComparison.Ordinal);
    }

    [Fact]
    public void PdfDocument_DefaultHorizontalRuleStyleAppliesToFollowingRulesAndSnapshotsInput() {
        const double fontSize = 10;
        var style = new PdfHorizontalRuleStyle {
            Thickness = 2,
            Color = PdfColor.FromRgb(10, 20, 30),
            SpacingBefore = 3,
            SpacingAfter = 15
        };

        byte[] bytes = PdfDocument.Create(new PdfOptions {
                PageWidth = 240,
                PageHeight = 180,
                MarginLeft = 20,
                MarginRight = 20,
                MarginTop = 20,
                MarginBottom = 20,
                DefaultFont = PdfStandardFont.Helvetica,
                DefaultFontSize = fontSize
            })
            .DefaultHorizontalRuleStyle(style)
            .HR()
            .Paragraph(p => p.Text("AfterDefaultRule"))
            .ToBytes();

        style.Thickness = 7;
        style.Color = PdfColor.FromRgb(200, 10, 10);
        style.SpacingAfter = 0;

        using var pdf = PdfPigDocument.Open(new MemoryStream(bytes));
        var page = pdf.GetPage(1);
        string rawPdf = Encoding.ASCII.GetString(bytes);
        double ruleBottomY = 180 - 20 - 2;
        double paragraphTopY = FindWordStartY(page, "AfterDefaultRule") + fontSize * 0.74;
        double clearance = ruleBottomY - paragraphTopY;

        Assert.True(clearance >= 14, $"Expected default horizontal rule spacing to leave visible rhythm. Clearance: {clearance:0.##}pt.");
        Assert.Contains("0.039 0.078 0.118 RG", rawPdf, StringComparison.Ordinal);
        Assert.Contains("2 w", rawPdf, StringComparison.Ordinal);
        Assert.DoesNotContain("0.784 0.039 0.039 RG", rawPdf, StringComparison.Ordinal);
    }

    [Fact]
    public void RowColumnHorizontalRule_UsesDefaultHorizontalRuleStyleWhenStyleIsNotProvided() {
        var style = new PdfHorizontalRuleStyle {
            Thickness = 1.5,
            Color = PdfColor.FromRgb(10, 20, 30),
            SpacingBefore = 4,
            SpacingAfter = 14
        };

        byte[] bytes = PdfDocument.Create(new PdfOptions {
                PageWidth = 240,
                PageHeight = 180,
                MarginLeft = 20,
                MarginRight = 20,
                MarginTop = 20,
                MarginBottom = 20,
                DefaultFont = PdfStandardFont.Helvetica,
                DefaultFontSize = 10
            })
            .DefaultHorizontalRuleStyle(style)
            .Compose(document =>
                document.Page(page =>
                    page.Content(content =>
                        content.Row(row =>
                            row.Column(100, column => column
                                .HR()
                                .Paragraph(p => p.Text("ColumnDefaultRule")))))))
            .ToBytes();

        string rawPdf = Encoding.ASCII.GetString(bytes);

        Assert.Contains("0.039 0.078 0.118 RG", rawPdf, StringComparison.Ordinal);
        Assert.Contains("1.5 w", rawPdf, StringComparison.Ordinal);
    }

    [Fact]
    public void PdfDocument_ThemeAppliesReusableDefaultsAndSnapshotsInput() {
        var textStyle = new PdfTextStyle {
            Font = PdfStandardFont.Helvetica,
            FontSize = 16,
            Color = PdfColor.FromRgb(10, 20, 30)
        };
        var paragraphStyle = new PdfParagraphStyle {
            FirstLineIndent = 24,
            SpacingAfter = 0
        };
        var tableStyle = TableStyles.Minimal();
        tableStyle.CellPaddingX = 22;
        var theme = new PdfTheme {
            TextStyle = textStyle,
            ParagraphStyle = paragraphStyle,
            TableStyle = tableStyle
        };

        var doc = PdfDocument.Create(new PdfOptions {
                PageWidth = 360,
                PageHeight = 260,
                MarginLeft = 30,
                MarginRight = 30,
                MarginTop = 30,
                MarginBottom = 30
            })
            .Theme(theme)
            .Paragraph(p => p
                .Text("ThemeFirst")
                .LineBreak()
                .Text("ThemeSecond"))
            .Table(new[] {
                new[] { "ThemeTable", "Value" },
                new[] { "Row", "1" }
            });

        textStyle.FontSize = 8;
        textStyle.Color = PdfColor.Black;
        paragraphStyle.FirstLineIndent = 0;
        tableStyle.CellPaddingX = 0;
        theme.TextStyle = new PdfTextStyle {
            Font = PdfStandardFont.Helvetica,
            FontSize = 8,
            Color = PdfColor.Black
        };

        byte[] bytes = doc.ToBytes();

        using var pdf = PdfPigDocument.Open(new MemoryStream(bytes));
        var page = pdf.GetPage(1);
        double pointSize = page.Letters
            .Where(letter => !string.IsNullOrWhiteSpace(letter.Value))
            .Select(letter => letter.PointSize)
            .First();
        double firstX = FindWordStartX(page, "ThemeFirst");
        double secondX = FindWordStartX(page, "ThemeSecond");
        double tableX = FindWordStartX(page, "ThemeTable");
        string rawPdf = Encoding.ASCII.GetString(bytes);

        Assert.InRange(pointSize, 15.5, 16.5);
        Assert.Contains("0.039 0.078 0.118 rg", rawPdf, StringComparison.Ordinal);
        Assert.True(firstX - secondX >= 22, $"Expected theme paragraph style to indent only the first paragraph line. First x: {firstX:0.##}, second x: {secondX:0.##}.");
        Assert.True(tableX - 30 >= 20, $"Expected theme table style padding to affect following tables. Table x: {tableX:0.##}.");
    }

    [Fact]
    public void PdfTheme_WordLikeProvidesGenericDocumentDefaultsAndSnapshotsInput() {
        PdfTheme theme = PdfTheme.WordLike();
        PdfOptions options = new PdfOptions().ApplyTheme(theme);

        theme.TextStyle = new PdfTextStyle {
            Font = PdfStandardFont.Courier,
            FontSize = 7,
            Color = PdfColor.Black
        };
        theme.ParagraphStyle = new PdfParagraphStyle {
            LineHeight = 2,
            SpacingAfter = 0
        };
        theme.TableStyle = TableStyles.Minimal();

        PdfTheme freshTheme = PdfTheme.WordLike();

        Assert.Equal(PdfStandardFont.Helvetica, options.DefaultFont);
        Assert.Equal(11, options.DefaultFontSize);
        Assert.Equal(PdfColor.FromRgb(31, 41, 55), options.DefaultTextColor);
        Assert.Equal(1.15, options.DefaultParagraphStyle!.LineHeight);
        Assert.Equal(8, options.DefaultParagraphStyle.SpacingAfter);
        Assert.True(options.DefaultParagraphStyle.WidowControl);
        Assert.Equal(20, options.DefaultHeadingStyles!.Level1!.FontSize);
        Assert.Equal(16, options.DefaultHeadingStyles.Level2!.FontSize);
        Assert.Equal(13.5, options.DefaultHeadingStyles.Level3!.FontSize);
        Assert.True(options.DefaultHeadingStyles.Level1.KeepWithNext);
        Assert.Equal(18, options.DefaultListStyle!.LeftIndent);
        Assert.Equal(6, options.DefaultListStyle.MarkerGap);
        Assert.Equal(8, options.DefaultListStyle.SpacingAfter);
        Assert.Equal(5, options.DefaultTableStyle!.CellPaddingX);
        Assert.Equal(5, options.DefaultTableStyle.CellPaddingY);
        Assert.Equal(8, options.DefaultPanelStyle!.PaddingX);
        Assert.Equal(8, options.DefaultPanelStyle.SpacingAfter);
        Assert.Equal(0.7, options.DefaultHorizontalRuleStyle!.Thickness);
        Assert.Equal(8, options.DefaultHorizontalRuleStyle.SpacingAfter);
        Assert.Equal(8, options.DefaultImageStyle!.SpacingAfter);
        Assert.Equal(8, options.DefaultDrawingStyle!.SpacingAfter);
        Assert.Equal(18, options.DefaultRowStyle!.Gap);
        Assert.Equal(8, options.DefaultRowStyle.SpacingAfter);
        Assert.Equal(11, freshTheme.TextStyle!.FontSize);
        Assert.Equal(8, freshTheme.ParagraphStyle!.SpacingAfter);
    }

    [Fact]
    public void PdfTheme_BuiltInVisualProfilesExposeReusableDocumentRhythm() {
        PdfOptions technical = new PdfOptions().ApplyTheme(PdfTheme.TechnicalDocument());
        PdfOptions compact = new PdfOptions().ApplyTheme(PdfTheme.Compact());
        PdfOptions report = new PdfOptions().ApplyTheme(PdfTheme.Report());

        Assert.Equal(PdfColor.FromRgb(15, 23, 42), technical.DefaultTableStyle!.HeaderFill);
        Assert.Equal(9.75, technical.DefaultTableStyle.FontSize);
        Assert.Equal(1.2, technical.DefaultTableStyle.LineHeight);
        Assert.True(technical.DefaultTableStyle.AutoFitColumns);
        Assert.Equal(9, technical.DefaultPanelStyle!.SpacingAfter);
        Assert.Equal(0.6, technical.DefaultHorizontalRuleStyle!.Thickness);

        Assert.Equal(10, compact.DefaultFontSize);
        Assert.Equal(1.08, compact.DefaultParagraphStyle!.LineHeight);
        Assert.Equal(4, compact.DefaultParagraphStyle.SpacingAfter);
        Assert.Equal(9, compact.DefaultTableStyle!.FontSize);
        Assert.Equal(14, compact.DefaultRowStyle!.Gap);

        Assert.Equal(PdfColor.FromRgb(30, 64, 175), report.DefaultTableStyle!.HeaderFill);
        Assert.Equal(PdfColor.FromRgb(239, 246, 255), report.DefaultTableStyle.RowStripeFill);
        Assert.Equal(9.25, report.DefaultTableStyle.FontSize);
        Assert.Equal(21, report.DefaultHeadingStyles!.Level1!.FontSize);
        Assert.Equal(PdfColor.FromRgb(30, 64, 175), report.DefaultHeadingStyles.Level2!.Color);
        Assert.Equal(10, report.DefaultPanelStyle!.SpacingAfter);
    }

    [Fact]
    public void PdfDocument_WordLikeThemeRendersReadableMixedFlowRhythm() {
        byte[] bytes = PdfDocument.Create(new PdfOptions {
                PageWidth = 420,
                PageHeight = 520,
                MarginLeft = 42,
                MarginRight = 42,
                MarginTop = 42,
                MarginBottom = 42
            })
            .Theme(PdfTheme.WordLike())
            .H1("WordLikeHeading")
            .Paragraph(p => p.Text("WordLikeBody keeps a comfortable default paragraph rhythm."))
            .Bullets(new[] {
                "WordLikeBulletOne",
                "WordLikeBulletTwo"
            })
            .PanelParagraph(p => p.Text("WordLikePanel"))
            .HR()
            .Table(new[] {
                new[] { "WordLikeTable", "Value" },
                new[] { "Alpha", "42" }
            })
            .ToBytes();

        using var pdf = PdfPigDocument.Open(new MemoryStream(bytes));
        var page = pdf.GetPage(1);
        double bodyPointSize = FindWordPointSize(page, "WordLikeBody");
        double headingY = FindWordStartY(page, "WordLikeHeading");
        double bodyY = FindWordStartY(page, "WordLikeBody");
        double bulletOneY = FindWordStartY(page, "WordLikeBulletOne");
        double bulletTwoY = FindWordStartY(page, "WordLikeBulletTwo");
        double panelY = FindWordStartY(page, "WordLikePanel");
        string rawPdf = Encoding.ASCII.GetString(bytes);

        Assert.InRange(bodyPointSize, 10.5, 11.5);
        Assert.True(headingY - bodyY >= 22, $"Expected Word-like heading/body rhythm. Gap: {headingY - bodyY:0.##}pt.");
        Assert.True(bodyY - bulletOneY >= 20, $"Expected Word-like paragraph/list rhythm. Gap: {bodyY - bulletOneY:0.##}pt.");
        Assert.True(bulletOneY - bulletTwoY >= 12, $"Expected Word-like list item rhythm. Gap: {bulletOneY - bulletTwoY:0.##}pt.");
        Assert.True(bulletTwoY - panelY >= 18, $"Expected Word-like list/panel rhythm. Gap: {bulletTwoY - panelY:0.##}pt.");
        Assert.Contains("0.122 0.161 0.216 rg", rawPdf, StringComparison.Ordinal);
        Assert.Contains("0.067 0.094 0.153 rg", rawPdf, StringComparison.Ordinal);
        AssertNoSameBaselineTextCollisions(page, "Word-like theme flow");
    }

    [Fact]
    public void PdfDocument_ReportThemeRendersStrongerTableAndPanelHierarchy() {
        byte[] bytes = PdfDocument.Create(new PdfOptions {
                PageWidth = 420,
                PageHeight = 520,
                MarginLeft = 42,
                MarginRight = 42,
                MarginTop = 42,
                MarginBottom = 42
            })
            .Theme(PdfTheme.Report())
            .H1("ReportThemeHeading")
            .H2("ReportThemeSection")
            .Paragraph(p => p.Text("ReportThemeBody keeps the body calm while tables carry the hierarchy."))
            .PanelParagraph(p => p.Text("ReportThemePanel"))
            .Table(new[] {
                new[] { "ReportThemeTable", "Value" },
                new[] { "Alpha", "42" },
                new[] { "Beta", "84" }
            })
            .ToBytes();

        using var pdf = PdfPigDocument.Open(new MemoryStream(bytes));
        var page = pdf.GetPage(1);
        string rawPdf = Encoding.ASCII.GetString(bytes);
        double headingY = FindWordStartY(page, "ReportThemeHeading");
        double sectionY = FindWordStartY(page, "ReportThemeSection");
        double bodyY = FindWordStartY(page, "ReportThemeBody");
        double tableY = FindWordStartY(page, "ReportThemeTable");

        Assert.True(headingY - sectionY >= 22, $"Expected report H1/H2 rhythm. Gap: {headingY - sectionY:0.##}pt.");
        Assert.True(sectionY - bodyY >= 18, $"Expected report H2/body rhythm. Gap: {sectionY - bodyY:0.##}pt.");
        Assert.True(bodyY - tableY >= 32, $"Expected report body/table rhythm. Gap: {bodyY - tableY:0.##}pt.");
        Assert.Contains("0.118 0.251 0.686 rg", rawPdf, StringComparison.Ordinal);
        Assert.Contains("0.937 0.965 1 rg", rawPdf, StringComparison.Ordinal);
        AssertNoSameBaselineTextCollisions(page, "report theme flow");
    }

    [Fact]
    public void PdfDocument_OptionsAppliedThemeRendersReusableDefaults() {
        var tableStyle = TableStyles.Minimal();
        tableStyle.CellPaddingX = 22;
        var options = new PdfOptions {
                PageWidth = 360,
                PageHeight = 260,
                MarginLeft = 30,
                MarginRight = 30,
                MarginTop = 30,
                MarginBottom = 30
            }
            .ApplyTheme(new PdfTheme {
                TextStyle = new PdfTextStyle {
                    Font = PdfStandardFont.Helvetica,
                    FontSize = 16,
                    Color = PdfColor.FromRgb(10, 20, 30)
                },
                ParagraphStyle = new PdfParagraphStyle {
                    FirstLineIndent = 24,
                    SpacingAfter = 0
                },
                TableStyle = tableStyle
            });

        byte[] bytes = PdfDocument.Create(options)
            .Paragraph(p => p
                .Text("OptionsThemeFirst")
                .LineBreak()
                .Text("OptionsThemeSecond"))
            .Table(new[] {
                new[] { "OptionsTable", "Value" },
                new[] { "Row", "1" }
            })
            .ToBytes();

        using var pdf = PdfPigDocument.Open(new MemoryStream(bytes));
        var page = pdf.GetPage(1);
        double pointSize = page.Letters
            .Where(letter => !string.IsNullOrWhiteSpace(letter.Value))
            .Select(letter => letter.PointSize)
            .First();
        double firstX = FindWordStartX(page, "OptionsThemeFirst");
        double secondX = FindWordStartX(page, "OptionsThemeSecond");
        double tableX = FindWordStartX(page, "OptionsTable");
        string rawPdf = Encoding.ASCII.GetString(bytes);

        Assert.InRange(pointSize, 15.5, 16.5);
        Assert.Contains("0.039 0.078 0.118 rg", rawPdf, StringComparison.Ordinal);
        Assert.True(firstX - secondX >= 22, $"Expected options theme paragraph style to indent only the first paragraph line. First x: {firstX:0.##}, second x: {secondX:0.##}.");
        Assert.True(tableX - 30 >= 20, $"Expected options theme table style padding to affect following tables. Table x: {tableX:0.##}.");
    }


}
