using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using OfficeIMO.Drawing;
using OfficeIMO.Pdf;
using UglyToad.PdfPig;
using Xunit;

namespace OfficeIMO.Tests.Pdf;

public class PdfDocVisualQualityTests {
    [Fact]
    public void Options_RejectInvalidPageGeometryAndTypography() {
        var widthException = Assert.Throws<ArgumentException>(() =>
            PdfDoc.Create(new PdfOptions {
                    PageWidth = 0
                })
                .Paragraph(p => p.Text("Invalid page width"))
                .ToBytes());

        Assert.Contains("PDF page width must be a positive finite value.", widthException.Message, StringComparison.Ordinal);

        var marginException = Assert.Throws<ArgumentException>(() =>
            PdfDoc.Create(new PdfOptions {
                    MarginLeft = -1
                })
                .Paragraph(p => p.Text("Invalid margin"))
                .ToBytes());

        Assert.Contains("PDF left margin must be a non-negative finite value.", marginException.Message, StringComparison.Ordinal);

        var contentWidthException = Assert.Throws<ArgumentException>(() =>
            PdfDoc.Create(new PdfOptions {
                    PageWidth = 100,
                    MarginLeft = 50,
                    MarginRight = 50
                })
                .Paragraph(p => p.Text("No content width"))
                .ToBytes());

        Assert.Contains("PDF margins must leave a positive content width.", contentWidthException.Message, StringComparison.Ordinal);

        var contentHeightException = Assert.Throws<ArgumentException>(() =>
            PdfDoc.Create(new PdfOptions {
                    PageHeight = 100,
                    MarginTop = 60,
                    MarginBottom = 40
                })
                .Paragraph(p => p.Text("No content height"))
                .ToBytes());

        Assert.Contains("PDF margins must leave a positive content height.", contentHeightException.Message, StringComparison.Ordinal);

        var fontException = Assert.Throws<ArgumentException>(() =>
            PdfDoc.Create(new PdfOptions {
                    DefaultFontSize = double.NaN
                })
                .Paragraph(p => p.Text("Invalid font size"))
                .ToBytes());

        Assert.Contains("PDF default font size must be a positive finite value.", fontException.Message, StringComparison.Ordinal);

        var defaultFontException = Assert.Throws<ArgumentOutOfRangeException>(() =>
            new PdfOptions {
                DefaultFont = (PdfStandardFont)99
            });

        Assert.Equal("DefaultFont", defaultFontException.ParamName);
        Assert.Contains("PDF default font must be one of the supported standard PDF fonts.", defaultFontException.Message, StringComparison.Ordinal);

        var headerException = Assert.Throws<ArgumentException>(() =>
            PdfDoc.Create(new PdfOptions {
                    ShowHeader = true,
                    HeaderFormat = "Header",
                    HeaderFontSize = double.NaN
                })
                .Paragraph(p => p.Text("Invalid header font size"))
                .ToBytes());

        Assert.Contains("PDF header font size must be a positive finite value.", headerException.Message, StringComparison.Ordinal);

        var headerFontException = Assert.Throws<ArgumentOutOfRangeException>(() =>
            new PdfOptions {
                HeaderFont = (PdfStandardFont)99
            });

        Assert.Equal("HeaderFont", headerFontException.ParamName);
        Assert.Contains("PDF header font must be one of the supported standard PDF fonts.", headerFontException.Message, StringComparison.Ordinal);

        var headerOffsetException = Assert.Throws<ArgumentException>(() =>
            PdfDoc.Create(new PdfOptions {
                    ShowHeader = true,
                    HeaderFormat = "Header",
                    MarginTop = 20,
                    HeaderOffsetY = 21
                })
                .Paragraph(p => p.Text("Header above page"))
                .ToBytes());

        Assert.Contains("PDF header offset must not exceed the top margin when header content is enabled.", headerOffsetException.Message, StringComparison.Ordinal);

        var headerFormatException = Assert.Throws<ArgumentException>(() =>
            PdfDoc.Create(new PdfOptions {
                    ShowHeader = true,
                    HeaderFormat = null!
                })
                .Paragraph(p => p.Text("Invalid header format"))
                .ToBytes());

        Assert.Contains("PDF header format cannot be null.", headerFormatException.Message, StringComparison.Ordinal);

        var headerAlignException = Assert.Throws<ArgumentException>(() =>
            PdfDoc.Create(new PdfOptions {
                    ShowHeader = true,
                    HeaderFormat = "Header",
                    HeaderAlign = (PdfAlign)99
                })
                .Paragraph(p => p.Text("Invalid header alignment"))
                .ToBytes());

        Assert.Contains("PDF header alignment must be Left, Center, or Right.", headerAlignException.Message, StringComparison.Ordinal);

        var headerJustifyException = Assert.Throws<ArgumentException>(() =>
            PdfDoc.Create(new PdfOptions {
                    ShowHeader = true,
                    HeaderFormat = "Header",
                    HeaderAlign = PdfAlign.Justify
                })
                .Paragraph(p => p.Text("Unsupported header alignment"))
                .ToBytes());

        Assert.Contains("PDF header alignment must be Left, Center, or Right.", headerJustifyException.Message, StringComparison.Ordinal);

        var footerException = Assert.Throws<ArgumentException>(() =>
            PdfDoc.Create(new PdfOptions {
                    ShowPageNumbers = true,
                    FooterFontSize = double.PositiveInfinity
                })
                .Paragraph(p => p.Text("Invalid footer font size"))
                .ToBytes());

        Assert.Contains("PDF footer font size must be a positive finite value.", footerException.Message, StringComparison.Ordinal);

        var footerFontException = Assert.Throws<ArgumentOutOfRangeException>(() =>
            new PdfOptions {
                FooterFont = (PdfStandardFont)99
            });

        Assert.Equal("FooterFont", footerFontException.ParamName);
        Assert.Contains("PDF footer font must be one of the supported standard PDF fonts.", footerFontException.Message, StringComparison.Ordinal);

        var footerOffsetException = Assert.Throws<ArgumentException>(() =>
            PdfDoc.Create(new PdfOptions {
                    ShowPageNumbers = true,
                    MarginBottom = 20,
                    FooterOffsetY = 21
                })
                .Paragraph(p => p.Text("Footer below page"))
                .ToBytes());

        Assert.Contains("PDF footer offset must not exceed the bottom margin when footer content is enabled.", footerOffsetException.Message, StringComparison.Ordinal);

        var footerAlignException = Assert.Throws<ArgumentException>(() =>
            PdfDoc.Create(new PdfOptions {
                    ShowPageNumbers = true,
                    FooterAlign = (PdfAlign)99
                })
                .Paragraph(p => p.Text("Invalid footer alignment"))
                .ToBytes());

        Assert.Contains("PDF footer alignment must be Left, Center, or Right.", footerAlignException.Message, StringComparison.Ordinal);

        var footerJustifyException = Assert.Throws<ArgumentException>(() =>
            PdfDoc.Create(new PdfOptions {
                    ShowPageNumbers = true,
                    FooterAlign = PdfAlign.Justify
                })
                .Paragraph(p => p.Text("Unsupported footer alignment"))
                .ToBytes());

        Assert.Contains("PDF footer alignment must be Left, Center, or Right.", footerJustifyException.Message, StringComparison.Ordinal);
    }

    [Fact]
    public void ComposePage_RejectsInvalidPageOptions() {
        var exception = Assert.Throws<ArgumentException>(() =>
            PdfDoc.Create()
                .Compose(compose =>
                    compose.Page(page => {
                        page.Size(200, 160);
                        page.Margin(left: 100, top: 20, right: 100, bottom: 20);
                        page.Content(content =>
                            content.Column(column =>
                                column.Item().Paragraph(p => p.Text("No content width"))));
                    }))
                .ToBytes());

        Assert.Contains("PDF margins must leave a positive content width.", exception.Message, StringComparison.Ordinal);
    }

    [Fact]
    public void DefaultOptions_UseProportionalHelveticaForPlainDocuments() {
        byte[] bytes = PdfDoc.Create(new PdfOptions {
                ShowHeader = true,
                HeaderFormat = "HeaderMarker",
                ShowPageNumbers = true,
                FooterFormat = "FooterMarker"
            })
            .Paragraph(p => p.Text("BodyMarker uses the built-in default font."))
            .ToBytes();

        string content = Encoding.ASCII.GetString(bytes);

        Assert.Contains("/BaseFont /Helvetica", content);
        Assert.DoesNotContain("/BaseFont /Courier", content);
    }

    [Fact]
    public void RichText_RejectsNullRunTextBeforeRendering() {
        Assert.Throws<ArgumentNullException>(() =>
            PdfDoc.Create().Paragraph(p => p.Text(null!)));

        Assert.Throws<ArgumentNullException>(() =>
            PdfDoc.Create().Paragraph(p => p.Bold(null!)));

        Assert.Throws<ArgumentNullException>(() =>
            PdfDoc.Create().Paragraph(p => p.Italic(null!)));

        Assert.Throws<ArgumentNullException>(() =>
            PdfDoc.Create().Paragraph(p => p.Underlined(null!)));

        Assert.Throws<ArgumentNullException>(() =>
            PdfDoc.Create().Paragraph(p => p.Strikethrough(null!)));

        Assert.Throws<ArgumentNullException>(() =>
            TextRun.Normal(null!));
    }

    [Fact]
    public void RichText_RejectsInvalidBaselineValuesBeforeRendering() {
        var exception = Assert.Throws<ArgumentException>(() =>
            new TextRun("Invalid baseline", baseline: (PdfTextBaseline)99));

        Assert.Equal("baseline", exception.ParamName);
        Assert.Contains("PDF text baseline must be Normal, Superscript, or Subscript.", exception.Message, StringComparison.Ordinal);

        var builderException = Assert.Throws<ArgumentException>(() =>
            PdfDoc.Create().Paragraph(p => p.Baseline((PdfTextBaseline)99).Text("Invalid baseline")));

        Assert.Equal("baseline", builderException.ParamName);
        Assert.Contains("PDF text baseline must be Normal, Superscript, or Subscript.", builderException.Message, StringComparison.Ordinal);
    }

    [Fact]
    public void RichText_RendersSuperscriptAndSubscriptWithTextRise() {
        byte[] bytes = PdfDoc.Create(new PdfOptions {
                DefaultFontSize = 14
            })
            .Paragraph(p => p
                .Text("Base")
                .Superscript()
                .Link("SUP", "https://example.com/sup")
                .Superscript(false)
                .Text("Mid")
                .Subscript("SUB")
                .Text("End"))
            .ToBytes();

        string content = Encoding.ASCII.GetString(bytes);
        Assert.Contains("/URI (https://example.com/sup)", content);
        Assert.Contains("4.9 Ts", content);
        Assert.Contains("-2.52 Ts", content);
        Assert.Contains("0 Ts", content);
        Assert.Contains("/F1 9.1 Tf", content);

        using var pdf = PdfDocument.Open(new MemoryStream(bytes));
        var letters = pdf.GetPage(1).Letters
            .Where(letter => !string.IsNullOrWhiteSpace(letter.Value))
            .OrderBy(letter => letter.StartBaseLine.X)
            .ToList();

        double baseY = AverageBaselineY(letters, "Base");
        double superY = AverageBaselineY(letters, "SUP");
        double midY = AverageBaselineY(letters, "Mid");
        double subY = AverageBaselineY(letters, "SUB");

        Assert.InRange(Math.Abs(baseY - midY), 0, 0.75);
        Assert.True(superY > baseY + 3.5, $"Expected superscript text to rise above the normal baseline. Base: {baseY:0.##}, super: {superY:0.##}.");
        Assert.True(subY < baseY - 1.5, $"Expected subscript text to sit below the normal baseline. Base: {baseY:0.##}, sub: {subY:0.##}.");
    }

    [Fact]
    public void ParagraphBlocks_SnapshotRunsIntoReadOnlyCollections() {
        var runs = new List<TextRun> {
            TextRun.Normal("Stable alpha"),
            TextRun.Bolded("Stable beta")
        };
        var paragraphStyle = new PdfParagraphStyle {
            LineHeight = 1.6,
            LeftIndent = 4,
            RightIndent = 5,
            FirstLineIndent = 3,
            SpacingBefore = 6,
            SpacingAfter = 7,
            KeepTogether = true,
            KeepWithNext = true,
            WidowControl = true
        };
        var panelStyle = new PanelStyle();

        var paragraph = new RichParagraphBlock(runs, PdfAlign.Left, null, paragraphStyle);
        var panel = new PanelParagraphBlock(runs, PdfAlign.Left, null, panelStyle);

        runs[0] = TextRun.Normal("Mutated alpha");
        runs.Add(TextRun.Normal("Late gamma"));
        paragraphStyle.LineHeight = 2.2;
        paragraphStyle.LeftIndent = 20;
        paragraphStyle.RightIndent = 21;
        paragraphStyle.FirstLineIndent = 22;
        paragraphStyle.SpacingBefore = 22;
        paragraphStyle.SpacingAfter = 23;
        paragraphStyle.KeepTogether = false;
        paragraphStyle.KeepWithNext = false;
        paragraphStyle.WidowControl = false;

        Assert.Equal(new[] { "Stable alpha", "Stable beta" }, paragraph.Runs.Select(run => run.Text).ToArray());
        Assert.Equal(new[] { "Stable alpha", "Stable beta" }, panel.Runs.Select(run => run.Text).ToArray());
        Assert.False(paragraph.Runs is List<TextRun>);
        Assert.False(panel.Runs is List<TextRun>);
        Assert.Equal(1.6, paragraph.Style!.LineHeight);
        Assert.Equal(4, paragraph.Style.LeftIndent);
        Assert.Equal(5, paragraph.Style.RightIndent);
        Assert.Equal(3, paragraph.Style.FirstLineIndent);
        Assert.Equal(6, paragraph.Style.SpacingBefore);
        Assert.Equal(7, paragraph.Style.SpacingAfter);
        Assert.True(paragraph.Style.KeepTogether);
        Assert.True(paragraph.Style.KeepWithNext);
        Assert.True(paragraph.Style.WidowControl);
    }

    [Fact]
    public void ListStyle_ClonePreservesPageFlowSettings() {
        var style = new PdfListStyle {
            KeepTogether = true,
            KeepWithNext = true
        };

        PdfListStyle clone = style.Clone();

        style.KeepTogether = false;
        style.KeepWithNext = false;

        Assert.True(clone.KeepTogether);
        Assert.True(clone.KeepWithNext);
    }

    [Fact]
    public void Options_SnapshotDefaultParagraphStyle() {
        var style = new PdfParagraphStyle {
            LineHeight = 1.7,
            LeftIndent = 8,
            FirstLineIndent = 4,
            SpacingAfter = 9,
            DefaultTabStopWidth = 72,
            WidowControl = true
        };
        var options = new PdfOptions {
            DefaultParagraphStyle = style
        };

        style.LineHeight = 2.1;
        style.LeftIndent = 20;
        style.FirstLineIndent = 21;
        style.SpacingAfter = 22;
        style.DefaultTabStopWidth = 18;
        style.WidowControl = false;

        PdfParagraphStyle readback = options.DefaultParagraphStyle!;
        readback.LeftIndent = 30;

        PdfOptions clone = options.Clone();

        Assert.Equal(1.7, options.DefaultParagraphStyle!.LineHeight);
        Assert.Equal(8, options.DefaultParagraphStyle.LeftIndent);
        Assert.Equal(4, options.DefaultParagraphStyle.FirstLineIndent);
        Assert.Equal(9, options.DefaultParagraphStyle.SpacingAfter);
        Assert.Equal(72, options.DefaultParagraphStyle.DefaultTabStopWidth);
        Assert.True(options.DefaultParagraphStyle.WidowControl);
        Assert.Equal(8, clone.DefaultParagraphStyle!.LeftIndent);
        Assert.Equal(72, clone.DefaultParagraphStyle.DefaultTabStopWidth);
    }

    [Fact]
    public void Options_SnapshotDefaultTableStyle() {
        var style = TableStyles.Minimal();
        style.BorderWidth = 0.8;
        style.CellPaddingX = 12;
        style.RowSeparatorColor = new PdfColor(0.11, 0.22, 0.33);
        style.RowSeparatorWidth = 0.7;
        style.FooterSeparatorColor = new PdfColor(0.21, 0.32, 0.43);
        style.FooterSeparatorWidth = 0.9;
        style.Alignments = new List<PdfColumnAlign> { PdfColumnAlign.Right };
        style.MaxWidth = 180;
        style.LeftIndent = 24;
        style.KeepWithNext = true;
        var options = new PdfOptions {
            DefaultTableStyle = style
        };

        style.BorderWidth = 2;
        style.CellPaddingX = 30;
        style.RowSeparatorColor = PdfColor.Black;
        style.RowSeparatorWidth = 1.5;
        style.FooterSeparatorColor = PdfColor.White;
        style.FooterSeparatorWidth = 2.1;
        style.Alignments[0] = PdfColumnAlign.Left;
        style.MaxWidth = 260;
        style.LeftIndent = 36;
        style.KeepWithNext = false;

        PdfTableStyle readback = options.DefaultTableStyle!;
        readback.CellPaddingX = 44;
        readback.Alignments![0] = PdfColumnAlign.Center;

        PdfOptions clone = options.Clone();

        Assert.Equal(0.8, options.DefaultTableStyle!.BorderWidth);
        Assert.Equal(12, options.DefaultTableStyle.CellPaddingX);
        Assert.Equal(new PdfColor(0.11, 0.22, 0.33), options.DefaultTableStyle.RowSeparatorColor);
        Assert.Equal(0.7, options.DefaultTableStyle.RowSeparatorWidth);
        Assert.Equal(new PdfColor(0.21, 0.32, 0.43), options.DefaultTableStyle.FooterSeparatorColor);
        Assert.Equal(0.9, options.DefaultTableStyle.FooterSeparatorWidth);
        Assert.Equal(PdfColumnAlign.Right, options.DefaultTableStyle.Alignments![0]);
        Assert.Equal(180, options.DefaultTableStyle.MaxWidth);
        Assert.Equal(24, options.DefaultTableStyle.LeftIndent);
        Assert.True(options.DefaultTableStyle.KeepWithNext);
        Assert.Equal(12, clone.DefaultTableStyle!.CellPaddingX);
        Assert.Equal(new PdfColor(0.21, 0.32, 0.43), clone.DefaultTableStyle.FooterSeparatorColor);
        Assert.Equal(0.9, clone.DefaultTableStyle.FooterSeparatorWidth);
        Assert.Equal(PdfColumnAlign.Right, clone.DefaultTableStyle.Alignments![0]);
        Assert.Equal(180, clone.DefaultTableStyle.MaxWidth);
        Assert.Equal(24, clone.DefaultTableStyle.LeftIndent);
        Assert.True(clone.DefaultTableStyle.KeepWithNext);
    }

    [Fact]
    public void Options_SnapshotDefaultHeadingStyles() {
        var style = new PdfHeadingStyle {
            FontSize = 16,
            LineHeight = 1.1,
            SpacingBefore = 4,
            SpacingAfter = 12,
            Color = PdfColor.FromRgb(10, 20, 30),
            KeepWithNext = false
        };
        var styles = new PdfHeadingStyles {
            Level1 = style
        };
        var options = new PdfOptions {
            DefaultHeadingStyles = styles
        };

        style.FontSize = 30;
        style.SpacingAfter = 2;
        styles.Level1 = new PdfHeadingStyle {
            FontSize = 9
        };

        PdfHeadingStyles readback = options.DefaultHeadingStyles!;
        readback.Level1 = new PdfHeadingStyle {
            FontSize = 8
        };

        PdfOptions clone = options.Clone();

        Assert.Equal(16, options.DefaultHeadingStyles!.Level1!.FontSize);
        Assert.Equal(1.1, options.DefaultHeadingStyles.Level1.LineHeight);
        Assert.Equal(4, options.DefaultHeadingStyles.Level1.SpacingBefore);
        Assert.Equal(12, options.DefaultHeadingStyles.Level1.SpacingAfter);
        Assert.Equal(PdfColor.FromRgb(10, 20, 30), options.DefaultHeadingStyles.Level1.Color);
        Assert.False(options.DefaultHeadingStyles.Level1.KeepWithNext);
        Assert.Equal(16, clone.DefaultHeadingStyles!.Level1!.FontSize);
    }

    [Fact]
    public void Options_SnapshotDefaultPanelStyle() {
        var style = new PanelStyle {
            Background = PdfColor.FromRgb(240, 248, 255),
            BorderColor = PdfColor.FromRgb(20, 40, 60),
            BorderWidth = 1.2,
            PaddingX = 12,
            PaddingY = 8,
            MaxWidth = 180,
            Align = PdfAlign.Center,
            SpacingBefore = 5,
            SpacingAfter = 13,
            KeepTogether = true,
            KeepWithNext = true
        };
        var options = new PdfOptions {
            DefaultPanelStyle = style
        };

        style.PaddingX = 30;
        style.Align = PdfAlign.Right;
        style.Background = PdfColor.Black;
        style.KeepWithNext = false;

        PanelStyle readback = options.DefaultPanelStyle!;
        readback.PaddingX = 44;

        PdfOptions clone = options.Clone();

        Assert.Equal(PdfColor.FromRgb(240, 248, 255), options.DefaultPanelStyle!.Background);
        Assert.Equal(PdfColor.FromRgb(20, 40, 60), options.DefaultPanelStyle.BorderColor);
        Assert.Equal(1.2, options.DefaultPanelStyle.BorderWidth);
        Assert.Equal(12, options.DefaultPanelStyle.PaddingX);
        Assert.Equal(8, options.DefaultPanelStyle.PaddingY);
        Assert.Equal(180, options.DefaultPanelStyle.MaxWidth);
        Assert.Equal(PdfAlign.Center, options.DefaultPanelStyle.Align);
        Assert.Equal(5, options.DefaultPanelStyle.SpacingBefore);
        Assert.Equal(13, options.DefaultPanelStyle.SpacingAfter);
        Assert.True(options.DefaultPanelStyle.KeepTogether);
        Assert.True(options.DefaultPanelStyle.KeepWithNext);
        Assert.Equal(12, clone.DefaultPanelStyle!.PaddingX);
        Assert.True(clone.DefaultPanelStyle.KeepWithNext);
    }

    [Fact]
    public void Options_SnapshotDefaultHorizontalRuleStyle() {
        var style = new PdfHorizontalRuleStyle {
            Thickness = 1.4,
            Color = PdfColor.FromRgb(20, 40, 60),
            SpacingBefore = 5,
            SpacingAfter = 13,
            KeepWithNext = true
        };
        var options = new PdfOptions {
            DefaultHorizontalRuleStyle = style
        };

        style.Thickness = 3;
        style.Color = PdfColor.Black;
        style.SpacingBefore = 1;
        style.SpacingAfter = 2;
        style.KeepWithNext = false;

        PdfHorizontalRuleStyle readback = options.DefaultHorizontalRuleStyle!;
        readback.Thickness = 4;
        readback.Color = PdfColor.FromRgb(200, 10, 10);

        PdfOptions clone = options.Clone();

        Assert.Equal(1.4, options.DefaultHorizontalRuleStyle!.Thickness);
        Assert.Equal(PdfColor.FromRgb(20, 40, 60), options.DefaultHorizontalRuleStyle.Color);
        Assert.Equal(5, options.DefaultHorizontalRuleStyle.SpacingBefore);
        Assert.Equal(13, options.DefaultHorizontalRuleStyle.SpacingAfter);
        Assert.True(options.DefaultHorizontalRuleStyle.KeepWithNext);
        Assert.Equal(1.4, clone.DefaultHorizontalRuleStyle!.Thickness);
        Assert.Equal(PdfColor.FromRgb(20, 40, 60), clone.DefaultHorizontalRuleStyle.Color);
        Assert.True(clone.DefaultHorizontalRuleStyle.KeepWithNext);
    }

    [Fact]
    public void Options_ApplyThemeSnapshotsDefaultStyles() {
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
        tableStyle.KeepWithNext = true;
        var headingStyle = new PdfHeadingStyle {
            FontSize = 19,
            SpacingAfter = 14,
            Color = PdfColor.FromRgb(40, 50, 60)
        };
        var listStyle = new PdfListStyle {
            FontSize = 13,
            LeftIndent = 18,
            Color = PdfColor.FromRgb(70, 80, 90),
            KeepTogether = true,
            KeepWithNext = true
        };
        var panelStyle = new PanelStyle {
            PaddingX = 16,
            MaxWidth = 190,
            Background = PdfColor.FromRgb(240, 248, 255),
            KeepWithNext = true
        };
        var horizontalRuleStyle = new PdfHorizontalRuleStyle {
            Thickness = 1.4,
            Color = PdfColor.FromRgb(80, 90, 100),
            SpacingAfter = 17,
            KeepWithNext = true
        };
        var rowStyle = new PdfRowStyle {
            Gap = 21,
            SpacingBefore = 5,
            SpacingAfter = 6,
            KeepTogether = true,
            KeepWithNext = true
        };
        var theme = new PdfTheme {
            TextStyle = textStyle,
            ParagraphStyle = paragraphStyle,
            TableStyle = tableStyle,
            HeadingStyles = new PdfHeadingStyles {
                Level1 = headingStyle
            },
            ListStyle = listStyle,
            PanelStyle = panelStyle,
            HorizontalRuleStyle = horizontalRuleStyle,
            RowStyle = rowStyle
        };
        var options = new PdfOptions().ApplyTheme(theme);

        textStyle.FontSize = 8;
        textStyle.Color = PdfColor.Black;
        paragraphStyle.FirstLineIndent = 0;
        tableStyle.CellPaddingX = 0;
        tableStyle.KeepWithNext = false;
        headingStyle.FontSize = 8;
        listStyle.FontSize = 8;
        listStyle.KeepTogether = false;
        listStyle.KeepWithNext = false;
        panelStyle.PaddingX = 4;
        panelStyle.KeepWithNext = false;
        horizontalRuleStyle.Thickness = 5;
        horizontalRuleStyle.Color = PdfColor.Black;
        horizontalRuleStyle.KeepWithNext = false;
        rowStyle.Gap = 3;
        rowStyle.SpacingBefore = 1;
        rowStyle.SpacingAfter = 2;
        rowStyle.KeepTogether = false;
        rowStyle.KeepWithNext = false;
        theme.TextStyle = new PdfTextStyle {
            Font = PdfStandardFont.Helvetica,
            FontSize = 7
        };

        PdfOptions clone = options.Clone();

        Assert.Equal(PdfStandardFont.Helvetica, options.DefaultFont);
        Assert.Equal(16, options.DefaultFontSize);
        Assert.Equal(PdfColor.FromRgb(10, 20, 30), options.DefaultTextColor);
        Assert.Equal(24, options.DefaultParagraphStyle!.FirstLineIndent);
        Assert.Equal(22, options.DefaultTableStyle!.CellPaddingX);
        Assert.True(options.DefaultTableStyle.KeepWithNext);
        Assert.Equal(19, options.DefaultHeadingStyles!.Level1!.FontSize);
        Assert.Equal(14, options.DefaultHeadingStyles.Level1.SpacingAfter);
        Assert.Equal(13, options.DefaultListStyle!.FontSize);
        Assert.Equal(18, options.DefaultListStyle.LeftIndent);
        Assert.True(options.DefaultListStyle.KeepTogether);
        Assert.True(options.DefaultListStyle.KeepWithNext);
        Assert.Equal(16, options.DefaultPanelStyle!.PaddingX);
        Assert.Equal(190, options.DefaultPanelStyle.MaxWidth);
        Assert.True(options.DefaultPanelStyle.KeepWithNext);
        Assert.Equal(1.4, options.DefaultHorizontalRuleStyle!.Thickness);
        Assert.Equal(PdfColor.FromRgb(80, 90, 100), options.DefaultHorizontalRuleStyle.Color);
        Assert.Equal(17, options.DefaultHorizontalRuleStyle.SpacingAfter);
        Assert.True(options.DefaultHorizontalRuleStyle.KeepWithNext);
        Assert.Equal(21, options.DefaultRowStyle!.Gap);
        Assert.Equal(5, options.DefaultRowStyle.SpacingBefore);
        Assert.Equal(6, options.DefaultRowStyle.SpacingAfter);
        Assert.True(options.DefaultRowStyle.KeepTogether);
        Assert.True(options.DefaultRowStyle.KeepWithNext);
        Assert.Equal(16, clone.DefaultFontSize);
        Assert.Equal(24, clone.DefaultParagraphStyle!.FirstLineIndent);
        Assert.Equal(22, clone.DefaultTableStyle!.CellPaddingX);
        Assert.True(clone.DefaultTableStyle.KeepWithNext);
        Assert.Equal(19, clone.DefaultHeadingStyles!.Level1!.FontSize);
        Assert.Equal(13, clone.DefaultListStyle!.FontSize);
        Assert.True(clone.DefaultListStyle.KeepTogether);
        Assert.True(clone.DefaultListStyle.KeepWithNext);
        Assert.Equal(16, clone.DefaultPanelStyle!.PaddingX);
        Assert.True(clone.DefaultPanelStyle.KeepWithNext);
        Assert.Equal(1.4, clone.DefaultHorizontalRuleStyle!.Thickness);
        Assert.True(clone.DefaultHorizontalRuleStyle.KeepWithNext);
        Assert.Equal(21, clone.DefaultRowStyle!.Gap);
        Assert.True(clone.DefaultRowStyle.KeepTogether);
        Assert.True(clone.DefaultRowStyle.KeepWithNext);
    }

    [Fact]
    public void PdfDoc_DefaultTextStyleAppliesToFollowingContent() {
        byte[] bytes = PdfDoc.Create(new PdfOptions {
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

        using var pdf = PdfDocument.Open(new MemoryStream(bytes));
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
    public void PdfDoc_DefaultTextStyleObjectAppliesToFollowingContentAndSnapshotsInput() {
        var style = new PdfTextStyle {
            Font = PdfStandardFont.Helvetica,
            FontSize = 16,
            Color = PdfColor.FromRgb(10, 20, 30)
        };

        byte[] bytes = PdfDoc.Create(new PdfOptions {
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

        using var pdf = PdfDocument.Open(new MemoryStream(bytes));
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
    public void PdfDoc_DefaultHeadingStyleAppliesToFollowingHeadingsAndSnapshotsInput() {
        var style = new PdfHeadingStyle {
            FontSize = 12,
            LineHeight = 1,
            SpacingAfter = 24,
            Color = PdfColor.FromRgb(10, 20, 30)
        };

        byte[] bytes = PdfDoc.Create(new PdfOptions {
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

        using var pdf = PdfDocument.Open(new MemoryStream(bytes));
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

        byte[] defaultBytes = PdfDoc.Create(options)
            .Paragraph(p => p.Text("BeforeMarker"), style: new PdfParagraphStyle { SpacingAfter = 0 })
            .H2("HeadingMarker", style: defaultStyle)
            .Paragraph(p => p.Text("AfterMarker"))
            .ToBytes();
        byte[] spacedBytes = PdfDoc.Create(options)
            .Paragraph(p => p.Text("BeforeMarker"), style: new PdfParagraphStyle { SpacingAfter = 0 })
            .H2("HeadingMarker", style: spacedStyle)
            .Paragraph(p => p.Text("AfterMarker"))
            .ToBytes();

        using var defaultPdf = PdfDocument.Open(new MemoryStream(defaultBytes));
        using var spacedPdf = PdfDocument.Open(new MemoryStream(spacedBytes));
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

        byte[] defaultBytes = PdfDoc.Create(options)
            .H2("TopHeadingMarker", style: defaultStyle)
            .ToBytes();
        byte[] spacedBytes = PdfDoc.Create(options)
            .H2("TopHeadingMarker", style: spacedStyle)
            .ToBytes();

        using var defaultPdf = PdfDocument.Open(new MemoryStream(defaultBytes));
        using var spacedPdf = PdfDocument.Open(new MemoryStream(spacedBytes));

        double defaultTopY = FindWordStartY(defaultPdf.GetPage(1), "TopHeadingMarker");
        double spacedTopY = FindWordStartY(spacedPdf.GetPage(1), "TopHeadingMarker");

        Assert.InRange(Math.Abs(defaultTopY - spacedTopY), 0, 1.5);
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

        byte[] defaultBytes = PdfDoc.Create(options)
            .Compose(document => document.Page(page => page.Content(content => content.Row(row => row.Column(100, column => column
                .H2("ColumnHeadingMarker", style: defaultStyle))))))
            .ToBytes();
        byte[] spacedBytes = PdfDoc.Create(options)
            .Compose(document => document.Page(page => page.Content(content => content.Row(row => row.Column(100, column => column
                .H2("ColumnHeadingMarker", style: spacedStyle))))))
            .ToBytes();

        using var defaultPdf = PdfDocument.Open(new MemoryStream(defaultBytes));
        using var spacedPdf = PdfDocument.Open(new MemoryStream(spacedBytes));

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

        using var defaultPdf = PdfDocument.Open(new MemoryStream(defaultBytes));
        using var spacedPdf = PdfDocument.Open(new MemoryStream(spacedBytes));

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

        using var defaultPdf = PdfDocument.Open(new MemoryStream(defaultBytes));
        using var spacedPdf = PdfDocument.Open(new MemoryStream(spacedBytes));

        double defaultTopY = FindWordStartY(defaultPdf.GetPage(1), marker);
        double spacedTopY = FindWordStartY(spacedPdf.GetPage(1), marker);

        Assert.InRange(Math.Abs(defaultTopY - spacedTopY), 0, 1.5);
    }

    [Fact]
    public void PdfDoc_DefaultPanelStyleAppliesToFollowingPanelsAndSnapshotsInput() {
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

        byte[] bytes = PdfDoc.Create(options)
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

        using var pdf = PdfDocument.Open(new MemoryStream(bytes));
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
    public void PdfDoc_DefaultHorizontalRuleStyleAppliesToFollowingRulesAndSnapshotsInput() {
        const double fontSize = 10;
        var style = new PdfHorizontalRuleStyle {
            Thickness = 2,
            Color = PdfColor.FromRgb(10, 20, 30),
            SpacingBefore = 3,
            SpacingAfter = 15
        };

        byte[] bytes = PdfDoc.Create(new PdfOptions {
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

        using var pdf = PdfDocument.Open(new MemoryStream(bytes));
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

        byte[] bytes = PdfDoc.Create(new PdfOptions {
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
    public void PdfDoc_ThemeAppliesReusableDefaultsAndSnapshotsInput() {
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

        var doc = PdfDoc.Create(new PdfOptions {
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

        using var pdf = PdfDocument.Open(new MemoryStream(bytes));
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
    public void PdfDoc_WordLikeThemeRendersReadableMixedFlowRhythm() {
        byte[] bytes = PdfDoc.Create(new PdfOptions {
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

        using var pdf = PdfDocument.Open(new MemoryStream(bytes));
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
    public void PdfDoc_OptionsAppliedThemeRendersReusableDefaults() {
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

        byte[] bytes = PdfDoc.Create(options)
            .Paragraph(p => p
                .Text("OptionsThemeFirst")
                .LineBreak()
                .Text("OptionsThemeSecond"))
            .Table(new[] {
                new[] { "OptionsTable", "Value" },
                new[] { "Row", "1" }
            })
            .ToBytes();

        using var pdf = PdfDocument.Open(new MemoryStream(bytes));
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

    [Fact]
    public void PdfDoc_DefaultTextStyleRejectsInvalidInputs() {
        Assert.Throws<ArgumentNullException>(() => PdfDoc.Create().DefaultTextStyle((Action<PdfTextStyleCompose>)null!));
        Assert.Throws<ArgumentNullException>(() => PdfDoc.Create().DefaultTextStyle((PdfTextStyle)null!));
        Assert.Throws<ArgumentNullException>(() => PdfDoc.Create().Theme(null!));
        Assert.Throws<ArgumentNullException>(() => new PdfOptions().ApplyTheme(null!));

        var fontException = Assert.Throws<ArgumentOutOfRangeException>(() =>
            PdfDoc.Create().DefaultTextStyle(style => style.Font((PdfStandardFont)99)));

        Assert.Equal("font", fontException.ParamName);
        Assert.Contains("PDF default font must be one of the supported standard PDF fonts.", fontException.Message, StringComparison.Ordinal);

        var fontSizeException = Assert.Throws<ArgumentOutOfRangeException>(() =>
            PdfDoc.Create().DefaultTextStyle(style => style.FontSize(double.NaN)));

        Assert.Equal("size", fontSizeException.ParamName);

        var textStyleFontException = Assert.Throws<ArgumentOutOfRangeException>(() =>
            new PdfTextStyle { Font = (PdfStandardFont)99 });

        Assert.Equal("Font", textStyleFontException.ParamName);
        Assert.Contains("PDF text style font must be one of the supported standard PDF fonts.", textStyleFontException.Message, StringComparison.Ordinal);

        var textStyleFontSizeException = Assert.Throws<ArgumentOutOfRangeException>(() =>
            new PdfTextStyle { FontSize = 0 });

        Assert.Equal("FontSize", textStyleFontSizeException.ParamName);

        Assert.Throws<ArgumentNullException>(() => PdfDoc.Create().DefaultHeadingStyle(1, null!));
        Assert.Throws<ArgumentOutOfRangeException>(() => PdfDoc.Create().DefaultHeadingStyle(4, new PdfHeadingStyle()));
        Assert.Throws<ArgumentNullException>(() => PdfDoc.Create().DefaultPanelStyle(null!));
        Assert.Throws<ArgumentNullException>(() => PdfDoc.Create().DefaultHorizontalRuleStyle(null!));
        Assert.Throws<ArgumentNullException>(() => PdfDoc.Create().DefaultDrawingStyle(null!));

        var headingSizeException = Assert.Throws<ArgumentException>(() =>
            new PdfHeadingStyle { FontSize = 0 });

        Assert.Equal("FontSize", headingSizeException.ParamName);

        var headingSpacingException = Assert.Throws<ArgumentException>(() =>
            new PdfHeadingStyle { SpacingAfter = -1 });

        Assert.Equal("SpacingAfter", headingSpacingException.ParamName);
    }

    [Fact]
    public void Heading_RejectsUnsupportedAlignmentBeforeRendering() {
        var invalidAlignException = Assert.Throws<ArgumentException>(() =>
            PdfDoc.Create().H1("Invalid heading", (PdfAlign)99));

        Assert.Contains("Heading alignment must be Left, Center, or Right.", invalidAlignException.Message, StringComparison.Ordinal);

        var justifyException = Assert.Throws<ArgumentException>(() =>
            PdfDoc.Create().H2("Unsupported heading", PdfAlign.Justify));

        Assert.Contains("Heading alignment must be Left, Center, or Right.", justifyException.Message, StringComparison.Ordinal);
    }

    [Fact]
    public void Lists_RejectUnsupportedAlignmentBeforeRendering() {
        var invalidBulletAlignException = Assert.Throws<ArgumentException>(() =>
            PdfDoc.Create().Bullets(new[] { "Invalid bullet" }, (PdfAlign)99));

        Assert.Contains("Bullet list alignment must be Left, Center, or Right.", invalidBulletAlignException.Message, StringComparison.Ordinal);

        var bulletJustifyException = Assert.Throws<ArgumentException>(() =>
            PdfDoc.Create().Bullets(new[] { "Unsupported bullet" }, PdfAlign.Justify));

        Assert.Contains("Bullet list alignment must be Left, Center, or Right.", bulletJustifyException.Message, StringComparison.Ordinal);

        var invalidNumberedAlignException = Assert.Throws<ArgumentException>(() =>
            PdfDoc.Create().Numbered(new[] { "Invalid numbered" }, (PdfAlign)99));

        Assert.Contains("Numbered list alignment must be Left, Center, or Right.", invalidNumberedAlignException.Message, StringComparison.Ordinal);

        var numberedJustifyException = Assert.Throws<ArgumentException>(() =>
            PdfDoc.Create().Numbered(new[] { "Unsupported numbered" }, PdfAlign.Justify));

        Assert.Contains("Numbered list alignment must be Left, Center, or Right.", numberedJustifyException.Message, StringComparison.Ordinal);
    }

    [Fact]
    public void ImageShapeAndDrawingBlocks_RejectUnsupportedAlignmentBeforeRendering() {
        byte[] png = CreateMinimalRgbPng();

        var invalidImageAlignException = Assert.Throws<ArgumentException>(() =>
            PdfDoc.Create().Image(png, 24, 24, (PdfAlign)99));

        Assert.Contains("Image alignment must be Left, Center, or Right.", invalidImageAlignException.Message, StringComparison.Ordinal);

        var imageJustifyException = Assert.Throws<ArgumentException>(() =>
            PdfDoc.Create().Image(png, 24, 24, PdfAlign.Justify));

        Assert.Contains("Image alignment must be Left, Center, or Right.", imageJustifyException.Message, StringComparison.Ordinal);

        var shape = OfficeShape.Rectangle(24, 12);

        var invalidShapeAlignException = Assert.Throws<ArgumentException>(() =>
            PdfDoc.Create().Shape(shape, (PdfAlign)99));

        Assert.Contains("Shape alignment must be Left, Center, or Right.", invalidShapeAlignException.Message, StringComparison.Ordinal);

        var shapeJustifyException = Assert.Throws<ArgumentException>(() =>
            PdfDoc.Create().Shape(shape, PdfAlign.Justify));

        Assert.Contains("Shape alignment must be Left, Center, or Right.", shapeJustifyException.Message, StringComparison.Ordinal);

        var drawing = new OfficeDrawing(24, 12)
            .AddShape(OfficeShape.Rectangle(24, 12), 0, 0);

        var invalidDrawingAlignException = Assert.Throws<ArgumentException>(() =>
            PdfDoc.Create().Drawing(drawing, (PdfAlign)99));

        Assert.Contains("Drawing alignment must be Left, Center, or Right.", invalidDrawingAlignException.Message, StringComparison.Ordinal);

        var drawingJustifyException = Assert.Throws<ArgumentException>(() =>
            PdfDoc.Create().Drawing(drawing, PdfAlign.Justify));

        Assert.Contains("Drawing alignment must be Left, Center, or Right.", drawingJustifyException.Message, StringComparison.Ordinal);
    }

    [Fact]
    public void ParagraphAndPanelBlocks_RejectInvalidAlignmentModelState() {
        var paragraphAlignException = Assert.Throws<ArgumentException>(() =>
            PdfDoc.Create().Paragraph(p => p.Text("Invalid paragraph"), (PdfAlign)99));

        Assert.Contains("Paragraph alignment must be Left, Center, Right, or Justify.", paragraphAlignException.Message, StringComparison.Ordinal);

        byte[] justifiedParagraph = PdfDoc.Create()
            .Paragraph(p => p.Text("Justified paragraph alignment remains supported for report text."), PdfAlign.Justify)
            .ToBytes();

        Assert.NotEmpty(justifiedParagraph);

        var panelParagraphAlignException = Assert.Throws<ArgumentException>(() =>
            PdfDoc.Create().PanelParagraph(p => p.Text("Invalid panel text"), new PanelStyle(), (PdfAlign)99));

        Assert.Contains("Panel paragraph alignment must be Left, Center, Right, or Justify.", panelParagraphAlignException.Message, StringComparison.Ordinal);

        byte[] justifiedPanelParagraph = PdfDoc.Create()
            .PanelParagraph(p => p.Text("Justified panel text remains supported inside the panel box."), new PanelStyle(), PdfAlign.Justify)
            .ToBytes();

        Assert.NotEmpty(justifiedPanelParagraph);

        var invalidPanelBoxAlignException = Assert.Throws<ArgumentException>(() =>
            PdfDoc.Create()
                .PanelParagraph(p => p.Text("Invalid panel box"), new PanelStyle {
                    MaxWidth = 120,
                    Align = (PdfAlign)99
                })
                .ToBytes());

        Assert.Contains("Panel box alignment must be Left, Center, or Right.", invalidPanelBoxAlignException.Message, StringComparison.Ordinal);

        var panelBoxJustifyException = Assert.Throws<ArgumentException>(() =>
            PdfDoc.Create()
                .PanelParagraph(p => p.Text("Unsupported panel box"), new PanelStyle {
                    MaxWidth = 120,
                    Align = PdfAlign.Justify
                })
                .ToBytes());

        Assert.Contains("Panel box alignment must be Left, Center, or Right.", panelBoxJustifyException.Message, StringComparison.Ordinal);
    }

    [Fact]
    public void MutableAlignmentProperties_RejectUnsupportedValuesOnAssignment() {
        var headerAlignException = Assert.Throws<ArgumentException>(() =>
            new PdfOptions {
                HeaderAlign = (PdfAlign)99
            });

        Assert.Contains("PDF header alignment must be Left, Center, or Right.", headerAlignException.Message, StringComparison.Ordinal);

        var headerJustifyException = Assert.Throws<ArgumentException>(() =>
            new PdfOptions {
                HeaderAlign = PdfAlign.Justify
            });

        Assert.Contains("PDF header alignment must be Left, Center, or Right.", headerJustifyException.Message, StringComparison.Ordinal);

        var footerAlignException = Assert.Throws<ArgumentException>(() =>
            new PdfOptions {
                FooterAlign = (PdfAlign)99
            });

        Assert.Contains("PDF footer alignment must be Left, Center, or Right.", footerAlignException.Message, StringComparison.Ordinal);

        var footerJustifyException = Assert.Throws<ArgumentException>(() =>
            new PdfOptions {
                FooterAlign = PdfAlign.Justify
            });

        Assert.Contains("PDF footer alignment must be Left, Center, or Right.", footerJustifyException.Message, StringComparison.Ordinal);

        var panelAlignException = Assert.Throws<ArgumentException>(() =>
            new PanelStyle {
                Align = (PdfAlign)99
            });

        Assert.Contains("Panel box alignment must be Left, Center, or Right.", panelAlignException.Message, StringComparison.Ordinal);

        var panelJustifyException = Assert.Throws<ArgumentException>(() =>
            new PanelStyle {
                Align = PdfAlign.Justify
            });

        Assert.Contains("Panel box alignment must be Left, Center, or Right.", panelJustifyException.Message, StringComparison.Ordinal);

        var captionStyle = TableStyles.Minimal();
        var captionAlignException = Assert.Throws<ArgumentException>(() =>
            captionStyle.CaptionAlign = (PdfAlign)99);

        Assert.Contains("Table caption alignment must be Left, Center, or Right.", captionAlignException.Message, StringComparison.Ordinal);

        var captionJustifyException = Assert.Throws<ArgumentException>(() =>
            captionStyle.CaptionAlign = PdfAlign.Justify);

        Assert.Contains("Table caption alignment must be Left, Center, or Right.", captionJustifyException.Message, StringComparison.Ordinal);
    }

    [Fact]
    public void LinkAnnotations_RejectInvalidUriModelStateBeforeRendering() {
        Assert.Throws<ArgumentNullException>(() =>
            PdfDoc.Create().Paragraph(p => p.Link("OfficeIMO", null!)));

        Assert.Throws<ArgumentException>(() =>
            PdfDoc.Create().Paragraph(p => p.Link("OfficeIMO", "relative/link")));

        Assert.Throws<ArgumentException>(() =>
            PdfDoc.Create().Paragraph(p => p.Link("", "https://evotec.xyz")));

        Assert.Throws<ArgumentException>(() =>
            PdfDoc.Create().Paragraph(p => p.Link("OfficeIMO", "https://evotec.xyz", contents: " ")));

        Assert.Throws<ArgumentException>(() =>
            PdfDoc.Create().H1("Linked heading", linkUri: "bookmark-only"));

        Assert.Throws<ArgumentException>(() =>
            PdfDoc.Create().H1("Linked heading", linkUri: "https://evotec.xyz", linkContents: " "));

        Assert.Throws<ArgumentException>(() =>
            PdfDoc.Create().H1("Plain heading", linkContents: "metadata without link"));

        byte[] png = CreateMinimalRgbPng();
        Assert.Throws<ArgumentException>(() =>
            PdfDoc.Create().Image(png, 24, 24, linkUri: "not-a-uri"));

        Assert.Throws<ArgumentException>(() =>
            PdfDoc.Create().Image(png, 24, 24, linkUri: "https://evotec.xyz", linkContents: " "));

        Assert.Throws<ArgumentException>(() =>
            PdfDoc.Create().Image(png, 24, 24, linkContents: "metadata without link"));

        var shape = OfficeShape.Rectangle(24, 12);
        Assert.Throws<ArgumentException>(() =>
            PdfDoc.Create().Shape(shape, linkUri: "not-a-uri"));

        Assert.Throws<ArgumentException>(() =>
            PdfDoc.Create().Shape(shape, linkUri: "https://evotec.xyz", linkContents: " "));

        Assert.Throws<ArgumentException>(() =>
            PdfDoc.Create().Shape(shape, linkContents: "metadata without link"));

        Assert.Throws<ArgumentException>(() =>
            PdfDoc.Create().Rectangle(24, 12, linkUri: "not-a-uri"));

        Assert.Throws<ArgumentException>(() =>
            PdfDoc.Create().Rectangle(24, 12, linkUri: "https://evotec.xyz", linkContents: " "));

        Assert.Throws<ArgumentException>(() =>
            PdfDoc.Create().Rectangle(24, 12, linkContents: "metadata without link"));

        var drawing = new OfficeDrawing(24, 12)
            .AddShape(OfficeShape.Rectangle(24, 12), 0, 0);
        Assert.Throws<ArgumentException>(() =>
            PdfDoc.Create().Drawing(drawing, linkUri: "not-a-uri"));

        Assert.Throws<ArgumentException>(() =>
            PdfDoc.Create().Drawing(drawing, linkUri: "https://evotec.xyz", linkContents: " "));

        Assert.Throws<ArgumentException>(() =>
            PdfDoc.Create().Drawing(drawing, linkContents: "metadata without link"));

        Assert.Throws<ArgumentException>(() =>
            PdfDoc.Create().TableWithLinks(
                new[] { new[] { "Name" }, new[] { "OfficeIMO" } },
                new Dictionary<(int Row, int Col), string> {
                    [(1, 0)] = "not-a-uri"
                }));

        Assert.Throws<ArgumentOutOfRangeException>(() =>
            PdfDoc.Create().TableWithLinks(
                new[] { new[] { "Name" }, new[] { "OfficeIMO" } },
                new Dictionary<(int Row, int Col), string> {
                    [(-1, 0)] = "https://evotec.xyz"
                }));

        Assert.Throws<ArgumentOutOfRangeException>(() =>
            PdfDoc.Create().TableWithLinks(
                new[] { new[] { "Name" }, new[] { "OfficeIMO" } },
                new Dictionary<(int Row, int Col), string> {
                    [(2, 0)] = "https://evotec.xyz"
                }));

        Assert.Throws<ArgumentOutOfRangeException>(() =>
            PdfDoc.Create().TableWithLinks(
                new[] { new[] { "Name", "Url" }, new[] { "OfficeIMO" } },
                new Dictionary<(int Row, int Col), string> {
                    [(1, 1)] = "https://evotec.xyz"
                }));
    }

    [Fact]
    public void Table_NormalizesCellsAndSnapshotsInputRowsBeforeRendering() {
        var body = new[] { "Original", (string)null! };
        var rows = new[] {
            new[] { "Name", "Value" },
            body
        };

        var doc = PdfDoc.Create()
            .Table(rows);

        body[0] = "Mutated";
        body[1] = "AlsoMutated";

        byte[] bytes = doc.ToBytes();

        using var pdf = PdfDocument.Open(new MemoryStream(bytes));
        string text = string.Concat(pdf.GetPage(1).Letters.Select(letter => letter.Value));

        Assert.Contains("Original", text, StringComparison.Ordinal);
        Assert.DoesNotContain("Mutated", text, StringComparison.Ordinal);
        Assert.DoesNotContain("AlsoMutated", text, StringComparison.Ordinal);
    }

    [Fact]
    public void TableBlock_SnapshotsRowsStyleAndLinksIntoReadOnlyModel() {
        var body = new[] { "Original", (string)null! };
        var style = TableStyles.Minimal();
        style.BorderWidth = 2;
        style.CellPaddingX = 7;
        style.RowSeparatorColor = new PdfColor(0.11, 0.22, 0.33);
        style.RowSeparatorWidth = 0.8;
        style.HeaderSeparatorColor = new PdfColor(0.44, 0.55, 0.66);
        style.HeaderSeparatorWidth = 1.2;
        style.FooterSeparatorColor = new PdfColor(0.22, 0.33, 0.44);
        style.FooterSeparatorWidth = 1.4;
        style.MaxWidth = 160;
        style.LeftIndent = 18;

        var block = new TableBlock(new[] {
            new[] { "Name", "Value" },
            body
        }, PdfAlign.Left, style);

        block.AddLink((1, 0), "https://evotec.xyz");
        body[0] = "Mutated";
        body[1] = "AlsoMutated";
        style.BorderWidth = 5;
        style.CellPaddingX = 20;
        style.RowSeparatorColor = PdfColor.Black;
        style.RowSeparatorWidth = 2;
        style.HeaderSeparatorColor = PdfColor.White;
        style.HeaderSeparatorWidth = 3;
        style.FooterSeparatorColor = PdfColor.White;
        style.FooterSeparatorWidth = 4;
        style.MaxWidth = 220;
        style.LeftIndent = 30;

        Assert.False(block.Rows is List<string[]>);
        Assert.False(block.Cells is List<IReadOnlyList<PdfTableCell>>);
        Assert.False(block.Links is Dictionary<(int Row, int Col), string>);
        Assert.Equal("Original", block.Rows[1][0]);
        Assert.Equal(string.Empty, block.Rows[1][1]);
        Assert.Equal(2, block.ColumnCount);
        Assert.Equal("Original", block.Cells[1][0].Text);
        Assert.Equal(1, block.Cells[1][0].ColumnSpan);
        Assert.Equal(1, block.Cells[1][0].RowSpan);
        Assert.Equal(2, block.Style!.BorderWidth);
        Assert.Equal(7, block.Style.CellPaddingX);
        Assert.Equal(new PdfColor(0.11, 0.22, 0.33), block.Style.RowSeparatorColor);
        Assert.Equal(0.8, block.Style.RowSeparatorWidth);
        Assert.Equal(new PdfColor(0.44, 0.55, 0.66), block.Style.HeaderSeparatorColor);
        Assert.Equal(1.2, block.Style.HeaderSeparatorWidth);
        Assert.Equal(new PdfColor(0.22, 0.33, 0.44), block.Style.FooterSeparatorColor);
        Assert.Equal(1.4, block.Style.FooterSeparatorWidth);
        Assert.Equal(160, block.Style.MaxWidth);
        Assert.Equal(18, block.Style.LeftIndent);
        Assert.True(block.Links.TryGetValue((1, 0), out string? uri));
        Assert.Equal("https://evotec.xyz", uri);
    }

    [Fact]
    public void PdfDoc_DefaultTableStyleAppliesToFollowingTablesAndSnapshotsInput() {
        var options = new PdfOptions {
            PageWidth = 300,
            PageHeight = 220,
            MarginLeft = 30,
            MarginRight = 30,
            MarginTop = 30,
            MarginBottom = 30,
            DefaultFont = PdfStandardFont.Helvetica,
            DefaultFontSize = 10
        };
        var style = TableStyles.Minimal();
        style.CellPaddingX = 22;

        byte[] bytes = PdfDoc.Create(options)
            .DefaultTableStyle(style)
            .Table(new[] {
                new[] { "DefaultPad", "Value" },
                new[] { "Row", "1" }
            })
            .ToBytes();

        style.CellPaddingX = 0;

        using var pdf = PdfDocument.Open(new MemoryStream(bytes));
        var page = pdf.GetPage(1);

        double markerX = FindWordStartX(page, "DefaultPad");

        Assert.True(markerX - options.MarginLeft >= 20, $"Expected fluent default table style padding to affect following tables and snapshot caller input. Marker x: {markerX:0.##}, margin: {options.MarginLeft:0.##}.");
    }

    [Fact]
    public void PdfDoc_ExplicitTableStyleOverridesDefaultTableStyle() {
        var options = new PdfOptions {
            PageWidth = 300,
            PageHeight = 220,
            MarginLeft = 30,
            MarginRight = 30,
            MarginTop = 30,
            MarginBottom = 30,
            DefaultFont = PdfStandardFont.Helvetica,
            DefaultFontSize = 10
        };
        var defaultStyle = TableStyles.Minimal();
        defaultStyle.CellPaddingX = 24;
        var explicitStyle = TableStyles.Minimal();
        explicitStyle.CellPaddingX = 2;

        byte[] bytes = PdfDoc.Create(options)
            .DefaultTableStyle(defaultStyle)
            .Table(new[] {
                new[] { "ExplicitPad", "Value" },
                new[] { "Row", "1" }
            }, style: explicitStyle)
            .ToBytes();

        using var pdf = PdfDocument.Open(new MemoryStream(bytes));
        var page = pdf.GetPage(1);

        double markerX = FindWordStartX(page, "ExplicitPad");

        Assert.True(markerX - options.MarginLeft <= 6, $"Expected explicit table style padding to override the document default. Marker x: {markerX:0.##}, margin: {options.MarginLeft:0.##}.");
    }

    [Fact]
    public void PdfDoc_DefaultTableStyleRejectsInvalidInputs() {
        Assert.Throws<ArgumentNullException>(() => PdfDoc.Create().DefaultTableStyle((PdfTableStyle)null!));
        Assert.Throws<ArgumentNullException>(() => PdfDoc.Create().DefaultTableStyle((string)null!));
        Assert.Throws<ArgumentException>(() => PdfDoc.Create().DefaultTableStyle("Missing Table Style"));
    }

    [Fact]
    public void TableWithLinks_SnapshotsInputLinkDictionaryBeforeRendering() {
        var links = new Dictionary<(int Row, int Col), string> {
            [(1, 0)] = "https://evotec.xyz"
        };

        var doc = PdfDoc.Create()
            .TableWithLinks(
                new[] { new[] { "Name" }, new[] { "OfficeIMO" } },
                links);

        links[(1, 0)] = "https://example.com";

        byte[] bytes = doc.ToBytes();
        string content = Encoding.ASCII.GetString(bytes);

        Assert.Contains("(https://evotec.xyz)", content);
        Assert.DoesNotContain("(https://example.com)", content);
    }

    [Fact]
    public void LinkAnnotations_RenderForParagraphHeadingAndTableCells() {
        byte[] bytes = PdfDoc.Create(new PdfOptions {
                DefaultFont = PdfStandardFont.Helvetica,
                DefaultFontSize = 11
            })
            .H1("Heading link", linkUri: "https://evotec.xyz/heading", linkContents: "Heading (metadata)")
            .Paragraph(p => p
                .Text("Visit ")
                .Link("paragraph link", "https://evotec.xyz/paragraph", contents: "Paragraph \\ metadata")
                .Text(" for details."))
            .TableWithLinks(
                new[] {
                    new[] { "Name", "Url" },
                    new[] { "OfficeIMO", "Open" }
                },
                new Dictionary<(int Row, int Col), string> {
                    [(1, 1)] = "https://evotec.xyz/table"
                })
            .ToBytes();

        string pdf = Encoding.ASCII.GetString(bytes);

        Assert.Contains("/Annots [", pdf, StringComparison.Ordinal);
        Assert.Equal(4, CountOccurrences(pdf, "/Subtype /Link"));
        Assert.Equal(4, CountOccurrences(pdf, "/S /URI"));
        Assert.Equal(1, CountOccurrences(pdf, "/URI (https://evotec.xyz/heading)"));
        Assert.Equal(2, CountOccurrences(pdf, "/URI (https://evotec.xyz/paragraph)"));
        Assert.Equal(1, CountOccurrences(pdf, "/URI (https://evotec.xyz/table)"));
        Assert.Equal(4, CountOccurrences(pdf, "/Contents ("));
        Assert.Equal(1, CountOccurrences(pdf, "/Contents (Heading \\(metadata\\))"));
        Assert.Equal(2, CountOccurrences(pdf, "/Contents (Paragraph \\\\ metadata)"));
        Assert.Equal(1, CountOccurrences(pdf, "/Contents (Open)"));

        var rectangles = ExtractLinkRectangles(pdf);
        Assert.Equal(4, rectangles.Count);
        foreach (var rect in rectangles) {
            Assert.True(rect.X2 > rect.X1, "Link annotation rectangle must have positive width.");
            Assert.True(rect.Y2 > rect.Y1, "Link annotation rectangle must have positive height.");
            Assert.InRange(rect.X1, 0, 612);
            Assert.InRange(rect.X2, 0, 612);
            Assert.InRange(rect.Y1, 0, 792);
            Assert.InRange(rect.Y2, 0, 792);
        }
    }

    [Fact]
    public void ImageLink_RendersAnnotationFromFinalImagePlacement() {
        var options = new PdfOptions {
            PageWidth = 220,
            PageHeight = 180,
            MarginLeft = 30,
            MarginRight = 30,
            MarginTop = 30,
            MarginBottom = 30,
            DefaultFont = PdfStandardFont.Helvetica,
            DefaultFontSize = 10
        };

        byte[] bytes = PdfDoc.Create(options)
            .Image(CreateMinimalRgbPng(), 80, 40, PdfAlign.Center, fit: OfficeImageFit.Contain, linkUri: "https://evotec.xyz/image", linkContents: "Image metadata")
            .ToBytes();

        string pdf = Encoding.ASCII.GetString(bytes);
        var rect = Assert.Single(ExtractLinkRectangles(pdf));

        Assert.Equal(1, CountOccurrences(pdf, "/Subtype /Link"));
        Assert.Equal(1, CountOccurrences(pdf, "/URI (https://evotec.xyz/image)"));
        Assert.Equal(1, CountOccurrences(pdf, "/Contents (Image metadata)"));
        Assert.InRange(rect.X1, 89.5, 90.5);
        Assert.InRange(rect.X2, 129.5, 130.5);
        Assert.InRange(rect.Y1, 109.5, 110.5);
        Assert.InRange(rect.Y2, 149.5, 150.5);
    }

    [Fact]
    public void RowColumnImageLink_RendersLinkAnnotation() {
        byte[] bytes = PdfDoc.Create(new PdfOptions {
                PageWidth = 220,
                PageHeight = 180,
                MarginLeft = 30,
                MarginRight = 30,
                MarginTop = 30,
                MarginBottom = 30,
                DefaultFont = PdfStandardFont.Helvetica,
                DefaultFontSize = 10
            })
            .Compose(document =>
                document.Page(page =>
                    page.Content(content =>
                        content.Row(row =>
                            row.Column(100, column =>
                                column.Image(CreateMinimalRgbPng(), 24, 24, PdfAlign.Right, linkUri: "https://evotec.xyz/column-image", linkContents: "Column image"))))))
            .ToBytes();

        string pdf = Encoding.ASCII.GetString(bytes);
        var rect = Assert.Single(ExtractLinkRectangles(pdf));

        Assert.Equal(1, CountOccurrences(pdf, "/Subtype /Link"));
        Assert.Equal(1, CountOccurrences(pdf, "/URI (https://evotec.xyz/column-image)"));
        Assert.Equal(1, CountOccurrences(pdf, "/Contents (Column image)"));
        Assert.True(rect.X2 > rect.X1, "Row-column image link annotation rectangle must have positive width.");
        Assert.True(rect.Y2 > rect.Y1, "Row-column image link annotation rectangle must have positive height.");
        Assert.InRange(rect.X2, 185.5, 190.5);
    }

    [Fact]
    public void ShapeLink_RendersAnnotationFromFinalShapePlacement() {
        var options = new PdfOptions {
            PageWidth = 220,
            PageHeight = 180,
            MarginLeft = 30,
            MarginRight = 30,
            MarginTop = 30,
            MarginBottom = 30,
            DefaultFont = PdfStandardFont.Helvetica,
            DefaultFontSize = 10
        };
        var shape = OfficeShape.Rectangle(40, 20);

        byte[] bytes = PdfDoc.Create(options)
            .Shape(shape, PdfAlign.Right, linkUri: "https://evotec.xyz/shape", linkContents: "Shape metadata")
            .ToBytes();

        string pdf = Encoding.ASCII.GetString(bytes);
        var rect = Assert.Single(ExtractLinkRectangles(pdf));

        Assert.Equal(1, CountOccurrences(pdf, "/Subtype /Link"));
        Assert.Equal(1, CountOccurrences(pdf, "/URI (https://evotec.xyz/shape)"));
        Assert.Equal(1, CountOccurrences(pdf, "/Contents (Shape metadata)"));
        Assert.InRange(rect.X1, 149.5, 150.5);
        Assert.InRange(rect.X2, 189.5, 190.5);
        Assert.InRange(rect.Y1, 129.5, 130.5);
        Assert.InRange(rect.Y2, 149.5, 150.5);
    }

    [Fact]
    public void ConvenienceVectorLink_RendersAnnotationFromFinalPlacement() {
        var options = new PdfOptions {
            PageWidth = 220,
            PageHeight = 180,
            MarginLeft = 30,
            MarginRight = 30,
            MarginTop = 30,
            MarginBottom = 30,
            DefaultFont = PdfStandardFont.Helvetica,
            DefaultFontSize = 10
        };

        byte[] bytes = PdfDoc.Create(options)
            .Rectangle(40, 20, align: PdfAlign.Center, linkUri: "https://evotec.xyz/rectangle", linkContents: "Rectangle metadata")
            .ToBytes();

        string pdf = Encoding.ASCII.GetString(bytes);
        var rect = Assert.Single(ExtractLinkRectangles(pdf));

        Assert.Equal(1, CountOccurrences(pdf, "/Subtype /Link"));
        Assert.Equal(1, CountOccurrences(pdf, "/URI (https://evotec.xyz/rectangle)"));
        Assert.Equal(1, CountOccurrences(pdf, "/Contents (Rectangle metadata)"));
        Assert.InRange(rect.X1, 89.5, 90.5);
        Assert.InRange(rect.X2, 129.5, 130.5);
        Assert.InRange(rect.Y1, 129.5, 130.5);
        Assert.InRange(rect.Y2, 149.5, 150.5);
    }

    [Fact]
    public void ComposeConvenienceVectorLinks_RenderLinkAnnotations() {
        byte[] bytes = PdfDoc.Create(new PdfOptions {
                PageWidth = 220,
                PageHeight = 220,
                MarginLeft = 30,
                MarginRight = 30,
                MarginTop = 30,
                MarginBottom = 30,
                DefaultFont = PdfStandardFont.Helvetica,
                DefaultFontSize = 10
            })
            .Compose(document =>
                document.Page(page =>
                    page.Content(content => content
                        .Item(item => item.Rectangle(40, 20, align: PdfAlign.Center, linkUri: "https://evotec.xyz/item-rectangle", linkContents: "Item rectangle"))
                        .Item(item => item.Element(element =>
                            element.Ellipse(30, 18, align: PdfAlign.Right, spacingBefore: 4, linkUri: "https://evotec.xyz/element-ellipse", linkContents: "Element ellipse"))))))
            .ToBytes();

        string pdf = Encoding.ASCII.GetString(bytes);
        var rectangles = ExtractLinkRectangles(pdf);

        Assert.Equal(2, CountOccurrences(pdf, "/Subtype /Link"));
        Assert.Equal(1, CountOccurrences(pdf, "/URI (https://evotec.xyz/item-rectangle)"));
        Assert.Equal(1, CountOccurrences(pdf, "/URI (https://evotec.xyz/element-ellipse)"));
        Assert.Equal(1, CountOccurrences(pdf, "/Contents (Item rectangle)"));
        Assert.Equal(1, CountOccurrences(pdf, "/Contents (Element ellipse)"));
        Assert.Equal(2, rectangles.Count);
        Assert.All(rectangles, rect => {
            Assert.True(rect.X2 > rect.X1, "Compose vector link annotation rectangle must have positive width.");
            Assert.True(rect.Y2 > rect.Y1, "Compose vector link annotation rectangle must have positive height.");
        });
    }

    [Fact]
    public void DrawingLink_RendersAnnotationFromFinalDrawingPlacement() {
        var options = new PdfOptions {
            PageWidth = 220,
            PageHeight = 180,
            MarginLeft = 30,
            MarginRight = 30,
            MarginTop = 30,
            MarginBottom = 30,
            DefaultFont = PdfStandardFont.Helvetica,
            DefaultFontSize = 10
        };
        var drawing = new OfficeDrawing(60, 30)
            .AddShape(OfficeShape.Rectangle(60, 30), 0, 0);

        byte[] bytes = PdfDoc.Create(options)
            .Drawing(drawing, PdfAlign.Center, linkUri: "https://evotec.xyz/drawing", linkContents: "Drawing metadata")
            .ToBytes();

        string pdf = Encoding.ASCII.GetString(bytes);
        var rect = Assert.Single(ExtractLinkRectangles(pdf));

        Assert.Equal(1, CountOccurrences(pdf, "/Subtype /Link"));
        Assert.Equal(1, CountOccurrences(pdf, "/URI (https://evotec.xyz/drawing)"));
        Assert.Equal(1, CountOccurrences(pdf, "/Contents (Drawing metadata)"));
        Assert.InRange(rect.X1, 79.5, 80.5);
        Assert.InRange(rect.X2, 139.5, 140.5);
        Assert.InRange(rect.Y1, 119.5, 120.5);
        Assert.InRange(rect.Y2, 149.5, 150.5);
    }

    [Fact]
    public void RowColumnConvenienceVectorLinks_RenderLinkAnnotations() {
        byte[] bytes = PdfDoc.Create(new PdfOptions {
                PageWidth = 220,
                PageHeight = 180,
                MarginLeft = 30,
                MarginRight = 30,
                MarginTop = 30,
                MarginBottom = 30,
                DefaultFont = PdfStandardFont.Helvetica,
                DefaultFontSize = 10
            })
            .Compose(document =>
                document.Page(page =>
                    page.Content(content =>
                        content.Row(row =>
                            row.Column(100, column => column
                                .Rectangle(24, 18, align: PdfAlign.Right, linkUri: "https://evotec.xyz/column-rectangle", linkContents: "Column rectangle")
                                .Ellipse(24, 18, align: PdfAlign.Right, spacingBefore: 4, linkUri: "https://evotec.xyz/column-ellipse", linkContents: "Column ellipse"))))))
            .ToBytes();

        string pdf = Encoding.ASCII.GetString(bytes);
        var rectangles = ExtractLinkRectangles(pdf);

        Assert.Equal(2, CountOccurrences(pdf, "/Subtype /Link"));
        Assert.Equal(1, CountOccurrences(pdf, "/URI (https://evotec.xyz/column-rectangle)"));
        Assert.Equal(1, CountOccurrences(pdf, "/URI (https://evotec.xyz/column-ellipse)"));
        Assert.Equal(1, CountOccurrences(pdf, "/Contents (Column rectangle)"));
        Assert.Equal(1, CountOccurrences(pdf, "/Contents (Column ellipse)"));
        Assert.Equal(2, rectangles.Count);
        Assert.All(rectangles, rect => {
            Assert.True(rect.X2 > rect.X1, "Row-column convenience vector link annotation rectangle must have positive width.");
            Assert.True(rect.Y2 > rect.Y1, "Row-column convenience vector link annotation rectangle must have positive height.");
        });
    }

    [Fact]
    public void RowColumnShapeAndDrawingLinks_RenderLinkAnnotations() {
        var drawing = new OfficeDrawing(24, 18)
            .AddShape(OfficeShape.Rectangle(24, 18), 0, 0);

        byte[] bytes = PdfDoc.Create(new PdfOptions {
                PageWidth = 220,
                PageHeight = 180,
                MarginLeft = 30,
                MarginRight = 30,
                MarginTop = 30,
                MarginBottom = 30,
                DefaultFont = PdfStandardFont.Helvetica,
                DefaultFontSize = 10
            })
            .Compose(document =>
                document.Page(page =>
                    page.Content(content =>
                        content.Row(row =>
                            row.Column(100, column => column
                                .Shape(OfficeShape.Rectangle(24, 18), PdfAlign.Right, linkUri: "https://evotec.xyz/column-shape", linkContents: "Column shape")
                                .Drawing(drawing, PdfAlign.Right, spacingBefore: 4, linkUri: "https://evotec.xyz/column-drawing", linkContents: "Column drawing"))))))
            .ToBytes();

        string pdf = Encoding.ASCII.GetString(bytes);
        var rectangles = ExtractLinkRectangles(pdf);

        Assert.Equal(2, CountOccurrences(pdf, "/Subtype /Link"));
        Assert.Equal(1, CountOccurrences(pdf, "/URI (https://evotec.xyz/column-shape)"));
        Assert.Equal(1, CountOccurrences(pdf, "/URI (https://evotec.xyz/column-drawing)"));
        Assert.Equal(1, CountOccurrences(pdf, "/Contents (Column shape)"));
        Assert.Equal(1, CountOccurrences(pdf, "/Contents (Column drawing)"));
        Assert.Equal(2, rectangles.Count);
        Assert.All(rectangles, rect => {
            Assert.True(rect.X2 > rect.X1, "Row-column vector link annotation rectangle must have positive width.");
            Assert.True(rect.Y2 > rect.Y1, "Row-column vector link annotation rectangle must have positive height.");
        });
    }

    [Fact]
    public void WrappedHeadingLink_RendersAnnotationForEachVisualLine() {
        var options = new PdfOptions {
            PageWidth = 140,
            PageHeight = 160,
            MarginLeft = 30,
            MarginRight = 30,
            MarginTop = 30,
            MarginBottom = 30,
            DefaultFont = PdfStandardFont.Helvetica,
            DefaultFontSize = 10
        };

        byte[] bytes = PdfDoc.Create(options)
            .H3("WWWWWWWW", linkUri: "https://evotec.xyz/wrapped-heading", linkContents: "Wrapped heading")
            .ToBytes();

        string pdf = Encoding.ASCII.GetString(bytes);
        int linkCount = CountOccurrences(pdf, "/URI (https://evotec.xyz/wrapped-heading)");
        var rectangles = ExtractLinkRectangles(pdf);

        Assert.True(linkCount > 1, "Expected a wrapped heading link to emit one annotation per visual line.");
        Assert.Equal(linkCount, rectangles.Count);
        Assert.Equal(linkCount, CountOccurrences(pdf, "/Contents (Wrapped heading)"));
        Assert.All(rectangles, rect => {
            Assert.True(rect.X2 > rect.X1, "Heading link annotation rectangle must have positive width.");
            Assert.True(rect.Y2 > rect.Y1, "Heading link annotation rectangle must have positive height.");
            Assert.InRange(rect.X1, options.MarginLeft - 0.5, options.PageWidth - options.MarginRight + 0.5);
            Assert.InRange(rect.X2, options.MarginLeft - 0.5, options.PageWidth - options.MarginRight + 0.5);
        });
    }

    [Fact]
    public void RowColumnHeadingLink_AlignsAnnotationWithRightAlignedText() {
        var options = new PdfOptions {
            PageWidth = 320,
            PageHeight = 180,
            MarginLeft = 30,
            MarginRight = 30,
            MarginTop = 30,
            MarginBottom = 30,
            DefaultFont = PdfStandardFont.Helvetica,
            DefaultFontSize = 10
        };

        byte[] bytes = PdfDoc.Create(options)
            .Compose(document =>
                document.Page(page =>
                    page.Content(content =>
                        content.Row(row =>
                            row.Column(100, column =>
                                column.H3("ColumnHead", PdfAlign.Right, linkUri: "https://evotec.xyz/right-heading", linkContents: "Right heading"))))))
            .ToBytes();

        string pdfText = Encoding.ASCII.GetString(bytes);
        var rect = Assert.Single(ExtractLinkRectangles(pdfText));

        using var pdf = PdfDocument.Open(new MemoryStream(bytes));
        var page = pdf.GetPage(1);
        double headingStartX = FindWordStartX(page, "ColumnHead");
        double headingEndX = FindWordEndX(page, "ColumnHead");
        double expectedRightEdge = options.PageWidth - options.MarginRight;

        Assert.InRange(Math.Abs(expectedRightEdge - headingEndX), 0, 5);
        Assert.InRange(Math.Abs(headingStartX - rect.X1), 0, 2.5);
        Assert.InRange(Math.Abs(headingEndX - rect.X2), 0, 2.5);
    }

    [Fact]
    public void RowColumnHeadingLink_RendersLinkAnnotation() {
        byte[] bytes = PdfDoc.Create(new PdfOptions {
                DefaultFont = PdfStandardFont.Helvetica,
                DefaultFontSize = 11
            })
            .Compose(document =>
                document.Page(page =>
                    page.Content(content =>
                        content.Row(row =>
                            row.Column(100, column =>
                                column.H3("Column heading", linkUri: "https://evotec.xyz/column-heading", linkContents: "Column heading metadata"))))))
            .ToBytes();

        string pdf = Encoding.ASCII.GetString(bytes);
        var rectangles = ExtractLinkRectangles(pdf);

        Assert.Equal(1, CountOccurrences(pdf, "/Subtype /Link"));
        Assert.Equal(1, CountOccurrences(pdf, "/URI (https://evotec.xyz/column-heading)"));
        Assert.Equal(1, CountOccurrences(pdf, "/Contents (Column heading metadata)"));
        var rect = Assert.Single(rectangles);
        Assert.True(rect.X2 > rect.X1, "Row-column heading link annotation rectangle must have positive width.");
        Assert.True(rect.Y2 > rect.Y1, "Row-column heading link annotation rectangle must have positive height.");
    }

    [Fact]
    public void RowColumnTableWithLinks_RendersTableCellLinkAnnotations() {
        byte[] bytes = PdfDoc.Create(new PdfOptions {
                DefaultFont = PdfStandardFont.Helvetica,
                DefaultFontSize = 11
            })
            .Compose(document =>
                document.Page(page =>
                    page.Content(content =>
                        content.Row(row =>
                            row.Column(100, column =>
                                column.TableWithLinks(
                                    new[] {
                                        new[] { "Name", "Url" },
                                        new[] { "OfficeIMO", "Open" }
                                    },
                                    new Dictionary<(int Row, int Col), string> {
                                        [(1, 1)] = "https://evotec.xyz/row-column-table"
                                    }))))))
            .ToBytes();

        string pdf = Encoding.ASCII.GetString(bytes);

        Assert.Contains("/Annots [", pdf, StringComparison.Ordinal);
        Assert.Equal(1, CountOccurrences(pdf, "/Subtype /Link"));
        Assert.Equal(1, CountOccurrences(pdf, "/S /URI"));
        Assert.Equal(1, CountOccurrences(pdf, "/URI (https://evotec.xyz/row-column-table)"));
        Assert.Equal(1, CountOccurrences(pdf, "/Contents (Open)"));

        var rectangles = ExtractLinkRectangles(pdf);
        Assert.Single(rectangles);
        var rect = rectangles[0];
        Assert.True(rect.X2 > rect.X1, "Row-column table link annotation rectangle must have positive width.");
        Assert.True(rect.Y2 > rect.Y1, "Row-column table link annotation rectangle must have positive height.");
        Assert.InRange(rect.X1, 0, 612);
        Assert.InRange(rect.X2, 0, 612);
        Assert.InRange(rect.Y1, 0, 792);
        Assert.InRange(rect.Y2, 0, 792);
    }

    [Fact]
    public void Bookmark_RejectsInvalidNamesAndDuplicateNames() {
        Assert.Throws<ArgumentNullException>(() =>
            PdfDoc.Create().Bookmark(null!));

        Assert.Throws<ArgumentException>(() =>
            PdfDoc.Create().Bookmark(" "));

        var duplicateException = Assert.Throws<ArgumentException>(() =>
            PdfDoc.Create()
                .Bookmark("Intro")
                .Paragraph(p => p.Text("First target."))
                .Bookmark("Intro")
                .Paragraph(p => p.Text("Second target."))
                .ToBytes());

        Assert.Contains("PDF bookmark names must be unique.", duplicateException.Message, StringComparison.Ordinal);
    }

    [Fact]
    public void BookmarkLinks_RejectInvalidTargetsAndMissingBookmarks() {
        Assert.Throws<ArgumentNullException>(() =>
            PdfDoc.Create().Paragraph(p => p.LinkToBookmark("Jump", null!)));

        Assert.Throws<ArgumentException>(() =>
            PdfDoc.Create().Paragraph(p => p.LinkToBookmark("Jump", " ")));

        Assert.Throws<ArgumentException>(() =>
            PdfDoc.Create().Paragraph(p => p.LinkToBookmark("", "Intro")));

        Assert.Throws<ArgumentException>(() =>
            PdfDoc.Create().Paragraph(p => p.LinkToBookmark("Jump", "Intro", contents: " ")));

        var missingTargetException = Assert.Throws<ArgumentException>(() =>
            PdfDoc.Create()
                .Paragraph(p => p.LinkToBookmark("Jump to missing bookmark", "MissingBookmark"))
                .ToBytes());

        Assert.Contains("PDF bookmark link target 'MissingBookmark' was not found.", missingTargetException.Message, StringComparison.Ordinal);
    }

    [Fact]
    public void Bookmark_RendersNamedDestinationNameTreeAndInspectorReadsIt() {
        byte[] bytes = PdfDoc.Create(new PdfOptions {
                PageWidth = 220,
                PageHeight = 180,
                MarginLeft = 30,
                MarginRight = 30,
                MarginTop = 30,
                MarginBottom = 30,
                DefaultFont = PdfStandardFont.Helvetica,
                DefaultFontSize = 10
            })
            .Bookmark("Intro (A)")
            .H1("Intro")
            .Paragraph(p => p.Text("Opening text."))
            .Bookmark("Details")
            .Paragraph(p => p.Text("Details text."))
            .ToBytes();

        string pdf = Encoding.ASCII.GetString(bytes);

        Assert.Contains("/Names << /Dests ", pdf, StringComparison.Ordinal);
        Assert.Contains("(Intro \\(A\\))", pdf, StringComparison.Ordinal);
        Assert.Contains("(Details)", pdf, StringComparison.Ordinal);

        PdfDocumentInfo info = PdfInspector.Inspect(bytes);
        Assert.True(info.HasNamedDestinations);
        Assert.True(info.HasCatalogNameTrees);
        Assert.Equal(2, info.NamedDestinationCount);
        Assert.Contains("Intro (A)", info.NamedDestinationNames);
        Assert.Contains("Details", info.NamedDestinationNames);
        Assert.All(info.NamedDestinations, destination => {
            Assert.Equal(1, destination.PageNumber);
            Assert.NotNull(destination.DestinationTop);
            Assert.InRange(destination.DestinationTop!.Value, 0, 180);
        });
    }

    [Fact]
    public void BookmarkLink_RendersGoToAnnotationAndInspectorReadsTarget() {
        byte[] bytes = PdfDoc.Create(new PdfOptions {
                PageWidth = 220,
                PageHeight = 220,
                MarginLeft = 30,
                MarginRight = 30,
                MarginTop = 30,
                MarginBottom = 30,
                DefaultFont = PdfStandardFont.Helvetica,
                DefaultFontSize = 10
            })
            .Paragraph(p => p
                .Text("See ")
                .LinkToBookmark("details", "Details", PdfColor.FromRgb(20, 90, 180), contents: "Jump to details")
                .Text("."))
            .Spacer(20)
            .Bookmark("Details")
            .H2("Details")
            .Paragraph(p => p.Text("Destination text."))
            .ToBytes();

        string pdf = Encoding.ASCII.GetString(bytes);
        var rect = Assert.Single(ExtractLinkRectangles(pdf));

        Assert.Equal(1, CountOccurrences(pdf, "/Subtype /Link"));
        Assert.Equal(1, CountOccurrences(pdf, "/S /GoTo"));
        Assert.Equal(1, CountOccurrences(pdf, "/D (Details)"));
        Assert.Equal(1, CountOccurrences(pdf, "/Contents (Jump to details)"));
        Assert.DoesNotContain("/S /URI", pdf, StringComparison.Ordinal);
        Assert.True(rect.X2 > rect.X1, "Bookmark link annotation rectangle must have positive width.");
        Assert.True(rect.Y2 > rect.Y1, "Bookmark link annotation rectangle must have positive height.");

        PdfDocumentInfo info = PdfInspector.Inspect(bytes);
        PdfLinkAnnotation link = Assert.Single(info.LinkAnnotations);
        Assert.True(link.IsNamedDestinationLink);
        Assert.False(link.IsUriLink);
        Assert.Null(link.Uri);
        Assert.Equal("Details", link.DestinationName);
        Assert.Equal("Jump to details", link.Contents);
        Assert.Equal(new[] { "Details" }, info.LinkDestinationNames);
        Assert.Equal(0, info.LinkUriCount);
        Assert.Empty(info.LinkUris);
    }

    [Fact]
    public void RowColumnBookmark_RendersNamedDestinationAtColumnFlowPosition() {
        byte[] bytes = PdfDoc.Create(new PdfOptions {
                PageWidth = 260,
                PageHeight = 180,
                MarginLeft = 30,
                MarginRight = 30,
                MarginTop = 30,
                MarginBottom = 30,
                DefaultFont = PdfStandardFont.Helvetica,
                DefaultFontSize = 10
            })
            .Compose(document =>
                document.Page(page =>
                    page.Content(content =>
                        content.Row(row =>
                            row.Column(100, column => column
                                .Bookmark("ColumnStart")
                                .H3("Column heading")
                                .Paragraph(p => p.Text("Column body.")))))))
            .ToBytes();

        PdfDocumentInfo info = PdfInspector.Inspect(bytes);
        PdfNamedDestination destination = Assert.Single(info.NamedDestinations);

        Assert.True(info.HasNamedDestinations);
        Assert.Equal("ColumnStart", destination.Name);
        Assert.Equal(1, destination.PageNumber);
        Assert.InRange(destination.DestinationTop!.Value, 149.5, 150.5);
    }

    [Fact]
    public void RowColumnBookmarkOnly_RendersZeroHeightNamedDestination() {
        byte[] bytes = PdfDoc.Create(new PdfOptions {
                PageWidth = 220,
                PageHeight = 180,
                MarginLeft = 30,
                MarginRight = 30,
                MarginTop = 30,
                MarginBottom = 30,
                DefaultFont = PdfStandardFont.Helvetica,
                DefaultFontSize = 10
            })
            .Compose(document =>
                document.Page(page =>
                    page.Content(content =>
                        content.Row(row =>
                            row.Column(100, column =>
                                column.Bookmark("InvisibleColumnAnchor"))))))
            .ToBytes();

        PdfDocumentInfo info = PdfInspector.Inspect(bytes);
        PdfNamedDestination destination = Assert.Single(info.NamedDestinations);

        Assert.Equal(1, info.PageCount);
        Assert.Equal("InvisibleColumnAnchor", destination.Name);
        Assert.Equal(1, destination.PageNumber);
        Assert.InRange(destination.DestinationTop!.Value, 149.5, 150.5);
    }

    [Fact]
    public void Paragraph_UsesNaturalWordSpacingForProportionalFonts() {
        var options = new PdfOptions {
            DefaultFont = PdfStandardFont.Helvetica,
            DefaultFontSize = 12
        };

        byte[] bytes = PdfDoc.Create(options)
            .Paragraph(p => p.Text("This sample uses proportional Helvetica text and should not stretch spaces between every word."))
            .ToBytes();

        using var pdf = PdfDocument.Open(new MemoryStream(bytes));
        var page = pdf.GetPage(1);
        var firstLineLetters = page.Letters
            .Where(letter => !string.IsNullOrWhiteSpace(letter.Value))
            .GroupBy(letter => Math.Round(letter.StartBaseLine.Y, 1))
            .OrderByDescending(group => group.Key)
            .First()
            .OrderBy(letter => letter.StartBaseLine.X)
            .ToList();

        var gaps = firstLineLetters
            .Zip(firstLineLetters.Skip(1), (left, right) => right.StartBaseLine.X - left.EndBaseLine.X)
            .Where(gap => gap > 1)
            .ToList();

        Assert.NotEmpty(gaps);
        Assert.True(gaps.Max() < 9, $"Expected natural word spacing, but found a {gaps.Max():0.##}pt gap.");
    }

    [Fact]
    public void Paragraph_UsesProportionalGlyphWidthsForWrapping() {
        var options = new PdfOptions {
            PageWidth = 100,
            PageHeight = 160,
            MarginLeft = 35,
            MarginRight = 35,
            MarginTop = 25,
            MarginBottom = 25,
            DefaultFont = PdfStandardFont.Helvetica,
            DefaultFontSize = 10
        };

        byte[] narrowBytes = PdfDoc.Create(options)
            .Paragraph(p => p.Text("Illii Illii"))
            .ToBytes();

        byte[] wideBytes = PdfDoc.Create(options)
            .Paragraph(p => p.Text("WWWW"))
            .ToBytes();

        using var narrowPdf = PdfDocument.Open(new MemoryStream(narrowBytes));
        using var widePdf = PdfDocument.Open(new MemoryStream(wideBytes));

        Assert.Equal(1, CountTextLines(narrowPdf.GetPage(1)));
        Assert.True(CountTextLines(widePdf.GetPage(1)) >= 2, "Expected wide Helvetica glyphs to wrap instead of overrunning the text frame.");
    }

    [Fact]
    public void Paragraph_JustifyExpandsWrappedLinesButNotFinalLine() {
        var options = new PdfOptions {
            PageWidth = 220,
            PageHeight = 240,
            MarginLeft = 30,
            MarginRight = 30,
            MarginTop = 30,
            MarginBottom = 30,
            DefaultFont = PdfStandardFont.Helvetica,
            DefaultFontSize = 12
        };

        byte[] bytes = PdfDoc.Create(options)
            .Paragraph(p => p.Text("Alpha beta gamma delta epsilon zeta eta theta iota kappa lambda mu nu xi omicron pi rho sigma tau."), PdfAlign.Justify)
            .ToBytes();

        using var pdf = PdfDocument.Open(new MemoryStream(bytes));
        var page = pdf.GetPage(1);
        var lines = GetNonWhitespaceLetterLines(page);

        Assert.True(lines.Count >= 2, "Expected the justified paragraph to wrap onto at least two lines.");

        var firstLineGaps = GetInterWordGaps(lines[0]);
        var lastLineGaps = GetInterWordGaps(lines[lines.Count - 1]);

        Assert.NotEmpty(firstLineGaps);
        Assert.NotEmpty(lastLineGaps);
        Assert.True(firstLineGaps.Max() > 9, $"Expected justification to expand wrapped-line gaps, but the largest gap was {firstLineGaps.Max():0.##}pt.");
        Assert.True(lastLineGaps.Max() < 9, $"Expected the final justified paragraph line to keep natural spacing, but found a {lastLineGaps.Max():0.##}pt gap.");

        string extracted = page.Text;
        Assert.Contains("Alpha", extracted, StringComparison.Ordinal);
        Assert.Contains("omicron", extracted, StringComparison.Ordinal);
    }

    [Fact]
    public void Paragraph_JustifyDoesNotStretchExplicitLineBreaks() {
        var options = new PdfOptions {
            PageWidth = 320,
            PageHeight = 220,
            MarginLeft = 30,
            MarginRight = 30,
            MarginTop = 30,
            MarginBottom = 30,
            DefaultFont = PdfStandardFont.Helvetica,
            DefaultFontSize = 12
        };

        byte[] bytes = PdfDoc.Create(options)
            .Paragraph(p => p
                .Text("Alpha beta gamma")
                .LineBreak()
                .Text("Second line continues with enough words to wrap naturally."), PdfAlign.Justify)
            .ToBytes();

        using var pdf = PdfDocument.Open(new MemoryStream(bytes));
        var page = pdf.GetPage(1);
        var lines = GetNonWhitespaceLetterLines(page);

        Assert.True(lines.Count >= 2, "Expected the explicit line break to create a second rendered line.");

        var hardBreakLineGaps = GetInterWordGaps(lines[0]);

        Assert.NotEmpty(hardBreakLineGaps);
        Assert.True(hardBreakLineGaps.Max() < 9, $"Expected explicit line-break text to keep natural spacing, but found a {hardBreakLineGaps.Max():0.##}pt gap.");
    }

    [Fact]
    public void Heading_RightAlignsUsingProportionalTextWidth() {
        var options = new PdfOptions {
            PageWidth = 320,
            PageHeight = 180,
            MarginLeft = 30,
            MarginRight = 30,
            MarginTop = 30,
            MarginBottom = 30,
            DefaultFont = PdfStandardFont.Helvetica,
            DefaultFontSize = 10
        };

        byte[] bytes = PdfDoc.Create(options)
            .H1("Illi", PdfAlign.Right)
            .ToBytes();

        using var pdf = PdfDocument.Open(new MemoryStream(bytes));
        var page = pdf.GetPage(1);
        double headingEndX = FindWordEndX(page, "Illi");
        double expectedRightEdge = options.PageWidth - options.MarginRight;

        Assert.InRange(Math.Abs(expectedRightEdge - headingEndX), 0, 5);
    }

    [Fact]
    public void ComposeItemHeading_AppliesExplicitAlignmentAndColor() {
        var options = new PdfOptions {
            PageWidth = 320,
            PageHeight = 180,
            MarginLeft = 30,
            MarginRight = 30,
            MarginTop = 30,
            MarginBottom = 30,
            DefaultFont = PdfStandardFont.Helvetica,
            DefaultFontSize = 10
        };

        byte[] bytes = PdfDoc.Create(options)
            .Compose(document =>
                document.Page(page =>
                    page.Content(content =>
                        content.Column(column =>
                            column.Item(item =>
                                item.H2("ComposeHead", PdfAlign.Center, PdfColor.FromRgb(10, 20, 30)))))))
            .ToBytes();

        using var pdf = PdfDocument.Open(new MemoryStream(bytes));
        var page = pdf.GetPage(1);
        var line = GetVisualTextLines(page, 0, options.PageWidth)
            .Single(line => line.Text.Contains("ComposeHead", StringComparison.Ordinal));
        double contentCenter = (options.MarginLeft + options.PageWidth - options.MarginRight) / 2;
        double lineCenter = (line.X1 + line.X2) / 2;
        string rawPdf = Encoding.ASCII.GetString(bytes);

        Assert.InRange(Math.Abs(contentCenter - lineCenter), 0, 5);
        Assert.Contains("0.039 0.078 0.118 rg", rawPdf, StringComparison.Ordinal);
    }

    [Fact]
    public void RowColumnHeading_AppliesExplicitAlignmentAndColor() {
        var options = new PdfOptions {
            PageWidth = 320,
            PageHeight = 180,
            MarginLeft = 30,
            MarginRight = 30,
            MarginTop = 30,
            MarginBottom = 30,
            DefaultFont = PdfStandardFont.Helvetica,
            DefaultFontSize = 10
        };

        byte[] bytes = PdfDoc.Create(options)
            .Compose(document =>
                document.Page(page =>
                    page.Content(content =>
                        content.Row(row =>
                            row.Column(100, column =>
                                column.H3("ColumnHead", PdfAlign.Right, PdfColor.FromRgb(10, 20, 30)))))))
            .ToBytes();

        using var pdf = PdfDocument.Open(new MemoryStream(bytes));
        var page = pdf.GetPage(1);
        double headingEndX = FindWordEndX(page, "ColumnHead");
        double expectedRightEdge = options.PageWidth - options.MarginRight;
        string rawPdf = Encoding.ASCII.GetString(bytes);

        Assert.InRange(Math.Abs(expectedRightEdge - headingEndX), 0, 5);
        Assert.Contains("0.039 0.078 0.118 rg", rawPdf, StringComparison.Ordinal);
    }

    [Fact]
    public void Heading_UsesProportionalGlyphWidthsForWideWrapping() {
        var options = new PdfOptions {
            PageWidth = 140,
            PageHeight = 160,
            MarginLeft = 30,
            MarginRight = 30,
            MarginTop = 30,
            MarginBottom = 30,
            DefaultFont = PdfStandardFont.Helvetica,
            DefaultFontSize = 10
        };

        byte[] bytes = PdfDoc.Create(options)
            .H3("WWWWWWWW")
            .ToBytes();

        using var pdf = PdfDocument.Open(new MemoryStream(bytes));
        var page = pdf.GetPage(1);
        int wideLineCount = page.Letters
            .Where(letter => letter.Value == "W")
            .GroupBy(letter => Math.Round(letter.StartBaseLine.Y, 1))
            .Count();
        double contentRight = options.PageWidth - options.MarginRight;
        double rightMostWideGlyph = page.Letters
            .Where(letter => letter.Value == "W")
            .Max(letter => letter.EndBaseLine.X);

        Assert.True(wideLineCount > 1, "Expected heading wide glyphs to wrap using their real Helvetica advance instead of an average character width.");
        Assert.InRange(rightMostWideGlyph, double.NegativeInfinity, contentRight + 1);
    }

    [Fact]
    public void Heading_UsesProportionalGlyphWidthsWithoutOverWrappingNarrowText() {
        var options = new PdfOptions {
            PageWidth = 140,
            PageHeight = 160,
            MarginLeft = 30,
            MarginRight = 30,
            MarginTop = 30,
            MarginBottom = 30,
            DefaultFont = PdfStandardFont.Helvetica,
            DefaultFontSize = 10
        };

        byte[] bytes = PdfDoc.Create(options)
            .H3(new string('i', 20))
            .ToBytes();

        using var pdf = PdfDocument.Open(new MemoryStream(bytes));
        var page = pdf.GetPage(1);
        int narrowLineCount = page.Letters
            .Where(letter => letter.Value == "i")
            .GroupBy(letter => Math.Round(letter.StartBaseLine.Y, 1))
            .Count();

        Assert.Equal(1, narrowLineCount);
    }

    [Fact]
    public void RowColumnHeading_UsesProportionalGlyphWidthsForWideWrapping() {
        var options = new PdfOptions {
            PageWidth = 140,
            PageHeight = 160,
            MarginLeft = 30,
            MarginRight = 30,
            MarginTop = 30,
            MarginBottom = 30,
            DefaultFont = PdfStandardFont.Helvetica,
            DefaultFontSize = 10
        };

        byte[] bytes = PdfDoc.Create(options)
            .Compose(document =>
                document.Page(page =>
                    page.Content(content =>
                        content.Row(row =>
                            row.Column(100, column =>
                                column.H3("WWWWWWWW"))))))
            .ToBytes();

        using var pdf = PdfDocument.Open(new MemoryStream(bytes));
        var page = pdf.GetPage(1);
        int wideLineCount = page.Letters
            .Where(letter => letter.Value == "W")
            .GroupBy(letter => Math.Round(letter.StartBaseLine.Y, 1))
            .Count();
        double contentRight = options.PageWidth - options.MarginRight;
        double rightMostWideGlyph = page.Letters
            .Where(letter => letter.Value == "W")
            .Max(letter => letter.EndBaseLine.X);

        Assert.True(wideLineCount > 1, "Expected row-column heading wide glyphs to wrap using their real Helvetica advance instead of an average character width.");
        Assert.InRange(rightMostWideGlyph, double.NegativeInfinity, contentRight + 1);
    }

    [Fact]
    public void RowColumnHeading_UsesProportionalGlyphWidthsWithoutOverWrappingNarrowText() {
        var options = new PdfOptions {
            PageWidth = 140,
            PageHeight = 160,
            MarginLeft = 30,
            MarginRight = 30,
            MarginTop = 30,
            MarginBottom = 30,
            DefaultFont = PdfStandardFont.Helvetica,
            DefaultFontSize = 10
        };

        byte[] bytes = PdfDoc.Create(options)
            .Compose(document =>
                document.Page(page =>
                    page.Content(content =>
                        content.Row(row =>
                            row.Column(100, column =>
                                column.H3(new string('i', 20)))))))
            .ToBytes();

        using var pdf = PdfDocument.Open(new MemoryStream(bytes));
        var page = pdf.GetPage(1);
        int narrowLineCount = page.Letters
            .Where(letter => letter.Value == "i")
            .GroupBy(letter => Math.Round(letter.StartBaseLine.Y, 1))
            .Count();

        Assert.Equal(1, narrowLineCount);
    }

    [Fact]
    public void Footer_RightAlignsUsingProportionalTextWidth() {
        var options = new PdfOptions {
            PageWidth = 320,
            PageHeight = 180,
            MarginLeft = 30,
            MarginRight = 30,
            MarginTop = 30,
            MarginBottom = 30,
            DefaultFont = PdfStandardFont.Helvetica,
            DefaultFontSize = 10,
            ShowPageNumbers = true,
            FooterFormat = "Illi",
            FooterFont = PdfStandardFont.Helvetica,
            FooterFontSize = 10,
            FooterAlign = PdfAlign.Right,
            FooterOffsetY = 12
        };

        byte[] bytes = PdfDoc.Create(options)
            .Paragraph(p => p.Text("Body"))
            .ToBytes();

        using var pdf = PdfDocument.Open(new MemoryStream(bytes));
        var page = pdf.GetPage(1);
        double footerEndX = FindWordEndX(page, "Illi");
        double expectedRightEdge = options.PageWidth - options.MarginRight;

        Assert.InRange(Math.Abs(expectedRightEdge - footerEndX), 0, 5);
    }

    [Fact]
    public void Paragraph_RendersExplicitLineBreaksInsideRichText() {
        var options = new PdfOptions {
            DefaultFont = PdfStandardFont.Helvetica,
            DefaultFontSize = 12
        };

        byte[] bytes = PdfDoc.Create(options)
            .Paragraph(p => p.Bold("Finding").LineBreak().Text("No critical issues detected."))
            .ToBytes();

        using var pdf = PdfDocument.Open(new MemoryStream(bytes));
        var page = pdf.GetPage(1);

        double findingY = FindWordStartY(page, "Finding");
        double noY = FindWordStartY(page, "No");

        Assert.True(findingY > noY + 12, $"Expected the explicit line break to move following text to the next line. Finding y: {findingY:0.##}, No y: {noY:0.##}.");
    }

    [Fact]
    public void Paragraph_UsesConfiguredSpacingBeforeAndAfter() {
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

        byte[] defaultBytes = CreateParagraphSpacingProbe(options, null);
        byte[] spacedBytes = CreateParagraphSpacingProbe(options, new PdfParagraphStyle {
            SpacingBefore = 12,
            SpacingAfter = 18
        });

        using var defaultPdf = PdfDocument.Open(new MemoryStream(defaultBytes));
        using var spacedPdf = PdfDocument.Open(new MemoryStream(spacedBytes));
        var defaultPage = defaultPdf.GetPage(1);
        var spacedPage = spacedPdf.GetPage(1);

        double defaultTargetY = FindWordStartY(defaultPage, "TargetMarker");
        double spacedTargetY = FindWordStartY(spacedPage, "TargetMarker");
        double defaultAfterY = FindWordStartY(defaultPage, "AfterMarker");
        double spacedAfterY = FindWordStartY(spacedPage, "AfterMarker");

        Assert.True(defaultTargetY - spacedTargetY >= 10, $"Expected paragraph spacing before to move target text down. Default y: {defaultTargetY:0.##}, spaced y: {spacedTargetY:0.##}.");
        Assert.True(defaultAfterY - spacedAfterY >= 24, $"Expected paragraph spacing before and after to move following text down. Default y: {defaultAfterY:0.##}, spaced y: {spacedAfterY:0.##}.");
    }

    [Fact]
    public void Paragraph_SuppressesSpacingBeforeAtPageTop() {
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

        byte[] defaultBytes = PdfDoc.Create(options)
            .Paragraph(p => p.Text("TopMarker"))
            .ToBytes();
        byte[] spacedBytes = PdfDoc.Create(options)
            .Paragraph(p => p.Text("TopMarker"), style: new PdfParagraphStyle {
                SpacingBefore = 28,
                SpacingAfter = 0
            })
            .ToBytes();

        using var defaultPdf = PdfDocument.Open(new MemoryStream(defaultBytes));
        using var spacedPdf = PdfDocument.Open(new MemoryStream(spacedBytes));

        double defaultTopY = FindWordStartY(defaultPdf.GetPage(1), "TopMarker");
        double spacedTopY = FindWordStartY(spacedPdf.GetPage(1), "TopMarker");

        Assert.InRange(Math.Abs(defaultTopY - spacedTopY), 0, 1.5);
    }

    [Fact]
    public void Paragraph_SuppressesSpacingBeforeAtRowColumnTop() {
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

        byte[] defaultBytes = PdfDoc.Create(options)
            .Compose(document => document.Page(page => page.Content(content => content.Row(row => row.Column(100, column => column
                .Paragraph(p => p.Text("ColumnTopMarker")))))))
            .ToBytes();
        byte[] spacedBytes = PdfDoc.Create(options)
            .Compose(document => document.Page(page => page.Content(content => content.Row(row => row.Column(100, column => column
                .Paragraph(p => p.Text("ColumnTopMarker"), style: new PdfParagraphStyle {
                    SpacingBefore = 28,
                    SpacingAfter = 0
                }))))))
            .ToBytes();

        using var defaultPdf = PdfDocument.Open(new MemoryStream(defaultBytes));
        using var spacedPdf = PdfDocument.Open(new MemoryStream(spacedBytes));

        double defaultTopY = FindWordStartY(defaultPdf.GetPage(1), "ColumnTopMarker");
        double spacedTopY = FindWordStartY(spacedPdf.GetPage(1), "ColumnTopMarker");

        Assert.InRange(Math.Abs(defaultTopY - spacedTopY), 0, 1.5);
    }

    [Fact]
    public void Spacer_AddsInvisibleVerticalSpaceWithoutExtractedText() {
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

        var paragraphStyle = new PdfParagraphStyle { SpacingBefore = 0, SpacingAfter = 0 };
        byte[] bytes = PdfDoc.Create(options)
            .Paragraph(p => p.Text("BeforeMarker"), style: paragraphStyle)
            .Spacer(24)
            .Paragraph(p => p.Text("AfterMarker"), style: paragraphStyle)
            .ToBytes();

        using var pdf = PdfDocument.Open(new MemoryStream(bytes));
        var page = pdf.GetPage(1);

        Assert.Equal(2, CountTextLines(page));
        Assert.Contains("BeforeMarker", page.Text, StringComparison.Ordinal);
        Assert.Contains("AfterMarker", page.Text, StringComparison.Ordinal);

        double beforeY = FindWordStartY(page, "BeforeMarker");
        double afterY = FindWordStartY(page, "AfterMarker");
        Assert.InRange(beforeY - afterY, 36, 42);
    }

    [Fact]
    public void Spacer_WorksInsideRowColumnFlow() {
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

        var paragraphStyle = new PdfParagraphStyle { SpacingBefore = 0, SpacingAfter = 0 };
        byte[] bytes = PdfDoc.Create(options)
            .Compose(document => document.Page(page => page.Content(content => content.Row(row => row.Column(100, column => column
                .Paragraph(p => p.Text("TopMarker"), style: paragraphStyle)
                .Spacer(20)
                .Paragraph(p => p.Text("BottomMarker"), style: paragraphStyle))))))
            .ToBytes();

        using var pdf = PdfDocument.Open(new MemoryStream(bytes));
        var page = pdf.GetPage(1);

        Assert.Equal(2, CountTextLines(page));

        double topY = FindWordStartY(page, "TopMarker");
        double bottomY = FindWordStartY(page, "BottomMarker");
        Assert.InRange(topY - bottomY, 32, 38);
    }

    [Fact]
    public void Spacer_RejectsInvalidHeights() {
        Assert.Throws<ArgumentOutOfRangeException>(() => PdfDoc.Create().Spacer(-1));
        Assert.Throws<ArgumentOutOfRangeException>(() => PdfDoc.Create().Spacer(double.NaN));
        Assert.Throws<ArgumentOutOfRangeException>(() => PdfDoc.Create().Spacer(double.PositiveInfinity));
    }

    [Fact]
    public void Paragraph_UsesConfiguredLineHeight() {
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

        byte[] defaultBytes = CreateParagraphLineHeightProbe(options, null);
        byte[] looseBytes = CreateParagraphLineHeightProbe(options, new PdfParagraphStyle {
            LineHeight = 2.0,
            SpacingAfter = 0
        });

        using var defaultPdf = PdfDocument.Open(new MemoryStream(defaultBytes));
        using var loosePdf = PdfDocument.Open(new MemoryStream(looseBytes));
        var defaultPage = defaultPdf.GetPage(1);
        var loosePage = loosePdf.GetPage(1);

        double defaultFirstY = FindWordStartY(defaultPage, "FirstLine");
        double defaultSecondY = FindWordStartY(defaultPage, "SecondLine");
        double defaultThirdY = FindWordStartY(defaultPage, "ThirdLine");
        double looseFirstY = FindWordStartY(loosePage, "FirstLine");
        double looseSecondY = FindWordStartY(loosePage, "SecondLine");
        double looseThirdY = FindWordStartY(loosePage, "ThirdLine");

        double defaultGapOne = defaultFirstY - defaultSecondY;
        double defaultGapTwo = defaultSecondY - defaultThirdY;
        double looseGapOne = looseFirstY - looseSecondY;
        double looseGapTwo = looseSecondY - looseThirdY;

        Assert.True(looseGapOne - defaultGapOne >= 5, $"Expected configured line height to increase the first line gap. Default gap: {defaultGapOne:0.##}, loose gap: {looseGapOne:0.##}.");
        Assert.True(looseGapTwo - defaultGapTwo >= 5, $"Expected configured line height to increase the second line gap. Default gap: {defaultGapTwo:0.##}, loose gap: {looseGapTwo:0.##}.");
    }

    [Fact]
    public void Paragraph_UsesConfiguredHorizontalIndents() {
        var options = new PdfOptions {
            PageWidth = 300,
            PageHeight = 280,
            MarginLeft = 30,
            MarginRight = 30,
            MarginTop = 30,
            MarginBottom = 30,
            DefaultFont = PdfStandardFont.Helvetica,
            DefaultFontSize = 10
        };

        byte[] defaultBytes = CreateParagraphIndentProbe(options, null);
        byte[] indentedBytes = CreateParagraphIndentProbe(options, new PdfParagraphStyle {
            LeftIndent = 24,
            RightIndent = 90,
            SpacingAfter = 0
        });

        using var defaultPdf = PdfDocument.Open(new MemoryStream(defaultBytes));
        using var indentedPdf = PdfDocument.Open(new MemoryStream(indentedBytes));
        var defaultPage = defaultPdf.GetPage(1);
        var indentedPage = indentedPdf.GetPage(1);

        double defaultX = FindWordStartX(defaultPage, "IndentedMarker");
        double indentedX = FindWordStartX(indentedPage, "IndentedMarker");
        int defaultLineCount = CountTextLines(defaultPage);
        int indentedLineCount = CountTextLines(indentedPage);

        Assert.True(indentedX - defaultX >= 22, $"Expected left indent to move paragraph text right. Default x: {defaultX:0.##}, indented x: {indentedX:0.##}.");
        Assert.True(indentedLineCount > defaultLineCount, $"Expected right indent to reduce text width and increase wrapping. Default lines: {defaultLineCount}, indented lines: {indentedLineCount}.");
    }

    [Fact]
    public void Paragraph_UsesConfiguredFirstLineIndent() {
        var options = new PdfOptions {
            PageWidth = 300,
            PageHeight = 220,
            MarginLeft = 30,
            MarginRight = 30,
            MarginTop = 30,
            MarginBottom = 30,
            DefaultFont = PdfStandardFont.Helvetica,
            DefaultFontSize = 10
        };

        byte[] bytes = PdfDoc.Create(options)
            .Paragraph(p => p
                .Text("FirstIndentMarker")
                .LineBreak()
                .Text("SecondIndentMarker"), style: new PdfParagraphStyle {
                    FirstLineIndent = 24,
                    SpacingAfter = 0
                })
            .ToBytes();

        using var pdf = PdfDocument.Open(new MemoryStream(bytes));
        var page = pdf.GetPage(1);

        double firstX = FindWordStartX(page, "FirstIndentMarker");
        double secondX = FindWordStartX(page, "SecondIndentMarker");

        Assert.True(firstX - secondX >= 22, $"Expected first line indent to move only the first line right. First x: {firstX:0.##}, second x: {secondX:0.##}.");
    }

    [Fact]
    public void Paragraph_UsesDefaultParagraphStyleWhenStyleIsNotProvided() {
        var options = new PdfOptions {
            PageWidth = 300,
            PageHeight = 220,
            MarginLeft = 30,
            MarginRight = 30,
            MarginTop = 30,
            MarginBottom = 30,
            DefaultFont = PdfStandardFont.Helvetica,
            DefaultFontSize = 10,
            DefaultParagraphStyle = new PdfParagraphStyle {
                FirstLineIndent = 24,
                SpacingAfter = 0
            }
        };

        byte[] bytes = PdfDoc.Create(options)
            .Paragraph(p => p
                .Text("DefaultFirstIndent")
                .LineBreak()
                .Text("DefaultSecondIndent"))
            .ToBytes();

        using var pdf = PdfDocument.Open(new MemoryStream(bytes));
        var page = pdf.GetPage(1);

        double firstX = FindWordStartX(page, "DefaultFirstIndent");
        double secondX = FindWordStartX(page, "DefaultSecondIndent");

        Assert.True(firstX - secondX >= 22, $"Expected default paragraph style to indent only the first line. First x: {firstX:0.##}, second x: {secondX:0.##}.");
    }

    [Fact]
    public void PdfDoc_DefaultParagraphStyleAppliesToFollowingParagraphsAndSnapshotsInput() {
        var options = new PdfOptions {
            PageWidth = 300,
            PageHeight = 220,
            MarginLeft = 30,
            MarginRight = 30,
            MarginTop = 30,
            MarginBottom = 30,
            DefaultFont = PdfStandardFont.Helvetica,
            DefaultFontSize = 10
        };
        var style = new PdfParagraphStyle {
            FirstLineIndent = 24,
            SpacingAfter = 0
        };

        byte[] bytes = PdfDoc.Create(options)
            .DefaultParagraphStyle(style)
            .Paragraph(p => p
                .Text("FluentDefaultFirst")
                .LineBreak()
                .Text("FluentDefaultSecond"))
            .ToBytes();

        style.FirstLineIndent = 0;

        using var pdf = PdfDocument.Open(new MemoryStream(bytes));
        var page = pdf.GetPage(1);

        double firstX = FindWordStartX(page, "FluentDefaultFirst");
        double secondX = FindWordStartX(page, "FluentDefaultSecond");

        Assert.True(firstX - secondX >= 22, $"Expected fluent default paragraph style to indent only the first line and snapshot caller input. First x: {firstX:0.##}, second x: {secondX:0.##}.");
    }

    [Fact]
    public void PdfDoc_DefaultParagraphStyleRejectsNull() {
        Assert.Throws<ArgumentNullException>(() => PdfDoc.Create().DefaultParagraphStyle(null!));
    }

    [Fact]
    public void Paragraph_ExplicitStyleOverridesDefaultParagraphStyle() {
        var options = new PdfOptions {
            PageWidth = 300,
            PageHeight = 220,
            MarginLeft = 30,
            MarginRight = 30,
            MarginTop = 30,
            MarginBottom = 30,
            DefaultFont = PdfStandardFont.Helvetica,
            DefaultFontSize = 10,
            DefaultParagraphStyle = new PdfParagraphStyle {
                LeftIndent = 40,
                SpacingAfter = 0
            }
        };

        byte[] bytes = PdfDoc.Create(options)
            .Paragraph(p => p.Text("ExplicitMarker"), style: new PdfParagraphStyle {
                SpacingAfter = 0
            })
            .ToBytes();

        using var pdf = PdfDocument.Open(new MemoryStream(bytes));
        var page = pdf.GetPage(1);

        double markerX = FindWordStartX(page, "ExplicitMarker");

        Assert.InRange(markerX, options.MarginLeft - 1, options.MarginLeft + 3);
    }

    [Fact]
    public void Paragraph_UsesConfiguredHangingIndent() {
        var options = new PdfOptions {
            PageWidth = 300,
            PageHeight = 220,
            MarginLeft = 30,
            MarginRight = 30,
            MarginTop = 30,
            MarginBottom = 30,
            DefaultFont = PdfStandardFont.Helvetica,
            DefaultFontSize = 10
        };

        byte[] bytes = PdfDoc.Create(options)
            .Paragraph(p => p
                .Text("HangingFirst")
                .LineBreak()
                .Text("HangingSecond"), style: new PdfParagraphStyle {
                    LeftIndent = 24,
                    FirstLineIndent = -24,
                    SpacingAfter = 0
                })
            .ToBytes();

        using var pdf = PdfDocument.Open(new MemoryStream(bytes));
        var page = pdf.GetPage(1);

        double firstX = FindWordStartX(page, "HangingFirst");
        double secondX = FindWordStartX(page, "HangingSecond");

        Assert.True(secondX - firstX >= 22, $"Expected hanging indent to move following lines right of the first line. First x: {firstX:0.##}, second x: {secondX:0.##}.");
    }

    [Fact]
    public void RowColumnParagraph_UsesConfiguredFirstLineIndent() {
        var options = new PdfOptions {
            PageWidth = 300,
            PageHeight = 220,
            MarginLeft = 30,
            MarginRight = 30,
            MarginTop = 30,
            MarginBottom = 30,
            DefaultFont = PdfStandardFont.Helvetica,
            DefaultFontSize = 10
        };

        byte[] bytes = PdfDoc.Create(options)
            .Compose(document =>
                document.Page(page =>
                    page.Content(content =>
                        content.Row(row =>
                            row.Column(100, column => column
                                .Paragraph(p => p
                                    .Text("ColumnFirstIndent")
                                    .LineBreak()
                                    .Text("ColumnSecondIndent"), style: new PdfParagraphStyle {
                                        FirstLineIndent = 24,
                                        SpacingAfter = 0
                                    }))))))
            .ToBytes();

        using var pdf = PdfDocument.Open(new MemoryStream(bytes));
        var page = pdf.GetPage(1);

        double firstX = FindWordStartX(page, "ColumnFirstIndent");
        double secondX = FindWordStartX(page, "ColumnSecondIndent");

        Assert.True(firstX - secondX >= 22, $"Expected row column first line indent to move only the first line right. First x: {firstX:0.##}, second x: {secondX:0.##}.");
    }

    [Fact]
    public void RowColumnParagraph_UsesDefaultParagraphStyleWhenStyleIsNotProvided() {
        var options = new PdfOptions {
            PageWidth = 300,
            PageHeight = 220,
            MarginLeft = 30,
            MarginRight = 30,
            MarginTop = 30,
            MarginBottom = 30,
            DefaultFont = PdfStandardFont.Helvetica,
            DefaultFontSize = 10,
            DefaultParagraphStyle = new PdfParagraphStyle {
                FirstLineIndent = 24,
                SpacingAfter = 0
            }
        };

        byte[] bytes = PdfDoc.Create(options)
            .Compose(document =>
                document.Page(page =>
                    page.Content(content =>
                        content.Row(row =>
                            row.Column(100, column => column
                                .Paragraph(p => p
                                    .Text("ColumnDefaultFirst")
                                    .LineBreak()
                                    .Text("ColumnDefaultSecond")))))))
            .ToBytes();

        using var pdf = PdfDocument.Open(new MemoryStream(bytes));
        var page = pdf.GetPage(1);

        double firstX = FindWordStartX(page, "ColumnDefaultFirst");
        double secondX = FindWordStartX(page, "ColumnDefaultSecond");

        Assert.True(firstX - secondX >= 22, $"Expected row column default paragraph style to indent only the first line. First x: {firstX:0.##}, second x: {secondX:0.##}.");
    }

    [Fact]
    public void Paragraph_SplitsLongContentAcrossPagesWithoutCrossingBottomMargin() {
        var options = new PdfOptions {
            PageWidth = 320,
            PageHeight = 170,
            MarginLeft = 30,
            MarginRight = 30,
            MarginTop = 30,
            MarginBottom = 30,
            DefaultFont = PdfStandardFont.Helvetica,
            DefaultFontSize = 9
        };

        string longText = string.Join(" ", Enumerable.Range(1, 180).Select(i => "segment" + i.ToString("000")));

        byte[] bytes = PdfDoc.Create(options)
            .Paragraph(p => p.Text(longText), style: new PdfParagraphStyle {
                LineHeight = 1.3,
                SpacingAfter = 0
            })
            .ToBytes();

        using var pdf = PdfDocument.Open(new MemoryStream(bytes));
        Assert.True(pdf.NumberOfPages > 1, "Expected a long rich paragraph to continue onto another page.");

        for (int pageNumber = 1; pageNumber <= pdf.NumberOfPages; pageNumber++) {
            var page = pdf.GetPage(pageNumber);
            double bottomMost = page.Letters
                .Where(letter => !string.IsNullOrWhiteSpace(letter.Value))
                .Min(letter => letter.StartBaseLine.Y);
            Assert.True(bottomMost >= options.MarginBottom - 2, $"Expected paragraph text to stay above the bottom margin on page {pageNumber}.");
        }

        Assert.Contains("segment001", pdf.GetPage(1).Text);
        Assert.Contains("segment180", pdf.GetPage(pdf.NumberOfPages).Text);
    }

    [Fact]
    public void Paragraph_KeepTogetherMovesWholeParagraphToNextPage() {
        var options = new PdfOptions {
            PageWidth = 260,
            PageHeight = 170,
            MarginLeft = 30,
            MarginRight = 30,
            MarginTop = 30,
            MarginBottom = 30,
            DefaultFont = PdfStandardFont.Helvetica,
            DefaultFontSize = 10
        };

        byte[] bytes = PdfDoc.Create(options)
            .Paragraph(p => p.Text("IntroMarker"), style: new PdfParagraphStyle {
                SpacingAfter = 70
            })
            .Paragraph(p => p
                .Text("KeepFirst")
                .LineBreak()
                .Text("KeepMiddle")
                .LineBreak()
                .Text("KeepLast"), style: new PdfParagraphStyle {
                    KeepTogether = true,
                    SpacingAfter = 0
                })
            .ToBytes();

        using var pdf = PdfDocument.Open(new MemoryStream(bytes));

        Assert.Equal(2, pdf.NumberOfPages);
        Assert.Contains("IntroMarker", pdf.GetPage(1).Text);
        Assert.DoesNotContain("KeepFirst", pdf.GetPage(1).Text);
        Assert.Contains("KeepFirst", pdf.GetPage(2).Text);
        Assert.Contains("KeepLast", pdf.GetPage(2).Text);
    }

    [Fact]
    public void List_KeepTogetherMovesWholeBulletListToNextPage() {
        var options = new PdfOptions {
            PageWidth = 260,
            PageHeight = 170,
            MarginLeft = 30,
            MarginRight = 30,
            MarginTop = 30,
            MarginBottom = 30,
            DefaultFont = PdfStandardFont.Helvetica,
            DefaultFontSize = 10
        };

        byte[] bytes = PdfDoc.Create(options)
            .Paragraph(p => p.Text("IntroMarker"), style: new PdfParagraphStyle {
                SpacingAfter = 70
            })
            .Bullets(new[] {
                "KeepListFirst",
                "KeepListMiddle",
                "KeepListLast"
            }, style: new PdfListStyle {
                KeepTogether = true,
                SpacingAfter = 0
            })
            .ToBytes();

        using var pdf = PdfDocument.Open(new MemoryStream(bytes));

        Assert.Equal(2, pdf.NumberOfPages);
        Assert.Contains("IntroMarker", pdf.GetPage(1).Text);
        Assert.DoesNotContain("KeepListFirst", pdf.GetPage(1).Text);
        Assert.Contains("KeepListFirst", pdf.GetPage(2).Text);
        Assert.Contains("KeepListLast", pdf.GetPage(2).Text);
    }

    [Fact]
    public void List_KeepWithNextMovesListWithFollowingParagraph() {
        var options = new PdfOptions {
            PageWidth = 260,
            PageHeight = 170,
            MarginLeft = 30,
            MarginRight = 30,
            MarginTop = 30,
            MarginBottom = 30,
            DefaultFont = PdfStandardFont.Helvetica,
            DefaultFontSize = 10
        };

        byte[] bytes = PdfDoc.Create(options)
            .Paragraph(p => p.Text("IntroMarker"), style: new PdfParagraphStyle {
                SpacingAfter = 64
            })
            .Bullets(new[] {
                "KeepListFirst",
                "KeepListSecond"
            }, style: new PdfListStyle {
                KeepWithNext = true,
                ItemSpacing = 0,
                SpacingAfter = 0
            })
            .Paragraph(p => p.Text("FollowingListBody"))
            .ToBytes();

        using var pdf = PdfDocument.Open(new MemoryStream(bytes));

        Assert.Equal(2, pdf.NumberOfPages);
        Assert.Contains("IntroMarker", pdf.GetPage(1).Text);
        Assert.DoesNotContain("KeepListFirst", pdf.GetPage(1).Text);
        Assert.Contains("KeepListFirst", pdf.GetPage(2).Text);
        Assert.Contains("KeepListSecond", pdf.GetPage(2).Text);
        Assert.Contains("FollowingListBody", pdf.GetPage(2).Text);
    }

    [Fact]
    public void Paragraph_KeepWithNextMovesParagraphWithFollowingParagraph() {
        var options = new PdfOptions {
            PageWidth = 260,
            PageHeight = 170,
            MarginLeft = 30,
            MarginRight = 30,
            MarginTop = 30,
            MarginBottom = 30,
            DefaultFont = PdfStandardFont.Helvetica,
            DefaultFontSize = 10
        };

        byte[] bytes = PdfDoc.Create(options)
            .Paragraph(p => p.Text("IntroMarker"), style: new PdfParagraphStyle {
                SpacingAfter = 70
            })
            .Paragraph(p => p.Text("KeepWithNextLabel"), style: new PdfParagraphStyle {
                KeepWithNext = true
            })
            .Paragraph(p => p.Text("FollowingBody"))
            .ToBytes();

        using var pdf = PdfDocument.Open(new MemoryStream(bytes));

        Assert.Equal(2, pdf.NumberOfPages);
        Assert.Contains("IntroMarker", pdf.GetPage(1).Text);
        Assert.DoesNotContain("KeepWithNextLabel", pdf.GetPage(1).Text);
        Assert.Contains("KeepWithNextLabel", pdf.GetPage(2).Text);
        Assert.Contains("FollowingBody", pdf.GetPage(2).Text);
    }

    [Fact]
    public void Paragraph_KeepWithNextMovesParagraphWithFollowingList() {
        var options = new PdfOptions {
            PageWidth = 260,
            PageHeight = 170,
            MarginLeft = 30,
            MarginRight = 30,
            MarginTop = 30,
            MarginBottom = 30,
            DefaultFont = PdfStandardFont.Helvetica,
            DefaultFontSize = 10
        };

        byte[] bytes = PdfDoc.Create(options)
            .Paragraph(p => p.Text("IntroMarker"), style: new PdfParagraphStyle {
                SpacingAfter = 70
            })
            .Paragraph(p => p.Text("KeepWithListLabel"), style: new PdfParagraphStyle {
                KeepWithNext = true
            })
            .Bullets(new[] { "FollowingBullet" })
            .ToBytes();

        using var pdf = PdfDocument.Open(new MemoryStream(bytes));

        Assert.Equal(2, pdf.NumberOfPages);
        Assert.Contains("IntroMarker", pdf.GetPage(1).Text);
        Assert.DoesNotContain("KeepWithListLabel", pdf.GetPage(1).Text);
        Assert.Contains("KeepWithListLabel", pdf.GetPage(2).Text);
        Assert.Contains("FollowingBullet", pdf.GetPage(2).Text);
    }

    [Fact]
    public void Paragraph_KeepWithNextMovesParagraphWithFollowingTable() {
        var options = new PdfOptions {
            PageWidth = 260,
            PageHeight = 170,
            MarginLeft = 30,
            MarginRight = 30,
            MarginTop = 30,
            MarginBottom = 30,
            DefaultFont = PdfStandardFont.Helvetica,
            DefaultFontSize = 10
        };

        byte[] bytes = PdfDoc.Create(options)
            .Paragraph(p => p.Text("IntroMarker"), style: new PdfParagraphStyle {
                SpacingAfter = 70
            })
            .Paragraph(p => p.Text("KeepWithTableLabel"), style: new PdfParagraphStyle {
                KeepWithNext = true
            })
            .Table(new[] {
                new[] { "FollowingTableCell", "Value" }
            })
            .ToBytes();

        using var pdf = PdfDocument.Open(new MemoryStream(bytes));

        Assert.Equal(2, pdf.NumberOfPages);
        Assert.Contains("IntroMarker", pdf.GetPage(1).Text);
        Assert.DoesNotContain("KeepWithTableLabel", pdf.GetPage(1).Text);
        Assert.Contains("KeepWithTableLabel", pdf.GetPage(2).Text);
        Assert.Contains("FollowingTableCell", pdf.GetPage(2).Text);
    }

    [Fact]
    public void Paragraph_KeepWithNextUsesConfiguredTableColumnWidthsForFollowingTable() {
        var style = TableStyles.Minimal();
        style.CellPaddingX = 4;
        style.CellPaddingY = 3;
        style.ColumnWidthPoints = new List<double?> { 44, null };

        var options = new PdfOptions {
            PageWidth = 260,
            PageHeight = 170,
            MarginLeft = 30,
            MarginRight = 30,
            MarginTop = 30,
            MarginBottom = 30,
            DefaultFont = PdfStandardFont.Helvetica,
            DefaultFontSize = 10
        };

        byte[] bytes = PdfDoc.Create(options)
            .Paragraph(p => p.Text("IntroMarker"), style: new PdfParagraphStyle {
                SpacingAfter = 35
            })
            .Paragraph(p => p.Text("KeepWithNarrowTableLabel"), style: new PdfParagraphStyle {
                KeepWithNext = true
            })
            .Table(new[] {
                new[] { "aa bb cc dd ee ff gg hh ii jj kk ll", "Value" }
            }, style: style)
            .ToBytes();

        using var pdf = PdfDocument.Open(new MemoryStream(bytes));

        Assert.Equal(2, pdf.NumberOfPages);
        Assert.Contains("IntroMarker", pdf.GetPage(1).Text);
        Assert.DoesNotContain("KeepWithNarrowTableLabel", pdf.GetPage(1).Text);
        Assert.Contains("KeepWithNarrowTableLabel", pdf.GetPage(2).Text);
        Assert.Contains("Value", pdf.GetPage(2).Text);
    }

    [Fact]
    public void Paragraph_KeepWithNextMovesParagraphWithFollowingRow() {
        var options = new PdfOptions {
            PageWidth = 260,
            PageHeight = 170,
            MarginLeft = 30,
            MarginRight = 30,
            MarginTop = 30,
            MarginBottom = 30,
            DefaultFont = PdfStandardFont.Helvetica,
            DefaultFontSize = 10,
            DefaultRowStyle = new PdfRowStyle {
                Gap = 18
            }
        };

        byte[] bytes = PdfDoc.Create(options)
            .Compose(document =>
                document.Page(page =>
                    page.Content(content => {
                        content.Column(column => {
                            column.Item().Paragraph(p => p.Text("IntroMarker"), style: new PdfParagraphStyle {
                                SpacingAfter = 70
                            });
                            column.Item().Paragraph(p => p.Text("KeepWithRowLabel"), style: new PdfParagraphStyle {
                                KeepWithNext = true
                            });
                        });
                        content.Row(row => row
                            .Column(50, column => column.Paragraph(p => p.Text("RowLeftBody")))
                            .Column(50, column => column.Paragraph(p => p.Text("RowRightBody"))));
                    })))
            .ToBytes();

        using var pdf = PdfDocument.Open(new MemoryStream(bytes));

        Assert.Equal(2, pdf.NumberOfPages);
        Assert.Contains("IntroMarker", pdf.GetPage(1).Text);
        Assert.DoesNotContain("KeepWithRowLabel", pdf.GetPage(1).Text);
        Assert.Contains("KeepWithRowLabel", pdf.GetPage(2).Text);
        Assert.Contains("RowLeftBody", pdf.GetPage(2).Text);
        Assert.Contains("RowRightBody", pdf.GetPage(2).Text);
    }

    [Fact]
    public void Heading_KeepsWithFollowingParagraphInTopLevelFlow() {
        var options = new PdfOptions {
            PageWidth = 260,
            PageHeight = 170,
            MarginLeft = 30,
            MarginRight = 30,
            MarginTop = 30,
            MarginBottom = 30,
            DefaultFont = PdfStandardFont.Helvetica,
            DefaultFontSize = 10
        };

        byte[] bytes = PdfDoc.Create(options)
            .Paragraph(p => p.Text("IntroMarker"), style: new PdfParagraphStyle {
                SpacingAfter = 70
            })
            .H3("SignalHeading")
            .Paragraph(p => p.Text("SignalBody"))
            .ToBytes();

        using var pdf = PdfDocument.Open(new MemoryStream(bytes));

        Assert.Equal(2, pdf.NumberOfPages);
        Assert.Contains("IntroMarker", pdf.GetPage(1).Text);
        Assert.DoesNotContain("SignalHeading", pdf.GetPage(1).Text);
        Assert.Contains("SignalHeading", pdf.GetPage(2).Text);
        Assert.Contains("SignalBody", pdf.GetPage(2).Text);
    }

    [Fact]
    public void Heading_KeepsWithFollowingPanelInTopLevelFlow() {
        var options = new PdfOptions {
            PageWidth = 260,
            PageHeight = 170,
            MarginLeft = 30,
            MarginRight = 30,
            MarginTop = 30,
            MarginBottom = 30,
            DefaultFont = PdfStandardFont.Helvetica,
            DefaultFontSize = 10
        };

        byte[] bytes = PdfDoc.Create(options)
            .Paragraph(p => p.Text("IntroMarker"), style: new PdfParagraphStyle {
                SpacingAfter = 70
            })
            .H3("PanelSignalHeading")
            .PanelParagraph(p => p.Text("PanelSignalBody"), new PanelStyle {
                PaddingY = 5,
                SpacingAfter = 0
            })
            .ToBytes();

        using var pdf = PdfDocument.Open(new MemoryStream(bytes));

        Assert.Equal(2, pdf.NumberOfPages);
        Assert.Contains("IntroMarker", pdf.GetPage(1).Text);
        Assert.DoesNotContain("PanelSignalHeading", pdf.GetPage(1).Text);
        Assert.Contains("PanelSignalHeading", pdf.GetPage(2).Text);
        Assert.Contains("PanelSignalBody", pdf.GetPage(2).Text);
    }

    [Fact]
    public void PanelParagraph_KeepWithNextMovesPanelWithFollowingParagraph() {
        var options = new PdfOptions {
            PageWidth = 260,
            PageHeight = 170,
            MarginLeft = 30,
            MarginRight = 30,
            MarginTop = 30,
            MarginBottom = 30,
            DefaultFont = PdfStandardFont.Helvetica,
            DefaultFontSize = 10
        };

        byte[] bytes = PdfDoc.Create(options)
            .Paragraph(p => p.Text("IntroMarker"), style: new PdfParagraphStyle {
                SpacingAfter = 61
            })
            .PanelParagraph(p => p.Text("PanelKeepWithNext"), new PanelStyle {
                KeepWithNext = true,
                PaddingY = 5,
                SpacingAfter = 0
            })
            .Paragraph(p => p.Text("FollowingPanelBody"))
            .ToBytes();

        using var pdf = PdfDocument.Open(new MemoryStream(bytes));

        Assert.Equal(2, pdf.NumberOfPages);
        Assert.Contains("IntroMarker", pdf.GetPage(1).Text);
        Assert.DoesNotContain("PanelKeepWithNext", pdf.GetPage(1).Text);
        Assert.Contains("PanelKeepWithNext", pdf.GetPage(2).Text);
        Assert.Contains("FollowingPanelBody", pdf.GetPage(2).Text);
    }

    [Fact]
    public void Heading_KeepsWithFollowingTableInTopLevelFlow() {
        var options = new PdfOptions {
            PageWidth = 260,
            PageHeight = 170,
            MarginLeft = 30,
            MarginRight = 30,
            MarginTop = 30,
            MarginBottom = 30,
            DefaultFont = PdfStandardFont.Helvetica,
            DefaultFontSize = 10
        };

        byte[] bytes = PdfDoc.Create(options)
            .Paragraph(p => p.Text("IntroMarker"), style: new PdfParagraphStyle {
                SpacingAfter = 70
            })
            .H3("TableSignalHeading")
            .Table(new[] {
                new[] { "TableSignalCell", "Value" }
            })
            .ToBytes();

        using var pdf = PdfDocument.Open(new MemoryStream(bytes));

        Assert.Equal(2, pdf.NumberOfPages);
        Assert.Contains("IntroMarker", pdf.GetPage(1).Text);
        Assert.DoesNotContain("TableSignalHeading", pdf.GetPage(1).Text);
        Assert.Contains("TableSignalHeading", pdf.GetPage(2).Text);
        Assert.Contains("TableSignalCell", pdf.GetPage(2).Text);
    }

    [Fact]
    public void Heading_KeepsWithFollowingRowInTopLevelFlow() {
        var options = new PdfOptions {
            PageWidth = 260,
            PageHeight = 170,
            MarginLeft = 30,
            MarginRight = 30,
            MarginTop = 30,
            MarginBottom = 30,
            DefaultFont = PdfStandardFont.Helvetica,
            DefaultFontSize = 10,
            DefaultRowStyle = new PdfRowStyle {
                Gap = 18
            }
        };

        byte[] bytes = PdfDoc.Create(options)
            .Compose(document =>
                document.Page(page =>
                    page.Content(content => {
                        content.Column(column => {
                            column.Item().Paragraph(p => p.Text("IntroMarker"), style: new PdfParagraphStyle {
                                SpacingAfter = 70
                            });
                            column.Item().H3("RowSignalHeading");
                        });
                        content.Row(row => row
                            .Column(50, column => column.Paragraph(p => p.Text("RowSignalLeft")))
                            .Column(50, column => column.Paragraph(p => p.Text("RowSignalRight"))));
                    })))
            .ToBytes();

        using var pdf = PdfDocument.Open(new MemoryStream(bytes));

        Assert.Equal(2, pdf.NumberOfPages);
        Assert.Contains("IntroMarker", pdf.GetPage(1).Text);
        Assert.DoesNotContain("RowSignalHeading", pdf.GetPage(1).Text);
        Assert.Contains("RowSignalHeading", pdf.GetPage(2).Text);
        Assert.Contains("RowSignalLeft", pdf.GetPage(2).Text);
        Assert.Contains("RowSignalRight", pdf.GetPage(2).Text);
    }

    [Fact]
    public void Paragraph_WidowControlAvoidsSingleLineAtPageBottom() {
        var options = new PdfOptions {
            PageWidth = 260,
            PageHeight = 170,
            MarginLeft = 30,
            MarginRight = 30,
            MarginTop = 30,
            MarginBottom = 30,
            DefaultFont = PdfStandardFont.Helvetica,
            DefaultFontSize = 10
        };

        byte[] bytes = PdfDoc.Create(options)
            .Paragraph(p => p.Text("IntroMarker"), style: new PdfParagraphStyle {
                SpacingAfter = 70
            })
            .Paragraph(p => p
                .Text("WidowFirst")
                .LineBreak()
                .Text("WidowSecond")
                .LineBreak()
                .Text("WidowThird"), style: new PdfParagraphStyle {
                    WidowControl = true,
                    SpacingAfter = 0
                })
            .ToBytes();

        using var pdf = PdfDocument.Open(new MemoryStream(bytes));

        Assert.Equal(2, pdf.NumberOfPages);
        Assert.Contains("IntroMarker", pdf.GetPage(1).Text);
        Assert.DoesNotContain("WidowFirst", pdf.GetPage(1).Text);
        Assert.Contains("WidowFirst", pdf.GetPage(2).Text);
        Assert.Contains("WidowThird", pdf.GetPage(2).Text);
    }

    [Fact]
    public void RowColumnParagraph_KeepTogetherMovesWholeParagraphToNextPage() {
        var options = new PdfOptions {
            PageWidth = 260,
            PageHeight = 170,
            MarginLeft = 30,
            MarginRight = 30,
            MarginTop = 30,
            MarginBottom = 30,
            DefaultFont = PdfStandardFont.Helvetica,
            DefaultFontSize = 10
        };

        byte[] bytes = PdfDoc.Create(options)
            .Paragraph(p => p.Text("IntroMarker"), style: new PdfParagraphStyle {
                SpacingAfter = 70
            })
            .Compose(document =>
                document.Page(page =>
                    page.Content(content =>
                        content.Row(row =>
                            row.Column(100, column => column
                                .Paragraph(p => p
                                    .Text("ColumnKeepFirst")
                                    .LineBreak()
                                    .Text("ColumnKeepMiddle")
                                    .LineBreak()
                                    .Text("ColumnKeepLast"), style: new PdfParagraphStyle {
                                        KeepTogether = true,
                                        SpacingAfter = 0
                                    }))))))
            .ToBytes();

        using var pdf = PdfDocument.Open(new MemoryStream(bytes));

        Assert.Equal(2, pdf.NumberOfPages);
        Assert.Contains("IntroMarker", pdf.GetPage(1).Text);
        Assert.DoesNotContain("ColumnKeepFirst", pdf.GetPage(1).Text);
        Assert.Contains("ColumnKeepFirst", pdf.GetPage(2).Text);
        Assert.Contains("ColumnKeepLast", pdf.GetPage(2).Text);
    }

    [Fact]
    public void RowColumnList_KeepTogetherMovesWholeBulletListToNextPage() {
        var options = new PdfOptions {
            PageWidth = 260,
            PageHeight = 170,
            MarginLeft = 30,
            MarginRight = 30,
            MarginTop = 30,
            MarginBottom = 30,
            DefaultFont = PdfStandardFont.Helvetica,
            DefaultFontSize = 10
        };

        byte[] bytes = PdfDoc.Create(options)
            .Paragraph(p => p.Text("IntroMarker"), style: new PdfParagraphStyle {
                SpacingAfter = 70
            })
            .Compose(document =>
                document.Page(page =>
                    page.Content(content =>
                        content.Row(row =>
                            row.Column(100, column => column
                                .Bullets(new[] {
                                    "ColumnKeepListFirst",
                                    "ColumnKeepListMiddle",
                                    "ColumnKeepListLast"
                                }, style: new PdfListStyle {
                                    KeepTogether = true,
                                    SpacingAfter = 0
                                }))))))
            .ToBytes();

        using var pdf = PdfDocument.Open(new MemoryStream(bytes));

        Assert.Equal(2, pdf.NumberOfPages);
        Assert.Contains("IntroMarker", pdf.GetPage(1).Text);
        Assert.DoesNotContain("ColumnKeepListFirst", pdf.GetPage(1).Text);
        Assert.Contains("ColumnKeepListFirst", pdf.GetPage(2).Text);
        Assert.Contains("ColumnKeepListLast", pdf.GetPage(2).Text);
    }

    [Fact]
    public void RowColumnList_KeepWithNextMovesNumberedListWithFollowingParagraph() {
        var options = new PdfOptions {
            PageWidth = 260,
            PageHeight = 170,
            MarginLeft = 30,
            MarginRight = 30,
            MarginTop = 30,
            MarginBottom = 30,
            DefaultFont = PdfStandardFont.Helvetica,
            DefaultFontSize = 10
        };

        byte[] bytes = PdfDoc.Create(options)
            .Paragraph(p => p.Text("IntroMarker"), style: new PdfParagraphStyle {
                SpacingAfter = 60
            })
            .Compose(document =>
                document.Page(page =>
                    page.Content(content =>
                        content.Row(row =>
                            row.Column(100, column => column
                                .Numbered(new[] {
                                    "ColumnKeepNumberOne",
                                    "ColumnKeepNumberTwo"
                                }, style: new PdfListStyle {
                                    KeepWithNext = true,
                                    ItemSpacing = 0,
                                    SpacingAfter = 0
                                })
                                .Paragraph(p => p.Text("ColumnFollowingListBody")))))))
            .ToBytes();

        using var pdf = PdfDocument.Open(new MemoryStream(bytes));

        Assert.Equal(2, pdf.NumberOfPages);
        Assert.Contains("IntroMarker", pdf.GetPage(1).Text);
        Assert.DoesNotContain("ColumnKeepNumberOne", pdf.GetPage(1).Text);
        Assert.Contains("ColumnKeepNumberOne", pdf.GetPage(2).Text);
        Assert.Contains("ColumnKeepNumberTwo", pdf.GetPage(2).Text);
        Assert.Contains("ColumnFollowingListBody", pdf.GetPage(2).Text);
    }

    [Fact]
    public void RowColumnParagraph_KeepWithNextMovesParagraphWithFollowingParagraph() {
        var options = new PdfOptions {
            PageWidth = 260,
            PageHeight = 170,
            MarginLeft = 30,
            MarginRight = 30,
            MarginTop = 30,
            MarginBottom = 30,
            DefaultFont = PdfStandardFont.Helvetica,
            DefaultFontSize = 10
        };

        byte[] bytes = PdfDoc.Create(options)
            .Paragraph(p => p.Text("IntroMarker"), style: new PdfParagraphStyle {
                SpacingAfter = 70
            })
            .Compose(document =>
                document.Page(page =>
                    page.Content(content =>
                        content.Row(row =>
                            row.Column(100, column => column
                                .Paragraph(p => p.Text("ColumnKeepWithNextLabel"), style: new PdfParagraphStyle {
                                    KeepWithNext = true
                                })
                                .Paragraph(p => p.Text("ColumnFollowingBody")))))))
            .ToBytes();

        using var pdf = PdfDocument.Open(new MemoryStream(bytes));

        Assert.Equal(2, pdf.NumberOfPages);
        Assert.Contains("IntroMarker", pdf.GetPage(1).Text);
        Assert.DoesNotContain("ColumnKeepWithNextLabel", pdf.GetPage(1).Text);
        Assert.Contains("ColumnKeepWithNextLabel", pdf.GetPage(2).Text);
        Assert.Contains("ColumnFollowingBody", pdf.GetPage(2).Text);
    }

    [Fact]
    public void RowColumnParagraph_KeepWithNextMovesParagraphWithFollowingList() {
        var options = new PdfOptions {
            PageWidth = 260,
            PageHeight = 170,
            MarginLeft = 30,
            MarginRight = 30,
            MarginTop = 30,
            MarginBottom = 30,
            DefaultFont = PdfStandardFont.Helvetica,
            DefaultFontSize = 10
        };

        byte[] bytes = PdfDoc.Create(options)
            .Paragraph(p => p.Text("IntroMarker"), style: new PdfParagraphStyle {
                SpacingAfter = 70
            })
            .Compose(document =>
                document.Page(page =>
                    page.Content(content =>
                        content.Row(row =>
                            row.Column(100, column => column
                                .Paragraph(p => p.Text("ColumnKeepWithListLabel"), style: new PdfParagraphStyle {
                                    KeepWithNext = true
                                })
                                .Bullets(new[] { "ColumnFollowingBullet" }))))))
            .ToBytes();

        using var pdf = PdfDocument.Open(new MemoryStream(bytes));

        Assert.Equal(2, pdf.NumberOfPages);
        Assert.Contains("IntroMarker", pdf.GetPage(1).Text);
        Assert.DoesNotContain("ColumnKeepWithListLabel", pdf.GetPage(1).Text);
        Assert.Contains("ColumnKeepWithListLabel", pdf.GetPage(2).Text);
        Assert.Contains("ColumnFollowingBullet", pdf.GetPage(2).Text);
    }

    [Fact]
    public void RowColumnHeading_KeepsWithFollowingParagraph() {
        var options = new PdfOptions {
            PageWidth = 260,
            PageHeight = 170,
            MarginLeft = 30,
            MarginRight = 30,
            MarginTop = 30,
            MarginBottom = 30,
            DefaultFont = PdfStandardFont.Helvetica,
            DefaultFontSize = 10
        };

        byte[] bytes = PdfDoc.Create(options)
            .Compose(document =>
                document.Page(page =>
                    page.Content(content =>
                        content.Row(row =>
                            row.Column(100, column => column
                                .Paragraph(p => p.Text("ColumnIntroMarker"), style: new PdfParagraphStyle {
                                    SpacingAfter = 70
                                })
                                .H3("ColumnSignalHeading")
                                .Paragraph(p => p.Text("ColumnSignalBody")))))))
            .ToBytes();

        using var pdf = PdfDocument.Open(new MemoryStream(bytes));

        Assert.Equal(2, pdf.NumberOfPages);
        Assert.Contains("ColumnIntroMarker", pdf.GetPage(1).Text);
        Assert.DoesNotContain("ColumnSignalHeading", pdf.GetPage(1).Text);
        Assert.Contains("ColumnSignalHeading", pdf.GetPage(2).Text);
        Assert.Contains("ColumnSignalBody", pdf.GetPage(2).Text);
    }

    [Fact]
    public void RowColumnParagraph_WidowControlAvoidsSingleLineAtPageBottom() {
        var options = new PdfOptions {
            PageWidth = 260,
            PageHeight = 170,
            MarginLeft = 30,
            MarginRight = 30,
            MarginTop = 30,
            MarginBottom = 30,
            DefaultFont = PdfStandardFont.Helvetica,
            DefaultFontSize = 10
        };

        byte[] bytes = PdfDoc.Create(options)
            .Compose(document =>
                document.Page(page =>
                    page.Content(content =>
                        content.Row(row =>
                            row.Column(100, column => column
                                .Paragraph(p => p.Text("ColumnIntroMarker"), style: new PdfParagraphStyle {
                                    SpacingAfter = 70
                                })
                                .Paragraph(p => p
                                    .Text("ColumnWidowFirst")
                                    .LineBreak()
                                    .Text("ColumnWidowSecond")
                                    .LineBreak()
                                    .Text("ColumnWidowThird"), style: new PdfParagraphStyle {
                                        WidowControl = true,
                                        SpacingAfter = 0
                                    }))))))
            .ToBytes();

        using var pdf = PdfDocument.Open(new MemoryStream(bytes));

        Assert.Equal(2, pdf.NumberOfPages);
        Assert.Contains("ColumnIntroMarker", pdf.GetPage(1).Text);
        Assert.DoesNotContain("ColumnWidowFirst", pdf.GetPage(1).Text);
        Assert.Contains("ColumnWidowFirst", pdf.GetPage(2).Text);
        Assert.Contains("ColumnWidowThird", pdf.GetPage(2).Text);
    }

    [Fact]
    public void RowColumnPanelParagraph_KeepWithNextMovesPanelWithFollowingParagraph() {
        var options = new PdfOptions {
            PageWidth = 260,
            PageHeight = 170,
            MarginLeft = 30,
            MarginRight = 30,
            MarginTop = 30,
            MarginBottom = 30,
            DefaultFont = PdfStandardFont.Helvetica,
            DefaultFontSize = 10
        };

        byte[] bytes = PdfDoc.Create(options)
            .Paragraph(p => p.Text("IntroMarker"), style: new PdfParagraphStyle {
                SpacingAfter = 60
            })
            .Compose(document =>
                document.Page(page =>
                    page.Content(content =>
                        content.Row(row =>
                            row.Column(100, column => column
                                .PanelParagraph(p => p.Text("ColumnPanelKeepWithNext"), new PanelStyle {
                                    KeepWithNext = true,
                                    PaddingY = 5,
                                    SpacingAfter = 0
                                })
                                .Paragraph(p => p.Text("ColumnFollowingPanelBody")))))))
            .ToBytes();

        using var pdf = PdfDocument.Open(new MemoryStream(bytes));

        Assert.Equal(2, pdf.NumberOfPages);
        Assert.Contains("IntroMarker", pdf.GetPage(1).Text);
        Assert.DoesNotContain("ColumnPanelKeepWithNext", pdf.GetPage(1).Text);
        Assert.Contains("ColumnPanelKeepWithNext", pdf.GetPage(2).Text);
        Assert.Contains("ColumnFollowingPanelBody", pdf.GetPage(2).Text);
    }

    [Fact]
    public void Paragraph_KeepTogetherRejectsContentTallerThanContentArea() {
        var options = new PdfOptions {
            PageWidth = 320,
            PageHeight = 170,
            MarginLeft = 30,
            MarginRight = 30,
            MarginTop = 30,
            MarginBottom = 30,
            DefaultFont = PdfStandardFont.Helvetica,
            DefaultFontSize = 9
        };
        string longText = string.Join(" ", Enumerable.Range(1, 180).Select(i => "paragraph" + i.ToString("000")));

        var exception = Assert.Throws<ArgumentException>(() =>
            PdfDoc.Create(options)
                .Paragraph(p => p.Text(longText), style: new PdfParagraphStyle {
                    KeepTogether = true
                })
                .ToBytes());

        Assert.Contains("Paragraph height exceeds the available page content height.", exception.Message, StringComparison.Ordinal);
    }

    [Fact]
    public void List_KeepTogetherRejectsContentTallerThanContentArea() {
        var options = new PdfOptions {
            PageWidth = 320,
            PageHeight = 170,
            MarginLeft = 30,
            MarginRight = 30,
            MarginTop = 30,
            MarginBottom = 30,
            DefaultFont = PdfStandardFont.Helvetica,
            DefaultFontSize = 10
        };

        var exception = Assert.Throws<ArgumentException>(() =>
            PdfDoc.Create(options)
                .Bullets(Enumerable.Range(1, 14).Select(i => "Keep list item " + i.ToString("00")), style: new PdfListStyle {
                    KeepTogether = true
                })
                .ToBytes());

        Assert.Contains("List height exceeds the available page content height.", exception.Message, StringComparison.Ordinal);
    }

    [Fact]
    public void PanelParagraph_KeepTogetherRejectsContentTallerThanContentArea() {
        var options = new PdfOptions {
            PageWidth = 320,
            PageHeight = 170,
            MarginLeft = 30,
            MarginRight = 30,
            MarginTop = 30,
            MarginBottom = 30,
            DefaultFont = PdfStandardFont.Helvetica,
            DefaultFontSize = 9
        };
        string longText = string.Join(" ", Enumerable.Range(1, 180).Select(i => "panel" + i.ToString("000")));

        var exception = Assert.Throws<ArgumentException>(() =>
            PdfDoc.Create(options)
                .PanelParagraph(p => p.Text(longText), new PanelStyle {
                    KeepTogether = true,
                    PaddingY = 8
                })
                .ToBytes());

        Assert.Contains("Panel height exceeds the available page content height.", exception.Message, StringComparison.Ordinal);
    }

    [Fact]
    public void PanelParagraph_LeavesBreathingRoomBeforeFollowingParagraph() {
        const double fontSize = 10;
        byte[] bytes = PdfDoc.Create(new PdfOptions {
                PageWidth = 260,
                PageHeight = 220,
                MarginLeft = 20,
                MarginRight = 20,
                MarginTop = 20,
                MarginBottom = 20,
                DefaultFont = PdfStandardFont.Helvetica,
                DefaultFontSize = fontSize
            })
            .PanelParagraph(p => p.Text("PanelMarker"), new PanelStyle {
                PaddingY = 6,
                BorderWidth = 0.5
            })
            .Paragraph(p => p.Text("AfterMarker"))
            .ToBytes();

        using var pdf = PdfDocument.Open(new MemoryStream(bytes));
        var page = pdf.GetPage(1);

        double ascender = fontSize * 0.74;
        double lineHeight = fontSize * 1.4;
        double paddingY = 6;
        double panelTextY = FindWordStartY(page, "PanelMarker");
        double panelTopY = panelTextY + paddingY + ascender;
        double panelBottomY = panelTopY - (paddingY + lineHeight + paddingY);
        double afterTopY = FindWordStartY(page, "AfterMarker") + ascender;
        double clearance = panelBottomY - afterTopY;

        Assert.True(clearance >= 5, $"Expected panel spacing to leave visible breathing room before following paragraph text. Clearance: {clearance:0.##}pt.");
    }

    [Fact]
    public void RowColumnPanelParagraph_LeavesBreathingRoomBeforeFollowingParagraph() {
        const double fontSize = 10;
        byte[] bytes = PdfDoc.Create(new PdfOptions {
                PageWidth = 260,
                PageHeight = 220,
                MarginLeft = 20,
                MarginRight = 20,
                MarginTop = 20,
                MarginBottom = 20,
                DefaultFont = PdfStandardFont.Helvetica,
                DefaultFontSize = fontSize
            })
            .Compose(document =>
                document.Page(page =>
                    page.Content(content =>
                        content.Row(row =>
                            row.Column(100, column => column
                                .PanelParagraph(p => p.Text("PanelMarker"), new PanelStyle {
                                    PaddingY = 6,
                                    BorderWidth = 0.5
                                })
                                .Paragraph(p => p.Text("AfterMarker")))))))
            .ToBytes();

        using var pdf = PdfDocument.Open(new MemoryStream(bytes));
        var page = pdf.GetPage(1);

        double ascender = fontSize * 0.74;
        double lineHeight = fontSize * 1.4;
        double paddingY = 6;
        double panelTextY = FindWordStartY(page, "PanelMarker");
        double panelTopY = panelTextY + paddingY + ascender;
        double panelBottomY = panelTopY - (paddingY + lineHeight + paddingY);
        double afterTopY = FindWordStartY(page, "AfterMarker") + ascender;
        double clearance = panelBottomY - afterTopY;

        Assert.True(clearance >= 5, $"Expected row-column panel spacing to leave visible breathing room before following paragraph text. Clearance: {clearance:0.##}pt.");
    }

    [Fact]
    public void RowColumnPanelParagraph_UsesDefaultPanelStyleWhenStyleIsNotProvided() {
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
        var style = new PanelStyle {
            Background = PdfColor.FromRgb(240, 248, 255),
            PaddingX = 16,
            MaxWidth = 120,
            Align = PdfAlign.Center
        };

        byte[] bytes = PdfDoc.Create(options)
            .DefaultPanelStyle(style)
            .Compose(document =>
                document.Page(page =>
                    page.Content(content =>
                        content.Row(row =>
                            row.Column(100, column => column
                                .PanelParagraph(p => p.Text("ColumnPanel")))))))
            .ToBytes();

        using var pdf = PdfDocument.Open(new MemoryStream(bytes));
        double panelTextX = FindWordStartX(pdf.GetPage(1), "ColumnPanel");
        string rawPdf = Encoding.ASCII.GetString(bytes);

        Assert.InRange(panelTextX, 135, 138);
        Assert.Contains("0.941 0.973 1 rg", rawPdf, StringComparison.Ordinal);
    }

    [Fact]
    public void PanelParagraph_UsesConfiguredSpacingBeforeAndAfter() {
        var options = new PdfOptions {
            PageWidth = 300,
            PageHeight = 260,
            MarginLeft = 30,
            MarginRight = 30,
            MarginTop = 30,
            MarginBottom = 30,
            DefaultFont = PdfStandardFont.Helvetica,
            DefaultFontSize = 10
        };

        byte[] defaultBytes = CreatePanelSpacingProbe(options, new PanelStyle {
            PaddingY = 6,
            SpacingBefore = 0,
            SpacingAfter = 0
        });
        byte[] spacedBytes = CreatePanelSpacingProbe(options, new PanelStyle {
            PaddingY = 6,
            SpacingBefore = 12,
            SpacingAfter = 18
        });

        using var defaultPdf = PdfDocument.Open(new MemoryStream(defaultBytes));
        using var spacedPdf = PdfDocument.Open(new MemoryStream(spacedBytes));
        var defaultPage = defaultPdf.GetPage(1);
        var spacedPage = spacedPdf.GetPage(1);

        double defaultPanelY = FindWordStartY(defaultPage, "PanelMarker");
        double spacedPanelY = FindWordStartY(spacedPage, "PanelMarker");
        double defaultAfterY = FindWordStartY(defaultPage, "AfterMarker");
        double spacedAfterY = FindWordStartY(spacedPage, "AfterMarker");

        Assert.True(defaultPanelY - spacedPanelY >= 10, $"Expected panel spacing before to move panel text down. Default y: {defaultPanelY:0.##}, spaced y: {spacedPanelY:0.##}.");
        Assert.True(defaultAfterY - spacedAfterY >= 28, $"Expected panel spacing before and after to move following text down. Default y: {defaultAfterY:0.##}, spaced y: {spacedAfterY:0.##}.");
    }

    [Fact]
    public void PanelParagraph_RejectsInvalidStyleValues() {
        var invalidBorderException = Assert.Throws<ArgumentException>(() =>
            new PanelStyle {
                BorderWidth = -0.5
            });

        Assert.Contains("Panel border width must be a non-negative finite value.", invalidBorderException.Message, StringComparison.Ordinal);

        var invalidHorizontalPaddingException = Assert.Throws<ArgumentException>(() =>
            new PanelStyle {
                PaddingX = double.PositiveInfinity
            });

        Assert.Contains("Panel horizontal padding must be a non-negative finite value.", invalidHorizontalPaddingException.Message, StringComparison.Ordinal);

        var invalidVerticalPaddingException = Assert.Throws<ArgumentException>(() =>
            new PanelStyle {
                PaddingY = -1
            });

        Assert.Contains("Panel vertical padding must be a non-negative finite value.", invalidVerticalPaddingException.Message, StringComparison.Ordinal);

        var invalidMaxWidthException = Assert.Throws<ArgumentException>(() =>
            new PanelStyle {
                MaxWidth = 0
            });

        Assert.Contains("Panel maximum width must be a positive finite value.", invalidMaxWidthException.Message, StringComparison.Ordinal);

        var invalidSpacingBeforeException = Assert.Throws<ArgumentException>(() =>
            new PanelStyle {
                SpacingBefore = -1
            });

        Assert.Contains("Panel spacing before must be a non-negative finite value.", invalidSpacingBeforeException.Message, StringComparison.Ordinal);

        var invalidSpacingAfterException = Assert.Throws<ArgumentException>(() =>
            new PanelStyle {
                SpacingAfter = double.NaN
            });

        Assert.Contains("Panel spacing after must be a non-negative finite value.", invalidSpacingAfterException.Message, StringComparison.Ordinal);

        var paddingException = Assert.Throws<ArgumentException>(() =>
            PdfDoc.Create(new PdfOptions {
                    PageWidth = 120,
                    MarginLeft = 20,
                    MarginRight = 20
                })
                .PanelParagraph(p => p.Text("No text frame"), new PanelStyle {
                    PaddingX = 40
                })
                .ToBytes());

        Assert.Contains("Panel horizontal padding must leave a positive text width.", paddingException.Message, StringComparison.Ordinal);
    }

    [Fact]
    public void PanelParagraph_SnapshotsStyleBeforeRendering() {
        var style = new PanelStyle {
            Background = PdfColor.FromRgb(26, 51, 77),
            BorderColor = PdfColor.FromRgb(40, 80, 120),
            BorderWidth = 2,
            PaddingX = 7,
            PaddingY = 8,
            MaxWidth = 140,
            Align = PdfAlign.Center,
            SpacingBefore = 3,
            SpacingAfter = 9,
            KeepTogether = true,
            KeepWithNext = true
        };

        var block = new PanelParagraphBlock(new[] { TextRun.Normal("Stable panel") }, PdfAlign.Left, null, style);

        style.Background = PdfColor.FromRgb(200, 10, 10);
        style.BorderColor = PdfColor.FromRgb(220, 10, 10);
        style.BorderWidth = 4;
        style.PaddingX = 20;
        style.PaddingY = 21;
        style.MaxWidth = 200;
        style.Align = PdfAlign.Right;
        style.SpacingBefore = 20;
        style.SpacingAfter = 21;
        style.KeepTogether = false;
        style.KeepWithNext = false;

        Assert.Equal(PdfColor.FromRgb(26, 51, 77), block.Style!.Background);
        Assert.Equal(PdfColor.FromRgb(40, 80, 120), block.Style.BorderColor);
        Assert.Equal(2, block.Style.BorderWidth);
        Assert.Equal(7, block.Style.PaddingX);
        Assert.Equal(8, block.Style.PaddingY);
        Assert.Equal(140, block.Style.MaxWidth);
        Assert.Equal(PdfAlign.Center, block.Style.Align);
        Assert.Equal(3, block.Style.SpacingBefore);
        Assert.Equal(9, block.Style.SpacingAfter);
        Assert.True(block.Style.KeepTogether);
        Assert.True(block.Style.KeepWithNext);

        var renderStyle = new PanelStyle {
            Background = PdfColor.FromRgb(26, 51, 77),
            BorderColor = PdfColor.FromRgb(40, 80, 120),
            BorderWidth = 2,
            PaddingX = 7,
            PaddingY = 8,
            MaxWidth = 140,
            Align = PdfAlign.Center,
            KeepTogether = true,
            KeepWithNext = true
        };

        var doc = PdfDoc.Create(new PdfOptions {
                PageWidth = 240,
                PageHeight = 160,
                MarginLeft = 20,
                MarginRight = 20,
                MarginTop = 20,
                MarginBottom = 20
            })
            .PanelParagraph(p => p.Text("Stable panel"), renderStyle);

        renderStyle.Background = PdfColor.FromRgb(200, 10, 10);
        renderStyle.BorderColor = PdfColor.FromRgb(220, 10, 10);
        renderStyle.BorderWidth = 4;
        renderStyle.PaddingX = 20;
        renderStyle.PaddingY = 21;
        renderStyle.MaxWidth = 200;
        renderStyle.Align = PdfAlign.Right;
        renderStyle.KeepTogether = false;
        renderStyle.KeepWithNext = false;

        byte[] bytes = doc.ToBytes();

        string content = Encoding.ASCII.GetString(bytes);

        Assert.Contains("0.102 0.2 0.302 rg", content);
        Assert.DoesNotContain("0.784 0.039 0.039 rg", content);
    }

    [Fact]
    public void HorizontalRule_RendersInTopLevelFlow() {
        byte[] bytes = PdfDoc.Create(new PdfOptions {
                PageWidth = 240,
                PageHeight = 180,
                MarginLeft = 20,
                MarginRight = 20,
                MarginTop = 20,
                MarginBottom = 20
            })
            .HR(
                thickness: 3,
                color: PdfColor.FromRgb(26, 51, 77),
                spacingBefore: 4,
                spacingAfter: 6)
            .ToBytes();

        string content = Encoding.ASCII.GetString(bytes);

        Assert.Contains("0.102 0.2 0.302 RG", content);
        Assert.Contains("3 w", content);
        Assert.Contains("20 158.5 m 220 158.5 l S", content);
    }

    [Fact]
    public void HorizontalRule_LeavesBreathingRoomBeforeFollowingParagraph() {
        const double fontSize = 10;
        byte[] bytes = PdfDoc.Create(new PdfOptions {
                PageWidth = 240,
                PageHeight = 180,
                MarginLeft = 20,
                MarginRight = 20,
                MarginTop = 20,
                MarginBottom = 20,
                DefaultFont = PdfStandardFont.Helvetica,
                DefaultFontSize = fontSize
            })
            .HR(
                thickness: 3,
                color: PdfColor.FromRgb(26, 51, 77),
                spacingBefore: 4,
                spacingAfter: 6)
            .Paragraph(p => p.Text("Guarded rhythm stays below the rule."))
            .ToBytes();

        using var pdf = PdfDocument.Open(new MemoryStream(bytes));
        var page = pdf.GetPage(1);

        double ruleBottomY = 180 - 20 - 3;
        double paragraphTopY = FindWordStartY(page, "Guarded") + fontSize * 0.74;
        double clearance = ruleBottomY - paragraphTopY;

        Assert.True(clearance >= 5, $"Expected rule spacing to leave visible breathing room before paragraph text. Clearance: {clearance:0.##}pt.");
    }

    [Fact]
    public void RowColumnHorizontalRule_LeavesBreathingRoomBeforeFollowingParagraph() {
        const double fontSize = 10;
        byte[] bytes = PdfDoc.Create(new PdfOptions {
                PageWidth = 240,
                PageHeight = 180,
                MarginLeft = 20,
                MarginRight = 20,
                MarginTop = 20,
                MarginBottom = 20,
                DefaultFont = PdfStandardFont.Helvetica,
                DefaultFontSize = fontSize
            })
            .Compose(document =>
                document.Page(page =>
                    page.Content(content =>
                        content.Row(row =>
                            row.Column(100, column => column
                                .HR(
                                    thickness: 3,
                                    color: PdfColor.FromRgb(26, 51, 77),
                                    spacingBefore: 4,
                                    spacingAfter: 6)
                                .Paragraph(p => p.Text("Guarded rhythm stays below the rule.")))))))
            .ToBytes();

        using var pdf = PdfDocument.Open(new MemoryStream(bytes));
        var page = pdf.GetPage(1);

        double ruleBottomY = 180 - 20 - 3;
        double paragraphTopY = FindWordStartY(page, "Guarded") + fontSize * 0.74;
        double clearance = ruleBottomY - paragraphTopY;

        Assert.True(clearance >= 5, $"Expected row-column rule spacing to leave visible breathing room before paragraph text. Clearance: {clearance:0.##}pt.");
    }

    [Fact]
    public void HorizontalRule_KeepWithNextMovesRuleWithFollowingParagraph() {
        var options = new PdfOptions {
            PageWidth = 260,
            PageHeight = 170,
            MarginLeft = 30,
            MarginRight = 30,
            MarginTop = 30,
            MarginBottom = 30,
            DefaultFont = PdfStandardFont.Helvetica,
            DefaultFontSize = 10
        };

        byte[] bytes = PdfDoc.Create(options)
            .Paragraph(p => p.Text("IntroMarker"), style: new PdfParagraphStyle {
                SpacingAfter = 81
            })
            .HR(style: new PdfHorizontalRuleStyle {
                Thickness = 3,
                SpacingBefore = 0,
                SpacingAfter = 0,
                KeepWithNext = true
            })
            .Paragraph(p => p.Text("FollowingRuleBody"))
            .ToBytes();

        using var pdf = PdfDocument.Open(new MemoryStream(bytes));
        string page1Content = string.Join("\n", GetPageContentStreams(bytes, 1));
        string page2Content = string.Join("\n", GetPageContentStreams(bytes, 2));

        Assert.Equal(2, pdf.NumberOfPages);
        Assert.Contains("IntroMarker", pdf.GetPage(1).Text);
        Assert.DoesNotContain("FollowingRuleBody", pdf.GetPage(1).Text);
        Assert.Contains("FollowingRuleBody", pdf.GetPage(2).Text);
        Assert.DoesNotContain("30 43.5 m 230 43.5 l S", page1Content);
        Assert.Contains("30 138.5 m 230 138.5 l S", page2Content);
    }

    [Fact]
    public void RowColumnHorizontalRule_KeepWithNextMovesRuleWithFollowingParagraph() {
        var options = new PdfOptions {
            PageWidth = 260,
            PageHeight = 170,
            MarginLeft = 30,
            MarginRight = 30,
            MarginTop = 30,
            MarginBottom = 30,
            DefaultFont = PdfStandardFont.Helvetica,
            DefaultFontSize = 10
        };

        byte[] bytes = PdfDoc.Create(options)
            .Paragraph(p => p.Text("IntroMarker"), style: new PdfParagraphStyle {
                SpacingAfter = 81
            })
            .Compose(document =>
                document.Page(page =>
                    page.Content(content =>
                        content.Row(row =>
                            row.Column(100, column => column
                                .HR(style: new PdfHorizontalRuleStyle {
                                    Thickness = 3,
                                    SpacingBefore = 0,
                                    SpacingAfter = 0,
                                    KeepWithNext = true
                                })
                                .Paragraph(p => p.Text("ColumnFollowingRuleBody")))))))
            .ToBytes();

        using var pdf = PdfDocument.Open(new MemoryStream(bytes));
        string page1Content = string.Join("\n", GetPageContentStreams(bytes, 1));
        string page2Content = string.Join("\n", GetPageContentStreams(bytes, 2));

        Assert.Equal(2, pdf.NumberOfPages);
        Assert.Contains("IntroMarker", pdf.GetPage(1).Text);
        Assert.DoesNotContain("ColumnFollowingRuleBody", pdf.GetPage(1).Text);
        Assert.Contains("ColumnFollowingRuleBody", pdf.GetPage(2).Text);
        Assert.DoesNotContain("30 43.5 m 230 43.5 l S", page1Content);
        Assert.Contains("30 138.5 m 230 138.5 l S", page2Content);
    }

    [Fact]
    public void Image_KeepWithNextMovesImageWithFollowingParagraph() {
        byte[] png = CreateMinimalRgbPng();
        var options = new PdfOptions {
            PageWidth = 260,
            PageHeight = 170,
            MarginLeft = 30,
            MarginRight = 30,
            MarginTop = 30,
            MarginBottom = 30,
            DefaultFont = PdfStandardFont.Helvetica,
            DefaultFontSize = 10
        };

        byte[] bytes = PdfDoc.Create(options)
            .Paragraph(p => p.Text("IntroMarker"), style: new PdfParagraphStyle {
                SpacingAfter = 66
            })
            .Image(png, 24, 24, style: new PdfImageStyle {
                KeepWithNext = true
            })
            .Paragraph(p => p.Text("FollowingImageBody"))
            .ToBytes();

        using var pdf = PdfDocument.Open(new MemoryStream(bytes));
        string page1Content = string.Join("\n", GetPageContentStreams(bytes, 1));
        string page2Content = string.Join("\n", GetPageContentStreams(bytes, 2));

        Assert.Equal(2, pdf.NumberOfPages);
        Assert.Contains("IntroMarker", pdf.GetPage(1).Text);
        Assert.DoesNotContain("FollowingImageBody", pdf.GetPage(1).Text);
        Assert.Contains("FollowingImageBody", pdf.GetPage(2).Text);
        Assert.DoesNotContain("/Im1 Do", page1Content);
        Assert.Contains("/Im1 Do", page2Content);
    }

    [Fact]
    public void RowColumnImage_KeepWithNextMovesImageWithFollowingParagraph() {
        byte[] png = CreateMinimalRgbPng();
        var options = new PdfOptions {
            PageWidth = 260,
            PageHeight = 170,
            MarginLeft = 30,
            MarginRight = 30,
            MarginTop = 30,
            MarginBottom = 30,
            DefaultFont = PdfStandardFont.Helvetica,
            DefaultFontSize = 10
        };

        byte[] bytes = PdfDoc.Create(options)
            .Paragraph(p => p.Text("IntroMarker"), style: new PdfParagraphStyle {
                SpacingAfter = 66
            })
            .Compose(document =>
                document.Page(page =>
                    page.Content(content =>
                        content.Row(row =>
                            row.Column(100, column => column
                                .Image(png, 24, 24, style: new PdfImageStyle {
                                    KeepWithNext = true
                                })
                                .Paragraph(p => p.Text("ColumnFollowingImageBody")))))))
            .ToBytes();

        using var pdf = PdfDocument.Open(new MemoryStream(bytes));
        string page1Content = string.Join("\n", GetPageContentStreams(bytes, 1));
        string page2Content = string.Join("\n", GetPageContentStreams(bytes, 2));

        Assert.Equal(2, pdf.NumberOfPages);
        Assert.Contains("IntroMarker", pdf.GetPage(1).Text);
        Assert.DoesNotContain("ColumnFollowingImageBody", pdf.GetPage(1).Text);
        Assert.Contains("ColumnFollowingImageBody", pdf.GetPage(2).Text);
        Assert.DoesNotContain("/Im1 Do", page1Content);
        Assert.Contains("/Im1 Do", page2Content);
    }

    [Fact]
    public void Shape_KeepWithNextMovesShapeWithFollowingParagraph() {
        var shape = OfficeShape.Rectangle(24, 24);
        shape.FillColor = OfficeColor.WhiteSmoke;
        var options = new PdfOptions {
            PageWidth = 260,
            PageHeight = 170,
            MarginLeft = 30,
            MarginRight = 30,
            MarginTop = 30,
            MarginBottom = 30,
            DefaultFont = PdfStandardFont.Helvetica,
            DefaultFontSize = 10
        };

        byte[] bytes = PdfDoc.Create(options)
            .Paragraph(p => p.Text("IntroMarker"), style: new PdfParagraphStyle {
                SpacingAfter = 66
            })
            .Shape(shape, style: new PdfDrawingStyle {
                KeepWithNext = true
            })
            .Paragraph(p => p.Text("FollowingShapeBody"))
            .ToBytes();

        using var pdf = PdfDocument.Open(new MemoryStream(bytes));
        string page1Content = string.Join("\n", GetPageContentStreams(bytes, 1));
        string page2Content = string.Join("\n", GetPageContentStreams(bytes, 2));

        Assert.Equal(2, pdf.NumberOfPages);
        Assert.Contains("IntroMarker", pdf.GetPage(1).Text);
        Assert.DoesNotContain("FollowingShapeBody", pdf.GetPage(1).Text);
        Assert.Contains("FollowingShapeBody", pdf.GetPage(2).Text);
        Assert.DoesNotContain("24 24 re f", page1Content);
        Assert.Contains("30 116 24 24 re f", page2Content);
    }

    [Fact]
    public void RowColumnShape_KeepWithNextMovesShapeWithFollowingParagraph() {
        var shape = OfficeShape.Rectangle(24, 24);
        shape.FillColor = OfficeColor.WhiteSmoke;
        var options = new PdfOptions {
            PageWidth = 260,
            PageHeight = 170,
            MarginLeft = 30,
            MarginRight = 30,
            MarginTop = 30,
            MarginBottom = 30,
            DefaultFont = PdfStandardFont.Helvetica,
            DefaultFontSize = 10
        };

        byte[] bytes = PdfDoc.Create(options)
            .Paragraph(p => p.Text("IntroMarker"), style: new PdfParagraphStyle {
                SpacingAfter = 66
            })
            .Compose(document =>
                document.Page(page =>
                    page.Content(content =>
                        content.Row(row =>
                            row.Column(100, column => column
                                .Shape(shape, style: new PdfDrawingStyle {
                                    KeepWithNext = true
                                })
                                .Paragraph(p => p.Text("ColumnFollowingShapeBody")))))))
            .ToBytes();

        using var pdf = PdfDocument.Open(new MemoryStream(bytes));
        string page1Content = string.Join("\n", GetPageContentStreams(bytes, 1));
        string page2Content = string.Join("\n", GetPageContentStreams(bytes, 2));

        Assert.Equal(2, pdf.NumberOfPages);
        Assert.Contains("IntroMarker", pdf.GetPage(1).Text);
        Assert.DoesNotContain("ColumnFollowingShapeBody", pdf.GetPage(1).Text);
        Assert.Contains("ColumnFollowingShapeBody", pdf.GetPage(2).Text);
        Assert.DoesNotContain("24 24 re f", page1Content);
        Assert.Contains("30 116 24 24 re f", page2Content);
    }

    [Fact]
    public void Drawing_KeepWithNextMovesDrawingWithFollowingParagraph() {
        var drawing = CreateKeepWithNextDrawingScene();
        var options = new PdfOptions {
            PageWidth = 260,
            PageHeight = 170,
            MarginLeft = 30,
            MarginRight = 30,
            MarginTop = 30,
            MarginBottom = 30,
            DefaultFont = PdfStandardFont.Helvetica,
            DefaultFontSize = 10
        };

        byte[] bytes = PdfDoc.Create(options)
            .Paragraph(p => p.Text("IntroMarker"), style: new PdfParagraphStyle {
                SpacingAfter = 66
            })
            .Drawing(drawing, style: new PdfDrawingStyle {
                KeepWithNext = true
            })
            .Paragraph(p => p.Text("FollowingDrawingBody"))
            .ToBytes();

        using var pdf = PdfDocument.Open(new MemoryStream(bytes));
        string page1Content = string.Join("\n", GetPageContentStreams(bytes, 1));
        string page2Content = string.Join("\n", GetPageContentStreams(bytes, 2));

        Assert.Equal(2, pdf.NumberOfPages);
        Assert.Contains("IntroMarker", pdf.GetPage(1).Text);
        Assert.DoesNotContain("FollowingDrawingBody", pdf.GetPage(1).Text);
        Assert.Contains("FollowingDrawingBody", pdf.GetPage(2).Text);
        Assert.DoesNotContain("24 24 re f", page1Content);
        Assert.Contains("30 116 24 24 re f", page2Content);
    }

    [Fact]
    public void RowColumnDrawing_KeepWithNextMovesDrawingWithFollowingParagraph() {
        var drawing = CreateKeepWithNextDrawingScene();
        var options = new PdfOptions {
            PageWidth = 260,
            PageHeight = 170,
            MarginLeft = 30,
            MarginRight = 30,
            MarginTop = 30,
            MarginBottom = 30,
            DefaultFont = PdfStandardFont.Helvetica,
            DefaultFontSize = 10
        };

        byte[] bytes = PdfDoc.Create(options)
            .Paragraph(p => p.Text("IntroMarker"), style: new PdfParagraphStyle {
                SpacingAfter = 66
            })
            .Compose(document =>
                document.Page(page =>
                    page.Content(content =>
                        content.Row(row =>
                            row.Column(100, column => column
                                .Drawing(drawing, style: new PdfDrawingStyle {
                                    KeepWithNext = true
                                })
                                .Paragraph(p => p.Text("ColumnFollowingDrawingBody")))))))
            .ToBytes();

        using var pdf = PdfDocument.Open(new MemoryStream(bytes));
        string page1Content = string.Join("\n", GetPageContentStreams(bytes, 1));
        string page2Content = string.Join("\n", GetPageContentStreams(bytes, 2));

        Assert.Equal(2, pdf.NumberOfPages);
        Assert.Contains("IntroMarker", pdf.GetPage(1).Text);
        Assert.DoesNotContain("ColumnFollowingDrawingBody", pdf.GetPage(1).Text);
        Assert.Contains("ColumnFollowingDrawingBody", pdf.GetPage(2).Text);
        Assert.DoesNotContain("24 24 re f", page1Content);
        Assert.Contains("30 116 24 24 re f", page2Content);
    }

    [Fact]
    public void Table_KeepWithNextMovesTableWithFollowingParagraph() {
        var style = TableStyles.Minimal();
        style.HeaderRowCount = 0;
        style.KeepWithNext = true;
        var options = new PdfOptions {
            PageWidth = 260,
            PageHeight = 170,
            MarginLeft = 30,
            MarginRight = 30,
            MarginTop = 30,
            MarginBottom = 30,
            DefaultFont = PdfStandardFont.Helvetica,
            DefaultFontSize = 10
        };

        byte[] bytes = PdfDoc.Create(options)
            .Paragraph(p => p.Text("IntroMarker"), style: new PdfParagraphStyle {
                SpacingAfter = 60
            })
            .Table(new[] {
                new[] { "TableKeepHeader", "Ready" },
                new[] { "TableKeepValue", "Ready" }
            }, style: style)
            .Paragraph(p => p.Text("FollowingTableBody"))
            .ToBytes();

        using var pdf = PdfDocument.Open(new MemoryStream(bytes));

        Assert.Equal(2, pdf.NumberOfPages);
        Assert.Contains("IntroMarker", pdf.GetPage(1).Text);
        Assert.DoesNotContain("TableKeepValue", pdf.GetPage(1).Text);
        Assert.DoesNotContain("FollowingTableBody", pdf.GetPage(1).Text);
        Assert.Contains("TableKeepValue", pdf.GetPage(2).Text);
        Assert.Contains("FollowingTableBody", pdf.GetPage(2).Text);
    }

    [Fact]
    public void RowColumnTable_KeepWithNextMovesTableWithFollowingParagraph() {
        var style = TableStyles.Minimal();
        style.HeaderRowCount = 0;
        style.KeepWithNext = true;
        var options = new PdfOptions {
            PageWidth = 260,
            PageHeight = 170,
            MarginLeft = 30,
            MarginRight = 30,
            MarginTop = 30,
            MarginBottom = 30,
            DefaultFont = PdfStandardFont.Helvetica,
            DefaultFontSize = 10
        };

        byte[] bytes = PdfDoc.Create(options)
            .Paragraph(p => p.Text("IntroMarker"), style: new PdfParagraphStyle {
                SpacingAfter = 60
            })
            .Compose(document =>
                document.Page(page =>
                    page.Content(content =>
                        content.Row(row =>
                            row.Column(100, column => column
                                .Table(new[] {
                                    new[] { "ColumnTableKeepHeader", "Ready" },
                                    new[] { "ColumnTableKeepValue", "Ready" }
                                }, style: style)
                                .Paragraph(p => p.Text("ColumnFollowingTableBody")))))))
            .ToBytes();

        using var pdf = PdfDocument.Open(new MemoryStream(bytes));

        Assert.Equal(2, pdf.NumberOfPages);
        Assert.Contains("IntroMarker", pdf.GetPage(1).Text);
        Assert.DoesNotContain("ColumnTableKeepValue", pdf.GetPage(1).Text);
        Assert.DoesNotContain("ColumnFollowingTableBody", pdf.GetPage(1).Text);
        Assert.Contains("ColumnTableKeepValue", pdf.GetPage(2).Text);
        Assert.Contains("ColumnFollowingTableBody", pdf.GetPage(2).Text);
    }

    [Fact]
    public void Row_KeepWithNextMovesRowWithFollowingParagraph() {
        var rowStyle = new PdfRowStyle {
            KeepWithNext = true
        };
        var options = new PdfOptions {
            PageWidth = 260,
            PageHeight = 170,
            MarginLeft = 30,
            MarginRight = 30,
            MarginTop = 30,
            MarginBottom = 30,
            DefaultFont = PdfStandardFont.Helvetica,
            DefaultFontSize = 10
        };

        byte[] bytes = PdfDoc.Create(options)
            .Compose(document =>
                document.Page(page =>
                    page.Content(content => {
                        content.Column(column =>
                            column.Item().Paragraph(p => p.Text("IntroMarker"), style: new PdfParagraphStyle {
                                SpacingAfter = 75
                            }));
                        content.Row(row => {
                            row.Style(rowStyle);
                            row.Column(100, column =>
                                column.Paragraph(p => p.Text("RowKeepColumn")));
                        });
                        content.Column(column =>
                            column.Item().Paragraph(p => p.Text("FollowingRowBody")));
                    })))
            .ToBytes();

        using var pdf = PdfDocument.Open(new MemoryStream(bytes));

        Assert.Equal(2, pdf.NumberOfPages);
        Assert.Contains("IntroMarker", pdf.GetPage(1).Text);
        Assert.DoesNotContain("RowKeepColumn", pdf.GetPage(1).Text);
        Assert.DoesNotContain("FollowingRowBody", pdf.GetPage(1).Text);
        Assert.Contains("RowKeepColumn", pdf.GetPage(2).Text);
        Assert.Contains("FollowingRowBody", pdf.GetPage(2).Text);
    }

    [Fact]
    public void RowColumns_KeepTextInsideColumnFramesWithReadableRhythm() {
        const double pageWidth = 420;
        const double margin = 30;
        const double gutter = 24;
        double contentWidth = pageWidth - margin - margin;
        double columnWidth = (contentWidth - gutter) / 2;
        double leftX = margin;
        double leftRightX = leftX + columnWidth;
        double rightX = leftRightX + gutter;
        double rightRightX = rightX + columnWidth;

        byte[] bytes = PdfDoc.Create(new PdfOptions {
                PageWidth = pageWidth,
                PageHeight = 280,
                MarginLeft = margin,
                MarginRight = margin,
                MarginTop = 30,
                MarginBottom = 30,
                DefaultFont = PdfStandardFont.Helvetica,
                DefaultFontSize = 10
            })
            .Compose(document =>
                document.Page(page =>
                    page.Content(content =>
                        content.Row(row =>
                            row.Gap(gutter)
                                .Column(50, column => column
                                    .H2("LeftFlow")
                                    .Paragraph(p => p.Text("LeftAlphaOne carries enough ordinary report text to wrap inside its column without touching the neighboring frame."))
                                    .Bullets(new[] {
                                        "LeftBulletOne stays inside the measure.",
                                        "LeftBulletTwo keeps a clear baseline."
                                    })
                                    .PanelParagraph(
                                        p => p.Bold("LeftPanel").Text(": spacing remains visible after the list."),
                                        new PanelStyle {
                                            BorderColor = PdfColor.FromRgb(191, 191, 191),
                                            BorderWidth = 0.5,
                                            PaddingX = 6,
                                            PaddingY = 5,
                                            Background = PdfColor.FromRgb(248, 250, 252)
                                        }))
                                .Column(50, column => column
                                    .H2("RightFlow")
                                    .Paragraph(p => p.Text("RightAlphaOne uses the same generic layout primitives and should start after the explicit gutter."))
                                    .Numbered(new[] {
                                        "RightStepOne composes content.",
                                        "RightStepTwo preserves reading rhythm."
                                    })
                                    .PanelParagraph(
                                        p => p.Bold("RightPanel").Text(": the final note avoids cramped text."),
                                        new PanelStyle {
                                            BorderColor = PdfColor.FromRgb(191, 191, 191),
                                            BorderWidth = 0.5,
                                            PaddingX = 6,
                                            PaddingY = 5,
                                            Background = PdfColor.FromRgb(248, 250, 252)
                                        }))))))
            .ToBytes();

        using var pdf = PdfDocument.Open(new MemoryStream(bytes));
        var renderedPage = pdf.GetPage(1);
        var leftLines = GetVisualTextLines(renderedPage, leftX - 1, leftRightX + 1);
        var rightLines = GetVisualTextLines(renderedPage, rightX - 1, rightRightX + 1);

        Assert.Contains(leftLines, line => line.Text.Contains("LeftAlphaOne", StringComparison.Ordinal));
        Assert.Contains(rightLines, line => line.Text.Contains("RightAlphaOne", StringComparison.Ordinal));
        Assert.True(leftLines.Count >= 7, $"Expected the left flow to produce multiple visual lines. Lines: {leftLines.Count}.");
        Assert.True(rightLines.Count >= 7, $"Expected the right flow to produce multiple visual lines. Lines: {rightLines.Count}.");

        Assert.All(leftLines, line =>
            Assert.True(line.X1 >= leftX - 1 && line.X2 <= leftRightX + 1.5,
                $"Expected left column line '{line.Text}' to stay inside {leftX:0.##}..{leftRightX:0.##}, but it rendered at {line.X1:0.##}..{line.X2:0.##}."));
        Assert.All(rightLines, line =>
            Assert.True(line.X1 >= rightX - 1.5 && line.X2 <= rightRightX + 1.5,
                $"Expected right column line '{line.Text}' to stay inside {rightX:0.##}..{rightRightX:0.##}, but it rendered at {line.X1:0.##}..{line.X2:0.##}."));

        foreach (var leftLine in leftLines) {
            foreach (var rightLine in rightLines.Where(line => Math.Abs(line.BaselineY - leftLine.BaselineY) <= 0.2)) {
                double clearance = rightLine.X1 - leftLine.X2;
                Assert.True(clearance >= gutter - 1,
                    $"Expected row columns to preserve the {gutter:0.##}pt gutter between '{leftLine.Text}' and '{rightLine.Text}'. Clearance: {clearance:0.##}pt.");
            }
        }

        AssertReadableTextRhythm(leftLines, "left column");
        AssertReadableTextRhythm(rightLines, "right column");
    }

    [Fact]
    public void RowColumns_UseBuiltInWordLikeGutterWhenNoGapIsConfigured() {
        const double pageWidth = 360;
        const double margin = 30;
        const double gutter = PdfRowStyle.DefaultGap;
        double contentWidth = pageWidth - margin - margin;
        double columnWidth = (contentWidth - gutter) / 2;
        double leftX = margin;
        double leftRightX = leftX + columnWidth;
        double rightX = leftRightX + gutter;
        double rightRightX = rightX + columnWidth;

        byte[] bytes = PdfDoc.Create(new PdfOptions {
                PageWidth = pageWidth,
                PageHeight = 180,
                MarginLeft = margin,
                MarginRight = margin,
                MarginTop = 30,
                MarginBottom = 30,
                DefaultFont = PdfStandardFont.Helvetica,
                DefaultFontSize = 10
            })
            .Compose(document =>
                document.Page(page =>
                    page.Content(content =>
                        content.Row(row => row
                            .Column(50, column => column.Paragraph(p => p.Text("LeftPlainColumn wraps in the first default column frame.")))
                            .Column(50, column => column.Paragraph(p => p.Text("RightPlainColumn starts after the built-in row gutter.")))))))
            .ToBytes();

        using var pdf = PdfDocument.Open(new MemoryStream(bytes));
        var renderedPage = pdf.GetPage(1);
        var leftLines = GetVisualTextLines(renderedPage, leftX - 1, leftRightX + 1);
        var rightLines = GetVisualTextLines(renderedPage, rightX - 1, rightRightX + 1);
        double rightStart = FindWordStartX(renderedPage, "RightPlainColumn");

        Assert.Contains(leftLines, line => line.Text.Contains("LeftPlainColumn", StringComparison.Ordinal));
        Assert.Contains(rightLines, line => line.Text.Contains("RightPlainColumn", StringComparison.Ordinal));
        Assert.True(rightStart >= rightX - 1,
            $"Expected an unstyled two-column row to use the built-in {gutter:0.##}pt gutter. Right column started at {rightStart:0.##}, expected at least {rightX:0.##}.");
        Assert.All(leftLines, line =>
            Assert.True(line.X1 >= leftX - 1 && line.X2 <= leftRightX + 1.5,
                $"Expected default-gutter left column line '{line.Text}' to stay inside the left column frame."));
        Assert.All(rightLines, line =>
            Assert.True(line.X1 >= rightX - 1.5 && line.X2 <= rightRightX + 1.5,
                $"Expected default-gutter right column line '{line.Text}' to stay inside the right column frame."));
        AssertNoSameBaselineTextCollisions(renderedPage, "default-gutter row columns");
    }

    [Fact]
    public void MixedWordLikeFlow_KeepsReadableRhythmAcrossGenericPrimitives() {
        const double pageWidth = 420;
        const double margin = 36;
        const double gutter = 24;
        double contentWidth = pageWidth - margin - margin;
        double columnWidth = (contentWidth - gutter) / 2;
        double leftX = margin;
        double leftRightX = leftX + columnWidth;
        double rightX = leftRightX + gutter;
        double rightRightX = rightX + columnWidth;
        byte[] png = CreateMinimalRgbPng();
        var shape = OfficeShape.Rectangle(72, 16);
        shape.FillColor = OfficeColor.LightBlue;
        shape.StrokeColor = OfficeColor.SteelBlue;
        shape.StrokeWidth = 0.75;

        var paragraphStyle = new PdfParagraphStyle {
            SpacingAfter = 10,
            LineHeight = 1.25
        };
        var panelStyle = new PanelStyle {
            Background = PdfColor.FromRgb(248, 250, 252),
            BorderColor = PdfColor.FromRgb(191, 191, 191),
            BorderWidth = 0.5,
            PaddingX = 8,
            PaddingY = 6,
            SpacingBefore = 2,
            SpacingAfter = 10
        };
        var listStyle = new PdfListStyle {
            LeftIndent = 10,
            MarkerGap = 6,
            SpacingAfter = 10,
            ItemSpacing = 2
        };
        var tableStyle = TableStyles.Minimal();
        tableStyle.HeaderRowCount = 0;
        tableStyle.CellPaddingX = 6;
        tableStyle.CellPaddingY = 4;
        tableStyle.SpacingBefore = 2;
        tableStyle.SpacingAfter = 10;
        var imageStyle = new PdfImageStyle {
            SpacingBefore = 4,
            SpacingAfter = 10
        };
        var drawingStyle = new PdfDrawingStyle {
            SpacingBefore = 2,
            SpacingAfter = 10
        };

        byte[] bytes = PdfDoc.Create(new PdfOptions {
                PageWidth = pageWidth,
                PageHeight = 620,
                MarginLeft = margin,
                MarginRight = margin,
                MarginTop = 36,
                MarginBottom = 36,
                DefaultFont = PdfStandardFont.Helvetica,
                DefaultFontSize = 10
            })
            .Compose(document =>
                document.Page(page =>
                    page.Content(content => {
                        content.Column(column => column.Item()
                            .H1("WordFlowGate", new PdfHeadingStyle {
                                FontSize = 18,
                                SpacingAfter = 12
                            })
                            .Paragraph(p => p.Text("LeadMarker introduces a generic report section without any invoice-specific shape."), style: paragraphStyle)
                            .PanelParagraph(p => p.Bold("PanelMarker").Text(": boxed notes keep visible padding and downstream rhythm."), panelStyle)
                            .Bullets(new[] {
                                "BulletMarker keeps list text in the normal document flow.",
                                "BulletSecond keeps a second readable baseline."
                            }, style: listStyle)
                            .Table(new[] {
                                new[] { "TableMarker", "Ready" },
                                new[] { "Rhythm", "Stable" }
                            }, style: tableStyle)
                            .Image(png, 24, 24, style: imageStyle)
                            .Shape(shape, style: drawingStyle)
                            .Paragraph(p => p.Text("AfterVisualMarker follows image and shape blocks with deliberate breathing room."), style: paragraphStyle));

                        content.Row(row => row
                            .Gap(gutter)
                            .Style(new PdfRowStyle {
                                SpacingBefore = 4,
                                SpacingAfter = 0
                            })
                            .Column(50, column => column
                                .H2("LeftMixed")
                                .Paragraph(p => p.Text("LeftMixedMarker wraps safely inside the left column measure."), style: paragraphStyle)
                                .Bullets(new[] { "LeftMixedBullet keeps rhythm." }, style: listStyle))
                            .Column(50, column => column
                                .H2("RightMixed")
                                .Paragraph(p => p.Text("RightMixedMarker starts after the explicit row gutter."), style: paragraphStyle)
                                .Numbered(new[] { "RightMixedStep keeps rhythm." }, style: listStyle)));
                    })))
            .ToBytes();

        using var pdf = PdfDocument.Open(new MemoryStream(bytes));
        var page = pdf.GetPage(1);
        string pageContent = string.Join("\n", GetPageContentStreams(bytes, 1));
        double leadY = FindWordStartY(page, "LeadMarker");
        double panelY = FindWordStartY(page, "PanelMarker");
        double bulletY = FindWordStartY(page, "BulletMarker");
        double tableY = FindWordStartY(page, "TableMarker");
        double afterVisualY = FindWordStartY(page, "AfterVisualMarker");
        var leftLines = GetVisualTextLines(page, leftX - 1, leftRightX + 1);
        var rightLines = GetVisualTextLines(page, rightX - 1, rightRightX + 1);
        var pageLines = GetVisualTextLines(page, 0, pageWidth);

        Assert.Equal(1, pdf.NumberOfPages);
        Assert.True(leadY - panelY >= 18, $"Expected panel content to sit below the lead paragraph with readable rhythm. Gap: {leadY - panelY:0.##}pt.");
        Assert.True(panelY - bulletY >= 18, $"Expected list content to sit below the panel with readable rhythm. Gap: {panelY - bulletY:0.##}pt.");
        Assert.True(bulletY - tableY >= 18, $"Expected table content to sit below the list with readable rhythm. Gap: {bulletY - tableY:0.##}pt.");
        Assert.True(tableY - afterVisualY >= 55, $"Expected text after image and shape blocks to preserve visual breathing room. Gap: {tableY - afterVisualY:0.##}pt.");
        Assert.Contains("/Im1 Do", pageContent);
        Assert.Contains("72 16 re B", pageContent);
        Assert.Contains(leftLines, line => line.Text.Contains("LeftMixedMarker", StringComparison.Ordinal));
        Assert.Contains(rightLines, line => line.Text.Contains("RightMixedMarker", StringComparison.Ordinal));
        Assert.All(leftLines, line =>
            Assert.True(line.X1 >= leftX - 1 && line.X2 <= leftRightX + 1.5,
                $"Expected mixed left-column line '{line.Text}' to stay inside the left column frame."));
        Assert.All(rightLines, line =>
            Assert.True(line.X1 >= rightX - 1.5 && line.X2 <= rightRightX + 1.5,
                $"Expected mixed right-column line '{line.Text}' to stay inside the right column frame."));
        AssertReadableTextRhythm(leftLines.Where(line => line.Text.Contains("Mixed", StringComparison.Ordinal)).ToList(), "mixed left column");
        AssertReadableTextRhythm(rightLines.Where(line => line.Text.Contains("Mixed", StringComparison.Ordinal)).ToList(), "mixed right column");
        AssertNoCrampedBaselines(pageLines, "mixed Word-like flow");
        AssertNoSameBaselineTextCollisions(page, "mixed Word-like flow");
        AssertNoAmbiguousSameBaselineRunGaps(page, "mixed Word-like flow");
    }

    [Fact]
    public void WordLikeLineItemTable_KeepsReadableColumnsWithoutTemplateApi() {
        const double pageWidth = 595;
        const double margin = 56;
        var style = TableStyles.ListTable1Light();
        style.RightAlignNumeric = true;
        style.CellPaddingX = 5;
        style.CellPaddingY = 5.5;
        style.ColumnWidthWeights = new List<double> { 0.45, 4.2, 1.15, 0.8, 1.2 };
        style.ColumnMinWidthPoints = new List<double?> { 22, 185, 62, 34, 68 };
        style.FooterRowCount = 1;
        style.FooterSeparatorColor = PdfColor.Black;
        style.FooterSeparatorWidth = 0.9;
        style.SpacingBefore = 8;
        style.SpacingAfter = 10;

        byte[] bytes = PdfDoc.Create(new PdfOptions {
                PageWidth = pageWidth,
                PageHeight = 760,
                MarginLeft = margin,
                MarginRight = margin,
                MarginTop = 56,
                MarginBottom = 56
            })
            .Theme(PdfTheme.WordLike())
            .H1("LineItemGate")
            .Paragraph(p => p.Text("A generic Word-like line item table protects table rhythm without adding invoice-specific engine APIs."))
            .Table(new[] {
                new[] { "#", "Product", "UnitPrice", "Qty", "Total" },
                new[] { "1", "MonitoringSeats", "31.80", "2", "63.60" },
                new[] { "2", "RadioInsulamPluviae", "62.57", "7", "437.99" },
                new[] { "3", "Long Wrapping Service Description For Column Rhythm", "42.50", "5", "212.50" },
                new[] { "4", "RexMaximeDixitque", "22.75", "5", "113.75" },
                new[] { "5", "ActumExemplumPrinceps", "6.41", "8", "51.28" },
                new[] { "6", "CustodiPuella", "79.05", "8", "632.40" },
                new[] { "", "TotalDue", "", "", "1499.12" }
            }, style: style)
            .Paragraph(p => p.Text("LineItemGateEnd"), PdfAlign.Center)
            .ToBytes();

        using var pdf = PdfDocument.Open(new MemoryStream(bytes));
        var page = pdf.GetPage(1);
        var allLines = GetVisualTextLines(page, 0, pageWidth);
        double tableStartY = FindWordStartY(page, "Product");
        double tableEndY = FindWordStartY(page, "TotalDue");
        var tableLines = allLines
            .Where(line => line.BaselineY <= tableStartY + 1 && line.BaselineY >= tableEndY - 1)
            .ToList();

        Assert.Equal(1, pdf.NumberOfPages);
        Assert.Contains(tableLines, line => line.Text.Contains("MonitoringSeats", StringComparison.Ordinal));
        Assert.Contains(tableLines, line => line.Text.Contains("LongWrappingService", StringComparison.Ordinal));
        Assert.Contains(tableLines, line => line.Text.Contains("Description", StringComparison.Ordinal));
        Assert.Contains(tableLines, line => line.Text.Contains("TotalDue", StringComparison.Ordinal));

        Assert.True(FindWordEndX(page, "MonitoringSeats") < FindWordStartX(page, "31.80") - 10,
            "Expected product text to end with visible space before the unit price column.");
        double monitoringBaselineY = FindWordStartY(page, "MonitoringSeats");
        Assert.True(FindWordEndXOnBaseline(page, "31.80", monitoringBaselineY) < FindWordStartXOnBaseline(page, "2", monitoringBaselineY) - 10,
            "Expected unit price text to stay separated from the quantity column.");
        Assert.True(FindWordEndXOnBaseline(page, "2", monitoringBaselineY) < FindWordStartXOnBaseline(page, "63.60", monitoringBaselineY) - 10,
            "Expected quantity text to stay separated from the total column.");
        double footerBaselineY = FindWordStartY(page, "TotalDue");
        Assert.True(FindWordEndXOnBaseline(page, "TotalDue", footerBaselineY) < FindWordStartXOnBaseline(page, "1499.12", footerBaselineY) - 10,
            "Expected footer label and numeric summary to stay visibly separated.");
        Assert.True(FindWordEndX(page, "632.40") <= pageWidth - margin + 1,
            "Expected the rightmost total to stay inside the document margin.");
        Assert.True(FindWordEndX(page, "1499.12") <= pageWidth - margin + 1,
            "Expected the footer total to stay inside the document margin.");
        Assert.True(FindWordStartY(page, "LineItemGateEnd") < tableEndY - 20,
            "Expected following content to retain breathing room after the table.");

        AssertNoCrampedBaselines(tableLines, "generic line item table");
        AssertNoSameBaselineTextCollisions(page, "generic line item table");
        AssertNoAmbiguousSameBaselineRunGaps(page, "generic line item table");
    }

    [Fact]
    public void ShowcaseDashboard_KeepsReadableGenericLayoutGeometry() {
        const double pageWidth = 841.89;
        const double marginLeft = 42;
        const double marginRight = 42;
        const double bodyGap = 18;
        double contentWidth = pageWidth - marginLeft - marginRight;
        double bodyColumnWidth = contentWidth - bodyGap;
        double leftColumnRightX = marginLeft + (bodyColumnWidth * 0.58);
        double rightColumnX = leftColumnRightX + bodyGap;
        double rightColumnRightX = pageWidth - marginRight;

        byte[] bytes = PdfDocRasterVisualBaselineTests.CreateShowcaseDashboard();

        using var pdf = PdfDocument.Open(new MemoryStream(bytes));
        var page = pdf.GetPage(1);
        var leftLines = GetVisualTextLines(page, marginLeft - 1, leftColumnRightX + 1);
        var rightLines = GetVisualTextLines(page, rightColumnX - 1, rightColumnRightX + 1);

        double titleY = FindWordStartY(page, "Quarterly");
        double leadY = FindWordStartY(page, "single-page");
        double firstMetricY = FindWordStartY(page, "92%");
        double secondMetricY = FindWordStartY(page, "1.8h");
        double thirdMetricY = FindWordStartY(page, "34");
        double fourthMetricY = FindWordStartY(page, "Critical");
        double leftHeadingY = FindWordStartY(page, "Delivery");
        double rightHeadingY = FindWordStartY(page, "Narrative");
        double riskHeaderY = FindWordStartY(page, "Area");
        double riskBodyY = FindWordStartY(page, "PDF");
        double decisionHeaderY = FindWordStartY(page, "Next");
        double decisionBodyY = FindWordStartY(page, "fixtures");

        Assert.Equal(1, pdf.NumberOfPages);
        Assert.True(titleY - leadY >= 21, $"Expected dashboard lead copy to sit comfortably below the title. Gap: {titleY - leadY:0.##}pt.");
        Assert.True(leadY - firstMetricY >= 24, $"Expected metric cards to start after the lead copy with visible breathing room. Gap: {leadY - firstMetricY:0.##}pt.");
        Assert.True(Math.Abs(firstMetricY - secondMetricY) <= 0.5, "Expected metric card values to align on the same visual baseline.");
        Assert.True(Math.Abs(firstMetricY - thirdMetricY) <= 0.5, "Expected metric card values to align on the same visual baseline.");
        Assert.True(firstMetricY - fourthMetricY >= 12, "Expected the long fourth metric label to wrap below its value instead of colliding with neighboring cards.");
        Assert.True(leftHeadingY - riskHeaderY >= 175, $"Expected the trend drawing to reserve vertical space before the risk table. Gap: {leftHeadingY - riskHeaderY:0.##}pt.");
        Assert.True(Math.Abs(leftHeadingY - rightHeadingY) <= 2, "Expected the two body columns to start on the same visual row.");
        Assert.True(riskHeaderY - riskBodyY >= 14, $"Expected dashboard table header and first row to retain readable rhythm. Gap: {riskHeaderY - riskBodyY:0.##}pt.");
        Assert.True(decisionHeaderY - decisionBodyY >= 14, $"Expected decision table header and first row to retain readable rhythm. Gap: {decisionHeaderY - decisionBodyY:0.##}pt.");

        Assert.Contains(leftLines, line => line.Text.Contains("Deliverytrend", StringComparison.Ordinal));
        Assert.Contains(rightLines, line => line.Text.Contains("Narrative", StringComparison.Ordinal));
        Assert.True(FindWordStartX(page, "Narrative") >= rightColumnX - 1,
            $"Expected the right dashboard column to start after the gutter. Narrative x: {FindWordStartX(page, "Narrative"):0.##}, expected at least {rightColumnX:0.##}.");
        Assert.True(FindWordEndX(page, "Roadmap") <= leftColumnRightX + 1,
            "Expected the left risk table owner column to stay inside the left dashboard column.");
        Assert.True(FindWordEndX(page, "slices") <= rightColumnRightX + 1,
            "Expected the right decision table text to stay inside the right dashboard column.");

        AssertNoCrampedBaselines(leftLines, "showcase dashboard left column");
        AssertNoCrampedBaselines(rightLines, "showcase dashboard right column");
        AssertNoSameBaselineTextCollisions(page, "showcase dashboard");
        AssertNoAmbiguousSameBaselineRunGaps(page, "showcase dashboard");
    }

    [Fact]
    public void RowStyle_DefaultsApplyGutterAndOuterRhythm() {
        const double pageWidth = 360;
        const double margin = 30;
        const double gutter = 36;
        const double fontSize = 10;
        double contentWidth = pageWidth - margin - margin;
        double expectedRightColumnX = margin + ((contentWidth - gutter) / 2) + gutter;
        var tightParagraph = new PdfParagraphStyle {
            SpacingAfter = 0
        };

        byte[] bytes = PdfDoc.Create(new PdfOptions {
                PageWidth = pageWidth,
                PageHeight = 220,
                MarginLeft = margin,
                MarginRight = margin,
                MarginTop = 30,
                MarginBottom = 30,
                DefaultFont = PdfStandardFont.Helvetica,
                DefaultFontSize = fontSize,
                DefaultRowStyle = new PdfRowStyle {
                    Gap = gutter,
                    SpacingBefore = 14,
                    SpacingAfter = 16
                }
            })
            .Compose(document =>
                document.Page(page =>
                    page.Content(content =>
                        content
                            .Column(column => column.Item().Paragraph(p => p.Text("BeforeRow"), style: tightParagraph))
                            .Row(row => row
                                .Column(50, column => column.Paragraph(p => p.Text("LeftDefaultGap"), style: tightParagraph))
                                .Column(50, column => column.Paragraph(p => p.Text("RightDefaultGap"), style: tightParagraph)))
                            .Column(column => column.Item().Paragraph(p => p.Text("AfterRow"), style: tightParagraph)))))
            .ToBytes();

        using var pdf = PdfDocument.Open(new MemoryStream(bytes));
        var page = pdf.GetPage(1);

        double beforeY = FindWordStartY(page, "BeforeRow");
        double leftY = FindWordStartY(page, "LeftDefaultGap");
        double afterY = FindWordStartY(page, "AfterRow");
        double rightX = FindWordStartX(page, "RightDefaultGap");

        Assert.True(rightX >= expectedRightColumnX - 1,
            $"Expected default row style to preserve a {gutter:0.##}pt gutter. Right column started at {rightX:0.##}, expected at least {expectedRightColumnX:0.##}.");
        Assert.True(beforeY - leftY >= 27,
            $"Expected row spacing before to create visible breathing room. Baseline gap: {beforeY - leftY:0.##}pt.");
        Assert.True(leftY - afterY >= 29,
            $"Expected row spacing after to create visible breathing room. Baseline gap: {leftY - afterY:0.##}pt.");
    }

    [Fact]
    public void RowStyle_KeepTogetherMovesRowToNextPage() {
        var tightParagraph = new PdfParagraphStyle {
            LineHeight = 1,
            SpacingAfter = 0
        };

        byte[] bytes = PdfDoc.Create(new PdfOptions {
                PageWidth = 260,
                PageHeight = 170,
                MarginLeft = 24,
                MarginRight = 24,
                MarginTop = 20,
                MarginBottom = 20,
                DefaultFont = PdfStandardFont.Helvetica,
                DefaultFontSize = 10,
                DefaultRowStyle = new PdfRowStyle {
                    Gap = 18,
                    KeepTogether = true
                }
            })
            .Compose(document =>
                document.Page(page =>
                    page.Content(content =>
                        content
                            .Column(column => {
                                for (int i = 0; i < 10; i++) {
                                    column.Item().Paragraph(p => p.Text("IntroLine" + i.ToString(CultureInfo.InvariantCulture)), style: tightParagraph);
                                }
                            })
                            .Row(row => row
                                .Column(50, column => column.Paragraph(p => p.Text("KeptRowLeft has enough text to wrap across several lines inside the first column."), style: tightParagraph))
                                .Column(50, column => column.Paragraph(p => p.Text("KeptRowRight should travel with the left column instead of starting on the first page."), style: tightParagraph))))))
            .ToBytes();

        using var pdf = PdfDocument.Open(new MemoryStream(bytes));

        Assert.True(pdf.NumberOfPages >= 2);
        Assert.DoesNotContain("KeptRowLeft", pdf.GetPage(1).Text, StringComparison.Ordinal);
        Assert.Contains("KeptRowLeft", pdf.GetPage(2).Text, StringComparison.Ordinal);
        Assert.Contains("KeptRowRight", pdf.GetPage(2).Text, StringComparison.Ordinal);
    }

    [Fact]
    public void RowStyle_KeepTogetherRejectsRowsTallerThanPageFrame() {
        string longText = string.Join(" ", Enumerable.Repeat("TooTallRowContent", 80));

        var exception = Assert.Throws<ArgumentException>(() =>
            PdfDoc.Create(new PdfOptions {
                    PageWidth = 220,
                    PageHeight = 120,
                    MarginLeft = 20,
                    MarginRight = 20,
                    MarginTop = 20,
                    MarginBottom = 20,
                    DefaultFont = PdfStandardFont.Helvetica,
                    DefaultFontSize = 10,
                    DefaultRowStyle = new PdfRowStyle {
                        KeepTogether = true
                    }
                })
                .Compose(document =>
                    document.Page(page =>
                        page.Content(content =>
                            content.Row(row =>
                                row.Column(100, column =>
                                    column.Paragraph(p => p.Text(longText)))))))
                .ToBytes());

        Assert.Contains("Row height exceeds the available page content height.", exception.Message, StringComparison.Ordinal);
    }

    [Fact]
    public void HorizontalRule_RejectsInvalidLayoutValues() {
        var thicknessException = Assert.Throws<ArgumentException>(() =>
            PdfDoc.Create()
                .HR(thickness: 0));

        Assert.Contains("Horizontal rule thickness must be a positive finite value.", thicknessException.Message, StringComparison.Ordinal);

        var spacingBeforeException = Assert.Throws<ArgumentException>(() =>
            PdfDoc.Create()
                .HR(spacingBefore: -1));

        Assert.Contains("Horizontal rule spacing before must be a non-negative finite value.", spacingBeforeException.Message, StringComparison.Ordinal);

        var spacingAfterException = Assert.Throws<ArgumentException>(() =>
            PdfDoc.Create()
                .HR(spacingAfter: double.PositiveInfinity));

        Assert.Contains("Horizontal rule spacing after must be a non-negative finite value.", spacingAfterException.Message, StringComparison.Ordinal);

        var columnException = Assert.Throws<ArgumentException>(() =>
            PdfDoc.Create()
                .Compose(compose =>
                    compose.Page(page =>
                        page.Content(content =>
                            content.Row(row =>
                                row.Column(100, column =>
                                    column.HR(thickness: double.NaN)))))));

        Assert.Contains("Horizontal rule thickness must be a positive finite value.", columnException.Message, StringComparison.Ordinal);
    }

    [Fact]
    public void HorizontalRule_RejectsHeightExceedingContentArea() {
        var exception = Assert.Throws<ArgumentException>(() =>
            PdfDoc.Create(new PdfOptions {
                    PageWidth = 220,
                    PageHeight = 140,
                    MarginLeft = 20,
                    MarginRight = 20,
                    MarginTop = 20,
                    MarginBottom = 20
                })
                .HR(thickness: 110, spacingBefore: 0, spacingAfter: 0)
                .ToBytes());

        Assert.Contains("Horizontal rule height exceeds the available page content height.", exception.Message, StringComparison.Ordinal);
    }

    [Fact]
    public void Paragraph_RejectsInvalidSpacing() {
        var spacingBeforeException = Assert.Throws<ArgumentException>(() =>
            new PdfParagraphStyle {
                SpacingBefore = -1
            });

        Assert.Contains("Paragraph spacing before must be a non-negative finite value.", spacingBeforeException.Message, StringComparison.Ordinal);

        var spacingAfterException = Assert.Throws<ArgumentException>(() =>
            new PdfParagraphStyle {
                SpacingAfter = double.PositiveInfinity
            });

        Assert.Contains("Paragraph spacing after must be a non-negative finite value.", spacingAfterException.Message, StringComparison.Ordinal);

        var lineHeightException = Assert.Throws<ArgumentException>(() =>
            new PdfParagraphStyle {
                LineHeight = 0
            });

        Assert.Contains("Paragraph line height must be a positive finite value.", lineHeightException.Message, StringComparison.Ordinal);

        var leftIndentException = Assert.Throws<ArgumentException>(() =>
            new PdfParagraphStyle {
                LeftIndent = -1
            });

        Assert.Contains("Paragraph left indent must be a non-negative finite value.", leftIndentException.Message, StringComparison.Ordinal);

        var rightIndentException = Assert.Throws<ArgumentException>(() =>
            new PdfParagraphStyle {
                RightIndent = double.NaN
            });

        Assert.Contains("Paragraph right indent must be a non-negative finite value.", rightIndentException.Message, StringComparison.Ordinal);

        var firstLineIndentException = Assert.Throws<ArgumentException>(() =>
            new PdfParagraphStyle {
                FirstLineIndent = double.PositiveInfinity
            });

        Assert.Contains("Paragraph first line indent must be a finite value.", firstLineIndentException.Message, StringComparison.Ordinal);

        var tabStopException = Assert.Throws<ArgumentException>(() =>
            new PdfParagraphStyle {
                DefaultTabStopWidth = double.NaN
            });

        Assert.Contains("Paragraph default tab stop width must be a positive finite value.", tabStopException.Message, StringComparison.Ordinal);

        var hangingOutsideFrameException = Assert.Throws<ArgumentException>(() =>
            PdfDoc.Create(new PdfOptions {
                    PageWidth = 160,
                    MarginLeft = 20,
                    MarginRight = 20
                })
                .Paragraph(p => p.Text("Invalid hanging indent"), style: new PdfParagraphStyle {
                    LeftIndent = 10,
                    FirstLineIndent = -12
                })
                .ToBytes());

        Assert.Contains("Paragraph first line indent must not move text outside the left content frame.", hangingOutsideFrameException.Message, StringComparison.Ordinal);

        var firstLineWidthException = Assert.Throws<ArgumentException>(() =>
            PdfDoc.Create(new PdfOptions {
                    PageWidth = 160,
                    MarginLeft = 20,
                    MarginRight = 20
                })
                .Paragraph(p => p.Text("Invalid first line width"), style: new PdfParagraphStyle {
                    FirstLineIndent = 120
                })
                .ToBytes());

        Assert.Contains("Paragraph first line indent must leave a positive text width.", firstLineWidthException.Message, StringComparison.Ordinal);

        Assert.Throws<ArgumentException>(() =>
            PdfDoc.Create(new PdfOptions {
                    PageWidth = 120,
                    MarginLeft = 20,
                    MarginRight = 20
                })
                .Paragraph(p => p.Text("Invalid text width"), style: new PdfParagraphStyle {
                    LeftIndent = 50,
                    RightIndent = 40
                })
                .ToBytes());
    }

    [Fact]
    public void VectorRectangle_RendersFillAndStrokeOperators() {
        byte[] bytes = PdfDoc.Create(new PdfOptions {
                PageWidth = 240,
                PageHeight = 180,
                MarginLeft = 20,
                MarginRight = 20,
                MarginTop = 20,
                MarginBottom = 20
            })
            .Rectangle(
                width: 100,
                height: 36,
                strokeColor: PdfColor.FromRgb(26, 51, 77),
                strokeWidth: 2.5,
                fillColor: PdfColor.FromRgb(204, 179, 153),
                align: PdfAlign.Center,
                spacingBefore: 4,
                spacingAfter: 6)
            .ToBytes();

        string content = Encoding.ASCII.GetString(bytes);

        Assert.Contains("0.8 0.702 0.6 rg", content);
        Assert.Contains("0.102 0.2 0.302 RG", content);
        Assert.Contains("2.5 w", content);
        Assert.Contains("70 124 100 36 re B", content);
    }

    [Fact]
    public void TableBodyText_ResetsFillColorAfterColoredHeader() {
        byte[] bytes = PdfDoc.Create(new PdfOptions {
                DefaultFont = PdfStandardFont.Helvetica,
                DefaultFontSize = 10
            })
            .Table(new[] {
                new[] { "Signal", "Evidence" },
                new[] { "DMARC", "Body text must stay readable after a white header." }
            }, style: new PdfTableStyle {
                HeaderFill = PdfColor.FromRgb(32, 76, 120),
                HeaderTextColor = PdfColor.White,
                TextColor = null,
                RowStripeFill = PdfColor.FromRgb(248, 250, 252)
            })
            .ToBytes();

        string content = Encoding.ASCII.GetString(bytes);
        int headerTextColorIndex = content.IndexOf("1 1 1 rg", StringComparison.Ordinal);
        int bodyTextIndex = content.IndexOf("<444D415243> Tj", StringComparison.Ordinal);
        int bodyTextColorIndex = content.LastIndexOf("0 0 0 rg", bodyTextIndex, StringComparison.Ordinal);

        Assert.True(headerTextColorIndex >= 0, "The header should use the configured white text color.");
        Assert.True(bodyTextIndex > headerTextColorIndex, "The body cell should be written after the colored header.");
        Assert.True(bodyTextColorIndex > headerTextColorIndex, "Body cells without an explicit color should reset fill to black.");
    }

    [Fact]
    public void TableStyle_CanControlGenericHeaderBodyAndFooterTypography() {
        var style = TableStyles.Minimal();
        style.FontSize = 8;
        style.LineHeight = 1.1;
        style.HeaderFontSize = 12;
        style.FooterFontSize = 10;
        style.FooterRowCount = 1;

        byte[] bytes = PdfDoc.Create(new PdfOptions {
                DefaultFont = PdfStandardFont.Helvetica,
                DefaultFontSize = 9
            })
            .Table(new[] {
                new[] { "Column", "Status" },
                new[] { "BodyRow", "Readable" },
                new[] { "Total", "Footer" }
            }, style: style)
            .ToBytes();

        string content = Encoding.ASCII.GetString(bytes);

        Assert.Contains("/F2 12 Tf", content);
        Assert.Contains("/F1 8 Tf", content);
        Assert.Contains("/F2 10 Tf", content);
    }

    [Fact]
    public void TableStyle_UsesConfiguredCellLineHeight() {
        var options = new PdfOptions {
            PageWidth = 260,
            PageHeight = 200,
            MarginLeft = 30,
            MarginRight = 30,
            MarginTop = 30,
            MarginBottom = 30,
            DefaultFont = PdfStandardFont.Helvetica,
            DefaultFontSize = 10
        };
        byte[] defaultBytes = CreateTableLineHeightProbe(options, null, useRowColumnFlow: false);
        byte[] looseBytes = CreateTableLineHeightProbe(options, 2.0, useRowColumnFlow: false);

        using var defaultPdf = PdfDocument.Open(new MemoryStream(defaultBytes));
        using var loosePdf = PdfDocument.Open(new MemoryStream(looseBytes));
        var defaultPage = defaultPdf.GetPage(1);
        var loosePage = loosePdf.GetPage(1);

        double defaultGap = FindWordStartY(defaultPage, "FirstLine") - FindWordStartY(defaultPage, "SecondLine");
        double looseGap = FindWordStartY(loosePage, "FirstLine") - FindWordStartY(loosePage, "SecondLine");

        Assert.True(looseGap > defaultGap + 4, $"Expected larger table line height to increase wrapped cell baseline gap. Default: {defaultGap:0.##}, loose: {looseGap:0.##}.");
    }

    [Fact]
    public void RowColumnTableStyle_UsesConfiguredCellLineHeight() {
        var options = new PdfOptions {
            PageWidth = 260,
            PageHeight = 200,
            MarginLeft = 30,
            MarginRight = 30,
            MarginTop = 30,
            MarginBottom = 30,
            DefaultFont = PdfStandardFont.Helvetica,
            DefaultFontSize = 10
        };
        byte[] defaultBytes = CreateTableLineHeightProbe(options, null, useRowColumnFlow: true);
        byte[] looseBytes = CreateTableLineHeightProbe(options, 2.0, useRowColumnFlow: true);

        using var defaultPdf = PdfDocument.Open(new MemoryStream(defaultBytes));
        using var loosePdf = PdfDocument.Open(new MemoryStream(looseBytes));
        var defaultPage = defaultPdf.GetPage(1);
        var loosePage = loosePdf.GetPage(1);

        double defaultGap = FindWordStartY(defaultPage, "FirstLine") - FindWordStartY(defaultPage, "SecondLine");
        double looseGap = FindWordStartY(loosePage, "FirstLine") - FindWordStartY(loosePage, "SecondLine");

        Assert.True(looseGap > defaultGap + 4, $"Expected larger row-column table line height to increase wrapped cell baseline gap. Default: {defaultGap:0.##}, loose: {looseGap:0.##}.");
    }

    [Fact]
    public void TableStyle_UsesConfiguredCellPaddingSides() {
        var options = new PdfOptions {
            PageWidth = 260,
            PageHeight = 180,
            MarginLeft = 30,
            MarginRight = 30,
            MarginTop = 30,
            MarginBottom = 30,
            DefaultFont = PdfStandardFont.Helvetica,
            DefaultFontSize = 10
        };
        byte[] defaultBytes = CreateTablePaddingProbe(options, useRowColumnFlow: false, useSidePadding: false);
        byte[] paddedBytes = CreateTablePaddingProbe(options, useRowColumnFlow: false, useSidePadding: true);

        using var defaultPdf = PdfDocument.Open(new MemoryStream(defaultBytes));
        using var paddedPdf = PdfDocument.Open(new MemoryStream(paddedBytes));
        var defaultPage = defaultPdf.GetPage(1);
        var paddedPage = paddedPdf.GetPage(1);

        double defaultX = FindWordStartX(defaultPage, "PadMarker");
        double paddedX = FindWordStartX(paddedPage, "PadMarker");
        double defaultY = FindWordStartY(defaultPage, "PadMarker");
        double paddedY = FindWordStartY(paddedPage, "PadMarker");

        Assert.True(paddedX > defaultX + 14, $"Expected left cell padding to move text right. Default x: {defaultX:0.##}, padded x: {paddedX:0.##}.");
        Assert.True(defaultY > paddedY + 10, $"Expected top cell padding to move text down. Default y: {defaultY:0.##}, padded y: {paddedY:0.##}.");
    }

    [Fact]
    public void RowColumnTableStyle_UsesConfiguredCellPaddingSides() {
        var options = new PdfOptions {
            PageWidth = 260,
            PageHeight = 180,
            MarginLeft = 30,
            MarginRight = 30,
            MarginTop = 30,
            MarginBottom = 30,
            DefaultFont = PdfStandardFont.Helvetica,
            DefaultFontSize = 10
        };
        byte[] defaultBytes = CreateTablePaddingProbe(options, useRowColumnFlow: true, useSidePadding: false);
        byte[] paddedBytes = CreateTablePaddingProbe(options, useRowColumnFlow: true, useSidePadding: true);

        using var defaultPdf = PdfDocument.Open(new MemoryStream(defaultBytes));
        using var paddedPdf = PdfDocument.Open(new MemoryStream(paddedBytes));
        var defaultPage = defaultPdf.GetPage(1);
        var paddedPage = paddedPdf.GetPage(1);

        double defaultX = FindWordStartX(defaultPage, "PadMarker");
        double paddedX = FindWordStartX(paddedPage, "PadMarker");
        double defaultY = FindWordStartY(defaultPage, "PadMarker");
        double paddedY = FindWordStartY(paddedPage, "PadMarker");

        Assert.True(paddedX > defaultX + 14, $"Expected row-column left cell padding to move text right. Default x: {defaultX:0.##}, padded x: {paddedX:0.##}.");
        Assert.True(defaultY > paddedY + 10, $"Expected row-column top cell padding to move text down. Default y: {defaultY:0.##}, padded y: {paddedY:0.##}.");
    }

    [Fact]
    public void TableStyle_CanDisableHeaderAndFooterBoldWithoutChangingDocumentFont() {
        var style = TableStyles.Minimal();
        style.HeaderBold = false;
        style.FooterBold = false;
        style.FooterRowCount = 1;

        byte[] bytes = PdfDoc.Create(new PdfOptions {
                DefaultFont = PdfStandardFont.Helvetica,
                DefaultFontSize = 9
            })
            .Table(new[] {
                new[] { "PlainHeader", "Status" },
                new[] { "BodyRow", "Readable" },
                new[] { "PlainFooter", "Ready" }
            }, style: style)
            .ToBytes();

        string content = Encoding.ASCII.GetString(bytes);

        Assert.Contains("/F1 9 Tf", content);
        Assert.DoesNotContain("/F2 9 Tf", content);
    }

    [Fact]
    public void VectorRoundedRectangle_RendersBezierCornersFromSharedShapeDescriptor() {
        byte[] bytes = PdfDoc.Create(new PdfOptions {
                PageWidth = 240,
                PageHeight = 180,
                MarginLeft = 20,
                MarginRight = 20,
                MarginTop = 20,
                MarginBottom = 20
            })
            .RoundedRectangle(
                width: 100,
                height: 36,
                cornerRadius: 8,
                strokeColor: PdfColor.FromRgb(26, 51, 77),
                strokeWidth: 2,
                fillColor: PdfColor.FromRgb(204, 179, 153),
                align: PdfAlign.Center,
                spacingBefore: 4,
                spacingAfter: 6)
            .ToBytes();

        string content = Encoding.ASCII.GetString(bytes);

        Assert.Contains("0.8 0.702 0.6 rg", content);
        Assert.Contains("0.102 0.2 0.302 RG", content);
        Assert.Contains("2 w", content);
        Assert.Contains("78 124 m", content);
        Assert.Contains("162 124 l", content);
        Assert.Contains("166.418 124 170 127.582 170 132 c", content);
        Assert.Contains("70 127.582 73.582 124 78 124 c h B", content);
    }

    [Fact]
    public void VectorLine_RendersStrokeOperatorFromSharedShapeDescriptor() {
        byte[] bytes = PdfDoc.Create(new PdfOptions {
                PageWidth = 240,
                PageHeight = 180,
                MarginLeft = 20,
                MarginRight = 20,
                MarginTop = 20,
                MarginBottom = 20
            })
            .Line(
                x1: 0,
                y1: 0,
                x2: 100,
                y2: 40,
                strokeColor: PdfColor.FromRgb(51, 102, 153),
                strokeWidth: 2,
                align: PdfAlign.Center,
                spacingBefore: 4,
                spacingAfter: 6,
                strokeDashStyle: OfficeStrokeDashStyle.Dash)
            .ToBytes();

        string content = Encoding.ASCII.GetString(bytes);

        Assert.Contains("0.2 0.4 0.6 RG", content);
        Assert.Contains("2 w", content);
        Assert.Contains("[6 3] 0 d", content);
        Assert.Contains("70 160 m 170 120 l S", content);
    }

    [Fact]
    public void VectorLine_RendersConfiguredStrokeCapAndJoin() {
        byte[] bytes = PdfDoc.Create(new PdfOptions {
                PageWidth = 240,
                PageHeight = 180,
                MarginLeft = 20,
                MarginRight = 20,
                MarginTop = 20,
                MarginBottom = 20
            })
            .Line(
                x1: 0,
                y1: 0,
                x2: 100,
                y2: 0,
                strokeColor: PdfColor.FromRgb(51, 102, 153),
                strokeWidth: 3,
                align: PdfAlign.Center,
                spacingBefore: 4,
                spacingAfter: 6,
                strokeLineCap: OfficeStrokeLineCap.Square,
                strokeLineJoin: OfficeStrokeLineJoin.Bevel)
            .ToBytes();

        string content = Encoding.ASCII.GetString(bytes);

        Assert.Contains("0.2 0.4 0.6 RG", content);
        Assert.Contains("3 w", content);
        Assert.Contains("2 J", content);
        Assert.Contains("2 j", content);
        Assert.Contains("70 160 m 170 160 l S", content);
    }

    [Fact]
    public void VectorShape_UsesSharedOfficeDrawingShapeDescriptor() {
        var shape = OfficeShape.Rectangle(90, 24);
        shape.FillColor = OfficeColor.WhiteSmoke;
        shape.StrokeColor = OfficeColor.SteelBlue;
        shape.StrokeWidth = 1.5;

        byte[] bytes = PdfDoc.Create(new PdfOptions {
                PageWidth = 220,
                PageHeight = 160,
                MarginLeft = 30,
                MarginRight = 30,
                MarginTop = 30,
                MarginBottom = 30
            })
            .Shape(shape, align: PdfAlign.Right)
            .ToBytes();

        string content = Encoding.ASCII.GetString(bytes);

        Assert.Contains("0.961 0.961 0.961 rg", content);
        Assert.Contains("0.275 0.51 0.706 RG", content);
        Assert.Contains("1.5 w", content);
        Assert.Contains("100 106 90 24 re B", content);
    }

    [Fact]
    public void VectorShape_RendersSharedTransformAsGraphicsStateMatrix() {
        var shape = OfficeShape.Rectangle(40, 20);
        shape.FillColor = OfficeColor.WhiteSmoke;
        shape.StrokeColor = OfficeColor.SteelBlue;
        shape.StrokeWidth = 1.5;
        shape.Transform = OfficeTransform.Translate(10, 5);

        byte[] bytes = PdfDoc.Create(new PdfOptions {
                PageWidth = 220,
                PageHeight = 160,
                MarginLeft = 30,
                MarginRight = 30,
                MarginTop = 30,
                MarginBottom = 30
            })
            .Shape(shape)
            .ToBytes();

        string content = Encoding.ASCII.GetString(bytes);

        Assert.Contains("q\n1 0 0 -1 40 125 cm", content);
        Assert.Contains("0.961 0.961 0.961 rg", content);
        Assert.Contains("0.275 0.51 0.706 RG", content);
        Assert.Contains("1.5 w", content);
        Assert.Contains("0 0 40 20 re B", content);
    }

    [Fact]
    public void VectorShape_RendersSharedOpacityAsExtGStateResource() {
        var shape = OfficeShape.Rectangle(90, 24);
        shape.FillColor = OfficeColor.WhiteSmoke;
        shape.StrokeColor = OfficeColor.SteelBlue;
        shape.StrokeWidth = 1.5;
        shape.FillOpacity = 0.35;
        shape.StrokeOpacity = 0.75;

        byte[] bytes = PdfDoc.Create(new PdfOptions {
                PageWidth = 220,
                PageHeight = 160,
                MarginLeft = 30,
                MarginRight = 30,
                MarginTop = 30,
                MarginBottom = 30
            })
            .Shape(shape)
            .ToBytes();

        string content = Encoding.ASCII.GetString(bytes);

        Assert.Contains("<< /Type /ExtGState /ca 0.35 /CA 0.75 >>", content);
        Assert.Contains("/ExtGState << /GS1 ", content);
        Assert.Contains("q\n/GS1 gs\nq\n0.961 0.961 0.961 rg", content);
        Assert.Contains("30 106 90 24 re B", content);
    }

    [Fact]
    public void VectorShape_RendersSharedClipPathBeforePainting() {
        var shape = OfficeShape.Rectangle(90, 40);
        shape.FillColor = OfficeColor.WhiteSmoke;
        shape.StrokeColor = OfficeColor.SteelBlue;
        shape.StrokeWidth = 1.5;
        shape.ClipPath = OfficeClipPath.Rectangle(45, 20);

        byte[] bytes = PdfDoc.Create(new PdfOptions {
                PageWidth = 220,
                PageHeight = 160,
                MarginLeft = 30,
                MarginRight = 30,
                MarginTop = 30,
                MarginBottom = 30
            })
            .Shape(shape)
            .ToBytes();

        string content = Encoding.ASCII.GetString(bytes);

        Assert.Contains("q\n30 110 45 20 re W n\nq\n0.961 0.961 0.961 rg", content);
        Assert.Contains("30 90 90 40 re B", content);
    }

    [Fact]
    public void VectorShape_RendersSharedClipPathInsideTransformGraphicsState() {
        var shape = OfficeShape.Rectangle(80, 40);
        shape.FillColor = OfficeColor.WhiteSmoke;
        shape.ClipPath = OfficeClipPath.RoundedRectangle(40, 20, 6);
        shape.Transform = OfficeTransform.Translate(10, 5);

        byte[] bytes = PdfDoc.Create(new PdfOptions {
                PageWidth = 220,
                PageHeight = 160,
                MarginLeft = 30,
                MarginRight = 30,
                MarginTop = 30,
                MarginBottom = 30
            })
            .Shape(shape)
            .ToBytes();

        string content = Encoding.ASCII.GetString(bytes);

        Assert.Contains("q\n1 0 0 -1 40 125 cm", content);
        Assert.Contains("6 0 m 34 0 l", content);
        Assert.Contains("W n\n0.961 0.961 0.961 rg", content);
        Assert.Contains("0 0 80 40 re f", content);
    }

    [Fact]
    public void VectorShape_RendersSharedLinearGradientAsAxialShadingResource() {
        var shape = OfficeShape.Rectangle(90, 24);
        shape.FillColor = OfficeColor.Red;
        shape.FillGradient = OfficeLinearGradient.Horizontal(OfficeColor.SteelBlue, OfficeColor.WhiteSmoke);
        shape.StrokeColor = OfficeColor.Black;
        shape.StrokeWidth = 1.25;

        byte[] bytes = PdfDoc.Create(new PdfOptions {
                PageWidth = 220,
                PageHeight = 160,
                MarginLeft = 30,
                MarginRight = 30,
                MarginTop = 30,
                MarginBottom = 30
            })
            .Shape(shape)
            .ToBytes();

        string content = Encoding.ASCII.GetString(bytes);

        Assert.Contains("/Shading << /SH1 ", content);
        Assert.Contains("/ShadingType 2 /ColorSpace /DeviceRGB /Coords [30 118 120 118]", content);
        Assert.Contains("/C0 [0.275 0.51 0.706] /C1 [0.961 0.961 0.961]", content);
        Assert.Contains("q\n30 106 90 24 re W n\n/SH1 sh\nQ", content);
        Assert.Contains("1.25 w", content);
        Assert.Contains("30 106 90 24 re S", content);
        Assert.DoesNotContain("1 0 0 rg", content, StringComparison.Ordinal);
    }

    [Fact]
    public void VectorShape_RendersSharedLinearGradientInsideTransformGraphicsState() {
        var shape = OfficeShape.Rectangle(40, 20);
        shape.FillGradient = OfficeLinearGradient.Vertical(OfficeColor.SteelBlue, OfficeColor.WhiteSmoke);
        shape.Transform = OfficeTransform.Translate(10, 5);

        byte[] bytes = PdfDoc.Create(new PdfOptions {
                PageWidth = 220,
                PageHeight = 160,
                MarginLeft = 30,
                MarginRight = 30,
                MarginTop = 30,
                MarginBottom = 30
            })
            .Shape(shape)
            .ToBytes();

        string content = Encoding.ASCII.GetString(bytes);

        Assert.Contains("/ShadingType 2 /ColorSpace /DeviceRGB /Coords [20 20 20 0]", content);
        Assert.Contains("q\n1 0 0 -1 40 125 cm", content);
        Assert.Contains("q\n0 0 40 20 re W n\n/SH1 sh\nQ", content);
    }

    [Fact]
    public void VectorShape_RendersSharedShadowBehindShapeGeometry() {
        var shape = OfficeShape.RoundedRectangle(90, 24, 6);
        shape.FillColor = OfficeColor.WhiteSmoke;
        shape.StrokeColor = OfficeColor.SteelBlue;
        shape.StrokeWidth = 1;
        shape.Shadow = new OfficeShadow(OfficeColor.Black, 0.22, 3, 4);

        byte[] bytes = PdfDoc.Create(new PdfOptions {
                PageWidth = 220,
                PageHeight = 160,
                MarginLeft = 30,
                MarginRight = 30,
                MarginTop = 30,
                MarginBottom = 30
            })
            .Shape(shape)
            .ToBytes();

        string content = Encoding.ASCII.GetString(bytes);

        Assert.Contains("/ExtGState << /GS1 ", content);
        Assert.Contains("/Type /ExtGState /ca 0.22 /CA 0.22", content);
        Assert.Contains("q\n/GS1 gs\nq\n0 0 0 rg", content);
        Assert.Contains("39 102", content);
        Assert.Contains("0.961 0.961 0.961 rg", content);
        Assert.True(content.IndexOf("/GS1 gs", StringComparison.Ordinal) < content.IndexOf("0.961 0.961 0.961 rg", StringComparison.Ordinal));
    }

    [Fact]
    public void VectorRectangle_RendersConfiguredDashStyle() {
        byte[] bytes = PdfDoc.Create(new PdfOptions {
                PageWidth = 220,
                PageHeight = 160,
                MarginLeft = 30,
                MarginRight = 30,
                MarginTop = 30,
                MarginBottom = 30
            })
            .Rectangle(
                width: 90,
                height: 24,
                strokeColor: PdfColor.FromRgb(51, 102, 153),
                strokeWidth: 2,
                align: PdfAlign.Left,
                strokeDashStyle: OfficeStrokeDashStyle.DashDot)
            .ToBytes();

        string content = Encoding.ASCII.GetString(bytes);

        Assert.Contains("0.2 0.4 0.6 RG", content);
        Assert.Contains("2 w", content);
        Assert.Contains("1 J", content);
        Assert.Contains("[6 3 2 3] 0 d", content);
        Assert.Contains("30 106 90 24 re S", content);
    }

    [Fact]
    public void VectorEllipse_RendersBezierPathFromSharedShapeDescriptor() {
        var shape = OfficeShape.Ellipse(80, 40);
        shape.FillColor = OfficeColor.WhiteSmoke;
        shape.StrokeColor = OfficeColor.SteelBlue;
        shape.StrokeWidth = 2;
        shape.StrokeDashStyle = OfficeStrokeDashStyle.Dot;

        byte[] bytes = PdfDoc.Create(new PdfOptions {
                PageWidth = 240,
                PageHeight = 180,
                MarginLeft = 20,
                MarginRight = 20,
                MarginTop = 20,
                MarginBottom = 20
            })
            .Shape(shape, align: PdfAlign.Center, spacingBefore: 4, spacingAfter: 6)
            .ToBytes();

        string content = Encoding.ASCII.GetString(bytes);

        Assert.Contains("0.961 0.961 0.961 rg", content);
        Assert.Contains("0.275 0.51 0.706 RG", content);
        Assert.Contains("2 w", content);
        Assert.Contains("1 J", content);
        Assert.Contains("[2 3] 0 d", content);
        Assert.Contains("160 140 m", content);
        Assert.Contains("160 151.046 142.091 160 120 160 c", content);
        Assert.Contains("142.091 120 160 128.954 160 140 c B", content);
    }

    [Fact]
    public void VectorPolygon_RendersClosedPathFromSharedShapeDescriptor() {
        var shape = OfficeShape.Polygon(
            new OfficePoint(0, 40),
            new OfficePoint(40, 0),
            new OfficePoint(80, 40));
        shape.FillColor = OfficeColor.WhiteSmoke;
        shape.StrokeColor = OfficeColor.SteelBlue;
        shape.StrokeWidth = 1.5;

        byte[] bytes = PdfDoc.Create(new PdfOptions {
                PageWidth = 240,
                PageHeight = 180,
                MarginLeft = 20,
                MarginRight = 20,
                MarginTop = 20,
                MarginBottom = 20
            })
            .Shape(shape, align: PdfAlign.Center, spacingBefore: 4, spacingAfter: 6)
            .ToBytes();

        string content = Encoding.ASCII.GetString(bytes);

        Assert.Contains("0.961 0.961 0.961 rg", content);
        Assert.Contains("0.275 0.51 0.706 RG", content);
        Assert.Contains("1.5 w", content);
        Assert.Contains("80 120 m", content);
        Assert.Contains("120 160 l", content);
        Assert.Contains("160 120 l", content);
        Assert.Contains("h B", content);
    }

    [Fact]
    public void VectorPolygon_RendersConfiguredStrokeJoinFromSharedShapeDescriptor() {
        var shape = OfficeShape.Polygon(
            new OfficePoint(0, 40),
            new OfficePoint(40, 0),
            new OfficePoint(80, 40));
        shape.StrokeColor = OfficeColor.SteelBlue;
        shape.StrokeWidth = 2;
        shape.StrokeLineJoin = OfficeStrokeLineJoin.Round;

        byte[] bytes = PdfDoc.Create(new PdfOptions {
                PageWidth = 240,
                PageHeight = 180,
                MarginLeft = 20,
                MarginRight = 20,
                MarginTop = 20,
                MarginBottom = 20
            })
            .Shape(shape, align: PdfAlign.Center, spacingBefore: 4, spacingAfter: 6)
            .ToBytes();

        string content = Encoding.ASCII.GetString(bytes);

        Assert.Contains("0.275 0.51 0.706 RG", content);
        Assert.Contains("2 w", content);
        Assert.Contains("1 j", content);
        Assert.Contains("80 120 m", content);
        Assert.Contains("h S", content);
    }

    [Fact]
    public void VectorPath_RendersMoveCurveAndCloseOperatorsFromSharedShapeDescriptor() {
        var shape = OfficeShape.Path(
            OfficePathCommand.MoveTo(0, 40),
            OfficePathCommand.CubicBezierTo(20, 0, 60, 0, 80, 40),
            OfficePathCommand.Close());
        shape.FillColor = OfficeColor.WhiteSmoke;
        shape.StrokeColor = OfficeColor.SteelBlue;
        shape.StrokeWidth = 1.5;

        byte[] bytes = PdfDoc.Create(new PdfOptions {
                PageWidth = 240,
                PageHeight = 180,
                MarginLeft = 20,
                MarginRight = 20,
                MarginTop = 20,
                MarginBottom = 20
            })
            .Shape(shape, align: PdfAlign.Center, spacingBefore: 4, spacingAfter: 6)
            .ToBytes();

        string content = Encoding.ASCII.GetString(bytes);

        Assert.Contains("0.961 0.961 0.961 rg", content);
        Assert.Contains("0.275 0.51 0.706 RG", content);
        Assert.Contains("1.5 w", content);
        Assert.Contains("80 120 m", content);
        Assert.Contains("100 160 140 160 160 120 c", content);
        Assert.Contains("h", content);
        Assert.Contains("B", content);
    }

    [Fact]
    public void VectorDrawing_RendersPositionedShapesFromSharedDrawingScene() {
        var background = OfficeShape.Rectangle(120, 60);
        background.FillColor = OfficeColor.WhiteSmoke;

        var marker = OfficeShape.Polygon(
            new OfficePoint(0, 30),
            new OfficePoint(40, 0),
            new OfficePoint(80, 30));
        marker.FillColor = OfficeColor.SteelBlue;
        marker.StrokeColor = OfficeColor.Black;
        marker.StrokeWidth = 1.25;

        var drawing = new OfficeDrawing(120, 60)
            .AddShape(background, 0, 0)
            .AddShape(marker, 20, 15);

        byte[] bytes = PdfDoc.Create(new PdfOptions {
                PageWidth = 240,
                PageHeight = 180,
                MarginLeft = 20,
                MarginRight = 20,
                MarginTop = 20,
                MarginBottom = 20
            })
            .Drawing(drawing, align: PdfAlign.Center, spacingBefore: 4, spacingAfter: 6)
            .ToBytes();

        string content = Encoding.ASCII.GetString(bytes);

        Assert.Contains("0.961 0.961 0.961 rg", content);
        Assert.Contains("60 100 120 60 re f", content);
        Assert.Contains("0.275 0.51 0.706 rg", content);
        Assert.Contains("0 0 0 RG", content);
        Assert.Contains("1.25 w", content);
        Assert.Contains("80 115 m", content);
        Assert.Contains("120 145 l", content);
        Assert.Contains("160 115 l", content);
        Assert.Contains("h B", content);
    }

    [Fact]
    public void VectorShape_RejectsInvalidGeometry() {
        Assert.Throws<ArgumentOutOfRangeException>(() =>
            PdfDoc.Create()
                .Rectangle(width: -1, height: 24));

        var invalidStroke = OfficeShape.Rectangle(90, 24);
        invalidStroke.StrokeWidth = -0.5;

        Assert.Throws<ArgumentOutOfRangeException>(() =>
            PdfDoc.Create()
                .Shape(invalidStroke));

        var invalidOpacity = OfficeShape.Rectangle(90, 24);
        invalidOpacity.FillOpacity = 1.1;

        Assert.Throws<ArgumentOutOfRangeException>(() =>
            PdfDoc.Create()
                .Shape(invalidOpacity));

        var invalidClipPath = OfficeShape.Rectangle(90, 24);
        invalidClipPath.ClipPath = OfficeClipPath.Rectangle(91, 24);

        Assert.Throws<ArgumentOutOfRangeException>(() =>
            PdfDoc.Create()
                .Shape(invalidClipPath));

        var invalidPolygon = new OfficeShape {
            Kind = OfficeShapeKind.Polygon,
            Width = 20,
            Height = 20
        };

        Assert.Throws<ArgumentException>(() =>
            PdfDoc.Create()
                .Shape(invalidPolygon));

        var invalidPath = new OfficeShape {
            Kind = OfficeShapeKind.Path,
            Width = 20,
            Height = 20
        };

        Assert.Throws<ArgumentException>(() =>
            PdfDoc.Create()
                .Shape(invalidPath));

        var invalidLine = new OfficeShape {
            Kind = OfficeShapeKind.Line,
            Width = 20,
            Height = 20
        };

        Assert.Throws<ArgumentException>(() =>
            PdfDoc.Create()
                .Shape(invalidLine));

        var invalidRoundedRectangle = OfficeShape.RoundedRectangle(40, 20, 4);
        invalidRoundedRectangle.CornerRadius = 11;

        Assert.Throws<ArgumentOutOfRangeException>(() =>
            PdfDoc.Create()
                .Shape(invalidRoundedRectangle));

        var emptyDrawing = new OfficeDrawing(40, 20);

        Assert.Throws<ArgumentException>(() =>
            PdfDoc.Create()
                .Drawing(emptyDrawing));

        var shapeSpacingBeforeException = Assert.Throws<ArgumentException>(() =>
            PdfDoc.Create()
                .Rectangle(width: 40, height: 20, spacingBefore: -1));

        Assert.Contains("Shape spacing before must be a non-negative finite value.", shapeSpacingBeforeException.Message, StringComparison.Ordinal);

        var shapeSpacingAfterException = Assert.Throws<ArgumentException>(() =>
            PdfDoc.Create()
                .Rectangle(width: 40, height: 20, spacingAfter: double.PositiveInfinity));

        Assert.Contains("Shape spacing after must be a non-negative finite value.", shapeSpacingAfterException.Message, StringComparison.Ordinal);

        var drawing = new OfficeDrawing(40, 20)
            .AddShape(OfficeShape.Rectangle(40, 20), 0, 0);

        var drawingSpacingBeforeException = Assert.Throws<ArgumentException>(() =>
            PdfDoc.Create()
                .Drawing(drawing, spacingBefore: double.NaN));

        Assert.Contains("Drawing spacing before must be a non-negative finite value.", drawingSpacingBeforeException.Message, StringComparison.Ordinal);

        var drawingSpacingAfterException = Assert.Throws<ArgumentException>(() =>
            PdfDoc.Create()
                .Compose(compose =>
                    compose.Page(page =>
                        page.Content(content =>
                            content.Row(row =>
                                row.Column(100, column =>
                                    column.Drawing(drawing, spacingAfter: -1)))))));

        Assert.Contains("Drawing spacing after must be a non-negative finite value.", drawingSpacingAfterException.Message, StringComparison.Ordinal);
    }

    [Fact]
    public void VectorShapeAndDrawing_RejectFlowBlocksTallerThanContentArea() {
        var options = new PdfOptions {
            PageWidth = 220,
            PageHeight = 140,
            MarginLeft = 20,
            MarginRight = 20,
            MarginTop = 20,
            MarginBottom = 20
        };

        var shapeException = Assert.Throws<ArgumentException>(() =>
            PdfDoc.Create(options)
                .Rectangle(width: 80, height: 130)
                .ToBytes());
        Assert.Contains("Shape height exceeds the available page content height.", shapeException.Message, StringComparison.Ordinal);

        var drawing = new OfficeDrawing(80, 130)
            .AddShape(OfficeShape.Rectangle(80, 130), 0, 0);

        var drawingException = Assert.Throws<ArgumentException>(() =>
            PdfDoc.Create(options)
                .Drawing(drawing)
                .ToBytes());
        Assert.Contains("Drawing height exceeds the available page content height.", drawingException.Message, StringComparison.Ordinal);
    }

    [Fact]
    public void VectorShapeAndDrawing_RejectFlowBlocksWiderThanContentArea() {
        var options = new PdfOptions {
            PageWidth = 220,
            PageHeight = 180,
            MarginLeft = 20,
            MarginRight = 20,
            MarginTop = 20,
            MarginBottom = 20
        };

        var shapeException = Assert.Throws<ArgumentException>(() =>
            PdfDoc.Create(options)
                .Rectangle(width: 190, height: 40)
                .ToBytes());
        Assert.Contains("Shape width exceeds the available page content width.", shapeException.Message, StringComparison.Ordinal);

        var drawing = new OfficeDrawing(190, 40)
            .AddShape(OfficeShape.Rectangle(190, 40), 0, 0);

        var drawingException = Assert.Throws<ArgumentException>(() =>
            PdfDoc.Create(options)
                .Drawing(drawing)
                .ToBytes());
        Assert.Contains("Drawing width exceeds the available page content width.", drawingException.Message, StringComparison.Ordinal);
    }

    [Fact]
    public void Table_WrapsLongCellTextInsideContentArea() {
        var options = new PdfOptions {
            DefaultFont = PdfStandardFont.Helvetica,
            DefaultFontSize = 10
        };

        byte[] bytes = PdfDoc.Create(options)
            .Table(new[] {
                new[] { "Area", "Status" },
                new[] {
                    "Generation",
                    "This is a long table cell value that should wrap instead of drawing across the next column or past the page margin."
                }
            })
            .ToBytes();

        using var pdf = PdfDocument.Open(new MemoryStream(bytes));
        var page = pdf.GetPage(1);
        double contentRight = options.PageWidth - options.MarginRight;
        double rightMost = page.Letters
            .Where(letter => !string.IsNullOrWhiteSpace(letter.Value))
            .Max(letter => letter.EndBaseLine.X);

        Assert.InRange(rightMost, double.NegativeInfinity, contentRight + 1);

        int statusLineCount = page.Letters
            .Where(letter => !string.IsNullOrWhiteSpace(letter.Value) && letter.StartBaseLine.X > options.MarginLeft + 250)
            .GroupBy(letter => Math.Round(letter.StartBaseLine.Y, 1))
            .Count();

        Assert.True(statusLineCount > 1, "Expected the long table cell to wrap to multiple visual lines.");
    }

    [Fact]
    public void Table_UsesProportionalGlyphWidthsForWideCellWrapping() {
        var options = new PdfOptions {
            PageWidth = 200,
            PageHeight = 160,
            MarginLeft = 30,
            MarginRight = 30,
            MarginTop = 30,
            MarginBottom = 30,
            DefaultFont = PdfStandardFont.Helvetica,
            DefaultFontSize = 10
        };
        var style = TableStyles.Minimal();
        style.HeaderRowCount = 0;
        style.ColumnWidthPoints = new List<double?> { 80 };

        byte[] bytes = PdfDoc.Create(options)
            .Table(new[] {
                new[] { "WWWWWWWW" }
            }, style: style)
            .ToBytes();

        using var pdf = PdfDocument.Open(new MemoryStream(bytes));
        var page = pdf.GetPage(1);
        int wideLineCount = page.Letters
            .Where(letter => letter.Value == "W")
            .GroupBy(letter => Math.Round(letter.StartBaseLine.Y, 1))
            .Count();
        double cellRight = options.MarginLeft + 80;
        double rightMostWideGlyph = page.Letters
            .Where(letter => letter.Value == "W")
            .Max(letter => letter.EndBaseLine.X);

        Assert.True(wideLineCount > 1, "Expected wide glyphs to wrap using their real Helvetica advance instead of an average character width.");
        Assert.InRange(rightMostWideGlyph, double.NegativeInfinity, cellRight + 1);
    }

    [Fact]
    public void Table_UsesProportionalGlyphWidthsWithoutOverWrappingNarrowCells() {
        var options = new PdfOptions {
            PageWidth = 200,
            PageHeight = 160,
            MarginLeft = 30,
            MarginRight = 30,
            MarginTop = 30,
            MarginBottom = 30,
            DefaultFont = PdfStandardFont.Helvetica,
            DefaultFontSize = 10
        };
        var style = TableStyles.Minimal();
        style.HeaderRowCount = 0;
        style.ColumnWidthPoints = new List<double?> { 80 };

        byte[] bytes = PdfDoc.Create(options)
            .Table(new[] {
                new[] { new string('i', 20) }
            }, style: style)
            .ToBytes();

        using var pdf = PdfDocument.Open(new MemoryStream(bytes));
        var page = pdf.GetPage(1);
        int narrowLineCount = page.Letters
            .Where(letter => letter.Value == "i")
            .GroupBy(letter => Math.Round(letter.StartBaseLine.Y, 1))
            .Count();

        Assert.Equal(1, narrowLineCount);
    }

    [Fact]
    public void RowColumnTable_UsesProportionalGlyphWidthsForWideCellWrapping() {
        var options = new PdfOptions {
            PageWidth = 200,
            PageHeight = 160,
            MarginLeft = 30,
            MarginRight = 30,
            MarginTop = 30,
            MarginBottom = 30,
            DefaultFont = PdfStandardFont.Helvetica,
            DefaultFontSize = 10
        };
        var style = TableStyles.Minimal();
        style.HeaderRowCount = 0;
        style.ColumnWidthPoints = new List<double?> { 80 };

        byte[] bytes = PdfDoc.Create(options)
            .Compose(document =>
                document.Page(page =>
                    page.Content(content =>
                        content.Row(row =>
                            row.Column(100, column =>
                                column.Table(new[] {
                                    new[] { "WWWWWWWW" }
                                }, style: style))))))
            .ToBytes();

        using var pdf = PdfDocument.Open(new MemoryStream(bytes));
        var page = pdf.GetPage(1);
        int wideLineCount = page.Letters
            .Where(letter => letter.Value == "W")
            .GroupBy(letter => Math.Round(letter.StartBaseLine.Y, 1))
            .Count();
        double cellRight = options.MarginLeft + 80;
        double rightMostWideGlyph = page.Letters
            .Where(letter => letter.Value == "W")
            .Max(letter => letter.EndBaseLine.X);

        Assert.True(wideLineCount > 1, "Expected row-column table wide glyphs to wrap using their real Helvetica advance instead of an average character width.");
        Assert.InRange(rightMostWideGlyph, double.NegativeInfinity, cellRight + 1);
    }

    [Fact]
    public void RowColumnTable_UsesProportionalGlyphWidthsWithoutOverWrappingNarrowCells() {
        var options = new PdfOptions {
            PageWidth = 200,
            PageHeight = 160,
            MarginLeft = 30,
            MarginRight = 30,
            MarginTop = 30,
            MarginBottom = 30,
            DefaultFont = PdfStandardFont.Helvetica,
            DefaultFontSize = 10
        };
        var style = TableStyles.Minimal();
        style.HeaderRowCount = 0;
        style.ColumnWidthPoints = new List<double?> { 80 };

        byte[] bytes = PdfDoc.Create(options)
            .Compose(document =>
                document.Page(page =>
                    page.Content(content =>
                        content.Row(row =>
                            row.Column(100, column =>
                                column.Table(new[] {
                                    new[] { new string('i', 20) }
                                }, style: style))))))
            .ToBytes();

        using var pdf = PdfDocument.Open(new MemoryStream(bytes));
        var page = pdf.GetPage(1);
        int narrowLineCount = page.Letters
            .Where(letter => letter.Value == "i")
            .GroupBy(letter => Math.Round(letter.StartBaseLine.Y, 1))
            .Count();

        Assert.Equal(1, narrowLineCount);
    }

    [Fact]
    public void List_UsesProportionalGlyphWidthsForWideBulletWrapping() {
        var options = new PdfOptions {
            PageWidth = 120,
            PageHeight = 160,
            MarginLeft = 30,
            MarginRight = 30,
            MarginTop = 30,
            MarginBottom = 30,
            DefaultFont = PdfStandardFont.Helvetica,
            DefaultFontSize = 10
        };

        byte[] bytes = PdfDoc.Create(options)
            .Bullets(new[] { "WWWWWWWW" })
            .ToBytes();

        using var pdf = PdfDocument.Open(new MemoryStream(bytes));
        var page = pdf.GetPage(1);
        int wideLineCount = page.Letters
            .Where(letter => letter.Value == "W")
            .GroupBy(letter => Math.Round(letter.StartBaseLine.Y, 1))
            .Count();
        double contentRight = options.PageWidth - options.MarginRight;
        double rightMostWideGlyph = page.Letters
            .Where(letter => letter.Value == "W")
            .Max(letter => letter.EndBaseLine.X);

        Assert.True(wideLineCount > 1, "Expected bullet-list wide glyphs to wrap using their real Helvetica advance instead of an average character width.");
        Assert.InRange(rightMostWideGlyph, double.NegativeInfinity, contentRight + 1);
    }

    [Fact]
    public void List_UsesProportionalGlyphWidthsWithoutOverWrappingNarrowBullets() {
        var options = new PdfOptions {
            PageWidth = 120,
            PageHeight = 160,
            MarginLeft = 30,
            MarginRight = 30,
            MarginTop = 30,
            MarginBottom = 30,
            DefaultFont = PdfStandardFont.Helvetica,
            DefaultFontSize = 10
        };

        byte[] bytes = PdfDoc.Create(options)
            .Bullets(new[] { new string('i', 20) })
            .ToBytes();

        using var pdf = PdfDocument.Open(new MemoryStream(bytes));
        var page = pdf.GetPage(1);
        int narrowLineCount = page.Letters
            .Where(letter => letter.Value == "i")
            .GroupBy(letter => Math.Round(letter.StartBaseLine.Y, 1))
            .Count();

        Assert.Equal(1, narrowLineCount);
    }

    [Fact]
    public void RowColumnList_UsesProportionalGlyphWidthsForWideNumberedWrapping() {
        var options = new PdfOptions {
            PageWidth = 120,
            PageHeight = 160,
            MarginLeft = 30,
            MarginRight = 30,
            MarginTop = 30,
            MarginBottom = 30,
            DefaultFont = PdfStandardFont.Helvetica,
            DefaultFontSize = 10
        };

        byte[] bytes = PdfDoc.Create(options)
            .Compose(document =>
                document.Page(page =>
                    page.Content(content =>
                        content.Row(row =>
                            row.Column(100, column =>
                                column.Numbered(new[] { "WWWWWWWW" }))))))
            .ToBytes();

        using var pdf = PdfDocument.Open(new MemoryStream(bytes));
        var page = pdf.GetPage(1);
        int wideLineCount = page.Letters
            .Where(letter => letter.Value == "W")
            .GroupBy(letter => Math.Round(letter.StartBaseLine.Y, 1))
            .Count();
        double contentRight = options.PageWidth - options.MarginRight;
        double rightMostWideGlyph = page.Letters
            .Where(letter => letter.Value == "W")
            .Max(letter => letter.EndBaseLine.X);

        Assert.True(wideLineCount > 1, "Expected row-column numbered-list wide glyphs to wrap using their real Helvetica advance instead of an average character width.");
        Assert.InRange(rightMostWideGlyph, double.NegativeInfinity, contentRight + 1);
    }

    [Fact]
    public void RowColumnList_UsesProportionalGlyphWidthsWithoutOverWrappingNarrowNumberedItems() {
        var options = new PdfOptions {
            PageWidth = 120,
            PageHeight = 160,
            MarginLeft = 30,
            MarginRight = 30,
            MarginTop = 30,
            MarginBottom = 30,
            DefaultFont = PdfStandardFont.Helvetica,
            DefaultFontSize = 10
        };

        byte[] bytes = PdfDoc.Create(options)
            .Compose(document =>
                document.Page(page =>
                    page.Content(content =>
                        content.Row(row =>
                            row.Column(100, column =>
                                column.Numbered(new[] { new string('i', 20) }))))))
            .ToBytes();

        using var pdf = PdfDocument.Open(new MemoryStream(bytes));
        var page = pdf.GetPage(1);
        int narrowLineCount = page.Letters
            .Where(letter => letter.Value == "i")
            .GroupBy(letter => Math.Round(letter.StartBaseLine.Y, 1))
            .Count();

        Assert.Equal(1, narrowLineCount);
    }

    [Fact]
    public void Table_BreaksLongUnspacedTokensAfterShortPrefixInsideContentArea() {
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
        var style = TableStyles.Minimal();
        style.ColumnWidthPoints = new List<double?> { 70, null };

        byte[] bytes = PdfDoc.Create(options)
            .Table(new[] {
                new[] { "Field", "Value" },
                new[] {
                    "Token",
                    "id " + new string('X', 72)
                }
            }, style: style)
            .ToBytes();

        using var pdf = PdfDocument.Open(new MemoryStream(bytes));
        var page = pdf.GetPage(1);
        double contentRight = options.PageWidth - options.MarginRight;
        double rightMost = page.Letters
            .Where(letter => !string.IsNullOrWhiteSpace(letter.Value))
            .Max(letter => letter.EndBaseLine.X);

        Assert.InRange(rightMost, double.NegativeInfinity, contentRight + 1);

        int tokenLineCount = page.Letters
            .Where(letter => letter.Value == "X")
            .GroupBy(letter => Math.Round(letter.StartBaseLine.Y, 1))
            .Count();

        Assert.True(tokenLineCount > 2, "Expected the long unspaced token to split across multiple table cell lines.");
    }

    [Fact]
    public void Table_CellTextThatEscapesCellRectanglesIsClipped() {
        var style = TableStyles.Minimal();
        style.CellPaddingX = 8;
        style.CellPaddingY = 5;
        style.RowBaselineOffset = 40;

        byte[] topLevelBytes = PdfDoc.Create(new PdfOptions {
                PageWidth = 320,
                PageHeight = 220,
                MarginLeft = 30,
                MarginRight = 30,
                MarginTop = 30,
                MarginBottom = 30,
                DefaultFont = PdfStandardFont.Helvetica,
                DefaultFontSize = 10
            })
            .Table(new[] {
                new[] { "Name", "Value" },
                new[] { "Long", "This table cell deliberately wraps so the writer has to emit more than one clipped text line." }
            }, style: style)
            .ToBytes();

        string topLevelContent = string.Join("\n", GetPageContentStreams(topLevelBytes, 1));
        int topLevelClipCount = Regex.Matches(topLevelContent, " re W n\\nBT\\n/F").Count;
        Assert.True(topLevelClipCount >= 5, "Expected top-level table cell text to be clipped by PDF cell rectangles.");

        byte[] rowColumnBytes = PdfDoc.Create(new PdfOptions {
                PageWidth = 320,
                PageHeight = 220,
                MarginLeft = 30,
                MarginRight = 30,
                MarginTop = 30,
                MarginBottom = 30,
                DefaultFont = PdfStandardFont.Helvetica,
                DefaultFontSize = 10
            })
            .Compose(compose =>
                compose.Page(page =>
                    page.Content(content =>
                        content.Row(row =>
                            row.Column(100, column =>
                                column.Table(new[] {
                                    new[] { "Name", "Value" },
                                    new[] { "Long", "Column-local table cells also get clipped to the cell content rectangle." }
                                }, style: style))))))
            .ToBytes();

        string rowColumnContent = string.Join("\n", GetPageContentStreams(rowColumnBytes, 1));
        int rowColumnClipCount = Regex.Matches(rowColumnContent, " re W n\\nBT\\n/F").Count;
        Assert.True(rowColumnClipCount >= 5, "Expected row-column table cell text to be clipped by PDF cell rectangles.");
    }

    [Fact]
    public void Table_PaginatesLongTablesAndRepeatsHeaderRows() {
        var options = new PdfOptions {
            PageWidth = 360,
            PageHeight = 240,
            MarginLeft = 30,
            MarginRight = 30,
            MarginTop = 30,
            MarginBottom = 30,
            DefaultFont = PdfStandardFont.Helvetica,
            DefaultFontSize = 9
        };

        var rows = new List<string[]> {
            new[] { "Metric", "Status" }
        };
        for (int i = 1; i <= 28; i++) {
            rows.Add(new[] { "Item " + i.ToString(), "Completed without clipping" });
        }

        byte[] bytes = PdfDoc.Create(options)
            .Table(rows)
            .ToBytes();

        using var pdf = PdfDocument.Open(new MemoryStream(bytes));
        Assert.True(pdf.NumberOfPages > 1, "Expected a long table to continue onto another page.");

        for (int pageNumber = 1; pageNumber <= pdf.NumberOfPages; pageNumber++) {
            var page = pdf.GetPage(pageNumber);
            Assert.Contains("Metric", page.Text);
            Assert.Contains("Status", page.Text);

            double bottomMost = page.Letters
                .Where(letter => !string.IsNullOrWhiteSpace(letter.Value))
                .Min(letter => letter.StartBaseLine.Y);
            Assert.True(bottomMost >= options.MarginBottom - 2, $"Expected table text to stay above the bottom margin on page {pageNumber}.");
        }

        Assert.Contains("Item 1", pdf.GetPage(1).Text);
        Assert.Contains("Item 28", pdf.GetPage(pdf.NumberOfPages).Text);
    }

    [Fact]
    public void Table_RepeatsConfiguredHeaderRowsAcrossPages() {
        var options = new PdfOptions {
            PageWidth = 360,
            PageHeight = 220,
            MarginLeft = 30,
            MarginRight = 30,
            MarginTop = 30,
            MarginBottom = 30,
            DefaultFont = PdfStandardFont.Helvetica,
            DefaultFontSize = 9
        };
        var style = TableStyles.Minimal();
        style.HeaderRowCount = 2;
        style.ColumnWidthWeights = new List<double> { 1, 1 };

        var rows = new List<string[]> {
            new[] { "Group", "State" },
            new[] { "Metric", "Owner" }
        };
        for (int i = 1; i <= 30; i++) {
            rows.Add(new[] { "Check " + i.ToString(), "Team " + i.ToString() });
        }

        byte[] bytes = PdfDoc.Create(options)
            .Table(rows, style: style)
            .ToBytes();

        using var pdf = PdfDocument.Open(new MemoryStream(bytes));
        Assert.True(pdf.NumberOfPages > 1, "Expected the two-row header table to continue onto another page.");

        for (int pageNumber = 1; pageNumber <= pdf.NumberOfPages; pageNumber++) {
            var pageText = pdf.GetPage(pageNumber).Text;
            Assert.Contains("Group", pageText);
            Assert.Contains("State", pageText);
            Assert.Contains("Metric", pageText);
            Assert.Contains("Owner", pageText);
        }

        Assert.Contains("Check 1", pdf.GetPage(1).Text);
        Assert.Contains("Check 30", pdf.GetPage(pdf.NumberOfPages).Text);
    }

    [Fact]
    public void RowColumnTable_RepeatsHeaderRowsAcrossPages() {
        var options = new PdfOptions {
            PageWidth = 360,
            PageHeight = 210,
            MarginLeft = 30,
            MarginRight = 30,
            MarginTop = 30,
            MarginBottom = 30,
            DefaultFont = PdfStandardFont.Helvetica,
            DefaultFontSize = 9
        };
        var style = TableStyles.Minimal();
        style.HeaderRowCount = 1;
        style.ColumnWidthWeights = new List<double> { 1, 1 };

        var rows = new List<string[]> {
            new[] { "ColMetric", "ColValue" }
        };
        for (int i = 1; i <= 28; i++) {
            rows.Add(new[] { "ColumnCheck " + i.ToString(CultureInfo.InvariantCulture), "Ready" });
        }

        byte[] bytes = PdfDoc.Create(options)
            .Compose(compose =>
                compose.Page(page =>
                    page.Content(content =>
                        content.Row(row =>
                            row.Column(100, column => column.Table(rows, style: style))))))
            .ToBytes();

        using var pdf = PdfDocument.Open(new MemoryStream(bytes));
        Assert.True(pdf.NumberOfPages > 1, "Expected the column-local table to continue onto another page.");

        for (int pageNumber = 1; pageNumber <= pdf.NumberOfPages; pageNumber++) {
            var pageText = pdf.GetPage(pageNumber).Text;
            Assert.Contains("ColMetric", pageText);
            Assert.Contains("ColValue", pageText);
        }

        Assert.Contains("ColumnCheck 1", pdf.GetPage(1).Text);
        Assert.Contains("ColumnCheck 28", pdf.GetPage(pdf.NumberOfPages).Text);
    }

    [Fact]
    public void Table_RendersConfiguredFooterRowsAtEndOfLongTables() {
        var options = new PdfOptions {
            PageWidth = 360,
            PageHeight = 220,
            MarginLeft = 30,
            MarginRight = 30,
            MarginTop = 30,
            MarginBottom = 30,
            DefaultFont = PdfStandardFont.Helvetica,
            DefaultFontSize = 9
        };
        var style = TableStyles.Minimal();
        style.FooterRowCount = 1;
        style.FooterFill = PdfColor.FromRgb(230, 230, 230);
        style.FooterTextColor = PdfColor.FromRgb(20, 20, 20);

        var rows = new List<string[]> {
            new[] { "Metric", "Value" }
        };
        for (int i = 1; i <= 30; i++) {
            rows.Add(new[] { "Item " + i.ToString(), i.ToString() });
        }
        rows.Add(new[] { "Total", "30" });

        byte[] bytes = PdfDoc.Create(options)
            .Table(rows, style: style)
            .ToBytes();

        using var pdf = PdfDocument.Open(new MemoryStream(bytes));
        Assert.True(pdf.NumberOfPages > 1, "Expected a long table with footer rows to continue onto another page.");
        Assert.DoesNotContain("Total", pdf.GetPage(1).Text);
        Assert.Contains("Total", pdf.GetPage(pdf.NumberOfPages).Text);
        Assert.Contains("30", pdf.GetPage(pdf.NumberOfPages).Text);
    }

    [Fact]
    public void Table_KeepTogetherMovesWholeTableToNextPage() {
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
        var style = TableStyles.Minimal();
        style.HeaderRowCount = 0;
        style.KeepTogether = true;

        byte[] bytes = PdfDoc.Create(options)
            .Paragraph(p => p.Text("IntroMarker"), style: new PdfParagraphStyle {
                SpacingAfter = 52
            })
            .Table(new[] {
                new[] { "KeepA", "Ready" },
                new[] { "KeepB", "Ready" },
                new[] { "KeepC", "Ready" }
            }, style: style)
            .ToBytes();

        using var pdf = PdfDocument.Open(new MemoryStream(bytes));
        Assert.Equal(2, pdf.NumberOfPages);
        Assert.Contains("IntroMarker", pdf.GetPage(1).Text);
        Assert.DoesNotContain("KeepA", pdf.GetPage(1).Text);
        Assert.Contains("KeepA", pdf.GetPage(2).Text);
        Assert.Contains("KeepC", pdf.GetPage(2).Text);
    }

    [Fact]
    public void RowColumnTable_KeepTogetherMovesWholeTableToNextPage() {
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
        var style = TableStyles.Minimal();
        style.HeaderRowCount = 0;
        style.KeepTogether = true;

        byte[] bytes = PdfDoc.Create(options)
            .Compose(compose =>
                compose.Page(page =>
                    page.Content(content =>
                        content.Row(row =>
                            row.Column(100, column => {
                                column.Paragraph(p => p.Text("ColumnIntroMarker"), style: new PdfParagraphStyle {
                                    SpacingAfter = 52
                                });
                                column.Table(new[] {
                                    new[] { "ColumnKeepA", "Ready" },
                                    new[] { "ColumnKeepB", "Ready" },
                                    new[] { "ColumnKeepC", "Ready" }
                                }, style: style);
                            })))))
            .ToBytes();

        using var pdf = PdfDocument.Open(new MemoryStream(bytes));
        Assert.Equal(2, pdf.NumberOfPages);
        Assert.Contains("ColumnIntroMarker", pdf.GetPage(1).Text);
        Assert.DoesNotContain("ColumnKeepA", pdf.GetPage(1).Text);
        Assert.Contains("ColumnKeepA", pdf.GetPage(2).Text);
        Assert.Contains("ColumnKeepC", pdf.GetPage(2).Text);
    }

    [Fact]
    public void Table_KeepTogetherRejectsTableTallerThanContentArea() {
        var style = TableStyles.Minimal();
        style.KeepTogether = true;
        style.HeaderRowCount = 0;
        style.LineHeight = 2.0;

        var rows = Enumerable.Range(1, 10)
            .Select(i => new[] { "KeepTooTall" + i.ToString(CultureInfo.InvariantCulture) })
            .ToArray();

        var exception = Assert.Throws<ArgumentException>(() =>
            PdfDoc.Create(new PdfOptions {
                    PageWidth = 260,
                    PageHeight = 160,
                    MarginLeft = 30,
                    MarginRight = 30,
                    MarginTop = 30,
                    MarginBottom = 30,
                    DefaultFont = PdfStandardFont.Helvetica,
                    DefaultFontSize = 10
                })
                .Table(rows, style: style)
                .ToBytes());

        Assert.Contains("Table height exceeds the available page content height.", exception.Message, StringComparison.Ordinal);
    }

    [Fact]
    public void Table_RendersConfiguredBodyColumnFills() {
        var style = TableStyles.Minimal();
        style.HeaderFill = null;
        style.RowStripeFill = null;
        style.BodyColumnFills = new List<PdfColor?> {
            null,
            new PdfColor(0.11, 0.22, 0.33)
        };

        byte[] bytes = PdfDoc.Create(new PdfOptions {
                PageWidth = 320,
                PageHeight = 220,
                MarginLeft = 30,
                MarginRight = 30,
                MarginTop = 30,
                MarginBottom = 30
            })
            .Table(new[] {
                new[] { "Metric", "Status" },
                new[] { "Queue", "Healthy" },
                new[] { "Latency", "Warning" }
            }, style: style)
            .ToBytes();

        string content = Encoding.ASCII.GetString(bytes);
        int fillCount = content.Split(new[] { "0.11 0.22 0.33 rg" }, StringSplitOptions.None).Length - 1;

        Assert.Equal(2, fillCount);
        Assert.Contains(" re f", content);
    }

    [Fact]
    public void RowColumnTable_RendersConfiguredBodyColumnFills() {
        var style = TableStyles.Minimal();
        style.HeaderFill = null;
        style.RowStripeFill = null;
        style.BodyColumnFills = new List<PdfColor?> {
            null,
            new PdfColor(0.11, 0.22, 0.33)
        };

        byte[] bytes = PdfDoc.Create(new PdfOptions {
                PageWidth = 320,
                PageHeight = 220,
                MarginLeft = 30,
                MarginRight = 30,
                MarginTop = 30,
                MarginBottom = 30
            })
            .Compose(document =>
                document.Page(page =>
                    page.Content(content =>
                        content.Row(row =>
                            row.Column(100, column =>
                                column.Table(new[] {
                                    new[] { "Metric", "Status" },
                                    new[] { "Queue", "Healthy" },
                                    new[] { "Latency", "Warning" }
                                }, style: style))))))
            .ToBytes();

        string contentStream = Encoding.ASCII.GetString(bytes);
        int fillCount = contentStream.Split(new[] { "0.11 0.22 0.33 rg" }, StringSplitOptions.None).Length - 1;

        Assert.Equal(2, fillCount);
        Assert.Contains(" re f", contentStream);
    }

    [Fact]
    public void Table_DoesNotApplyBodyRowStripeFillToHeaderRows() {
        var style = TableStyles.Minimal();
        style.HeaderRowCount = 2;
        style.HeaderFill = null;
        style.RowStripeFill = new PdfColor(0.19, 0.29, 0.39);

        byte[] bytes = PdfDoc.Create(new PdfOptions {
                PageWidth = 320,
                PageHeight = 220,
                MarginLeft = 30,
                MarginRight = 30,
                MarginTop = 30,
                MarginBottom = 30
            })
            .Table(new[] {
                new[] { "Group", "State" },
                new[] { "Metric", "Owner" },
                new[] { "Queue", "Healthy" },
                new[] { "Latency", "Warning" },
                new[] { "Errors", "None" }
            }, style: style)
            .ToBytes();

        string content = Encoding.ASCII.GetString(bytes);
        int fillCount = content.Split(new[] { "0.19 0.29 0.39 rg" }, StringSplitOptions.None).Length - 1;

        Assert.Equal(1, fillCount);
    }

    [Fact]
    public void RowColumnTable_DoesNotApplyBodyRowStripeFillToHeaderRows() {
        var style = TableStyles.Minimal();
        style.HeaderRowCount = 2;
        style.HeaderFill = null;
        style.RowStripeFill = new PdfColor(0.19, 0.29, 0.39);

        byte[] bytes = PdfDoc.Create(new PdfOptions {
                PageWidth = 320,
                PageHeight = 220,
                MarginLeft = 30,
                MarginRight = 30,
                MarginTop = 30,
                MarginBottom = 30
            })
            .Compose(document =>
                document.Page(page =>
                    page.Content(content =>
                        content.Row(row =>
                            row.Column(100, column =>
                                column.Table(new[] {
                                    new[] { "Group", "State" },
                                    new[] { "Metric", "Owner" },
                                    new[] { "Queue", "Healthy" },
                                    new[] { "Latency", "Warning" },
                                    new[] { "Errors", "None" }
                                }, style: style))))))
            .ToBytes();

        string content = Encoding.ASCII.GetString(bytes);
        int fillCount = content.Split(new[] { "0.19 0.29 0.39 rg" }, StringSplitOptions.None).Length - 1;

        Assert.Equal(1, fillCount);
    }

    [Fact]
    public void Table_StripesBodyRowsRelativeToFirstBodyRow() {
        var style = TableStyles.Minimal();
        style.HeaderRowCount = 1;
        style.HeaderFill = null;
        style.RowStripeFill = new PdfColor(0.21, 0.31, 0.41);

        byte[] bytes = PdfDoc.Create(new PdfOptions {
                PageWidth = 320,
                PageHeight = 220,
                MarginLeft = 30,
                MarginRight = 30,
                MarginTop = 30,
                MarginBottom = 30
            })
            .Table(new[] {
                new[] { "Metric", "Status" },
                new[] { "Queue", "Healthy" },
                new[] { "Latency", "Warning" },
                new[] { "Errors", "None" }
            }, style: style)
            .ToBytes();

        string content = Encoding.ASCII.GetString(bytes);
        int fillCount = content.Split(new[] { "0.21 0.31 0.41 rg" }, StringSplitOptions.None).Length - 1;

        Assert.Equal(1, fillCount);
    }

    [Fact]
    public void RowColumnTable_StripesBodyRowsRelativeToFirstBodyRow() {
        var style = TableStyles.Minimal();
        style.HeaderRowCount = 1;
        style.HeaderFill = null;
        style.RowStripeFill = new PdfColor(0.21, 0.31, 0.41);

        byte[] bytes = PdfDoc.Create(new PdfOptions {
                PageWidth = 320,
                PageHeight = 220,
                MarginLeft = 30,
                MarginRight = 30,
                MarginTop = 30,
                MarginBottom = 30
            })
            .Compose(document =>
                document.Page(page =>
                    page.Content(content =>
                        content.Row(row =>
                            row.Column(100, column =>
                                column.Table(new[] {
                                    new[] { "Metric", "Status" },
                                    new[] { "Queue", "Healthy" },
                                    new[] { "Latency", "Warning" },
                                    new[] { "Errors", "None" }
                                }, style: style))))))
            .ToBytes();

        string content = Encoding.ASCII.GetString(bytes);
        int fillCount = content.Split(new[] { "0.21 0.31 0.41 rg" }, StringSplitOptions.None).Length - 1;

        Assert.Equal(1, fillCount);
    }

    [Fact]
    public void Table_RendersConfiguredCellFills() {
        var style = TableStyles.Minimal();
        style.HeaderFill = null;
        style.RowStripeFill = null;
        style.CellFills = new Dictionary<(int Row, int Column), PdfColor> {
            [(2, 1)] = new PdfColor(0.42, 0.18, 0.66)
        };

        byte[] bytes = PdfDoc.Create(new PdfOptions {
                PageWidth = 320,
                PageHeight = 220,
                MarginLeft = 30,
                MarginRight = 30,
                MarginTop = 30,
                MarginBottom = 30
            })
            .Table(new[] {
                new[] { "Metric", "Status" },
                new[] { "Queue", "Healthy" },
                new[] { "Latency", "Warning" }
            }, style: style)
            .ToBytes();

        string content = Encoding.ASCII.GetString(bytes);
        int fillCount = content.Split(new[] { "0.42 0.18 0.66 rg" }, StringSplitOptions.None).Length - 1;

        Assert.Equal(1, fillCount);
        Assert.Contains(" re f", content);
    }

    [Fact]
    public void RowColumnTable_RendersConfiguredCellFills() {
        var style = TableStyles.Minimal();
        style.HeaderFill = null;
        style.RowStripeFill = null;
        style.CellFills = new Dictionary<(int Row, int Column), PdfColor> {
            [(2, 1)] = new PdfColor(0.42, 0.18, 0.66)
        };

        byte[] bytes = PdfDoc.Create(new PdfOptions {
                PageWidth = 320,
                PageHeight = 220,
                MarginLeft = 30,
                MarginRight = 30,
                MarginTop = 30,
                MarginBottom = 30
            })
            .Compose(document =>
                document.Page(page =>
                    page.Content(content =>
                        content.Row(row =>
                            row.Column(100, column =>
                                column.Table(new[] {
                                    new[] { "Metric", "Status" },
                                    new[] { "Queue", "Healthy" },
                                    new[] { "Latency", "Warning" }
                                }, style: style))))))
            .ToBytes();

        string contentStream = Encoding.ASCII.GetString(bytes);
        int fillCount = contentStream.Split(new[] { "0.42 0.18 0.66 rg" }, StringSplitOptions.None).Length - 1;

        Assert.Equal(1, fillCount);
        Assert.Contains(" re f", contentStream);
    }

    [Fact]
    public void Table_RendersConfiguredCellBorders() {
        var style = TableStyles.Minimal();
        style.BorderColor = null;
        style.CellBorders = new Dictionary<(int Row, int Column), PdfCellBorder> {
            [(2, 1)] = new PdfCellBorder {
                Color = new PdfColor(0.12, 0.34, 0.56),
                Width = 1.7
            }
        };

        byte[] bytes = PdfDoc.Create(new PdfOptions {
                PageWidth = 320,
                PageHeight = 220,
                MarginLeft = 30,
                MarginRight = 30,
                MarginTop = 30,
                MarginBottom = 30
            })
            .Table(new[] {
                new[] { "Metric", "Status" },
                new[] { "Queue", "Healthy" },
                new[] { "Latency", "Warning" }
            }, style: style)
            .ToBytes();

        string content = Encoding.ASCII.GetString(bytes);
        int borderColorCount = content.Split(new[] { "0.12 0.34 0.56 RG" }, StringSplitOptions.None).Length - 1;

        Assert.Equal(1, borderColorCount);
        Assert.Contains("1.7 w", content);
        Assert.Contains(" re S", content);
    }

    [Fact]
    public void RowColumnTable_RendersConfiguredCellBorders() {
        var style = TableStyles.Minimal();
        style.BorderColor = null;
        style.CellBorders = new Dictionary<(int Row, int Column), PdfCellBorder> {
            [(2, 1)] = new PdfCellBorder {
                Color = new PdfColor(0.12, 0.34, 0.56),
                Width = 1.7
            }
        };

        byte[] bytes = PdfDoc.Create(new PdfOptions {
                PageWidth = 320,
                PageHeight = 220,
                MarginLeft = 30,
                MarginRight = 30,
                MarginTop = 30,
                MarginBottom = 30
            })
            .Compose(document =>
                document.Page(page =>
                    page.Content(content =>
                        content.Row(row =>
                            row.Column(100, column =>
                                column.Table(new[] {
                                    new[] { "Metric", "Status" },
                                    new[] { "Queue", "Healthy" },
                                    new[] { "Latency", "Warning" }
                                }, style: style))))))
            .ToBytes();

        string content = Encoding.ASCII.GetString(bytes);
        int borderColorCount = content.Split(new[] { "0.12 0.34 0.56 RG" }, StringSplitOptions.None).Length - 1;

        Assert.Equal(1, borderColorCount);
        Assert.Contains("1.7 w", content);
        Assert.Contains(" re S", content);
    }

    [Fact]
    public void Table_RendersConfiguredCellBorderSides() {
        var style = TableStyles.Minimal();
        style.BorderColor = null;
        style.CellBorders = new Dictionary<(int Row, int Column), PdfCellBorder> {
            [(2, 1)] = new PdfCellBorder {
                Color = new PdfColor(0.2, 0.3, 0.4),
                Width = 2.2,
                Right = false,
                Bottom = false,
                Left = false
            }
        };

        byte[] bytes = PdfDoc.Create(new PdfOptions {
                PageWidth = 320,
                PageHeight = 220,
                MarginLeft = 30,
                MarginRight = 30,
                MarginTop = 30,
                MarginBottom = 30
            })
            .Table(new[] {
                new[] { "Metric", "Status" },
                new[] { "Queue", "Healthy" },
                new[] { "Total", "Warning" }
            }, style: style)
            .ToBytes();

        string content = Encoding.ASCII.GetString(bytes);
        int borderColorCount = content.Split(new[] { "0.2 0.3 0.4 RG" }, StringSplitOptions.None).Length - 1;

        Assert.Equal(1, borderColorCount);
        Assert.Contains("2.2 w", content);
        Assert.Contains(" l S", content);
        Assert.DoesNotContain(" re S", content);
    }

    [Fact]
    public void Table_RendersConfiguredRowSeparatorsWithoutCellBorderDictionary() {
        var style = TableStyles.Minimal();
        style.BorderColor = null;
        style.RowSeparatorColor = new PdfColor(0.12, 0.34, 0.56);
        style.RowSeparatorWidth = 0.6;
        style.HeaderSeparatorColor = new PdfColor(0.7, 0.2, 0.1);
        style.HeaderSeparatorWidth = 1.1;

        byte[] bytes = PdfDoc.Create(new PdfOptions {
                PageWidth = 320,
                PageHeight = 220,
                MarginLeft = 30,
                MarginRight = 30,
                MarginTop = 30,
                MarginBottom = 30
            })
            .Table(new[] {
                new[] { "Metric", "Status" },
                new[] { "Queue", "Healthy" },
                new[] { "Latency", "Warning" }
            }, style: style)
            .ToBytes();

        string content = Encoding.ASCII.GetString(bytes);
        int bodySeparatorCount = content.Split(new[] { "0.12 0.34 0.56 RG" }, StringSplitOptions.None).Length - 1;
        int headerSeparatorCount = content.Split(new[] { "0.7 0.2 0.1 RG" }, StringSplitOptions.None).Length - 1;

        Assert.Equal(2, bodySeparatorCount);
        Assert.Equal(1, headerSeparatorCount);
        Assert.Contains("0.6 w", content);
        Assert.Contains("1.1 w", content);
        Assert.Contains(" l S", content);
        Assert.DoesNotContain(" re S", content);
    }

    [Fact]
    public void RowColumnTable_RendersConfiguredRowSeparatorsWithoutCellBorderDictionary() {
        var style = TableStyles.Minimal();
        style.BorderColor = null;
        style.RowSeparatorColor = new PdfColor(0.12, 0.34, 0.56);
        style.RowSeparatorWidth = 0.6;
        style.HeaderSeparatorColor = new PdfColor(0.7, 0.2, 0.1);
        style.HeaderSeparatorWidth = 1.1;

        byte[] bytes = PdfDoc.Create(new PdfOptions {
                PageWidth = 320,
                PageHeight = 220,
                MarginLeft = 30,
                MarginRight = 30,
                MarginTop = 30,
                MarginBottom = 30
            })
            .Compose(document =>
                document.Page(page =>
                    page.Content(content =>
                        content.Row(row =>
                            row.Column(100, column =>
                                column.Table(new[] {
                                    new[] { "Metric", "Status" },
                                    new[] { "Queue", "Healthy" },
                                    new[] { "Latency", "Warning" }
                                }, style: style))))))
            .ToBytes();

        string content = Encoding.ASCII.GetString(bytes);
        int bodySeparatorCount = content.Split(new[] { "0.12 0.34 0.56 RG" }, StringSplitOptions.None).Length - 1;
        int headerSeparatorCount = content.Split(new[] { "0.7 0.2 0.1 RG" }, StringSplitOptions.None).Length - 1;

        Assert.Equal(2, bodySeparatorCount);
        Assert.Equal(1, headerSeparatorCount);
        Assert.Contains("0.6 w", content);
        Assert.Contains("1.1 w", content);
        Assert.Contains(" l S", content);
        Assert.DoesNotContain(" re S", content);
    }

    [Fact]
    public void Table_RendersConfiguredFooterSeparatorAboveFooterRows() {
        var style = TableStyles.Minimal();
        style.BorderColor = null;
        style.RowSeparatorColor = new PdfColor(0.12, 0.34, 0.56);
        style.RowSeparatorWidth = 0.6;
        style.FooterRowCount = 1;
        style.FooterSeparatorColor = new PdfColor(0.2, 0.7, 0.3);
        style.FooterSeparatorWidth = 1.3;

        byte[] bytes = PdfDoc.Create(new PdfOptions {
                PageWidth = 320,
                PageHeight = 220,
                MarginLeft = 30,
                MarginRight = 30,
                MarginTop = 30,
                MarginBottom = 30
            })
            .Table(new[] {
                new[] { "Metric", "Status" },
                new[] { "Queue", "Healthy" },
                new[] { "Latency", "Warning" },
                new[] { "Total", "Ready" }
            }, style: style)
            .ToBytes();

        string content = Encoding.ASCII.GetString(bytes);
        int footerSeparatorCount = content.Split(new[] { "0.2 0.7 0.3 RG" }, StringSplitOptions.None).Length - 1;

        Assert.Equal(1, footerSeparatorCount);
        Assert.Contains("1.3 w", content);
        Assert.Contains(" l S", content);
        Assert.DoesNotContain(" re S", content);
    }

    [Fact]
    public void RowColumnTable_RendersConfiguredFooterSeparatorAboveFooterRows() {
        var style = TableStyles.Minimal();
        style.BorderColor = null;
        style.RowSeparatorColor = new PdfColor(0.12, 0.34, 0.56);
        style.RowSeparatorWidth = 0.6;
        style.FooterRowCount = 1;
        style.FooterSeparatorColor = new PdfColor(0.2, 0.7, 0.3);
        style.FooterSeparatorWidth = 1.3;

        byte[] bytes = PdfDoc.Create(new PdfOptions {
                PageWidth = 320,
                PageHeight = 220,
                MarginLeft = 30,
                MarginRight = 30,
                MarginTop = 30,
                MarginBottom = 30
            })
            .Compose(document =>
                document.Page(page =>
                    page.Content(content =>
                        content.Row(row =>
                            row.Column(100, column =>
                                column.Table(new[] {
                                    new[] { "Metric", "Status" },
                                    new[] { "Queue", "Healthy" },
                                    new[] { "Latency", "Warning" },
                                    new[] { "Total", "Ready" }
                                }, style: style))))))
            .ToBytes();

        string content = Encoding.ASCII.GetString(bytes);
        int footerSeparatorCount = content.Split(new[] { "0.2 0.7 0.3 RG" }, StringSplitOptions.None).Length - 1;

        Assert.Equal(1, footerSeparatorCount);
        Assert.Contains("1.3 w", content);
        Assert.Contains(" l S", content);
        Assert.DoesNotContain(" re S", content);
    }

    [Fact]
    public void TableStyles_ExposeWordLikeGenericPresetsWithoutSemanticAlignment() {
        var tableGrid = TableStyles.TableGrid();
        var plainTable = TableStyles.PlainTable1();
        var gridTable = TableStyles.GridTable1Light();
        var listTable = TableStyles.ListTable1Light();

        Assert.Equal(PdfColor.FromRgb(191, 191, 191), tableGrid.BorderColor);
        Assert.Equal(0.5, tableGrid.BorderWidth);
        Assert.Null(tableGrid.HeaderFill);
        Assert.Null(tableGrid.RowStripeFill);

        Assert.Null(plainTable.BorderColor);
        Assert.Equal(0, plainTable.BorderWidth);
        Assert.Null(plainTable.RowSeparatorColor);
        Assert.Null(plainTable.HeaderSeparatorColor);

        Assert.Equal(PdfColor.FromRgb(217, 217, 217), gridTable.BorderColor);
        Assert.Equal(PdfColor.FromRgb(127, 127, 127), gridTable.HeaderSeparatorColor);
        Assert.Equal(0.8, gridTable.HeaderSeparatorWidth);
        Assert.Equal(PdfColor.FromRgb(127, 127, 127), gridTable.FooterSeparatorColor);
        Assert.Equal(0.8, gridTable.FooterSeparatorWidth);

        Assert.Null(listTable.BorderColor);
        Assert.Equal(PdfColor.Black, listTable.HeaderSeparatorColor);
        Assert.Equal(PdfColor.Black, listTable.FooterSeparatorColor);
        Assert.Equal(0.8, listTable.FooterSeparatorWidth);
        Assert.Equal(PdfColor.FromRgb(224, 224, 224), listTable.RowSeparatorColor);

        Assert.False(tableGrid.RightAlignNumeric);
        Assert.False(plainTable.RightAlignNumeric);
        Assert.False(gridTable.RightAlignNumeric);
        Assert.False(listTable.RightAlignNumeric);

        var independentGridTable = TableStyles.GridTable1Light();
        gridTable.CellPaddingX = 20;
        Assert.Equal(5, independentGridTable.CellPaddingX);
    }

    [Fact]
    public void TableStyles_ResolveSupportedWordStyleNamesToFreshPdfStyles() {
        Assert.Equal(new[] {
            "TableNormal",
            "TableGrid",
            "PlainTable1",
            "GridTable1Light",
            "GridTable1LightAccent1",
            "GridTable1LightAccent2",
            "GridTable1LightAccent3",
            "GridTable1LightAccent4",
            "GridTable1LightAccent5",
            "GridTable1LightAccent6",
            "GridTable1Light-Accent1",
            "GridTable1Light-Accent2",
            "GridTable1Light-Accent3",
            "GridTable1Light-Accent4",
            "GridTable1Light-Accent5",
            "GridTable1Light-Accent6",
            "ListTable1Light",
            "ListTable1LightAccent1",
            "ListTable1LightAccent2",
            "ListTable1LightAccent3",
            "ListTable1LightAccent4",
            "ListTable1LightAccent5",
            "ListTable1LightAccent6",
            "ListTable1Light-Accent1",
            "ListTable1Light-Accent2",
            "ListTable1Light-Accent3",
            "ListTable1Light-Accent4",
            "ListTable1Light-Accent5",
            "ListTable1Light-Accent6"
        }, TableStyles.SupportedWordStyleNames);

        PdfTableStyle tableNormal = TableStyles.FromWordTableStyle("Table Normal");
        PdfTableStyle tableGrid = TableStyles.FromWordTableStyle("Table Grid");
        PdfTableStyle plainTable = TableStyles.FromWordTableStyle("plain_table_1");
        bool resolvedGridLight = TableStyles.TryFromWordTableStyle("grid-table-1-light", out PdfTableStyle? gridLight);
        PdfTableStyle gridLightAccent = TableStyles.FromWordTableStyle("GridTable1Light-Accent2");
        PdfTableStyle listTable = TableStyles.FromWordTableStyle(" list table 1 light ");
        PdfTableStyle listTableAccent = TableStyles.FromWordTableStyle("ListTable1LightAccent5");

        Assert.Null(tableNormal.BorderColor);
        Assert.Equal(PdfColor.FromRgb(191, 191, 191), tableGrid.BorderColor);
        Assert.Null(plainTable.BorderColor);
        Assert.True(resolvedGridLight);
        Assert.NotNull(gridLight);
        Assert.Equal(PdfColor.FromRgb(217, 217, 217), gridLight!.BorderColor);
        Assert.Equal(PdfColor.FromRgb(127, 127, 127), gridLight.FooterSeparatorColor);
        Assert.Equal(PdfColor.FromRgb(247, 202, 172), gridLightAccent.BorderColor);
        Assert.Equal(PdfColor.FromRgb(244, 176, 131), gridLightAccent.HeaderSeparatorColor);
        Assert.Equal(PdfColor.FromRgb(224, 224, 224), listTable.RowSeparatorColor);
        Assert.Equal(PdfColor.Black, listTable.FooterSeparatorColor);
        Assert.Equal(PdfColor.FromRgb(222, 234, 246), listTableAccent.RowStripeFill);
        Assert.Equal(PdfColor.FromRgb(224, 224, 224), listTableAccent.RowSeparatorColor);
        Assert.Equal(PdfColor.FromRgb(156, 194, 229), listTableAccent.HeaderSeparatorColor);

        PdfTableStyle independentListTable = TableStyles.FromWordTableStyle("ListTable1Light");
        listTable.CellPaddingX = 20;
        Assert.Equal(4, independentListTable.CellPaddingX);

        Assert.False(TableStyles.TryFromWordTableStyle("GridTable7Colorful", out PdfTableStyle? missingStyle));
        Assert.Null(missingStyle);

        var exception = Assert.Throws<ArgumentException>(() => TableStyles.FromWordTableStyle("GridTable7Colorful"));
        Assert.Equal("styleName", exception.ParamName);
        Assert.Contains("Unsupported Word table style 'GridTable7Colorful'.", exception.Message, StringComparison.Ordinal);
        Assert.Contains("Supported styles: TableNormal, TableGrid, PlainTable1, GridTable1Light", exception.Message, StringComparison.Ordinal);
        Assert.Contains("GridTable1Light-Accent6", exception.Message, StringComparison.Ordinal);
        Assert.Contains("ListTable1Light-Accent6", exception.Message, StringComparison.Ordinal);

        Assert.Throws<ArgumentNullException>(() => TableStyles.FromWordTableStyle(null!));
        Assert.Throws<ArgumentNullException>(() => TableStyles.TryFromWordTableStyle(null!, out _));
    }

    [Theory]
    [InlineData(1, 180, 198, 231, 142, 170, 219, 217, 226, 243)]
    [InlineData(2, 247, 202, 172, 244, 176, 131, 251, 228, 213)]
    [InlineData(3, 219, 219, 219, 201, 201, 201, 237, 237, 237)]
    [InlineData(4, 255, 229, 153, 255, 217, 102, 255, 242, 204)]
    [InlineData(5, 189, 214, 238, 156, 194, 229, 222, 234, 246)]
    [InlineData(6, 197, 224, 179, 168, 208, 141, 226, 239, 217)]
    public void TableStyles_ResolveWordAccentVariantsWithDefaultThemeColors(
        int accent,
        int lightR,
        int lightG,
        int lightB,
        int strongR,
        int strongG,
        int strongB,
        int paleR,
        int paleG,
        int paleB) {
        PdfTableStyle grid = TableStyles.FromWordTableStyle("GridTable1Light-Accent" + accent.ToString(CultureInfo.InvariantCulture));
        PdfTableStyle list = TableStyles.FromWordTableStyle("ListTable1LightAccent" + accent.ToString(CultureInfo.InvariantCulture));

        PdfColor ExpectedRgb(int r, int g, int b) => PdfColor.FromRgb((byte)r, (byte)g, (byte)b);

        Assert.Equal(ExpectedRgb(lightR, lightG, lightB), grid.BorderColor);
        Assert.Equal(ExpectedRgb(strongR, strongG, strongB), grid.HeaderSeparatorColor);
        Assert.Equal(ExpectedRgb(strongR, strongG, strongB), grid.FooterSeparatorColor);

        Assert.Equal(ExpectedRgb(paleR, paleG, paleB), list.RowStripeFill);
        Assert.Equal(ExpectedRgb(strongR, strongG, strongB), list.HeaderSeparatorColor);
        Assert.Equal(ExpectedRgb(strongR, strongG, strongB), list.FooterSeparatorColor);
    }

    [Fact]
    public void TableStyles_WordLikePresetsRenderDistinctGridAndListGeometry() {
        string plainContent = RenderTableStyleContent(TableStyles.PlainTable1());
        string gridContent = RenderTableStyleContent(TableStyles.TableGrid());
        string gridLightContent = RenderTableStyleContent(TableStyles.GridTable1Light());
        string listContent = RenderTableStyleContent(TableStyles.ListTable1Light());

        Assert.DoesNotContain(" re S", plainContent);
        Assert.DoesNotContain(" l S", plainContent);

        Assert.Contains(" re S", gridContent);
        Assert.Contains(" l S", gridContent);

        Assert.Contains(" re S", gridLightContent);
        Assert.Contains(" l S", gridLightContent);
        Assert.Contains("0.8 w", gridLightContent);

        Assert.DoesNotContain(" re S", listContent);
        Assert.Contains(" l S", listContent);
        Assert.Contains("0.45 w", listContent);
        Assert.Contains("0.8 w", listContent);
    }

    [Fact]
    public void WordLikeTablePresets_ProvideFooterSeparatorsForSummaryRows() {
        PdfTableStyle gridLight = TableStyles.GridTable1Light();
        PdfTableStyle listTable = TableStyles.ListTable1Light();

        Assert.Equal(PdfColor.FromRgb(127, 127, 127), gridLight.FooterSeparatorColor);
        Assert.Equal(0.8, gridLight.FooterSeparatorWidth);
        Assert.Equal(PdfColor.Black, listTable.FooterSeparatorColor);
        Assert.Equal(0.8, listTable.FooterSeparatorWidth);
    }

    [Fact]
    public void PdfTheme_WordLikeTableStyleIncludesFooterSeparatorByDefault() {
        PdfTheme theme = PdfTheme.WordLike();
        PdfTableStyle style = theme.TableStyle!;

        Assert.Equal(PdfColor.FromRgb(17, 24, 39), style.FooterSeparatorColor);
        Assert.Equal(0.8, style.FooterSeparatorWidth);
    }

    [Fact]
    public void Table_UsesConfiguredMinimumRowHeight() {
        var options = new PdfOptions {
            PageWidth = 320,
            PageHeight = 260,
            MarginLeft = 30,
            MarginRight = 30,
            MarginTop = 30,
            MarginBottom = 30,
            DefaultFont = PdfStandardFont.Helvetica,
            DefaultFontSize = 9
        };
        var style = TableStyles.Minimal();
        style.HeaderRowCount = 0;
        style.MinRowHeight = 36;

        byte[] bytes = PdfDoc.Create(options)
            .Table(new[] {
                new[] { "Alpha", "Ready" },
                new[] { "Beta", "Ready" },
                new[] { "Gamma", "Ready" }
            }, style: style)
            .ToBytes();

        using var pdf = PdfDocument.Open(new MemoryStream(bytes));
        var page = pdf.GetPage(1);

        double alphaY = FindWordStartY(page, "Alpha");
        double betaY = FindWordStartY(page, "Beta");
        double gammaY = FindWordStartY(page, "Gamma");

        Assert.True(alphaY - betaY >= 34, $"Expected minimum row height spacing between first and second row. Alpha y: {alphaY:0.##}, Beta y: {betaY:0.##}.");
        Assert.True(betaY - gammaY >= 34, $"Expected minimum row height spacing between second and third row. Beta y: {betaY:0.##}, Gamma y: {gammaY:0.##}.");
    }

    [Fact]
    public void Table_UsesConfiguredSpacingBeforeAndAfter() {
        var options = new PdfOptions {
            PageWidth = 320,
            PageHeight = 260,
            MarginLeft = 30,
            MarginRight = 30,
            MarginTop = 30,
            MarginBottom = 30,
            DefaultFont = PdfStandardFont.Helvetica,
            DefaultFontSize = 9
        };

        byte[] defaultBytes = CreateTableSpacingProbe(options, spacingBefore: 0, spacingAfter: 0);
        byte[] spacedBytes = CreateTableSpacingProbe(options, spacingBefore: 12, spacingAfter: 18);

        using var defaultPdf = PdfDocument.Open(new MemoryStream(defaultBytes));
        using var spacedPdf = PdfDocument.Open(new MemoryStream(spacedBytes));
        var defaultPage = defaultPdf.GetPage(1);
        var spacedPage = spacedPdf.GetPage(1);

        double defaultTableY = FindWordStartY(defaultPage, "Alpha");
        double spacedTableY = FindWordStartY(spacedPage, "Alpha");
        double defaultAfterY = FindWordStartY(defaultPage, "AfterMarker");
        double spacedAfterY = FindWordStartY(spacedPage, "AfterMarker");

        Assert.True(defaultTableY - spacedTableY >= 10, $"Expected table spacing before to move table content down. Default y: {defaultTableY:0.##}, spaced y: {spacedTableY:0.##}.");
        Assert.True(defaultAfterY - spacedAfterY >= 28, $"Expected table spacing before and after to move following content down. Default y: {defaultAfterY:0.##}, spaced y: {spacedAfterY:0.##}.");
    }

    [Fact]
    public void Table_SuppressesSpacingBeforeAtPageTop() {
        var options = new PdfOptions {
            PageWidth = 320,
            PageHeight = 220,
            MarginLeft = 30,
            MarginRight = 30,
            MarginTop = 30,
            MarginBottom = 30,
            DefaultFont = PdfStandardFont.Helvetica,
            DefaultFontSize = 9
        };
        var defaultStyle = TableStyles.Minimal();
        defaultStyle.HeaderRowCount = 0;
        var spacedStyle = TableStyles.Minimal();
        spacedStyle.HeaderRowCount = 0;
        spacedStyle.SpacingBefore = 28;
        spacedStyle.SpacingAfter = 0;

        byte[] defaultBytes = PdfDoc.Create(options)
            .Table(new[] {
                new[] { "TopTableMarker", "Ready" },
                new[] { "Beta", "Ready" }
            }, style: defaultStyle)
            .ToBytes();
        byte[] spacedBytes = PdfDoc.Create(options)
            .Table(new[] {
                new[] { "TopTableMarker", "Ready" },
                new[] { "Beta", "Ready" }
            }, style: spacedStyle)
            .ToBytes();

        using var defaultPdf = PdfDocument.Open(new MemoryStream(defaultBytes));
        using var spacedPdf = PdfDocument.Open(new MemoryStream(spacedBytes));

        double defaultTopY = FindWordStartY(defaultPdf.GetPage(1), "TopTableMarker");
        double spacedTopY = FindWordStartY(spacedPdf.GetPage(1), "TopTableMarker");

        Assert.InRange(Math.Abs(defaultTopY - spacedTopY), 0, 1.5);
    }

    [Fact]
    public void RowColumnTable_SuppressesSpacingBeforeAtColumnTop() {
        var options = new PdfOptions {
            PageWidth = 320,
            PageHeight = 220,
            MarginLeft = 30,
            MarginRight = 30,
            MarginTop = 30,
            MarginBottom = 30,
            DefaultFont = PdfStandardFont.Helvetica,
            DefaultFontSize = 9
        };
        var defaultStyle = TableStyles.Minimal();
        defaultStyle.HeaderRowCount = 0;
        var spacedStyle = TableStyles.Minimal();
        spacedStyle.HeaderRowCount = 0;
        spacedStyle.SpacingBefore = 28;
        spacedStyle.SpacingAfter = 0;
        string[][] rows = {
            new[] { "ColumnTableMarker", "Ready" },
            new[] { "Beta", "Ready" }
        };

        byte[] defaultBytes = PdfDoc.Create(options)
            .Compose(document => document.Page(page => page.Content(content => content.Row(row => row.Column(100, column => column
                .Table(rows, style: defaultStyle))))))
            .ToBytes();
        byte[] spacedBytes = PdfDoc.Create(options)
            .Compose(document => document.Page(page => page.Content(content => content.Row(row => row.Column(100, column => column
                .Table(rows, style: spacedStyle))))))
            .ToBytes();

        using var defaultPdf = PdfDocument.Open(new MemoryStream(defaultBytes));
        using var spacedPdf = PdfDocument.Open(new MemoryStream(spacedBytes));

        double defaultTopY = FindWordStartY(defaultPdf.GetPage(1), "ColumnTableMarker");
        double spacedTopY = FindWordStartY(spacedPdf.GetPage(1), "ColumnTableMarker");

        Assert.InRange(Math.Abs(defaultTopY - spacedTopY), 0, 1.5);
    }

    [Fact]
    public void Table_RendersConfiguredCaptionAboveGrid() {
        var options = new PdfOptions {
            PageWidth = 320,
            PageHeight = 260,
            MarginLeft = 30,
            MarginRight = 30,
            MarginTop = 30,
            MarginBottom = 30,
            DefaultFont = PdfStandardFont.Helvetica,
            DefaultFontSize = 9
        };
        var style = TableStyles.Minimal();
        style.HeaderRowCount = 0;
        style.Caption = "SignalCaption";
        style.CaptionAlign = PdfAlign.Right;
        style.CaptionColor = PdfColor.FromRgb(80, 90, 100);
        style.CaptionFontSize = 8;
        style.CaptionSpacingAfter = 10;

        byte[] bytes = PdfDoc.Create(options)
            .Table(new[] {
                new[] { "Alpha", "Ready" },
                new[] { "Beta", "Ready" }
            }, style: style)
            .ToBytes();

        using var pdf = PdfDocument.Open(new MemoryStream(bytes));
        var page = pdf.GetPage(1);

        double captionY = FindWordStartY(page, "SignalCaption");
        double alphaY = FindWordStartY(page, "Alpha");
        double captionX = FindWordStartX(page, "SignalCaption");
        double alphaX = FindWordStartX(page, "Alpha");

        Assert.True(captionY > alphaY + 14, $"Expected the table caption above the first row. Caption y: {captionY:0.##}, Alpha y: {alphaY:0.##}.");
        Assert.True(captionX > alphaX + 120, $"Expected the right-aligned caption to render near the table's right edge. Caption x: {captionX:0.##}, Alpha x: {alphaX:0.##}.");

        string content = Encoding.ASCII.GetString(bytes);
        Assert.Contains("0.314 0.353 0.392 rg", content);
    }

    [Fact]
    public void RowColumnTable_RendersConfiguredCaptionAboveGrid() {
        var options = new PdfOptions {
            PageWidth = 320,
            PageHeight = 260,
            MarginLeft = 30,
            MarginRight = 30,
            MarginTop = 30,
            MarginBottom = 30,
            DefaultFont = PdfStandardFont.Helvetica,
            DefaultFontSize = 9
        };
        var style = TableStyles.Minimal();
        style.HeaderRowCount = 0;
        style.Caption = "SignalCaption";
        style.CaptionAlign = PdfAlign.Right;
        style.CaptionColor = PdfColor.FromRgb(80, 90, 100);
        style.CaptionFontSize = 8;
        style.CaptionSpacingAfter = 10;

        byte[] bytes = PdfDoc.Create(options)
            .Compose(document =>
                document.Page(page =>
                    page.Content(content =>
                        content.Row(row =>
                            row.Column(100, column =>
                                column.Table(new[] {
                                    new[] { "Alpha", "Ready" },
                                    new[] { "Beta", "Ready" }
                                }, style: style))))))
            .ToBytes();

        using var pdf = PdfDocument.Open(new MemoryStream(bytes));
        var page = pdf.GetPage(1);

        double captionY = FindWordStartY(page, "SignalCaption");
        double alphaY = FindWordStartY(page, "Alpha");
        double captionX = FindWordStartX(page, "SignalCaption");
        double alphaX = FindWordStartX(page, "Alpha");

        Assert.True(captionY > alphaY + 14, $"Expected the row-column table caption above the first row. Caption y: {captionY:0.##}, Alpha y: {alphaY:0.##}.");
        Assert.True(captionX > alphaX + 120, $"Expected the right-aligned row-column caption to render near the table's right edge. Caption x: {captionX:0.##}, Alpha x: {alphaX:0.##}.");

        string content = Encoding.ASCII.GetString(bytes);
        Assert.Contains("0.314 0.353 0.392 rg", content);
    }

    [Fact]
    public void Table_RejectsCaptionAndFirstRowTallerThanContentArea() {
        var options = new PdfOptions {
            PageWidth = 260,
            PageHeight = 180,
            MarginLeft = 30,
            MarginRight = 30,
            MarginTop = 30,
            MarginBottom = 30,
            DefaultFont = PdfStandardFont.Helvetica,
            DefaultFontSize = 9
        };
        var style = TableStyles.Minimal();
        style.Caption = string.Join(" ", Enumerable.Repeat("TallCaption", 40));
        style.CaptionFontSize = 14;
        style.CaptionSpacingAfter = 8;

        var exception = Assert.Throws<ArgumentException>(() =>
            PdfDoc.Create(options)
                .Table(new[] {
                    new[] { "Alpha", "Ready" },
                    new[] { "Beta", "Ready" }
                }, style: style)
                .ToBytes());

        Assert.Contains("Table caption and first row exceed the available page content height.", exception.Message, StringComparison.Ordinal);
    }

    [Fact]
    public void RowColumnTable_RejectsCaptionAndFirstRowTallerThanContentArea() {
        var options = new PdfOptions {
            PageWidth = 260,
            PageHeight = 180,
            MarginLeft = 30,
            MarginRight = 30,
            MarginTop = 30,
            MarginBottom = 30,
            DefaultFont = PdfStandardFont.Helvetica,
            DefaultFontSize = 9
        };
        var style = TableStyles.Minimal();
        style.Caption = string.Join(" ", Enumerable.Repeat("TallCaption", 40));
        style.CaptionFontSize = 14;
        style.CaptionSpacingAfter = 8;

        var exception = Assert.Throws<ArgumentException>(() =>
            PdfDoc.Create(options)
                .Compose(document =>
                    document.Page(page =>
                        page.Content(content =>
                            content.Row(row =>
                                row.Column(100, column =>
                                    column.Table(new[] {
                                        new[] { "Alpha", "Ready" },
                                        new[] { "Beta", "Ready" }
                                    }, style: style))))))
                .ToBytes());

        Assert.Contains("Table caption and first row exceed the available page content height.", exception.Message, StringComparison.Ordinal);
    }

    [Fact]
    public void Table_UsesRelativeColumnWidthWeights() {
        var options = new PdfOptions {
            PageWidth = 360,
            PageHeight = 240,
            MarginLeft = 30,
            MarginRight = 30,
            MarginTop = 30,
            MarginBottom = 30,
            DefaultFont = PdfStandardFont.Helvetica,
            DefaultFontSize = 9
        };
        var style = TableStyles.Minimal();
        style.ColumnWidthWeights = new List<double> { 1, 3, 1 };

        byte[] bytes = PdfDoc.Create(options)
            .Table(new[] {
                new[] { "ID", "Description", "Score" },
                new[] { "A1", "Longer descriptive value", "100" }
            }, style: style)
            .ToBytes();

        using var pdf = PdfDocument.Open(new MemoryStream(bytes));
        var page = pdf.GetPage(1);

        double idX = FindWordStartX(page, "ID");
        double descriptionX = FindWordStartX(page, "Description");
        double scoreX = FindWordStartX(page, "Score");

        double firstColumnWidth = descriptionX - idX;
        double secondColumnWidth = scoreX - descriptionX;
        Assert.True(secondColumnWidth > firstColumnWidth * 2, $"Expected the middle table column to be visibly wider. First gap: {firstColumnWidth:0.##}, second gap: {secondColumnWidth:0.##}.");
    }

    [Fact]
    public void RowColumnTable_UsesRelativeColumnWidthWeights() {
        var options = new PdfOptions {
            PageWidth = 360,
            PageHeight = 240,
            MarginLeft = 30,
            MarginRight = 30,
            MarginTop = 30,
            MarginBottom = 30,
            DefaultFont = PdfStandardFont.Helvetica,
            DefaultFontSize = 9
        };
        var style = TableStyles.Minimal();
        style.ColumnWidthWeights = new List<double> { 1, 3, 1 };

        byte[] bytes = PdfDoc.Create(options)
            .Compose(document =>
                document.Page(page =>
                    page.Content(content =>
                        content.Row(row =>
                            row.Column(100, column =>
                                column.Table(new[] {
                                    new[] { "ID", "Description", "Score" },
                                    new[] { "A1", "Longer descriptive value", "100" }
                                }, style: style))))))
            .ToBytes();

        using var pdf = PdfDocument.Open(new MemoryStream(bytes));
        var page = pdf.GetPage(1);

        double idX = FindWordStartX(page, "ID");
        double descriptionX = FindWordStartX(page, "Description");
        double scoreX = FindWordStartX(page, "Score");

        double firstColumnWidth = descriptionX - idX;
        double secondColumnWidth = scoreX - descriptionX;
        Assert.True(secondColumnWidth > firstColumnWidth * 2, $"Expected the row-column middle table column to be visibly wider. First gap: {firstColumnWidth:0.##}, second gap: {secondColumnWidth:0.##}.");
    }

    [Fact]
    public void Table_MaxWidthCapsWeightedColumnsAndHonorsAlignment() {
        var options = new PdfOptions {
            PageWidth = 360,
            PageHeight = 240,
            MarginLeft = 30,
            MarginRight = 30,
            MarginTop = 30,
            MarginBottom = 30,
            DefaultFont = PdfStandardFont.Helvetica,
            DefaultFontSize = 9
        };
        var style = TableStyles.Minimal();
        style.MaxWidth = 180;
        style.ColumnWidthWeights = new List<double> { 1, 2 };
        style.HeaderRowCount = 0;

        byte[] bytes = PdfDoc.Create(options)
            .Table(new[] {
                new[] { "Alpha", "Beta" }
            }, align: PdfAlign.Right, style: style)
            .ToBytes();

        using var pdf = PdfDocument.Open(new MemoryStream(bytes));
        var page = pdf.GetPage(1);

        double alphaX = FindWordStartX(page, "Alpha");
        double betaX = FindWordStartX(page, "Beta");

        Assert.InRange(alphaX, 152, 158);
        Assert.InRange(betaX - alphaX, 58, 68);
    }

    [Fact]
    public void RowColumnTable_MaxWidthCapsWeightedColumnsAndHonorsAlignment() {
        var options = new PdfOptions {
            PageWidth = 360,
            PageHeight = 240,
            MarginLeft = 30,
            MarginRight = 30,
            MarginTop = 30,
            MarginBottom = 30,
            DefaultFont = PdfStandardFont.Helvetica,
            DefaultFontSize = 9
        };
        var style = TableStyles.Minimal();
        style.MaxWidth = 180;
        style.ColumnWidthWeights = new List<double> { 1, 2 };
        style.HeaderRowCount = 0;

        byte[] bytes = PdfDoc.Create(options)
            .Compose(document =>
                document.Page(page =>
                    page.Content(content =>
                        content.Row(row =>
                            row.Column(100, column =>
                                column.Table(new[] {
                                    new[] { "Alpha", "Beta" }
                                }, align: PdfAlign.Center, style: style))))))
            .ToBytes();

        using var pdf = PdfDocument.Open(new MemoryStream(bytes));
        var page = pdf.GetPage(1);

        double alphaX = FindWordStartX(page, "Alpha");
        double betaX = FindWordStartX(page, "Beta");

        Assert.InRange(alphaX, 92, 98);
        Assert.InRange(betaX - alphaX, 58, 68);
    }

    [Fact]
    public void Table_LeftIndentOffsetsTableFrameBeforeColumnSizing() {
        var options = new PdfOptions {
            PageWidth = 360,
            PageHeight = 240,
            MarginLeft = 30,
            MarginRight = 30,
            MarginTop = 30,
            MarginBottom = 30,
            DefaultFont = PdfStandardFont.Helvetica,
            DefaultFontSize = 9
        };
        var style = TableStyles.Minimal();
        style.LeftIndent = 60;
        style.MaxWidth = 180;
        style.ColumnWidthWeights = new List<double> { 1, 2 };
        style.HeaderRowCount = 0;

        byte[] bytes = PdfDoc.Create(options)
            .Table(new[] {
                new[] { "Alpha", "Beta" }
            }, align: PdfAlign.Left, style: style)
            .ToBytes();

        using var pdf = PdfDocument.Open(new MemoryStream(bytes));
        var page = pdf.GetPage(1);

        double alphaX = FindWordStartX(page, "Alpha");
        double betaX = FindWordStartX(page, "Beta");

        Assert.InRange(alphaX, 92, 98);
        Assert.InRange(betaX - alphaX, 58, 68);
    }

    [Fact]
    public void RowColumnTable_LeftIndentOffsetsTableFrameBeforeColumnSizing() {
        var options = new PdfOptions {
            PageWidth = 360,
            PageHeight = 240,
            MarginLeft = 30,
            MarginRight = 30,
            MarginTop = 30,
            MarginBottom = 30,
            DefaultFont = PdfStandardFont.Helvetica,
            DefaultFontSize = 9
        };
        var style = TableStyles.Minimal();
        style.LeftIndent = 40;
        style.MaxWidth = 120;
        style.ColumnWidthWeights = new List<double> { 1, 2 };
        style.HeaderRowCount = 0;

        byte[] bytes = PdfDoc.Create(options)
            .Compose(document =>
                document.Page(page =>
                    page.Content(content =>
                        content.Row(row =>
                            row.Column(100, column =>
                                column.Table(new[] {
                                    new[] { "Alpha", "Beta" }
                                }, align: PdfAlign.Left, style: style))))))
            .ToBytes();

        using var pdf = PdfDocument.Open(new MemoryStream(bytes));
        var page = pdf.GetPage(1);

        double alphaX = FindWordStartX(page, "Alpha");
        double betaX = FindWordStartX(page, "Beta");

        Assert.InRange(alphaX, 72, 78);
        Assert.InRange(betaX - alphaX, 38, 48);
    }

    [Fact]
    public void Table_ColumnSpanUsesCombinedColumnWidthAndSnapshotsCells() {
        var options = new PdfOptions {
            PageWidth = 320,
            PageHeight = 220,
            MarginLeft = 30,
            MarginRight = 30,
            MarginTop = 30,
            MarginBottom = 30,
            DefaultFont = PdfStandardFont.Helvetica,
            DefaultFontSize = 9
        };
        var style = TableStyles.Minimal();
        style.HeaderRowCount = 0;
        style.CellPaddingX = 4;
        style.ColumnWidthPoints = new List<double?> { 70, 70, 60 };
        var spanned = PdfTableCell.Span("SpannedTitle", 2);
        var rows = new[] {
            new[] { spanned, PdfTableCell.TextCell("Tail") },
            new[] { PdfTableCell.TextCell("A"), PdfTableCell.TextCell("B"), PdfTableCell.TextCell("C") }
        };

        PdfDoc doc = PdfDoc.Create(options)
            .Table(rows, style: style);

        rows[0][0] = PdfTableCell.TextCell("Mutated");
        byte[] bytes = doc.ToBytes();

        using var pdf = PdfDocument.Open(new MemoryStream(bytes));
        var page = pdf.GetPage(1);

        double spannedX = FindWordStartX(page, "SpannedTitle");
        double tailX = FindWordStartX(page, "Tail");
        double bX = FindWordStartX(page, "B");
        double cX = FindWordStartX(page, "C");

        Assert.InRange(spannedX, 33, 38);
        Assert.InRange(tailX, 173, 178);
        Assert.InRange(bX, 103, 108);
        Assert.InRange(cX, 173, 178);
        Assert.DoesNotContain("Mutated", page.Text, StringComparison.Ordinal);
    }

    [Fact]
    public void RowColumnTable_ColumnSpanUsesCombinedColumnWidth() {
        var options = new PdfOptions {
            PageWidth = 360,
            PageHeight = 240,
            MarginLeft = 30,
            MarginRight = 30,
            MarginTop = 30,
            MarginBottom = 30,
            DefaultFont = PdfStandardFont.Helvetica,
            DefaultFontSize = 9
        };
        var style = TableStyles.Minimal();
        style.HeaderRowCount = 0;
        style.CellPaddingX = 4;
        style.ColumnWidthPoints = new List<double?> { 50, 50, 40 };

        byte[] bytes = PdfDoc.Create(options)
            .Compose(document =>
                document.Page(page =>
                    page.Content(content =>
                        content.Row(row =>
                            row.Column(100, column =>
                                column.Table(new[] {
                                    new[] { PdfTableCell.Span("Merged", 2), PdfTableCell.TextCell("Tail") },
                                    new[] { PdfTableCell.TextCell("A"), PdfTableCell.TextCell("B"), PdfTableCell.TextCell("C") }
                                }, style: style))))))
            .ToBytes();

        using var pdf = PdfDocument.Open(new MemoryStream(bytes));
        var page = pdf.GetPage(1);

        double mergedX = FindWordStartX(page, "Merged");
        double tailX = FindWordStartX(page, "Tail");
        double bX = FindWordStartX(page, "B");
        double cX = FindWordStartX(page, "C");

        Assert.InRange(mergedX, 33, 38);
        Assert.InRange(tailX, 133, 138);
        Assert.InRange(bX, 83, 88);
        Assert.InRange(cX, 133, 138);
    }

    [Fact]
    public void Table_RowSpanOccupiesFollowingRowGridColumn() {
        var options = new PdfOptions {
            PageWidth = 320,
            PageHeight = 220,
            MarginLeft = 30,
            MarginRight = 30,
            MarginTop = 30,
            MarginBottom = 30,
            DefaultFont = PdfStandardFont.Helvetica,
            DefaultFontSize = 9
        };
        var style = TableStyles.Minimal();
        style.HeaderRowCount = 0;
        style.CellPaddingX = 4;
        style.CellPaddingY = 4;
        style.ColumnWidthPoints = new List<double?> { 60, 70, 70 };
        style.VerticalAlignments = new List<PdfCellVerticalAlign> { PdfCellVerticalAlign.Middle };

        byte[] bytes = PdfDoc.Create(options)
            .Table(new[] {
                new[] { PdfTableCell.Merge("GroupOne", rowSpan: 2), PdfTableCell.TextCell("A1"), PdfTableCell.TextCell("B1") },
                new[] { PdfTableCell.TextCell("A2"), PdfTableCell.TextCell("B2") },
                new[] { PdfTableCell.TextCell("Tail0"), PdfTableCell.TextCell("Tail1"), PdfTableCell.TextCell("Tail2") }
            }, style: style)
            .ToBytes();

        using var pdf = PdfDocument.Open(new MemoryStream(bytes));
        var page = pdf.GetPage(1);

        double groupX = FindWordStartX(page, "GroupOne");
        double a1X = FindWordStartX(page, "A1");
        double b1X = FindWordStartX(page, "B1");
        double a2X = FindWordStartX(page, "A2");
        double b2X = FindWordStartX(page, "B2");
        double groupY = FindWordStartY(page, "GroupOne");
        double a1Y = FindWordStartY(page, "A1");
        double a2Y = FindWordStartY(page, "A2");

        Assert.InRange(groupX, 33, 38);
        Assert.InRange(a1X, 93, 99);
        Assert.InRange(b1X, 163, 169);
        Assert.InRange(a2X, 93, 99);
        Assert.InRange(b2X, 163, 169);
        Assert.True(groupY < a1Y - 2 && groupY > a2Y + 2,
            $"Expected vertically centered row-spanned text between row baselines. Group={groupY:0.##}, A1={a1Y:0.##}, A2={a2Y:0.##}.");
    }

    [Fact]
    public void RowColumnTable_RowSpanOccupiesFollowingRowGridColumn() {
        var options = new PdfOptions {
            PageWidth = 360,
            PageHeight = 240,
            MarginLeft = 30,
            MarginRight = 30,
            MarginTop = 30,
            MarginBottom = 30,
            DefaultFont = PdfStandardFont.Helvetica,
            DefaultFontSize = 9
        };
        var style = TableStyles.Minimal();
        style.HeaderRowCount = 0;
        style.CellPaddingX = 4;
        style.CellPaddingY = 4;
        style.ColumnWidthPoints = new List<double?> { 45, 45, 45 };
        style.VerticalAlignments = new List<PdfCellVerticalAlign> { PdfCellVerticalAlign.Middle };

        byte[] bytes = PdfDoc.Create(options)
            .Compose(document =>
                document.Page(page =>
                    page.Content(content =>
                        content.Row(row =>
                            row.Column(100, column =>
                                column.Table(new[] {
                                    new[] { PdfTableCell.Merge("Merge", rowSpan: 2), PdfTableCell.TextCell("A1"), PdfTableCell.TextCell("B1") },
                                    new[] { PdfTableCell.TextCell("A2"), PdfTableCell.TextCell("B2") },
                                    new[] { PdfTableCell.TextCell("C0"), PdfTableCell.TextCell("C1"), PdfTableCell.TextCell("C2") }
                                }, style: style))))))
            .ToBytes();

        using var pdf = PdfDocument.Open(new MemoryStream(bytes));
        var page = pdf.GetPage(1);

        double mergedX = FindWordStartX(page, "Merge");
        double a1X = FindWordStartX(page, "A1");
        double b1X = FindWordStartX(page, "B1");
        double a2X = FindWordStartX(page, "A2");
        double b2X = FindWordStartX(page, "B2");
        double mergedY = FindWordStartY(page, "Merge");
        double a1Y = FindWordStartY(page, "A1");
        double a2Y = FindWordStartY(page, "A2");

        Assert.InRange(mergedX, 33, 38);
        Assert.InRange(a1X, 78, 84);
        Assert.InRange(b1X, 123, 129);
        Assert.InRange(a2X, 78, 84);
        Assert.InRange(b2X, 123, 129);
        Assert.True(mergedY < a1Y - 2 && mergedY > a2Y + 2,
            $"Expected row-column row-spanned text between row baselines. Merged={mergedY:0.##}, A1={a1Y:0.##}, A2={a2Y:0.##}.");
    }

    [Fact]
    public void Table_RowSpanCellFillAndBorderUseCombinedRowHeight() {
        var style = TableStyles.Minimal();
        style.HeaderRowCount = 0;
        style.BorderColor = null;
        style.RowSeparatorColor = null;
        style.HeaderFill = null;
        style.RowStripeFill = null;
        style.CellPaddingX = 4;
        style.CellPaddingY = 4;
        style.MinRowHeight = 24;
        style.ColumnWidthPoints = new List<double?> { 60, 70, 70 };
        style.CellFills = new Dictionary<(int Row, int Column), PdfColor> {
            [(0, 0)] = new PdfColor(0.31, 0.41, 0.51)
        };
        style.CellBorders = new Dictionary<(int Row, int Column), PdfCellBorder> {
            [(0, 0)] = new PdfCellBorder {
                Color = new PdfColor(0.61, 0.21, 0.11),
                Width = 1.4
            }
        };

        byte[] bytes = PdfDoc.Create(new PdfOptions {
                PageWidth = 320,
                PageHeight = 220,
                MarginLeft = 30,
                MarginRight = 30,
                MarginTop = 30,
                MarginBottom = 30,
                DefaultFontSize = 9
            })
            .Table(new[] {
                new[] { PdfTableCell.Merge("Group", rowSpan: 2), PdfTableCell.TextCell("A1"), PdfTableCell.TextCell("B1") },
                new[] { PdfTableCell.TextCell("A2"), PdfTableCell.TextCell("B2") },
                new[] { PdfTableCell.TextCell("C0"), PdfTableCell.TextCell("C1"), PdfTableCell.TextCell("C2") }
            }, style: style)
            .ToBytes();

        string content = string.Join("\n", GetPageContentStreams(bytes, 1));
        var fill = Assert.Single(ExtractPaintedRectangles(content, "0.31 0.41 0.51 rg", "f"));
        var border = Assert.Single(ExtractPaintedRectangles(content, "0.61 0.21 0.11 RG", "S"));

        Assert.InRange(fill.W, 59, 61);
        Assert.True(fill.H > 45, $"Expected row-spanned cell fill to use combined row height. Height: {fill.H:0.##}.");
        Assert.InRange(border.W, 59, 61);
        Assert.True(border.H > 45, $"Expected row-spanned cell border to use combined row height. Height: {border.H:0.##}.");
    }

    [Fact]
    public void RowColumnTable_RowSpanCellFillAndBorderUseCombinedRowHeight() {
        var style = TableStyles.Minimal();
        style.HeaderRowCount = 0;
        style.BorderColor = null;
        style.RowSeparatorColor = null;
        style.HeaderFill = null;
        style.RowStripeFill = null;
        style.CellPaddingX = 4;
        style.CellPaddingY = 4;
        style.MinRowHeight = 24;
        style.ColumnWidthPoints = new List<double?> { 45, 45, 45 };
        style.CellFills = new Dictionary<(int Row, int Column), PdfColor> {
            [(0, 0)] = new PdfColor(0.31, 0.41, 0.51)
        };
        style.CellBorders = new Dictionary<(int Row, int Column), PdfCellBorder> {
            [(0, 0)] = new PdfCellBorder {
                Color = new PdfColor(0.61, 0.21, 0.11),
                Width = 1.4
            }
        };

        byte[] bytes = PdfDoc.Create(new PdfOptions {
                PageWidth = 320,
                PageHeight = 220,
                MarginLeft = 30,
                MarginRight = 30,
                MarginTop = 30,
                MarginBottom = 30,
                DefaultFontSize = 9
            })
            .Compose(document =>
                document.Page(page =>
                    page.Content(content =>
                        content.Row(row =>
                            row.Column(100, column =>
                                column.Table(new[] {
                                    new[] { PdfTableCell.Merge("Group", rowSpan: 2), PdfTableCell.TextCell("A1"), PdfTableCell.TextCell("B1") },
                                    new[] { PdfTableCell.TextCell("A2"), PdfTableCell.TextCell("B2") },
                                    new[] { PdfTableCell.TextCell("C0"), PdfTableCell.TextCell("C1"), PdfTableCell.TextCell("C2") }
                                }, style: style))))))
            .ToBytes();

        string content = string.Join("\n", GetPageContentStreams(bytes, 1));
        var fill = Assert.Single(ExtractPaintedRectangles(content, "0.31 0.41 0.51 rg", "f"));
        var border = Assert.Single(ExtractPaintedRectangles(content, "0.61 0.21 0.11 RG", "S"));

        Assert.InRange(fill.W, 44, 46);
        Assert.True(fill.H > 45, $"Expected row-column row-spanned cell fill to use combined row height. Height: {fill.H:0.##}.");
        Assert.InRange(border.W, 44, 46);
        Assert.True(border.H > 45, $"Expected row-column row-spanned cell border to use combined row height. Height: {border.H:0.##}.");
    }

    [Fact]
    public void Table_RowSpanIgnoresContinuationCellFillAndBorderCoordinates() {
        var style = TableStyles.Minimal();
        style.HeaderRowCount = 0;
        style.BorderColor = null;
        style.RowSeparatorColor = null;
        style.HeaderFill = null;
        style.RowStripeFill = null;
        style.CellPaddingX = 4;
        style.CellPaddingY = 4;
        style.MinRowHeight = 24;
        style.ColumnWidthPoints = new List<double?> { 60, 70, 70 };
        style.CellFills = new Dictionary<(int Row, int Column), PdfColor> {
            [(1, 0)] = new PdfColor(0.31, 0.41, 0.51)
        };
        style.CellBorders = new Dictionary<(int Row, int Column), PdfCellBorder> {
            [(1, 0)] = new PdfCellBorder {
                Color = new PdfColor(0.61, 0.21, 0.11),
                Width = 1.4
            }
        };

        byte[] bytes = PdfDoc.Create(new PdfOptions {
                PageWidth = 320,
                PageHeight = 220,
                MarginLeft = 30,
                MarginRight = 30,
                MarginTop = 30,
                MarginBottom = 30,
                DefaultFontSize = 9
            })
            .Table(new[] {
                new[] { PdfTableCell.Merge("Group", rowSpan: 2), PdfTableCell.TextCell("A1"), PdfTableCell.TextCell("B1") },
                new[] { PdfTableCell.TextCell("A2"), PdfTableCell.TextCell("B2") },
                new[] { PdfTableCell.TextCell("C0"), PdfTableCell.TextCell("C1"), PdfTableCell.TextCell("C2") }
            }, style: style)
            .ToBytes();

        string content = string.Join("\n", GetPageContentStreams(bytes, 1));

        Assert.Empty(ExtractPaintedRectangles(content, "0.31 0.41 0.51 rg", "f"));
        Assert.Empty(ExtractPaintedRectangles(content, "0.61 0.21 0.11 RG", "S"));
    }

    [Fact]
    public void RowColumnTable_RowSpanIgnoresContinuationCellFillAndBorderCoordinates() {
        var style = TableStyles.Minimal();
        style.HeaderRowCount = 0;
        style.BorderColor = null;
        style.RowSeparatorColor = null;
        style.HeaderFill = null;
        style.RowStripeFill = null;
        style.CellPaddingX = 4;
        style.CellPaddingY = 4;
        style.MinRowHeight = 24;
        style.ColumnWidthPoints = new List<double?> { 45, 45, 45 };
        style.CellFills = new Dictionary<(int Row, int Column), PdfColor> {
            [(1, 0)] = new PdfColor(0.31, 0.41, 0.51)
        };
        style.CellBorders = new Dictionary<(int Row, int Column), PdfCellBorder> {
            [(1, 0)] = new PdfCellBorder {
                Color = new PdfColor(0.61, 0.21, 0.11),
                Width = 1.4
            }
        };

        byte[] bytes = PdfDoc.Create(new PdfOptions {
                PageWidth = 320,
                PageHeight = 220,
                MarginLeft = 30,
                MarginRight = 30,
                MarginTop = 30,
                MarginBottom = 30,
                DefaultFontSize = 9
            })
            .Compose(document =>
                document.Page(page =>
                    page.Content(content =>
                        content.Row(row =>
                            row.Column(100, column =>
                                column.Table(new[] {
                                    new[] { PdfTableCell.Merge("Group", rowSpan: 2), PdfTableCell.TextCell("A1"), PdfTableCell.TextCell("B1") },
                                    new[] { PdfTableCell.TextCell("A2"), PdfTableCell.TextCell("B2") },
                                    new[] { PdfTableCell.TextCell("C0"), PdfTableCell.TextCell("C1"), PdfTableCell.TextCell("C2") }
                                }, style: style))))))
            .ToBytes();

        string content = string.Join("\n", GetPageContentStreams(bytes, 1));

        Assert.Empty(ExtractPaintedRectangles(content, "0.31 0.41 0.51 rg", "f"));
        Assert.Empty(ExtractPaintedRectangles(content, "0.61 0.21 0.11 RG", "S"));
    }

    [Fact]
    public void Table_ColumnSpanIgnoresContinuationCellFillAndBorderCoordinates() {
        var style = TableStyles.Minimal();
        style.HeaderRowCount = 0;
        style.BorderColor = null;
        style.RowSeparatorColor = null;
        style.HeaderFill = null;
        style.RowStripeFill = null;
        style.CellPaddingX = 4;
        style.CellPaddingY = 4;
        style.MinRowHeight = 24;
        style.ColumnWidthPoints = new List<double?> { 60, 70, 70 };
        style.CellFills = new Dictionary<(int Row, int Column), PdfColor> {
            [(0, 1)] = new PdfColor(0.31, 0.41, 0.51)
        };
        style.CellBorders = new Dictionary<(int Row, int Column), PdfCellBorder> {
            [(0, 1)] = new PdfCellBorder {
                Color = new PdfColor(0.61, 0.21, 0.11),
                Width = 1.4
            }
        };

        byte[] bytes = PdfDoc.Create(new PdfOptions {
                PageWidth = 320,
                PageHeight = 220,
                MarginLeft = 30,
                MarginRight = 30,
                MarginTop = 30,
                MarginBottom = 30,
                DefaultFontSize = 9
            })
            .Table(new[] {
                new[] { PdfTableCell.Span("Group", 2), PdfTableCell.TextCell("B1") },
                new[] { PdfTableCell.TextCell("A2"), PdfTableCell.TextCell("B2"), PdfTableCell.TextCell("C2") }
            }, style: style)
            .ToBytes();

        string content = string.Join("\n", GetPageContentStreams(bytes, 1));

        Assert.Empty(ExtractPaintedRectangles(content, "0.31 0.41 0.51 rg", "f"));
        Assert.Empty(ExtractPaintedRectangles(content, "0.61 0.21 0.11 RG", "S"));
    }

    [Fact]
    public void RowColumnTable_ColumnSpanIgnoresContinuationCellFillAndBorderCoordinates() {
        var style = TableStyles.Minimal();
        style.HeaderRowCount = 0;
        style.BorderColor = null;
        style.RowSeparatorColor = null;
        style.HeaderFill = null;
        style.RowStripeFill = null;
        style.CellPaddingX = 4;
        style.CellPaddingY = 4;
        style.MinRowHeight = 24;
        style.ColumnWidthPoints = new List<double?> { 45, 45, 45 };
        style.CellFills = new Dictionary<(int Row, int Column), PdfColor> {
            [(0, 1)] = new PdfColor(0.31, 0.41, 0.51)
        };
        style.CellBorders = new Dictionary<(int Row, int Column), PdfCellBorder> {
            [(0, 1)] = new PdfCellBorder {
                Color = new PdfColor(0.61, 0.21, 0.11),
                Width = 1.4
            }
        };

        byte[] bytes = PdfDoc.Create(new PdfOptions {
                PageWidth = 320,
                PageHeight = 220,
                MarginLeft = 30,
                MarginRight = 30,
                MarginTop = 30,
                MarginBottom = 30,
                DefaultFontSize = 9
            })
            .Compose(document =>
                document.Page(page =>
                    page.Content(content =>
                        content.Row(row =>
                            row.Column(100, column =>
                                column.Table(new[] {
                                    new[] { PdfTableCell.Span("Group", 2), PdfTableCell.TextCell("B1") },
                                    new[] { PdfTableCell.TextCell("A2"), PdfTableCell.TextCell("B2"), PdfTableCell.TextCell("C2") }
                                }, style: style))))))
            .ToBytes();

        string content = string.Join("\n", GetPageContentStreams(bytes, 1));

        Assert.Empty(ExtractPaintedRectangles(content, "0.31 0.41 0.51 rg", "f"));
        Assert.Empty(ExtractPaintedRectangles(content, "0.61 0.21 0.11 RG", "S"));
    }

    [Fact]
    public void Table_RowSpanSkipsContinuationRowStripeFillAcrossMergedCell() {
        var style = TableStyles.Minimal();
        style.HeaderRowCount = 0;
        style.BorderColor = null;
        style.RowSeparatorColor = null;
        style.HeaderFill = null;
        style.RowStripeFill = new PdfColor(0.21, 0.31, 0.41);
        style.CellPaddingX = 4;
        style.CellPaddingY = 4;
        style.MinRowHeight = 24;
        style.ColumnWidthPoints = new List<double?> { 60, 70, 70 };

        byte[] bytes = PdfDoc.Create(new PdfOptions {
                PageWidth = 320,
                PageHeight = 220,
                MarginLeft = 30,
                MarginRight = 30,
                MarginTop = 30,
                MarginBottom = 30,
                DefaultFontSize = 9
            })
            .Table(new[] {
                new[] { PdfTableCell.Merge("Group", rowSpan: 2), PdfTableCell.TextCell("A1"), PdfTableCell.TextCell("B1") },
                new[] { PdfTableCell.TextCell("A2"), PdfTableCell.TextCell("B2") },
                new[] { PdfTableCell.TextCell("C0"), PdfTableCell.TextCell("C1"), PdfTableCell.TextCell("C2") }
            }, style: style)
            .ToBytes();

        string content = string.Join("\n", GetPageContentStreams(bytes, 1));
        var stripe = Assert.Single(ExtractPaintedRectangles(content, "0.21 0.31 0.41 rg", "f"));

        Assert.InRange(stripe.X, 89, 91);
        Assert.InRange(stripe.W, 139, 141);
    }

    [Fact]
    public void RowColumnTable_RowSpanSkipsContinuationRowStripeFillAcrossMergedCell() {
        var style = TableStyles.Minimal();
        style.HeaderRowCount = 0;
        style.BorderColor = null;
        style.RowSeparatorColor = null;
        style.HeaderFill = null;
        style.RowStripeFill = new PdfColor(0.21, 0.31, 0.41);
        style.CellPaddingX = 4;
        style.CellPaddingY = 4;
        style.MinRowHeight = 24;
        style.ColumnWidthPoints = new List<double?> { 45, 45, 45 };

        byte[] bytes = PdfDoc.Create(new PdfOptions {
                PageWidth = 320,
                PageHeight = 220,
                MarginLeft = 30,
                MarginRight = 30,
                MarginTop = 30,
                MarginBottom = 30,
                DefaultFontSize = 9
            })
            .Compose(document =>
                document.Page(page =>
                    page.Content(content =>
                        content.Row(row =>
                            row.Column(100, column =>
                                column.Table(new[] {
                                    new[] { PdfTableCell.Merge("Group", rowSpan: 2), PdfTableCell.TextCell("A1"), PdfTableCell.TextCell("B1") },
                                    new[] { PdfTableCell.TextCell("A2"), PdfTableCell.TextCell("B2") },
                                    new[] { PdfTableCell.TextCell("C0"), PdfTableCell.TextCell("C1"), PdfTableCell.TextCell("C2") }
                                }, style: style))))))
            .ToBytes();

        string content = string.Join("\n", GetPageContentStreams(bytes, 1));
        var stripe = Assert.Single(ExtractPaintedRectangles(content, "0.21 0.31 0.41 rg", "f"));

        Assert.InRange(stripe.X, 74, 76);
        Assert.InRange(stripe.W, 89, 91);
    }

    [Fact]
    public void Table_RowSpanSkipsContinuationBodyColumnFillAcrossMergedCell() {
        var style = TableStyles.Minimal();
        style.HeaderRowCount = 0;
        style.BorderColor = null;
        style.RowSeparatorColor = null;
        style.HeaderFill = null;
        style.RowStripeFill = null;
        style.CellPaddingX = 4;
        style.CellPaddingY = 4;
        style.MinRowHeight = 24;
        style.ColumnWidthPoints = new List<double?> { 60, 70, 70 };
        style.BodyColumnFills = new List<PdfColor?> {
            new PdfColor(0.11, 0.22, 0.33)
        };

        byte[] bytes = PdfDoc.Create(new PdfOptions {
                PageWidth = 320,
                PageHeight = 220,
                MarginLeft = 30,
                MarginRight = 30,
                MarginTop = 30,
                MarginBottom = 30,
                DefaultFontSize = 9
            })
            .Table(new[] {
                new[] { PdfTableCell.Merge("Group", rowSpan: 2), PdfTableCell.TextCell("A1"), PdfTableCell.TextCell("B1") },
                new[] { PdfTableCell.TextCell("A2"), PdfTableCell.TextCell("B2") },
                new[] { PdfTableCell.TextCell("C0"), PdfTableCell.TextCell("C1"), PdfTableCell.TextCell("C2") }
            }, style: style)
            .ToBytes();

        string content = string.Join("\n", GetPageContentStreams(bytes, 1));
        var fills = ExtractPaintedRectangles(content, "0.11 0.22 0.33 rg", "f");

        Assert.Equal(2, fills.Count);
        Assert.All(fills, fill => {
            Assert.InRange(fill.X, 29, 31);
            Assert.InRange(fill.W, 59, 61);
        });
    }

    [Fact]
    public void RowColumnTable_RowSpanSkipsContinuationBodyColumnFillAcrossMergedCell() {
        var style = TableStyles.Minimal();
        style.HeaderRowCount = 0;
        style.BorderColor = null;
        style.RowSeparatorColor = null;
        style.HeaderFill = null;
        style.RowStripeFill = null;
        style.CellPaddingX = 4;
        style.CellPaddingY = 4;
        style.MinRowHeight = 24;
        style.ColumnWidthPoints = new List<double?> { 45, 45, 45 };
        style.BodyColumnFills = new List<PdfColor?> {
            new PdfColor(0.11, 0.22, 0.33)
        };

        byte[] bytes = PdfDoc.Create(new PdfOptions {
                PageWidth = 320,
                PageHeight = 220,
                MarginLeft = 30,
                MarginRight = 30,
                MarginTop = 30,
                MarginBottom = 30,
                DefaultFontSize = 9
            })
            .Compose(document =>
                document.Page(page =>
                    page.Content(content =>
                        content.Row(row =>
                            row.Column(100, column =>
                                column.Table(new[] {
                                    new[] { PdfTableCell.Merge("Group", rowSpan: 2), PdfTableCell.TextCell("A1"), PdfTableCell.TextCell("B1") },
                                    new[] { PdfTableCell.TextCell("A2"), PdfTableCell.TextCell("B2") },
                                    new[] { PdfTableCell.TextCell("C0"), PdfTableCell.TextCell("C1"), PdfTableCell.TextCell("C2") }
                                }, style: style))))))
            .ToBytes();

        string content = string.Join("\n", GetPageContentStreams(bytes, 1));
        var fills = ExtractPaintedRectangles(content, "0.11 0.22 0.33 rg", "f");

        Assert.Equal(2, fills.Count);
        Assert.All(fills, fill => {
            Assert.InRange(fill.X, 29, 31);
            Assert.InRange(fill.W, 44, 46);
        });
    }

    [Fact]
    public void Table_ColumnSpanSkipsContinuationBodyColumnFillAcrossMergedCell() {
        var style = TableStyles.Minimal();
        style.HeaderRowCount = 0;
        style.BorderColor = null;
        style.RowSeparatorColor = null;
        style.HeaderFill = null;
        style.RowStripeFill = null;
        style.CellPaddingX = 4;
        style.CellPaddingY = 4;
        style.MinRowHeight = 24;
        style.ColumnWidthPoints = new List<double?> { 60, 70, 70 };
        style.BodyColumnFills = new List<PdfColor?> {
            null,
            new PdfColor(0.11, 0.22, 0.33)
        };

        byte[] bytes = PdfDoc.Create(new PdfOptions {
                PageWidth = 320,
                PageHeight = 220,
                MarginLeft = 30,
                MarginRight = 30,
                MarginTop = 30,
                MarginBottom = 30,
                DefaultFontSize = 9
            })
            .Table(new[] {
                new[] { PdfTableCell.Span("Group", 2), PdfTableCell.TextCell("B1") },
                new[] { PdfTableCell.TextCell("A2"), PdfTableCell.TextCell("B2"), PdfTableCell.TextCell("C2") }
            }, style: style)
            .ToBytes();

        string content = string.Join("\n", GetPageContentStreams(bytes, 1));
        var fill = Assert.Single(ExtractPaintedRectangles(content, "0.11 0.22 0.33 rg", "f"));

        Assert.InRange(fill.X, 89, 91);
        Assert.InRange(fill.W, 69, 71);
    }

    [Fact]
    public void RowColumnTable_ColumnSpanSkipsContinuationBodyColumnFillAcrossMergedCell() {
        var style = TableStyles.Minimal();
        style.HeaderRowCount = 0;
        style.BorderColor = null;
        style.RowSeparatorColor = null;
        style.HeaderFill = null;
        style.RowStripeFill = null;
        style.CellPaddingX = 4;
        style.CellPaddingY = 4;
        style.MinRowHeight = 24;
        style.ColumnWidthPoints = new List<double?> { 45, 45, 45 };
        style.BodyColumnFills = new List<PdfColor?> {
            null,
            new PdfColor(0.11, 0.22, 0.33)
        };

        byte[] bytes = PdfDoc.Create(new PdfOptions {
                PageWidth = 320,
                PageHeight = 220,
                MarginLeft = 30,
                MarginRight = 30,
                MarginTop = 30,
                MarginBottom = 30,
                DefaultFontSize = 9
            })
            .Compose(document =>
                document.Page(page =>
                    page.Content(content =>
                        content.Row(row =>
                            row.Column(100, column =>
                                column.Table(new[] {
                                    new[] { PdfTableCell.Span("Group", 2), PdfTableCell.TextCell("B1") },
                                    new[] { PdfTableCell.TextCell("A2"), PdfTableCell.TextCell("B2"), PdfTableCell.TextCell("C2") }
                                }, style: style))))))
            .ToBytes();

        string content = string.Join("\n", GetPageContentStreams(bytes, 1));
        var fill = Assert.Single(ExtractPaintedRectangles(content, "0.11 0.22 0.33 rg", "f"));

        Assert.InRange(fill.X, 74, 76);
        Assert.InRange(fill.W, 44, 46);
    }

    [Fact]
    public void Table_RowSpanSkipsInternalRowSeparatorAcrossMergedCell() {
        var style = TableStyles.Minimal();
        style.HeaderRowCount = 0;
        style.BorderColor = null;
        style.HeaderFill = null;
        style.RowStripeFill = null;
        style.RowSeparatorColor = new PdfColor(0.12, 0.34, 0.56);
        style.RowSeparatorWidth = 0.6;
        style.CellPaddingX = 4;
        style.CellPaddingY = 4;
        style.MinRowHeight = 24;
        style.ColumnWidthPoints = new List<double?> { 60, 70, 70 };

        byte[] bytes = PdfDoc.Create(new PdfOptions {
                PageWidth = 320,
                PageHeight = 220,
                MarginLeft = 30,
                MarginRight = 30,
                MarginTop = 30,
                MarginBottom = 30,
                DefaultFontSize = 9
            })
            .Table(new[] {
                new[] { PdfTableCell.Merge("Group", rowSpan: 2), PdfTableCell.TextCell("A1"), PdfTableCell.TextCell("B1") },
                new[] { PdfTableCell.TextCell("A2"), PdfTableCell.TextCell("B2") },
                new[] { PdfTableCell.TextCell("C0"), PdfTableCell.TextCell("C1"), PdfTableCell.TextCell("C2") }
            }, style: style)
            .ToBytes();

        string content = string.Join("\n", GetPageContentStreams(bytes, 1));
        var separators = ExtractStrokedLineSegments(content, "0.12 0.34 0.56 RG");
        var partial = Assert.Single(separators, segment => segment.X1 > 88 && segment.X1 < 92 && segment.X2 > 228 && segment.X2 < 232);

        Assert.Contains(separators, segment => segment.X1 > 28 && segment.X1 < 32 && segment.X2 > 228 && segment.X2 < 232);
        Assert.DoesNotContain(separators, segment => Math.Abs(segment.Y1 - partial.Y1) < 0.01 && segment.X1 > 28 && segment.X1 < 32);
    }

    [Fact]
    public void RowColumnTable_RowSpanSkipsInternalRowSeparatorAcrossMergedCell() {
        var style = TableStyles.Minimal();
        style.HeaderRowCount = 0;
        style.BorderColor = null;
        style.HeaderFill = null;
        style.RowStripeFill = null;
        style.RowSeparatorColor = new PdfColor(0.12, 0.34, 0.56);
        style.RowSeparatorWidth = 0.6;
        style.CellPaddingX = 4;
        style.CellPaddingY = 4;
        style.MinRowHeight = 24;
        style.ColumnWidthPoints = new List<double?> { 45, 45, 45 };

        byte[] bytes = PdfDoc.Create(new PdfOptions {
                PageWidth = 320,
                PageHeight = 220,
                MarginLeft = 30,
                MarginRight = 30,
                MarginTop = 30,
                MarginBottom = 30,
                DefaultFontSize = 9
            })
            .Compose(document =>
                document.Page(page =>
                    page.Content(content =>
                        content.Row(row =>
                            row.Column(100, column =>
                                column.Table(new[] {
                                    new[] { PdfTableCell.Merge("Group", rowSpan: 2), PdfTableCell.TextCell("A1"), PdfTableCell.TextCell("B1") },
                                    new[] { PdfTableCell.TextCell("A2"), PdfTableCell.TextCell("B2") },
                                    new[] { PdfTableCell.TextCell("C0"), PdfTableCell.TextCell("C1"), PdfTableCell.TextCell("C2") }
                                }, style: style))))))
            .ToBytes();

        string content = string.Join("\n", GetPageContentStreams(bytes, 1));
        var separators = ExtractStrokedLineSegments(content, "0.12 0.34 0.56 RG");
        var partial = Assert.Single(separators, segment => segment.X1 > 73 && segment.X1 < 77 && segment.X2 > 163 && segment.X2 < 167);

        Assert.Contains(separators, segment => segment.X1 > 28 && segment.X1 < 32 && segment.X2 > 163 && segment.X2 < 167);
        Assert.DoesNotContain(separators, segment => Math.Abs(segment.Y1 - partial.Y1) < 0.01 && segment.X1 > 28 && segment.X1 < 32);
    }

    [Fact]
    public void Table_RowSpanSkipsInternalBorderLineAcrossMergedCell() {
        var style = TableStyles.Minimal();
        style.HeaderRowCount = 0;
        style.BorderColor = new PdfColor(0.12, 0.34, 0.56);
        style.BorderWidth = 0.6;
        style.RowSeparatorColor = null;
        style.HeaderFill = null;
        style.RowStripeFill = null;
        style.CellPaddingX = 4;
        style.CellPaddingY = 4;
        style.MinRowHeight = 24;
        style.ColumnWidthPoints = new List<double?> { 60, 70, 70 };

        byte[] bytes = PdfDoc.Create(new PdfOptions {
                PageWidth = 320,
                PageHeight = 180,
                MarginLeft = 30,
                MarginRight = 30,
                MarginTop = 30,
                MarginBottom = 30,
                DefaultFontSize = 9
            })
            .Table(new[] {
                new[] { PdfTableCell.Merge("Group", rowSpan: 2), PdfTableCell.TextCell("A1"), PdfTableCell.TextCell("B1") },
                new[] { PdfTableCell.TextCell("A2"), PdfTableCell.TextCell("B2") }
            }, style: style)
            .ToBytes();

        string content = string.Join("\n", GetPageContentStreams(bytes, 1));
        var horizontalBorders = ExtractStrokedLineSegments(content, "0.12 0.34 0.56 RG")
            .Where(segment => Math.Abs(segment.Y1 - segment.Y2) < 0.01)
            .ToList();
        var partials = horizontalBorders
            .Where(segment => segment.X1 > 88 && segment.X1 < 92 && segment.X2 > 228 && segment.X2 < 232)
            .ToList();

        Assert.NotEmpty(partials);
        Assert.Contains(horizontalBorders, segment => segment.X1 > 28 && segment.X1 < 32 && segment.X2 > 228 && segment.X2 < 232);
        Assert.DoesNotContain(horizontalBorders, segment => Math.Abs(segment.Y1 - partials[0].Y1) < 0.01 && segment.X1 > 28 && segment.X1 < 32);
        Assert.DoesNotContain(" re S", content);
    }

    [Fact]
    public void RowColumnTable_RowSpanSkipsInternalBorderLineAcrossMergedCell() {
        var style = TableStyles.Minimal();
        style.HeaderRowCount = 0;
        style.BorderColor = new PdfColor(0.12, 0.34, 0.56);
        style.BorderWidth = 0.6;
        style.RowSeparatorColor = null;
        style.HeaderFill = null;
        style.RowStripeFill = null;
        style.CellPaddingX = 4;
        style.CellPaddingY = 4;
        style.MinRowHeight = 24;
        style.ColumnWidthPoints = new List<double?> { 45, 45, 45 };

        byte[] bytes = PdfDoc.Create(new PdfOptions {
                PageWidth = 320,
                PageHeight = 180,
                MarginLeft = 30,
                MarginRight = 30,
                MarginTop = 30,
                MarginBottom = 30,
                DefaultFontSize = 9
            })
            .Compose(document =>
                document.Page(page =>
                    page.Content(content =>
                        content.Row(row =>
                            row.Column(100, column =>
                                column.Table(new[] {
                                    new[] { PdfTableCell.Merge("Group", rowSpan: 2), PdfTableCell.TextCell("A1"), PdfTableCell.TextCell("B1") },
                                    new[] { PdfTableCell.TextCell("A2"), PdfTableCell.TextCell("B2") }
                                }, style: style))))))
            .ToBytes();

        string content = string.Join("\n", GetPageContentStreams(bytes, 1));
        var horizontalBorders = ExtractStrokedLineSegments(content, "0.12 0.34 0.56 RG")
            .Where(segment => Math.Abs(segment.Y1 - segment.Y2) < 0.01)
            .ToList();
        var partials = horizontalBorders
            .Where(segment => segment.X1 > 73 && segment.X1 < 77 && segment.X2 > 163 && segment.X2 < 167)
            .ToList();

        Assert.NotEmpty(partials);
        Assert.Contains(horizontalBorders, segment => segment.X1 > 28 && segment.X1 < 32 && segment.X2 > 163 && segment.X2 < 167);
        Assert.DoesNotContain(horizontalBorders, segment => Math.Abs(segment.Y1 - partials[0].Y1) < 0.01 && segment.X1 > 28 && segment.X1 < 32);
        Assert.DoesNotContain(" re S", content);
    }

    [Fact]
    public void Table_LinkedColumnSpanRendersAnnotationAcrossCombinedCellWidth() {
        var options = new PdfOptions {
            PageWidth = 320,
            PageHeight = 220,
            MarginLeft = 30,
            MarginRight = 30,
            MarginTop = 30,
            MarginBottom = 30,
            DefaultFont = PdfStandardFont.Helvetica,
            DefaultFontSize = 9
        };
        var style = TableStyles.Minimal();
        style.HeaderRowCount = 0;
        style.CellPaddingX = 4;
        style.ColumnWidthPoints = new List<double?> { 70, 70, 60 };

        byte[] bytes = PdfDoc.Create(options)
            .Table(new[] {
                new[] {
                    PdfTableCell.Span("SpannedLink", 2, "https://evotec.xyz/spanned", "Spanned cell metadata"),
                    PdfTableCell.TextCell("Tail")
                },
                new[] { PdfTableCell.TextCell("A"), PdfTableCell.TextCell("B"), PdfTableCell.TextCell("C") }
            }, style: style)
            .ToBytes();

        string pdf = Encoding.ASCII.GetString(bytes);
        var rectangles = ExtractLinkRectangles(pdf);
        var rect = Assert.Single(rectangles);

        Assert.Equal(1, CountOccurrences(pdf, "/Subtype /Link"));
        Assert.Equal(1, CountOccurrences(pdf, "/URI (https://evotec.xyz/spanned)"));
        Assert.Equal(1, CountOccurrences(pdf, "/Contents (Spanned cell metadata)"));
        Assert.InRange(rect.X1, 33, 38);
        Assert.True(rect.X2 - rect.X1 > 120, $"Expected linked spanned cell annotation to cover the combined cell width. Width: {rect.X2 - rect.X1:0.##}.");
    }

    [Fact]
    public void RowColumnTable_LinkedColumnSpanRendersAnnotationAcrossCombinedCellWidth() {
        var options = new PdfOptions {
            PageWidth = 360,
            PageHeight = 240,
            MarginLeft = 30,
            MarginRight = 30,
            MarginTop = 30,
            MarginBottom = 30,
            DefaultFont = PdfStandardFont.Helvetica,
            DefaultFontSize = 9
        };
        var style = TableStyles.Minimal();
        style.HeaderRowCount = 0;
        style.CellPaddingX = 4;
        style.ColumnWidthPoints = new List<double?> { 50, 50, 40 };

        byte[] bytes = PdfDoc.Create(options)
            .Compose(document =>
                document.Page(page =>
                    page.Content(content =>
                        content.Row(row =>
                            row.Column(100, column =>
                                column.Table(new[] {
                                    new[] {
                                        PdfTableCell.Span("MergedLink", 2, "https://evotec.xyz/row-column-spanned", "Row-column spanned metadata"),
                                        PdfTableCell.TextCell("Tail")
                                    },
                                    new[] { PdfTableCell.TextCell("A"), PdfTableCell.TextCell("B"), PdfTableCell.TextCell("C") }
                                }, style: style))))))
            .ToBytes();

        string pdf = Encoding.ASCII.GetString(bytes);
        var rectangles = ExtractLinkRectangles(pdf);
        var rect = Assert.Single(rectangles);

        Assert.Equal(1, CountOccurrences(pdf, "/Subtype /Link"));
        Assert.Equal(1, CountOccurrences(pdf, "/URI (https://evotec.xyz/row-column-spanned)"));
        Assert.Equal(1, CountOccurrences(pdf, "/Contents (Row-column spanned metadata)"));
        Assert.InRange(rect.X1, 33, 38);
        Assert.True(rect.X2 - rect.X1 > 80, $"Expected row-column linked spanned cell annotation to cover the combined cell width. Width: {rect.X2 - rect.X1:0.##}.");
    }

    [Fact]
    public void Table_LinkedRowSpanRendersAnnotationAcrossCombinedCellHeight() {
        var options = new PdfOptions {
            PageWidth = 320,
            PageHeight = 240,
            MarginLeft = 30,
            MarginRight = 30,
            MarginTop = 30,
            MarginBottom = 30,
            DefaultFont = PdfStandardFont.Helvetica,
            DefaultFontSize = 9
        };
        var style = TableStyles.Minimal();
        style.HeaderRowCount = 0;
        style.CellPaddingX = 4;
        style.CellPaddingY = 4;
        style.MinRowHeight = 28;
        style.ColumnWidthPoints = new List<double?> { 60, 70, 70 };

        byte[] bytes = PdfDoc.Create(options)
            .Table(new[] {
                new[] {
                    PdfTableCell.Merge("TallLink", rowSpan: 2, linkUri: "https://evotec.xyz/row-spanned", linkContents: "Row-spanned cell metadata"),
                    PdfTableCell.TextCell("A1"),
                    PdfTableCell.TextCell("B1")
                },
                new[] { PdfTableCell.TextCell("A2"), PdfTableCell.TextCell("B2") },
                new[] { PdfTableCell.TextCell("C0"), PdfTableCell.TextCell("C1"), PdfTableCell.TextCell("C2") }
            }, style: style)
            .ToBytes();

        string pdf = Encoding.ASCII.GetString(bytes);
        var rectangles = ExtractLinkRectangles(pdf);
        var rect = Assert.Single(rectangles);

        Assert.Equal(1, CountOccurrences(pdf, "/Subtype /Link"));
        Assert.Equal(1, CountOccurrences(pdf, "/URI (https://evotec.xyz/row-spanned)"));
        Assert.Equal(1, CountOccurrences(pdf, "/Contents (Row-spanned cell metadata)"));
        Assert.InRange(rect.X1, 33, 38);
        Assert.True(rect.Y2 - rect.Y1 > 40, $"Expected linked row-spanned cell annotation to cover the combined cell height. Height: {rect.Y2 - rect.Y1:0.##}.");
    }

    [Fact]
    public void RowColumnTable_LinkedRowSpanRendersAnnotationAcrossCombinedCellHeight() {
        var options = new PdfOptions {
            PageWidth = 360,
            PageHeight = 260,
            MarginLeft = 30,
            MarginRight = 30,
            MarginTop = 30,
            MarginBottom = 30,
            DefaultFont = PdfStandardFont.Helvetica,
            DefaultFontSize = 9
        };
        var style = TableStyles.Minimal();
        style.HeaderRowCount = 0;
        style.CellPaddingX = 4;
        style.CellPaddingY = 4;
        style.MinRowHeight = 28;
        style.ColumnWidthPoints = new List<double?> { 50, 50, 40 };

        byte[] bytes = PdfDoc.Create(options)
            .Compose(document =>
                document.Page(page =>
                    page.Content(content =>
                        content.Row(row =>
                            row.Column(100, column =>
                                column.Table(new[] {
                                    new[] {
                                        PdfTableCell.Merge("TallLink", rowSpan: 2, linkUri: "https://evotec.xyz/row-column-row-spanned", linkContents: "Row-column row-spanned metadata"),
                                        PdfTableCell.TextCell("A1"),
                                        PdfTableCell.TextCell("B1")
                                    },
                                    new[] { PdfTableCell.TextCell("A2"), PdfTableCell.TextCell("B2") },
                                    new[] { PdfTableCell.TextCell("C0"), PdfTableCell.TextCell("C1"), PdfTableCell.TextCell("C2") }
                                }, style: style))))))
            .ToBytes();

        string pdf = Encoding.ASCII.GetString(bytes);
        var rectangles = ExtractLinkRectangles(pdf);
        var rect = Assert.Single(rectangles);

        Assert.Equal(1, CountOccurrences(pdf, "/Subtype /Link"));
        Assert.Equal(1, CountOccurrences(pdf, "/URI (https://evotec.xyz/row-column-row-spanned)"));
        Assert.Equal(1, CountOccurrences(pdf, "/Contents (Row-column row-spanned metadata)"));
        Assert.InRange(rect.X1, 33, 38);
        Assert.True(rect.Y2 - rect.Y1 > 40, $"Expected row-column linked row-spanned cell annotation to cover the combined cell height. Height: {rect.Y2 - rect.Y1:0.##}.");
    }

    [Fact]
    public void Table_RectangularMergedCellUsesCombinedBoxForFillBorderAndLink() {
        var options = new PdfOptions {
            PageWidth = 340,
            PageHeight = 250,
            MarginLeft = 30,
            MarginRight = 30,
            MarginTop = 30,
            MarginBottom = 30,
            DefaultFont = PdfStandardFont.Helvetica,
            DefaultFontSize = 9
        };
        var style = TableStyles.Minimal();
        style.HeaderRowCount = 0;
        style.BorderColor = null;
        style.RowSeparatorColor = null;
        style.HeaderFill = null;
        style.RowStripeFill = null;
        style.CellPaddingX = 4;
        style.CellPaddingY = 4;
        style.MinRowHeight = 24;
        style.ColumnWidthPoints = new List<double?> { 50, 60, 70 };
        style.CellFills = new Dictionary<(int Row, int Column), PdfColor> {
            [(0, 0)] = new PdfColor(0.23, 0.34, 0.45)
        };
        style.CellBorders = new Dictionary<(int Row, int Column), PdfCellBorder> {
            [(0, 0)] = new PdfCellBorder {
                Color = new PdfColor(0.63, 0.24, 0.14),
                Width = 1.4
            }
        };

        byte[] bytes = PdfDoc.Create(options)
            .Table(new[] {
                new[] {
                    PdfTableCell.Merge("MergedBox", columnSpan: 2, rowSpan: 2, linkUri: "https://evotec.xyz/rectangular-merged", linkContents: "Rectangular merged metadata"),
                    PdfTableCell.TextCell("C1")
                },
                new[] { PdfTableCell.TextCell("C2") },
                new[] { PdfTableCell.TextCell("A3"), PdfTableCell.TextCell("B3"), PdfTableCell.TextCell("C3") }
            }, style: style)
            .ToBytes();

        string pdf = Encoding.ASCII.GetString(bytes);
        string content = string.Join("\n", GetPageContentStreams(bytes, 1));
        var fill = Assert.Single(ExtractPaintedRectangles(content, "0.23 0.34 0.45 rg", "f"));
        var border = Assert.Single(ExtractPaintedRectangles(content, "0.63 0.24 0.14 RG", "S"));
        var link = Assert.Single(ExtractLinkRectangles(pdf));

        Assert.InRange(fill.W, 109, 111);
        Assert.True(fill.H > 45, $"Expected rectangular merged cell fill to use combined row height. Height: {fill.H:0.##}.");
        Assert.InRange(border.W, 109, 111);
        Assert.True(border.H > 45, $"Expected rectangular merged cell border to use combined row height. Height: {border.H:0.##}.");
        Assert.True(link.X2 - link.X1 > 100, $"Expected linked rectangular merged cell to cover combined width. Width: {link.X2 - link.X1:0.##}.");
        Assert.True(link.Y2 - link.Y1 >= 39, $"Expected linked rectangular merged cell to cover combined text-frame height. Height: {link.Y2 - link.Y1:0.##}.");
        Assert.Equal(1, CountOccurrences(pdf, "/URI (https://evotec.xyz/rectangular-merged)"));
        Assert.Equal(1, CountOccurrences(pdf, "/Contents (Rectangular merged metadata)"));
    }

    [Fact]
    public void RowColumnTable_RectangularMergedCellUsesCombinedBoxForFillBorderAndLink() {
        var options = new PdfOptions {
            PageWidth = 360,
            PageHeight = 260,
            MarginLeft = 30,
            MarginRight = 30,
            MarginTop = 30,
            MarginBottom = 30,
            DefaultFont = PdfStandardFont.Helvetica,
            DefaultFontSize = 9
        };
        var style = TableStyles.Minimal();
        style.HeaderRowCount = 0;
        style.BorderColor = null;
        style.RowSeparatorColor = null;
        style.HeaderFill = null;
        style.RowStripeFill = null;
        style.CellPaddingX = 4;
        style.CellPaddingY = 4;
        style.MinRowHeight = 24;
        style.ColumnWidthPoints = new List<double?> { 45, 45, 45 };
        style.CellFills = new Dictionary<(int Row, int Column), PdfColor> {
            [(0, 0)] = new PdfColor(0.23, 0.34, 0.45)
        };
        style.CellBorders = new Dictionary<(int Row, int Column), PdfCellBorder> {
            [(0, 0)] = new PdfCellBorder {
                Color = new PdfColor(0.63, 0.24, 0.14),
                Width = 1.4
            }
        };

        byte[] bytes = PdfDoc.Create(options)
            .Compose(document =>
                document.Page(page =>
                    page.Content(content =>
                        content.Row(row =>
                            row.Column(100, column =>
                                column.Table(new[] {
                                    new[] {
                                        PdfTableCell.Merge("MergedBox", columnSpan: 2, rowSpan: 2, linkUri: "https://evotec.xyz/row-column-rectangular-merged", linkContents: "Row-column rectangular merged metadata"),
                                        PdfTableCell.TextCell("C1")
                                    },
                                    new[] { PdfTableCell.TextCell("C2") },
                                    new[] { PdfTableCell.TextCell("A3"), PdfTableCell.TextCell("B3"), PdfTableCell.TextCell("C3") }
                                }, style: style))))))
            .ToBytes();

        string pdf = Encoding.ASCII.GetString(bytes);
        string content = string.Join("\n", GetPageContentStreams(bytes, 1));
        var fill = Assert.Single(ExtractPaintedRectangles(content, "0.23 0.34 0.45 rg", "f"));
        var border = Assert.Single(ExtractPaintedRectangles(content, "0.63 0.24 0.14 RG", "S"));
        var link = Assert.Single(ExtractLinkRectangles(pdf));

        Assert.InRange(fill.W, 89, 91);
        Assert.True(fill.H > 45, $"Expected row-column rectangular merged cell fill to use combined row height. Height: {fill.H:0.##}.");
        Assert.InRange(border.W, 89, 91);
        Assert.True(border.H > 45, $"Expected row-column rectangular merged cell border to use combined row height. Height: {border.H:0.##}.");
        Assert.True(link.X2 - link.X1 > 80, $"Expected row-column linked rectangular merged cell to cover combined width. Width: {link.X2 - link.X1:0.##}.");
        Assert.True(link.Y2 - link.Y1 >= 39, $"Expected row-column linked rectangular merged cell to cover combined text-frame height. Height: {link.Y2 - link.Y1:0.##}.");
        Assert.Equal(1, CountOccurrences(pdf, "/URI (https://evotec.xyz/row-column-rectangular-merged)"));
        Assert.Equal(1, CountOccurrences(pdf, "/Contents (Row-column rectangular merged metadata)"));
    }

    [Fact]
    public void Table_RectangularMergedCellSkipsInternalVerticalGridOnContinuationRow() {
        var options = new PdfOptions {
            PageWidth = 340,
            PageHeight = 250,
            MarginLeft = 30,
            MarginRight = 30,
            MarginTop = 30,
            MarginBottom = 30,
            DefaultFont = PdfStandardFont.Helvetica,
            DefaultFontSize = 9
        };
        var style = TableStyles.Minimal();
        style.HeaderRowCount = 0;
        style.BorderColor = new PdfColor(0.12, 0.34, 0.56);
        style.BorderWidth = 0.6;
        style.RowSeparatorColor = null;
        style.HeaderFill = null;
        style.RowStripeFill = null;
        style.CellPaddingX = 4;
        style.CellPaddingY = 4;
        style.MinRowHeight = 24;
        style.ColumnWidthPoints = new List<double?> { 50, 60, 70 };

        byte[] bytes = PdfDoc.Create(options)
            .Table(new[] {
                new[] {
                    PdfTableCell.Merge("MergedBox", columnSpan: 2, rowSpan: 2),
                    PdfTableCell.TextCell("C1")
                },
                new[] { PdfTableCell.TextCell("C2") },
                new[] { PdfTableCell.TextCell("A3"), PdfTableCell.TextCell("B3"), PdfTableCell.TextCell("C3") }
            }, style: style)
            .ToBytes();

        string content = string.Join("\n", GetPageContentStreams(bytes, 1));
        var verticalBorders = ExtractStrokedLineSegments(content, "0.12 0.34 0.56 RG")
            .Where(segment => Math.Abs(segment.X1 - segment.X2) < 0.01)
            .ToList();

        Assert.DoesNotContain(verticalBorders, segment => Math.Abs(segment.X1 - 80) < 0.01 && (segment.Y1 + segment.Y2) / 2 > 172);
        Assert.Contains(verticalBorders, segment => Math.Abs(segment.X1 - 80) < 0.01 && (segment.Y1 + segment.Y2) / 2 > 148 && (segment.Y1 + segment.Y2) / 2 < 172);
        Assert.Contains(verticalBorders, segment => Math.Abs(segment.X1 - 140) < 0.01 && (segment.Y1 + segment.Y2) / 2 > 172);
    }

    [Fact]
    public void RowColumnTable_RectangularMergedCellSkipsInternalVerticalGridOnContinuationRow() {
        var options = new PdfOptions {
            PageWidth = 360,
            PageHeight = 260,
            MarginLeft = 30,
            MarginRight = 30,
            MarginTop = 30,
            MarginBottom = 30,
            DefaultFont = PdfStandardFont.Helvetica,
            DefaultFontSize = 9
        };
        var style = TableStyles.Minimal();
        style.HeaderRowCount = 0;
        style.BorderColor = new PdfColor(0.12, 0.34, 0.56);
        style.BorderWidth = 0.6;
        style.RowSeparatorColor = null;
        style.HeaderFill = null;
        style.RowStripeFill = null;
        style.CellPaddingX = 4;
        style.CellPaddingY = 4;
        style.MinRowHeight = 24;
        style.ColumnWidthPoints = new List<double?> { 45, 45, 45 };

        byte[] bytes = PdfDoc.Create(options)
            .Compose(document =>
                document.Page(page =>
                    page.Content(content =>
                        content.Row(row =>
                            row.Column(100, column =>
                                column.Table(new[] {
                                    new[] {
                                        PdfTableCell.Merge("MergedBox", columnSpan: 2, rowSpan: 2),
                                        PdfTableCell.TextCell("C1")
                                    },
                                    new[] { PdfTableCell.TextCell("C2") },
                                    new[] { PdfTableCell.TextCell("A3"), PdfTableCell.TextCell("B3"), PdfTableCell.TextCell("C3") }
                                }, style: style))))))
            .ToBytes();

        string content = string.Join("\n", GetPageContentStreams(bytes, 1));
        var verticalBorders = ExtractStrokedLineSegments(content, "0.12 0.34 0.56 RG")
            .Where(segment => Math.Abs(segment.X1 - segment.X2) < 0.01)
            .ToList();

        Assert.DoesNotContain(verticalBorders, segment => Math.Abs(segment.X1 - 75) < 0.01 && (segment.Y1 + segment.Y2) / 2 > 182);
        Assert.Contains(verticalBorders, segment => Math.Abs(segment.X1 - 75) < 0.01 && (segment.Y1 + segment.Y2) / 2 > 158 && (segment.Y1 + segment.Y2) / 2 < 182);
        Assert.Contains(verticalBorders, segment => Math.Abs(segment.X1 - 120) < 0.01 && (segment.Y1 + segment.Y2) / 2 > 182);
    }

    [Fact]
    public void Table_RectangularMergedCellAlignsTextInsideCombinedBox() {
        var options = new PdfOptions {
            PageWidth = 340,
            PageHeight = 250,
            MarginLeft = 30,
            MarginRight = 30,
            MarginTop = 30,
            MarginBottom = 30,
            DefaultFont = PdfStandardFont.Helvetica,
            DefaultFontSize = 9
        };
        var style = TableStyles.Minimal();
        style.HeaderRowCount = 0;
        style.CellPaddingX = 4;
        style.CellPaddingY = 4;
        style.MinRowHeight = 24;
        style.ColumnWidthPoints = new List<double?> { 50, 60, 70 };
        style.Alignments = new List<PdfColumnAlign> { PdfColumnAlign.Center, PdfColumnAlign.Center, PdfColumnAlign.Center };
        style.VerticalAlignments = new List<PdfCellVerticalAlign> { PdfCellVerticalAlign.Bottom, PdfCellVerticalAlign.Bottom, PdfCellVerticalAlign.Bottom };

        byte[] bytes = PdfDoc.Create(options)
            .Table(new[] {
                new[] {
                    PdfTableCell.Merge("MergedBox", columnSpan: 2, rowSpan: 2),
                    PdfTableCell.TextCell("C1")
                },
                new[] { PdfTableCell.TextCell("C2") },
                new[] { PdfTableCell.TextCell("A3"), PdfTableCell.TextCell("B3"), PdfTableCell.TextCell("C3") }
            }, style: style)
            .ToBytes();

        using var pdf = PdfDocument.Open(new MemoryStream(bytes));
        var page = pdf.GetPage(1);

        double mergedX = FindWordStartX(page, "MergedBox");
        double c2Y = FindWordStartY(page, "C2");
        double mergedY = FindWordStartY(page, "MergedBox");

        Assert.InRange(mergedX, 56, 68);
        Assert.True(Math.Abs(mergedY - c2Y) <= 3,
            $"Expected rectangular merged-cell bottom alignment to place text with the second row baseline. Merged={mergedY:0.##}, C2={c2Y:0.##}.");
    }

    [Fact]
    public void RowColumnTable_RectangularMergedCellAlignsTextInsideCombinedBox() {
        var options = new PdfOptions {
            PageWidth = 360,
            PageHeight = 260,
            MarginLeft = 30,
            MarginRight = 30,
            MarginTop = 30,
            MarginBottom = 30,
            DefaultFont = PdfStandardFont.Helvetica,
            DefaultFontSize = 9
        };
        var style = TableStyles.Minimal();
        style.HeaderRowCount = 0;
        style.CellPaddingX = 4;
        style.CellPaddingY = 4;
        style.MinRowHeight = 24;
        style.ColumnWidthPoints = new List<double?> { 45, 45, 45 };
        style.Alignments = new List<PdfColumnAlign> { PdfColumnAlign.Center, PdfColumnAlign.Center, PdfColumnAlign.Center };
        style.VerticalAlignments = new List<PdfCellVerticalAlign> { PdfCellVerticalAlign.Bottom, PdfCellVerticalAlign.Bottom, PdfCellVerticalAlign.Bottom };

        byte[] bytes = PdfDoc.Create(options)
            .Compose(document =>
                document.Page(page =>
                    page.Content(content =>
                        content.Row(row =>
                            row.Column(100, column =>
                                column.Table(new[] {
                                    new[] {
                                        PdfTableCell.Merge("MergedBox", columnSpan: 2, rowSpan: 2),
                                        PdfTableCell.TextCell("C1")
                                    },
                                    new[] { PdfTableCell.TextCell("C2") },
                                    new[] { PdfTableCell.TextCell("A3"), PdfTableCell.TextCell("B3"), PdfTableCell.TextCell("C3") }
                                }, style: style))))))
            .ToBytes();

        using var pdf = PdfDocument.Open(new MemoryStream(bytes));
        var page = pdf.GetPage(1);

        double mergedX = FindWordStartX(page, "MergedBox");
        double c2Y = FindWordStartY(page, "C2");
        double mergedY = FindWordStartY(page, "MergedBox");

        Assert.InRange(mergedX, 46, 58);
        Assert.True(Math.Abs(mergedY - c2Y) <= 3,
            $"Expected row-column rectangular merged-cell bottom alignment to place text with the second row baseline. Merged={mergedY:0.##}, C2={c2Y:0.##}.");
    }

    [Fact]
    public void Table_UsesFixedColumnWidthPointsWithRemainingWeightedColumns() {
        var options = new PdfOptions {
            PageWidth = 360,
            PageHeight = 240,
            MarginLeft = 30,
            MarginRight = 30,
            MarginTop = 30,
            MarginBottom = 30,
            DefaultFont = PdfStandardFont.Helvetica,
            DefaultFontSize = 9
        };
        var style = TableStyles.Minimal();
        style.ColumnWidthPoints = new List<double?> { 60, null, 50 };

        byte[] bytes = PdfDoc.Create(options)
            .Table(new[] {
                new[] { "ID", "Description", "Score" },
                new[] { "A1", "Longer descriptive value", "100" }
            }, style: style)
            .ToBytes();

        using var pdf = PdfDocument.Open(new MemoryStream(bytes));
        var page = pdf.GetPage(1);

        double idX = FindWordStartX(page, "ID");
        double descriptionX = FindWordStartX(page, "Description");
        double scoreX = FindWordStartX(page, "Score");

        double firstColumnWidth = descriptionX - idX;
        double secondColumnWidth = scoreX - descriptionX;
        Assert.InRange(firstColumnWidth, 55, 65);
        Assert.True(secondColumnWidth > 170, $"Expected the unfixed middle table column to consume remaining width. Second gap: {secondColumnWidth:0.##}.");
    }

    [Fact]
    public void RowColumnTable_UsesFixedColumnWidthPointsWithRemainingWeightedColumns() {
        var options = new PdfOptions {
            PageWidth = 360,
            PageHeight = 240,
            MarginLeft = 30,
            MarginRight = 30,
            MarginTop = 30,
            MarginBottom = 30,
            DefaultFont = PdfStandardFont.Helvetica,
            DefaultFontSize = 9
        };
        var style = TableStyles.Minimal();
        style.ColumnWidthPoints = new List<double?> { 60, null, 50 };

        byte[] bytes = PdfDoc.Create(options)
            .Compose(document =>
                document.Page(page =>
                    page.Content(content =>
                        content.Row(row =>
                            row.Column(100, column =>
                                column.Table(new[] {
                                    new[] { "ID", "Description", "Score" },
                                    new[] { "A1", "Longer descriptive value", "100" }
                                }, style: style))))))
            .ToBytes();

        using var pdf = PdfDocument.Open(new MemoryStream(bytes));
        var page = pdf.GetPage(1);

        double idX = FindWordStartX(page, "ID");
        double descriptionX = FindWordStartX(page, "Description");
        double scoreX = FindWordStartX(page, "Score");

        double firstColumnWidth = descriptionX - idX;
        double secondColumnWidth = scoreX - descriptionX;
        Assert.InRange(firstColumnWidth, 55, 65);
        Assert.True(secondColumnWidth > 170, $"Expected the row-column unfixed middle table column to consume remaining width. Second gap: {secondColumnWidth:0.##}.");
    }

    [Fact]
    public void Table_UsesMinimumColumnWidthPointsForWeightedColumns() {
        var options = new PdfOptions {
            PageWidth = 360,
            PageHeight = 240,
            MarginLeft = 30,
            MarginRight = 30,
            MarginTop = 30,
            MarginBottom = 30,
            DefaultFont = PdfStandardFont.Helvetica,
            DefaultFontSize = 9
        };
        var style = TableStyles.Minimal();
        style.ColumnWidthWeights = new List<double> { 1, 10, 1 };
        style.ColumnMinWidthPoints = new List<double?> { 80, null, null };

        byte[] bytes = PdfDoc.Create(options)
            .Table(new[] {
                new[] { "ID", "Description", "Score" },
                new[] { "A1", "Longer descriptive value", "100" }
            }, style: style)
            .ToBytes();

        using var pdf = PdfDocument.Open(new MemoryStream(bytes));
        var page = pdf.GetPage(1);

        double idX = FindWordStartX(page, "ID");
        double descriptionX = FindWordStartX(page, "Description");
        double firstColumnWidth = descriptionX - idX;

        Assert.InRange(firstColumnWidth, 75, 85);
    }

    [Fact]
    public void RowColumnTable_UsesMinimumColumnWidthPointsForWeightedColumns() {
        var options = new PdfOptions {
            PageWidth = 360,
            PageHeight = 240,
            MarginLeft = 30,
            MarginRight = 30,
            MarginTop = 30,
            MarginBottom = 30,
            DefaultFont = PdfStandardFont.Helvetica,
            DefaultFontSize = 9
        };
        var style = TableStyles.Minimal();
        style.ColumnWidthWeights = new List<double> { 1, 10, 1 };
        style.ColumnMinWidthPoints = new List<double?> { 80, null, null };

        byte[] bytes = PdfDoc.Create(options)
            .Compose(document =>
                document.Page(page =>
                    page.Content(content =>
                        content.Row(row =>
                            row.Column(100, column =>
                                column.Table(new[] {
                                    new[] { "ID", "Description", "Score" },
                                    new[] { "A1", "Longer descriptive value", "100" }
                                }, style: style))))))
            .ToBytes();

        using var pdf = PdfDocument.Open(new MemoryStream(bytes));
        var page = pdf.GetPage(1);

        double idX = FindWordStartX(page, "ID");
        double descriptionX = FindWordStartX(page, "Description");
        double firstColumnWidth = descriptionX - idX;

        Assert.InRange(firstColumnWidth, 75, 85);
    }

    [Fact]
    public void Table_UsesMaximumColumnWidthPointsForWeightedColumns() {
        var options = new PdfOptions {
            PageWidth = 360,
            PageHeight = 240,
            MarginLeft = 30,
            MarginRight = 30,
            MarginTop = 30,
            MarginBottom = 30,
            DefaultFont = PdfStandardFont.Helvetica,
            DefaultFontSize = 9
        };
        var style = TableStyles.Minimal();
        style.ColumnWidthWeights = new List<double> { 1, 10, 1 };
        style.ColumnMaxWidthPoints = new List<double?> { null, 120, null };

        byte[] bytes = PdfDoc.Create(options)
            .Table(new[] {
                new[] { "ID", "Description", "Score" },
                new[] { "A1", "Longer descriptive value", "100" }
            }, style: style)
            .ToBytes();

        using var pdf = PdfDocument.Open(new MemoryStream(bytes));
        var page = pdf.GetPage(1);

        double descriptionX = FindWordStartX(page, "Description");
        double scoreX = FindWordStartX(page, "Score");
        double secondColumnWidth = scoreX - descriptionX;

        Assert.InRange(secondColumnWidth, 115, 125);
    }

    [Fact]
    public void RowColumnTable_UsesMaximumColumnWidthPointsForWeightedColumns() {
        var options = new PdfOptions {
            PageWidth = 360,
            PageHeight = 240,
            MarginLeft = 30,
            MarginRight = 30,
            MarginTop = 30,
            MarginBottom = 30,
            DefaultFont = PdfStandardFont.Helvetica,
            DefaultFontSize = 9
        };
        var style = TableStyles.Minimal();
        style.ColumnWidthWeights = new List<double> { 1, 10, 1 };
        style.ColumnMaxWidthPoints = new List<double?> { null, 120, null };

        byte[] bytes = PdfDoc.Create(options)
            .Compose(document =>
                document.Page(page =>
                    page.Content(content =>
                        content.Row(row =>
                            row.Column(100, column =>
                                column.Table(new[] {
                                    new[] { "ID", "Description", "Score" },
                                    new[] { "A1", "Longer descriptive value", "100" }
                                }, style: style))))))
            .ToBytes();

        using var pdf = PdfDocument.Open(new MemoryStream(bytes));
        var page = pdf.GetPage(1);

        double descriptionX = FindWordStartX(page, "Description");
        double scoreX = FindWordStartX(page, "Score");
        double secondColumnWidth = scoreX - descriptionX;

        Assert.InRange(secondColumnWidth, 115, 125);
    }

    [Fact]
    public void Table_UsesConfiguredVerticalColumnAlignment() {
        var options = new PdfOptions {
            PageWidth = 360,
            PageHeight = 240,
            MarginLeft = 30,
            MarginRight = 30,
            MarginTop = 30,
            MarginBottom = 30,
            DefaultFont = PdfStandardFont.Helvetica,
            DefaultFontSize = 9
        };
        var style = TableStyles.Minimal();
        style.ColumnWidthPoints = new List<double?> { 80, null };
        style.VerticalAlignments = new List<PdfCellVerticalAlign> { PdfCellVerticalAlign.Bottom, PdfCellVerticalAlign.Top };

        byte[] bytes = PdfDoc.Create(options)
            .Table(new[] {
                new[] { "Name", "Notes" },
                new[] {
                    "BottomValue",
                    "This note wraps across several lines so the row becomes tall enough to make vertical alignment visible in the first cell."
                }
            }, style: style)
            .ToBytes();

        using var pdf = PdfDocument.Open(new MemoryStream(bytes));
        var page = pdf.GetPage(1);

        double bottomValueY = FindWordStartY(page, "BottomValue");
        double wrappedFirstLineY = FindWordStartY(page, "This");

        Assert.True(bottomValueY < wrappedFirstLineY - 10, $"Expected the first-column value to sit lower than the top-aligned wrapped text. BottomValue y: {bottomValueY:0.##}, wrapped y: {wrappedFirstLineY:0.##}.");
    }

    [Fact]
    public void RowColumnTable_UsesConfiguredVerticalColumnAlignment() {
        var options = new PdfOptions {
            PageWidth = 360,
            PageHeight = 240,
            MarginLeft = 30,
            MarginRight = 30,
            MarginTop = 30,
            MarginBottom = 30,
            DefaultFont = PdfStandardFont.Helvetica,
            DefaultFontSize = 9
        };
        var style = TableStyles.Minimal();
        style.ColumnWidthPoints = new List<double?> { 80, null };
        style.VerticalAlignments = new List<PdfCellVerticalAlign> { PdfCellVerticalAlign.Bottom, PdfCellVerticalAlign.Top };

        byte[] bytes = PdfDoc.Create(options)
            .Compose(document =>
                document.Page(page =>
                    page.Content(content =>
                        content.Row(row =>
                            row.Column(100, column =>
                                column.Table(new[] {
                                    new[] { "Name", "Notes" },
                                    new[] {
                                        "BottomValue",
                                        "This note wraps across several lines so the row becomes tall enough to make vertical alignment visible in the first cell."
                                    }
                                }, style: style))))))
            .ToBytes();

        using var pdf = PdfDocument.Open(new MemoryStream(bytes));
        var page = pdf.GetPage(1);

        double bottomValueY = FindWordStartY(page, "BottomValue");
        double wrappedFirstLineY = FindWordStartY(page, "This");

        Assert.True(bottomValueY < wrappedFirstLineY - 10, $"Expected the first row-column cell value to sit lower than the top-aligned wrapped text. BottomValue y: {bottomValueY:0.##}, wrapped y: {wrappedFirstLineY:0.##}.");
    }

    [Fact]
    public void Table_AutoFitsFlexibleColumnsFromMeasuredContent() {
        var options = new PdfOptions {
            PageWidth = 360,
            PageHeight = 240,
            MarginLeft = 30,
            MarginRight = 30,
            MarginTop = 30,
            MarginBottom = 30,
            DefaultFont = PdfStandardFont.Helvetica,
            DefaultFontSize = 9
        };
        var style = TableStyles.Minimal();
        style.AutoFitColumns = true;

        byte[] bytes = PdfDoc.Create(options)
            .Table(new[] {
                new[] { "SKU", "Description", "Amount" },
                new[] { "A1", "Managed service renewal with monitoring and incident response", "1250" }
            }, style: style)
            .ToBytes();

        using var pdf = PdfDocument.Open(new MemoryStream(bytes));
        var page = pdf.GetPage(1);

        double skuX = FindWordStartX(page, "SKU");
        double descriptionX = FindWordStartX(page, "Description");
        double amountX = FindWordStartX(page, "Amount");
        double firstColumnWidth = descriptionX - skuX;
        double secondColumnWidth = amountX - descriptionX;
        double rightMost = page.Letters
            .Where(letter => !string.IsNullOrWhiteSpace(letter.Value))
            .Max(letter => letter.EndBaseLine.X);

        Assert.True(secondColumnWidth > firstColumnWidth * 3, $"Expected measured content to make the description column much wider. First gap: {firstColumnWidth:0.##}, second gap: {secondColumnWidth:0.##}.");
        Assert.True(secondColumnWidth > 190, $"Expected measured content to reserve substantial width for the description column. Second gap: {secondColumnWidth:0.##}.");
        Assert.InRange(rightMost, double.NegativeInfinity, options.PageWidth - options.MarginRight + 3);
    }

    [Fact]
    public void RowColumnTable_AutoFitsFlexibleColumnsFromMeasuredContent() {
        var options = new PdfOptions {
            PageWidth = 360,
            PageHeight = 240,
            MarginLeft = 30,
            MarginRight = 30,
            MarginTop = 30,
            MarginBottom = 30,
            DefaultFont = PdfStandardFont.Helvetica,
            DefaultFontSize = 9
        };
        var style = TableStyles.Minimal();
        style.AutoFitColumns = true;

        byte[] bytes = PdfDoc.Create(options)
            .Compose(document =>
                document.Page(page =>
                    page.Content(content =>
                        content.Row(row =>
                            row.Column(100, column =>
                                column.Table(new[] {
                                    new[] { "SKU", "Description", "Amount" },
                                    new[] { "A1", "Managed service renewal with monitoring and incident response", "1250" }
                                }, style: style))))))
            .ToBytes();

        using var pdf = PdfDocument.Open(new MemoryStream(bytes));
        var page = pdf.GetPage(1);

        double skuX = FindWordStartX(page, "SKU");
        double descriptionX = FindWordStartX(page, "Description");
        double amountX = FindWordStartX(page, "Amount");
        double firstColumnWidth = descriptionX - skuX;
        double secondColumnWidth = amountX - descriptionX;
        double rightMost = page.Letters
            .Where(letter => !string.IsNullOrWhiteSpace(letter.Value))
            .Max(letter => letter.EndBaseLine.X);

        Assert.True(secondColumnWidth > firstColumnWidth * 3, $"Expected measured content to make the row-column description column much wider. First gap: {firstColumnWidth:0.##}, second gap: {secondColumnWidth:0.##}.");
        Assert.True(secondColumnWidth > 190, $"Expected measured content to reserve substantial width for the row-column description column. Second gap: {secondColumnWidth:0.##}.");
        Assert.InRange(rightMost, double.NegativeInfinity, options.PageWidth - options.MarginRight + 3);
    }

    [Fact]
    public void Table_RightAlignsCurrencyPercentAndParenthesizedNumbers() {
        var options = new PdfOptions {
            PageWidth = 320,
            PageHeight = 240,
            MarginLeft = 30,
            MarginRight = 30,
            MarginTop = 30,
            MarginBottom = 30,
            DefaultFont = PdfStandardFont.Helvetica,
            DefaultFontSize = 9
        };
        var style = TableStyles.RightAlignedNumbers();
        style.ColumnWidthPoints = new List<double?> { 120, 100 };

        byte[] bytes = PdfDoc.Create(options)
            .Table(new[] {
                new[] { "Metric", "Amount" },
                new[] { "Revenue", "$1,234.50" },
                new[] { "Refund", "(45.20)" },
                new[] { "Margin", "99%" },
                new[] { "EU", "€1,234.50" }
            }, style: style)
            .ToBytes();

        using var pdf = PdfDocument.Open(new MemoryStream(bytes));
        var page = pdf.GetPage(1);

        double dollarEnd = FindWordEndX(page, "$1,234.50");
        double refundEnd = FindWordEndX(page, "(45.20)");
        double percentEnd = FindWordEndX(page, "99%");
        double euroEnd = FindWordEndX(page, "€1,234.50");

        Assert.InRange(Math.Abs(refundEnd - dollarEnd), 0, 3);
        Assert.InRange(Math.Abs(percentEnd - dollarEnd), 0, 3);
        Assert.InRange(Math.Abs(euroEnd - dollarEnd), 0, 3);
    }

    [Fact]
    public void RowColumnTable_RightAlignsCurrencyPercentAndParenthesizedNumbers() {
        var options = new PdfOptions {
            PageWidth = 320,
            PageHeight = 240,
            MarginLeft = 30,
            MarginRight = 30,
            MarginTop = 30,
            MarginBottom = 30,
            DefaultFont = PdfStandardFont.Helvetica,
            DefaultFontSize = 9
        };
        var style = TableStyles.RightAlignedNumbers();
        style.ColumnWidthPoints = new List<double?> { 120, 100 };

        byte[] bytes = PdfDoc.Create(options)
            .Compose(document =>
                document.Page(page =>
                    page.Content(content =>
                        content.Row(row =>
                            row.Column(100, column =>
                                column.Table(new[] {
                                    new[] { "Metric", "Amount" },
                                    new[] { "Revenue", "$1,234.50" },
                                    new[] { "Refund", "(45.20)" },
                                    new[] { "Margin", "99%" },
                                    new[] { "EU", "€1,234.50" }
                                }, style: style))))))
            .ToBytes();

        using var pdf = PdfDocument.Open(new MemoryStream(bytes));
        var page = pdf.GetPage(1);

        double dollarEnd = FindWordEndX(page, "$1,234.50");
        double refundEnd = FindWordEndX(page, "(45.20)");
        double percentEnd = FindWordEndX(page, "99%");
        double euroEnd = FindWordEndX(page, "€1,234.50");

        Assert.InRange(Math.Abs(refundEnd - dollarEnd), 0, 3);
        Assert.InRange(Math.Abs(percentEnd - dollarEnd), 0, 3);
        Assert.InRange(Math.Abs(euroEnd - dollarEnd), 0, 3);
    }

    [Fact]
    public void Table_SplitsSingleTallRowsAcrossPages() {
        var options = new PdfOptions {
            PageWidth = 360,
            PageHeight = 180,
            MarginLeft = 30,
            MarginRight = 30,
            MarginTop = 30,
            MarginBottom = 30,
            DefaultFont = PdfStandardFont.Helvetica,
            DefaultFontSize = 9
        };
        var style = TableStyles.Minimal();
        style.ColumnWidthPoints = new List<double?> { 70, null };

        string longValue = string.Join(" ", Enumerable.Range(1, 60).Select(i => "segment" + i.ToString("00")));

        byte[] bytes = PdfDoc.Create(options)
            .Table(new[] {
                new[] { "Type", "Description" },
                new[] { "Finding", longValue }
            }, style: style)
            .ToBytes();

        using var pdf = PdfDocument.Open(new MemoryStream(bytes));
        Assert.True(pdf.NumberOfPages > 1, "Expected one very tall table row to continue onto another page.");

        for (int pageNumber = 1; pageNumber <= pdf.NumberOfPages; pageNumber++) {
            var page = pdf.GetPage(pageNumber);
            Assert.Contains("Type", page.Text);
            Assert.Contains("Description", page.Text);

            double bottomMost = page.Letters
                .Where(letter => !string.IsNullOrWhiteSpace(letter.Value))
                .Min(letter => letter.StartBaseLine.Y);
            Assert.True(bottomMost >= options.MarginBottom - 2, $"Expected split row text to stay above the bottom margin on page {pageNumber}.");
        }

        Assert.Contains("segment01", pdf.GetPage(1).Text);
        Assert.Contains("segment60", pdf.GetPage(pdf.NumberOfPages).Text);
    }

    [Fact]
    public void RowColumnTable_SplitsSingleTallRowsAcrossPages() {
        var options = new PdfOptions {
            PageWidth = 360,
            PageHeight = 180,
            MarginLeft = 30,
            MarginRight = 30,
            MarginTop = 30,
            MarginBottom = 30,
            DefaultFont = PdfStandardFont.Helvetica,
            DefaultFontSize = 9
        };
        var style = TableStyles.Minimal();
        style.ColumnWidthPoints = new List<double?> { 70, null };

        string longValue = string.Join(" ", Enumerable.Range(1, 60).Select(i => "segment" + i.ToString("00")));

        byte[] bytes = PdfDoc.Create(options)
            .Compose(document =>
                document.Page(page =>
                    page.Content(content =>
                        content.Row(row =>
                            row.Column(100, column =>
                                column.Table(new[] {
                                    new[] { "Type", "Description" },
                                    new[] { "Finding", longValue }
                                }, style: style))))))
            .ToBytes();

        using var pdf = PdfDocument.Open(new MemoryStream(bytes));
        Assert.True(pdf.NumberOfPages > 1, "Expected one very tall row-column table row to continue onto another page.");

        for (int pageNumber = 1; pageNumber <= pdf.NumberOfPages; pageNumber++) {
            var page = pdf.GetPage(pageNumber);
            Assert.Contains("Type", page.Text);
            Assert.Contains("Description", page.Text);

            double bottomMost = page.Letters
                .Where(letter => !string.IsNullOrWhiteSpace(letter.Value))
                .Min(letter => letter.StartBaseLine.Y);
            Assert.True(bottomMost >= options.MarginBottom - 2, $"Expected split row-column row text to stay above the bottom margin on page {pageNumber}.");
        }

        Assert.Contains("segment01", pdf.GetPage(1).Text);
        Assert.Contains("segment60", pdf.GetPage(pdf.NumberOfPages).Text);
    }

    [Fact]
    public void Table_DisallowRowBreakRejectsSingleTallRows() {
        var style = TableStyles.Minimal();
        style.AllowRowBreakAcrossPages = false;
        style.ColumnWidthPoints = new List<double?> { 70, null };

        string longValue = string.Join(" ", Enumerable.Range(1, 60).Select(i => "segment" + i.ToString("00")));

        var exception = Assert.Throws<ArgumentException>(() =>
            PdfDoc.Create(new PdfOptions {
                    PageWidth = 360,
                    PageHeight = 180,
                    MarginLeft = 30,
                    MarginRight = 30,
                    MarginTop = 30,
                    MarginBottom = 30,
                    DefaultFont = PdfStandardFont.Helvetica,
                    DefaultFontSize = 9
                })
                .Table(new[] {
                    new[] { "Type", "Description" },
                    new[] { "Finding", longValue }
                }, style: style)
                .ToBytes());

        Assert.Contains("Table row height exceeds the available page content height and row splitting is disabled.", exception.Message, StringComparison.Ordinal);
    }

    [Fact]
    public void RowColumnTable_DisallowRowBreakRejectsSingleTallRows() {
        var style = TableStyles.Minimal();
        style.AllowRowBreakAcrossPages = false;
        style.ColumnWidthPoints = new List<double?> { 70, null };

        string longValue = string.Join(" ", Enumerable.Range(1, 60).Select(i => "segment" + i.ToString("00")));

        var exception = Assert.Throws<ArgumentException>(() =>
            PdfDoc.Create(new PdfOptions {
                    PageWidth = 360,
                    PageHeight = 180,
                    MarginLeft = 30,
                    MarginRight = 30,
                    MarginTop = 30,
                    MarginBottom = 30,
                    DefaultFont = PdfStandardFont.Helvetica,
                    DefaultFontSize = 9
                })
                .Compose(document =>
                    document.Page(page =>
                        page.Content(content =>
                            content.Row(row =>
                                row.Column(100, column =>
                                    column.Table(new[] {
                                        new[] { "Type", "Description" },
                                        new[] { "Finding", longValue }
                                    }, style: style))))))
                .ToBytes());

        Assert.Contains("Table row height exceeds the available page content height and row splitting is disabled.", exception.Message, StringComparison.Ordinal);
    }

    [Fact]
    public void Table_RejectsInvalidRelativeColumnWidthWeights() {
        var style = TableStyles.Minimal();

        var exception = Assert.Throws<ArgumentException>(() =>
            style.ColumnWidthWeights = new List<double> { 1, 0, 1 });

        Assert.Contains("Table column width weights must be positive finite values.", exception.Message, StringComparison.Ordinal);
    }

    [Fact]
    public void Table_RejectsInvalidFixedColumnWidthPoints() {
        var invalidStyle = TableStyles.Minimal();

        var invalidException = Assert.Throws<ArgumentException>(() =>
            invalidStyle.ColumnWidthPoints = new List<double?> { 1, -5 });

        Assert.Contains("Table fixed column widths must be positive finite values.", invalidException.Message, StringComparison.Ordinal);

        var tooWideStyle = TableStyles.Minimal();
        tooWideStyle.ColumnWidthPoints = new List<double?> { 400, 400 };

        Assert.Throws<ArgumentException>(() =>
            PdfDoc.Create()
                .Table(new[] {
                    new[] { "A", "B" },
                    new[] { "1", "2" }
                }, style: tooWideStyle)
                .ToBytes());
    }

    [Fact]
    public void Table_RejectsInvalidMinimumAndMaximumColumnWidthPoints() {
        var invalidMinimum = TableStyles.Minimal();

        var invalidMinimumException = Assert.Throws<ArgumentException>(() =>
            invalidMinimum.ColumnMinWidthPoints = new List<double?> { 0 });

        Assert.Contains("Table minimum column widths must be positive finite values.", invalidMinimumException.Message, StringComparison.Ordinal);

        var invertedRange = TableStyles.Minimal();
        invertedRange.ColumnMinWidthPoints = new List<double?> { 90 };
        invertedRange.ColumnMaxWidthPoints = new List<double?> { 60 };

        Assert.Throws<ArgumentException>(() =>
            PdfDoc.Create()
                .Table(new[] {
                    new[] { "A", "B" },
                    new[] { "1", "2" }
                }, style: invertedRange)
                .ToBytes());

        var impossibleMinimums = TableStyles.Minimal();
        impossibleMinimums.ColumnMinWidthPoints = new List<double?> { 400, 400 };

        Assert.Throws<ArgumentException>(() =>
            PdfDoc.Create()
                .Table(new[] {
                    new[] { "A", "B" },
                    new[] { "1", "2" }
                }, style: impossibleMinimums)
                .ToBytes());

        var invalidHeaderRows = TableStyles.Minimal();

        var invalidHeaderRowsException = Assert.Throws<ArgumentException>(() =>
            invalidHeaderRows.HeaderRowCount = -1);

        Assert.Contains("Table header row count cannot be negative.", invalidHeaderRowsException.Message, StringComparison.Ordinal);

        var invalidFooterRows = TableStyles.Minimal();

        var invalidFooterRowsException = Assert.Throws<ArgumentException>(() =>
            invalidFooterRows.FooterRowCount = -1);

        Assert.Contains("Table footer row count cannot be negative.", invalidFooterRowsException.Message, StringComparison.Ordinal);

        var invalidMinimumRowHeight = TableStyles.Minimal();

        var invalidMinimumRowHeightException = Assert.Throws<ArgumentException>(() =>
            invalidMinimumRowHeight.MinRowHeight = -1);

        Assert.Contains("Table minimum row height must be a non-negative finite value.", invalidMinimumRowHeightException.Message, StringComparison.Ordinal);

        var invalidSpacingBefore = TableStyles.Minimal();

        var invalidSpacingBeforeException = Assert.Throws<ArgumentException>(() =>
            invalidSpacingBefore.SpacingBefore = -1);

        Assert.Contains("Table spacing before must be a non-negative finite value.", invalidSpacingBeforeException.Message, StringComparison.Ordinal);

        var invalidSpacingAfter = TableStyles.Minimal();

        var invalidSpacingAfterException = Assert.Throws<ArgumentException>(() =>
            invalidSpacingAfter.SpacingAfter = double.PositiveInfinity);

        Assert.Contains("Table spacing after must be a non-negative finite value.", invalidSpacingAfterException.Message, StringComparison.Ordinal);

        var invalidCaptionFontSize = TableStyles.Minimal();
        invalidCaptionFontSize.Caption = "Caption";

        var invalidCaptionFontSizeException = Assert.Throws<ArgumentException>(() =>
            invalidCaptionFontSize.CaptionFontSize = 0);

        Assert.Contains("Table caption font size must be a positive finite value.", invalidCaptionFontSizeException.Message, StringComparison.Ordinal);

        var invalidBodyFontSize = TableStyles.Minimal();

        var invalidBodyFontSizeException = Assert.Throws<ArgumentException>(() =>
            invalidBodyFontSize.FontSize = 0);

        Assert.Contains("Table body font size must be a positive finite value.", invalidBodyFontSizeException.Message, StringComparison.Ordinal);

        var invalidLineHeight = TableStyles.Minimal();

        var invalidLineHeightException = Assert.Throws<ArgumentException>(() =>
            invalidLineHeight.LineHeight = double.NaN);

        Assert.Contains("Table line height must be a positive finite value.", invalidLineHeightException.Message, StringComparison.Ordinal);

        var invalidHeaderFontSize = TableStyles.Minimal();

        var invalidHeaderFontSizeException = Assert.Throws<ArgumentException>(() =>
            invalidHeaderFontSize.HeaderFontSize = double.NaN);

        Assert.Contains("Table header font size must be a positive finite value.", invalidHeaderFontSizeException.Message, StringComparison.Ordinal);

        var invalidFooterFontSize = TableStyles.Minimal();

        var invalidFooterFontSizeException = Assert.Throws<ArgumentException>(() =>
            invalidFooterFontSize.FooterFontSize = double.PositiveInfinity);

        Assert.Contains("Table footer font size must be a positive finite value.", invalidFooterFontSizeException.Message, StringComparison.Ordinal);

        var whitespaceCaption = TableStyles.Minimal();
        whitespaceCaption.Caption = "   ";

        var whitespaceCaptionException = Assert.Throws<ArgumentException>(() =>
            PdfDoc.Create()
                .Table(new[] {
                    new[] { "A", "B" },
                    new[] { "1", "2" }
                }, style: whitespaceCaption)
                .ToBytes());

        Assert.Contains("Table caption cannot be empty or whitespace.", whitespaceCaptionException.Message, StringComparison.Ordinal);

        var invalidCaptionSpacing = TableStyles.Minimal();
        invalidCaptionSpacing.Caption = "Caption";

        var invalidCaptionSpacingException = Assert.Throws<ArgumentException>(() =>
            invalidCaptionSpacing.CaptionSpacingAfter = -1);

        Assert.Contains("Table caption spacing after must be a non-negative finite value.", invalidCaptionSpacingException.Message, StringComparison.Ordinal);

        var invalidCellBorder = TableStyles.Minimal();

        var invalidCellBorderException = Assert.Throws<ArgumentException>(() =>
            invalidCellBorder.CellBorders = new Dictionary<(int Row, int Column), PdfCellBorder> {
                [(1, 1)] = new PdfCellBorder {
                    Width = -0.5
                }
            });

        Assert.Contains("Table cell border widths must be non-negative finite values.", invalidCellBorderException.Message, StringComparison.Ordinal);
    }

    [Theory]
    [InlineData(-0.1)]
    [InlineData(double.NaN)]
    [InlineData(double.PositiveInfinity)]
    public void CellBorder_RejectsInvalidWidthOnAssignment(double width) {
        var border = new PdfCellBorder();

        var exception = Assert.Throws<ArgumentException>(() =>
            border.Width = width);

        Assert.Equal("Width", exception.ParamName);
        Assert.Contains("Table cell border widths must be non-negative finite values.", exception.Message, StringComparison.Ordinal);
    }

    [Fact]
    public void TableCell_RejectsInvalidColumnSpan() {
        var exception = Assert.Throws<ArgumentOutOfRangeException>(() =>
            new PdfTableCell("Invalid", 0));

        Assert.Equal("columnSpan", exception.ParamName);
        Assert.Contains("Table cell column span must be at least 1.", exception.Message, StringComparison.Ordinal);
    }

    [Fact]
    public void TableCell_RejectsInvalidRowSpan() {
        var exception = Assert.Throws<ArgumentOutOfRangeException>(() =>
            new PdfTableCell("Invalid", rowSpan: 0));

        Assert.Equal("rowSpan", exception.ParamName);
        Assert.Contains("Table cell row span must be at least 1.", exception.Message, StringComparison.Ordinal);
    }

    [Fact]
    public void Table_RejectsRowSpanBeyondAvailableRows() {
        var exception = Assert.Throws<ArgumentException>(() =>
            PdfDoc.Create()
                .Table(new[] {
                    new[] { PdfTableCell.Merge("TooTall", rowSpan: 3), PdfTableCell.TextCell("A1") },
                    new[] { PdfTableCell.TextCell("A2") }
                }));

        Assert.Equal("rows", exception.ParamName);
        Assert.Contains("Table cell row span cannot extend beyond the available table rows.", exception.Message, StringComparison.Ordinal);
    }

    [Fact]
    public void RowColumnTable_RejectsRowSpanBeyondAvailableRows() {
        var exception = Assert.Throws<ArgumentException>(() =>
            PdfDoc.Create()
                .Compose(document =>
                    document.Page(page =>
                        page.Content(content =>
                            content.Row(row =>
                                row.Column(100, column =>
                                    column.Table(new[] {
                                        new[] { PdfTableCell.Merge("TooTall", rowSpan: 3), PdfTableCell.TextCell("A1") },
                                        new[] { PdfTableCell.TextCell("A2") }
                                    })))))));

        Assert.Equal("rows", exception.ParamName);
        Assert.Contains("Table cell row span cannot extend beyond the available table rows.", exception.Message, StringComparison.Ordinal);
    }

    [Fact]
    public void Table_RejectsRowSpanCrossingHeaderBoundary() {
        var style = TableStyles.Minimal();
        style.HeaderRowCount = 1;

        var exception = Assert.Throws<ArgumentException>(() =>
            PdfDoc.Create()
                .Table(new[] {
                    new[] { PdfTableCell.Merge("HeaderBody", rowSpan: 2), PdfTableCell.TextCell("H1") },
                    new[] { PdfTableCell.TextCell("B1") },
                    new[] { PdfTableCell.TextCell("B2"), PdfTableCell.TextCell("B3") }
                }, style: style)
                .ToBytes());

        Assert.Contains("Table cell row span cannot cross the table header boundary.", exception.Message, StringComparison.Ordinal);
    }

    [Fact]
    public void Table_RejectsRowSpanCrossingFooterBoundary() {
        var style = TableStyles.Minimal();
        style.HeaderRowCount = 0;
        style.FooterRowCount = 1;

        var exception = Assert.Throws<ArgumentException>(() =>
            PdfDoc.Create()
                .Table(new[] {
                    new[] { PdfTableCell.TextCell("B0"), PdfTableCell.TextCell("B1") },
                    new[] { PdfTableCell.Merge("BodyFooter", rowSpan: 2), PdfTableCell.TextCell("B2") },
                    new[] { PdfTableCell.TextCell("F1") }
                }, style: style)
                .ToBytes());

        Assert.Contains("Table cell row span cannot cross the table footer boundary.", exception.Message, StringComparison.Ordinal);
    }

    [Fact]
    public void RowColumnTable_RejectsRowSpanCrossingHeaderBoundary() {
        var style = TableStyles.Minimal();
        style.HeaderRowCount = 1;

        var exception = Assert.Throws<ArgumentException>(() =>
            PdfDoc.Create()
                .Compose(document =>
                    document.Page(page =>
                        page.Content(content =>
                            content.Row(row =>
                                row.Column(100, column =>
                                    column.Table(new[] {
                                        new[] { PdfTableCell.Merge("HeaderBody", rowSpan: 2), PdfTableCell.TextCell("H1") },
                                        new[] { PdfTableCell.TextCell("B1") },
                                        new[] { PdfTableCell.TextCell("B2"), PdfTableCell.TextCell("B3") }
                                    }, style: style))))))
                .ToBytes());

        Assert.Contains("Table cell row span cannot cross the table header boundary.", exception.Message, StringComparison.Ordinal);
    }

    [Fact]
    public void RowColumnTable_RejectsRowSpanCrossingFooterBoundary() {
        var style = TableStyles.Minimal();
        style.HeaderRowCount = 0;
        style.FooterRowCount = 1;

        var exception = Assert.Throws<ArgumentException>(() =>
            PdfDoc.Create()
                .Compose(document =>
                    document.Page(page =>
                        page.Content(content =>
                            content.Row(row =>
                                row.Column(100, column =>
                                    column.Table(new[] {
                                        new[] { PdfTableCell.TextCell("B0"), PdfTableCell.TextCell("B1") },
                                        new[] { PdfTableCell.Merge("BodyFooter", rowSpan: 2), PdfTableCell.TextCell("B2") },
                                        new[] { PdfTableCell.TextCell("F1") }
                                    }, style: style))))))
                .ToBytes());

        Assert.Contains("Table cell row span cannot cross the table footer boundary.", exception.Message, StringComparison.Ordinal);
    }

    [Fact]
    public void Table_RejectsHeaderRowCountBeyondRows() {
        var style = TableStyles.Minimal();
        style.HeaderRowCount = 3;

        var exception = Assert.Throws<ArgumentException>(() =>
            PdfDoc.Create()
                .Table(new[] {
                    new[] { "H1", "H2" },
                    new[] { "B1", "B2" }
                }, style: style)
                .ToBytes());

        Assert.Contains("Table header row count cannot exceed the table row count.", exception.Message, StringComparison.Ordinal);
    }

    [Fact]
    public void Table_RejectsCombinedHeaderAndFooterRowsBeyondRows() {
        var style = TableStyles.Minimal();
        style.HeaderRowCount = 1;
        style.FooterRowCount = 2;

        var exception = Assert.Throws<ArgumentException>(() =>
            PdfDoc.Create()
                .Table(new[] {
                    new[] { "H1", "H2" },
                    new[] { "B1", "B2" }
                }, style: style)
                .ToBytes());

        Assert.Contains("Table header and footer row counts cannot exceed the table row count.", exception.Message, StringComparison.Ordinal);
    }

    [Fact]
    public void RowColumnTable_RejectsFooterRowCountBeyondRows() {
        var style = TableStyles.Minimal();
        style.HeaderRowCount = 0;
        style.FooterRowCount = 3;

        var exception = Assert.Throws<ArgumentException>(() =>
            PdfDoc.Create()
                .Compose(document =>
                    document.Page(page =>
                        page.Content(content =>
                            content.Row(row =>
                                row.Column(100, column =>
                                    column.Table(new[] {
                                        new[] { "B1", "B2" },
                                        new[] { "B3", "B4" }
                                    }, style: style))))))
                .ToBytes());

        Assert.Contains("Table footer row count cannot exceed the table row count.", exception.Message, StringComparison.Ordinal);
    }

    [Fact]
    public void RowColumnTable_RejectsCombinedHeaderAndFooterRowsBeyondRows() {
        var style = TableStyles.Minimal();
        style.HeaderRowCount = 1;
        style.FooterRowCount = 2;

        var exception = Assert.Throws<ArgumentException>(() =>
            PdfDoc.Create()
                .Compose(document =>
                    document.Page(page =>
                        page.Content(content =>
                            content.Row(row =>
                                row.Column(100, column =>
                                    column.Table(new[] {
                                        new[] { "H1", "H2" },
                                        new[] { "B1", "B2" }
                                    }, style: style))))))
                .ToBytes());

        Assert.Contains("Table header and footer row counts cannot exceed the table row count.", exception.Message, StringComparison.Ordinal);
    }

    [Fact]
    public void ParagraphKeepWithNext_RejectsFollowingTableWithHeaderRowCountBeyondRows() {
        var style = TableStyles.Minimal();
        style.HeaderRowCount = 3;

        var exception = Assert.Throws<ArgumentException>(() =>
            PdfDoc.Create()
                .Paragraph(p => p.Text("KeepWithNextPrelude"), style: new PdfParagraphStyle {
                    KeepWithNext = true
                })
                .Table(new[] {
                    new[] { "H1", "H2" },
                    new[] { "B1", "B2" }
                }, style: style)
                .ToBytes());

        Assert.Contains("Table header row count cannot exceed the table row count.", exception.Message, StringComparison.Ordinal);
    }

    [Fact]
    public void ParagraphKeepWithNext_RejectsFollowingTableWithRowSpanCrossingHeaderBoundary() {
        var style = TableStyles.Minimal();
        style.HeaderRowCount = 1;

        var exception = Assert.Throws<ArgumentException>(() =>
            PdfDoc.Create()
                .Paragraph(p => p.Text("KeepWithNextPrelude"), style: new PdfParagraphStyle {
                    KeepWithNext = true
                })
                .Table(new[] {
                    new[] { PdfTableCell.Merge("HeaderBody", rowSpan: 2), PdfTableCell.TextCell("H1") },
                    new[] { PdfTableCell.TextCell("B1") },
                    new[] { PdfTableCell.TextCell("B2"), PdfTableCell.TextCell("B3") }
                }, style: style)
                .ToBytes());

        Assert.Contains("Table cell row span cannot cross the table header boundary.", exception.Message, StringComparison.Ordinal);
    }

    [Fact]
    public void Table_RejectsOutOfRangeCellFillCoordinates() {
        var style = TableStyles.Minimal();
        style.CellFills = new Dictionary<(int Row, int Column), PdfColor> {
            [(2, 0)] = PdfColor.Gray
        };

        var exception = Assert.Throws<ArgumentException>(() =>
            PdfDoc.Create()
                .Table(new[] {
                    new[] { "A", "B" },
                    new[] { "1", "2" }
                }, style: style)
                .ToBytes());

        Assert.Contains("Table cell fill coordinates must fit inside the table grid.", exception.Message, StringComparison.Ordinal);
    }

    [Fact]
    public void RowColumnTable_RejectsOutOfRangeCellBorderCoordinates() {
        var style = TableStyles.Minimal();
        style.CellBorders = new Dictionary<(int Row, int Column), PdfCellBorder> {
            [(0, 2)] = new PdfCellBorder()
        };

        var exception = Assert.Throws<ArgumentException>(() =>
            PdfDoc.Create()
                .Compose(document =>
                    document.Page(page =>
                        page.Content(content =>
                            content.Row(row =>
                                row.Column(100, column =>
                                    column.Table(new[] {
                                        new[] { "A", "B" },
                                        new[] { "1", "2" }
                                    }, style: style))))))
                .ToBytes());

        Assert.Contains("Table cell border coordinates must fit inside the table grid.", exception.Message, StringComparison.Ordinal);
    }

    [Fact]
    public void ParagraphKeepWithNext_RejectsFollowingTableWithOutOfRangeCellFillCoordinates() {
        var style = TableStyles.Minimal();
        style.CellFills = new Dictionary<(int Row, int Column), PdfColor> {
            [(0, 2)] = PdfColor.Gray
        };

        var exception = Assert.Throws<ArgumentException>(() =>
            PdfDoc.Create()
                .Paragraph(p => p.Text("KeepWithNextPrelude"), style: new PdfParagraphStyle {
                    KeepWithNext = true
                })
                .Table(new[] {
                    new[] { "A", "B" },
                    new[] { "1", "2" }
                }, style: style)
                .ToBytes());

        Assert.Contains("Table cell fill coordinates must fit inside the table grid.", exception.Message, StringComparison.Ordinal);
    }

    [Fact]
    public void Table_RejectsOutOfRangeBodyColumnFill() {
        var style = TableStyles.Minimal();
        style.BodyColumnFills = new List<PdfColor?> {
            null,
            null,
            PdfColor.Gray
        };

        var exception = Assert.Throws<ArgumentException>(() =>
            PdfDoc.Create()
                .Table(new[] {
                    new[] { "A", "B" },
                    new[] { "1", "2" }
                }, style: style)
                .ToBytes());

        Assert.Contains("Table body column fills must fit inside the table grid.", exception.Message, StringComparison.Ordinal);
    }

    [Fact]
    public void Table_RejectsOutOfRangeColumnAlignment() {
        var style = TableStyles.Minimal();
        style.Alignments = new List<PdfColumnAlign> {
            PdfColumnAlign.Left,
            PdfColumnAlign.Right,
            PdfColumnAlign.Center
        };

        var exception = Assert.Throws<ArgumentException>(() =>
            PdfDoc.Create()
                .Table(new[] {
                    new[] { "A", "B" },
                    new[] { "1", "2" }
                }, style: style)
                .ToBytes());

        Assert.Contains("Table column alignments must fit inside the table grid.", exception.Message, StringComparison.Ordinal);
    }

    [Fact]
    public void RowColumnTable_RejectsOutOfRangeColumnAlignment() {
        var style = TableStyles.Minimal();
        style.Alignments = new List<PdfColumnAlign> {
            PdfColumnAlign.Left,
            PdfColumnAlign.Right,
            PdfColumnAlign.Center
        };

        var exception = Assert.Throws<ArgumentException>(() =>
            PdfDoc.Create()
                .Compose(document =>
                    document.Page(page =>
                        page.Content(content =>
                            content.Row(row =>
                                row.Column(100, column =>
                                    column.Table(new[] {
                                        new[] { "A", "B" },
                                        new[] { "1", "2" }
                                    }, style: style))))))
                .ToBytes());

        Assert.Contains("Table column alignments must fit inside the table grid.", exception.Message, StringComparison.Ordinal);
    }

    [Fact]
    public void ParagraphKeepWithNext_RejectsFollowingTableWithOutOfRangeColumnAlignment() {
        var style = TableStyles.Minimal();
        style.Alignments = new List<PdfColumnAlign> {
            PdfColumnAlign.Left,
            PdfColumnAlign.Right,
            PdfColumnAlign.Center
        };

        var exception = Assert.Throws<ArgumentException>(() =>
            PdfDoc.Create()
                .Paragraph(p => p.Text("KeepWithNextPrelude"), style: new PdfParagraphStyle {
                    KeepWithNext = true
                })
                .Table(new[] {
                    new[] { "A", "B" },
                    new[] { "1", "2" }
                }, style: style)
                .ToBytes());

        Assert.Contains("Table column alignments must fit inside the table grid.", exception.Message, StringComparison.Ordinal);
    }

    [Fact]
    public void Table_RejectsOutOfRangeVerticalAlignment() {
        var style = TableStyles.Minimal();
        style.VerticalAlignments = new List<PdfCellVerticalAlign> {
            PdfCellVerticalAlign.Top,
            PdfCellVerticalAlign.Middle,
            PdfCellVerticalAlign.Bottom
        };

        var exception = Assert.Throws<ArgumentException>(() =>
            PdfDoc.Create()
                .Table(new[] {
                    new[] { "A", "B" },
                    new[] { "1", "2" }
                }, style: style)
                .ToBytes());

        Assert.Contains("Table vertical alignments must fit inside the table grid.", exception.Message, StringComparison.Ordinal);
    }

    [Fact]
    public void RowColumnTable_RejectsOutOfRangeVerticalAlignment() {
        var style = TableStyles.Minimal();
        style.VerticalAlignments = new List<PdfCellVerticalAlign> {
            PdfCellVerticalAlign.Top,
            PdfCellVerticalAlign.Middle,
            PdfCellVerticalAlign.Bottom
        };

        var exception = Assert.Throws<ArgumentException>(() =>
            PdfDoc.Create()
                .Compose(document =>
                    document.Page(page =>
                        page.Content(content =>
                            content.Row(row =>
                                row.Column(100, column =>
                                    column.Table(new[] {
                                        new[] { "A", "B" },
                                        new[] { "1", "2" }
                                    }, style: style))))))
                .ToBytes());

        Assert.Contains("Table vertical alignments must fit inside the table grid.", exception.Message, StringComparison.Ordinal);
    }

    [Fact]
    public void ParagraphKeepWithNext_RejectsFollowingTableWithOutOfRangeVerticalAlignment() {
        var style = TableStyles.Minimal();
        style.VerticalAlignments = new List<PdfCellVerticalAlign> {
            PdfCellVerticalAlign.Top,
            PdfCellVerticalAlign.Middle,
            PdfCellVerticalAlign.Bottom
        };

        var exception = Assert.Throws<ArgumentException>(() =>
            PdfDoc.Create()
                .Paragraph(p => p.Text("KeepWithNextPrelude"), style: new PdfParagraphStyle {
                    KeepWithNext = true
                })
                .Table(new[] {
                    new[] { "A", "B" },
                    new[] { "1", "2" }
                }, style: style)
                .ToBytes());

        Assert.Contains("Table vertical alignments must fit inside the table grid.", exception.Message, StringComparison.Ordinal);
    }

    [Fact]
    public void RowColumnTable_RejectsOutOfRangeColumnWidthWeight() {
        var style = TableStyles.Minimal();
        style.ColumnWidthWeights = new List<double> {
            1,
            1,
            1
        };

        var exception = Assert.Throws<ArgumentException>(() =>
            PdfDoc.Create()
                .Compose(document =>
                    document.Page(page =>
                        page.Content(content =>
                            content.Row(row =>
                                row.Column(100, column =>
                                    column.Table(new[] {
                                        new[] { "A", "B" },
                                        new[] { "1", "2" }
                                    }, style: style))))))
                .ToBytes());

        Assert.Contains("Table column width weights must fit inside the table grid.", exception.Message, StringComparison.Ordinal);
    }

    [Fact]
    public void ParagraphKeepWithNext_RejectsFollowingTableWithOutOfRangeFixedColumnWidth() {
        var style = TableStyles.Minimal();
        style.ColumnWidthPoints = new List<double?> {
            null,
            null,
            42
        };

        var exception = Assert.Throws<ArgumentException>(() =>
            PdfDoc.Create()
                .Paragraph(p => p.Text("KeepWithNextPrelude"), style: new PdfParagraphStyle {
                    KeepWithNext = true
                })
                .Table(new[] {
                    new[] { "A", "B" },
                    new[] { "1", "2" }
                }, style: style)
                .ToBytes());

        Assert.Contains("Table fixed column widths must fit inside the table grid.", exception.Message, StringComparison.Ordinal);
    }

    [Fact]
    public void TableCell_RejectsInvalidLinkUri() {
        var exception = Assert.Throws<ArgumentException>(() =>
            PdfTableCell.TextCell("Invalid", "not-a-uri"));

        Assert.Equal("linkUri", exception.ParamName);
        Assert.Contains("Parameter 'linkUri' must be an absolute URI.", exception.Message, StringComparison.Ordinal);
    }

    [Fact]
    public void Table_RejectsInvalidStylePrimitiveValues() {
        var invalidBorder = TableStyles.Minimal();

        var borderException = Assert.Throws<ArgumentException>(() =>
            invalidBorder.BorderWidth = double.NaN);

        Assert.Contains("Table border width must be a non-negative finite value.", borderException.Message, StringComparison.Ordinal);

        var invalidRowSeparator = TableStyles.Minimal();

        var rowSeparatorException = Assert.Throws<ArgumentException>(() =>
            invalidRowSeparator.RowSeparatorWidth = double.PositiveInfinity);

        Assert.Contains("Table row separator width must be a non-negative finite value.", rowSeparatorException.Message, StringComparison.Ordinal);

        var invalidHeaderSeparator = TableStyles.Minimal();

        var headerSeparatorException = Assert.Throws<ArgumentException>(() =>
            invalidHeaderSeparator.HeaderSeparatorWidth = -0.1);

        Assert.Contains("Table header separator width must be a non-negative finite value.", headerSeparatorException.Message, StringComparison.Ordinal);

        var invalidFooterSeparator = TableStyles.Minimal();

        var footerSeparatorException = Assert.Throws<ArgumentException>(() =>
            invalidFooterSeparator.FooterSeparatorWidth = -0.1);

        Assert.Contains("Table footer separator width must be a non-negative finite value.", footerSeparatorException.Message, StringComparison.Ordinal);

        var invalidMaxWidth = TableStyles.Minimal();

        var maxWidthException = Assert.Throws<ArgumentException>(() =>
            invalidMaxWidth.MaxWidth = 0);

        Assert.Contains("Table max width must be a positive finite value.", maxWidthException.Message, StringComparison.Ordinal);

        var invalidLeftIndent = TableStyles.Minimal();

        var leftIndentException = Assert.Throws<ArgumentException>(() =>
            invalidLeftIndent.LeftIndent = -1);

        Assert.Contains("Table left indent must be a non-negative finite value.", leftIndentException.Message, StringComparison.Ordinal);

        var invalidHorizontalPadding = TableStyles.Minimal();

        var horizontalPaddingException = Assert.Throws<ArgumentException>(() =>
            invalidHorizontalPadding.CellPaddingX = -1);

        Assert.Contains("Table horizontal cell padding must be a non-negative finite value.", horizontalPaddingException.Message, StringComparison.Ordinal);

        var invalidVerticalPadding = TableStyles.Minimal();

        var verticalPaddingException = Assert.Throws<ArgumentException>(() =>
            invalidVerticalPadding.CellPaddingY = double.PositiveInfinity);

        Assert.Contains("Table vertical cell padding must be a non-negative finite value.", verticalPaddingException.Message, StringComparison.Ordinal);

        var invalidLeftPadding = TableStyles.Minimal();

        var leftPaddingException = Assert.Throws<ArgumentException>(() =>
            invalidLeftPadding.CellPaddingLeft = -1);

        Assert.Contains("Table left cell padding must be a non-negative finite value.", leftPaddingException.Message, StringComparison.Ordinal);

        var invalidRightPadding = TableStyles.Minimal();

        var rightPaddingException = Assert.Throws<ArgumentException>(() =>
            invalidRightPadding.CellPaddingRight = double.NaN);

        Assert.Contains("Table right cell padding must be a non-negative finite value.", rightPaddingException.Message, StringComparison.Ordinal);

        var invalidTopPadding = TableStyles.Minimal();

        var topPaddingException = Assert.Throws<ArgumentException>(() =>
            invalidTopPadding.CellPaddingTop = double.PositiveInfinity);

        Assert.Contains("Table top cell padding must be a non-negative finite value.", topPaddingException.Message, StringComparison.Ordinal);

        var invalidBottomPadding = TableStyles.Minimal();

        var bottomPaddingException = Assert.Throws<ArgumentException>(() =>
            invalidBottomPadding.CellPaddingBottom = -0.1);

        Assert.Contains("Table bottom cell padding must be a non-negative finite value.", bottomPaddingException.Message, StringComparison.Ordinal);

        var excessiveHorizontalPadding = TableStyles.Minimal();
        excessiveHorizontalPadding.ColumnWidthPoints = new List<double?> { 12 };
        excessiveHorizontalPadding.CellPaddingX = 6;

        var textWidthException = Assert.Throws<ArgumentException>(() =>
            PdfDoc.Create()
                .Table(new[] {
                    new[] { "A" },
                    new[] { "1" }
                }, style: excessiveHorizontalPadding)
                .ToBytes());

        Assert.Contains("Table horizontal cell padding must leave a positive text width.", textWidthException.Message, StringComparison.Ordinal);

        var excessiveLeftIndent = TableStyles.Minimal();
        excessiveLeftIndent.LeftIndent = 400;

        var leftIndentWidthException = Assert.Throws<ArgumentException>(() =>
            PdfDoc.Create(new PdfOptions {
                    PageWidth = 180,
                    PageHeight = 180,
                    MarginLeft = 30,
                    MarginRight = 30,
                    MarginTop = 30,
                    MarginBottom = 30
                })
                .Table(new[] {
                    new[] { "Only" }
                }, style: excessiveLeftIndent)
                .ToBytes());

        Assert.Contains("Table left indent must leave a positive table width.", leftIndentWidthException.Message, StringComparison.Ordinal);

        var invalidBaselineOffset = TableStyles.Minimal();

        var baselineException = Assert.Throws<ArgumentException>(() =>
            invalidBaselineOffset.RowBaselineOffset = double.NaN);

        Assert.Contains("Table row baseline offset must be a finite value.", baselineException.Message, StringComparison.Ordinal);

        var invalidCellFill = TableStyles.Minimal();

        var fillException = Assert.Throws<ArgumentException>(() =>
            invalidCellFill.CellFills = new Dictionary<(int Row, int Column), PdfColor> {
                [(-1, 0)] = PdfColor.Gray
            });

        Assert.Contains("Table cell fill coordinates cannot be negative.", fillException.Message, StringComparison.Ordinal);
    }

    [Fact]
    public void Table_RejectsInvalidAlignmentEnumValues() {
        var invalidTableAlignException = Assert.Throws<ArgumentException>(() =>
            PdfDoc.Create()
                .Table(new[] {
                    new[] { "A", "B" },
                    new[] { "1", "2" }
                }, align: (PdfAlign)99)
                .ToBytes());

        Assert.Contains("Table alignment must be Left, Center, or Right.", invalidTableAlignException.Message, StringComparison.Ordinal);

        var unsupportedTableAlignException = Assert.Throws<ArgumentException>(() =>
            PdfDoc.Create()
                .Table(new[] {
                    new[] { "A", "B" },
                    new[] { "1", "2" }
                }, align: PdfAlign.Justify)
                .ToBytes());

        Assert.Contains("Table alignment must be Left, Center, or Right.", unsupportedTableAlignException.Message, StringComparison.Ordinal);

        var composeTableAlignException = Assert.Throws<ArgumentException>(() =>
            PdfDoc.Create()
                .Compose(compose =>
                    compose.Page(page =>
                        page.Content(content =>
                            content.Column(column =>
                                column.Item().Table(new[] {
                                    new[] { "A", "B" },
                                    new[] { "1", "2" }
                                }, align: (PdfAlign)99)))))
                .ToBytes());

        Assert.Contains("Table alignment must be Left, Center, or Right.", composeTableAlignException.Message, StringComparison.Ordinal);

        var tableWithLinksAlignException = Assert.Throws<ArgumentException>(() =>
            PdfDoc.Create()
                .TableWithLinks(
                    new[] {
                        new[] { "A", "B" },
                        new[] { "1", "2" }
                    },
                    new Dictionary<(int Row, int Col), string> {
                        [(1, 0)] = "https://evotec.xyz"
                    },
                    align: PdfAlign.Justify));

        Assert.Contains("Table alignment must be Left, Center, or Right.", tableWithLinksAlignException.Message, StringComparison.Ordinal);

        var invalidCaptionAlign = TableStyles.Minimal();
        invalidCaptionAlign.Caption = "Caption";

        var invalidCaptionAlignException = Assert.Throws<ArgumentException>(() =>
            invalidCaptionAlign.CaptionAlign = (PdfAlign)99);

        Assert.Contains("Table caption alignment must be Left, Center, or Right.", invalidCaptionAlignException.Message, StringComparison.Ordinal);

        var unsupportedCaptionAlign = TableStyles.Minimal();
        unsupportedCaptionAlign.Caption = "Caption";

        var unsupportedCaptionAlignException = Assert.Throws<ArgumentException>(() =>
            unsupportedCaptionAlign.CaptionAlign = PdfAlign.Justify);

        Assert.Contains("Table caption alignment must be Left, Center, or Right.", unsupportedCaptionAlignException.Message, StringComparison.Ordinal);

        var invalidColumnAlign = TableStyles.Minimal();

        var invalidColumnAlignException = Assert.Throws<ArgumentException>(() =>
            invalidColumnAlign.Alignments = new List<PdfColumnAlign> { (PdfColumnAlign)99 });

        Assert.Contains("Table column alignments must be Left, Center, or Right.", invalidColumnAlignException.Message, StringComparison.Ordinal);

        var invalidVerticalAlign = TableStyles.Minimal();

        var invalidVerticalAlignException = Assert.Throws<ArgumentException>(() =>
            invalidVerticalAlign.VerticalAlignments = new List<PdfCellVerticalAlign> { (PdfCellVerticalAlign)99 });

        Assert.Contains("Table vertical alignments must be defined PDF cell vertical alignment values.", invalidVerticalAlignException.Message, StringComparison.Ordinal);
    }

    [Fact]
    public void TableStyle_AlignmentListsSnapshotAssignedCollections() {
        var horizontal = new List<PdfColumnAlign> { PdfColumnAlign.Left, PdfColumnAlign.Center };
        var vertical = new List<PdfCellVerticalAlign> { PdfCellVerticalAlign.Top, PdfCellVerticalAlign.Middle };
        var style = TableStyles.Minimal();

        style.Alignments = horizontal;
        style.VerticalAlignments = vertical;

        horizontal[0] = PdfColumnAlign.Right;
        vertical[0] = PdfCellVerticalAlign.Bottom;

        Assert.NotNull(style.Alignments);
        Assert.NotNull(style.VerticalAlignments);
        Assert.Equal(PdfColumnAlign.Left, style.Alignments![0]);
        Assert.Equal(PdfCellVerticalAlign.Top, style.VerticalAlignments![0]);
    }

    [Fact]
    public void TableStyle_ColumnSizingListsSnapshotAssignedCollections() {
        var fixedWidths = new List<double?> { 60, null };
        var minWidths = new List<double?> { 40, null };
        var maxWidths = new List<double?> { null, 120 };
        var weights = new List<double> { 1, 2 };
        var style = TableStyles.Minimal();

        style.ColumnWidthPoints = fixedWidths;
        style.ColumnMinWidthPoints = minWidths;
        style.ColumnMaxWidthPoints = maxWidths;
        style.ColumnWidthWeights = weights;

        fixedWidths[0] = 10;
        minWidths[0] = 10;
        maxWidths[1] = 10;
        weights[1] = 10;

        Assert.NotNull(style.ColumnWidthPoints);
        Assert.NotNull(style.ColumnMinWidthPoints);
        Assert.NotNull(style.ColumnMaxWidthPoints);
        Assert.NotNull(style.ColumnWidthWeights);
        Assert.Equal(60, style.ColumnWidthPoints![0]);
        Assert.Equal(40, style.ColumnMinWidthPoints![0]);
        Assert.Equal(120, style.ColumnMaxWidthPoints![1]);
        Assert.Equal(2, style.ColumnWidthWeights![1]);
    }

    [Fact]
    public void TableStyle_FillAndBorderCollectionsSnapshotAssignedValues() {
        var columnFills = new List<PdfColor?> { PdfColor.Gray, PdfColor.LightGray };
        var cellFills = new Dictionary<(int Row, int Column), PdfColor> {
            [(1, 1)] = new PdfColor(0.1, 0.2, 0.3)
        };
        var cellBorder = new PdfCellBorder {
            Color = new PdfColor(0.4, 0.5, 0.6),
            Width = 1.25,
            Left = false
        };
        var cellBorders = new Dictionary<(int Row, int Column), PdfCellBorder> {
            [(1, 1)] = cellBorder
        };
        var style = TableStyles.Minimal();

        style.BodyColumnFills = columnFills;
        style.CellFills = cellFills;
        style.CellBorders = cellBorders;

        columnFills[0] = PdfColor.White;
        cellFills[(1, 1)] = PdfColor.Black;
        cellBorder.Width = 4;
        cellBorder.Left = true;

        Assert.NotNull(style.BodyColumnFills);
        Assert.NotNull(style.CellFills);
        Assert.NotNull(style.CellBorders);
        Assert.Equal(PdfColor.Gray, style.BodyColumnFills![0]);
        Assert.Equal(new PdfColor(0.1, 0.2, 0.3), style.CellFills![(1, 1)]);
        Assert.Equal(1.25, style.CellBorders![(1, 1)].Width);
        Assert.False(style.CellBorders![(1, 1)].Left);
    }

    [Fact]
    public void TableStyle_TypographySettingsSurviveClone() {
        var style = TableStyles.Minimal();
        style.FontSize = 8;
        style.LineHeight = 1.6;
        style.HeaderFontSize = 12;
        style.FooterFontSize = 10;
        style.HeaderBold = false;
        style.FooterBold = false;
        style.KeepTogether = true;
        style.KeepWithNext = true;
        style.AllowRowBreakAcrossPages = false;
        style.MaxWidth = 180;
        style.LeftIndent = 24;
        style.CellPaddingLeft = 7;
        style.CellPaddingRight = 8;
        style.CellPaddingTop = 9;
        style.CellPaddingBottom = 10;

        PdfTableStyle clone = style.Clone();

        Assert.Equal(8, clone.FontSize);
        Assert.Equal(1.6, clone.LineHeight);
        Assert.Equal(12, clone.HeaderFontSize);
        Assert.Equal(10, clone.FooterFontSize);
        Assert.False(clone.HeaderBold);
        Assert.False(clone.FooterBold);
        Assert.True(clone.KeepTogether);
        Assert.True(clone.KeepWithNext);
        Assert.False(clone.AllowRowBreakAcrossPages);
        Assert.Equal(180, clone.MaxWidth);
        Assert.Equal(24, clone.LeftIndent);
        Assert.Equal(7, clone.CellPaddingLeft);
        Assert.Equal(8, clone.CellPaddingRight);
        Assert.Equal(9, clone.CellPaddingTop);
        Assert.Equal(10, clone.CellPaddingBottom);
    }

    private static byte[] CreateTablePaddingProbe(PdfOptions options, bool useRowColumnFlow, bool useSidePadding) {
        var style = TableStyles.Minimal();
        style.HeaderRowCount = 0;
        style.CellPaddingX = 0;
        style.CellPaddingY = 0;
        style.ColumnWidthPoints = new List<double?> { 90 };
        if (useSidePadding) {
            style.CellPaddingLeft = 18;
            style.CellPaddingRight = 3;
            style.CellPaddingTop = 16;
            style.CellPaddingBottom = 4;
        }

        var rows = new[] {
            new[] { "PadMarker" }
        };

        if (useRowColumnFlow) {
            return PdfDoc.Create(options)
                .Compose(compose =>
                    compose.Page(page =>
                        page.Content(content =>
                            content.Row(row =>
                                row.Column(100, column => column.Table(rows, style: style))))))
                .ToBytes();
        }

        return PdfDoc.Create(options)
            .Table(rows, style: style)
            .ToBytes();
    }

    private static byte[] CreateTableLineHeightProbe(PdfOptions options, double? lineHeight, bool useRowColumnFlow) {
        var style = TableStyles.Minimal();
        style.HeaderRowCount = 0;
        style.FontSize = 10;
        style.LineHeight = lineHeight;
        style.ColumnWidthPoints = new List<double?> { 72 };

        var rows = new[] {
            new[] { "FirstLine SecondLine" }
        };

        if (useRowColumnFlow) {
            return PdfDoc.Create(options)
                .Compose(compose =>
                    compose.Page(page =>
                        page.Content(content =>
                            content.Row(row =>
                                row.Column(100, column => column.Table(rows, style: style))))))
                .ToBytes();
        }

        return PdfDoc.Create(options)
            .Table(rows, style: style)
            .ToBytes();
    }

    private static byte[] CreateTableSpacingProbe(PdfOptions options, double spacingBefore, double spacingAfter) {
        var style = TableStyles.Minimal();
        style.HeaderRowCount = 0;
        style.SpacingBefore = spacingBefore;
        style.SpacingAfter = spacingAfter;

        return PdfDoc.Create(options)
            .Paragraph(p => p.Text("BeforeMarker"))
            .Table(new[] {
                new[] { "Alpha", "Ready" },
                new[] { "Beta", "Ready" }
            }, style: style)
            .Paragraph(p => p.Text("AfterMarker"))
            .ToBytes();
    }

    private static string RenderTableStyleContent(PdfTableStyle style) {
        byte[] bytes = PdfDoc.Create(new PdfOptions {
                PageWidth = 320,
                PageHeight = 220,
                MarginLeft = 30,
                MarginRight = 30,
                MarginTop = 30,
                MarginBottom = 30
            })
            .Table(new[] {
                new[] { "Metric", "Status" },
                new[] { "Queue", "Healthy" },
                new[] { "Latency", "Warning" }
            }, style: style)
            .ToBytes();

        return Encoding.ASCII.GetString(bytes);
    }

    private static int CountOccurrences(string text, string value) {
        int count = 0;
        int startIndex = 0;
        while (true) {
            int index = text.IndexOf(value, startIndex, StringComparison.Ordinal);
            if (index < 0) {
                return count;
            }

            count++;
            startIndex = index + value.Length;
        }
    }

    private static List<(double X1, double Y1, double X2, double Y2)> ExtractLinkRectangles(string pdf) {
        var rectangles = new List<(double X1, double Y1, double X2, double Y2)>();
        var matches = Regex.Matches(pdf, @"/Subtype /Link\b.*?/Rect \[(?<x1>-?\d+(?:\.\d+)?) (?<y1>-?\d+(?:\.\d+)?) (?<x2>-?\d+(?:\.\d+)?) (?<y2>-?\d+(?:\.\d+)?)\]");
        foreach (Match match in matches) {
            rectangles.Add((
                double.Parse(match.Groups["x1"].Value, CultureInfo.InvariantCulture),
                double.Parse(match.Groups["y1"].Value, CultureInfo.InvariantCulture),
                double.Parse(match.Groups["x2"].Value, CultureInfo.InvariantCulture),
                double.Parse(match.Groups["y2"].Value, CultureInfo.InvariantCulture)));
        }

        return rectangles;
    }

    private static List<(double X, double Y, double W, double H)> ExtractPaintedRectangles(string content, string colorOperator, string paintOperator) {
        var rectangles = new List<(double X, double Y, double W, double H)>();
        string pattern = Regex.Escape(colorOperator) +
            @"[\s\S]*?(?<x>-?\d+(?:\.\d+)?) (?<y>-?\d+(?:\.\d+)?) (?<w>-?\d+(?:\.\d+)?) (?<h>-?\d+(?:\.\d+)?) re\s+" +
            Regex.Escape(paintOperator);
        var matches = Regex.Matches(content, pattern);
        foreach (Match match in matches) {
            rectangles.Add((
                double.Parse(match.Groups["x"].Value, CultureInfo.InvariantCulture),
                double.Parse(match.Groups["y"].Value, CultureInfo.InvariantCulture),
                double.Parse(match.Groups["w"].Value, CultureInfo.InvariantCulture),
                double.Parse(match.Groups["h"].Value, CultureInfo.InvariantCulture)));
        }

        return rectangles;
    }

    private static List<(double X1, double Y1, double X2, double Y2)> ExtractStrokedLineSegments(string content, string colorOperator) {
        var segments = new List<(double X1, double Y1, double X2, double Y2)>();
        const string number = @"-?\d+(?:\.\d+)?";
        string pattern = Regex.Escape(colorOperator) +
            @"\s+(?:" + number + @" w\s+)?" +
            @"(?<x1>" + number + @") (?<y1>" + number + @") m\s+" +
            @"(?<x2>" + number + @") (?<y2>" + number + @") l\s+S";
        var matches = Regex.Matches(content, pattern);
        foreach (Match match in matches) {
            segments.Add((
                double.Parse(match.Groups["x1"].Value, CultureInfo.InvariantCulture),
                double.Parse(match.Groups["y1"].Value, CultureInfo.InvariantCulture),
                double.Parse(match.Groups["x2"].Value, CultureInfo.InvariantCulture),
                double.Parse(match.Groups["y2"].Value, CultureInfo.InvariantCulture)));
        }

        return segments;
    }

    private static byte[] CreateParagraphSpacingProbe(PdfOptions options, PdfParagraphStyle? style) {
        return PdfDoc.Create(options)
            .Paragraph(p => p.Text("BeforeMarker"))
            .Paragraph(p => p.Text("TargetMarker"), style: style)
            .Paragraph(p => p.Text("AfterMarker"))
            .ToBytes();
    }

    private static byte[] CreatePanelSpacingProbe(PdfOptions options, PanelStyle style) {
        return PdfDoc.Create(options)
            .Paragraph(p => p.Text("BeforeMarker"), style: new PdfParagraphStyle {
                SpacingAfter = 0
            })
            .PanelParagraph(p => p.Text("PanelMarker"), style)
            .Paragraph(p => p.Text("AfterMarker"))
            .ToBytes();
    }

    private static byte[] CreateParagraphLineHeightProbe(PdfOptions options, PdfParagraphStyle? style) {
        return PdfDoc.Create(options)
            .Paragraph(p => p
                .Text("FirstLine")
                .LineBreak()
                .Text("SecondLine")
                .LineBreak()
                .Text("ThirdLine"), style: style)
            .ToBytes();
    }

    private static byte[] CreateParagraphIndentProbe(PdfOptions options, PdfParagraphStyle? style) {
        return PdfDoc.Create(options)
            .Paragraph(p => p.Text("IndentedMarker alpha beta gamma delta epsilon zeta eta theta iota kappa lambda mu nu xi omicron pi rho sigma tau."), style: style)
            .ToBytes();
    }

    private static byte[] CreateTopLevelFlowSpacingBeforeProbe(string blockKind, double spacingBefore) {
        var options = CreateFlowSpacingProbeOptions();
        var paragraphStyle = new PdfParagraphStyle { SpacingBefore = 0, SpacingAfter = 0 };
        var doc = PdfDoc.Create(options);

        switch (blockKind) {
            case "bullet-list":
                doc.Bullets(new[] { "ListTopMarker" }, style: new PdfListStyle { SpacingBefore = spacingBefore, SpacingAfter = 0, ItemSpacing = 0 });
                break;
            case "numbered-list":
                doc.Numbered(new[] { "ListTopMarker" }, style: new PdfListStyle { SpacingBefore = spacingBefore, SpacingAfter = 0, ItemSpacing = 0 });
                break;
            case "panel":
                doc.PanelParagraph(p => p.Text("PanelTopMarker"), new PanelStyle { SpacingBefore = spacingBefore, SpacingAfter = 0, PaddingX = 4, PaddingY = 4 });
                break;
            case "horizontal-rule":
                doc.HR(style: new PdfHorizontalRuleStyle { Thickness = 2, SpacingBefore = spacingBefore, SpacingAfter = 0 })
                    .Paragraph(p => p.Text("AfterFixedMarker"), style: paragraphStyle);
                break;
            case "image":
                doc.Image(CreateMinimalRgbPng(), 24, 12, style: new PdfImageStyle { SpacingBefore = spacingBefore, SpacingAfter = 0 })
                    .Paragraph(p => p.Text("AfterFixedMarker"), style: paragraphStyle);
                break;
            case "shape":
                doc.Shape(OfficeShape.Rectangle(24, 12), style: new PdfDrawingStyle { SpacingBefore = spacingBefore, SpacingAfter = 0 })
                    .Paragraph(p => p.Text("AfterFixedMarker"), style: paragraphStyle);
                break;
            case "drawing":
                doc.Drawing(new OfficeDrawing(24, 12).AddShape(OfficeShape.Rectangle(24, 12), 0, 0), style: new PdfDrawingStyle { SpacingBefore = spacingBefore, SpacingAfter = 0 })
                    .Paragraph(p => p.Text("AfterFixedMarker"), style: paragraphStyle);
                break;
            case "row":
                doc.Compose(document => document.Page(page => page.Content(content => content.Row(row => row
                    .Style(new PdfRowStyle { SpacingBefore = spacingBefore, SpacingAfter = 0 })
                    .Column(100, column => column.Paragraph(p => p.Text("RowTopMarker"), style: paragraphStyle))))));
                break;
            default:
                throw new ArgumentOutOfRangeException(nameof(blockKind), blockKind, "Unknown flow block kind.");
        }

        return doc.ToBytes();
    }

    private static byte[] CreateColumnFlowSpacingBeforeProbe(string blockKind, double spacingBefore) {
        var options = CreateFlowSpacingProbeOptions();
        return PdfDoc.Create(options)
            .Compose(document => document.Page(page => page.Content(content => content.Row(row => row
                .Column(100, column => AddColumnFlowSpacingBeforeProbe(column, blockKind, spacingBefore))))))
            .ToBytes();
    }

    private static void AddColumnFlowSpacingBeforeProbe(PdfRowColumnCompose column, string blockKind, double spacingBefore) {
        var paragraphStyle = new PdfParagraphStyle { SpacingBefore = 0, SpacingAfter = 0 };

        switch (blockKind) {
            case "bullet-list":
                column.Bullets(new[] { "ColumnListMarker" }, style: new PdfListStyle { SpacingBefore = spacingBefore, SpacingAfter = 0, ItemSpacing = 0 });
                break;
            case "numbered-list":
                column.Numbered(new[] { "ColumnListMarker" }, style: new PdfListStyle { SpacingBefore = spacingBefore, SpacingAfter = 0, ItemSpacing = 0 });
                break;
            case "panel":
                column.PanelParagraph(p => p.Text("ColumnPanelMarker"), new PanelStyle { SpacingBefore = spacingBefore, SpacingAfter = 0, PaddingX = 4, PaddingY = 4 });
                break;
            case "horizontal-rule":
                column.HR(style: new PdfHorizontalRuleStyle { Thickness = 2, SpacingBefore = spacingBefore, SpacingAfter = 0 })
                    .Paragraph(p => p.Text("ColumnAfterFixedMarker"), style: paragraphStyle);
                break;
            case "image":
                column.Image(CreateMinimalRgbPng(), 24, 12, style: new PdfImageStyle { SpacingBefore = spacingBefore, SpacingAfter = 0 })
                    .Paragraph(p => p.Text("ColumnAfterFixedMarker"), style: paragraphStyle);
                break;
            case "shape":
                column.Shape(OfficeShape.Rectangle(24, 12), style: new PdfDrawingStyle { SpacingBefore = spacingBefore, SpacingAfter = 0 })
                    .Paragraph(p => p.Text("ColumnAfterFixedMarker"), style: paragraphStyle);
                break;
            case "drawing":
                column.Drawing(new OfficeDrawing(24, 12).AddShape(OfficeShape.Rectangle(24, 12), 0, 0), style: new PdfDrawingStyle { SpacingBefore = spacingBefore, SpacingAfter = 0 })
                    .Paragraph(p => p.Text("ColumnAfterFixedMarker"), style: paragraphStyle);
                break;
            default:
                throw new ArgumentOutOfRangeException(nameof(blockKind), blockKind, "Unknown flow block kind.");
        }
    }

    private static PdfOptions CreateFlowSpacingProbeOptions() {
        return new PdfOptions {
            PageWidth = 320,
            PageHeight = 220,
            MarginLeft = 30,
            MarginRight = 30,
            MarginTop = 30,
            MarginBottom = 30,
            DefaultFont = PdfStandardFont.Helvetica,
            DefaultFontSize = 10
        };
    }

    private static OfficeDrawing CreateKeepWithNextDrawingScene() {
        var shape = OfficeShape.Rectangle(24, 24);
        shape.FillColor = OfficeColor.WhiteSmoke;

        return new OfficeDrawing(24, 24)
            .AddShape(shape, 0, 0);
    }

    private static byte[] CreateMinimalRgbPng() {
        return new byte[] {
            137, 80, 78, 71, 13, 10, 26, 10,
            0, 0, 0, 13,
            73, 72, 68, 82,
            0, 0, 0, 1,
            0, 0, 0, 1,
            8, 2, 0, 0, 0,
            0, 0, 0, 0,
            0, 0, 0, 12,
            73, 68, 65, 84,
            0x78, 0x9C, 0x63, 0xF8, 0xCF, 0xC0, 0x00, 0x00, 0x03, 0x01, 0x01, 0x00,
            0, 0, 0, 0,
            0, 0, 0, 0,
            73, 69, 78, 68,
            0, 0, 0, 0
        };
    }

    private static double FindWordStartX(UglyToad.PdfPig.Content.Page page, string word) {
        var lines = page.Letters
            .Where(letter => !string.IsNullOrWhiteSpace(letter.Value))
            .GroupBy(letter => Math.Round(letter.StartBaseLine.Y, 1));

        foreach (var line in lines) {
            var ordered = line.OrderBy(letter => letter.StartBaseLine.X).ToList();
            string text = string.Concat(ordered.Select(letter => letter.Value));
            int index = text.IndexOf(word, StringComparison.Ordinal);
            if (index >= 0) {
                return ordered[index].StartBaseLine.X;
            }
        }

        throw new InvalidOperationException("Could not find word '" + word + "' in rendered PDF text.");
    }

    private static double FindWordEndX(UglyToad.PdfPig.Content.Page page, string word) {
        var lines = page.Letters
            .Where(letter => !string.IsNullOrWhiteSpace(letter.Value))
            .GroupBy(letter => Math.Round(letter.StartBaseLine.Y, 1));

        foreach (var line in lines) {
            var ordered = line.OrderBy(letter => letter.StartBaseLine.X).ToList();
            string text = string.Concat(ordered.Select(letter => letter.Value));
            int index = text.IndexOf(word, StringComparison.Ordinal);
            if (index >= 0) {
                return ordered[index + word.Length - 1].EndBaseLine.X;
            }
        }

        throw new InvalidOperationException("Could not find word '" + word + "' in rendered PDF text.");
    }

    private static double FindWordStartXOnBaseline(UglyToad.PdfPig.Content.Page page, string word, double baselineY) {
        var ordered = GetLettersOnBaseline(page, baselineY);
        string text = string.Concat(ordered.Select(letter => letter.Value));
        int index = text.IndexOf(word, StringComparison.Ordinal);
        if (index >= 0) {
            return ordered[index].StartBaseLine.X;
        }

        throw new InvalidOperationException("Could not find word '" + word + "' on baseline " + baselineY.ToString("0.##", CultureInfo.InvariantCulture) + ".");
    }

    private static double FindWordEndXOnBaseline(UglyToad.PdfPig.Content.Page page, string word, double baselineY) {
        var ordered = GetLettersOnBaseline(page, baselineY);
        string text = string.Concat(ordered.Select(letter => letter.Value));
        int index = text.IndexOf(word, StringComparison.Ordinal);
        if (index >= 0) {
            return ordered[index + word.Length - 1].EndBaseLine.X;
        }

        throw new InvalidOperationException("Could not find word '" + word + "' on baseline " + baselineY.ToString("0.##", CultureInfo.InvariantCulture) + ".");
    }

    private static List<UglyToad.PdfPig.Content.Letter> GetLettersOnBaseline(UglyToad.PdfPig.Content.Page page, double baselineY) {
        double roundedBaselineY = Math.Round(baselineY, 1);
        return page.Letters
            .Where(letter =>
                !string.IsNullOrWhiteSpace(letter.Value) &&
                Math.Round(letter.StartBaseLine.Y, 1).Equals(roundedBaselineY))
            .OrderBy(letter => letter.StartBaseLine.X)
            .ToList();
    }

    private static double FindWordStartY(UglyToad.PdfPig.Content.Page page, string word) {
        var lines = page.Letters
            .Where(letter => !string.IsNullOrWhiteSpace(letter.Value))
            .GroupBy(letter => Math.Round(letter.StartBaseLine.Y, 1));

        foreach (var line in lines) {
            var ordered = line.OrderBy(letter => letter.StartBaseLine.X).ToList();
            string text = string.Concat(ordered.Select(letter => letter.Value));
            int index = text.IndexOf(word, StringComparison.Ordinal);
            if (index >= 0) {
                return ordered[index].StartBaseLine.Y;
            }
        }

        throw new InvalidOperationException("Could not find word '" + word + "' in rendered PDF text.");
    }

    private static double AverageBaselineY(IReadOnlyList<UglyToad.PdfPig.Content.Letter> letters, string word) {
        string text = string.Concat(letters.Select(letter => letter.Value));
        int index = text.IndexOf(word, StringComparison.Ordinal);
        if (index < 0) {
            throw new InvalidOperationException("Could not find word '" + word + "' in rendered PDF text.");
        }

        return letters
            .Skip(index)
            .Take(word.Length)
            .Average(letter => letter.StartBaseLine.Y);
    }

    private static double FindWordPointSize(UglyToad.PdfPig.Content.Page page, string word) {
        var lines = page.Letters
            .Where(letter => !string.IsNullOrWhiteSpace(letter.Value))
            .GroupBy(letter => Math.Round(letter.StartBaseLine.Y, 1));

        foreach (var line in lines) {
            var ordered = line.OrderBy(letter => letter.StartBaseLine.X).ToList();
            string text = string.Concat(ordered.Select(letter => letter.Value));
            int index = text.IndexOf(word, StringComparison.Ordinal);
            if (index >= 0) {
                return ordered[index].PointSize;
            }
        }

        throw new InvalidOperationException("Could not find word '" + word + "' in rendered PDF text.");
    }

    private static int CountTextLines(UglyToad.PdfPig.Content.Page page) {
        return page.Letters
            .Where(letter => !string.IsNullOrWhiteSpace(letter.Value))
            .GroupBy(letter => Math.Round(letter.StartBaseLine.Y, 1))
            .Count();
    }

    private static List<List<UglyToad.PdfPig.Content.Letter>> GetNonWhitespaceLetterLines(UglyToad.PdfPig.Content.Page page) {
        return page.Letters
            .Where(letter => !string.IsNullOrWhiteSpace(letter.Value))
            .GroupBy(letter => Math.Round(letter.StartBaseLine.Y, 1))
            .OrderByDescending(group => group.Key)
            .Select(group => group.OrderBy(letter => letter.StartBaseLine.X).ToList())
            .ToList();
    }

    private static List<VisualTextLine> GetVisualTextLines(UglyToad.PdfPig.Content.Page page, double minX, double maxX) {
        return page.Letters
            .Where(letter =>
                !string.IsNullOrWhiteSpace(letter.Value) &&
                letter.StartBaseLine.X >= minX &&
                letter.EndBaseLine.X <= maxX)
            .GroupBy(letter => Math.Round(letter.StartBaseLine.Y, 1))
            .OrderByDescending(group => group.Key)
            .Select(group => {
                var letters = group.OrderBy(letter => letter.StartBaseLine.X).ToList();
                return new VisualTextLine(
                    string.Concat(letters.Select(letter => letter.Value)),
                    letters.Min(letter => letter.StartBaseLine.X),
                    letters.Max(letter => letter.EndBaseLine.X),
                    letters[0].StartBaseLine.Y);
            })
            .ToList();
    }

    private static void AssertReadableTextRhythm(IReadOnlyList<VisualTextLine> lines, string areaName) {
        for (int i = 1; i < lines.Count; i++) {
            double baselineGap = lines[i - 1].BaselineY - lines[i].BaselineY;
            Assert.True(baselineGap >= 9.5,
                $"Expected {areaName} text to keep readable baseline rhythm between '{lines[i - 1].Text}' and '{lines[i].Text}'. Gap: {baselineGap:0.##}pt.");
        }
    }

    private static void AssertNoCrampedBaselines(IReadOnlyList<VisualTextLine> lines, string areaName) {
        for (int i = 1; i < lines.Count; i++) {
            double baselineGap = lines[i - 1].BaselineY - lines[i].BaselineY;
            Assert.True(baselineGap >= 7.5,
                $"Expected {areaName} to avoid cramped or overlapping baselines between '{lines[i - 1].Text}' and '{lines[i].Text}'. Gap: {baselineGap:0.##}pt.");
        }
    }

    private static void AssertNoSameBaselineTextCollisions(UglyToad.PdfPig.Content.Page page, string areaName) {
        foreach (var line in page.Letters
            .Where(letter => !string.IsNullOrWhiteSpace(letter.Value))
            .GroupBy(letter => Math.Round(letter.StartBaseLine.Y, 1))) {
            var ordered = line.OrderBy(letter => letter.StartBaseLine.X).ToList();
            for (int i = 1; i < ordered.Count; i++) {
                double gap = ordered[i].StartBaseLine.X - ordered[i - 1].EndBaseLine.X;
                Assert.True(gap >= -0.25,
                    $"Expected {areaName} text not to collide on baseline {line.Key:0.##}, but '{ordered[i - 1].Value}' and '{ordered[i].Value}' overlap by {-gap:0.##}pt.");
            }
        }
    }

    private static void AssertNoAmbiguousSameBaselineRunGaps(UglyToad.PdfPig.Content.Page page, string areaName) {
        const double sameRunGap = 0.75;
        const double minimumReadableRunGap = 1.8;

        foreach (var line in page.Letters
            .Where(letter => !string.IsNullOrWhiteSpace(letter.Value))
            .GroupBy(letter => Math.Round(letter.StartBaseLine.Y, 1))) {
            var ordered = line.OrderBy(letter => letter.StartBaseLine.X).ToList();
            for (int i = 1; i < ordered.Count; i++) {
                double gap = ordered[i].StartBaseLine.X - ordered[i - 1].EndBaseLine.X;
                if (gap <= sameRunGap) {
                    continue;
                }

                Assert.True(gap >= minimumReadableRunGap,
                    $"Expected {areaName} to avoid visually ambiguous same-baseline run gaps on baseline {line.Key:0.##} between '{ordered[i - 1].Value}' and '{ordered[i].Value}'. Gap: {gap:0.##}pt.");
            }
        }
    }

    private static List<double> GetInterWordGaps(List<UglyToad.PdfPig.Content.Letter> letters) {
        return letters
            .Zip(letters.Skip(1), (left, right) => right.StartBaseLine.X - left.EndBaseLine.X)
            .Where(gap => gap > 1)
            .ToList();
    }

    private static IReadOnlyList<string> GetPageContentStreams(byte[] pdf, int pageNumber) {
        var document = PdfReadDocument.Load(pdf);
        var (objects, _) = PdfSyntax.ParseObjects(pdf);
        int pageObjectNumber = document.Pages[pageNumber - 1].ObjectNumber;
        if (!objects.TryGetValue(pageObjectNumber, out var pageObject) || pageObject.Value is not PdfDictionary pageDictionary) {
            throw new InvalidOperationException("Page object was not found.");
        }

        if (!pageDictionary.Items.TryGetValue("Contents", out var contents)) {
            throw new InvalidOperationException("Page contents were not found.");
        }

        var streams = new List<string>();
        AppendContentStreams(objects, contents, streams);
        return streams;
    }

    private static void AppendContentStreams(Dictionary<int, PdfIndirectObject> objects, PdfObject contents, List<string> streams) {
        if (contents is PdfReference reference) {
            if (objects.TryGetValue(reference.ObjectNumber, out var indirect) && indirect.Value is PdfStream stream) {
                streams.Add(Encoding.GetEncoding("ISO-8859-1").GetString(stream.Data));
            }

            return;
        }

        if (contents is PdfArray array) {
            foreach (var item in array.Items) {
                AppendContentStreams(objects, item, streams);
            }
        }
    }

    private sealed class VisualTextLine {
        internal VisualTextLine(string text, double x1, double x2, double baselineY) {
            Text = text;
            X1 = x1;
            X2 = x2;
            BaselineY = baselineY;
        }

        internal string Text { get; }

        internal double X1 { get; }

        internal double X2 { get; }

        internal double BaselineY { get; }
    }
}
