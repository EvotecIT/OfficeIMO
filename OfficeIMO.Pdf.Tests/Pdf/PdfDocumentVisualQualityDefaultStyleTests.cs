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
    public void ListStyle_ClonePreservesPageFlowSettings() {
        var style = new PdfListStyle {
            MarkerFontFamily = "Premium Marker",
            KeepTogether = true,
            KeepWithNext = true
        };

        PdfListStyle clone = style.Clone();

        style.KeepTogether = false;
        style.KeepWithNext = false;
        style.MarkerFontFamily = "Changed Marker";

        Assert.Equal("Premium Marker", clone.MarkerFontFamily);
        Assert.True(clone.KeepTogether);
        Assert.True(clone.KeepWithNext);
    }

    [Fact]
    public void Options_SnapshotDefaultParagraphStyle() {
        var style = new PdfParagraphStyle {
            LineHeight = 1.7,
            LeftIndent = global::OfficeIMO.OfficeLength.Points(8),
            FirstLineIndent = global::OfficeIMO.OfficeLength.Points(4),
            SpacingAfter = global::OfficeIMO.OfficeLength.Points(9),
            DefaultTabStopWidth = global::OfficeIMO.OfficeLength.Points(72),
            WidowControl = true
        };
        var options = new PdfCreateOptions {
            DefaultParagraphStyle = style
        };

        style.LineHeight = 2.1;
        style.LeftIndent = global::OfficeIMO.OfficeLength.Points(20);
        style.FirstLineIndent = global::OfficeIMO.OfficeLength.Points(21);
        style.SpacingAfter = global::OfficeIMO.OfficeLength.Points(22);
        style.DefaultTabStopWidth = global::OfficeIMO.OfficeLength.Points(18);
        style.WidowControl = false;

        PdfParagraphStyle readback = options.DefaultParagraphStyle!;
        readback.LeftIndent = global::OfficeIMO.OfficeLength.Points(30);

        PdfCreateOptions clone = options.Clone();

        Assert.Equal(1.7, options.DefaultParagraphStyle!.LineHeight);
        Assert.Equal(8, options.DefaultParagraphStyle.LeftIndent.ToPoints());
        Assert.Equal(4, options.DefaultParagraphStyle.FirstLineIndent.ToPoints());
        Assert.Equal(9, (options.DefaultParagraphStyle.SpacingAfter?.ToPoints()));
        Assert.Equal(72, (options.DefaultParagraphStyle.DefaultTabStopWidth?.ToPoints()));
        Assert.True(options.DefaultParagraphStyle.WidowControl);
        Assert.Equal(8, clone.DefaultParagraphStyle!.LeftIndent.ToPoints());
        Assert.Equal(72, (clone.DefaultParagraphStyle.DefaultTabStopWidth?.ToPoints()));
    }

    [Fact]
    public void Options_SnapshotDefaultTableStyle() {
        var style = PdfTableStyles.Minimal();
        style.BorderWidth = global::OfficeIMO.OfficeLength.Points(0.8);
        style.CellPaddingX = global::OfficeIMO.OfficeLength.Points(12);
        style.RowSeparatorColor = global::OfficeIMO.OfficeColor.FromNormalizedRgb(0.11, 0.22, 0.33);
        style.RowSeparatorWidth = global::OfficeIMO.OfficeLength.Points(0.7);
        style.FooterSeparatorColor = global::OfficeIMO.OfficeColor.FromNormalizedRgb(0.21, 0.32, 0.43);
        style.FooterSeparatorWidth = global::OfficeIMO.OfficeLength.Points(0.9);
        style.Alignments = new List<PdfColumnAlign> { PdfColumnAlign.Right };
        style.MaxWidth = global::OfficeIMO.OfficeLength.Points(180);
        style.LeftIndent = global::OfficeIMO.OfficeLength.Points(24);
        style.KeepWithNext = true;
        var options = new PdfCreateOptions {
            DefaultTableStyle = style
        };

        style.BorderWidth = global::OfficeIMO.OfficeLength.Points(2);
        style.CellPaddingX = global::OfficeIMO.OfficeLength.Points(30);
        style.RowSeparatorColor = global::OfficeIMO.OfficeColor.Black;
        style.RowSeparatorWidth = global::OfficeIMO.OfficeLength.Points(1.5);
        style.FooterSeparatorColor = global::OfficeIMO.OfficeColor.White;
        style.FooterSeparatorWidth = global::OfficeIMO.OfficeLength.Points(2.1);
        style.Alignments[0] = PdfColumnAlign.Left;
        style.MaxWidth = global::OfficeIMO.OfficeLength.Points(260);
        style.LeftIndent = global::OfficeIMO.OfficeLength.Points(36);
        style.KeepWithNext = false;

        PdfTableStyle readback = options.DefaultTableStyle!;
        readback.CellPaddingX = global::OfficeIMO.OfficeLength.Points(44);
        readback.Alignments![0] = PdfColumnAlign.Center;

        PdfCreateOptions clone = options.Clone();

        Assert.True(options.HasExplicitDefaultTableStyle);
        Assert.Equal(0.8, options.DefaultTableStyle!.BorderWidth.ToPoints());
        Assert.Equal(12, options.DefaultTableStyle.CellPaddingX.ToPoints());
        Assert.Equal(global::OfficeIMO.OfficeColor.FromNormalizedRgb(0.11, 0.22, 0.33), options.DefaultTableStyle.RowSeparatorColor);
        Assert.Equal(0.7, options.DefaultTableStyle.RowSeparatorWidth.ToPoints());
        Assert.Equal(global::OfficeIMO.OfficeColor.FromNormalizedRgb(0.21, 0.32, 0.43), options.DefaultTableStyle.FooterSeparatorColor);
        Assert.Equal(0.9, options.DefaultTableStyle.FooterSeparatorWidth.ToPoints());
        Assert.Equal(PdfColumnAlign.Right, options.DefaultTableStyle.Alignments![0]);
        Assert.Equal(180, options.DefaultTableStyle.MaxWidth?.ToPoints());
        Assert.Equal(24, options.DefaultTableStyle.LeftIndent.ToPoints());
        Assert.True(options.DefaultTableStyle.KeepWithNext);
        Assert.Equal(12, clone.DefaultTableStyle!.CellPaddingX.ToPoints());
        Assert.Equal(global::OfficeIMO.OfficeColor.FromNormalizedRgb(0.21, 0.32, 0.43), clone.DefaultTableStyle.FooterSeparatorColor);
        Assert.Equal(0.9, clone.DefaultTableStyle.FooterSeparatorWidth.ToPoints());
        Assert.Equal(PdfColumnAlign.Right, clone.DefaultTableStyle.Alignments![0]);
        Assert.Equal(180, clone.DefaultTableStyle.MaxWidth?.ToPoints());
        Assert.Equal(24, clone.DefaultTableStyle.LeftIndent.ToPoints());
        Assert.True(clone.DefaultTableStyle.KeepWithNext);
        Assert.True(clone.HasExplicitDefaultTableStyle);
        Assert.False(new PdfCreateOptions().HasExplicitDefaultTableStyle);
        Assert.False(new PdfCreateOptions().Clone().HasExplicitDefaultTableStyle);
    }

    [Fact]
    public void Options_SnapshotDefaultHeadingStyles() {
        var style = new PdfHeadingStyle {
            FontSize = global::OfficeIMO.OfficeLength.Points(16),
            LineHeight = 1.1,
            SpacingBefore = global::OfficeIMO.OfficeLength.Points(4),
            SpacingAfter = global::OfficeIMO.OfficeLength.Points(12),
            Color = global::OfficeIMO.OfficeColor.FromRgb(10, 20, 30),
            Bold = false,
            ApplySpacingBeforeAtTop = true,
            KeepWithNext = false
        };
        var styles = new PdfHeadingStyles {
            Level1 = style
        };
        var options = new PdfCreateOptions {
            DefaultHeadingStyles = styles
        };

        style.FontSize = global::OfficeIMO.OfficeLength.Points(30);
        style.SpacingAfter = global::OfficeIMO.OfficeLength.Points(2);
        styles.Level1 = new PdfHeadingStyle {
            FontSize = global::OfficeIMO.OfficeLength.Points(9)
        };

        PdfHeadingStyles readback = options.DefaultHeadingStyles!;
        readback.Level1 = new PdfHeadingStyle {
            FontSize = global::OfficeIMO.OfficeLength.Points(8)
        };

        PdfCreateOptions clone = options.Clone();

        Assert.Equal(16, (options.DefaultHeadingStyles!.Level1!.FontSize?.ToPoints()));
        Assert.Equal(1.1, options.DefaultHeadingStyles.Level1.LineHeight);
        Assert.Equal(4, options.DefaultHeadingStyles.Level1.SpacingBefore.ToPoints());
        Assert.Equal(12, (options.DefaultHeadingStyles.Level1.SpacingAfter?.ToPoints()));
        Assert.Equal(global::OfficeIMO.OfficeColor.FromRgb(10, 20, 30), options.DefaultHeadingStyles.Level1.Color);
        Assert.False(options.DefaultHeadingStyles.Level1.Bold);
        Assert.True(options.DefaultHeadingStyles.Level1.ApplySpacingBeforeAtTop);
        Assert.False(options.DefaultHeadingStyles.Level1.KeepWithNext);
        Assert.Equal(16, (clone.DefaultHeadingStyles!.Level1!.FontSize?.ToPoints()));
        Assert.False(clone.DefaultHeadingStyles.Level1.Bold);
        Assert.True(clone.DefaultHeadingStyles.Level1.ApplySpacingBeforeAtTop);
    }

    [Fact]
    public void Options_SnapshotDefaultPanelStyle() {
        var style = new PdfPanelStyle {
            Background = global::OfficeIMO.OfficeColor.FromRgb(240, 248, 255),
            BorderColor = global::OfficeIMO.OfficeColor.FromRgb(20, 40, 60),
            BorderWidth = global::OfficeIMO.OfficeLength.Points(1.2),
            PaddingX = global::OfficeIMO.OfficeLength.Points(12),
            PaddingY = global::OfficeIMO.OfficeLength.Points(8),
            MaxWidth = global::OfficeIMO.OfficeLength.Points(180),
            Align = PdfAlign.Center,
            SpacingBefore = global::OfficeIMO.OfficeLength.Points(5),
            SpacingAfter = global::OfficeIMO.OfficeLength.Points(13),
            KeepTogether = true,
            KeepWithNext = true
        };
        var options = new PdfCreateOptions {
            DefaultPanelStyle = style
        };

        style.PaddingX = global::OfficeIMO.OfficeLength.Points(30);
        style.Align = PdfAlign.Right;
        style.Background = global::OfficeIMO.OfficeColor.Black;
        style.KeepWithNext = false;

        PdfPanelStyle readback = options.DefaultPanelStyle!;
        readback.PaddingX = global::OfficeIMO.OfficeLength.Points(44);

        PdfCreateOptions clone = options.Clone();

        Assert.Equal(global::OfficeIMO.OfficeColor.FromRgb(240, 248, 255), options.DefaultPanelStyle!.Background);
        Assert.Equal(global::OfficeIMO.OfficeColor.FromRgb(20, 40, 60), options.DefaultPanelStyle.BorderColor);
        Assert.Equal(1.2, options.DefaultPanelStyle.BorderWidth.ToPoints());
        Assert.Equal(12, options.DefaultPanelStyle.PaddingX.ToPoints());
        Assert.Equal(8, options.DefaultPanelStyle.PaddingY.ToPoints());
        Assert.Equal(180, (options.DefaultPanelStyle.MaxWidth?.ToPoints()));
        Assert.Equal(PdfAlign.Center, options.DefaultPanelStyle.Align);
        Assert.Equal(5, options.DefaultPanelStyle.SpacingBefore.ToPoints());
        Assert.Equal(13, options.DefaultPanelStyle.SpacingAfter.ToPoints());
        Assert.True(options.DefaultPanelStyle.KeepTogether);
        Assert.True(options.DefaultPanelStyle.KeepWithNext);
        Assert.Equal(12, clone.DefaultPanelStyle!.PaddingX.ToPoints());
        Assert.True(clone.DefaultPanelStyle.KeepWithNext);
    }

    [Fact]
    public void PanelParagraph_RendersAndSnapshotsSideSpecificPanelBorders() {
        global::OfficeIMO.OfficeColor red = global::OfficeIMO.OfficeColor.FromRgb(255, 0, 0);
        global::OfficeIMO.OfficeColor blue = global::OfficeIMO.OfficeColor.FromRgb(0, 0, 255);
        var style = new PdfPanelStyle {
            Background = global::OfficeIMO.OfficeColor.FromRgb(245, 245, 245),
            TopBorder = new PdfPanelBorder {
                Color = red,
                Width = global::OfficeIMO.OfficeLength.Points(2)
            },
            LeftBorder = new PdfPanelBorder {
                Color = blue,
                Width = global::OfficeIMO.OfficeLength.Points(1.5)
            },
            PaddingX = global::OfficeIMO.OfficeLength.Points(8),
            PaddingY = global::OfficeIMO.OfficeLength.Points(6),
            SpacingAfter = global::OfficeIMO.OfficeLength.Points(0)
        };
        var options = new PdfCreateOptions {
            PageWidth = global::OfficeIMO.OfficeLength.Points(260),
            PageHeight = global::OfficeIMO.OfficeLength.Points(180),
            MarginLeft = global::OfficeIMO.OfficeLength.Points(30),
            MarginRight = global::OfficeIMO.OfficeLength.Points(30),
            MarginTop = global::OfficeIMO.OfficeLength.Points(30),
            MarginBottom = global::OfficeIMO.OfficeLength.Points(30),
            DefaultPanelStyle = style
        };

        style.TopBorder = new PdfPanelBorder {
            Color = global::OfficeIMO.OfficeColor.FromRgb(0, 128, 0),
            Width = global::OfficeIMO.OfficeLength.Points(4)
        };
        PdfPanelStyle readback = options.DefaultPanelStyle!;
        readback.LeftBorder = new PdfPanelBorder {
            Color = global::OfficeIMO.OfficeColor.Black,
            Width = global::OfficeIMO.OfficeLength.Points(3)
        };

        PdfCreateOptions clone = options.Clone();
        byte[] bytes = PdfDocument.Create(options)
            .PanelParagraph(p => p.Text("PanelSideBorders"))
            .ToBytes();

        string raw = PdfTestContentText.ReadDecoded(bytes);
        using var pdf = PdfPigDocument.Open(new MemoryStream(bytes));

        Assert.Equal(red, options.DefaultPanelStyle!.TopBorder!.Color);
        Assert.Equal(2, options.DefaultPanelStyle.TopBorder.Width.ToPoints());
        Assert.Equal(blue, options.DefaultPanelStyle.LeftBorder!.Color);
        Assert.Equal(1.5, options.DefaultPanelStyle.LeftBorder.Width.ToPoints());
        Assert.Equal(red, clone.DefaultPanelStyle!.TopBorder!.Color);
        Assert.Equal(blue, clone.DefaultPanelStyle.LeftBorder!.Color);
        Assert.Contains("PanelSideBorders", pdf.GetPage(1).Text);
        Assert.Contains("1 0 0 RG", raw);
        Assert.Contains("2 w", raw);
        Assert.Contains("0 0 1 RG", raw);
        Assert.Contains("1.5 w", raw);
    }

    [Fact]
    public void Options_SnapshotDefaultHorizontalRuleStyle() {
        var style = new PdfHorizontalRuleStyle {
            Thickness = global::OfficeIMO.OfficeLength.Points(1.4),
            Color = global::OfficeIMO.OfficeColor.FromRgb(20, 40, 60),
            SpacingBefore = global::OfficeIMO.OfficeLength.Points(5),
            SpacingAfter = global::OfficeIMO.OfficeLength.Points(13),
            KeepWithNext = true
        };
        var options = new PdfCreateOptions {
            DefaultHorizontalRuleStyle = style
        };

        style.Thickness = global::OfficeIMO.OfficeLength.Points(3);
        style.Color = global::OfficeIMO.OfficeColor.Black;
        style.SpacingBefore = global::OfficeIMO.OfficeLength.Points(1);
        style.SpacingAfter = global::OfficeIMO.OfficeLength.Points(2);
        style.KeepWithNext = false;

        PdfHorizontalRuleStyle readback = options.DefaultHorizontalRuleStyle!;
        readback.Thickness = global::OfficeIMO.OfficeLength.Points(4);
        readback.Color = global::OfficeIMO.OfficeColor.FromRgb(200, 10, 10);

        PdfCreateOptions clone = options.Clone();

        Assert.Equal(1.4, options.DefaultHorizontalRuleStyle!.Thickness.ToPoints());
        Assert.Equal(global::OfficeIMO.OfficeColor.FromRgb(20, 40, 60), options.DefaultHorizontalRuleStyle.Color);
        Assert.Equal(5, options.DefaultHorizontalRuleStyle.SpacingBefore.ToPoints());
        Assert.Equal(13, options.DefaultHorizontalRuleStyle.SpacingAfter.ToPoints());
        Assert.True(options.DefaultHorizontalRuleStyle.KeepWithNext);
        Assert.Equal(1.4, clone.DefaultHorizontalRuleStyle!.Thickness.ToPoints());
        Assert.Equal(global::OfficeIMO.OfficeColor.FromRgb(20, 40, 60), clone.DefaultHorizontalRuleStyle.Color);
        Assert.True(clone.DefaultHorizontalRuleStyle.KeepWithNext);
    }

    [Fact]
    public void Options_ApplyThemeSnapshotsDefaultStyles() {
        var textStyle = new PdfTextStyle {
            Font = PdfStandardFont.Helvetica,
            FontSize = global::OfficeIMO.OfficeLength.Points(16),
            Color = global::OfficeIMO.OfficeColor.FromRgb(10, 20, 30)
        };
        var paragraphStyle = new PdfParagraphStyle {
            FirstLineIndent = global::OfficeIMO.OfficeLength.Points(24),
            SpacingAfter = global::OfficeIMO.OfficeLength.Points(0)
        };
        var tableStyle = PdfTableStyles.Minimal();
        tableStyle.CellPaddingX = global::OfficeIMO.OfficeLength.Points(22);
        tableStyle.KeepWithNext = true;
        var headingStyle = new PdfHeadingStyle {
            FontSize = global::OfficeIMO.OfficeLength.Points(19),
            SpacingAfter = global::OfficeIMO.OfficeLength.Points(14),
            Color = global::OfficeIMO.OfficeColor.FromRgb(40, 50, 60)
        };
        var listStyle = new PdfListStyle {
            FontSize = global::OfficeIMO.OfficeLength.Points(13),
            LeftIndent = global::OfficeIMO.OfficeLength.Points(18),
            Color = global::OfficeIMO.OfficeColor.FromRgb(70, 80, 90),
            KeepTogether = true,
            KeepWithNext = true
        };
        var panelStyle = new PdfPanelStyle {
            PaddingX = global::OfficeIMO.OfficeLength.Points(16),
            MaxWidth = global::OfficeIMO.OfficeLength.Points(190),
            Background = global::OfficeIMO.OfficeColor.FromRgb(240, 248, 255),
            KeepWithNext = true
        };
        var horizontalRuleStyle = new PdfHorizontalRuleStyle {
            Thickness = global::OfficeIMO.OfficeLength.Points(1.4),
            Color = global::OfficeIMO.OfficeColor.FromRgb(80, 90, 100),
            SpacingAfter = global::OfficeIMO.OfficeLength.Points(17),
            KeepWithNext = true
        };
        var rowStyle = new PdfRowStyle {
            Gap = global::OfficeIMO.OfficeLength.Points(21),
            SpacingBefore = global::OfficeIMO.OfficeLength.Points(5),
            SpacingAfter = global::OfficeIMO.OfficeLength.Points(6),
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
            PdfPanelStyle = panelStyle,
            HorizontalRuleStyle = horizontalRuleStyle,
            RowStyle = rowStyle
        };
        var options = new PdfCreateOptions().ApplyTheme(theme);

        textStyle.FontSize = global::OfficeIMO.OfficeLength.Points(8);
        textStyle.Color = global::OfficeIMO.OfficeColor.Black;
        paragraphStyle.FirstLineIndent = global::OfficeIMO.OfficeLength.Points(0);
        tableStyle.CellPaddingX = global::OfficeIMO.OfficeLength.Points(0);
        tableStyle.KeepWithNext = false;
        headingStyle.FontSize = global::OfficeIMO.OfficeLength.Points(8);
        listStyle.FontSize = global::OfficeIMO.OfficeLength.Points(8);
        listStyle.KeepTogether = false;
        listStyle.KeepWithNext = false;
        panelStyle.PaddingX = global::OfficeIMO.OfficeLength.Points(4);
        panelStyle.KeepWithNext = false;
        horizontalRuleStyle.Thickness = global::OfficeIMO.OfficeLength.Points(5);
        horizontalRuleStyle.Color = global::OfficeIMO.OfficeColor.Black;
        horizontalRuleStyle.KeepWithNext = false;
        rowStyle.Gap = global::OfficeIMO.OfficeLength.Points(3);
        rowStyle.SpacingBefore = global::OfficeIMO.OfficeLength.Points(1);
        rowStyle.SpacingAfter = global::OfficeIMO.OfficeLength.Points(2);
        rowStyle.KeepTogether = false;
        rowStyle.KeepWithNext = false;
        theme.TextStyle = new PdfTextStyle {
            Font = PdfStandardFont.Helvetica,
            FontSize = global::OfficeIMO.OfficeLength.Points(7)
        };

        PdfCreateOptions clone = options.Clone();

        Assert.Equal(PdfStandardFont.Helvetica, options.DefaultFont);
        Assert.Equal(16, options.DefaultFontSize.ToPoints());
        Assert.Equal(global::OfficeIMO.OfficeColor.FromRgb(10, 20, 30), options.DefaultTextColor);
        Assert.Equal(24, options.DefaultParagraphStyle!.FirstLineIndent.ToPoints());
        Assert.Equal(22, options.DefaultTableStyle!.CellPaddingX.ToPoints());
        Assert.True(options.DefaultTableStyle.KeepWithNext);
        Assert.Equal(19, (options.DefaultHeadingStyles!.Level1!.FontSize?.ToPoints()));
        Assert.Equal(14, (options.DefaultHeadingStyles.Level1.SpacingAfter?.ToPoints()));
        Assert.Equal(13, (options.DefaultListStyle!.FontSize?.ToPoints()));
        Assert.Equal(18, options.DefaultListStyle.LeftIndent.ToPoints());
        Assert.True(options.DefaultListStyle.KeepTogether);
        Assert.True(options.DefaultListStyle.KeepWithNext);
        Assert.Equal(16, options.DefaultPanelStyle!.PaddingX.ToPoints());
        Assert.Equal(190, (options.DefaultPanelStyle.MaxWidth?.ToPoints()));
        Assert.True(options.DefaultPanelStyle.KeepWithNext);
        Assert.Equal(1.4, options.DefaultHorizontalRuleStyle!.Thickness.ToPoints());
        Assert.Equal(global::OfficeIMO.OfficeColor.FromRgb(80, 90, 100), options.DefaultHorizontalRuleStyle.Color);
        Assert.Equal(17, options.DefaultHorizontalRuleStyle.SpacingAfter.ToPoints());
        Assert.True(options.DefaultHorizontalRuleStyle.KeepWithNext);
        Assert.Equal(21, (options.DefaultRowStyle!.Gap?.ToPoints()));
        Assert.Equal(5, options.DefaultRowStyle.SpacingBefore.ToPoints());
        Assert.Equal(6, options.DefaultRowStyle.SpacingAfter.ToPoints());
        Assert.True(options.DefaultRowStyle.KeepTogether);
        Assert.True(options.DefaultRowStyle.KeepWithNext);
        Assert.Equal(16, clone.DefaultFontSize.ToPoints());
        Assert.Equal(24, clone.DefaultParagraphStyle!.FirstLineIndent.ToPoints());
        Assert.Equal(22, clone.DefaultTableStyle!.CellPaddingX.ToPoints());
        Assert.True(clone.DefaultTableStyle.KeepWithNext);
        Assert.Equal(19, (clone.DefaultHeadingStyles!.Level1!.FontSize?.ToPoints()));
        Assert.Equal(13, (clone.DefaultListStyle!.FontSize?.ToPoints()));
        Assert.True(clone.DefaultListStyle.KeepTogether);
        Assert.True(clone.DefaultListStyle.KeepWithNext);
        Assert.Equal(16, clone.DefaultPanelStyle!.PaddingX.ToPoints());
        Assert.True(clone.DefaultPanelStyle.KeepWithNext);
        Assert.Equal(1.4, clone.DefaultHorizontalRuleStyle!.Thickness.ToPoints());
        Assert.True(clone.DefaultHorizontalRuleStyle.KeepWithNext);
        Assert.Equal(21, (clone.DefaultRowStyle!.Gap?.ToPoints()));
        Assert.True(clone.DefaultRowStyle.KeepTogether);
        Assert.True(clone.DefaultRowStyle.KeepWithNext);
    }


}
