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

        Assert.True(options.HasExplicitDefaultTableStyle);
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
        Assert.True(clone.HasExplicitDefaultTableStyle);
        Assert.False(new PdfOptions().HasExplicitDefaultTableStyle);
        Assert.False(new PdfOptions().Clone().HasExplicitDefaultTableStyle);
    }

    [Fact]
    public void Options_SnapshotDefaultHeadingStyles() {
        var style = new PdfHeadingStyle {
            FontSize = 16,
            LineHeight = 1.1,
            SpacingBefore = 4,
            SpacingAfter = 12,
            Color = PdfColor.FromRgb(10, 20, 30),
            Bold = false,
            ApplySpacingBeforeAtTop = true,
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
        Assert.False(options.DefaultHeadingStyles.Level1.Bold);
        Assert.True(options.DefaultHeadingStyles.Level1.ApplySpacingBeforeAtTop);
        Assert.False(options.DefaultHeadingStyles.Level1.KeepWithNext);
        Assert.Equal(16, clone.DefaultHeadingStyles!.Level1!.FontSize);
        Assert.False(clone.DefaultHeadingStyles.Level1.Bold);
        Assert.True(clone.DefaultHeadingStyles.Level1.ApplySpacingBeforeAtTop);
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
    public void PanelParagraph_RendersAndSnapshotsSideSpecificPanelBorders() {
        PdfColor red = PdfColor.FromRgb(255, 0, 0);
        PdfColor blue = PdfColor.FromRgb(0, 0, 255);
        var style = new PanelStyle {
            Background = PdfColor.FromRgb(245, 245, 245),
            TopBorder = new PdfPanelBorder {
                Color = red,
                Width = 2
            },
            LeftBorder = new PdfPanelBorder {
                Color = blue,
                Width = 1.5
            },
            PaddingX = 8,
            PaddingY = 6,
            SpacingAfter = 0
        };
        var options = new PdfOptions {
            PageWidth = 260,
            PageHeight = 180,
            MarginLeft = 30,
            MarginRight = 30,
            MarginTop = 30,
            MarginBottom = 30,
            DefaultPanelStyle = style
        };

        style.TopBorder = new PdfPanelBorder {
            Color = PdfColor.FromRgb(0, 128, 0),
            Width = 4
        };
        PanelStyle readback = options.DefaultPanelStyle!;
        readback.LeftBorder = new PdfPanelBorder {
            Color = PdfColor.Black,
            Width = 3
        };

        PdfOptions clone = options.Clone();
        byte[] bytes = PdfDocument.Create(options)
            .PanelParagraph(p => p.Text("PanelSideBorders"))
            .ToBytes();

        string raw = Encoding.ASCII.GetString(bytes);
        using var pdf = PdfPigDocument.Open(new MemoryStream(bytes));

        Assert.Equal(red, options.DefaultPanelStyle!.TopBorder!.Color);
        Assert.Equal(2, options.DefaultPanelStyle.TopBorder.Width);
        Assert.Equal(blue, options.DefaultPanelStyle.LeftBorder!.Color);
        Assert.Equal(1.5, options.DefaultPanelStyle.LeftBorder.Width);
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


}
