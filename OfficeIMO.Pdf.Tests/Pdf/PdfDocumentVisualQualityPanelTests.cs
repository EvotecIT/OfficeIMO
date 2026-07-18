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
    public void PanelParagraph_LeavesBreathingRoomBeforeFollowingParagraph() {
        const double fontSize = 10;
        byte[] bytes = PdfDocument.Create(new PdfOptions {
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

        using var pdf = PdfPigDocument.Open(new MemoryStream(bytes));
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
        byte[] bytes = PdfDocument.Create(new PdfOptions {
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

        using var pdf = PdfPigDocument.Open(new MemoryStream(bytes));
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

        byte[] bytes = PdfDocument.Create(options)
            .DefaultPanelStyle(style)
            .Compose(document =>
                document.Page(page =>
                    page.Content(content =>
                        content.Row(row =>
                            row.Column(100, column => column
                                .PanelParagraph(p => p.Text("ColumnPanel")))))))
            .ToBytes();

        using var pdf = PdfPigDocument.Open(new MemoryStream(bytes));
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

        using var defaultPdf = PdfPigDocument.Open(new MemoryStream(defaultBytes));
        using var spacedPdf = PdfPigDocument.Open(new MemoryStream(spacedBytes));
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
            PdfDocument.Create(new PdfOptions {
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
    public void PanelParagraph_WithTightSplitPadding_RendersForwardProgressAcrossPages() {
        byte[] bytes = PdfDocument.Create(new PdfOptions {
                PageWidth = 180,
                PageHeight = 110,
                MarginLeft = 20,
                MarginRight = 20,
                MarginTop = 20,
                MarginBottom = 20,
                DefaultFont = PdfStandardFont.Helvetica,
                DefaultFontSize = 10
            })
            .PanelParagraph(p => p
                .Text("FirstSegment")
                .LineBreak()
                .Text("SecondSegment"), new PanelStyle {
                    PaddingY = 56,
                    BorderWidth = 0,
                    SpacingAfter = 0
                })
            .ToBytes();

        using var pdf = PdfPigDocument.Open(new MemoryStream(bytes));
        string text = PdfReadDocument.Open(bytes).ExtractText();

        Assert.Equal(2, pdf.NumberOfPages);
        Assert.Contains("FirstSegment", text, StringComparison.Ordinal);
        Assert.Contains("SecondSegment", text, StringComparison.Ordinal);
    }

    [Fact]
    public void PanelParagraph_KeepTogetherOversizedContentSplitsAcrossPages() {
        byte[] bytes = PdfDocument.Create(new PdfOptions {
                PageWidth = 180,
                PageHeight = 130,
                MarginLeft = 20,
                MarginRight = 20,
                MarginTop = 20,
                MarginBottom = 20,
                DefaultFont = PdfStandardFont.Helvetica,
                DefaultFontSize = 10
            })
            .PanelParagraph(paragraph => {
                for (int index = 1; index <= 16; index++) {
                    if (index > 1) {
                        paragraph.LineBreak();
                    }

                    paragraph.Text("PanelLine" + index.ToString(CultureInfo.InvariantCulture));
                }
            }, new PanelStyle {
                KeepTogether = true,
                PaddingY = 6,
                BorderWidth = 0.5,
                SpacingAfter = 0
            })
            .ToBytes();

        using var pdf = PdfPigDocument.Open(new MemoryStream(bytes));
        string text = PdfReadDocument.Open(bytes).ExtractText();

        Assert.True(pdf.NumberOfPages > 1);
        Assert.Contains("PanelLine1", text, StringComparison.Ordinal);
        Assert.Contains("PanelLine16", text, StringComparison.Ordinal);
    }

    [Fact]
    public void PanelParagraph_RejectsVerticalPaddingThatCannotFitFirstLine() {
        var exception = Assert.Throws<ArgumentException>(() =>
            PdfDocument.Create(new PdfOptions {
                    PageWidth = 180,
                    PageHeight = 100,
                    MarginLeft = 20,
                    MarginRight = 20,
                    MarginTop = 20,
                    MarginBottom = 20,
                    DefaultFont = PdfStandardFont.Helvetica,
                    DefaultFontSize = 10
                })
                .PanelParagraph(p => p.Text("CannotFit"), new PanelStyle {
                    PaddingY = 70,
                    BorderWidth = 0
                })
                .ToBytes());

        Assert.Contains("Panel vertical padding and first line height exceed the available page content height.", exception.Message, StringComparison.Ordinal);
    }

    [Fact]
    public void PanelParagraph_KeepTogetherRejectsSingleLineWhenBottomPaddingCannotFit() {
        var exception = Assert.Throws<ArgumentException>(() =>
            PdfDocument.Create(new PdfOptions {
                    PageWidth = 180,
                    PageHeight = 110,
                    MarginLeft = 20,
                    MarginRight = 20,
                    MarginTop = 20,
                    MarginBottom = 20,
                    DefaultFont = PdfStandardFont.Helvetica,
                    DefaultFontSize = 10
                })
                .PanelParagraph(p => p.Text("CannotFitWithBottomPadding"), new PanelStyle {
                    KeepTogether = true,
                    PaddingY = 36,
                    BorderWidth = 0
                })
                .ToBytes());

        Assert.Contains("Panel vertical padding and first line height exceed the available page content height.", exception.Message, StringComparison.Ordinal);
    }

    [Fact]
    public void RowColumnPanelParagraph_RejectsVerticalPaddingThatCannotFitFirstLine() {
        var exception = Assert.Throws<ArgumentException>(() =>
            PdfDocument.Create(new PdfOptions {
                    PageWidth = 180,
                    PageHeight = 100,
                    MarginLeft = 20,
                    MarginRight = 20,
                    MarginTop = 20,
                    MarginBottom = 20,
                    DefaultFont = PdfStandardFont.Helvetica,
                    DefaultFontSize = 10
                })
                .Compose(document =>
                    document.Page(page =>
                        page.Content(content =>
                            content.Row(row =>
                                row.Column(100, column => column
                                    .PanelParagraph(p => p.Text("CannotFit"), new PanelStyle {
                                        PaddingY = 70,
                                        BorderWidth = 0
                                    }))))))
                .ToBytes());

        Assert.Contains("Panel vertical padding and first line height exceed the available page content height.", exception.Message, StringComparison.Ordinal);
    }

    [Fact]
    public void Panel_ComposesCommonFlowBlocksIntoStyledPanel() {
        byte[] pdf = PdfDocument.Create()
            .Panel(panel => panel
                    .H3("Panel Snapshot")
                    .Paragraph(p => p.Text("A reusable panel can combine regular document blocks."))
                    .RichBullets(new[] {
                        PdfListItem.Rich(new[] {
                            TextRun.Bolded("Reusable"),
                            TextRun.Normal(" core behavior")
                        })
                    })
                    .Table(new[] {
                        new[] { "Area", "State" },
                        new[] { "Markdown", "Ready" }
                    }, style: new PdfTableStyle {
                        HeaderRowCount = 1
                    })
                    .HR()
                    .PanelParagraph(p => p.Bold("Nested note").Text(": still uses panel text rendering.")),
                new PanelStyle {
                    Background = PdfColor.FromRgb(248, 250, 252),
                    BorderColor = PdfColor.FromRgb(37, 99, 235),
                    BorderWidth = 0.8,
                    PaddingX = 10,
                    PaddingY = 8
                })
            .ToBytes();

        string text = PdfReadDocument.Open(pdf).ExtractText();

        Assert.Contains("Panel Snapshot", text);
        Assert.Contains("reusable panel", text, StringComparison.OrdinalIgnoreCase);
        Assert.Contains("Reusable core behavior", text);
        Assert.Contains("Area: Markdown", text);
        Assert.Contains("State: Ready", text);
        Assert.Contains("Nested note", text);
    }

    [Fact]
    public void Panel_ComposesChecklistTablesAsReadableTaskStates() {
        var checklistStyle = TableStyles.Minimal();
        checklistStyle.CellIcons = new Dictionary<(int Row, int Column), PdfCellIcon> {
            [(0, 0)] = new PdfCellIcon {
                Kind = PdfCellIconKind.CheckBoxChecked,
                Color = PdfColor.FromRgb(22, 163, 74)
            },
            [(1, 0)] = new PdfCellIcon {
                Kind = PdfCellIconKind.CheckBoxUnchecked,
                Color = PdfColor.FromRgb(100, 116, 139)
            }
        };

        byte[] pdf = PdfDocument.Create()
            .Panel(panel => panel.Table(new[] {
                new[] { string.Empty, "Ship polished Markdown checklist visuals" },
                new[] { string.Empty, "Keep literal task markers out of the PDF text" }
            }, style: checklistStyle))
            .ToBytes();

        string text = PdfReadDocument.Open(pdf).ExtractText();

        Assert.Contains("Done: Ship polished Markdown checklist visuals", text, StringComparison.Ordinal);
        Assert.Contains("Open: Keep literal task markers out of the PDF text", text, StringComparison.Ordinal);
        Assert.DoesNotContain("[x]", text, StringComparison.OrdinalIgnoreCase);
        Assert.DoesNotContain("[ ]", text, StringComparison.Ordinal);
    }

    [Fact]
    public void ElementCompose_CanComposePanel() {
        byte[] pdf = PdfDocument.Create()
            .Compose(document =>
                document.Page(page =>
                    page.Content(content =>
                        content.Item(item =>
                            item.Element(element =>
                                element.Panel(panel => panel
                                        .H3("Element Panel")
                                        .Paragraph(p => p.Text("Nested element groups can use composed panels.")),
                                    new PanelStyle {
                                        Background = PdfColor.FromRgb(248, 250, 252),
                                        BorderColor = PdfColor.FromRgb(148, 163, 184),
                                        PaddingX = 8,
                                        PaddingY = 6
                                    }))))))
            .ToBytes();

        string text = PdfReadDocument.Open(pdf).ExtractText();

        Assert.Contains("Element Panel", text, StringComparison.Ordinal);
        Assert.Contains("Nested element groups can use composed panels.", text, StringComparison.Ordinal);
    }

    [Fact]
    public void ElementCompose_CanComposeCommonDocumentPrimitives() {
        byte[] pdf = PdfDocument.Create()
            .Compose(document =>
                document.Page(page =>
                    page.Content(content =>
                        content.Item(item =>
                            item.Element(element => element
                                .RichBullets(new[] {
                                    PdfListItem.Rich(new[] {
                                        TextRun.Bolded("Rich"),
                                        TextRun.Normal(" bullet")
                                    })
                                })
                                .RichNumbered(new[] {
                                    PdfListItem.Rich(new[] {
                                        TextRun.Normal("Numbered element item")
                                    })
                                })
                                .HR()
                                .PanelParagraph(p => p.Bold("Element note").Text(": composed in a grouped flow.")))))))
            .ToBytes();

        string text = PdfReadDocument.Open(pdf).ExtractText();

        Assert.Contains("Rich bullet", text, StringComparison.Ordinal);
        Assert.Contains("Numbered element item", text, StringComparison.Ordinal);
        Assert.Contains("Element note", text, StringComparison.Ordinal);
    }

    [Fact]
    public void Panel_RejectsUnsupportedNestedFlowBlocks() {
        var exception = Assert.Throws<NotSupportedException>(() =>
            PdfDocument.Create()
                .Panel(panel => panel.TextField("InsidePanel")));

        Assert.Contains("Panel currently supports paragraphs, headings, lists, simple tables, horizontal rules, spacers, bookmarks, and nested panel paragraphs.", exception.Message, StringComparison.Ordinal);
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

        var doc = PdfDocument.Create(new PdfOptions {
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


}
