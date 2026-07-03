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

        byte[] bytes = PdfDocument.Create(new PdfOptions {
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

        using var pdf = PdfPigDocument.Open(new MemoryStream(bytes));
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

        byte[] bytes = PdfDocument.Create(new PdfOptions {
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

        using var pdf = PdfPigDocument.Open(new MemoryStream(bytes));

        Assert.True(pdf.NumberOfPages >= 2);
        Assert.DoesNotContain("KeptRowLeft", pdf.GetPage(1).Text, StringComparison.Ordinal);
        Assert.Contains("KeptRowLeft", pdf.GetPage(2).Text, StringComparison.Ordinal);
        Assert.Contains("KeptRowRight", pdf.GetPage(2).Text, StringComparison.Ordinal);
    }

    [Fact]
    public void RowStyle_KeepTogetherRejectsRowsTallerThanPageFrame() {
        string longText = string.Join(" ", Enumerable.Repeat("TooTallRowContent", 80));

        var exception = Assert.Throws<ArgumentException>(() =>
            PdfDocument.Create(new PdfOptions {
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
            PdfDocument.Create()
                .HR(thickness: 0));

        Assert.Contains("Horizontal rule thickness must be a positive finite value.", thicknessException.Message, StringComparison.Ordinal);

        var spacingBeforeException = Assert.Throws<ArgumentException>(() =>
            PdfDocument.Create()
                .HR(spacingBefore: -1));

        Assert.Contains("Horizontal rule spacing before must be a non-negative finite value.", spacingBeforeException.Message, StringComparison.Ordinal);

        var spacingAfterException = Assert.Throws<ArgumentException>(() =>
            PdfDocument.Create()
                .HR(spacingAfter: double.PositiveInfinity));

        Assert.Contains("Horizontal rule spacing after must be a non-negative finite value.", spacingAfterException.Message, StringComparison.Ordinal);

        var columnException = Assert.Throws<ArgumentException>(() =>
            PdfDocument.Create()
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
            PdfDocument.Create(new PdfOptions {
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
            PdfDocument.Create(new PdfOptions {
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
            PdfDocument.Create(new PdfOptions {
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
            PdfDocument.Create(new PdfOptions {
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


}
