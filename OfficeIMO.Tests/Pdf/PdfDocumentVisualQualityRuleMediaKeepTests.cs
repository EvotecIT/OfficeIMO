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
    public void HorizontalRule_RendersInTopLevelFlow() {
        byte[] bytes = PdfDocument.Create(new PdfOptions {
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
            .HR(
                thickness: 3,
                color: PdfColor.FromRgb(26, 51, 77),
                spacingBefore: 4,
                spacingAfter: 6)
            .Paragraph(p => p.Text("Guarded rhythm stays below the rule."))
            .ToBytes();

        using var pdf = PdfPigDocument.Open(new MemoryStream(bytes));
        var page = pdf.GetPage(1);

        double ruleBottomY = 180 - 20 - 3;
        double paragraphTopY = FindWordStartY(page, "Guarded") + fontSize * 0.74;
        double clearance = ruleBottomY - paragraphTopY;

        Assert.True(clearance >= 5, $"Expected rule spacing to leave visible breathing room before paragraph text. Clearance: {clearance:0.##}pt.");
    }

    [Fact]
    public void RowColumnHorizontalRule_LeavesBreathingRoomBeforeFollowingParagraph() {
        const double fontSize = 10;
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

        using var pdf = PdfPigDocument.Open(new MemoryStream(bytes));
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

        byte[] bytes = PdfDocument.Create(options)
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

        using var pdf = PdfPigDocument.Open(new MemoryStream(bytes));
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

        byte[] bytes = PdfDocument.Create(options)
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

        using var pdf = PdfPigDocument.Open(new MemoryStream(bytes));
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

        byte[] bytes = PdfDocument.Create(options)
            .Paragraph(p => p.Text("IntroMarker"), style: new PdfParagraphStyle {
                SpacingAfter = 66
            })
            .Image(png, 24, 24, style: new PdfImageStyle {
                KeepWithNext = true
            })
            .Paragraph(p => p.Text("FollowingImageBody"))
            .ToBytes();

        using var pdf = PdfPigDocument.Open(new MemoryStream(bytes));
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
    public void Image_KeepWithNextMeasuresFollowingHeadingChain() {
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

        byte[] bytes = PdfDocument.Create(options)
            .Paragraph(p => p.Text("IntroMarker"), style: new PdfParagraphStyle {
                SpacingAfter = 56
            })
            .Image(png, 24, 24, style: new PdfImageStyle {
                KeepWithNext = true
            })
            .H3("FollowingImageHeading")
            .Paragraph(p => p.Text("FollowingImageChainBody"))
            .ToBytes();

        using var pdf = PdfPigDocument.Open(new MemoryStream(bytes));
        string page1Content = string.Join("\n", GetPageContentStreams(bytes, 1));
        string page2Content = string.Join("\n", GetPageContentStreams(bytes, 2));

        Assert.Equal(2, pdf.NumberOfPages);
        Assert.Contains("IntroMarker", pdf.GetPage(1).Text);
        Assert.DoesNotContain("FollowingImageHeading", pdf.GetPage(1).Text);
        Assert.DoesNotContain("FollowingImageChainBody", pdf.GetPage(1).Text);
        Assert.Contains("FollowingImageHeading", pdf.GetPage(2).Text);
        Assert.Contains("FollowingImageChainBody", pdf.GetPage(2).Text);
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

        byte[] bytes = PdfDocument.Create(options)
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

        using var pdf = PdfPigDocument.Open(new MemoryStream(bytes));
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

        byte[] bytes = PdfDocument.Create(options)
            .Paragraph(p => p.Text("IntroMarker"), style: new PdfParagraphStyle {
                SpacingAfter = 66
            })
            .Shape(shape, style: new PdfDrawingStyle {
                KeepWithNext = true
            })
            .Paragraph(p => p.Text("FollowingShapeBody"))
            .ToBytes();

        using var pdf = PdfPigDocument.Open(new MemoryStream(bytes));
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

        byte[] bytes = PdfDocument.Create(options)
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

        using var pdf = PdfPigDocument.Open(new MemoryStream(bytes));
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

        byte[] bytes = PdfDocument.Create(options)
            .Paragraph(p => p.Text("IntroMarker"), style: new PdfParagraphStyle {
                SpacingAfter = 66
            })
            .Drawing(drawing, style: new PdfDrawingStyle {
                KeepWithNext = true
            })
            .Paragraph(p => p.Text("FollowingDrawingBody"))
            .ToBytes();

        using var pdf = PdfPigDocument.Open(new MemoryStream(bytes));
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

        byte[] bytes = PdfDocument.Create(options)
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

        using var pdf = PdfPigDocument.Open(new MemoryStream(bytes));
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

        byte[] bytes = PdfDocument.Create(options)
            .Paragraph(p => p.Text("IntroMarker"), style: new PdfParagraphStyle {
                SpacingAfter = 60
            })
            .Table(new[] {
                new[] { "TableKeepHeader", "Ready" },
                new[] { "TableKeepValue", "Ready" }
            }, style: style)
            .Paragraph(p => p.Text("FollowingTableBody"))
            .ToBytes();

        using var pdf = PdfPigDocument.Open(new MemoryStream(bytes));

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

        byte[] bytes = PdfDocument.Create(options)
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

        using var pdf = PdfPigDocument.Open(new MemoryStream(bytes));

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

        byte[] bytes = PdfDocument.Create(options)
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

        using var pdf = PdfPigDocument.Open(new MemoryStream(bytes));

        Assert.Equal(2, pdf.NumberOfPages);
        Assert.Contains("IntroMarker", pdf.GetPage(1).Text);
        Assert.DoesNotContain("RowKeepColumn", pdf.GetPage(1).Text);
        Assert.DoesNotContain("FollowingRowBody", pdf.GetPage(1).Text);
        Assert.Contains("RowKeepColumn", pdf.GetPage(2).Text);
        Assert.Contains("FollowingRowBody", pdf.GetPage(2).Text);
    }


}
