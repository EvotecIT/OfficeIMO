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

        byte[] bytes = PdfDocument.Create(options)
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

        using var pdf = PdfPigDocument.Open(new MemoryStream(bytes));

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

        byte[] bytes = PdfDocument.Create(options)
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

        using var pdf = PdfPigDocument.Open(new MemoryStream(bytes));

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

        byte[] bytes = PdfDocument.Create(options)
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

        using var pdf = PdfPigDocument.Open(new MemoryStream(bytes));

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

        byte[] bytes = PdfDocument.Create(options)
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

        using var pdf = PdfPigDocument.Open(new MemoryStream(bytes));

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

        byte[] bytes = PdfDocument.Create(options)
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

        using var pdf = PdfPigDocument.Open(new MemoryStream(bytes));

        Assert.Equal(2, pdf.NumberOfPages);
        Assert.Contains("IntroMarker", pdf.GetPage(1).Text);
        Assert.DoesNotContain("ColumnKeepWithListLabel", pdf.GetPage(1).Text);
        Assert.Contains("ColumnKeepWithListLabel", pdf.GetPage(2).Text);
        Assert.Contains("ColumnFollowingBullet", pdf.GetPage(2).Text);
    }

    [Fact]
    public void RowColumnParagraph_KeepWithNextMeasuresFollowingHeadingChain() {
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
            .Compose(document =>
                document.Page(page =>
                    page.Content(content =>
                        content.Row(row =>
                            row.Column(100, column => column
                                .Paragraph(p => p.Text("ColumnChainKeepWithNext"), style: new PdfParagraphStyle {
                                    KeepWithNext = true
                                })
                                .H3("ColumnFollowingHeading")
                                .Paragraph(p => p.Text("ColumnFollowingHeadingBody")))))))
            .ToBytes();

        using var pdf = PdfPigDocument.Open(new MemoryStream(bytes));

        Assert.Equal(2, pdf.NumberOfPages);
        Assert.Contains("IntroMarker", pdf.GetPage(1).Text);
        Assert.DoesNotContain("ColumnChainKeepWithNext", pdf.GetPage(1).Text);
        Assert.DoesNotContain("ColumnFollowingHeading", pdf.GetPage(1).Text);
        Assert.DoesNotContain("ColumnFollowingHeadingBody", pdf.GetPage(1).Text);
        Assert.Contains("ColumnChainKeepWithNext", pdf.GetPage(2).Text);
        Assert.Contains("ColumnFollowingHeading", pdf.GetPage(2).Text);
        Assert.Contains("ColumnFollowingHeadingBody", pdf.GetPage(2).Text);
    }

    [Fact]
    public void RowColumnParagraph_LongKeepWithNextChainRendersWithBoundedMeasurement() {
        byte[] bytes = PdfDocument.Create(new PdfOptions {
                PageWidth = 300,
                PageHeight = 400,
                MarginLeft = 30,
                MarginRight = 30,
                MarginTop = 30,
                MarginBottom = 30,
                DefaultFont = PdfStandardFont.Helvetica,
                DefaultFontSize = 8
            })
            .Compose(document =>
                document.Page(page =>
                    page.Content(content =>
                        content.Row(row =>
                            row.Column(100, column => {
                                for (int i = 0; i < 768; i++) {
                                    column.Paragraph(
                                        paragraph => paragraph.Text("Keep" + i.ToString(CultureInfo.InvariantCulture)),
                                        style: new PdfParagraphStyle { KeepWithNext = true });
                                }

                                column.Paragraph(paragraph => paragraph.Text("KeepChainTail"));
                            })))))
            .ToBytes();

        using var pdf = PdfPigDocument.Open(new MemoryStream(bytes));
        Assert.True(pdf.NumberOfPages > 1);
        Assert.Contains("KeepChainTail", pdf.GetPage(pdf.NumberOfPages).Text);
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

        byte[] bytes = PdfDocument.Create(options)
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

        using var pdf = PdfPigDocument.Open(new MemoryStream(bytes));

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

        byte[] bytes = PdfDocument.Create(options)
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

        using var pdf = PdfPigDocument.Open(new MemoryStream(bytes));

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

        byte[] bytes = PdfDocument.Create(options)
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

        using var pdf = PdfPigDocument.Open(new MemoryStream(bytes));

        Assert.Equal(2, pdf.NumberOfPages);
        Assert.Contains("IntroMarker", pdf.GetPage(1).Text);
        Assert.DoesNotContain("ColumnPanelKeepWithNext", pdf.GetPage(1).Text);
        Assert.Contains("ColumnPanelKeepWithNext", pdf.GetPage(2).Text);
        Assert.Contains("ColumnFollowingPanelBody", pdf.GetPage(2).Text);
    }


}
