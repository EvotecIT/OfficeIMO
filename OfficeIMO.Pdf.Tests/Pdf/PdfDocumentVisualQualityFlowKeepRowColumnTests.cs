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
        var options = new PdfCreateOptions {
            PageWidth = global::OfficeIMO.OfficeLength.Points(260),
            PageHeight = global::OfficeIMO.OfficeLength.Points(170),
            MarginLeft = global::OfficeIMO.OfficeLength.Points(30),
            MarginRight = global::OfficeIMO.OfficeLength.Points(30),
            MarginTop = global::OfficeIMO.OfficeLength.Points(30),
            MarginBottom = global::OfficeIMO.OfficeLength.Points(30),
            DefaultFont = PdfStandardFont.Helvetica,
            DefaultFontSize = global::OfficeIMO.OfficeLength.Points(10)
        };

        byte[] bytes = PdfDocument.Create(options)
            .Paragraph(p => p.AddText("IntroMarker"), style: new PdfParagraphStyle {
                SpacingAfter = global::OfficeIMO.OfficeLength.Points(70)
            })
            .Compose(document =>
                document.Page(page =>
                    page.Content(content =>
                        content.Row(row =>
                            row.PercentColumn(100, column => column
                                .AddParagraph(p => {
                                    p.AddText("ColumnKeepFirst");
                                    p.AddLineBreak();
                                    p.AddText("ColumnKeepMiddle");
                                    p.AddLineBreak();
                                    p.AddText("ColumnKeepLast");
                                }, style: new PdfParagraphStyle {
                                        KeepTogether = true,
                                        SpacingAfter = global::OfficeIMO.OfficeLength.Points(0)
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
        var options = new PdfCreateOptions {
            PageWidth = global::OfficeIMO.OfficeLength.Points(260),
            PageHeight = global::OfficeIMO.OfficeLength.Points(170),
            MarginLeft = global::OfficeIMO.OfficeLength.Points(30),
            MarginRight = global::OfficeIMO.OfficeLength.Points(30),
            MarginTop = global::OfficeIMO.OfficeLength.Points(30),
            MarginBottom = global::OfficeIMO.OfficeLength.Points(30),
            DefaultFont = PdfStandardFont.Helvetica,
            DefaultFontSize = global::OfficeIMO.OfficeLength.Points(10)
        };

        byte[] bytes = PdfDocument.Create(options)
            .Paragraph(p => p.AddText("IntroMarker"), style: new PdfParagraphStyle {
                SpacingAfter = global::OfficeIMO.OfficeLength.Points(70)
            })
            .Compose(document =>
                document.Page(page =>
                    page.Content(content =>
                        content.Row(row =>
                            row.PercentColumn(100, column => column
                                .AddList(new[] {
                                    "ColumnKeepListFirst",
                                    "ColumnKeepListMiddle",
                                    "ColumnKeepListLast"
                                }, style: new PdfListStyle {
                                    KeepTogether = true,
                                    SpacingAfter = global::OfficeIMO.OfficeLength.Points(0)
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
        var options = new PdfCreateOptions {
            PageWidth = global::OfficeIMO.OfficeLength.Points(260),
            PageHeight = global::OfficeIMO.OfficeLength.Points(170),
            MarginLeft = global::OfficeIMO.OfficeLength.Points(30),
            MarginRight = global::OfficeIMO.OfficeLength.Points(30),
            MarginTop = global::OfficeIMO.OfficeLength.Points(30),
            MarginBottom = global::OfficeIMO.OfficeLength.Points(30),
            DefaultFont = PdfStandardFont.Helvetica,
            DefaultFontSize = global::OfficeIMO.OfficeLength.Points(10)
        };

        byte[] bytes = PdfDocument.Create(options)
            .Paragraph(p => p.AddText("IntroMarker"), style: new PdfParagraphStyle {
                SpacingAfter = global::OfficeIMO.OfficeLength.Points(60)
            })
            .Compose(document =>
                document.Page(page =>
                    page.Content(content =>
                        content.Row(row =>
                            row.PercentColumn(100, column => column
                                .AddNumberedList(new[] {
                                    "ColumnKeepNumberOne",
                                    "ColumnKeepNumberTwo"
                                }, style: new PdfListStyle {
                                    KeepWithNext = true,
                                    ItemSpacing = global::OfficeIMO.OfficeLength.Points(0),
                                    SpacingAfter = global::OfficeIMO.OfficeLength.Points(0)
                                })
                                .AddParagraph(p => p.AddText("ColumnFollowingListBody")))))))
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
        var options = new PdfCreateOptions {
            PageWidth = global::OfficeIMO.OfficeLength.Points(260),
            PageHeight = global::OfficeIMO.OfficeLength.Points(170),
            MarginLeft = global::OfficeIMO.OfficeLength.Points(30),
            MarginRight = global::OfficeIMO.OfficeLength.Points(30),
            MarginTop = global::OfficeIMO.OfficeLength.Points(30),
            MarginBottom = global::OfficeIMO.OfficeLength.Points(30),
            DefaultFont = PdfStandardFont.Helvetica,
            DefaultFontSize = global::OfficeIMO.OfficeLength.Points(10)
        };

        byte[] bytes = PdfDocument.Create(options)
            .Paragraph(p => p.AddText("IntroMarker"), style: new PdfParagraphStyle {
                SpacingAfter = global::OfficeIMO.OfficeLength.Points(70)
            })
            .Compose(document =>
                document.Page(page =>
                    page.Content(content =>
                        content.Row(row =>
                            row.PercentColumn(100, column => column
                                .AddParagraph(p => p.AddText("ColumnKeepWithNextLabel"), style: new PdfParagraphStyle {
                                    KeepWithNext = true
                                })
                                .AddParagraph(p => p.AddText("ColumnFollowingBody")))))))
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
        var options = new PdfCreateOptions {
            PageWidth = global::OfficeIMO.OfficeLength.Points(260),
            PageHeight = global::OfficeIMO.OfficeLength.Points(170),
            MarginLeft = global::OfficeIMO.OfficeLength.Points(30),
            MarginRight = global::OfficeIMO.OfficeLength.Points(30),
            MarginTop = global::OfficeIMO.OfficeLength.Points(30),
            MarginBottom = global::OfficeIMO.OfficeLength.Points(30),
            DefaultFont = PdfStandardFont.Helvetica,
            DefaultFontSize = global::OfficeIMO.OfficeLength.Points(10)
        };

        byte[] bytes = PdfDocument.Create(options)
            .Paragraph(p => p.AddText("IntroMarker"), style: new PdfParagraphStyle {
                SpacingAfter = global::OfficeIMO.OfficeLength.Points(70)
            })
            .Compose(document =>
                document.Page(page =>
                    page.Content(content =>
                        content.Row(row =>
                            row.PercentColumn(100, column => column
                                .AddParagraph(p => p.AddText("ColumnKeepWithListLabel"), style: new PdfParagraphStyle {
                                    KeepWithNext = true
                                })
                                .AddList(new[] { "ColumnFollowingBullet" }))))))
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
        var options = new PdfCreateOptions {
            PageWidth = global::OfficeIMO.OfficeLength.Points(260),
            PageHeight = global::OfficeIMO.OfficeLength.Points(170),
            MarginLeft = global::OfficeIMO.OfficeLength.Points(30),
            MarginRight = global::OfficeIMO.OfficeLength.Points(30),
            MarginTop = global::OfficeIMO.OfficeLength.Points(30),
            MarginBottom = global::OfficeIMO.OfficeLength.Points(30),
            DefaultFont = PdfStandardFont.Helvetica,
            DefaultFontSize = global::OfficeIMO.OfficeLength.Points(10)
        };

        byte[] bytes = PdfDocument.Create(options)
            .Paragraph(p => p.AddText("IntroMarker"), style: new PdfParagraphStyle {
                SpacingAfter = global::OfficeIMO.OfficeLength.Points(56)
            })
            .Compose(document =>
                document.Page(page =>
                    page.Content(content =>
                        content.Row(row =>
                            row.PercentColumn(100, column => column
                                .AddParagraph(p => p.AddText("ColumnChainKeepWithNext"), style: new PdfParagraphStyle {
                                    KeepWithNext = true
                                })
                                .H3("ColumnFollowingHeading")
                                .AddParagraph(p => p.AddText("ColumnFollowingHeadingBody")))))))
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
        byte[] bytes = PdfDocument.Create(new PdfCreateOptions {
                PageWidth = global::OfficeIMO.OfficeLength.Points(300),
                PageHeight = global::OfficeIMO.OfficeLength.Points(400),
                MarginLeft = global::OfficeIMO.OfficeLength.Points(30),
                MarginRight = global::OfficeIMO.OfficeLength.Points(30),
                MarginTop = global::OfficeIMO.OfficeLength.Points(30),
                MarginBottom = global::OfficeIMO.OfficeLength.Points(30),
                DefaultFont = PdfStandardFont.Helvetica,
                DefaultFontSize = global::OfficeIMO.OfficeLength.Points(8)
            })
            .Compose(document =>
                document.Page(page =>
                    page.Content(content =>
                        content.Row(row =>
                            row.PercentColumn(100, column => {
                                for (int i = 0; i < 768; i++) {
                                    column.AddParagraph(
                                        paragraph => paragraph.AddText("Keep" + i.ToString(CultureInfo.InvariantCulture)),
                                        style: new PdfParagraphStyle { KeepWithNext = true });
                                }

                                column.AddParagraph(paragraph => paragraph.AddText("KeepChainTail"));
                            })))))
            .ToBytes();

        using var pdf = PdfPigDocument.Open(new MemoryStream(bytes));
        Assert.True(pdf.NumberOfPages > 1);
        Assert.Contains("KeepChainTail", pdf.GetPage(pdf.NumberOfPages).Text);
    }

    [Fact]
    public void RowColumnHeading_KeepsWithFollowingParagraph() {
        var options = new PdfCreateOptions {
            PageWidth = global::OfficeIMO.OfficeLength.Points(260),
            PageHeight = global::OfficeIMO.OfficeLength.Points(170),
            MarginLeft = global::OfficeIMO.OfficeLength.Points(30),
            MarginRight = global::OfficeIMO.OfficeLength.Points(30),
            MarginTop = global::OfficeIMO.OfficeLength.Points(30),
            MarginBottom = global::OfficeIMO.OfficeLength.Points(30),
            DefaultFont = PdfStandardFont.Helvetica,
            DefaultFontSize = global::OfficeIMO.OfficeLength.Points(10)
        };

        byte[] bytes = PdfDocument.Create(options)
            .Compose(document =>
                document.Page(page =>
                    page.Content(content =>
                        content.Row(row =>
                            row.PercentColumn(100, column => column
                                .AddParagraph(p => p.AddText("ColumnIntroMarker"), style: new PdfParagraphStyle {
                                    SpacingAfter = global::OfficeIMO.OfficeLength.Points(70)
                                })
                                .H3("ColumnSignalHeading")
                                .AddParagraph(p => p.AddText("ColumnSignalBody")))))))
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
        var options = new PdfCreateOptions {
            PageWidth = global::OfficeIMO.OfficeLength.Points(260),
            PageHeight = global::OfficeIMO.OfficeLength.Points(170),
            MarginLeft = global::OfficeIMO.OfficeLength.Points(30),
            MarginRight = global::OfficeIMO.OfficeLength.Points(30),
            MarginTop = global::OfficeIMO.OfficeLength.Points(30),
            MarginBottom = global::OfficeIMO.OfficeLength.Points(30),
            DefaultFont = PdfStandardFont.Helvetica,
            DefaultFontSize = global::OfficeIMO.OfficeLength.Points(10)
        };

        byte[] bytes = PdfDocument.Create(options)
            .Compose(document =>
                document.Page(page =>
                    page.Content(content =>
                        content.Row(row =>
                            row.PercentColumn(100, column => column
                                .AddParagraph(p => p.AddText("ColumnIntroMarker"), style: new PdfParagraphStyle {
                                    SpacingAfter = global::OfficeIMO.OfficeLength.Points(70)
                                })
                                .AddParagraph(p => {
                                    p.AddText("ColumnWidowFirst");
                                    p.AddLineBreak();
                                    p.AddText("ColumnWidowSecond");
                                    p.AddLineBreak();
                                    p.AddText("ColumnWidowThird");
                                }, style: new PdfParagraphStyle {
                                        WidowControl = true,
                                        SpacingAfter = global::OfficeIMO.OfficeLength.Points(0)
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
        var options = new PdfCreateOptions {
            PageWidth = global::OfficeIMO.OfficeLength.Points(260),
            PageHeight = global::OfficeIMO.OfficeLength.Points(170),
            MarginLeft = global::OfficeIMO.OfficeLength.Points(30),
            MarginRight = global::OfficeIMO.OfficeLength.Points(30),
            MarginTop = global::OfficeIMO.OfficeLength.Points(30),
            MarginBottom = global::OfficeIMO.OfficeLength.Points(30),
            DefaultFont = PdfStandardFont.Helvetica,
            DefaultFontSize = global::OfficeIMO.OfficeLength.Points(10)
        };

        byte[] bytes = PdfDocument.Create(options)
            .Paragraph(p => p.AddText("IntroMarker"), style: new PdfParagraphStyle {
                SpacingAfter = global::OfficeIMO.OfficeLength.Points(60)
            })
            .Compose(document =>
                document.Page(page =>
                    page.Content(content =>
                        content.Row(row =>
                            row.PercentColumn(100, column => column
                                .AddPanelParagraph(p => p.AddText("ColumnPanelKeepWithNext"), new PdfPanelStyle {
                                    KeepWithNext = true,
                                    PaddingY = global::OfficeIMO.OfficeLength.Points(5),
                                    SpacingAfter = global::OfficeIMO.OfficeLength.Points(0)
                                })
                                .AddParagraph(p => p.AddText("ColumnFollowingPanelBody")))))))
            .ToBytes();

        using var pdf = PdfPigDocument.Open(new MemoryStream(bytes));

        Assert.Equal(2, pdf.NumberOfPages);
        Assert.Contains("IntroMarker", pdf.GetPage(1).Text);
        Assert.DoesNotContain("ColumnPanelKeepWithNext", pdf.GetPage(1).Text);
        Assert.Contains("ColumnPanelKeepWithNext", pdf.GetPage(2).Text);
        Assert.Contains("ColumnFollowingPanelBody", pdf.GetPage(2).Text);
    }


}
