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
    public void Paragraph_SplitsLongContentAcrossPagesWithoutCrossingBottomMargin() {
        var options = new PdfCreateOptions {
            PageWidth = global::OfficeIMO.OfficeLength.Points(320),
            PageHeight = global::OfficeIMO.OfficeLength.Points(170),
            MarginLeft = global::OfficeIMO.OfficeLength.Points(30),
            MarginRight = global::OfficeIMO.OfficeLength.Points(30),
            MarginTop = global::OfficeIMO.OfficeLength.Points(30),
            MarginBottom = global::OfficeIMO.OfficeLength.Points(30),
            DefaultFont = PdfStandardFont.Helvetica,
            DefaultFontSize = global::OfficeIMO.OfficeLength.Points(9)
        };

        string longText = string.Join(" ", Enumerable.Range(1, 180).Select(i => "segment" + i.ToString("000")));

        byte[] bytes = PdfDocument.Create(options)
            .Paragraph(p => p.Text(longText), style: new PdfParagraphStyle {
                LineHeight = 1.3,
                SpacingAfter = global::OfficeIMO.OfficeLength.Points(0)
            })
            .ToBytes();

        using var pdf = PdfPigDocument.Open(new MemoryStream(bytes));
        Assert.True(pdf.NumberOfPages > 1, "Expected a long rich paragraph to continue onto another page.");

        for (int pageNumber = 1; pageNumber <= pdf.NumberOfPages; pageNumber++) {
            var page = pdf.GetPage(pageNumber);
            double bottomMost = page.Letters
                .Where(letter => !string.IsNullOrWhiteSpace(letter.Value))
                .Min(letter => letter.StartBaseLine.Y);
            Assert.True(bottomMost >= options.MarginBottom.ToPoints() - 2, $"Expected paragraph text to stay above the bottom margin on page {pageNumber}.");
        }

        Assert.Contains("segment001", pdf.GetPage(1).Text);
        Assert.Contains("segment180", pdf.GetPage(pdf.NumberOfPages).Text);
    }

    [Fact]
    public void Paragraph_KeepTogetherMovesWholeParagraphToNextPage() {
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
            .Paragraph(p => p.Text("IntroMarker"), style: new PdfParagraphStyle {
                SpacingAfter = global::OfficeIMO.OfficeLength.Points(70)
            })
            .Paragraph(p => p
                .Text("KeepFirst")
                .LineBreak()
                .Text("KeepMiddle")
                .LineBreak()
                .Text("KeepLast"), style: new PdfParagraphStyle {
                    KeepTogether = true,
                    SpacingAfter = global::OfficeIMO.OfficeLength.Points(0)
                })
            .ToBytes();

        using var pdf = PdfPigDocument.Open(new MemoryStream(bytes));

        Assert.Equal(2, pdf.NumberOfPages);
        Assert.Contains("IntroMarker", pdf.GetPage(1).Text);
        Assert.DoesNotContain("KeepFirst", pdf.GetPage(1).Text);
        Assert.Contains("KeepFirst", pdf.GetPage(2).Text);
        Assert.Contains("KeepLast", pdf.GetPage(2).Text);
    }

    [Fact]
    public void List_KeepTogetherMovesWholeBulletListToNextPage() {
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
            .Paragraph(p => p.Text("IntroMarker"), style: new PdfParagraphStyle {
                SpacingAfter = global::OfficeIMO.OfficeLength.Points(70)
            })
            .Bullets(new[] {
                "KeepListFirst",
                "KeepListMiddle",
                "KeepListLast"
            }, style: new PdfListStyle {
                KeepTogether = true,
                SpacingAfter = global::OfficeIMO.OfficeLength.Points(0)
            })
            .ToBytes();

        using var pdf = PdfPigDocument.Open(new MemoryStream(bytes));

        Assert.Equal(2, pdf.NumberOfPages);
        Assert.Contains("IntroMarker", pdf.GetPage(1).Text);
        Assert.DoesNotContain("KeepListFirst", pdf.GetPage(1).Text);
        Assert.Contains("KeepListFirst", pdf.GetPage(2).Text);
        Assert.Contains("KeepListLast", pdf.GetPage(2).Text);
    }

    [Fact]
    public void List_KeepWithNextMovesListWithFollowingParagraph() {
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
            .Paragraph(p => p.Text("IntroMarker"), style: new PdfParagraphStyle {
                SpacingAfter = global::OfficeIMO.OfficeLength.Points(64)
            })
            .Bullets(new[] {
                "KeepListFirst",
                "KeepListSecond"
            }, style: new PdfListStyle {
                KeepWithNext = true,
                ItemSpacing = global::OfficeIMO.OfficeLength.Points(0),
                SpacingAfter = global::OfficeIMO.OfficeLength.Points(0)
            })
            .Paragraph(p => p.Text("FollowingListBody"))
            .ToBytes();

        using var pdf = PdfPigDocument.Open(new MemoryStream(bytes));

        Assert.Equal(2, pdf.NumberOfPages);
        Assert.Contains("IntroMarker", pdf.GetPage(1).Text);
        Assert.DoesNotContain("KeepListFirst", pdf.GetPage(1).Text);
        Assert.Contains("KeepListFirst", pdf.GetPage(2).Text);
        Assert.Contains("KeepListSecond", pdf.GetPage(2).Text);
        Assert.Contains("FollowingListBody", pdf.GetPage(2).Text);
    }

    [Fact]
    public void Paragraph_KeepWithNextMovesParagraphWithFollowingParagraph() {
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
            .Paragraph(p => p.Text("IntroMarker"), style: new PdfParagraphStyle {
                SpacingAfter = global::OfficeIMO.OfficeLength.Points(70)
            })
            .Paragraph(p => p.Text("KeepWithNextLabel"), style: new PdfParagraphStyle {
                KeepWithNext = true
            })
            .Paragraph(p => p.Text("FollowingBody"))
            .ToBytes();

        using var pdf = PdfPigDocument.Open(new MemoryStream(bytes));

        Assert.Equal(2, pdf.NumberOfPages);
        Assert.Contains("IntroMarker", pdf.GetPage(1).Text);
        Assert.DoesNotContain("KeepWithNextLabel", pdf.GetPage(1).Text);
        Assert.Contains("KeepWithNextLabel", pdf.GetPage(2).Text);
        Assert.Contains("FollowingBody", pdf.GetPage(2).Text);
    }

    [Fact]
    public void Paragraph_KeepWithNextMovesParagraphWithFollowingList() {
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
            .Paragraph(p => p.Text("IntroMarker"), style: new PdfParagraphStyle {
                SpacingAfter = global::OfficeIMO.OfficeLength.Points(70)
            })
            .Paragraph(p => p.Text("KeepWithListLabel"), style: new PdfParagraphStyle {
                KeepWithNext = true
            })
            .Bullets(new[] { "FollowingBullet" })
            .ToBytes();

        using var pdf = PdfPigDocument.Open(new MemoryStream(bytes));

        Assert.Equal(2, pdf.NumberOfPages);
        Assert.Contains("IntroMarker", pdf.GetPage(1).Text);
        Assert.DoesNotContain("KeepWithListLabel", pdf.GetPage(1).Text);
        Assert.Contains("KeepWithListLabel", pdf.GetPage(2).Text);
        Assert.Contains("FollowingBullet", pdf.GetPage(2).Text);
    }

    [Fact]
    public void Paragraph_KeepWithNextMovesParagraphWithFollowingTable() {
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
            .Paragraph(p => p.Text("IntroMarker"), style: new PdfParagraphStyle {
                SpacingAfter = global::OfficeIMO.OfficeLength.Points(70)
            })
            .Paragraph(p => p.Text("KeepWithTableLabel"), style: new PdfParagraphStyle {
                KeepWithNext = true
            })
            .Table(new[] {
                new[] { "FollowingTableCell", "Value" }
            })
            .ToBytes();

        using var pdf = PdfPigDocument.Open(new MemoryStream(bytes));

        Assert.Equal(2, pdf.NumberOfPages);
        Assert.Contains("IntroMarker", pdf.GetPage(1).Text);
        Assert.DoesNotContain("KeepWithTableLabel", pdf.GetPage(1).Text);
        Assert.Contains("KeepWithTableLabel", pdf.GetPage(2).Text);
        Assert.Contains("FollowingTableCell", pdf.GetPage(2).Text);
    }

    [Fact]
    public void Paragraph_KeepWithNextUsesConfiguredTableColumnWidthsForFollowingTable() {
        var style = PdfTableStyles.Minimal();
        style.CellPaddingX = global::OfficeIMO.OfficeLength.Points(4);
        style.CellPaddingY = global::OfficeIMO.OfficeLength.Points(3);
        style.ColumnWidths = new List<global::OfficeIMO.OfficeLength?> { global::OfficeIMO.OfficeLength.Points(44), null };

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
            .Paragraph(p => p.Text("IntroMarker"), style: new PdfParagraphStyle {
                SpacingAfter = global::OfficeIMO.OfficeLength.Points(35)
            })
            .Paragraph(p => p.Text("KeepWithNarrowTableLabel"), style: new PdfParagraphStyle {
                KeepWithNext = true
            })
            .Table(new[] {
                new[] { "aa bb cc dd ee ff gg hh ii jj kk ll", "Value" }
            }, style: style)
            .ToBytes();

        using var pdf = PdfPigDocument.Open(new MemoryStream(bytes));

        Assert.Equal(2, pdf.NumberOfPages);
        Assert.Contains("IntroMarker", pdf.GetPage(1).Text);
        Assert.DoesNotContain("KeepWithNarrowTableLabel", pdf.GetPage(1).Text);
        Assert.Contains("KeepWithNarrowTableLabel", pdf.GetPage(2).Text);
        Assert.Contains("Value", pdf.GetPage(2).Text);
    }

    [Fact]
    public void Paragraph_KeepWithNextMovesParagraphWithFollowingRow() {
        var options = new PdfCreateOptions {
            PageWidth = global::OfficeIMO.OfficeLength.Points(260),
            PageHeight = global::OfficeIMO.OfficeLength.Points(170),
            MarginLeft = global::OfficeIMO.OfficeLength.Points(30),
            MarginRight = global::OfficeIMO.OfficeLength.Points(30),
            MarginTop = global::OfficeIMO.OfficeLength.Points(30),
            MarginBottom = global::OfficeIMO.OfficeLength.Points(30),
            DefaultFont = PdfStandardFont.Helvetica,
            DefaultFontSize = global::OfficeIMO.OfficeLength.Points(10),
            DefaultRowStyle = new PdfRowStyle {
                Gap = global::OfficeIMO.OfficeLength.Points(18)
            }
        };

        byte[] bytes = PdfDocument.Create(options)
            .Compose(document =>
                document.Page(page =>
                    page.Content(content => {
                        content.Column(column => {
                            column.Item().Paragraph(p => p.Text("IntroMarker"), style: new PdfParagraphStyle {
                                SpacingAfter = global::OfficeIMO.OfficeLength.Points(70)
                            });
                            column.Item().Paragraph(p => p.Text("KeepWithRowLabel"), style: new PdfParagraphStyle {
                                KeepWithNext = true
                            });
                        });
                        content.Row(row => row
                            .PercentColumn(50, column => column.Paragraph(p => p.Text("RowLeftBody")))
                            .PercentColumn(50, column => column.Paragraph(p => p.Text("RowRightBody"))));
                    })))
            .ToBytes();

        using var pdf = PdfPigDocument.Open(new MemoryStream(bytes));

        Assert.Equal(2, pdf.NumberOfPages);
        Assert.Contains("IntroMarker", pdf.GetPage(1).Text);
        Assert.DoesNotContain("KeepWithRowLabel", pdf.GetPage(1).Text);
        Assert.Contains("KeepWithRowLabel", pdf.GetPage(2).Text);
        Assert.Contains("RowLeftBody", pdf.GetPage(2).Text);
        Assert.Contains("RowRightBody", pdf.GetPage(2).Text);
    }

    [Fact]
    public void Heading_KeepsWithFollowingParagraphInTopLevelFlow() {
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
            .Paragraph(p => p.Text("IntroMarker"), style: new PdfParagraphStyle {
                SpacingAfter = global::OfficeIMO.OfficeLength.Points(70)
            })
            .H3("SignalHeading")
            .Paragraph(p => p.Text("SignalBody"))
            .ToBytes();

        using var pdf = PdfPigDocument.Open(new MemoryStream(bytes));

        Assert.Equal(2, pdf.NumberOfPages);
        Assert.Contains("IntroMarker", pdf.GetPage(1).Text);
        Assert.DoesNotContain("SignalHeading", pdf.GetPage(1).Text);
        Assert.Contains("SignalHeading", pdf.GetPage(2).Text);
        Assert.Contains("SignalBody", pdf.GetPage(2).Text);
    }

    [Fact]
    public void Heading_KeepWithNextMeasuresFollowingHeadingChainInTopLevelFlow() {
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
            .Paragraph(p => p.Text("IntroMarker"), style: new PdfParagraphStyle {
                SpacingAfter = global::OfficeIMO.OfficeLength.Points(45)
            })
            .H2("ChainTopHeading")
            .H3("ChainSubHeading")
            .Paragraph(p => p.Text("ChainBody"))
            .ToBytes();

        using var pdf = PdfPigDocument.Open(new MemoryStream(bytes));

        Assert.Equal(2, pdf.NumberOfPages);
        Assert.Contains("IntroMarker", pdf.GetPage(1).Text);
        Assert.DoesNotContain("ChainTopHeading", pdf.GetPage(1).Text);
        Assert.Contains("ChainTopHeading", pdf.GetPage(2).Text);
        Assert.Contains("ChainSubHeading", pdf.GetPage(2).Text);
        Assert.Contains("ChainBody", pdf.GetPage(2).Text);
    }

    [Fact]
    public void Heading_KeepWithNextReservesWidowControlledFollowingParagraph() {
        var options = new PdfCreateOptions {
            PageWidth = global::OfficeIMO.OfficeLength.Points(260),
            PageHeight = global::OfficeIMO.OfficeLength.Points(180),
            MarginLeft = global::OfficeIMO.OfficeLength.Points(30),
            MarginRight = global::OfficeIMO.OfficeLength.Points(30),
            MarginTop = global::OfficeIMO.OfficeLength.Points(30),
            MarginBottom = global::OfficeIMO.OfficeLength.Points(30),
            DefaultFont = PdfStandardFont.Helvetica,
            DefaultFontSize = global::OfficeIMO.OfficeLength.Points(10)
        };

        var heading2 = new PdfHeadingStyle {
            FontSize = global::OfficeIMO.OfficeLength.Points(12),
            LineHeight = 1.2,
            SpacingAfter = global::OfficeIMO.OfficeLength.Points(2),
            KeepWithNext = true
        };
        var heading3 = new PdfHeadingStyle {
            FontSize = global::OfficeIMO.OfficeLength.Points(11),
            LineHeight = 1.2,
            SpacingAfter = global::OfficeIMO.OfficeLength.Points(2),
            KeepWithNext = true
        };

        byte[] bytes = PdfDocument.Create(options)
            .Paragraph(p => p.Text("IntroMarker"), style: new PdfParagraphStyle {
                SpacingAfter = global::OfficeIMO.OfficeLength.Points(50)
            })
            .H2("WidowChainHeading", style: heading2)
            .H3("WidowChainSubHeading", style: heading3)
            .Paragraph(p => p.Text("Following body text wraps enough to require multiple rendered lines in this narrow frame."), style: new PdfParagraphStyle {
                WidowControl = true,
                SpacingAfter = global::OfficeIMO.OfficeLength.Points(0)
            })
            .ToBytes();

        using var pdf = PdfPigDocument.Open(new MemoryStream(bytes));

        Assert.Equal(2, pdf.NumberOfPages);
        Assert.Contains("IntroMarker", pdf.GetPage(1).Text);
        Assert.DoesNotContain("WidowChainHeading", pdf.GetPage(1).Text);
        Assert.Contains("WidowChainHeading", pdf.GetPage(2).Text);
        Assert.Contains("WidowChainSubHeading", pdf.GetPage(2).Text);
        Assert.Contains("Following body text", pdf.GetPage(2).Text);
    }

    [Fact]
    public void Heading_KeepWithNextMeasuresRuleAndFollowingHeadingChainInTopLevelFlow() {
        var options = new PdfCreateOptions {
            PageWidth = global::OfficeIMO.OfficeLength.Points(260),
            PageHeight = global::OfficeIMO.OfficeLength.Points(170),
            MarginLeft = global::OfficeIMO.OfficeLength.Points(30),
            MarginRight = global::OfficeIMO.OfficeLength.Points(30),
            MarginTop = global::OfficeIMO.OfficeLength.Points(30),
            MarginBottom = global::OfficeIMO.OfficeLength.Points(30),
            DefaultFont = PdfStandardFont.Helvetica,
            DefaultFontSize = OfficeLength.Points(10),
            // The page geometry is sized for the built-in heading metrics, not the document theme.
            DefaultHeadingStyles = new PdfHeadingStyles()
        };

        byte[] bytes = PdfDocument.Create(options)
            .Paragraph(p => p.Text("IntroMarker"), style: new PdfParagraphStyle {
                SpacingAfter = global::OfficeIMO.OfficeLength.Points(44)
            })
            .H2("RuleChainTopHeading")
            .HR(style: new PdfHorizontalRuleStyle {
                Thickness = global::OfficeIMO.OfficeLength.Points(0.5),
                SpacingBefore = global::OfficeIMO.OfficeLength.Points(0),
                SpacingAfter = global::OfficeIMO.OfficeLength.Points(4),
                KeepWithNext = true
            })
            .H3("RuleChainSubHeading")
            .Paragraph(p => p.Text("RuleChainBody"))
            .ToBytes();

        using var pdf = PdfPigDocument.Open(new MemoryStream(bytes));

        Assert.Equal(2, pdf.NumberOfPages);
        Assert.Contains("IntroMarker", pdf.GetPage(1).Text);
        Assert.DoesNotContain("RuleChainTopHeading", pdf.GetPage(1).Text);
        Assert.Contains("RuleChainTopHeading", pdf.GetPage(2).Text);
        Assert.Contains("RuleChainSubHeading", pdf.GetPage(2).Text);
        Assert.Contains("RuleChainBody", pdf.GetPage(2).Text);
    }

    [Fact]
    public void Heading_KeepWithNextSkipsBookmarkMarkersInFollowingHeadingChain() {
        var options = new PdfCreateOptions {
            PageWidth = global::OfficeIMO.OfficeLength.Points(260),
            PageHeight = global::OfficeIMO.OfficeLength.Points(170),
            MarginLeft = global::OfficeIMO.OfficeLength.Points(30),
            MarginRight = global::OfficeIMO.OfficeLength.Points(30),
            MarginTop = global::OfficeIMO.OfficeLength.Points(30),
            MarginBottom = global::OfficeIMO.OfficeLength.Points(30),
            DefaultFont = PdfStandardFont.Helvetica,
            DefaultFontSize = OfficeLength.Points(10),
            // The page geometry is sized for the built-in heading metrics, not the document theme.
            DefaultHeadingStyles = new PdfHeadingStyles()
        };

        PdfDocument document = PdfDocument.Create(options)
            .Paragraph(p => p.Text("IntroMarker"), style: new PdfParagraphStyle {
                SpacingAfter = global::OfficeIMO.OfficeLength.Points(45)
            })
            .Bookmark("bookmark-chain-top-heading")
            .H2("BookmarkChainTopHeading");
        for (int index = 0; index < 300; index++) {
            document.Bookmark("bookmark-chain-marker-" + index.ToString(CultureInfo.InvariantCulture));
        }

        byte[] bytes = document
            .H3("BookmarkChainSubHeading")
            .Paragraph(p => p.Text("BookmarkChainBody"))
            .ToBytes();

        using var pdf = PdfPigDocument.Open(new MemoryStream(bytes));

        Assert.Equal(2, pdf.NumberOfPages);
        Assert.Contains("IntroMarker", pdf.GetPage(1).Text);
        Assert.DoesNotContain("BookmarkChainTopHeading", pdf.GetPage(1).Text);
        Assert.Contains("BookmarkChainTopHeading", pdf.GetPage(2).Text);
        Assert.Contains("BookmarkChainSubHeading", pdf.GetPage(2).Text);
        Assert.Contains("BookmarkChainBody", pdf.GetPage(2).Text);
        var destination = Assert.Single(PdfInspector.Inspect(bytes).NamedDestinations,
            item => item.Name == "bookmark-chain-top-heading");
        Assert.Equal(2, destination.PageNumber);
    }

    [Fact]
    public void Heading_KeepsWithFollowingPanelInTopLevelFlow() {
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
            .Paragraph(p => p.Text("IntroMarker"), style: new PdfParagraphStyle {
                SpacingAfter = global::OfficeIMO.OfficeLength.Points(70)
            })
            .H3("PanelSignalHeading")
            .PanelParagraph(p => p.Text("PanelSignalBody"), new PdfPanelStyle {
                PaddingY = global::OfficeIMO.OfficeLength.Points(5),
                SpacingAfter = global::OfficeIMO.OfficeLength.Points(0)
            })
            .ToBytes();

        using var pdf = PdfPigDocument.Open(new MemoryStream(bytes));

        Assert.Equal(2, pdf.NumberOfPages);
        Assert.Contains("IntroMarker", pdf.GetPage(1).Text);
        Assert.DoesNotContain("PanelSignalHeading", pdf.GetPage(1).Text);
        Assert.Contains("PanelSignalHeading", pdf.GetPage(2).Text);
        Assert.Contains("PanelSignalBody", pdf.GetPage(2).Text);
    }

    [Fact]
    public void PanelParagraph_KeepWithNextMovesPanelWithFollowingParagraph() {
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
            .Paragraph(p => p.Text("IntroMarker"), style: new PdfParagraphStyle {
                SpacingAfter = global::OfficeIMO.OfficeLength.Points(61)
            })
            .PanelParagraph(p => p.Text("PanelKeepWithNext"), new PdfPanelStyle {
                KeepWithNext = true,
                PaddingY = global::OfficeIMO.OfficeLength.Points(5),
                SpacingAfter = global::OfficeIMO.OfficeLength.Points(0)
            })
            .Paragraph(p => p.Text("FollowingPanelBody"))
            .ToBytes();

        using var pdf = PdfPigDocument.Open(new MemoryStream(bytes));

        Assert.Equal(2, pdf.NumberOfPages);
        Assert.Contains("IntroMarker", pdf.GetPage(1).Text);
        Assert.DoesNotContain("PanelKeepWithNext", pdf.GetPage(1).Text);
        Assert.Contains("PanelKeepWithNext", pdf.GetPage(2).Text);
        Assert.Contains("FollowingPanelBody", pdf.GetPage(2).Text);
    }

    [Fact]
    public void PanelParagraph_KeepWithNextMeasuresFollowingHeadingChain() {
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
            .Paragraph(p => p.Text("IntroMarker"), style: new PdfParagraphStyle {
                SpacingAfter = global::OfficeIMO.OfficeLength.Points(56)
            })
            .PanelParagraph(p => p.Text("PanelChainKeepWithNext"), new PdfPanelStyle {
                KeepWithNext = true,
                PaddingY = global::OfficeIMO.OfficeLength.Points(5),
                SpacingAfter = global::OfficeIMO.OfficeLength.Points(0)
            })
            .H3("FollowingPanelHeading")
            .Paragraph(p => p.Text("FollowingPanelChainBody"))
            .ToBytes();

        using var pdf = PdfPigDocument.Open(new MemoryStream(bytes));

        Assert.Equal(2, pdf.NumberOfPages);
        Assert.Contains("IntroMarker", pdf.GetPage(1).Text);
        Assert.DoesNotContain("PanelChainKeepWithNext", pdf.GetPage(1).Text);
        Assert.DoesNotContain("FollowingPanelHeading", pdf.GetPage(1).Text);
        Assert.DoesNotContain("FollowingPanelChainBody", pdf.GetPage(1).Text);
        Assert.Contains("PanelChainKeepWithNext", pdf.GetPage(2).Text);
        Assert.Contains("FollowingPanelHeading", pdf.GetPage(2).Text);
        Assert.Contains("FollowingPanelChainBody", pdf.GetPage(2).Text);
    }

    [Fact]
    public void Heading_KeepsWithFollowingTableInTopLevelFlow() {
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
            .Paragraph(p => p.Text("IntroMarker"), style: new PdfParagraphStyle {
                SpacingAfter = global::OfficeIMO.OfficeLength.Points(70)
            })
            .H3("TableSignalHeading")
            .Table(new[] {
                new[] { "TableSignalCell", "Value" }
            })
            .ToBytes();

        using var pdf = PdfPigDocument.Open(new MemoryStream(bytes));

        Assert.Equal(2, pdf.NumberOfPages);
        Assert.Contains("IntroMarker", pdf.GetPage(1).Text);
        Assert.DoesNotContain("TableSignalHeading", pdf.GetPage(1).Text);
        Assert.Contains("TableSignalHeading", pdf.GetPage(2).Text);
        Assert.Contains("TableSignalCell", pdf.GetPage(2).Text);
    }

    [Fact]
    public void Heading_KeepsWithFollowingRowInTopLevelFlow() {
        var options = new PdfCreateOptions {
            PageWidth = global::OfficeIMO.OfficeLength.Points(260),
            PageHeight = global::OfficeIMO.OfficeLength.Points(170),
            MarginLeft = global::OfficeIMO.OfficeLength.Points(30),
            MarginRight = global::OfficeIMO.OfficeLength.Points(30),
            MarginTop = global::OfficeIMO.OfficeLength.Points(30),
            MarginBottom = global::OfficeIMO.OfficeLength.Points(30),
            DefaultFont = PdfStandardFont.Helvetica,
            DefaultFontSize = global::OfficeIMO.OfficeLength.Points(10),
            DefaultRowStyle = new PdfRowStyle {
                Gap = global::OfficeIMO.OfficeLength.Points(18)
            }
        };

        byte[] bytes = PdfDocument.Create(options)
            .Compose(document =>
                document.Page(page =>
                    page.Content(content => {
                        content.Column(column => {
                            column.Item().Paragraph(p => p.Text("IntroMarker"), style: new PdfParagraphStyle {
                                SpacingAfter = global::OfficeIMO.OfficeLength.Points(70)
                            });
                            column.Item().H3("RowSignalHeading");
                        });
                        content.Row(row => row
                            .PercentColumn(50, column => column.Paragraph(p => p.Text("RowSignalLeft")))
                            .PercentColumn(50, column => column.Paragraph(p => p.Text("RowSignalRight"))));
                    })))
            .ToBytes();

        using var pdf = PdfPigDocument.Open(new MemoryStream(bytes));

        Assert.Equal(2, pdf.NumberOfPages);
        Assert.Contains("IntroMarker", pdf.GetPage(1).Text);
        Assert.DoesNotContain("RowSignalHeading", pdf.GetPage(1).Text);
        Assert.Contains("RowSignalHeading", pdf.GetPage(2).Text);
        Assert.Contains("RowSignalLeft", pdf.GetPage(2).Text);
        Assert.Contains("RowSignalRight", pdf.GetPage(2).Text);
    }

    [Fact]
    public void Paragraph_WidowControlAvoidsSingleLineAtPageBottom() {
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
            .Paragraph(p => p.Text("IntroMarker"), style: new PdfParagraphStyle {
                SpacingAfter = global::OfficeIMO.OfficeLength.Points(70)
            })
            .Paragraph(p => p
                .Text("WidowFirst")
                .LineBreak()
                .Text("WidowSecond")
                .LineBreak()
                .Text("WidowThird"), style: new PdfParagraphStyle {
                    WidowControl = true,
                    SpacingAfter = global::OfficeIMO.OfficeLength.Points(0)
                })
            .ToBytes();

        using var pdf = PdfPigDocument.Open(new MemoryStream(bytes));

        Assert.Equal(2, pdf.NumberOfPages);
        Assert.Contains("IntroMarker", pdf.GetPage(1).Text);
        Assert.DoesNotContain("WidowFirst", pdf.GetPage(1).Text);
        Assert.Contains("WidowFirst", pdf.GetPage(2).Text);
        Assert.Contains("WidowThird", pdf.GetPage(2).Text);
    }


}
