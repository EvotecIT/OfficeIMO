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
    public void LinkAnnotations_RenderForParagraphHeadingAndTableCells() {
        byte[] bytes = PdfDocument.Create(new PdfOptions {
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
        Assert.Equal(3, CountOccurrences(pdf, "/Subtype /Link"));
        Assert.Equal(3, CountOccurrences(pdf, "/S /URI"));
        Assert.Equal(1, CountOccurrences(pdf, "/URI (https://evotec.xyz/heading)"));
        Assert.Equal(1, CountOccurrences(pdf, "/URI (https://evotec.xyz/paragraph)"));
        Assert.Equal(1, CountOccurrences(pdf, "/URI (https://evotec.xyz/table)"));
        Assert.Equal(3, CountOccurrences(pdf, "/Contents ("));
        Assert.Equal(1, CountOccurrences(pdf, "/Contents (Heading \\(metadata\\))"));
        Assert.Equal(1, CountOccurrences(pdf, "/Contents (Paragraph \\\\ metadata)"));
        Assert.Equal(1, CountOccurrences(pdf, "/Contents (Open)"));

        var rectangles = ExtractLinkRectangles(pdf);
        Assert.Equal(3, rectangles.Count);
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

        byte[] bytes = PdfDocument.Create(options)
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
        byte[] bytes = PdfDocument.Create(new PdfOptions {
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

        byte[] bytes = PdfDocument.Create(options)
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

        byte[] bytes = PdfDocument.Create(options)
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
        byte[] bytes = PdfDocument.Create(new PdfOptions {
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

        byte[] bytes = PdfDocument.Create(options)
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
        byte[] bytes = PdfDocument.Create(new PdfOptions {
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

        byte[] bytes = PdfDocument.Create(new PdfOptions {
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

        byte[] bytes = PdfDocument.Create(options)
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

        byte[] bytes = PdfDocument.Create(options)
            .Compose(document =>
                document.Page(page =>
                    page.Content(content =>
                        content.Row(row =>
                            row.Column(100, column =>
                                column.H3("ColumnHead", PdfAlign.Right, linkUri: "https://evotec.xyz/right-heading", linkContents: "Right heading"))))))
            .ToBytes();

        string pdfText = Encoding.ASCII.GetString(bytes);
        var rect = Assert.Single(ExtractLinkRectangles(pdfText));

        using var pdf = PdfPigDocument.Open(new MemoryStream(bytes));
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
        byte[] bytes = PdfDocument.Create(new PdfOptions {
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
        byte[] bytes = PdfDocument.Create(new PdfOptions {
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
    public void Tables_DeclareRichCellRunFontsBeforeRendering() {
        byte[] bytes = PdfDocument.Create(new PdfOptions {
                DefaultFont = PdfStandardFont.Helvetica,
                DefaultFontSize = 11
            })
            .Table(new[] {
                new[] {
                    PdfTableCell.RichTextCell(new[] { TextRun.Normal("Direct table Times", font: PdfStandardFont.TimesRoman) })
                }
            })
            .Compose(document =>
                document.Page(page =>
                    page.Content(content =>
                        content.Row(row =>
                            row.Column(100, column =>
                                column.Table(new[] {
                                    new[] {
                                        PdfTableCell.RichTextCell(new[] { TextRun.Normal("Column table Courier", font: PdfStandardFont.Courier) })
                                    }
                                }))))))
            .ToBytes();

        string pdf = Encoding.ASCII.GetString(bytes);

        Assert.Contains("/BaseFont /Times-Roman", pdf, StringComparison.Ordinal);
        Assert.Contains("/BaseFont /Courier", pdf, StringComparison.Ordinal);
    }

    [Fact]
    public void Bookmark_RejectsInvalidNamesAndDuplicateNames() {
        Assert.Throws<ArgumentNullException>(() =>
            PdfDocument.Create().Bookmark(null!));

        Assert.Throws<ArgumentException>(() =>
            PdfDocument.Create().Bookmark(" "));

        var duplicateException = Assert.Throws<ArgumentException>(() =>
            PdfDocument.Create()
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
            PdfDocument.Create().Paragraph(p => p.LinkToBookmark("Jump", null!)));

        Assert.Throws<ArgumentException>(() =>
            PdfDocument.Create().Paragraph(p => p.LinkToBookmark("Jump", " ")));

        Assert.Throws<ArgumentException>(() =>
            PdfDocument.Create().Paragraph(p => p.LinkToBookmark("", "Intro")));

        Assert.Throws<ArgumentException>(() =>
            PdfDocument.Create().Paragraph(p => p.LinkToBookmark("Jump", "Intro", contents: " ")));

        var missingTargetException = Assert.Throws<ArgumentException>(() =>
            PdfDocument.Create()
                .Paragraph(p => p.LinkToBookmark("Jump to missing bookmark", "MissingBookmark"))
                .ToBytes());

        Assert.Contains("PDF bookmark link target 'MissingBookmark' was not found.", missingTargetException.Message, StringComparison.Ordinal);

        var missingHeadingTargetException = Assert.Throws<ArgumentException>(() =>
            PdfDocument.Create()
                .H1("Jump to missing bookmark", linkDestinationName: "MissingHeadingBookmark")
                .ToBytes());

        Assert.Contains("PDF bookmark link target 'MissingHeadingBookmark' was not found.", missingHeadingTargetException.Message, StringComparison.Ordinal);
    }

    [Fact]
    public void Bookmark_RendersNamedDestinationNameTreeAndInspectorReadsIt() {
        byte[] bytes = PdfDocument.Create(new PdfOptions {
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
        byte[] bytes = PdfDocument.Create(new PdfOptions {
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
        byte[] bytes = PdfDocument.Create(new PdfOptions {
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
        byte[] bytes = PdfDocument.Create(new PdfOptions {
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


}
