using System;
using System.IO;
using System.Linq;
using OfficeIMO.Drawing;
using OfficeIMO.Pdf;
using Xunit;

namespace OfficeIMO.Tests.Pdf;

public partial class PdfInspectorTests {
    [Fact]
    public void Preflight_ReportsAnnotationsWithoutBlockingRewrite() {
        byte[] bytes = BuildAnnotatedPdf();

        PdfDocumentPreflight report = PdfInspector.Preflight(bytes);

        Assert.True(report.CanRead);
        Assert.True(report.CanRewrite);
        Assert.True(report.Probe.HasAnnotations);
        Assert.NotNull(report.DocumentInfo);
        Assert.True(report.DocumentInfo!.HasAnnotations);
        Assert.Empty(report.Diagnostics);
        Assert.Empty(report.ReadBlockers);
        Assert.Empty(report.RewriteBlockers);
        Assert.False(report.HasReadBlocker(PdfReadBlockerKind.ParserUnsupported));
        Assert.False(report.HasRewriteBlocker(PdfRewriteBlockerKind.Outlines));
    }

    [Fact]
    public void Inspect_ReadsSimpleUriLinkAnnotationsWithContents() {
        byte[] bytes = BuildAnnotatedPdf();

        PdfDocumentInfo info = PdfInspector.Inspect(bytes);

        Assert.True(info.HasAnnotations);
        Assert.True(info.HasLinkAnnotations);
        Assert.Equal(1, info.LinkAnnotationCount);
        Assert.Single(info.LinkAnnotations);
        Assert.Equal(1, info.LinkUriCount);
        Assert.Equal(new[] { "https://evotec.xyz" }, info.LinkUris);
        Assert.Same(info.LinkAnnotations[0], Assert.Single(info.LinkAnnotationsByUri["https://evotec.xyz"]));
        Assert.Same(info.LinkAnnotations[0], Assert.Single(info.GetLinkAnnotationsByUri("https://evotec.xyz")));
        Assert.Empty(info.GetLinkAnnotationsByUri("https://evotec.xyz/missing"));
        Assert.Single(info.Pages);
        Assert.Single(info.Pages[0].LinkAnnotations);
        Assert.All(info.LinkAnnotations, link => Assert.Equal(1, link.PageNumber));
        foreach (var link in info.Pages[0].LinkAnnotations) {
            Assert.Equal(1, link.PageNumber);
            Assert.Equal("https://evotec.xyz", link.Uri);
            Assert.Equal("OfficeIMO link", link.Contents);
            Assert.True(link.Width > 0);
            Assert.True(link.Height > 0);
            Assert.InRange(link.X1, 0, info.Pages[0].Width);
            Assert.InRange(link.X2, 0, info.Pages[0].Width);
            Assert.InRange(link.Y1, 0, info.Pages[0].Height);
            Assert.InRange(link.Y2, 0, info.Pages[0].Height);
        }

        PdfReadDocument document = PdfReadDocument.Open(bytes);
        var pageLinks = document.Pages[0].GetLinkAnnotations();
        Assert.Equal(info.Pages[0].LinkAnnotations.Count, pageLinks.Count);
        Assert.Equal(info.Pages[0].LinkAnnotations[0].Uri, pageLinks[0].Uri);
        Assert.Equal(info.Pages[0].LinkAnnotations[0].Contents, pageLinks[0].Contents);
    }

    [Fact]
    public void Inspect_ReadsGeneratedWrappedAndRowColumnHeadingLinks() {
        var wrappedOptions = new PdfOptions {
            PageWidth = 140,
            PageHeight = 160,
            MarginLeft = 30,
            MarginRight = 30,
            MarginTop = 30,
            MarginBottom = 30,
            DefaultFont = PdfStandardFont.Helvetica,
            DefaultFontSize = 10
        };

        byte[] wrappedBytes = PdfDocument.Create(wrappedOptions)
            .H3("WWWWWWWW", linkUri: "https://evotec.xyz/wrapped-heading", linkContents: "Wrapped heading")
            .ToBytes();

        PdfDocumentInfo wrappedInfo = PdfInspector.Inspect(wrappedBytes);

        Assert.True(wrappedInfo.LinkAnnotationCount > 1);
        Assert.Equal(wrappedInfo.LinkAnnotationCount, wrappedInfo.Pages[0].LinkAnnotations.Count);
        Assert.All(wrappedInfo.LinkAnnotations, link => {
            Assert.Equal(1, link.PageNumber);
            Assert.Equal("https://evotec.xyz/wrapped-heading", link.Uri);
            Assert.Equal("Wrapped heading", link.Contents);
            Assert.True(link.Width > 0);
            Assert.True(link.Height > 0);
        });

        byte[] rowBytes = PdfDocument.Create(new PdfOptions {
                DefaultFont = PdfStandardFont.Helvetica,
                DefaultFontSize = 10
            })
            .Compose(document =>
                document.Page(page =>
                    page.Content(content =>
                        content.Row(row =>
                            row.Column(100, column =>
                                column.H3("ColumnHead", PdfAlign.Right, linkUri: "https://evotec.xyz/right-heading", linkContents: "Right heading"))))))
            .ToBytes();

        PdfDocumentInfo rowInfo = PdfInspector.Inspect(rowBytes);
        PdfLinkAnnotation rowLink = Assert.Single(rowInfo.LinkAnnotations);

        Assert.Equal(1, rowLink.PageNumber);
        Assert.Equal("https://evotec.xyz/right-heading", rowLink.Uri);
        Assert.Equal("Right heading", rowLink.Contents);
        Assert.True(rowLink.Width > 0);
        Assert.True(rowLink.Height > 0);
        Assert.Equal(rowLink.Uri, rowInfo.LinkUris.Single());
    }

    [Fact]
    public void Inspect_ReadsGeneratedImageLinks() {
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
            .Image(CreateMinimalRgbPng(), 80, 40, PdfAlign.Center, fit: OfficeImageFit.Contain, linkUri: "https://evotec.xyz/image", linkContents: "Image metadata")
            .ToBytes();

        PdfDocumentInfo info = PdfInspector.Inspect(bytes);
        PdfLinkAnnotation link = Assert.Single(info.LinkAnnotations);

        Assert.Equal(1, link.PageNumber);
        Assert.Equal("https://evotec.xyz/image", link.Uri);
        Assert.Equal("Image metadata", link.Contents);
        Assert.InRange(link.X1, 89.5, 90.5);
        Assert.InRange(link.X2, 129.5, 130.5);
        Assert.InRange(link.Y1, 109.5, 110.5);
        Assert.InRange(link.Y2, 149.5, 150.5);
        Assert.Equal(link.Uri, info.LinkUris.Single());
    }

    [Fact]
    public void Inspect_ReadsGeneratedShapeAndDrawingLinks() {
        var drawing = new OfficeDrawing(60, 30)
            .AddShape(OfficeShape.Rectangle(60, 30), 0, 0);

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
            .Shape(OfficeShape.Rectangle(40, 20), PdfAlign.Right, linkUri: "https://evotec.xyz/shape", linkContents: "Shape metadata")
            .Drawing(drawing, PdfAlign.Center, spacingBefore: 6, linkUri: "https://evotec.xyz/drawing", linkContents: "Drawing metadata")
            .ToBytes();

        PdfDocumentInfo info = PdfInspector.Inspect(bytes);

        Assert.Equal(2, info.LinkAnnotationCount);
        Assert.Equal(2, info.LinkUriCount);
        Assert.Equal(new[] { "https://evotec.xyz/shape", "https://evotec.xyz/drawing" }, info.LinkUris);

        PdfLinkAnnotation shapeLink = Assert.Single(info.LinkAnnotations, link => link.Uri == "https://evotec.xyz/shape");
        Assert.Equal(1, shapeLink.PageNumber);
        Assert.Equal("Shape metadata", shapeLink.Contents);
        Assert.InRange(shapeLink.X1, 149.5, 150.5);
        Assert.InRange(shapeLink.X2, 189.5, 190.5);
        Assert.InRange(shapeLink.Y1, 169.5, 170.5);
        Assert.InRange(shapeLink.Y2, 189.5, 190.5);

        PdfLinkAnnotation drawingLink = Assert.Single(info.LinkAnnotations, link => link.Uri == "https://evotec.xyz/drawing");
        Assert.Equal(1, drawingLink.PageNumber);
        Assert.Equal("Drawing metadata", drawingLink.Contents);
        Assert.InRange(drawingLink.X1, 79.5, 80.5);
        Assert.InRange(drawingLink.X2, 139.5, 140.5);
        Assert.InRange(drawingLink.Y1, 133.5, 134.5);
        Assert.InRange(drawingLink.Y2, 163.5, 164.5);
    }

    [Fact]
    public void Inspect_ReadsGeneratedConvenienceVectorLinks() {
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
            .Rectangle(40, 20, align: PdfAlign.Center, linkUri: "https://evotec.xyz/rectangle", linkContents: "Rectangle metadata")
            .ToBytes();

        PdfDocumentInfo info = PdfInspector.Inspect(bytes);
        PdfLinkAnnotation link = Assert.Single(info.LinkAnnotations);

        Assert.Equal(1, info.LinkAnnotationCount);
        Assert.Equal(1, info.LinkUriCount);
        Assert.Equal("https://evotec.xyz/rectangle", link.Uri);
        Assert.Equal("Rectangle metadata", link.Contents);
        Assert.InRange(link.X1, 89.5, 90.5);
        Assert.InRange(link.X2, 129.5, 130.5);
        Assert.InRange(link.Y1, 129.5, 130.5);
        Assert.InRange(link.Y2, 149.5, 150.5);
        Assert.Equal(link.Uri, info.LinkUris.Single());
    }

    [Fact]
    public void Preflight_AllowsGeneratedBookmarkNamedDestinationsForReadAndRewrite() {
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
            .Bookmark("Intro")
            .Paragraph(p => p.Text("Bookmarked paragraph."))
            .ToBytes();

        PdfDocumentPreflight report = PdfInspector.Preflight(bytes);

        Assert.True(report.CanRead);
        Assert.True(report.CanRewrite);
        Assert.True(report.Probe.HasNamedDestinations);
        Assert.True(report.Probe.HasCatalogNameTrees);
        Assert.NotNull(report.DocumentInfo);
        Assert.True(report.DocumentInfo!.HasNamedDestinations);
        Assert.True(report.DocumentInfo.HasCatalogNameTrees);
        AssertNamedDestination(report.DocumentInfo, "Intro", 1, 150);
        Assert.Empty(report.Diagnostics);
        Assert.Empty(report.ReadBlockers);
        Assert.Empty(report.RewriteBlockers);
        Assert.False(report.HasRewriteBlocker(PdfRewriteBlockerKind.NamedDestinations));
    }

    [Fact]
    public void Inspect_ReadsGeneratedBookmarkLinksAsNamedDestinationLinks() {
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
            .Paragraph(p => p.LinkToBookmark("Details", "Details", contents: "Internal jump"))
            .Spacer(20)
            .Bookmark("Details")
            .H2("Details")
            .Paragraph(p => p.Text("Destination paragraph."))
            .ToBytes();

        PdfDocumentInfo info = PdfInspector.Inspect(bytes);
        PdfLinkAnnotation link = Assert.Single(info.LinkAnnotations);

        Assert.True(info.HasLinkAnnotations);
        Assert.Equal(1, info.LinkAnnotationCount);
        Assert.Equal(0, info.LinkUriCount);
        Assert.Empty(info.LinkUris);
        Assert.Equal(1, info.LinkDestinationCount);
        Assert.Equal(new[] { "Details" }, info.LinkDestinationNames);
        Assert.Same(link, Assert.Single(info.LinkAnnotationsByDestinationName["Details"]));
        Assert.Same(link, Assert.Single(info.GetLinkAnnotationsByDestinationName("Details")));
        Assert.Empty(info.GetLinkAnnotationsByDestinationName("Missing"));
        Assert.True(link.IsNamedDestinationLink);
        Assert.False(link.IsUriLink);
        Assert.Null(link.Uri);
        Assert.Equal("Details", link.DestinationName);
        Assert.Equal("Internal jump", link.Contents);
        Assert.Equal(1, link.PageNumber);
        Assert.True(link.Width > 0);
        Assert.True(link.Height > 0);
    }

    [Fact]
    public void Inspect_SummarizesDirectDestinationLinkTargetPages() {
        PdfDocumentInfo info = PdfInspector.Inspect(BuildDirectDestinationLinkPdf());
        PdfLinkAnnotation link = Assert.Single(info.LinkAnnotations);

        Assert.True(info.HasLinkAnnotations);
        Assert.Equal(1, info.LinkAnnotationCount);
        Assert.Equal(0, info.LinkUriCount);
        Assert.Empty(info.LinkUris);
        Assert.Equal(0, info.LinkDestinationCount);
        Assert.Empty(info.LinkDestinationNames);
        Assert.Equal(1, info.LinkDestinationPageNumberCount);
        Assert.Equal(new[] { 1 }, info.LinkDestinationPageNumbers);
        Assert.Same(link, Assert.Single(info.LinkAnnotationsByDestinationPageNumber[1]));
        Assert.Same(link, Assert.Single(info.GetLinkAnnotationsByDestinationPageNumber(1)));
        Assert.Empty(info.GetLinkAnnotationsByDestinationPageNumber(2));
        Assert.Throws<ArgumentOutOfRangeException>(() => info.GetLinkAnnotationsByDestinationPageNumber(0));
        Assert.True(link.IsInternalDestinationLink);
        Assert.False(link.IsUriLink);
        Assert.False(link.IsNamedDestinationLink);
        Assert.Equal(1, link.PageNumber);
        Assert.Equal(1, link.DestinationPageNumber);
        Assert.Equal(PdfOpenActionDestinationMode.FitRectangle, link.DestinationMode);
        Assert.Equal(10D, link.DestinationLeft);
        Assert.Equal(20D, link.DestinationBottom);
        Assert.Equal(90D, link.DestinationRight);
        Assert.Equal(144D, link.DestinationTop);
    }

    [Fact]
    public void Inspect_ReadsNamedActionLinkAnnotations() {
        byte[] bytes = BuildNamedActionLinkPdf();

        PdfDocumentInfo info = PdfInspector.Inspect(bytes);
        PdfLinkAnnotation link = Assert.Single(info.LinkAnnotations);

        Assert.True(info.HasLinkAnnotations);
        Assert.Equal(1, info.LinkAnnotationCount);
        Assert.Equal(0, info.LinkUriCount);
        Assert.Empty(info.LinkUris);
        Assert.Equal(0, info.LinkDestinationCount);
        Assert.Empty(info.LinkDestinationNames);
        Assert.Equal(0, info.LinkDestinationPageNumberCount);
        Assert.Empty(info.LinkDestinationPageNumbers);
        Assert.Equal(1, info.LinkNamedActionCount);
        Assert.Equal(new[] { "NextPage" }, info.LinkNamedActions);
        Assert.Same(link, Assert.Single(info.LinkAnnotationsByNamedAction["NextPage"]));
        Assert.Same(link, Assert.Single(info.GetLinkAnnotationsByNamedAction("NextPage")));
        Assert.Empty(info.GetLinkAnnotationsByNamedAction("PrevPage"));
        Assert.True(link.IsNamedActionLink);
        Assert.False(link.IsUriLink);
        Assert.False(link.IsNamedDestinationLink);
        Assert.False(link.IsInternalDestinationLink);
        Assert.Null(link.Uri);
        Assert.Null(link.DestinationName);
        Assert.Null(link.DestinationPageNumber);
        Assert.Equal("NextPage", link.NamedAction);
        Assert.Equal("Next page action", link.Contents);
        Assert.Equal(1, link.PageNumber);

        PdfReadDocument document = PdfReadDocument.Open(bytes);
        PdfLinkAnnotation pageLink = Assert.Single(document.Pages[0].GetLinkAnnotations());
        Assert.True(pageLink.IsNamedActionLink);
        Assert.Equal("NextPage", pageLink.NamedAction);
    }

    [Fact]
    public void Inspect_ReadsRemoteGoToLinkAnnotations() {
        byte[] bytes = BuildRemoteGoToLinkPdf();

        PdfDocumentInfo info = PdfInspector.Inspect(bytes);
        PdfLinkAnnotation link = Assert.Single(info.LinkAnnotations);

        Assert.True(info.HasLinkAnnotations);
        Assert.Equal(1, info.LinkAnnotationCount);
        Assert.Equal(0, info.LinkUriCount);
        Assert.Empty(info.LinkUris);
        Assert.Equal(0, info.LinkDestinationCount);
        Assert.Empty(info.LinkDestinationNames);
        Assert.Equal(0, info.LinkDestinationPageNumberCount);
        Assert.Empty(info.LinkDestinationPageNumbers);
        Assert.Equal(0, info.LinkNamedActionCount);
        Assert.Empty(info.LinkNamedActions);
        Assert.Equal(1, info.LinkRemoteFileCount);
        Assert.Equal(new[] { "remote-report.pdf" }, info.LinkRemoteFiles);
        Assert.Same(link, Assert.Single(info.LinkAnnotationsByRemoteFile["remote-report.pdf"]));
        Assert.Same(link, Assert.Single(info.GetLinkAnnotationsByRemoteFile("remote-report.pdf")));
        Assert.Empty(info.GetLinkAnnotationsByRemoteFile("missing.pdf"));
        Assert.True(link.IsRemoteGoToLink);
        Assert.False(link.IsUriLink);
        Assert.False(link.IsNamedDestinationLink);
        Assert.False(link.IsInternalDestinationLink);
        Assert.False(link.IsNamedActionLink);
        Assert.Null(link.Uri);
        Assert.Null(link.DestinationName);
        Assert.Null(link.DestinationPageNumber);
        Assert.Null(link.NamedAction);
        Assert.Equal("remote-report.pdf", link.RemoteFile);
        Assert.Null(link.RemoteDestinationName);
        Assert.Equal(2, link.RemoteDestinationPageNumber);
        Assert.Equal(PdfOpenActionDestinationMode.FitHorizontal, link.RemoteDestinationMode);
        Assert.Equal(144D, link.RemoteDestinationTop);
        Assert.Equal("Remote report link", link.Contents);
        Assert.Equal(1, link.PageNumber);

        PdfReadDocument document = PdfReadDocument.Open(bytes);
        PdfLinkAnnotation pageLink = Assert.Single(document.Pages[0].GetLinkAnnotations());
        Assert.True(pageLink.IsRemoteGoToLink);
        Assert.Equal("remote-report.pdf", pageLink.RemoteFile);
        Assert.Equal(2, pageLink.RemoteDestinationPageNumber);
    }


}
