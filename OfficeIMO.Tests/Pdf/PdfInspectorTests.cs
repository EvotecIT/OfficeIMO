using System;
using System.IO;
using System.Linq;
using OfficeIMO.Drawing;
using OfficeIMO.Pdf;
using Xunit;

namespace OfficeIMO.Tests.Pdf;

public class PdfInspectorTests {
    [Fact]
    public void Inspect_ReturnsPageCountSizesAndMetadata() {
        byte[] bytes = BuildTwoPagePdf();

        PdfDocumentInfo info = PdfInspector.Inspect(bytes);

        Assert.Equal(2, info.PageCount);
        Assert.Equal("1.4", info.HeaderVersion);
        Assert.Equal("Inspection sample", info.Metadata.Title);
        Assert.Equal("OfficeIMO", info.Metadata.Author);
        Assert.Equal("Roadmap", info.Metadata.Subject);
        Assert.Equal("pdf,inspect", info.Metadata.Keywords);
        Assert.Null(info.CatalogPageMode);
        Assert.Null(info.CatalogPageLayout);
        Assert.Null(info.CatalogVersion);
        Assert.Null(info.CatalogLanguage);
        Assert.False(info.HasLinkAnnotations);
        Assert.Equal(0, info.LinkAnnotationCount);
        Assert.Equal(0, info.LinkUriCount);
        Assert.Equal(0, info.LinkDestinationCount);
        Assert.Empty(info.LinkAnnotations);
        Assert.Empty(info.LinkUris);
        Assert.Empty(info.LinkDestinationNames);
        Assert.False(info.HasNamedDestinations);
        Assert.Equal(0, info.NamedDestinationCount);
        Assert.Empty(info.NamedDestinations);
        Assert.Empty(info.NamedDestinationNames);
        Assert.False(info.HasReadableOpenAction);
        Assert.Null(info.OpenAction);
        Assert.False(info.HasReadableViewerPreferences);
        Assert.Null(info.ViewerPreferences);
        Assert.False(info.HasReadablePageLabels);
        Assert.Equal(0, info.PageLabelCount);
        Assert.Empty(info.PageLabels);

        Assert.Equal(1, info.Pages[0].PageNumber);
        Assert.Equal(595, info.Pages[0].Width);
        Assert.Equal(842, info.Pages[0].Height);
        Assert.False(info.Pages[0].IsLandscape);

        Assert.Equal(2, info.Pages[1].PageNumber);
        Assert.Equal(792, info.Pages[1].Width);
        Assert.Equal(612, info.Pages[1].Height);
        Assert.True(info.Pages[1].IsLandscape);
    }

    [Fact]
    public void Inspect_ReadsFromPathAndStream() {
        byte[] bytes = BuildTwoPagePdf();
        string path = Path.Combine(Path.GetTempPath(), "officeimo-pdf-inspect-" + Guid.NewGuid().ToString("N") + ".pdf");

        try {
            File.WriteAllBytes(path, bytes);

            PdfDocumentInfo fromPath = PdfInspector.Inspect(path);
            using var stream = new MemoryStream(bytes);
            PdfDocumentInfo fromStream = PdfInspector.Inspect(stream);

            Assert.Equal(2, fromPath.PageCount);
            Assert.Equal(2, fromStream.PageCount);
            Assert.Equal(fromPath.Pages[1].Width, fromStream.Pages[1].Width);
            Assert.Equal(fromPath.Pages[1].Height, fromStream.Pages[1].Height);
        } finally {
            if (File.Exists(path)) {
                File.Delete(path);
            }
        }
    }

    [Fact]
    public void Inspect_ReportsSignatureMarkersWithoutFailingRead() {
        PdfDocumentInfo info = PdfInspector.Inspect(BuildSignedPdfMarker());

        Assert.True(info.HasSignatures);
        Assert.Equal(1, info.PageCount);
        Assert.Equal(200, info.Pages[0].Width);
        Assert.Equal(200, info.Pages[0].Height);
    }

    [Fact]
    public void Probe_ReportsVersionSecurityMarkersAndDoesNotRequireFullParsing() {
        byte[] bytes = BuildEncryptedPdfMarker();

        PdfDocumentProbe probe = PdfInspector.Probe(bytes);

        Assert.Equal("1.7", probe.HeaderVersion);
        Assert.True(probe.HasEncryption);
        Assert.False(probe.HasSignatures);
        Assert.False(probe.HasForms);
        Assert.False(probe.HasAnnotations);
        Assert.False(probe.HasOutlines);
        Assert.False(probe.HasCatalogViewSettings);
        Assert.False(probe.HasPageLabels);
        Assert.False(probe.HasCatalogNameTrees);
        Assert.False(probe.HasNamedDestinations);
        Assert.False(probe.HasOpenActions);
        Assert.False(probe.HasViewerPreferences);
        Assert.False(probe.HasTaggedContent);
        Assert.False(probe.HasXmpMetadata);
        Assert.False(probe.HasCatalogUri);
        Assert.False(probe.HasOutputIntents);
        Assert.False(probe.HasEmbeddedFiles);
        Assert.False(probe.HasOptionalContent);
        Assert.False(probe.HasActiveContent);
        Assert.Throws<NotSupportedException>(() => PdfInspector.Inspect(bytes));
    }

    [Fact]
    public void Probe_ReadsFromPathAndStream() {
        byte[] bytes = BuildSignedPdfMarker();
        string path = Path.Combine(Path.GetTempPath(), "officeimo-pdf-probe-" + Guid.NewGuid().ToString("N") + ".pdf");

        try {
            File.WriteAllBytes(path, bytes);

            PdfDocumentProbe fromPath = PdfInspector.Probe(path);
            using var stream = new MemoryStream(bytes);
            PdfDocumentProbe fromStream = PdfInspector.Probe(stream);

            Assert.Equal("1.4", fromPath.HeaderVersion);
            Assert.True(fromPath.HasSignatures);
            Assert.True(fromPath.HasForms);
            Assert.False(fromPath.HasEncryption);
            Assert.False(fromPath.HasAnnotations);
            Assert.False(fromPath.HasOutlines);
            Assert.False(fromPath.HasCatalogViewSettings);
            Assert.False(fromPath.HasPageLabels);
            Assert.False(fromPath.HasCatalogNameTrees);
            Assert.False(fromPath.HasNamedDestinations);
            Assert.False(fromPath.HasOpenActions);
            Assert.False(fromPath.HasViewerPreferences);
            Assert.False(fromPath.HasTaggedContent);
            Assert.False(fromPath.HasXmpMetadata);
            Assert.False(fromPath.HasCatalogUri);
            Assert.False(fromPath.HasOutputIntents);
            Assert.False(fromPath.HasEmbeddedFiles);
            Assert.False(fromPath.HasOptionalContent);
            Assert.False(fromPath.HasActiveContent);
            Assert.Equal(fromPath.HeaderVersion, fromStream.HeaderVersion);
            Assert.Equal(fromPath.HasSignatures, fromStream.HasSignatures);
            Assert.Equal(fromPath.HasForms, fromStream.HasForms);
            Assert.Equal(fromPath.HasEncryption, fromStream.HasEncryption);
            Assert.Equal(fromPath.HasAnnotations, fromStream.HasAnnotations);
            Assert.Equal(fromPath.HasOutlines, fromStream.HasOutlines);
            Assert.Equal(fromPath.HasCatalogViewSettings, fromStream.HasCatalogViewSettings);
            Assert.Equal(fromPath.HasPageLabels, fromStream.HasPageLabels);
            Assert.Equal(fromPath.HasCatalogNameTrees, fromStream.HasCatalogNameTrees);
            Assert.Equal(fromPath.HasNamedDestinations, fromStream.HasNamedDestinations);
            Assert.Equal(fromPath.HasOpenActions, fromStream.HasOpenActions);
            Assert.Equal(fromPath.HasViewerPreferences, fromStream.HasViewerPreferences);
            Assert.Equal(fromPath.HasTaggedContent, fromStream.HasTaggedContent);
            Assert.Equal(fromPath.HasXmpMetadata, fromStream.HasXmpMetadata);
            Assert.Equal(fromPath.HasCatalogUri, fromStream.HasCatalogUri);
            Assert.Equal(fromPath.HasOutputIntents, fromStream.HasOutputIntents);
            Assert.Equal(fromPath.HasEmbeddedFiles, fromStream.HasEmbeddedFiles);
            Assert.Equal(fromPath.HasOptionalContent, fromStream.HasOptionalContent);
            Assert.Equal(fromPath.HasActiveContent, fromStream.HasActiveContent);
        } finally {
            if (File.Exists(path)) {
                File.Delete(path);
            }
        }
    }

    [Fact]
    public void Probe_RejectsInvalidInputs() {
        Assert.Throws<ArgumentNullException>(() => PdfInspector.Probe((byte[])null!));
        Assert.Throws<ArgumentNullException>(() => PdfInspector.Probe((string)null!));
        Assert.Throws<ArgumentException>(() => PdfInspector.Probe(" "));
        Assert.Throws<ArgumentNullException>(() => PdfInspector.Probe((Stream)null!));
        Assert.Throws<ArgumentException>(() => PdfInspector.Probe(new WriteOnlyStream()));
    }

    [Fact]
    public void Preflight_AllowsGeneratedPdfForReadAndRewrite() {
        PdfDocumentPreflight report = PdfInspector.Preflight(BuildTwoPagePdf());

        Assert.True(report.CanRead);
        Assert.True(report.CanRewrite);
        Assert.Empty(report.Diagnostics);
        Assert.Empty(report.ReadBlockers);
        Assert.Empty(report.RewriteBlockers);
        Assert.False(report.HasReadBlocker(PdfReadBlockerKind.Encryption));
        Assert.False(report.HasRewriteBlocker(PdfRewriteBlockerKind.Forms));
        Assert.NotNull(report.DocumentInfo);
        Assert.Equal(2, report.DocumentInfo!.PageCount);
        Assert.Equal("1.4", report.Probe.HeaderVersion);
    }

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
        Assert.Equal(2, info.LinkAnnotationCount);
        Assert.Equal(2, info.LinkAnnotations.Count);
        Assert.Equal(1, info.LinkUriCount);
        Assert.Equal(new[] { "https://evotec.xyz" }, info.LinkUris);
        Assert.Single(info.Pages);
        Assert.Equal(2, info.Pages[0].LinkAnnotations.Count);
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

        PdfReadDocument document = PdfReadDocument.Load(bytes);
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

        byte[] wrappedBytes = PdfDoc.Create(wrappedOptions)
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

        byte[] rowBytes = PdfDoc.Create(new PdfOptions {
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
        byte[] bytes = PdfDoc.Create(new PdfOptions {
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

        byte[] bytes = PdfDoc.Create(new PdfOptions {
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
        byte[] bytes = PdfDoc.Create(new PdfOptions {
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
        byte[] bytes = PdfDoc.Create(new PdfOptions {
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
        byte[] bytes = PdfDoc.Create(new PdfOptions {
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
    public void Preflight_AllowsSimpleOutlinePdfReadAndRewrite() {
        PdfDocumentPreflight report = PdfInspector.Preflight(BuildOutlinePdf());

        Assert.True(report.CanRead);
        Assert.True(report.CanRewrite);
        Assert.True(report.Probe.HasOutlines);
        Assert.True(report.Probe.HasCatalogViewSettings);
        Assert.NotNull(report.DocumentInfo);
        Assert.True(report.DocumentInfo!.HasOutlines);
        Assert.True(report.DocumentInfo.HasCatalogViewSettings);
        Assert.NotEmpty(report.DocumentInfo.Outlines);
        Assert.Empty(report.ReadBlockers);
        Assert.Empty(report.Diagnostics);
        Assert.Empty(report.RewriteBlockers);
        Assert.False(report.HasRewriteBlocker(PdfRewriteBlockerKind.CatalogViewSettings));
    }

    [Fact]
    public void Preflight_AllowsComplexOutlinePdfReadButBlocksRewrite() {
        PdfDocumentPreflight report = PdfInspector.Preflight(BuildComplexOutlinePdf());

        Assert.True(report.CanRead);
        Assert.False(report.CanRewrite);
        Assert.True(report.Probe.HasOutlines);
        Assert.True(report.Probe.HasCatalogViewSettings);
        Assert.NotNull(report.DocumentInfo);
        Assert.True(report.DocumentInfo!.HasOutlines);
        Assert.True(report.DocumentInfo.HasCatalogViewSettings);
        Assert.NotEmpty(report.DocumentInfo.Outlines);
        Assert.Empty(report.ReadBlockers);
        Assert.Contains("PDF outlines are not supported for rewriting by OfficeIMO.Pdf yet.", report.Diagnostics);
        AssertRewriteBlocker(report, PdfRewriteBlockerKind.Outlines, "PDF outlines are not supported for rewriting by OfficeIMO.Pdf yet.");
        Assert.False(report.HasRewriteBlocker(PdfRewriteBlockerKind.CatalogViewSettings));
    }

    [Fact]
    public void Preflight_AllowsCatalogViewSettingPdfReadAndRewrite() {
        PdfDocumentPreflight report = PdfInspector.Preflight(BuildCatalogViewSettingPdf());

        Assert.True(report.CanRead);
        Assert.True(report.CanRewrite);
        Assert.True(report.Probe.HasCatalogViewSettings);
        Assert.False(report.Probe.HasOutlines);
        Assert.NotNull(report.DocumentInfo);
        Assert.True(report.DocumentInfo!.HasCatalogViewSettings);
        Assert.False(report.DocumentInfo.HasOutlines);
        Assert.Equal("FullScreen", report.DocumentInfo.CatalogPageMode);
        Assert.Equal("TwoColumnLeft", report.DocumentInfo.CatalogPageLayout);
        Assert.Null(report.DocumentInfo.CatalogVersion);
        Assert.Null(report.DocumentInfo.CatalogLanguage);
        Assert.Empty(report.DocumentInfo.Outlines);
        Assert.Empty(report.ReadBlockers);
        Assert.Empty(report.Diagnostics);
        Assert.Empty(report.RewriteBlockers);
        Assert.False(report.HasRewriteBlocker(PdfRewriteBlockerKind.CatalogViewSettings));
    }

    [Fact]
    public void Preflight_ReadsCatalogLanguageAndVersion() {
        PdfDocumentPreflight report = PdfInspector.Preflight(BuildCatalogIdentityPdf());

        Assert.True(report.CanRead);
        Assert.True(report.CanRewrite);
        Assert.NotNull(report.DocumentInfo);
        Assert.Equal("1.7", report.DocumentInfo!.CatalogVersion);
        Assert.Equal("pl-PL", report.DocumentInfo.CatalogLanguage);
        Assert.Null(report.DocumentInfo.CatalogPageMode);
        Assert.Null(report.DocumentInfo.CatalogPageLayout);
        Assert.Empty(report.ReadBlockers);
        Assert.Empty(report.Diagnostics);
        Assert.Empty(report.RewriteBlockers);
    }

    [Fact]
    public void Preflight_AllowsPageLabelPdfReadAndRewrite() {
        PdfDocumentPreflight report = PdfInspector.Preflight(BuildPageLabelPdf());

        Assert.True(report.CanRead);
        Assert.True(report.CanRewrite);
        Assert.True(report.Probe.HasPageLabels);
        Assert.False(report.Probe.HasCatalogNameTrees);
        Assert.False(report.Probe.HasNamedDestinations);
        Assert.NotNull(report.DocumentInfo);
        Assert.True(report.DocumentInfo!.HasPageLabels);
        Assert.False(report.DocumentInfo.HasCatalogNameTrees);
        Assert.False(report.DocumentInfo.HasNamedDestinations);
        AssertPageLabel(report.DocumentInfo, 0, 1, "D", null, 1);
        Assert.Empty(report.ReadBlockers);
        Assert.Empty(report.Diagnostics);
        Assert.Empty(report.RewriteBlockers);
        Assert.False(report.HasRewriteBlocker(PdfRewriteBlockerKind.PageLabels));
    }

    [Fact]
    public void Preflight_AllowsComplexPageLabelPdfReadButBlocksRewrite() {
        PdfDocumentPreflight report = PdfInspector.Preflight(BuildComplexPageLabelPdf());

        Assert.True(report.CanRead);
        Assert.False(report.CanRewrite);
        Assert.True(report.Probe.HasPageLabels);
        Assert.NotNull(report.DocumentInfo);
        Assert.True(report.DocumentInfo!.HasPageLabels);
        Assert.False(report.DocumentInfo.HasReadablePageLabels);
        Assert.Equal(0, report.DocumentInfo.PageLabelCount);
        Assert.Empty(report.DocumentInfo.PageLabels);
        Assert.Empty(report.ReadBlockers);
        Assert.Contains(report.Diagnostics, diagnostic => diagnostic.Contains("page labels", StringComparison.OrdinalIgnoreCase));
        Assert.Contains(report.RewriteBlockers, blocker => blocker.Kind == PdfRewriteBlockerKind.PageLabels);
        Assert.True(report.HasRewriteBlocker(PdfRewriteBlockerKind.PageLabels));
    }

    [Fact]
    public void Preflight_AllowsDirectNamedDestinationPdfReadAndRewrite() {
        PdfDocumentPreflight report = PdfInspector.Preflight(BuildNamedDestinationPdf());

        Assert.True(report.CanRead);
        Assert.True(report.CanRewrite);
        Assert.True(report.Probe.HasNamedDestinations);
        Assert.False(report.Probe.HasCatalogNameTrees);
        Assert.False(report.Probe.HasPageLabels);
        Assert.NotNull(report.DocumentInfo);
        Assert.True(report.DocumentInfo!.HasNamedDestinations);
        Assert.False(report.DocumentInfo.HasCatalogNameTrees);
        Assert.False(report.DocumentInfo.HasPageLabels);
        AssertNamedDestination(report.DocumentInfo, "Chapter1", 1, 200);
        Assert.Empty(report.ReadBlockers);
        Assert.Empty(report.Diagnostics);
        Assert.Empty(report.RewriteBlockers);
        Assert.False(report.HasRewriteBlocker(PdfRewriteBlockerKind.NamedDestinations));
    }

    [Fact]
    public void Preflight_AllowsNamedDestinationNameTreeReadAndRewrite() {
        PdfDocumentPreflight report = PdfInspector.Preflight(BuildNamedDestinationNameTreePdf());

        Assert.True(report.CanRead);
        Assert.True(report.CanRewrite);
        Assert.True(report.Probe.HasNamedDestinations);
        Assert.True(report.Probe.HasCatalogNameTrees);
        Assert.NotNull(report.DocumentInfo);
        Assert.True(report.DocumentInfo!.HasNamedDestinations);
        Assert.True(report.DocumentInfo.HasCatalogNameTrees);
        AssertNamedDestination(report.DocumentInfo, "Chapter1", 1, 200);
        Assert.Empty(report.ReadBlockers);
        Assert.Empty(report.Diagnostics);
        Assert.Empty(report.RewriteBlockers);
        Assert.False(report.HasRewriteBlocker(PdfRewriteBlockerKind.NamedDestinations));
    }

    [Fact]
    public void Preflight_AllowsComplexNamedDestinationNameTreeReadButBlocksRewrite() {
        PdfDocumentPreflight report = PdfInspector.Preflight(BuildComplexNamedDestinationNameTreePdf());

        Assert.True(report.CanRead);
        Assert.False(report.CanRewrite);
        Assert.True(report.Probe.HasNamedDestinations);
        Assert.True(report.Probe.HasCatalogNameTrees);
        Assert.NotNull(report.DocumentInfo);
        Assert.True(report.DocumentInfo!.HasNamedDestinations);
        Assert.True(report.DocumentInfo.HasCatalogNameTrees);
        Assert.Empty(report.ReadBlockers);
        Assert.Contains("PDF named destinations are not supported for rewriting by OfficeIMO.Pdf yet.", report.Diagnostics);
        AssertRewriteBlocker(report, PdfRewriteBlockerKind.NamedDestinations, "PDF named destinations are not supported for rewriting by OfficeIMO.Pdf yet.");
    }

    [Fact]
    public void Preflight_AllowsDestinationOpenActionPdfReadAndRewrite() {
        PdfDocumentPreflight report = PdfInspector.Preflight(BuildOpenActionPdf());

        Assert.True(report.CanRead);
        Assert.True(report.CanRewrite);
        Assert.True(report.Probe.HasOpenActions);
        Assert.False(report.Probe.HasViewerPreferences);
        Assert.NotNull(report.DocumentInfo);
        Assert.True(report.DocumentInfo!.HasOpenActions);
        Assert.False(report.DocumentInfo.HasViewerPreferences);
        AssertOpenAction(report.DocumentInfo, "Destination", 1, null);
        Assert.Empty(report.ReadBlockers);
        Assert.Empty(report.Diagnostics);
        Assert.Empty(report.RewriteBlockers);
        Assert.False(report.HasRewriteBlocker(PdfRewriteBlockerKind.OpenActions));
    }

    [Fact]
    public void Preflight_AllowsGoToOpenActionDictionaryReadAndRewrite() {
        PdfDocumentPreflight report = PdfInspector.Preflight(BuildOpenActionDictionaryPdf());

        Assert.True(report.CanRead);
        Assert.True(report.CanRewrite);
        Assert.True(report.Probe.HasOpenActions);
        Assert.NotNull(report.DocumentInfo);
        Assert.True(report.DocumentInfo!.HasOpenActions);
        AssertOpenAction(report.DocumentInfo, "GoTo", 1, null);
        Assert.Empty(report.ReadBlockers);
        Assert.Empty(report.Diagnostics);
        Assert.Empty(report.RewriteBlockers);
        Assert.False(report.HasRewriteBlocker(PdfRewriteBlockerKind.OpenActions));
    }

    [Fact]
    public void Preflight_AllowsComplexOpenActionDictionaryReadButBlocksRewrite() {
        PdfDocumentPreflight report = PdfInspector.Preflight(BuildComplexOpenActionDictionaryPdf());

        Assert.True(report.CanRead);
        Assert.False(report.CanRewrite);
        Assert.True(report.Probe.HasOpenActions);
        Assert.NotNull(report.DocumentInfo);
        Assert.True(report.DocumentInfo!.HasOpenActions);
        Assert.False(report.DocumentInfo.HasReadableOpenAction);
        Assert.Null(report.DocumentInfo.OpenAction);
        Assert.Empty(report.ReadBlockers);
        Assert.Contains("PDF open actions are not supported for rewriting by OfficeIMO.Pdf yet.", report.Diagnostics);
        AssertRewriteBlocker(report, PdfRewriteBlockerKind.OpenActions, "PDF open actions are not supported for rewriting by OfficeIMO.Pdf yet.");
    }

    [Fact]
    public void Preflight_AllowsSimpleViewerPreferencePdfReadAndRewrite() {
        PdfDocumentPreflight report = PdfInspector.Preflight(BuildViewerPreferencePdf());

        Assert.True(report.CanRead);
        Assert.True(report.CanRewrite);
        Assert.True(report.Probe.HasViewerPreferences);
        Assert.False(report.Probe.HasOpenActions);
        Assert.NotNull(report.DocumentInfo);
        Assert.True(report.DocumentInfo!.HasViewerPreferences);
        Assert.False(report.DocumentInfo.HasOpenActions);
        AssertViewerPreferences(report.DocumentInfo);
        Assert.Empty(report.ReadBlockers);
        Assert.Empty(report.RewriteBlockers);
        Assert.DoesNotContain("PDF viewer preferences are not supported for rewriting by OfficeIMO.Pdf yet.", report.Diagnostics);
    }

    [Fact]
    public void Preflight_AllowsComplexViewerPreferencePdfReadButBlocksRewrite() {
        PdfDocumentPreflight report = PdfInspector.Preflight(BuildComplexViewerPreferencePdf());

        Assert.True(report.CanRead);
        Assert.False(report.CanRewrite);
        Assert.True(report.Probe.HasViewerPreferences);
        Assert.NotNull(report.DocumentInfo);
        Assert.True(report.DocumentInfo!.HasViewerPreferences);
        Assert.False(report.DocumentInfo.HasReadableViewerPreferences);
        Assert.Null(report.DocumentInfo.ViewerPreferences);
        Assert.Empty(report.ReadBlockers);
        Assert.Contains("PDF viewer preferences are not supported for rewriting by OfficeIMO.Pdf yet.", report.Diagnostics);
        AssertRewriteBlocker(report, PdfRewriteBlockerKind.ViewerPreferences, "PDF viewer preferences are not supported for rewriting by OfficeIMO.Pdf yet.");
    }

    [Fact]
    public void Preflight_AllowsTaggedPdfReadButBlocksRewrite() {
        PdfDocumentPreflight report = PdfInspector.Preflight(BuildTaggedPdf());

        Assert.True(report.CanRead);
        Assert.False(report.CanRewrite);
        Assert.True(report.Probe.HasTaggedContent);
        Assert.NotNull(report.DocumentInfo);
        Assert.True(report.DocumentInfo!.HasTaggedContent);
        Assert.Empty(report.ReadBlockers);
        Assert.Contains("PDF tagged content structure is not supported for rewriting by OfficeIMO.Pdf yet.", report.Diagnostics);
        AssertRewriteBlocker(report, PdfRewriteBlockerKind.TaggedContent, "PDF tagged content structure is not supported for rewriting by OfficeIMO.Pdf yet.");
    }

    [Fact]
    public void Preflight_AllowsSimpleXmpMetadataPdfReadAndRewrite() {
        PdfDocumentPreflight report = PdfInspector.Preflight(BuildXmpMetadataPdf());

        Assert.True(report.CanRead);
        Assert.True(report.CanRewrite);
        Assert.True(report.Probe.HasXmpMetadata);
        Assert.False(report.Probe.HasCatalogUri);
        Assert.False(report.Probe.HasOutputIntents);
        Assert.NotNull(report.DocumentInfo);
        Assert.True(report.DocumentInfo!.HasXmpMetadata);
        Assert.False(report.DocumentInfo.HasCatalogUri);
        Assert.False(report.DocumentInfo.HasOutputIntents);
        Assert.Empty(report.ReadBlockers);
        Assert.Empty(report.RewriteBlockers);
        Assert.DoesNotContain("PDF XMP metadata is not supported for rewriting by OfficeIMO.Pdf yet.", report.Diagnostics);
    }

    [Fact]
    public void Preflight_AllowsComplexXmpMetadataPdfReadButBlocksRewrite() {
        PdfDocumentPreflight report = PdfInspector.Preflight(BuildComplexXmpMetadataPdf());

        Assert.True(report.CanRead);
        Assert.False(report.CanRewrite);
        Assert.True(report.Probe.HasXmpMetadata);
        Assert.NotNull(report.DocumentInfo);
        Assert.True(report.DocumentInfo!.HasXmpMetadata);
        Assert.Empty(report.ReadBlockers);
        Assert.Contains("PDF XMP metadata is not supported for rewriting by OfficeIMO.Pdf yet.", report.Diagnostics);
        AssertRewriteBlocker(report, PdfRewriteBlockerKind.XmpMetadata, "PDF XMP metadata is not supported for rewriting by OfficeIMO.Pdf yet.");
    }

    [Fact]
    public void Preflight_AllowsSimpleCatalogUriPdfReadAndRewrite() {
        PdfDocumentPreflight report = PdfInspector.Preflight(BuildCatalogUriPdf());

        Assert.True(report.CanRead);
        Assert.True(report.CanRewrite);
        Assert.True(report.Probe.HasCatalogUri);
        Assert.False(report.Probe.HasOutputIntents);
        Assert.NotNull(report.DocumentInfo);
        Assert.True(report.DocumentInfo!.HasCatalogUri);
        Assert.False(report.DocumentInfo.HasOutputIntents);
        Assert.Empty(report.ReadBlockers);
        Assert.Empty(report.RewriteBlockers);
        Assert.DoesNotContain("PDF catalog URI dictionaries are not supported for rewriting by OfficeIMO.Pdf yet.", report.Diagnostics);
    }

    [Fact]
    public void Preflight_DoesNotTreatLinkAnnotationUriAsCatalogUri() {
        PdfDocumentPreflight report = PdfInspector.Preflight(BuildAnnotatedPdf());

        Assert.True(report.CanRead);
        Assert.True(report.CanRewrite);
        Assert.True(report.Probe.HasAnnotations);
        Assert.False(report.Probe.HasCatalogUri);
        Assert.NotNull(report.DocumentInfo);
        Assert.True(report.DocumentInfo!.HasAnnotations);
        Assert.False(report.DocumentInfo.HasCatalogUri);
        Assert.Empty(report.ReadBlockers);
        Assert.Empty(report.RewriteBlockers);
    }

    [Fact]
    public void Preflight_AllowsComplexCatalogUriPdfReadButBlocksRewrite() {
        PdfDocumentPreflight report = PdfInspector.Preflight(BuildComplexCatalogUriPdf());

        Assert.True(report.CanRead);
        Assert.False(report.CanRewrite);
        Assert.True(report.Probe.HasCatalogUri);
        Assert.NotNull(report.DocumentInfo);
        Assert.True(report.DocumentInfo!.HasCatalogUri);
        Assert.Empty(report.ReadBlockers);
        Assert.Contains("PDF catalog URI dictionaries are not supported for rewriting by OfficeIMO.Pdf yet.", report.Diagnostics);
        AssertRewriteBlocker(report, PdfRewriteBlockerKind.CatalogUri, "PDF catalog URI dictionaries are not supported for rewriting by OfficeIMO.Pdf yet.");
    }

    [Fact]
    public void Preflight_AllowsSimpleOutputIntentPdfReadAndRewrite() {
        PdfDocumentPreflight report = PdfInspector.Preflight(BuildOutputIntentPdf());

        Assert.True(report.CanRead);
        Assert.True(report.CanRewrite);
        Assert.True(report.Probe.HasOutputIntents);
        Assert.False(report.Probe.HasXmpMetadata);
        Assert.False(report.Probe.HasCatalogUri);
        Assert.NotNull(report.DocumentInfo);
        Assert.True(report.DocumentInfo!.HasOutputIntents);
        Assert.False(report.DocumentInfo.HasXmpMetadata);
        Assert.False(report.DocumentInfo.HasCatalogUri);
        Assert.Empty(report.ReadBlockers);
        Assert.Empty(report.RewriteBlockers);
        Assert.DoesNotContain("PDF output intents are not supported for rewriting by OfficeIMO.Pdf yet.", report.Diagnostics);
    }

    [Fact]
    public void Preflight_AllowsComplexOutputIntentPdfReadButBlocksRewrite() {
        PdfDocumentPreflight report = PdfInspector.Preflight(BuildComplexOutputIntentPdf());

        Assert.True(report.CanRead);
        Assert.False(report.CanRewrite);
        Assert.True(report.Probe.HasOutputIntents);
        Assert.NotNull(report.DocumentInfo);
        Assert.True(report.DocumentInfo!.HasOutputIntents);
        Assert.Empty(report.ReadBlockers);
        Assert.Contains("PDF output intents are not supported for rewriting by OfficeIMO.Pdf yet.", report.Diagnostics);
        AssertRewriteBlocker(report, PdfRewriteBlockerKind.OutputIntents, "PDF output intents are not supported for rewriting by OfficeIMO.Pdf yet.");
    }

    [Fact]
    public void Preflight_AllowsSimpleEmbeddedFilePdfReadAndRewrite() {
        PdfDocumentPreflight report = PdfInspector.Preflight(BuildEmbeddedFilePdf());

        Assert.True(report.CanRead);
        Assert.True(report.CanRewrite);
        Assert.True(report.Probe.HasEmbeddedFiles);
        Assert.True(report.Probe.HasCatalogNameTrees);
        Assert.False(report.Probe.HasOptionalContent);
        Assert.NotNull(report.DocumentInfo);
        Assert.True(report.DocumentInfo!.HasEmbeddedFiles);
        Assert.True(report.DocumentInfo.HasCatalogNameTrees);
        Assert.False(report.DocumentInfo.HasOptionalContent);
        Assert.Empty(report.ReadBlockers);
        Assert.Empty(report.RewriteBlockers);
        Assert.DoesNotContain("PDF embedded files are not supported for rewriting by OfficeIMO.Pdf yet.", report.Diagnostics);
    }

    [Fact]
    public void Preflight_AllowsSimpleAssociatedFilePdfReadAndRewrite() {
        PdfDocumentPreflight report = PdfInspector.Preflight(BuildAssociatedFilePdf());

        Assert.True(report.CanRead);
        Assert.True(report.CanRewrite);
        Assert.True(report.Probe.HasEmbeddedFiles);
        Assert.False(report.Probe.HasCatalogNameTrees);
        Assert.False(report.Probe.HasOptionalContent);
        Assert.NotNull(report.DocumentInfo);
        Assert.True(report.DocumentInfo!.HasEmbeddedFiles);
        Assert.False(report.DocumentInfo.HasCatalogNameTrees);
        Assert.False(report.DocumentInfo.HasOptionalContent);
        Assert.Empty(report.ReadBlockers);
        Assert.Empty(report.RewriteBlockers);
        Assert.DoesNotContain("PDF embedded files are not supported for rewriting by OfficeIMO.Pdf yet.", report.Diagnostics);
    }

    [Fact]
    public void Preflight_AllowsCombinedDestinationAndEmbeddedFileNameTreesReadAndRewrite() {
        PdfDocumentPreflight report = PdfInspector.Preflight(BuildCombinedDestinationAndEmbeddedFileNameTreePdf());

        Assert.True(report.CanRead);
        Assert.True(report.CanRewrite);
        Assert.True(report.Probe.HasNamedDestinations);
        Assert.True(report.Probe.HasEmbeddedFiles);
        Assert.True(report.Probe.HasCatalogNameTrees);
        Assert.NotNull(report.DocumentInfo);
        Assert.True(report.DocumentInfo!.HasNamedDestinations);
        Assert.True(report.DocumentInfo.HasEmbeddedFiles);
        Assert.True(report.DocumentInfo.HasCatalogNameTrees);
        Assert.Empty(report.ReadBlockers);
        Assert.Empty(report.RewriteBlockers);
        Assert.DoesNotContain("PDF named destinations are not supported for rewriting by OfficeIMO.Pdf yet.", report.Diagnostics);
        Assert.DoesNotContain("PDF embedded files are not supported for rewriting by OfficeIMO.Pdf yet.", report.Diagnostics);
    }

    [Fact]
    public void Preflight_AllowsUnsupportedCatalogNameTreePdfReadButBlocksRewrite() {
        PdfDocumentPreflight report = PdfInspector.Preflight(BuildUnsupportedCatalogNameTreePdf());

        Assert.True(report.CanRead);
        Assert.False(report.CanRewrite);
        Assert.True(report.Probe.HasCatalogNameTrees);
        Assert.False(report.Probe.HasNamedDestinations);
        Assert.False(report.Probe.HasEmbeddedFiles);
        Assert.NotNull(report.DocumentInfo);
        Assert.True(report.DocumentInfo!.HasCatalogNameTrees);
        Assert.False(report.DocumentInfo.HasNamedDestinations);
        Assert.False(report.DocumentInfo.HasEmbeddedFiles);
        Assert.Empty(report.ReadBlockers);
        Assert.Contains("PDF catalog name trees are not supported for rewriting by OfficeIMO.Pdf yet.", report.Diagnostics);
        AssertRewriteBlocker(report, PdfRewriteBlockerKind.CatalogNameTrees, "PDF catalog name trees are not supported for rewriting by OfficeIMO.Pdf yet.");
    }

    [Fact]
    public void Preflight_AllowsComplexEmbeddedFilePdfReadButBlocksRewrite() {
        PdfDocumentPreflight report = PdfInspector.Preflight(BuildComplexEmbeddedFilePdf());

        Assert.True(report.CanRead);
        Assert.False(report.CanRewrite);
        Assert.True(report.Probe.HasEmbeddedFiles);
        Assert.True(report.Probe.HasCatalogNameTrees);
        Assert.NotNull(report.DocumentInfo);
        Assert.True(report.DocumentInfo!.HasEmbeddedFiles);
        Assert.True(report.DocumentInfo.HasCatalogNameTrees);
        Assert.Empty(report.ReadBlockers);
        Assert.Contains("PDF embedded files are not supported for rewriting by OfficeIMO.Pdf yet.", report.Diagnostics);
        AssertRewriteBlocker(report, PdfRewriteBlockerKind.EmbeddedFiles, "PDF embedded files are not supported for rewriting by OfficeIMO.Pdf yet.");
    }

    [Fact]
    public void Preflight_AllowsComplexAssociatedFilePdfReadButBlocksRewrite() {
        PdfDocumentPreflight report = PdfInspector.Preflight(BuildComplexAssociatedFilePdf());

        Assert.True(report.CanRead);
        Assert.False(report.CanRewrite);
        Assert.True(report.Probe.HasEmbeddedFiles);
        Assert.False(report.Probe.HasCatalogNameTrees);
        Assert.NotNull(report.DocumentInfo);
        Assert.True(report.DocumentInfo!.HasEmbeddedFiles);
        Assert.False(report.DocumentInfo.HasCatalogNameTrees);
        Assert.Empty(report.ReadBlockers);
        Assert.Contains("PDF embedded files are not supported for rewriting by OfficeIMO.Pdf yet.", report.Diagnostics);
        AssertRewriteBlocker(report, PdfRewriteBlockerKind.EmbeddedFiles, "PDF embedded files are not supported for rewriting by OfficeIMO.Pdf yet.");
    }

    [Fact]
    public void Preflight_AllowsSimpleOptionalContentPdfReadAndRewrite() {
        PdfDocumentPreflight report = PdfInspector.Preflight(BuildOptionalContentPdf());

        Assert.True(report.CanRead);
        Assert.True(report.CanRewrite);
        Assert.True(report.Probe.HasOptionalContent);
        Assert.False(report.Probe.HasEmbeddedFiles);
        Assert.NotNull(report.DocumentInfo);
        Assert.True(report.DocumentInfo!.HasOptionalContent);
        Assert.False(report.DocumentInfo.HasEmbeddedFiles);
        Assert.Empty(report.ReadBlockers);
        Assert.Empty(report.RewriteBlockers);
        Assert.DoesNotContain("PDF optional content layers are not supported for rewriting by OfficeIMO.Pdf yet.", report.Diagnostics);
    }

    [Fact]
    public void Preflight_AllowsComplexOptionalContentPdfReadButBlocksRewrite() {
        PdfDocumentPreflight report = PdfInspector.Preflight(BuildComplexOptionalContentPdf());

        Assert.True(report.CanRead);
        Assert.False(report.CanRewrite);
        Assert.True(report.Probe.HasOptionalContent);
        Assert.NotNull(report.DocumentInfo);
        Assert.True(report.DocumentInfo!.HasOptionalContent);
        Assert.Empty(report.ReadBlockers);
        Assert.Contains("PDF optional content layers are not supported for rewriting by OfficeIMO.Pdf yet.", report.Diagnostics);
        AssertRewriteBlocker(report, PdfRewriteBlockerKind.OptionalContent, "PDF optional content layers are not supported for rewriting by OfficeIMO.Pdf yet.");
    }

    [Fact]
    public void Preflight_AllowsActiveContentPdfReadButBlocksRewrite() {
        PdfDocumentPreflight report = PdfInspector.Preflight(BuildActiveContentPdf());

        Assert.True(report.CanRead);
        Assert.False(report.CanRewrite);
        Assert.True(report.Probe.HasActiveContent);
        Assert.True(report.Probe.HasCatalogNameTrees);
        Assert.NotNull(report.DocumentInfo);
        Assert.True(report.DocumentInfo!.HasActiveContent);
        Assert.True(report.DocumentInfo.HasCatalogNameTrees);
        Assert.Empty(report.ReadBlockers);
        Assert.Contains("PDF active content is not supported for rewriting by OfficeIMO.Pdf yet.", report.Diagnostics);
        AssertRewriteBlocker(report, PdfRewriteBlockerKind.ActiveContent, "PDF active content is not supported for rewriting by OfficeIMO.Pdf yet.");
    }

    [Fact]
    public void Preflight_BlocksEncryptedPdfBeforeFullInspection() {
        PdfDocumentPreflight report = PdfInspector.Preflight(BuildEncryptedPdfMarker());

        Assert.False(report.CanRead);
        Assert.False(report.CanRewrite);
        Assert.Null(report.DocumentInfo);
        Assert.True(report.Probe.HasEncryption);
        Assert.Contains("Encrypted PDF files are not supported by OfficeIMO.Pdf yet.", report.Diagnostics);
        AssertReadBlocker(report, PdfReadBlockerKind.Encryption, "Encrypted PDF files are not supported by OfficeIMO.Pdf yet.");
        AssertRewriteBlocker(report, PdfRewriteBlockerKind.Encryption, "Encrypted PDF files are not supported by OfficeIMO.Pdf yet.");
    }

    [Fact]
    public void Preflight_AllowsSignedPdfReadButBlocksRewrite() {
        PdfDocumentPreflight report = PdfInspector.Preflight(BuildSignedPdfMarker());

        Assert.True(report.CanRead);
        Assert.False(report.CanRewrite);
        Assert.NotNull(report.DocumentInfo);
        Assert.True(report.Probe.HasSignatures);
        Assert.True(report.DocumentInfo!.HasSignatures);
        Assert.True(report.Probe.HasForms);
        Assert.True(report.DocumentInfo.HasForms);
        Assert.Empty(report.ReadBlockers);
        Assert.Contains("Signed PDF files are not supported for rewriting by OfficeIMO.Pdf yet.", report.Diagnostics);
        AssertRewriteBlocker(report, PdfRewriteBlockerKind.Signatures, "Signed PDF files are not supported for rewriting by OfficeIMO.Pdf yet.");
        AssertRewriteBlocker(report, PdfRewriteBlockerKind.Forms, "PDF form fields are not supported for rewriting by OfficeIMO.Pdf yet.");
    }

    [Fact]
    public void Preflight_AllowsFormPdfReadButBlocksRewrite() {
        PdfDocumentPreflight report = PdfInspector.Preflight(BuildFormPdfMarker());

        Assert.True(report.CanRead);
        Assert.False(report.CanRewrite);
        Assert.NotNull(report.DocumentInfo);
        Assert.False(report.Probe.HasSignatures);
        Assert.True(report.Probe.HasForms);
        Assert.True(report.DocumentInfo!.HasForms);
        Assert.True(report.DocumentInfo.HasReadableFormFields);
        Assert.Equal(1, report.DocumentInfo.FormFieldCount);
        Assert.Equal("Name", report.DocumentInfo.FormFields[0].Name);
        Assert.Equal("Tx", report.DocumentInfo.FormFields[0].FieldType);
        Assert.Equal("OfficeIMO", report.DocumentInfo.FormFields[0].Value);
        Assert.Equal(new[] { "Name" }, report.DocumentInfo.FormFieldNames);
        Assert.Empty(report.ReadBlockers);
        Assert.Contains("PDF form fields are not supported for rewriting by OfficeIMO.Pdf yet.", report.Diagnostics);
        AssertRewriteBlocker(report, PdfRewriteBlockerKind.Forms, "PDF form fields are not supported for rewriting by OfficeIMO.Pdf yet.");
    }

    [Fact]
    public void Inspect_ReadsHierarchicalAcroFormFieldNamesForWrappers() {
        PdfDocumentInfo info = PdfInspector.Inspect(BuildHierarchicalFormPdf());

        Assert.True(info.HasForms);
        Assert.True(info.HasReadableFormFields);
        Assert.Equal(2, info.FormFieldCount);
        Assert.Equal(new[] { "Person.Name", "AcceptTerms" }, info.FormFieldNames);
        Assert.Equal("Person.Name", info.FormFields[0].Name);
        Assert.Equal("Name", info.FormFields[0].PartialName);
        Assert.Equal("Tx", info.FormFields[0].FieldType);
        Assert.Equal("OfficeIMO", info.FormFields[0].Value);
        Assert.Equal("Display name", info.FormFields[0].AlternateName);
        Assert.Equal("ExportName", info.FormFields[0].MappingName);
        Assert.Equal(1, info.FormFields[0].Flags);
        Assert.Equal("AcceptTerms", info.FormFields[1].Name);
        Assert.Equal("Btn", info.FormFields[1].FieldType);
        Assert.Equal("Yes", info.FormFields[1].Value);
    }

    [Fact]
    public void Preflight_ReportsInvalidHeaderWithoutParserException() {
        PdfDocumentPreflight report = PdfInspector.Preflight(System.Text.Encoding.ASCII.GetBytes("not a pdf"));

        Assert.False(report.CanRead);
        Assert.False(report.CanRewrite);
        Assert.Null(report.DocumentInfo);
        Assert.Null(report.Probe.HeaderVersion);
        Assert.Contains("PDF header was not found.", report.Diagnostics);
        AssertReadBlocker(report, PdfReadBlockerKind.MissingHeader, "PDF header was not found.");
        Assert.Empty(report.RewriteBlockers);
    }

    [Fact]
    public void Preflight_ReportsNoPagesWithReadBlocker() {
        PdfDocumentPreflight report = PdfInspector.Preflight(BuildNoPagesPdf());

        Assert.False(report.CanRead);
        Assert.False(report.CanRewrite);
        Assert.NotNull(report.DocumentInfo);
        Assert.Empty(report.DocumentInfo!.Pages);
        Assert.Contains("No PDF pages were discovered.", report.Diagnostics);
        AssertReadBlocker(report, PdfReadBlockerKind.NoPages, "No PDF pages were discovered.");
        Assert.Empty(report.RewriteBlockers);
    }

    [Fact]
    public void Preflight_ReadsFromPathAndStream() {
        byte[] bytes = BuildTwoPagePdf();
        string path = Path.Combine(Path.GetTempPath(), "officeimo-pdf-preflight-" + Guid.NewGuid().ToString("N") + ".pdf");

        try {
            File.WriteAllBytes(path, bytes);

            PdfDocumentPreflight fromPath = PdfInspector.Preflight(path);
            using var stream = new MemoryStream(bytes);
            PdfDocumentPreflight fromStream = PdfInspector.Preflight(stream);

            Assert.True(fromPath.CanRewrite);
            Assert.True(fromStream.CanRewrite);
            Assert.Empty(fromPath.ReadBlockers);
            Assert.Empty(fromStream.ReadBlockers);
            Assert.Empty(fromPath.RewriteBlockers);
            Assert.Empty(fromStream.RewriteBlockers);
            Assert.Equal(2, fromPath.DocumentInfo!.PageCount);
            Assert.Equal(fromPath.DocumentInfo.PageCount, fromStream.DocumentInfo!.PageCount);
        } finally {
            if (File.Exists(path)) {
                File.Delete(path);
            }
        }
    }

    [Fact]
    public void Preflight_RejectsInvalidInputs() {
        Assert.Throws<ArgumentNullException>(() => PdfInspector.Preflight((byte[])null!));
        Assert.Throws<ArgumentNullException>(() => PdfInspector.Preflight((string)null!));
        Assert.Throws<ArgumentException>(() => PdfInspector.Preflight(" "));
        Assert.Throws<ArgumentNullException>(() => PdfInspector.Preflight((Stream)null!));
        Assert.Throws<ArgumentException>(() => PdfInspector.Preflight(new WriteOnlyStream()));
    }

    private static byte[] BuildTwoPagePdf() {
        var doc = PdfDoc.Create()
            .Meta(
                title: "Inspection sample",
                author: "OfficeIMO",
                subject: "Roadmap",
                keywords: "pdf,inspect");

        doc.Compose(compose => {
            compose.Page(page => {
                page.Size(PageSizes.A4);
                page.Content(content => content.Column(column => column.Item().Paragraph(p => p.Text("First page."))));
            });

            compose.Page(page => {
                page.Size(new PageSize(792, 612));
                page.Content(content => content.Column(column => column.Item().Paragraph(p => p.Text("Landscape page."))));
            });
        });

        return doc.ToBytes();
    }

    private static void AssertNamedDestination(PdfDocumentInfo info, string name, int pageNumber, double destinationTop) {
        Assert.Equal(1, info.NamedDestinationCount);
        Assert.Equal(new[] { name }, info.NamedDestinationNames);

        PdfNamedDestination destination = Assert.Single(info.NamedDestinations);
        Assert.Equal(name, destination.Name);
        Assert.Equal(pageNumber, destination.PageNumber);
        Assert.Equal(destinationTop, destination.DestinationTop);
    }

    private static void AssertOpenAction(PdfDocumentInfo info, string actionType, int pageNumber, double? destinationTop) {
        Assert.True(info.HasReadableOpenAction);
        Assert.NotNull(info.OpenAction);
        Assert.Equal(actionType, info.OpenAction!.ActionType);
        Assert.Equal(pageNumber, info.OpenAction.PageNumber);
        Assert.Equal(destinationTop, info.OpenAction.DestinationTop);
    }

    private static void AssertViewerPreferences(PdfDocumentInfo info) {
        Assert.True(info.HasReadableViewerPreferences);
        Assert.NotNull(info.ViewerPreferences);
        Assert.Equal(2, info.ViewerPreferences!.Count);
        Assert.Equal("true", info.ViewerPreferences.GetValue("HideToolbar"));
        Assert.Equal("true", info.ViewerPreferences.GetValue("DisplayDocTitle"));
        Assert.True(info.ViewerPreferences.GetBoolean("HideToolbar"));
        Assert.True(info.ViewerPreferences.GetBoolean("DisplayDocTitle"));
        Assert.Null(info.ViewerPreferences.GetValue("Missing"));
        Assert.Null(info.ViewerPreferences.GetBoolean("Missing"));
    }

    private static void AssertPageLabel(PdfDocumentInfo info, int startPageIndex, int startPageNumber, string? style, string? prefix, int? startNumber) {
        Assert.True(info.HasReadablePageLabels);
        Assert.Equal(1, info.PageLabelCount);

        PdfPageLabel label = Assert.Single(info.PageLabels);
        Assert.Equal(startPageIndex, label.StartPageIndex);
        Assert.Equal(startPageNumber, label.StartPageNumber);
        Assert.Equal(style, label.Style);
        Assert.Equal(prefix, label.Prefix);
        Assert.Equal(startNumber, label.StartNumber);
    }

    private static void AssertRewriteBlocker(PdfDocumentPreflight report, PdfRewriteBlockerKind kind, string message) {
        PdfRewriteBlocker? blocker = null;
        for (int i = 0; i < report.RewriteBlockers.Count; i++) {
            if (report.RewriteBlockers[i].Kind == kind) {
                blocker = report.RewriteBlockers[i];
                break;
            }
        }

        Assert.NotNull(blocker);
        Assert.Equal(message, blocker!.Message);
        Assert.True(report.HasRewriteBlocker(kind));
    }

    private static void AssertReadBlocker(PdfDocumentPreflight report, PdfReadBlockerKind kind, string message) {
        PdfReadBlocker? blocker = null;
        for (int i = 0; i < report.ReadBlockers.Count; i++) {
            if (report.ReadBlockers[i].Kind == kind) {
                blocker = report.ReadBlockers[i];
                break;
            }
        }

        Assert.NotNull(blocker);
        Assert.Equal(message, blocker!.Message);
        Assert.True(report.HasReadBlocker(kind));
    }

    private static byte[] BuildAnnotatedPdf() {
        return PdfDoc.Create()
            .Paragraph(p => p.Link("OfficeIMO link", "https://evotec.xyz"))
            .ToBytes();
    }

    private static byte[] BuildOutlinePdf() {
        return PdfDoc.Create(new PdfOptions { CreateOutlineFromHeadings = true })
            .H1("Executive summary")
            .Paragraph(p => p.Text("Outline sample"))
            .ToBytes();
    }

    private static byte[] BuildComplexOutlinePdf() {
        string pdf = string.Join("\n", new[] {
            "%PDF-1.4",
            "1 0 obj",
            "<< /Type /Catalog /Pages 2 0 R /Outlines 5 0 R /PageMode /UseOutlines >>",
            "endobj",
            "2 0 obj",
            "<< /Type /Pages /Count 1 /Kids [3 0 R] >>",
            "endobj",
            "3 0 obj",
            "<< /Type /Page /Parent 2 0 R /MediaBox [0 0 200 200] /Contents 4 0 R >>",
            "endobj",
            "4 0 obj",
            "<< /Length 0 >>",
            "stream",
            "",
            "endstream",
            "endobj",
            "5 0 obj",
            "<< /Type /Outlines /First 6 0 R /Last 6 0 R /Count 1 >>",
            "endobj",
            "6 0 obj",
            "<< /Title (Chapter 1) /Parent 5 0 R /A << /S /GoTo /D [3 0 R /XYZ 0 200 0] >> >>",
            "endobj",
            "trailer",
            "<< /Root 1 0 R /Size 7 >>",
            "%%EOF"
        });

        return System.Text.Encoding.ASCII.GetBytes(pdf);
    }

    private static byte[] BuildCatalogViewSettingPdf() {
        string pdf = string.Join("\n", new[] {
            "%PDF-1.4",
            "1 0 obj",
            "<< /Type /Catalog /Pages 2 0 R /PageMode /FullScreen /PageLayout /TwoColumnLeft >>",
            "endobj",
            "2 0 obj",
            "<< /Type /Pages /Count 1 /Kids [3 0 R] >>",
            "endobj",
            "3 0 obj",
            "<< /Type /Page /Parent 2 0 R /MediaBox [0 0 200 200] /Contents 4 0 R >>",
            "endobj",
            "4 0 obj",
            "<< /Length 0 >>",
            "stream",
            "",
            "endstream",
            "endobj",
            "trailer",
            "<< /Root 1 0 R /Size 5 >>",
            "%%EOF"
        });

        return System.Text.Encoding.ASCII.GetBytes(pdf);
    }

    private static byte[] BuildCatalogIdentityPdf() {
        string pdf = string.Join("\n", new[] {
            "%PDF-1.4",
            "1 0 obj",
            "<< /Type /Catalog /Pages 2 0 R /Version /1.7 /Lang (pl-PL) >>",
            "endobj",
            "2 0 obj",
            "<< /Type /Pages /Count 1 /Kids [3 0 R] >>",
            "endobj",
            "3 0 obj",
            "<< /Type /Page /Parent 2 0 R /MediaBox [0 0 200 200] /Contents 4 0 R >>",
            "endobj",
            "4 0 obj",
            "<< /Length 0 >>",
            "stream",
            "",
            "endstream",
            "endobj",
            "trailer",
            "<< /Root 1 0 R /Size 5 >>",
            "%%EOF"
        });

        return System.Text.Encoding.ASCII.GetBytes(pdf);
    }

    private static byte[] BuildSignedPdfMarker() {
        string pdf = string.Join("\n", new[] {
            "%PDF-1.4",
            "1 0 obj",
            "<< /Type /Catalog /Pages 2 0 R /AcroForm 6 0 R >>",
            "endobj",
            "2 0 obj",
            "<< /Type /Pages /Count 1 /Kids [3 0 R] >>",
            "endobj",
            "3 0 obj",
            "<< /Type /Page /Parent 2 0 R /MediaBox [0 0 200 200] /Contents 4 0 R >>",
            "endobj",
            "4 0 obj",
            "<< /Length 0 >>",
            "stream",
            "",
            "endstream",
            "endobj",
            "5 0 obj",
            "<< /Type /Sig /ByteRange [0 0 0 0] /Contents <> >>",
            "endobj",
            "6 0 obj",
            "<< /SigFlags 3 /Fields [7 0 R] >>",
            "endobj",
            "7 0 obj",
            "<< /FT /Sig /V 5 0 R >>",
            "endobj",
            "trailer",
            "<< /Root 1 0 R /Size 8 >>",
            "%%EOF"
        });

        return System.Text.Encoding.ASCII.GetBytes(pdf);
    }

    private static byte[] BuildEncryptedPdfMarker() {
        string pdf = string.Join("\n", new[] {
            "%PDF-1.7",
            "1 0 obj",
            "<< /Type /Catalog /Pages 2 0 R >>",
            "endobj",
            "trailer",
            "<< /Root 1 0 R /Size 2 /Encrypt 3 0 R >>",
            "%%EOF"
        });

        return System.Text.Encoding.ASCII.GetBytes(pdf);
    }

    private static byte[] BuildNoPagesPdf() {
        string pdf = string.Join("\n", new[] {
            "%PDF-1.4",
            "1 0 obj",
            "<< /Type /Catalog /Pages 2 0 R >>",
            "endobj",
            "2 0 obj",
            "<< /Type /Pages /Count 0 /Kids [] >>",
            "endobj",
            "trailer",
            "<< /Root 1 0 R /Size 3 >>",
            "%%EOF"
        });

        return System.Text.Encoding.ASCII.GetBytes(pdf);
    }

    private static byte[] BuildFormPdfMarker() {
        string pdf = string.Join("\n", new[] {
            "%PDF-1.4",
            "1 0 obj",
            "<< /Type /Catalog /Pages 2 0 R /AcroForm 5 0 R >>",
            "endobj",
            "2 0 obj",
            "<< /Type /Pages /Count 1 /Kids [3 0 R] >>",
            "endobj",
            "3 0 obj",
            "<< /Type /Page /Parent 2 0 R /MediaBox [0 0 200 200] /Contents 4 0 R >>",
            "endobj",
            "4 0 obj",
            "<< /Length 0 >>",
            "stream",
            "",
            "endstream",
            "endobj",
            "5 0 obj",
            "<< /Fields [6 0 R] >>",
            "endobj",
            "6 0 obj",
            "<< /FT /Tx /T (Name) /V (OfficeIMO) >>",
            "endobj",
            "trailer",
            "<< /Root 1 0 R /Size 7 >>",
            "%%EOF"
        });

        return System.Text.Encoding.ASCII.GetBytes(pdf);
    }

    private static byte[] BuildHierarchicalFormPdf() {
        string pdf = string.Join("\n", new[] {
            "%PDF-1.4",
            "1 0 obj",
            "<< /Type /Catalog /Pages 2 0 R /AcroForm 5 0 R >>",
            "endobj",
            "2 0 obj",
            "<< /Type /Pages /Count 1 /Kids [3 0 R] >>",
            "endobj",
            "3 0 obj",
            "<< /Type /Page /Parent 2 0 R /MediaBox [0 0 200 200] /Contents 4 0 R >>",
            "endobj",
            "4 0 obj",
            "<< /Length 0 >>",
            "stream",
            "",
            "endstream",
            "endobj",
            "5 0 obj",
            "<< /Fields [6 0 R 8 0 R] >>",
            "endobj",
            "6 0 obj",
            "<< /T (Person) /Kids [7 0 R] >>",
            "endobj",
            "7 0 obj",
            "<< /FT /Tx /T (Name) /TU (Display name) /TM (ExportName) /V (OfficeIMO) /Ff 1 >>",
            "endobj",
            "8 0 obj",
            "<< /FT /Btn /T (AcceptTerms) /V /Yes >>",
            "endobj",
            "trailer",
            "<< /Root 1 0 R /Size 9 >>",
            "%%EOF"
        });

        return System.Text.Encoding.ASCII.GetBytes(pdf);
    }

    private static byte[] BuildPageLabelPdf() {
        string pdf = string.Join("\n", new[] {
            "%PDF-1.4",
            "1 0 obj",
            "<< /Type /Catalog /Pages 2 0 R /PageLabels 5 0 R >>",
            "endobj",
            "2 0 obj",
            "<< /Type /Pages /Count 1 /Kids [3 0 R] >>",
            "endobj",
            "3 0 obj",
            "<< /Type /Page /Parent 2 0 R /MediaBox [0 0 200 200] /Contents 4 0 R >>",
            "endobj",
            "4 0 obj",
            "<< /Length 0 >>",
            "stream",
            "",
            "endstream",
            "endobj",
            "5 0 obj",
            "<< /Nums [0 << /S /D /St 1 >>] >>",
            "endobj",
            "trailer",
            "<< /Root 1 0 R /Size 6 >>",
            "%%EOF"
        });

        return System.Text.Encoding.ASCII.GetBytes(pdf);
    }

    private static byte[] BuildComplexPageLabelPdf() {
        string pdf = string.Join("\n", new[] {
            "%PDF-1.4",
            "1 0 obj",
            "<< /Type /Catalog /Pages 2 0 R /PageLabels 5 0 R >>",
            "endobj",
            "2 0 obj",
            "<< /Type /Pages /Count 1 /Kids [3 0 R] >>",
            "endobj",
            "3 0 obj",
            "<< /Type /Page /Parent 2 0 R /MediaBox [0 0 200 200] /Contents 4 0 R >>",
            "endobj",
            "4 0 obj",
            "<< /Length 0 >>",
            "stream",
            "",
            "endstream",
            "endobj",
            "5 0 obj",
            "<< /Kids [6 0 R] >>",
            "endobj",
            "6 0 obj",
            "<< /Nums [0 << /S /D /St 1 >>] >>",
            "endobj",
            "trailer",
            "<< /Root 1 0 R /Size 7 >>",
            "%%EOF"
        });

        return System.Text.Encoding.ASCII.GetBytes(pdf);
    }

    private static byte[] BuildNamedDestinationPdf() {
        string pdf = string.Join("\n", new[] {
            "%PDF-1.4",
            "1 0 obj",
            "<< /Type /Catalog /Pages 2 0 R /Dests 5 0 R >>",
            "endobj",
            "2 0 obj",
            "<< /Type /Pages /Count 1 /Kids [3 0 R] >>",
            "endobj",
            "3 0 obj",
            "<< /Type /Page /Parent 2 0 R /MediaBox [0 0 200 200] /Contents 4 0 R >>",
            "endobj",
            "4 0 obj",
            "<< /Length 0 >>",
            "stream",
            "",
            "endstream",
            "endobj",
            "5 0 obj",
            "<< /Chapter1 [3 0 R /XYZ 0 200 0] >>",
            "endobj",
            "trailer",
            "<< /Root 1 0 R /Size 6 >>",
            "%%EOF"
        });

        return System.Text.Encoding.ASCII.GetBytes(pdf);
    }

    private static byte[] BuildNamedDestinationNameTreePdf() {
        string pdf = string.Join("\n", new[] {
            "%PDF-1.4",
            "1 0 obj",
            "<< /Type /Catalog /Pages 2 0 R /Names << /Dests << /Names [(Chapter1) [3 0 R /XYZ 0 200 0]] >> >> >>",
            "endobj",
            "2 0 obj",
            "<< /Type /Pages /Count 1 /Kids [3 0 R] >>",
            "endobj",
            "3 0 obj",
            "<< /Type /Page /Parent 2 0 R /MediaBox [0 0 200 200] /Contents 4 0 R >>",
            "endobj",
            "4 0 obj",
            "<< /Length 0 >>",
            "stream",
            "",
            "endstream",
            "endobj",
            "trailer",
            "<< /Root 1 0 R /Size 5 >>",
            "%%EOF"
        });

        return System.Text.Encoding.ASCII.GetBytes(pdf);
    }

    private static byte[] BuildComplexNamedDestinationNameTreePdf() {
        string pdf = string.Join("\n", new[] {
            "%PDF-1.4",
            "1 0 obj",
            "<< /Type /Catalog /Pages 2 0 R /Names << /Dests << /Kids [5 0 R] >> >> >>",
            "endobj",
            "2 0 obj",
            "<< /Type /Pages /Count 1 /Kids [3 0 R] >>",
            "endobj",
            "3 0 obj",
            "<< /Type /Page /Parent 2 0 R /MediaBox [0 0 200 200] /Contents 4 0 R >>",
            "endobj",
            "4 0 obj",
            "<< /Length 0 >>",
            "stream",
            "",
            "endstream",
            "endobj",
            "5 0 obj",
            "<< /Names [(Chapter1) [3 0 R /XYZ 0 200 0]] >>",
            "endobj",
            "trailer",
            "<< /Root 1 0 R /Size 6 >>",
            "%%EOF"
        });

        return System.Text.Encoding.ASCII.GetBytes(pdf);
    }

    private static byte[] BuildOpenActionPdf() {
        string pdf = string.Join("\n", new[] {
            "%PDF-1.4",
            "1 0 obj",
            "<< /Type /Catalog /Pages 2 0 R /OpenAction 5 0 R >>",
            "endobj",
            "2 0 obj",
            "<< /Type /Pages /Count 1 /Kids [3 0 R] >>",
            "endobj",
            "3 0 obj",
            "<< /Type /Page /Parent 2 0 R /MediaBox [0 0 200 200] /Contents 4 0 R >>",
            "endobj",
            "4 0 obj",
            "<< /Length 0 >>",
            "stream",
            "",
            "endstream",
            "endobj",
            "5 0 obj",
            "[3 0 R /Fit]",
            "endobj",
            "trailer",
            "<< /Root 1 0 R /Size 6 >>",
            "%%EOF"
        });

        return System.Text.Encoding.ASCII.GetBytes(pdf);
    }

    private static byte[] BuildOpenActionDictionaryPdf() {
        string pdf = string.Join("\n", new[] {
            "%PDF-1.4",
            "1 0 obj",
            "<< /Type /Catalog /Pages 2 0 R /OpenAction 5 0 R >>",
            "endobj",
            "2 0 obj",
            "<< /Type /Pages /Count 1 /Kids [3 0 R] >>",
            "endobj",
            "3 0 obj",
            "<< /Type /Page /Parent 2 0 R /MediaBox [0 0 200 200] /Contents 4 0 R >>",
            "endobj",
            "4 0 obj",
            "<< /Length 0 >>",
            "stream",
            "",
            "endstream",
            "endobj",
            "5 0 obj",
            "<< /S /GoTo /D [3 0 R /Fit] >>",
            "endobj",
            "trailer",
            "<< /Root 1 0 R /Size 6 >>",
            "%%EOF"
        });

        return System.Text.Encoding.ASCII.GetBytes(pdf);
    }

    private static byte[] BuildComplexOpenActionDictionaryPdf() {
        string pdf = string.Join("\n", new[] {
            "%PDF-1.4",
            "1 0 obj",
            "<< /Type /Catalog /Pages 2 0 R /OpenAction 5 0 R >>",
            "endobj",
            "2 0 obj",
            "<< /Type /Pages /Count 1 /Kids [3 0 R] >>",
            "endobj",
            "3 0 obj",
            "<< /Type /Page /Parent 2 0 R /MediaBox [0 0 200 200] /Contents 4 0 R >>",
            "endobj",
            "4 0 obj",
            "<< /Length 0 >>",
            "stream",
            "",
            "endstream",
            "endobj",
            "5 0 obj",
            "<< /S /URI /URI (https://example.com) >>",
            "endobj",
            "trailer",
            "<< /Root 1 0 R /Size 6 >>",
            "%%EOF"
        });

        return System.Text.Encoding.ASCII.GetBytes(pdf);
    }

    private static byte[] BuildViewerPreferencePdf() {
        string pdf = string.Join("\n", new[] {
            "%PDF-1.4",
            "1 0 obj",
            "<< /Type /Catalog /Pages 2 0 R /ViewerPreferences 5 0 R >>",
            "endobj",
            "2 0 obj",
            "<< /Type /Pages /Count 1 /Kids [3 0 R] >>",
            "endobj",
            "3 0 obj",
            "<< /Type /Page /Parent 2 0 R /MediaBox [0 0 200 200] /Contents 4 0 R >>",
            "endobj",
            "4 0 obj",
            "<< /Length 0 >>",
            "stream",
            "",
            "endstream",
            "endobj",
            "5 0 obj",
            "<< /HideToolbar true /DisplayDocTitle true >>",
            "endobj",
            "trailer",
            "<< /Root 1 0 R /Size 6 >>",
            "%%EOF"
        });

        return System.Text.Encoding.ASCII.GetBytes(pdf);
    }

    private static byte[] BuildComplexViewerPreferencePdf() {
        string pdf = string.Join("\n", new[] {
            "%PDF-1.4",
            "1 0 obj",
            "<< /Type /Catalog /Pages 2 0 R /ViewerPreferences 5 0 R >>",
            "endobj",
            "2 0 obj",
            "<< /Type /Pages /Count 1 /Kids [3 0 R] >>",
            "endobj",
            "3 0 obj",
            "<< /Type /Page /Parent 2 0 R /MediaBox [0 0 200 200] /Contents 4 0 R >>",
            "endobj",
            "4 0 obj",
            "<< /Length 0 >>",
            "stream",
            "",
            "endstream",
            "endobj",
            "5 0 obj",
            "<< /HideToolbar true /ViewArea 3 0 R >>",
            "endobj",
            "trailer",
            "<< /Root 1 0 R /Size 6 >>",
            "%%EOF"
        });

        return System.Text.Encoding.ASCII.GetBytes(pdf);
    }

    private static byte[] BuildTaggedPdf() {
        string pdf = string.Join("\n", new[] {
            "%PDF-1.4",
            "1 0 obj",
            "<< /Type /Catalog /Pages 2 0 R /MarkInfo << /Marked true >> /StructTreeRoot 5 0 R >>",
            "endobj",
            "2 0 obj",
            "<< /Type /Pages /Count 1 /Kids [3 0 R] >>",
            "endobj",
            "3 0 obj",
            "<< /Type /Page /Parent 2 0 R /MediaBox [0 0 200 200] /Contents 4 0 R /StructParents 0 >>",
            "endobj",
            "4 0 obj",
            "<< /Length 0 >>",
            "stream",
            "",
            "endstream",
            "endobj",
            "5 0 obj",
            "<< /Type /StructTreeRoot /K [6 0 R] /ParentTree 7 0 R >>",
            "endobj",
            "6 0 obj",
            "<< /Type /StructElem /S /Document /P 5 0 R >>",
            "endobj",
            "7 0 obj",
            "<< /Nums [0 [6 0 R]] >>",
            "endobj",
            "trailer",
            "<< /Root 1 0 R /Size 8 >>",
            "%%EOF"
        });

        return System.Text.Encoding.ASCII.GetBytes(pdf);
    }

    private static byte[] BuildXmpMetadataPdf() {
        string pdf = string.Join("\n", new[] {
            "%PDF-1.4",
            "1 0 obj",
            "<< /Type /Catalog /Pages 2 0 R /Metadata 5 0 R >>",
            "endobj",
            "2 0 obj",
            "<< /Type /Pages /Count 1 /Kids [3 0 R] >>",
            "endobj",
            "3 0 obj",
            "<< /Type /Page /Parent 2 0 R /MediaBox [0 0 200 200] /Contents 4 0 R >>",
            "endobj",
            "4 0 obj",
            "<< /Length 0 >>",
            "stream",
            "",
            "endstream",
            "endobj",
            "5 0 obj",
            "<< /Type /Metadata /Subtype /XML /Length 12 >>",
            "stream",
            "<x:xmpmeta/>",
            "endstream",
            "endobj",
            "trailer",
            "<< /Root 1 0 R /Size 6 >>",
            "%%EOF"
        });

        return System.Text.Encoding.ASCII.GetBytes(pdf);
    }

    private static byte[] BuildComplexXmpMetadataPdf() {
        string pdf = string.Join("\n", new[] {
            "%PDF-1.4",
            "1 0 obj",
            "<< /Type /Catalog /Pages 2 0 R /Metadata 5 0 R >>",
            "endobj",
            "2 0 obj",
            "<< /Type /Pages /Count 1 /Kids [3 0 R] >>",
            "endobj",
            "3 0 obj",
            "<< /Type /Page /Parent 2 0 R /MediaBox [0 0 200 200] /Contents 4 0 R >>",
            "endobj",
            "4 0 obj",
            "<< /Length 0 >>",
            "stream",
            "",
            "endstream",
            "endobj",
            "5 0 obj",
            "<< /Type /Metadata /Subtype /XML /Source 3 0 R /Length 12 >>",
            "stream",
            "<x:xmpmeta/>",
            "endstream",
            "endobj",
            "trailer",
            "<< /Root 1 0 R /Size 6 >>",
            "%%EOF"
        });

        return System.Text.Encoding.ASCII.GetBytes(pdf);
    }

    private static byte[] BuildCatalogUriPdf() {
        string pdf = string.Join("\n", new[] {
            "%PDF-1.4",
            "1 0 obj",
            "<< /Type /Catalog /Pages 2 0 R /URI << /Base (https://evotec.xyz/docs/) >> >>",
            "endobj",
            "2 0 obj",
            "<< /Type /Pages /Count 1 /Kids [3 0 R] >>",
            "endobj",
            "3 0 obj",
            "<< /Type /Page /Parent 2 0 R /MediaBox [0 0 200 200] /Contents 4 0 R >>",
            "endobj",
            "4 0 obj",
            "<< /Length 0 >>",
            "stream",
            "",
            "endstream",
            "endobj",
            "trailer",
            "<< /Root 1 0 R /Size 5 >>",
            "%%EOF"
        });

        return System.Text.Encoding.ASCII.GetBytes(pdf);
    }

    private static byte[] BuildComplexCatalogUriPdf() {
        string pdf = string.Join("\n", new[] {
            "%PDF-1.4",
            "1 0 obj",
            "<< /Type /Catalog /Pages 2 0 R /URI << /Base (https://evotec.xyz/docs/) /SourcePage 3 0 R >> >>",
            "endobj",
            "2 0 obj",
            "<< /Type /Pages /Count 1 /Kids [3 0 R] >>",
            "endobj",
            "3 0 obj",
            "<< /Type /Page /Parent 2 0 R /MediaBox [0 0 200 200] /Contents 4 0 R >>",
            "endobj",
            "4 0 obj",
            "<< /Length 0 >>",
            "stream",
            "",
            "endstream",
            "endobj",
            "trailer",
            "<< /Root 1 0 R /Size 5 >>",
            "%%EOF"
        });

        return System.Text.Encoding.ASCII.GetBytes(pdf);
    }

    private static byte[] BuildOutputIntentPdf() {
        string pdf = string.Join("\n", new[] {
            "%PDF-1.4",
            "1 0 obj",
            "<< /Type /Catalog /Pages 2 0 R /OutputIntents [5 0 R] >>",
            "endobj",
            "2 0 obj",
            "<< /Type /Pages /Count 1 /Kids [3 0 R] >>",
            "endobj",
            "3 0 obj",
            "<< /Type /Page /Parent 2 0 R /MediaBox [0 0 200 200] /Contents 4 0 R >>",
            "endobj",
            "4 0 obj",
            "<< /Length 0 >>",
            "stream",
            "",
            "endstream",
            "endobj",
            "5 0 obj",
            "<< /Type /OutputIntent /S /GTS_PDFA1 /OutputConditionIdentifier (sRGB) >>",
            "endobj",
            "trailer",
            "<< /Root 1 0 R /Size 6 >>",
            "%%EOF"
        });

        return System.Text.Encoding.ASCII.GetBytes(pdf);
    }

    private static byte[] BuildComplexOutputIntentPdf() {
        string pdf = string.Join("\n", new[] {
            "%PDF-1.4",
            "1 0 obj",
            "<< /Type /Catalog /Pages 2 0 R /OutputIntents [5 0 R] >>",
            "endobj",
            "2 0 obj",
            "<< /Type /Pages /Count 1 /Kids [3 0 R] >>",
            "endobj",
            "3 0 obj",
            "<< /Type /Page /Parent 2 0 R /MediaBox [0 0 200 200] /Contents 4 0 R >>",
            "endobj",
            "4 0 obj",
            "<< /Length 0 >>",
            "stream",
            "",
            "endstream",
            "endobj",
            "5 0 obj",
            "<< /Type /OutputIntent /S /GTS_PDFA1 /OutputConditionIdentifier (sRGB) /DestOutputProfile 3 0 R >>",
            "endobj",
            "trailer",
            "<< /Root 1 0 R /Size 6 >>",
            "%%EOF"
        });

        return System.Text.Encoding.ASCII.GetBytes(pdf);
    }

    private static byte[] BuildEmbeddedFilePdf() {
        string pdf = string.Join("\n", new[] {
            "%PDF-1.4",
            "1 0 obj",
            "<< /Type /Catalog /Pages 2 0 R /Names << /EmbeddedFiles << /Names [(note.txt) 5 0 R] >> >> >>",
            "endobj",
            "2 0 obj",
            "<< /Type /Pages /Count 1 /Kids [3 0 R] >>",
            "endobj",
            "3 0 obj",
            "<< /Type /Page /Parent 2 0 R /MediaBox [0 0 200 200] /Contents 4 0 R >>",
            "endobj",
            "4 0 obj",
            "<< /Length 0 >>",
            "stream",
            "",
            "endstream",
            "endobj",
            "5 0 obj",
            "<< /Type /Filespec /F (note.txt) /EF << /F 6 0 R >> >>",
            "endobj",
            "6 0 obj",
            "<< /Type /EmbeddedFile /Length 4 >>",
            "stream",
            "note",
            "endstream",
            "endobj",
            "trailer",
            "<< /Root 1 0 R /Size 7 >>",
            "%%EOF"
        });

        return System.Text.Encoding.ASCII.GetBytes(pdf);
    }

    private static byte[] BuildAssociatedFilePdf() {
        string pdf = string.Join("\n", new[] {
            "%PDF-1.7",
            "1 0 obj",
            "<< /Type /Catalog /Pages 2 0 R /AF [5 0 R] >>",
            "endobj",
            "2 0 obj",
            "<< /Type /Pages /Count 1 /Kids [3 0 R] >>",
            "endobj",
            "3 0 obj",
            "<< /Type /Page /Parent 2 0 R /MediaBox [0 0 200 200] /Contents 4 0 R >>",
            "endobj",
            "4 0 obj",
            "<< /Length 0 >>",
            "stream",
            "",
            "endstream",
            "endobj",
            "5 0 obj",
            "<< /Type /Filespec /F (data.xml) /AFRelationship /Data /EF << /F 6 0 R >> >>",
            "endobj",
            "6 0 obj",
            "<< /Type /EmbeddedFile /Subtype /text#2Fxml /Length 4 >>",
            "stream",
            "data",
            "endstream",
            "endobj",
            "trailer",
            "<< /Root 1 0 R /Size 7 >>",
            "%%EOF"
        });

        return System.Text.Encoding.ASCII.GetBytes(pdf);
    }

    private static byte[] BuildCombinedDestinationAndEmbeddedFileNameTreePdf() {
        string pdf = string.Join("\n", new[] {
            "%PDF-1.4",
            "1 0 obj",
            "<< /Type /Catalog /Pages 2 0 R /Names << /Dests << /Names [(Chapter1) [3 0 R /XYZ 0 200 0]] >> /EmbeddedFiles << /Names [(note.txt) 5 0 R] >> >> >>",
            "endobj",
            "2 0 obj",
            "<< /Type /Pages /Count 1 /Kids [3 0 R] >>",
            "endobj",
            "3 0 obj",
            "<< /Type /Page /Parent 2 0 R /MediaBox [0 0 200 200] /Contents 4 0 R >>",
            "endobj",
            "4 0 obj",
            "<< /Length 0 >>",
            "stream",
            "",
            "endstream",
            "endobj",
            "5 0 obj",
            "<< /Type /Filespec /F (note.txt) /EF << /F 6 0 R >> >>",
            "endobj",
            "6 0 obj",
            "<< /Type /EmbeddedFile /Length 4 >>",
            "stream",
            "note",
            "endstream",
            "endobj",
            "trailer",
            "<< /Root 1 0 R /Size 7 >>",
            "%%EOF"
        });

        return System.Text.Encoding.ASCII.GetBytes(pdf);
    }

    private static byte[] BuildUnsupportedCatalogNameTreePdf() {
        string pdf = string.Join("\n", new[] {
            "%PDF-1.4",
            "1 0 obj",
            "<< /Type /Catalog /Pages 2 0 R /Names << /Templates << /Names [(Layout) 5 0 R] >> >> >>",
            "endobj",
            "2 0 obj",
            "<< /Type /Pages /Count 1 /Kids [3 0 R] >>",
            "endobj",
            "3 0 obj",
            "<< /Type /Page /Parent 2 0 R /MediaBox [0 0 200 200] /Contents 4 0 R >>",
            "endobj",
            "4 0 obj",
            "<< /Length 0 >>",
            "stream",
            "",
            "endstream",
            "endobj",
            "5 0 obj",
            "<< /Type /Template /Name (Layout) >>",
            "endobj",
            "trailer",
            "<< /Root 1 0 R /Size 6 >>",
            "%%EOF"
        });

        return System.Text.Encoding.ASCII.GetBytes(pdf);
    }

    private static byte[] BuildComplexEmbeddedFilePdf() {
        string pdf = string.Join("\n", new[] {
            "%PDF-1.4",
            "1 0 obj",
            "<< /Type /Catalog /Pages 2 0 R /Names << /EmbeddedFiles << /Names [(note.txt) 5 0 R] >> >> >>",
            "endobj",
            "2 0 obj",
            "<< /Type /Pages /Count 1 /Kids [3 0 R] >>",
            "endobj",
            "3 0 obj",
            "<< /Type /Page /Parent 2 0 R /MediaBox [0 0 200 200] /Contents 4 0 R >>",
            "endobj",
            "4 0 obj",
            "<< /Length 0 >>",
            "stream",
            "",
            "endstream",
            "endobj",
            "5 0 obj",
            "<< /Type /Filespec /F (note.txt) /EF << /F 3 0 R >> >>",
            "endobj",
            "trailer",
            "<< /Root 1 0 R /Size 6 >>",
            "%%EOF"
        });

        return System.Text.Encoding.ASCII.GetBytes(pdf);
    }

    private static byte[] BuildComplexAssociatedFilePdf() {
        string pdf = string.Join("\n", new[] {
            "%PDF-1.7",
            "1 0 obj",
            "<< /Type /Catalog /Pages 2 0 R /AF [5 0 R] >>",
            "endobj",
            "2 0 obj",
            "<< /Type /Pages /Count 1 /Kids [3 0 R] >>",
            "endobj",
            "3 0 obj",
            "<< /Type /Page /Parent 2 0 R /MediaBox [0 0 200 200] /Contents 4 0 R >>",
            "endobj",
            "4 0 obj",
            "<< /Length 0 >>",
            "stream",
            "",
            "endstream",
            "endobj",
            "5 0 obj",
            "<< /Type /Filespec /F (data.xml) /AFRelationship /Data /EF << /F 3 0 R >> >>",
            "endobj",
            "trailer",
            "<< /Root 1 0 R /Size 6 >>",
            "%%EOF"
        });

        return System.Text.Encoding.ASCII.GetBytes(pdf);
    }

    private static byte[] BuildOptionalContentPdf() {
        string pdf = string.Join("\n", new[] {
            "%PDF-1.5",
            "1 0 obj",
            "<< /Type /Catalog /Pages 2 0 R /OCProperties << /OCGs [5 0 R] /D << /ON [5 0 R] /Order [5 0 R] >> >> >>",
            "endobj",
            "2 0 obj",
            "<< /Type /Pages /Count 1 /Kids [3 0 R] >>",
            "endobj",
            "3 0 obj",
            "<< /Type /Page /Parent 2 0 R /MediaBox [0 0 200 200] /Contents 4 0 R >>",
            "endobj",
            "4 0 obj",
            "<< /Length 0 >>",
            "stream",
            "",
            "endstream",
            "endobj",
            "5 0 obj",
            "<< /Type /OCG /Name (Layer 1) >>",
            "endobj",
            "trailer",
            "<< /Root 1 0 R /Size 6 >>",
            "%%EOF"
        });

        return System.Text.Encoding.ASCII.GetBytes(pdf);
    }

    private static byte[] BuildComplexOptionalContentPdf() {
        string pdf = string.Join("\n", new[] {
            "%PDF-1.5",
            "1 0 obj",
            "<< /Type /Catalog /Pages 2 0 R /OCProperties << /OCGs [3 0 R] /D << /ON [3 0 R] /Order [3 0 R] >> >> >>",
            "endobj",
            "2 0 obj",
            "<< /Type /Pages /Count 1 /Kids [3 0 R] >>",
            "endobj",
            "3 0 obj",
            "<< /Type /Page /Parent 2 0 R /MediaBox [0 0 200 200] /Contents 4 0 R >>",
            "endobj",
            "4 0 obj",
            "<< /Length 0 >>",
            "stream",
            "",
            "endstream",
            "endobj",
            "trailer",
            "<< /Root 1 0 R /Size 5 >>",
            "%%EOF"
        });

        return System.Text.Encoding.ASCII.GetBytes(pdf);
    }

    private static byte[] BuildActiveContentPdf() {
        string pdf = string.Join("\n", new[] {
            "%PDF-1.4",
            "1 0 obj",
            "<< /Type /Catalog /Pages 2 0 R /Names << /JavaScript << /Names [(Open) 5 0 R] >> >> >>",
            "endobj",
            "2 0 obj",
            "<< /Type /Pages /Count 1 /Kids [3 0 R] >>",
            "endobj",
            "3 0 obj",
            "<< /Type /Page /Parent 2 0 R /MediaBox [0 0 200 200] /Contents 4 0 R >>",
            "endobj",
            "4 0 obj",
            "<< /Length 0 >>",
            "stream",
            "",
            "endstream",
            "endobj",
            "5 0 obj",
            "<< /S /JavaScript /JS (app.alert('OfficeIMO')) >>",
            "endobj",
            "trailer",
            "<< /Root 1 0 R /Size 6 >>",
            "%%EOF"
        });

        return System.Text.Encoding.ASCII.GetBytes(pdf);
    }

    private static byte[] CreateMinimalRgbPng() {
        return new byte[] {
            137, 80, 78, 71, 13, 10, 26, 10,
            0, 0, 0, 13,
            73, 72, 68, 82,
            0, 0, 0, 1,
            0, 0, 0, 1,
            8, 2, 0, 0, 0,
            0, 0, 0, 0,
            0, 0, 0, 12,
            73, 68, 65, 84,
            0x78, 0x9C, 0x63, 0xF8, 0xCF, 0xC0, 0x00, 0x00, 0x03, 0x01, 0x01, 0x00,
            0, 0, 0, 0,
            0, 0, 0, 0,
            73, 69, 78, 68,
            0, 0, 0, 0
        };
    }

    private sealed class WriteOnlyStream : MemoryStream {
        public override bool CanRead => false;
    }
}
