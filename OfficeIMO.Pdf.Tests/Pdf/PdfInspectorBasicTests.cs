using System;
using System.IO;
using System.Linq;
using OfficeIMO.Drawing;
using OfficeIMO.Pdf;
using Xunit;

namespace OfficeIMO.Tests.Pdf;

public partial class PdfInspectorTests {
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
        Assert.Equal(0, info.AnnotationActionTypeCount);
        Assert.Empty(info.AnnotationActionTypes);
        Assert.Empty(info.AnnotationsByActionType);
        Assert.Empty(info.GetAnnotationsByActionType("JavaScript"));
        Assert.Equal(0, info.LinkAnnotationCount);
        Assert.Equal(0, info.LinkUriCount);
        Assert.Equal(0, info.LinkDestinationCount);
        Assert.Equal(0, info.LinkDestinationPageNumberCount);
        Assert.Equal(0, info.LinkNamedActionCount);
        Assert.Equal(0, info.LinkRemoteFileCount);
        Assert.Empty(info.LinkAnnotations);
        Assert.Empty(info.LinkUris);
        Assert.Empty(info.LinkDestinationNames);
        Assert.Empty(info.LinkDestinationPageNumbers);
        Assert.Empty(info.LinkNamedActions);
        Assert.Empty(info.LinkRemoteFiles);
        Assert.Empty(info.LinkAnnotationsByUri);
        Assert.Empty(info.LinkAnnotationsByDestinationName);
        Assert.Empty(info.LinkAnnotationsByDestinationPageNumber);
        Assert.Empty(info.LinkAnnotationsByNamedAction);
        Assert.Empty(info.LinkAnnotationsByRemoteFile);
        Assert.Empty(info.GetLinkAnnotationsByUri("https://evotec.xyz/missing"));
        Assert.Empty(info.GetLinkAnnotationsByDestinationName("Missing"));
        Assert.Empty(info.GetLinkAnnotationsByDestinationPageNumber(1));
        Assert.Empty(info.GetLinkAnnotationsByNamedAction("NextPage"));
        Assert.Empty(info.GetLinkAnnotationsByRemoteFile("remote-report.pdf"));
        Assert.Throws<ArgumentOutOfRangeException>(() => info.GetLinkAnnotationsByDestinationPageNumber(0));
        Assert.False(info.HasNamedDestinations);
        Assert.Equal(0, info.NamedDestinationCount);
        Assert.Empty(info.NamedDestinations);
        Assert.Empty(info.NamedDestinationNames);
        Assert.False(info.HasCatalogActions);
        Assert.Equal(0, info.CatalogActionCount);
        Assert.Empty(info.CatalogActions);
        Assert.Empty(info.CatalogActionNames);
        Assert.Empty(info.CatalogActionTypes);
        Assert.Empty(info.CatalogActionSources);
        Assert.Empty(info.CatalogActionsByActionType);
        Assert.Empty(info.CatalogActionsBySource);
        Assert.Empty(info.GetCatalogActionsByActionType("JavaScript"));
        Assert.Empty(info.GetCatalogActionsBySource("OpenAction"));
        Assert.False(info.HasAttachments);
        Assert.Equal(0, info.AttachmentCount);
        Assert.Empty(info.Attachments);
        Assert.Empty(info.AttachmentNames);
        Assert.Empty(info.AttachmentFileNames);
        Assert.Empty(info.AttachmentSources);
        Assert.Empty(info.AttachmentsByName);
        Assert.Empty(info.AttachmentsByFileName);
        Assert.Empty(info.AttachmentsBySource);
        Assert.Empty(info.AttachmentsByRelationship);
        Assert.Empty(info.GetAttachmentsByName("note.txt"));
        Assert.Empty(info.GetAttachmentsByFileName("note.txt"));
        Assert.Empty(info.GetAttachmentsBySource("AF"));
        Assert.Empty(info.GetAttachmentsByRelationship(PdfAssociatedFileRelationship.Data));
        Assert.False(info.HasReadableOptionalContent);
        Assert.False(info.HasOptionalContentGroups);
        Assert.Equal(0, info.OptionalContentGroupCount);
        Assert.Null(info.OptionalContent);
        Assert.Empty(info.OptionalContentGroups);
        Assert.Empty(info.OptionalContentGroupNames);
        Assert.Empty(info.OptionalContentGroupsByName);
        Assert.Empty(info.GetOptionalContentGroupsByName("Layer 1"));
        Assert.False(info.HasPageActions);
        Assert.Equal(0, info.PageActionCount);
        Assert.Empty(info.PageActions);
        Assert.Empty(info.PageActionTypes);
        Assert.Empty(info.PageActionTriggerNames);
        Assert.Empty(info.PageActionPaths);
        Assert.Empty(info.PageActionsByActionType);
        Assert.Empty(info.PageActionsByTriggerName);
        Assert.Empty(info.PageActionsByActionPath);
        Assert.Empty(info.PageActionsByPageNumber);
        Assert.Empty(info.GetPageActionsByActionType("JavaScript"));
        Assert.Empty(info.GetPageActionsByTriggerName("O"));
        Assert.Empty(info.GetPageActionsByActionPath("O.Next"));
        Assert.Empty(info.GetPageActions(1));
        Assert.Throws<ArgumentOutOfRangeException>(() => info.GetPageActions(0));
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
        Assert.False(info.Pages[0].HasPageActions);
        Assert.Equal(0, info.Pages[0].PageActionCount);
        Assert.Empty(info.Pages[0].PageActions);

        Assert.Equal(2, info.Pages[1].PageNumber);
        Assert.Equal(792, info.Pages[1].Width);
        Assert.Equal(612, info.Pages[1].Height);
        Assert.True(info.Pages[1].IsLandscape);
        Assert.False(info.Pages[1].HasPageActions);
        Assert.Empty(info.Pages[1].PageActions);
    }

    [Fact]
    public void Inspect_ReadsPageGeometryAndPresentationMetadata() {
        PdfDocumentInfo info = PdfInspector.Inspect(PdfPageGeometrySupport.BuildPageGeometryPdf());

        PdfPageInfo page = Assert.Single(info.Pages);
        Assert.Equal(380, page.Width);
        Assert.Equal(260, page.Height);
        Assert.Equal(400, page.MediaBox!.Width);
        Assert.Equal(300, page.MediaBox.Height);
        Assert.Equal(10, page.CropBox!.Left);
        Assert.Equal(20, page.CropBox.Bottom);
        Assert.Same(page.CropBox, page.Geometry.EffectiveBox);
        Assert.Equal(5, page.BleedBox!.Left);
        Assert.Equal(280, page.BleedBox.Height);
        Assert.Equal(20, page.TrimBox!.Left);
        Assert.Equal(240, page.TrimBox.Height);
        Assert.Equal(25, page.ArtBox!.Left);
        Assert.Equal(230, page.ArtBox.Height);
        Assert.True(page.Geometry.HasNonDefaultBoundaryBoxes);
        Assert.Equal(2, page.UserUnit);
        Assert.Equal("S", page.TabOrder);
        Assert.Equal(5, page.DurationSeconds);
        Assert.True(page.Geometry.HasTransition);

        PdfPageTransition transition = page.Transition!;
        Assert.Equal("Fly", transition.Style);
        Assert.Equal(1.5, transition.DurationSeconds);
        Assert.Equal("H", transition.Dimension);
        Assert.Equal("I", transition.Motion);
        Assert.Equal(90, transition.Direction);
        Assert.Equal(0.75, transition.Scale);
        Assert.True(transition.IsFlyAreaOpaque);
        Assert.True(page.HasPageMetadata);
        Assert.Equal(5, page.Geometry.MetadataObjectNumber);
        Assert.True(page.HasPieceInfo);
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
    public void InspectPageRanges_ReturnsSelectedPagesInCallerOrder() {
        byte[] bytes = BuildTwoPagePdf();
        string path = Path.Combine(Path.GetTempPath(), "officeimo-pdf-inspect-ranges-" + Guid.NewGuid().ToString("N") + ".pdf");
        byte[] prefix = System.Text.Encoding.ASCII.GetBytes("prefix");

        try {
            File.WriteAllBytes(path, bytes);

            PdfDocumentInfo selected = PdfInspector.InspectPageRanges(bytes, PdfPageRange.ParseMany("2,1,2"));
            PdfDocumentInfo fromPath = PdfInspector.InspectPageRanges(path, PdfPageRange.From(2, 2));
            using var stream = new MemoryStream(prefix.Concat(bytes).ToArray());
            stream.Position = prefix.Length;
            PdfDocumentInfo fromStream = PdfInspector.InspectPageRanges(stream, PdfPageRange.From(1, 1));

            Assert.Equal(3, selected.PageCount);
            Assert.Equal(new[] { 2, 1, 2 }, selected.Pages.Select(page => page.PageNumber).ToArray());
            Assert.Equal(new[] { 792d, 595d, 792d }, selected.Pages.Select(page => page.Width).ToArray());
            Assert.Equal(2, Assert.Single(fromPath.Pages).PageNumber);
            Assert.Equal(1, Assert.Single(fromStream.Pages).PageNumber);
        } finally {
            if (File.Exists(path)) {
                File.Delete(path);
            }
        }
    }

    [Fact]
    public void InspectPageRanges_FiltersPageScopedObjectsToSelectedSourcePages() {
        byte[] pdf = BuildThreePageInspectMetadataPdf();

        PdfDocumentInfo selected = PdfInspector.InspectPageRanges(pdf, PdfPageRange.ParseMany("2,1,2"));

        Assert.Equal(new[] { 2, 1, 2 }, selected.Pages.Select(page => page.PageNumber).ToArray());
        Assert.Equal(new[] { "First", "Second" }, selected.NamedDestinationNames.OrderBy(name => name).ToArray());
        Assert.Equal(new[] { 1, 2 }, selected.NamedDestinations.Select(destination => destination.PageNumber!.Value).OrderBy(pageNumber => pageNumber).ToArray());
        Assert.Equal(new[] { "First outline", "Second outline" }, selected.Outlines.Select(outline => outline.Title).OrderBy(title => title).ToArray());
        Assert.Equal(new[] { 1, 2 }, selected.Outlines.Select(outline => outline.PageNumber!.Value).OrderBy(pageNumber => pageNumber).ToArray());
        PdfOutlineItem secondOutline = Assert.Single(selected.Outlines, outline => outline.Title == "Second outline");
        Assert.Equal(PdfOpenActionDestinationMode.FitRectangle, secondOutline.DestinationMode);
        Assert.Equal(10D, secondOutline.DestinationLeft);
        Assert.Equal(20D, secondOutline.DestinationBottom);
        Assert.Equal(90D, secondOutline.DestinationRight);
        Assert.Equal(144D, secondOutline.DestinationTop);
        Assert.Single(selected.PageLabels);
        Assert.Equal(0, selected.PageLabels[0].StartPageIndex);
        Assert.Equal("A-", selected.PageLabels[0].Prefix);
        Assert.Equal(10, selected.PageLabels[0].StartNumber);
        Assert.False(selected.HasReadableOpenAction);
        Assert.Null(selected.OpenAction);

        PdfDocumentInfo third = PdfInspector.InspectPageRanges(pdf, PdfPageRange.From(3, 3));

        PdfNamedDestination thirdDestination = Assert.Single(third.NamedDestinations);
        Assert.Equal("Third", thirdDestination.Name);
        Assert.Equal(3, thirdDestination.PageNumber);
        Assert.Equal("Third outline", Assert.Single(third.Outlines).Title);
        Assert.Equal(3, third.OpenAction!.PageNumber);
        PdfPageLabel thirdLabel = Assert.Single(third.PageLabels);
        Assert.Equal(2, thirdLabel.StartPageIndex);
        Assert.Equal("B-", thirdLabel.Prefix);
        Assert.Equal(3, thirdLabel.StartNumber);
    }

    [Fact]
    public void InspectPageRanges_FiltersAcroFormFieldsToSelectedSourcePages() {
        byte[] pdf = PdfDocument.Create(new PdfOptions {
                PageWidth = 320,
                PageHeight = 220,
                MarginLeft = 36,
                MarginRight = 36,
                MarginTop = 36,
                MarginBottom = 36
            })
            .TextField("First.Page", width: 120, height: 20, value: "one")
            .PageBreak()
            .TextField("Second.Page", width: 120, height: 20, value: "two")
            .PageBreak()
            .TextField("Third.Page", width: 120, height: 20, value: "three")
            .ToBytes();

        PdfDocumentInfo info = PdfInspector.InspectPageRanges(pdf, PdfPageRange.ParseMany("2,1,2"));

        Assert.Equal(2, info.FormFields.Count);
        Assert.Contains("First.Page", info.FormFieldNames);
        Assert.Contains("Second.Page", info.FormFieldNames);
        Assert.DoesNotContain("Third.Page", info.FormFieldNames);
        Assert.Equal(2, info.GetFormWidgets("Second.Page").Count);
        Assert.All(info.GetFormWidgets("Second.Page"), widget => Assert.Equal(2, widget.PageNumber));
        Assert.Equal(new[] { 1, 1, 1 }, info.Pages.Select(page => page.FormWidgets.Count).ToArray());
        Assert.Equal(new[] { "Second.Page", "First.Page", "Second.Page" }, info.Pages.Select(page => Assert.Single(page.FormWidgets).FieldName).ToArray());
        Assert.Equal(2, info.GetFormWidgets(2).Count);
        Assert.Empty(info.GetFormWidgets("Third.Page"));
        Assert.Equal(3, info.FormWidgetCount);
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
        Assert.Throws<PdfUnsupportedEncryptionException>(() => PdfInspector.Inspect(bytes));
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
        Assert.True(report.CanExtractText);
        Assert.True(report.CanExtractImages);
        Assert.True(report.CanReadLogicalObjects);
        Assert.True(report.CanManipulatePages);
        Assert.False(report.CanFillSimpleFormFields);
        Assert.False(report.CanFlattenSimpleFormFields);
        Assert.False(report.CanFillAndFlattenSimpleFormFields);
        Assert.True(report.Can(PdfPreflightCapability.ExtractText));
        Assert.True(report.Can(PdfPreflightCapability.ExtractImages));
        Assert.True(report.Can(PdfPreflightCapability.ReadLogicalObjects));
        Assert.True(report.Can(PdfPreflightCapability.ManipulatePages));
        Assert.False(report.Can(PdfPreflightCapability.FillSimpleFormFields));
        Assert.Empty(report.GetCapabilityDiagnostics(PdfPreflightCapability.ExtractText));
        Assert.Empty(report.GetCapabilityDiagnostics(PdfPreflightCapability.ExtractImages));
        Assert.Empty(report.GetCapabilityDiagnostics(PdfPreflightCapability.ReadLogicalObjects));
        Assert.Empty(report.GetCapabilityDiagnostics(PdfPreflightCapability.ManipulatePages));
        Assert.Contains(
            "PDF does not contain named text, choice, or button AcroForm fields supported for simple form filling by OfficeIMO.Pdf.",
            report.GetCapabilityDiagnostics(PdfPreflightCapability.FillSimpleFormFields));
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
    public void Validator_ReturnsReadableResultForGeneratedPdf() {
        PdfValidationResult result = PdfValidator.Validate(BuildTwoPagePdf());

        Assert.True(result.IsValid);
        Assert.True(result.CanRead);
        Assert.True(result.CanRewrite);
        Assert.True(result.CanExtractText);
        Assert.True(result.CanExtractImages);
        Assert.True(result.CanReadLogicalObjects);
        Assert.True(result.CanManipulatePages);
        Assert.True(result.Can(PdfPreflightCapability.ExtractText));
        Assert.Empty(result.GetCapabilityDiagnostics(PdfPreflightCapability.ExtractText));
        Assert.Equal("1.4", result.HeaderVersion);
        Assert.Equal(2, result.PageCount);
        Assert.Empty(result.Diagnostics);
        Assert.Empty(result.ReadBlockers);
        Assert.Empty(result.RewriteBlockers);
        Assert.NotNull(result.DocumentInfo);
        Assert.Same(result.Preflight.DocumentInfo, result.DocumentInfo);
    }

    [Fact]
    public void Validator_ReportsMalformedPdfWithoutThrowing() {
        byte[] bytes = System.Text.Encoding.ASCII.GetBytes("not a pdf");

        PdfValidationResult result = PdfValidator.Validate(bytes);

        Assert.False(result.IsValid);
        Assert.False(result.CanRead);
        Assert.False(result.CanRewrite);
        Assert.Equal(0, result.PageCount);
        Assert.Null(result.HeaderVersion);
        Assert.True(result.HasReadBlocker(PdfReadBlockerKind.MissingHeader));
        Assert.Contains("PDF header was not found.", result.Diagnostics);
        Assert.Null(result.DocumentInfo);
    }

    [Fact]
    public void Validator_ReadsFromPathAndCurrentStreamPosition() {
        byte[] bytes = BuildTwoPagePdf();
        byte[] prefix = System.Text.Encoding.ASCII.GetBytes("prefix");
        string path = Path.Combine(Path.GetTempPath(), "officeimo-pdf-validate-" + Guid.NewGuid().ToString("N") + ".pdf");

        try {
            File.WriteAllBytes(path, bytes);

            PdfValidationResult fromPath = PdfValidator.Validate(path);
            using var stream = new MemoryStream(prefix.Concat(bytes).ToArray());
            stream.Position = prefix.Length;
            PdfValidationResult fromStream = PdfValidator.Validate(stream);

            Assert.True(fromPath.IsValid);
            Assert.True(fromStream.IsValid);
            Assert.Equal(2, fromPath.PageCount);
            Assert.Equal(2, fromStream.PageCount);
            Assert.Equal(fromPath.DocumentInfo!.Pages[1].Width, fromStream.DocumentInfo!.Pages[1].Width);
        } finally {
            if (File.Exists(path)) {
                File.Delete(path);
            }
        }
    }


}
