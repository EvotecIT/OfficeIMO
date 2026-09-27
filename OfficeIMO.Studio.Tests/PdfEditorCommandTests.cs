using OfficeIMO.Pdf;
using OfficeIMO.Studio.Features.Editor;

namespace OfficeIMO.Studio.Tests;

public sealed class PdfEditorCommandTests {
    public static TheoryData<PdfEditorTool, string> AnnotationTools => new() {
        { PdfEditorTool.Note, "Text" },
        { PdfEditorTool.FreeText, "FreeText" },
        { PdfEditorTool.Highlight, "Highlight" },
        { PdfEditorTool.Underline, "Underline" },
        { PdfEditorTool.StrikeOut, "StrikeOut" },
        { PdfEditorTool.Rectangle, "Square" },
        { PdfEditorTool.Ellipse, "Circle" },
        { PdfEditorTool.Line, "Line" },
        { PdfEditorTool.Ink, "Ink" },
        { PdfEditorTool.Stamp, "Stamp" },
        { PdfEditorTool.SignatureAppearance, "FreeText" }
    };

    [Theory]
    [MemberData(nameof(AnnotationTools))]
    public void AnnotationTool_CreateSaveReopenInspectAndRender(PdfEditorTool tool, string subtype) {
        byte[] source = CreateSource();
        PdfEditorCommand command = PdfEditorCommandFactory.Create(source, tool, CreateGesture(), CreateProperties());

        byte[] edited = PdfEditorCommandExecutor.Apply(source, command);
        PdfDocument reopened = PdfDocument.Load(edited);
        PdfAnnotation annotation = Assert.Single(reopened.Inspect().GetAnnotationsBySubtype(subtype));
        PdfPageRenderResult render = Assert.Single(reopened.Render.Pages("1", new PdfPageRenderOptions {
            Format = PdfPageRenderFormat.Svg,
            ContinueOnError = false
        }));

        Assert.Equal(1, annotation.PageNumber);
        Assert.True(render.Succeeded, string.Join(Environment.NewLine, render.Diagnostics));
        Assert.NotEmpty(render.Bytes!);
    }

    [Fact]
    public void MarkupKeepsSeparateQuadsForSelectedLines() {
        byte[] source = CreateSource();
        var gesture = new PdfEditorGesture(1, 40D, 50D, 300D, 90D,
            Array.Empty<PdfEditorVisualPoint>(), TextQuads: [
                new PdfEditorVisualBounds(200D, 50D, 300D, 65D),
                new PdfEditorVisualBounds(40D, 75D, 120D, 90D)
            ]);
        PdfEditorCommand command = PdfEditorCommandFactory.Create(source, PdfEditorTool.Highlight, gesture, CreateProperties());

        byte[] edited = PdfEditorCommandExecutor.Apply(source, command);
        PdfAnnotation annotation = Assert.Single(PdfDocument.Load(edited).Inspect().GetAnnotationsBySubtype("Highlight"));

        Assert.Equal(16, annotation.QuadPoints.Count);
        Assert.Equal(100D, annotation.QuadPoints[2] - annotation.QuadPoints[0], 3);
        Assert.Equal(80D, annotation.QuadPoints[10] - annotation.QuadPoints[8], 3);
    }

    [Fact]
    public void SelectionGeometryGroupsAdjacentGlyphsWithoutFillingTheGapBetweenLines() {
        byte[] source = CreateSource(string.Join(" ", Enumerable.Repeat("selection geometry", 30)));
        IReadOnlyList<PdfPageInteractionRegion> regions = PdfDocument.Load(source).Render.Interactions(1).TextRegions;

        IReadOnlyList<PdfEditorVisualBounds> quads = OfficeIMO.Studio.Features.Reader.PdfPageCanvas.MergeSelectedTextRegions(regions);

        Assert.True(quads.Count >= 2);
        Assert.All(quads, quad => Assert.True(quad.Width > 0 && quad.Height > 0));
        Assert.True(quads[0].Bottom < quads[1].Top);
    }

    [Fact]
    public void LinkTool_CreatesUriActionReturnedByReader() {
        byte[] source = CreateSource();
        PdfEditorCommand command = PdfEditorCommandFactory.Create(source, PdfEditorTool.Link, CreateGesture(), CreateProperties());

        byte[] edited = PdfEditorCommandExecutor.Apply(source, command);

        PdfLinkAnnotation link = Assert.Single(PdfDocument.Load(edited).Inspect().GetLinkAnnotationsByUri("https://officeimo.com"));
        Assert.Equal(1, link.PageNumber);
    }

    [Fact]
    public void AddedTextAndImage_AreRealPageContentAfterReopen() {
        byte[] source = CreateSource();
        PdfEditorProperties textProperties = CreateProperties() with { Text = "Added overlay" };
        byte[] withText = PdfEditorCommandExecutor.Apply(
            source,
            PdfEditorCommandFactory.Create(source, PdfEditorTool.AddText, CreateGesture(), textProperties));
        PdfEditorProperties imageProperties = CreateProperties() with { ImageBytes = TinyPng };
        byte[] withImage = PdfEditorCommandExecutor.Apply(
            withText,
            PdfEditorCommandFactory.Create(withText, PdfEditorTool.AddImage, CreateGesture(), imageProperties));

        PdfDocument reopened = PdfDocument.Load(withImage);
        Assert.Contains("Added overlay", reopened.Read().Text, StringComparison.Ordinal);
        Assert.NotEmpty(reopened.Read().Images);
        Assert.True(Assert.Single(reopened.Render.Pages("1", new PdfPageRenderOptions { Format = PdfPageRenderFormat.Svg })).Succeeded);
    }

    [Fact]
    public void VerifiedRedaction_RemovesMatchedTextFromExtractedRawAndDecodedContent() {
        const string secret = "Secret account 123-45";
        byte[] source = CreateSource(secret);
        PdfPageInteractionMap interactions = PdfDocument.Load(source).Render.Interactions(1);
        PdfPageInteractionRegion[] regions = interactions.TextRegions
            .Where(region => !string.IsNullOrWhiteSpace(region.Text))
            .ToArray();
        var gesture = new PdfEditorGesture(
            1,
            regions.Min(region => region.Quad.Left),
            regions.Min(region => region.Quad.Top),
            regions.Max(region => region.Quad.Right),
            regions.Max(region => region.Quad.Bottom),
            Array.Empty<PdfEditorVisualPoint>());
        PdfEditorCommand command = PdfEditorCommandFactory.Create(source, PdfEditorTool.Redact, gesture, CreateProperties());

        PdfVerifiedRedactionResult result = PdfEditorCommandExecutor.ApplyVerifiedRedaction(source, command, secret);

        Assert.True(result.Plan.HasMatches);
        Assert.True(result.Evidence.IsVerified, result.Evidence.Summary);
        Assert.DoesNotContain(secret, PdfDocument.Load(result.Bytes).Read().Text, StringComparison.Ordinal);
        Assert.All(result.Evidence.Items, item => Assert.Equal(PdfRedactionEvidenceStatus.VerifiedAbsent, item.Status));
        Assert.True(result.Evidence.Verification.RawPdfBytesChecked);
        Assert.True(result.Evidence.Verification.DecodedPdfStreamsChecked);
        Assert.True(result.Evidence.Verification.ManagedRenderingChecked);
    }

    [Fact]
    public void AreaRedaction_AllowsIdenticalTextOutsideTheReviewedArea() {
        const string repeated = "Confidential";
        byte[] source = PdfDocument.Create(compose => {
            compose.Page(page => page.Size(600D, 800D).Content(content => content.Item(item => item.Paragraph(paragraph => paragraph.Text(repeated)))));
            compose.Page(page => page.Size(600D, 800D).Content(content => content.Item(item => item.Paragraph(paragraph => paragraph.Text(repeated)))));
        }).ToBytes();
        PdfPageInteractionRegion[] regions = PdfDocument.Load(source).Render.Interactions(1).TextRegions
            .Where(region => !string.IsNullOrWhiteSpace(region.Text))
            .ToArray();
        var gesture = new PdfEditorGesture(
            1,
            regions.Min(region => region.Quad.Left),
            regions.Min(region => region.Quad.Top),
            regions.Max(region => region.Quad.Right),
            regions.Max(region => region.Quad.Bottom),
            Array.Empty<PdfEditorVisualPoint>());

        PdfVerifiedRedactionResult result = PdfEditorCommandExecutor.ApplyVerifiedRedaction(
            source,
            PdfEditorCommandFactory.Create(source, PdfEditorTool.Redact, gesture, CreateProperties()));

        Assert.True(result.Evidence.IsVerified, result.Evidence.Summary);
        Assert.Contains(repeated, PdfDocument.Load(result.Bytes).Read().Text, StringComparison.Ordinal);
        Assert.Empty(result.Evidence.ResidualMatches);
        Assert.All(result.Evidence.Items, item => Assert.Equal(PdfRedactionEvidenceStatus.VerifiedAbsent, item.Status));
    }

    [Fact]
    public void VerifiedRedaction_ProvesImageOnlyAndAnnotationOnlyAreas() {
        byte[] blank = PdfDocument.Create(compose => compose.Page(page => page.Size(600D, 800D))).ToBytes();
        PdfEditorGesture gesture = CreateGesture();
        byte[] withImage = PdfEditorCommandExecutor.Apply(
            blank,
            PdfEditorCommandFactory.Create(blank, PdfEditorTool.AddImage, gesture, CreateProperties() with { ImageBytes = TinyPng }));
        PdfPageInteractionRegion imageRegion = Assert.Single(
            PdfDocument.Load(withImage).Render.Interactions(1).Regions,
            static region => region.Kind == PdfInteractionKind.Image);
        PdfEditorGesture imageGesture = Gesture(imageRegion);
        PdfVerifiedRedactionResult imageResult = PdfEditorCommandExecutor.ApplyVerifiedRedaction(
            withImage,
            PdfEditorCommandFactory.Create(withImage, PdfEditorTool.Redact, imageGesture, CreateProperties()));

        Assert.Contains(imageResult.Plan.Matches, match => match.Kind == PdfRedactionMatchKind.ImagePlacement);
        Assert.True(imageResult.Evidence.IsVerified, imageResult.Evidence.Summary);
        Assert.Empty(imageResult.Evidence.ResidualMatches);
        Assert.All(imageResult.Evidence.Items, item => Assert.Equal(PdfRedactionEvidenceStatus.VerifiedAbsent, item.Status));

        byte[] withAnnotation = PdfEditorCommandExecutor.Apply(
            blank,
            PdfEditorCommandFactory.Create(blank, PdfEditorTool.Rectangle, gesture, CreateProperties()));
        PdfPageInteractionRegion annotationRegion = Assert.Single(
            PdfDocument.Load(withAnnotation).Render.Interactions(1).Regions,
            static region => region.Kind == PdfInteractionKind.Annotation);
        PdfEditorGesture annotationGesture = Gesture(annotationRegion);
        PdfVerifiedRedactionResult annotationResult = PdfEditorCommandExecutor.ApplyVerifiedRedaction(
            withAnnotation,
            PdfEditorCommandFactory.Create(withAnnotation, PdfEditorTool.Redact, annotationGesture, CreateProperties()));

        Assert.Contains(annotationResult.Plan.Matches, match => match.Kind == PdfRedactionMatchKind.Annotation);
        Assert.True(annotationResult.Evidence.IsVerified, annotationResult.Evidence.Summary);
        Assert.Empty(annotationResult.Evidence.ResidualMatches);
        Assert.All(annotationResult.Evidence.Items, item => Assert.Equal(PdfRedactionEvidenceStatus.VerifiedAbsent, item.Status));
    }

    private static byte[] CreateSource(string text = "Existing content") =>
        PdfDocument.Create(compose => compose.Page(page => page
            .Size(600D, 800D)
            .Content(content => content.Item(item => item.Paragraph(paragraph => paragraph.Text(text)))))).ToBytes();

    private static PdfEditorGesture CreateGesture() => new(
        1,
        40D,
        50D,
        180D,
        100D,
        new[] { new PdfEditorVisualPoint(40D, 50D), new PdfEditorVisualPoint(180D, 100D) });

    private static PdfEditorGesture Gesture(PdfPageInteractionRegion region) => new(
        1,
        region.Quad.Left,
        region.Quad.Top,
        region.Quad.Right,
        region.Quad.Bottom,
        Array.Empty<PdfEditorVisualPoint>());

    private static PdfEditorProperties CreateProperties() => new(
        "Review annotation",
        "OfficeIMO Studio",
        PdfColor.FromRgb(229, 72, 77),
        "Approved",
        "https://officeimo.com",
        14D);

    private static readonly byte[] TinyPng = Convert.FromBase64String(
        "iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAQAAAC1HAwCAAAAC0lEQVR42mNk+A8AAQUBAScY42YAAAAASUVORK5CYII=");
}
