using OfficeIMO.Pdf;
using Xunit;

namespace OfficeIMO.Tests.Pdf;

public sealed class PdfDocumentViewingTests {
    [Fact]
    public void DefaultLoadOptionsRemainSourceCompatibleWithViewingInspection() {
        var document = PdfDocument.Load(PdfDocument.Create().Paragraph(p => p.AddText("Viewing contract")).ToBytes());
        Assert.NotNull(document.InspectForViewing(default).LogicalContent);
        Assert.NotNull(document.InspectForViewing(default, default).LogicalContent);
        Func<PdfLoadOptions?, CancellationToken, PdfDocumentViewInfo> inspect = document.InspectForViewing;
        Assert.NotNull(inspect(default, default).LogicalContent);
        Assert.Null(document.InspectGeometryForViewing(default, default).LogicalContent);
    }

    [Theory]
    [InlineData(PdfStandardPermissions.None)]
    [InlineData(PdfStandardPermissions.Accessibility)]
    [InlineData(PdfStandardPermissions.Print)]
    public void RestrictedViewingPreservesRasterAndDoesNotGrantExtraction(PdfStandardPermissions permissions) {
        byte[] bytes = CreatePdf(permissions);
        var restricted = PdfDocument.Load(bytes, new PdfLoadOptions { Password = "reader" });
        PdfDocumentViewInfo view = restricted.InspectForViewing();
        Assert.Equal(PdfPasswordAuthenticationRole.User, view.Security.PasswordAuthenticationRole);
        Assert.False(view.CanExtractContent); Assert.Null(view.LogicalContent);
        Assert.Equal(permissions.HasFlag(PdfStandardPermissions.Accessibility), view.CanExtractText);
        Assert.Single(view.Pages); Assert.Empty(view.Pages[0].Annotations); Assert.Empty(view.Pages[0].FormWidgets);
        Assert.Empty(view.Pages[0].LinkAnnotations); Assert.Empty(view.Pages[0].PageActions);
        var owner = PdfDocument.Load(bytes, new PdfLoadOptions { Password = "owner" });
        PdfDocumentViewInfo authorized = owner.InspectForViewing();
        Assert.True(authorized.CanExtractContent); Assert.NotNull(authorized.LogicalContent);
        PdfDocumentViewInfo geometry = owner.InspectGeometryForViewing();
        Assert.True(geometry.CanExtractContent); Assert.Null(geometry.LogicalContent);
        Assert.Equal(authorized.PageCount, geometry.PageCount);
        Assert.Equal(authorized.Pages[0].Width, geometry.Pages[0].Width);
        Assert.False(restricted.InspectGeometryForViewing().CanExtractContent);
        Assert.Equal("Private metadata", authorized.LogicalContent.Metadata.Title);
        Assert.Equal(authorized.Pages[0].Width, view.Pages[0].Width);
        PdfPageRenderResult raster = restricted.Render.DisplayPage(1);
        Assert.True(raster.Succeeded); Assert.Equal(global::OfficeIMO.Drawing.OfficeImageExportFormat.Png, raster.Format);
        Assert.Equal(owner.Render.DisplayPage(1).Bytes, raster.Bytes);
        string? capture = Environment.GetEnvironmentVariable("OFFICEIMO_PDF_VIEWING_CAPTURE_DIR");
        if (!string.IsNullOrEmpty(capture)) {
            Directory.CreateDirectory(capture); File.WriteAllBytes(Path.Combine(capture, "view-" + permissions + ".png"), raster.Bytes!);
        }
        Assert.Equal(owner.Render.Pages("1", new PdfPageRenderOptions { ContinueOnError = false })[0].Bytes, raster.Bytes);
        Assert.Throws<PdfPermissionDeniedException>(() => restricted.Inspect());
        Assert.Throws<PdfPermissionDeniedException>(() => restricted.Read());
        Assert.Throws<PdfPermissionDeniedException>(() => restricted.Render.Drawing(1));
        // Raster rendering is viewing; SVG output carries text and source images, so it stays an extraction.
        Assert.Equal(raster.Bytes, restricted.Render.Pages("1", new PdfPageRenderOptions { ContinueOnError = false })[0].Bytes);
        Assert.Throws<PdfPermissionDeniedException>(() => restricted.Render.Pages("1", new PdfPageRenderOptions {
            ContinueOnError = false, Format = global::OfficeIMO.Drawing.OfficeImageExportFormat.Svg }));
        // A provider-loaded font program is caller code that would receive the page text.
        var callerFonts = new global::OfficeIMO.Drawing.OfficeFontFaceCollection { FontProgramProvider = new CallerFontProgramProvider() };
        callerFonts.Add("Helvetica", File.ReadAllBytes(Path.Combine(RepositoryTestPaths.Find(), "OfficeIMO.Drawing.Tests", "TestAssets", "OpenSans-Bold.ttf")));
        Assert.Throws<PdfPermissionDeniedException>(() => restricted.Render.Pages("1", new PdfPageRenderOptions {
            ContinueOnError = false, Fonts = callerFonts }));
        Assert.False(restricted.PlanMutation(PdfMutationOperation.ExtractPages).CanExecute);
        Assert.False(restricted.PlanMutation(PdfMutationOperation.ChangeEncryption).CanExecute);
        if (!permissions.HasFlag(PdfStandardPermissions.Print))
            Assert.Throws<PdfPermissionDeniedException>(() => restricted.GetPrintablePageLayouts());
        // A display call must not relax cached read options for any subsequent operation.
        Assert.Throws<PdfPermissionDeniedException>(() => restricted.Render.Interactions(1));
    }

    [Theory]
    [InlineData("display")]
    [InlineData("ranges")]
    [InlineData("selection")]
    [InlineData("images")]
    public void CancellationDuringGeneratedSerializationStopsFurtherComposition(string route) {
        using var cancellation = new CancellationTokenSource();
        int callbacks = 0;
        var options = new PdfCreateOptions {
            TextLineBreakCallback = text => {
                callbacks++; cancellation.Cancel(); return new[] { text.Length / 2 };
            }
        };
        var document = PdfDocument.Create(options);
        for (int index = 0; index < 30; index++) document.Paragraph(paragraph => paragraph.AddText(new string('W', 600)));
        Assert.ThrowsAny<OperationCanceledException>(() => {
            switch (route) {
                case "display": document.Render.DisplayPage(1, cancellationToken: cancellation.Token); break;
                case "ranges": document.Render.Pages("1", cancellationToken: cancellation.Token); break;
                case "selection": document.Render.Pages(PdfPageSelection.Parse("1"), cancellationToken: cancellation.Token); break;
                case "images": document.ToImages(cancellationToken: cancellation.Token); break;
            }
        });
        Assert.InRange(callbacks, 1, 2);
    }

    [Fact]
    public void RestrictedViewOmitsExistingLogicalMetadataLinksAndAttachments() {
        byte[] source = PdfRewritePreservationTestSupport.BuildPreservationProofPdf();
        byte[] encrypted = PdfDocument.Load(source).Security.Encrypt(new PdfStandardEncryptionOptions("reader") {
            OwnerPassword = "owner", AllowedPermissions = PdfStandardPermissions.None
        }).Pdf;
        var owner = PdfDocument.Load(encrypted, new PdfLoadOptions { Password = "owner" }).InspectForViewing();
        Assert.NotNull(owner.LogicalContent); Assert.NotEmpty(owner.LogicalContent.Attachments);
        Assert.NotEmpty(owner.LogicalContent.NamedDestinations); Assert.NotEmpty(owner.Pages[0].LinkAnnotations);
        var user = PdfDocument.Load(encrypted, new PdfLoadOptions { Password = "reader" }).InspectForViewing();
        Assert.Null(user.LogicalContent); Assert.Equal(owner.PageCount, user.PageCount);
        Assert.All(user.Pages, page => { Assert.Empty(page.LinkAnnotations); Assert.Empty(page.Annotations); Assert.Empty(page.FormWidgets); });
    }

    [Fact]
    public void ViewingStillRequiresAuthenticationAndHonorsCancellationAndBudgets() {
        byte[] bytes = CreatePdf(PdfStandardPermissions.None);
        Assert.Throws<PdfPasswordRequiredException>(() => PdfDocument.Load(bytes).InspectForViewing());
        Assert.Throws<PdfInvalidPasswordException>(() => PdfDocument.Load(bytes, new PdfLoadOptions { Password = "wrong" }).Render.DisplayPage(1));
        var document = PdfDocument.Load(bytes, new PdfLoadOptions { Password = "reader" });
        using var cancellation = new CancellationTokenSource(); cancellation.Cancel();
        Assert.ThrowsAny<OperationCanceledException>(() => document.InspectForViewing(cancellationToken: cancellation.Token));
        Assert.ThrowsAny<OperationCanceledException>(() => document.Render.DisplayPage(1, cancellationToken: cancellation.Token));
        Assert.Throws<PdfReadLimitException>(() => document.Render.DisplayPage(1, new PdfPageDisplayOptions { MaximumPixels = 4 }));
        Assert.Throws<PdfReadLimitException>(() => document.Render.DisplayPage(1, new PdfPageDisplayOptions { MaximumOutputBytes = 8 }));
        Assert.Throws<ArgumentOutOfRangeException>(() => document.Render.DisplayPage(0));
        Assert.Throws<ArgumentOutOfRangeException>(() => document.Render.DisplayPage(2));
        var thumbnail = document.Render.DisplayPage(1, new PdfPageDisplayOptions { MaximumDimension = 80 });
        Assert.InRange(Math.Max(thumbnail.Width, thumbnail.Height), 1, 80);
    }

    private sealed class CallerFontProgramProvider : global::OfficeIMO.Drawing.IOfficeFontProgramProvider {
        public global::OfficeIMO.Drawing.OfficeFontProgramLoadResult? TryLoad(global::OfficeIMO.Drawing.OfficeFontProgramLoadRequest request) =>
            new(new CallerFontProgram(), request.Data.Length);
    }

    private sealed class CallerFontProgram : global::OfficeIMO.Drawing.IOfficeFontProgram {
        private static Exception Reached() => new InvalidOperationException("Page content reached a caller font program.");
        public string Fingerprint => "caller-font-program";
        public string? DisplayName => null;
        public int? CollectionIndex => null;
        public int UnitsPerEm => 1000;
        public bool IsOpenTypeCff => false;
        public bool ProvidesComplexTextLayout => false;
        public double LineSpacingRatio => 1D;
        public byte[] GetFontDataForShaping() => throw Reached();
        public bool HasGlyphs(string text) => throw Reached();
        public double Measure(string text, double fontSize) => throw Reached();
        public IReadOnlyList<double> MeasureTextElements(IReadOnlyList<string> elements, double fontSize) => throw Reached();
        public double LineHeight(double fontSize) => throw Reached();
        public List<List<global::OfficeIMO.Drawing.OfficePoint>> GetTextContours(string text, double x, double y, double fontSize) => throw Reached();
        public bool TryGetGlyphMetrics(int scalar, out int glyphId, out int advanceWidth) => throw Reached();
        public double MeasureShapedText(string text, global::OfficeIMO.Drawing.OfficeTextShapingResult result, double fontSize) => throw Reached();
        public List<List<global::OfficeIMO.Drawing.OfficePoint>> GetShapedTextContours(string text, global::OfficeIMO.Drawing.OfficeTextShapingResult result,
            double x, double y, double fontSize) => throw Reached();
    }

    private static byte[] CreatePdf(PdfStandardPermissions permissions) =>
        PdfDocument.Create(new PdfCreateOptions().SetEncryption(new PdfStandardEncryptionOptions("reader") {
            OwnerPassword = "owner", AllowedPermissions = permissions
        })).Meta(title: "Private metadata")
          .Paragraph(paragraph => paragraph.AddText("Visible content with restricted extraction"))
          .Canvas(canvas => canvas.Image(PdfPngTestImages.CreateRgbPng(30, 90, 180), global::OfficeIMO.OfficeLength.Points(20D), global::OfficeIMO.OfficeLength.Points(20D), global::OfficeIMO.OfficeLength.Points(40D), global::OfficeIMO.OfficeLength.Points(40D))).ToBytes();
}
