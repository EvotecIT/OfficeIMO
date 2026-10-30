using OfficeIMO.Pdf;
using OfficeIMO.Studio.Native;
using OfficeIMO.Studio.Features.Workspace;

namespace OfficeIMO.Studio.Core.Tests;

public sealed class StudioNativePermissionFactsTests : IDisposable {
    private readonly string _root = Path.Combine(Path.GetTempPath(), "officeimo-native-permissions-" + Guid.NewGuid().ToString("N"));
    public StudioNativePermissionFactsTests() {
        Directory.CreateDirectory(_root);
        StudioNativeEngine.Initialize(Path.Combine(_root, "data"));
    }
    public void Dispose() => Directory.Delete(_root, true);

    private string Source(PdfStandardPermissions permissions) {
        string path = Path.Combine(_root, "restricted.pdf");
        File.WriteAllBytes(path, PdfDocument.Create(compose => compose.Page(page => page.Size(300, 400)))
            .Security.Encrypt(new PdfStandardEncryptionOptions("reader") { OwnerPassword = "owner", AllowedPermissions = permissions }).Pdf);
        return path;
    }

    [Theory]
    [InlineData(false, false)]
    [InlineData(true, false)]
    [InlineData(true, true)]
    public async Task PrintingFactsDescribeTheAuthenticatedResolutionLimit(bool highQuality, bool owner) {
        var permissions = PdfStandardPermissions.Print | (highQuality ? PdfStandardPermissions.HighQualityPrint : PdfStandardPermissions.None);
        using var document = await StudioNativeDocument.OpenAsync(Source(permissions), owner ? "owner" : "reader", default);
        var facts = await document.GetInfoAsync(default);
        Assert.True(facts.CanPrint);
        Assert.Equal(owner || highQuality ? null : (double?)150, facts.MaxPrintDpi);
        Assert.Equal(!owner, facts.RequiresOwnerPasswordForProtection);
    }

    [Fact]
    public async Task AssemblyOnlyDocumentCanRotateAndInsertBlankWithoutCopyPermission() {
        using var document = await StudioNativeDocument.OpenAsync(Source(PdfStandardPermissions.AssembleDocument), "reader", default);
        var facts = await document.GetInfoAsync(default);
        Assert.True(facts.CanRotatePages);
        Assert.True(facts.CanInsertBlankPages);
        Assert.False(facts.CanExtractPages);
        await document.InsertBlankPageAsync(1, 300, 400, default);
        Assert.Equal(2, (await document.GetInfoAsync(default)).Pages.Count);
        await document.RotatePagesAsync([1, 2], 90, default);
        Assert.All((await document.GetInfoAsync(default)).Pages, page => Assert.Equal(90, page.Rotation));
    }

    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public async Task AccessibilityTextAndDisplayDoNotGrantContentExportOrPrinting(bool owner) {
        byte[] png = Convert.FromBase64String("iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAIAAACQd1PeAAAADElEQVR4nGP4z8AAAAMBAQDJ/pLvAAAAAElFTkSuQmCC");
        string path = Path.Combine(_root, "accessible-image.pdf");
        File.WriteAllBytes(path, PdfDocument.Create(document => document.Content(content => {
                content.Paragraph(paragraph => paragraph.Text("Accessible text"));
                content.Image(png, global::OfficeIMO.OfficeLength.Points(24), global::OfficeIMO.OfficeLength.Points(24));
            }))
            .Security.Encrypt(new PdfStandardEncryptionOptions("reader") {
                OwnerPassword = "owner", AllowedPermissions = PdfStandardPermissions.Accessibility
            }).Pdf);
        byte[] original = File.ReadAllBytes(path);
        using var document = await StudioNativeDocument.OpenAsync(path, owner ? "owner" : "reader", default);
        var facts = await document.GetInfoAsync(default);
        Assert.True(facts.CanExtractText);
        Assert.Equal(owner, facts.CanExtractContent);
        Assert.Equal(owner, facts.CanPrint);
        var display = await document.RenderPageAsync(1, 2, default);
        Assert.Equal(new byte[] { 137, 80, 78, 71, 13, 10, 26, 10 }, display.Bytes[..8]);
        string text = Path.Combine(_root, "accessible.txt");
        await document.WriteAsync(new(Export: new("text", text)), default);
        Assert.Contains("Accessible text", File.ReadAllText(text));
        string images = Path.Combine(_root, "images");
        string html = Path.Combine(_root, "converted.html");
        if (owner) {
            var result = await document.WriteAsync(new(Export: new("images", images)), default);
            Assert.True(new FileInfo(Assert.Single(result.Files)).Length > 0);
            await document.WriteAsync(new(Export: new("html", html)), default);
            Assert.Contains("Accessible text", File.ReadAllText(html));
        } else {
            foreach (string format in new[] { "images", "html", "word", "excel", "powerpoint", "markdown", "json" }) {
                string blocked = Path.Combine(_root, "blocked-" + format);
                await Assert.ThrowsAsync<NotSupportedException>(() => document.WriteAsync(new(Export: new(format, blocked)), default));
                Assert.False(File.Exists(blocked));
                Assert.False(Directory.Exists(blocked));
            }
        }
        Assert.Equal(original, File.ReadAllBytes(path));
    }

    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public async Task EncryptedAnnotationCreationAndBoundsRetainAuthentication(bool owner) {
        using var document = await StudioNativeDocument.OpenAsync(Source(PdfStandardPermissions.ModifyAnnotations), owner ? "owner" : "reader", default);
        await document.AddAnnotationsAsync([
            new("Note", 1, new(20, 20, 30, 30), Text: "Note"),
            new("Rectangle", 1, new(60, 60, 40, 30))
        ], default);
        var square = Assert.Single((await document.GetAnnotationsAsync(1, default)).Annotations, annotation => annotation.Subtype == "Square");
        var moved = square.Bounds with { X = square.Bounds.X + 5 };
        await document.UpdateAnnotationAsync(new(square.ObjectNumber, Bounds: moved), default);
        Assert.Equal(moved, Assert.Single((await document.GetAnnotationsAsync(1, default)).Annotations, annotation => annotation.Subtype == "Square").Bounds);
        await document.UndoAsync(default);
        Assert.Equal(square.Bounds, Assert.Single((await document.GetAnnotationsAsync(1, default)).Annotations, annotation => annotation.Subtype == "Square").Bounds);
        await document.UndoAsync(default);
        Assert.Empty((await document.GetAnnotationsAsync(1, default)).Annotations);
        Assert.True((await document.GetInfoAsync(default)).IsEncrypted);
    }

    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public async Task RestrictedStructureAndFillOnlyFieldsDoNotExtractPageContents(bool fillAllowed) {
        string path = Path.Combine(_root, "restricted-form.pdf");
        var source = PdfDocument.Create(document => document.Content(content => content.Canvas(canvas =>
            canvas.TextField("Name", null, global::OfficeIMO.OfficeLength.Points(72), global::OfficeIMO.OfficeLength.Points(72), global::OfficeIMO.OfficeLength.Points(200), global::OfficeIMO.OfficeLength.Points(24)))));
        File.WriteAllBytes(path, source.Security.Encrypt(new PdfStandardEncryptionOptions("reader") {
            OwnerPassword = "owner", AllowedPermissions = fillAllowed ? PdfStandardPermissions.FillForms : PdfStandardPermissions.None
        }).Pdf);
        using var document = await StudioNativeDocument.OpenAsync(path, "reader", default);
        Assert.True((await document.GetStructureAsync(default)).HasForms);
        Assert.False((await document.GetInfoAsync(default)).CanExtractContent);
        Assert.Throws<PdfPermissionDeniedException>(() => PdfDocument.Load(path, new PdfLoadOptions { Password = "reader" }).Inspect());
        if (!fillAllowed) {
            await Assert.ThrowsAsync<PdfPermissionDeniedException>(() => document.GetFormFieldsAsync(1, default));
            return;
        }
        var field = Assert.Single((await document.GetFormFieldsAsync(1, default)).Fields);
        Assert.Equal("Name", field.Name);
        Assert.Equal(200, Assert.Single(field.Widgets).Bounds.Width);
        // Metadata authorization does not grant a write path the PDF engine has not qualified.
        Assert.False((await document.GetInfoAsync(default)).CanFillForms);
        await Assert.ThrowsAsync<NotSupportedException>(() => document.FillFormAsync(new(
            new Dictionary<string, IReadOnlyList<string>> { ["Name"] = ["Ada"] }), default));
        using var authorized = await StudioNativeDocument.OpenAsync(path, "owner", default);
        await authorized.FillFormAsync(new(new Dictionary<string, IReadOnlyList<string>> { ["Name"] = ["Ada"] }), default);
        Assert.Equal(["Ada"], Assert.Single((await authorized.GetFormFieldsAsync(1, default)).Fields).Values);
    }

    [Fact]
    public async Task Form_authoring_metadata_uses_its_own_authorization_without_copy_permission() {
        string path = Path.Combine(_root, "author-form.pdf");
        byte[] original = PdfDocument.Create(document => document.Content(content => content.Canvas(canvas =>
            canvas.TextField("Name", "Ada", global::OfficeIMO.OfficeLength.Points(72), global::OfficeIMO.OfficeLength.Points(72), global::OfficeIMO.OfficeLength.Points(200), global::OfficeIMO.OfficeLength.Points(24))))).Security.Encrypt(new("reader") {
                OwnerPassword = "owner", AllowedPermissions = PdfStandardPermissions.ModifyContents | PdfStandardPermissions.ModifyAnnotations
            }).Pdf;
        File.WriteAllBytes(path, original);
        using var workspace = await PdfWorkspace.OpenAsync(path, default, StudioNativeEngine.CreateRecoveryStore(), "reader");
        Assert.True(workspace.CanAuthorForms);
        Assert.Null(workspace.DocumentInfo);
        await workspace.UpdateFormDefinitionAsync("Name", "Recipient", true, false, default);
        var field = Assert.Single(workspace.FormMetadata);
        Assert.Equal("Recipient", field.Name);
        Assert.True(field.IsRequired);
        Assert.Equal("Ada", field.Value);
        await workspace.UndoAsync(default);
        Assert.Equal("Name", Assert.Single(workspace.FormMetadata).Name);
        Assert.Equal(original, File.ReadAllBytes(path));
    }

    [Fact]
    public async Task Authorized_sanitized_copy_does_not_require_copy_permission_for_internal_preservation_proof() {
        string path = Source(PdfStandardPermissions.ModifyContents);
        byte[] original = File.ReadAllBytes(path);
        using var document = await StudioNativeDocument.OpenAsync(path, "reader", default);
        string output = Path.Combine(_root, "sanitized.pdf");
        await document.WriteAsync(new(Sanitize: new(output)), default);
        Assert.Equal(original, File.ReadAllBytes(path));
        using var reopened = await StudioNativeDocument.OpenAsync(output, "reader", default);
        var info = await reopened.GetInfoAsync(default);
        Assert.True(info.IsEncrypted);
        Assert.False(info.CanExtractContent);
        Assert.Single(info.Pages);
    }
}
