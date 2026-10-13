using System.Runtime.InteropServices;
using System.Security.Cryptography;
using System.Security.Cryptography.X509Certificates;
using System.Text;
using System.Text.Json;
using OfficeIMO.Studio.Features.Conversion;
using OfficeIMO.Studio.Features.Reader;
using OfficeIMO.Pdf;
using OfficeIMO.Security;
using OfficeIMO.Studio.Features.Scan;
using OfficeIMO.Studio.Features.Workspace;
using OfficeIMO.Studio.Native;

namespace OfficeIMO.Studio.Core.Tests;

/// <summary>The engine library native Apple apps call: the same document journey as the desktop, through its C interface.</summary>
public sealed class StudioNativeEngineTests : IDisposable {
    private readonly string _root = Path.Combine(Path.GetTempPath(), "officeimo-native-engine-" + Guid.NewGuid().ToString("N"));

    public StudioNativeEngineTests() {
        Directory.CreateDirectory(_root);
        StudioNativeEngine.Initialize(Path.Combine(_root, "data"));
    }

    public void Dispose() => Directory.Delete(_root, recursive: true);

    private string CreatePdf(int pages = 2) {
        string path = Path.Combine(_root, "source.pdf");
        PdfDocument.Create(document => {
            for (int page = 0; page < pages; page++) document.Page(content => content.Size(300, 400));
        }).Save(path);
        return path;
    }

    [Fact]
    public async Task A_document_opens_renders_rotates_undoes_and_saves_a_copy() {
        string source = CreatePdf();
        byte[] original = File.ReadAllBytes(source);
        using StudioNativeDocument document = await StudioNativeDocument.OpenAsync(source, null, default);

        StudioNativeDocumentInfo info = await document.GetInfoAsync(default);
        Assert.Equal(new[] { 1, 2 }, info.Pages.Select(page => page.Number));
        Assert.False(info.IsDirty);

        var rendered = await document.RenderPageAsync(1, 1.5, default);
        Assert.Equal(new byte[] { 0x89, 0x50, 0x4E, 0x47 }, rendered.Bytes[..4]);
        Assert.Equal(450, rendered.PixelWidth);

        await document.RotatePagesAsync([1], 90, default);
        info = await document.GetInfoAsync(default);
        Assert.Equal(90, info.Pages[0].Rotation);
        Assert.True(info.IsDirty && info.CanUndo);
        await document.UndoAsync(default);
        Assert.Equal(0, (await document.GetInfoAsync(default)).Pages[0].Rotation);
        await document.RedoAsync(default);
        // Rendering follows the edited revision, with the page now landscape.
        Assert.Equal(600, (await document.RenderPageAsync(1, 1.5, default)).PixelWidth);

        string copy = Path.Combine(_root, "rotated.pdf");
        await document.SaveAsync(copy, default);
        info = await document.GetInfoAsync(default);
        Assert.Equal(copy, info.Path);
        Assert.False(info.IsDirty);
        Assert.Equal(original, File.ReadAllBytes(source));
        using StudioNativeDocument reopened = await StudioNativeDocument.OpenAsync(copy, null, default);
        Assert.Equal(90, (await reopened.GetInfoAsync(default)).Pages[0].Rotation);
    }

    // Pages of distinct widths, so a page's width identifies it after an edit.
    private string CreatePagesOfWidths(params double[] widths) => CreatePagesOfWidthsAt("widths.pdf", widths);

    private string CreatePagesOfWidthsAt(string name, params double[] widths) {
        string path = Path.Combine(_root, name);
        PdfDocument.Create(document => {
            foreach (double width in widths) document.Page(content => content.Size(width, 400));
        }).Save(path);
        return path;
    }

    private static async Task<double[]> WidthsAsync(StudioNativeDocument document) =>
        (await document.GetInfoAsync(default)).Pages.Select(page => page.Width).ToArray();

    [Fact]
    public async Task Pages_are_deleted_duplicated_reordered_and_inserted_as_undoable_changes() {
        using StudioNativeDocument document = await StudioNativeDocument.OpenAsync(CreatePagesOfWidths(300, 310, 320), null, default);
        Assert.True((await document.GetInfoAsync(default)).CanMutatePages);

        await document.DeletePagesAsync([2], default);
        Assert.Equal(new[] { 300D, 320D }, await WidthsAsync(document));
        await document.UndoAsync(default);
        Assert.Equal(new[] { 300D, 310D, 320D }, await WidthsAsync(document));

        await document.DuplicatePagesAsync([1], default);
        Assert.Equal(new[] { 300D, 300D, 310D, 320D }, await WidthsAsync(document));

        await document.ReorderPagesAsync([4, 3, 1, 2], default);
        Assert.Equal(new[] { 320D, 310D, 300D, 300D }, await WidthsAsync(document));

        await document.InsertBlankPageAsync(5, 500, 600, default);
        StudioNativeDocumentInfo info = await document.GetInfoAsync(default);
        Assert.Equal((500D, 600D), (info.Pages[^1].Width, info.Pages[^1].Height));

        // Each edit undoes on its own, back to the original pages.
        for (int edit = 0; edit < 3; edit++) await document.UndoAsync(default);
        Assert.Equal(new[] { 300D, 310D, 320D }, await WidthsAsync(document));
    }

    [Fact]
    public async Task Page_edits_refuse_what_would_damage_the_document() {
        using StudioNativeDocument document = await StudioNativeDocument.OpenAsync(CreatePagesOfWidths(300, 310), null, default);

        await Assert.ThrowsAsync<ArgumentException>(() => document.DeletePagesAsync([1, 2], default));
        await Assert.ThrowsAsync<ArgumentException>(() => document.ReorderPagesAsync([1, 1], default));
        await Assert.ThrowsAsync<ArgumentException>(() => document.ReorderPagesAsync([2], default));
        await Assert.ThrowsAsync<ArgumentOutOfRangeException>(() => document.DuplicatePagesAsync([3], default));
        await Assert.ThrowsAsync<ArgumentOutOfRangeException>(() => document.InsertBlankPageAsync(4, 300, 400, default));
        await Assert.ThrowsAsync<ArgumentOutOfRangeException>(() => document.InsertBlankPageAsync(1, 0, 400, default));

        Assert.False((await document.GetInfoAsync(default)).IsDirty);
    }

    [Theory]
    [InlineData("open", false)]
    [InlineData("owner", true)]
    public async Task Printing_follows_the_document_permissions(string password, bool allowed) {
        string path = Path.Combine(_root, "no-print.pdf");
        var encryption = new PdfStandardEncryptionOptions("open") { OwnerPassword = "owner", AllowedPermissions = PdfStandardPermissions.CopyContents };
        File.WriteAllBytes(path, PdfDocument.Load(File.ReadAllBytes(CreatePagesOfWidths(300))).Security.Encrypt(encryption).Pdf);

        using StudioNativeDocument document = await StudioNativeDocument.OpenAsync(path, password, default);
        Assert.Equal(allowed, (await document.GetInfoAsync(default)).CanPrint);
        using StudioNativeDocument open = await StudioNativeDocument.OpenAsync(CreatePagesOfWidths(300), null, default);
        Assert.True((await open.GetInfoAsync(default)).CanPrint);
    }

    [Fact]
    public async Task Pages_from_other_PDFs_are_inserted_as_one_undoable_change() {
        using StudioNativeDocument document = await StudioNativeDocument.OpenAsync(CreatePagesOfWidths(300, 310), null, default);
        string first = CreatePagesOfWidthsAt("first.pdf", 400, 410, 420);
        string second = CreatePagesOfWidthsAt("second.pdf", 500);
        Assert.True((await document.GetInfoAsync(default)).CanImportPages);

        StudioNativeImportResult result = await document.ImportPagesAsync(new StudioNativeImportRequest(
            [new StudioNativeImportSource(first, Pages: [3, 1]), new StudioNativeImportSource(second)], InsertBefore: 2), default);

        Assert.Equal(3, result.Imported);
        Assert.Equal(new[] { 300D, 420D, 400D, 500D, 310D }, await WidthsAsync(document));
        await document.UndoAsync(default);
        Assert.Equal(new[] { 300D, 310D }, await WidthsAsync(document));
    }

    [Fact]
    public async Task Importing_reports_which_source_needs_its_password() {
        using StudioNativeDocument document = await StudioNativeDocument.OpenAsync(CreatePagesOfWidths(300), null, default);
        string plain = CreatePagesOfWidthsAt("plain.pdf", 400);
        string locked = Path.Combine(_root, "locked.pdf");
        File.WriteAllBytes(locked, PdfDocument.Load(File.ReadAllBytes(CreatePagesOfWidthsAt("unlocked.pdf", 500)))
            .Security.Encrypt(new PdfStandardEncryptionOptions("open") { OwnerPassword = "owner" }).Pdf);

        var missing = await Assert.ThrowsAsync<PdfImportPasswordException>(() => document.ImportPagesAsync(new StudioNativeImportRequest(
            [new StudioNativeImportSource(plain), new StudioNativeImportSource(locked)], 2), default));
        Assert.Equal((1, "locked.pdf", false), (missing.SourceIndex, missing.Name, missing.Invalid));
        Assert.Equal(StudioNativeAbi.PasswordRequired, StudioNativeAbi.StatusOf(missing));
        var wrong = await Assert.ThrowsAsync<PdfImportPasswordException>(() => document.ImportPagesAsync(new StudioNativeImportRequest(
            [new StudioNativeImportSource(locked, "nope")], 2), default));
        Assert.Equal(StudioNativeAbi.InvalidPassword, StudioNativeAbi.StatusOf(wrong));
        Assert.Equal(new[] { 300D }, await WidthsAsync(document));

        await document.ImportPagesAsync(new StudioNativeImportRequest([new StudioNativeImportSource(plain), new StudioNativeImportSource(locked, "open")], 2), default);
        Assert.Equal(new[] { 300D, 400D, 500D }, await WidthsAsync(document));
    }

    [Fact]
    public async Task Importing_refuses_requests_outside_the_document() {
        using StudioNativeDocument document = await StudioNativeDocument.OpenAsync(CreatePagesOfWidths(300), null, default);
        string source = CreatePagesOfWidthsAt("source.pdf", 400, 410);

        await Assert.ThrowsAsync<ArgumentException>(() => document.ImportPagesAsync(new StudioNativeImportRequest([], 1), default));
        await Assert.ThrowsAnyAsync<ArgumentException>(() => document.ImportPagesAsync(new StudioNativeImportRequest([new StudioNativeImportSource(source, Pages: [3])], 1), default));
        await Assert.ThrowsAnyAsync<ArgumentException>(() => document.ImportPagesAsync(new StudioNativeImportRequest([new StudioNativeImportSource(source)], 3), default));
        await Assert.ThrowsAnyAsync<Exception>(() => document.ImportPagesAsync(new StudioNativeImportRequest([new StudioNativeImportSource(Path.Combine(_root, "missing.pdf"))], 1), default));
        Assert.False((await document.GetInfoAsync(default)).IsDirty);
    }

    [Fact]
    public unsafe void The_C_interface_imports_pages_and_names_a_source_that_needs_its_password() {
        string target = CreatePagesOfWidths(300);
        string plain = CreatePagesOfWidthsAt("plain.pdf", 400);
        string locked = Path.Combine(_root, "locked.pdf");
        File.WriteAllBytes(locked, PdfDocument.Load(File.ReadAllBytes(plain)).Security.Encrypt(new PdfStandardEncryptionOptions("open")).Pdf);
        delegate* unmanaged<byte*, byte*, nint, nint*, byte**, int> open = &StudioNativeExports.DocumentOpen;
        delegate* unmanaged<nint, byte*, nint, byte**, byte**, int> import = &StudioNativeExports.DocumentImportPagesJson;
        delegate* unmanaged<nint, void> close = &StudioNativeExports.DocumentClose;
        delegate* unmanaged<void*, void> free = &StudioNativeExports.Free;

        nint document = 0; byte* error = null; byte* json = null;
        fixed (byte* location = Encoding.UTF8.GetBytes(target + "\0")) Assert.Equal(StudioNativeAbi.Ok, open(location, null, 0, &document, &error));
        try {
            string request = JsonSerializer.Serialize(new { insertBefore = 2, sources = new object[] { new { path = plain }, new { path = locked } } });
            fixed (byte* body = Encoding.UTF8.GetBytes(request + "\0")) Assert.Equal(StudioNativeAbi.PasswordRequired, import(document, body, 0, &json, &error));
            using (JsonDocument needed = JsonDocument.Parse(Marshal.PtrToStringUTF8((nint)json)!)) {
                Assert.Equal(1, needed.RootElement.GetProperty("source").GetInt32());
                Assert.Equal("locked.pdf", needed.RootElement.GetProperty("name").GetString());
            }
            free(json); free(error);

            request = JsonSerializer.Serialize(new { insertBefore = 2, sources = new object[] { new { path = plain }, new { path = locked, password = "open" } } });
            fixed (byte* body = Encoding.UTF8.GetBytes(request + "\0")) Assert.Equal(StudioNativeAbi.Ok, import(document, body, 0, &json, &error));
            using JsonDocument done = JsonDocument.Parse(Marshal.PtrToStringUTF8((nint)json)!);
            free(json);
            Assert.Equal(2, done.RootElement.GetProperty("imported").GetInt32());
        } finally { close(document); }
    }

    private string CreateFormPdf() {
        string path = Path.Combine(_root, "form.pdf");
        File.WriteAllBytes(path, PdfDocument.Create(document => document.Content(content => content.Canvas(canvas => {
            canvas.TextField("Name", null, 72, 72, 200, 24);
            canvas.CheckBox("Subscribe", false, 72, 120, 16, 16);
            canvas.RadioButton("Plan", "Basic", true, 72, 160, 16, 16);
            canvas.RadioButton("Plan", "Pro", false, 120, 160, 16, 16);
            canvas.ChoiceField("Country", ["Poland", "Germany"], null, 72, 200, 150, 22);
        }))).ToBytes());
        return path;
    }

    [Fact]
    public async Task Form_fields_list_their_kinds_widgets_and_values_and_fill_as_one_undoable_change() {
        using StudioNativeDocument document = await StudioNativeDocument.OpenAsync(CreateFormPdf(), null, default);
        Assert.True((await document.GetInfoAsync(default)).CanFillForms);

        StudioNativeFormFields form = await document.GetFormFieldsAsync(1, default);
        Assert.True(form.CanFill);
        Assert.Equal(new[] { "Country", "Name", "Plan", "Subscribe" }, form.Fields.Select(static field => field.Name).Order());
        StudioNativeFormField name = form.Fields.Single(static field => field.Name == "Name");
        Assert.Equal("text", name.Kind);
        Assert.True(name.Widgets[0].Bounds.Width > 150 && name.Widgets[0].Bounds.Height > 10);
        StudioNativeFormField subscribe = form.Fields.Single(static field => field.Name == "Subscribe");
        Assert.Equal("checkbox", subscribe.Kind);
        Assert.False(subscribe.Widgets[0].IsOn);
        StudioNativeFormField plan = form.Fields.Single(static field => field.Name == "Plan");
        Assert.Equal("radio", plan.Kind);
        Assert.Equal(2, plan.Widgets.Count);
        Assert.Equal(new[] { "Poland", "Germany" }, form.Fields.Single(static field => field.Name == "Country").Options.Select(static option => option.Value));

        await document.FillFormAsync(new StudioNativeFormFill(new Dictionary<string, IReadOnlyList<string>> {
            ["Name"] = ["Ada Lovelace"], ["Subscribe"] = [subscribe.Widgets[0].OnState!], ["Plan"] = [plan.Widgets[1].OnState!], ["Country"] = ["Germany"]
        }), default);

        StudioNativeFormFields filled = await document.GetFormFieldsAsync(1, default);
        Assert.Equal("Ada Lovelace", filled.Fields.Single(static field => field.Name == "Name").Values.Single());
        Assert.True(filled.Fields.Single(static field => field.Name == "Subscribe").Widgets[0].IsOn);
        Assert.Equal(new[] { false, true }, filled.Fields.Single(static field => field.Name == "Plan").Widgets.Select(static widget => widget.IsOn));
        Assert.Equal("Germany", filled.Fields.Single(static field => field.Name == "Country").Values.Single());

        await document.UndoAsync(default);
        Assert.DoesNotContain((await document.GetFormFieldsAsync(1, default)).Fields.Single(static field => field.Name == "Name").Values, static value => value.Length > 0);
    }

    [Fact]
    public async Task Filling_refuses_fields_the_form_does_not_have() {
        using StudioNativeDocument document = await StudioNativeDocument.OpenAsync(CreateFormPdf(), null, default);
        await Assert.ThrowsAsync<ArgumentException>(() => document.FillFormAsync(new StudioNativeFormFill(null), default));
        await Assert.ThrowsAsync<ArgumentException>(() => document.FillFormAsync(new StudioNativeFormFill(
            new Dictionary<string, IReadOnlyList<string>> { ["Missing"] = ["x"] }), default));
        await Assert.ThrowsAsync<ArgumentOutOfRangeException>(() => document.GetFormFieldsAsync(9, default));
        Assert.False((await document.GetInfoAsync(default)).IsDirty);
    }

    [Fact]
    public unsafe void The_C_interface_lists_and_fills_form_fields() {
        string path = CreateFormPdf();
        delegate* unmanaged<byte*, byte*, nint, nint*, byte**, int> open = &StudioNativeExports.DocumentOpen;
        delegate* unmanaged<nint, int, nint, byte**, byte**, int> fields = &StudioNativeExports.DocumentFormFieldsJson;
        delegate* unmanaged<nint, byte*, nint, byte**, int> fill = &StudioNativeExports.DocumentFillFormJson;
        delegate* unmanaged<nint, void> close = &StudioNativeExports.DocumentClose;
        delegate* unmanaged<void*, void> free = &StudioNativeExports.Free;
        nint document = 0; byte* error = null; byte* json = null;
        fixed (byte* location = Encoding.UTF8.GetBytes(path + "\0")) Assert.Equal(StudioNativeAbi.Ok, open(location, null, 0, &document, &error));
        try {
            fixed (byte* body = Encoding.UTF8.GetBytes("{\"values\":{\"Name\":[\"Grace\"]}}\0")) Assert.Equal(StudioNativeAbi.Ok, fill(document, body, 0, &error));
            fixed (byte* body = Encoding.UTF8.GetBytes("{\"values\":{\"Nope\":[\"x\"]}}\0")) Assert.Equal(StudioNativeAbi.InvalidArgument, fill(document, body, 0, &error));
            free(error);
            Assert.Equal(StudioNativeAbi.Ok, fields(document, 1, 0, &json, &error));
            using JsonDocument form = JsonDocument.Parse(Marshal.PtrToStringUTF8((nint)json)!);
            free(json);
            JsonElement name = form.RootElement.GetProperty("fields").EnumerateArray().Single(field => field.GetProperty("name").GetString() == "Name");
            Assert.Equal("Grace", name.GetProperty("values")[0].GetString());
            Assert.Equal("text", name.GetProperty("kind").GetString());
        } finally { close(document); }
    }

    [Fact]
    public async Task Extracted_pages_form_a_new_document_and_leave_the_open_one_unchanged() {
        using StudioNativeDocument document = await StudioNativeDocument.OpenAsync(CreatePagesOfWidths(300, 310, 320), null, default);
        long revision = (await document.GetInfoAsync(default)).Revision;
        string extracted = Path.Combine(_root, "extracted.pdf");

        await document.ExtractPagesAsync([3, 1], extracted, default);

        using StudioNativeDocument copy = await StudioNativeDocument.OpenAsync(extracted, null, default);
        Assert.Equal(new[] { 300D, 320D }, await WidthsAsync(copy));
        StudioNativeDocumentInfo info = await document.GetInfoAsync(default);
        Assert.Equal(revision, info.Revision);
        Assert.False(info.IsDirty);
    }

    [Fact]
    public async Task Search_finds_text_with_rectangles_on_the_page_as_displayed() {
        string path = Path.Combine(_root, "text.pdf");
        PdfDocument.Create(document => {
            foreach (string text in new[] { "Alpha needle here", "Nothing to see", "Another NEEDLE" })
                document.Page(page => page.Size(400, 300).Content(content => content.Column(column => column.Item().Paragraph(p => p.Text(text)))));
        }).Save(path);
        using StudioNativeDocument document = await StudioNativeDocument.OpenAsync(path, null, default);

        StudioNativeSearchResult result = await document.SearchAsync("  needle ", default);

        Assert.Equal("needle", result.Query);
        Assert.Equal(new[] { (1, 1), (3, 2) }, result.Hits.Select(hit => (hit.Page, hit.Occurrence)));
        Assert.All(result.Hits, hit => {
            Assert.True(hit.Bounds.Width > 0 && hit.Bounds.Height > 0);
            Assert.InRange(hit.Bounds.X + hit.Bounds.Width, 0D, 400D);
            Assert.InRange(hit.Bounds.Y + hit.Bounds.Height, 0D, 300D);
            Assert.Single(hit.LineBounds);
        });

        // Rotating the page turns the rectangles with it, into the displayed page's space.
        await document.RotatePagesAsync([1], 90, default);
        StudioNativeSearchHit rotated = (await document.SearchAsync("needle", default)).Hits[0];
        Assert.True(rotated.Bounds.Height > rotated.Bounds.Width);
        Assert.InRange(rotated.Bounds.Y + rotated.Bounds.Height, 0D, 400D);
    }

    [Fact]
    public async Task Streamed_search_reports_each_page_in_order_with_the_hits_of_the_full_result() {
        string path = Path.Combine(_root, "text.pdf");
        PdfDocument.Create(document => {
            foreach (string text in new[] { "Alpha needle here", "Nothing to see", "Another NEEDLE and a needle" })
                document.Page(page => page.Size(400, 300).Content(content => content.Column(column => column.Item().Paragraph(p => p.Text(text)))));
        }).Save(path);
        using StudioNativeDocument document = await StudioNativeDocument.OpenAsync(path, null, default);
        var pages = new List<StudioNativeSearchPage>();

        StudioNativeSearchResult result = await document.SearchAsync("needle", default, pages.Add);

        Assert.Equal(new[] { 1, 2, 3 }, pages.Select(static page => page.PagesSearched));
        Assert.All(pages, static page => Assert.Equal(3, page.PageCount));
        Assert.Equal(new[] { 1, 0, 2 }, pages.Select(static page => page.Hits.Count));
        Assert.Equal(result.Hits.Select(static hit => (hit.Page, hit.Occurrence)),
            pages.SelectMany(static page => page.Hits).Select(static hit => (hit.Page, hit.Occurrence)));
        Assert.Equal(new[] { (1, 1), (3, 2), (3, 3) }, result.Hits.Select(static hit => (hit.Page, hit.Occurrence)));
    }

    [Fact]
    public async Task A_copy_carries_unsaved_changes_and_leaves_the_document_where_it_is() {
        string source = CreatePdf();
        byte[] original = File.ReadAllBytes(source);
        using StudioNativeDocument document = await StudioNativeDocument.OpenAsync(source, null, default);
        await document.RotatePagesAsync([1], 90, default);
        string copy = Path.Combine(_root, "copy.pdf");

        await document.SaveCopyAsync(copy, default);

        StudioNativeDocumentInfo info = await document.GetInfoAsync(default);
        Assert.True(info.IsDirty);
        Assert.Equal(source, info.Path);
        using (StudioNativeDocument reopened = await StudioNativeDocument.OpenAsync(copy, null, default))
            Assert.Equal(90, (await reopened.GetInfoAsync(default)).Pages[0].Rotation);
        Assert.Equal(original, File.ReadAllBytes(source));
        await Assert.ThrowsAsync<ArgumentException>(() => document.SaveCopyAsync(source, default));
    }

    [Fact]
    public async Task A_protected_copy_needs_its_password_and_an_unprotected_copy_of_it_does_not() {
        using StudioNativeDocument document = await StudioNativeDocument.OpenAsync(CreatePdf(), null, default);
        Assert.False((await document.GetInfoAsync(default)).IsEncrypted);
        string locked = Path.Combine(_root, "locked.pdf");

        await document.SaveProtectedCopyAsync(new StudioNativeProtectionRequest(locked, "open sesame", "owner", null), default);

        await Assert.ThrowsAsync<PdfPasswordRequiredException>(() => StudioNativeDocument.OpenAsync(locked, null, default));
        await Assert.ThrowsAsync<PdfInvalidPasswordException>(() => StudioNativeDocument.OpenAsync(locked, "wrong", default));
        using StudioNativeDocument opened = await StudioNativeDocument.OpenAsync(locked, "open sesame", default);
        Assert.True((await opened.GetInfoAsync(default)).IsEncrypted);

        string unlocked = Path.Combine(_root, "unlocked.pdf");
        await opened.SaveProtectedCopyAsync(new StudioNativeProtectionRequest(unlocked, null, null, "owner"), default);
        using StudioNativeDocument plain = await StudioNativeDocument.OpenAsync(unlocked, null, default);
        Assert.False((await plain.GetInfoAsync(default)).IsEncrypted);
        await Assert.ThrowsAsync<ArgumentException>(() => plain.SaveProtectedCopyAsync(new StudioNativeProtectionRequest(locked, null, null, null), default));
    }

    [Fact]
    public async Task A_signed_copy_verifies_and_a_wrong_certificate_password_is_reported() {
        using RSA key = RSA.Create(2048);
        var request = new CertificateRequest("CN=Studio Signer", key, HashAlgorithmName.SHA256, RSASignaturePadding.Pkcs1);
        request.CertificateExtensions.Add(new X509KeyUsageExtension(X509KeyUsageFlags.DigitalSignature | X509KeyUsageFlags.NonRepudiation, true));
        using X509Certificate2 certificate = request.CreateSelfSigned(DateTimeOffset.Now.AddDays(-1), DateTimeOffset.Now.AddYears(1));
        string identity = Path.Combine(_root, "signer.p12");
        File.WriteAllBytes(identity, certificate.Export(X509ContentType.Pkcs12, "p12 password"));
        using StudioNativeDocument document = await StudioNativeDocument.OpenAsync(CreatePdf(), null, default);
        Assert.True((await document.GetInfoAsync(default)).CanSign);
        string signed = Path.Combine(_root, "signed.pdf");

        await Assert.ThrowsAsync<StudioNativeCertificatePasswordException>(() =>
            document.SaveSignedCopyAsync(new StudioNativeSigningRequest(signed, identity, "wrong", null, null), default));
        await document.SaveSignedCopyAsync(new StudioNativeSigningRequest(signed, identity, "p12 password", "Approved", "Kraków"), default);

        PdfSignatureValidationReport report = PdfDocument.Load(signed).Security
            .ValidateSignatures(new PdfCmsSignatureCryptographyProvider(OfficeSecurityProvider.Default));
        PdfSignatureValidationResult signature = Assert.Single(report.Signatures);
        Assert.True(signature.IsStructurallyValid);
        Assert.True(signature.CryptographicResult?.IsMathematicallyValid);
        Assert.False((await document.GetInfoAsync(default)).IsDirty);
    }

    [Fact]
    public unsafe void The_C_interface_reorders_pages_and_returns_search_hits_as_json() {
        string path = Path.Combine(_root, "text.pdf");
        PdfDocument.Create(document => {
            foreach (string text in new[] { "first page", "second needle" })
                document.Page(page => page.Size(400, 300).Content(content => content.Column(column => column.Item().Paragraph(p => p.Text(text)))));
        }).Save(path);
        delegate* unmanaged<byte*, byte*, nint, nint*, byte**, int> open = &StudioNativeExports.DocumentOpen;
        delegate* unmanaged<nint, int*, int, nint, byte**, int> reorder = &StudioNativeExports.DocumentReorderPages;
        delegate* unmanaged<nint, byte*, nint, byte**, byte**, int> search = &StudioNativeExports.DocumentSearchJson;
        delegate* unmanaged<nint, void> close = &StudioNativeExports.DocumentClose;
        delegate* unmanaged<void*, void> free = &StudioNativeExports.Free;

        nint document = 0; byte* error = null;
        fixed (byte* location = Encoding.UTF8.GetBytes(path + "\0")) Assert.Equal(StudioNativeAbi.Ok, open(location, null, 0, &document, &error));
        try {
            int* order = stackalloc int[] { 2, 1 };
            Assert.Equal(StudioNativeAbi.Ok, reorder(document, order, 2, 0, &error));
            Assert.Equal(StudioNativeAbi.InvalidArgument, reorder(document, order, 1, 0, &error));
            free(error);

            byte* json = null;
            fixed (byte* query = Encoding.UTF8.GetBytes("needle\0")) Assert.Equal(StudioNativeAbi.Ok, search(document, query, 0, &json, &error));
            using JsonDocument hits = JsonDocument.Parse(Marshal.PtrToStringUTF8((nint)json)!);
            free(json);
            JsonElement hit = Assert.Single(hits.RootElement.GetProperty("hits").EnumerateArray());
            Assert.Equal(1, hit.GetProperty("page").GetInt32());
            Assert.True(hit.GetProperty("bounds").GetProperty("width").GetDouble() > 0);
            Assert.Equal(1, hits.RootElement.GetProperty("abiVersion").GetInt32());
        } finally { close(document); }
    }

    [Fact]
    public async Task Ink_markup_keeps_its_width_and_opacity_and_lists_where_it_was_drawn() {
        using StudioNativeDocument document = await StudioNativeDocument.OpenAsync(CreatePdf(1), null, default);
        Assert.True((await document.GetInfoAsync(default)).CanEditAnnotations);
        StudioNativePoint[] stroke = [new(50, 60), new(120, 80), new(200, 70)];

        await document.AddAnnotationAsync(new StudioNativeAnnotationRequest("ink", 1, Strokes: [stroke],
            Color: new StudioNativeColor(0.9, 0.1, 0.1), StrokeWidth: 6, Opacity: 0.5, Author: "Reviewer"), default);

        StudioNativeAnnotation ink = Assert.Single((await document.GetAnnotationsAsync(1, default)).Annotations);
        Assert.Equal("Ink", ink.Subtype);
        Assert.Equal("Reviewer", ink.Author);
        // The listed area covers the stroke in the displayed page's points, widened by half the stroke.
        Assert.InRange(ink.Bounds.X, 40D, 50D);
        Assert.InRange(ink.Bounds.Y, 50D, 60D);
        Assert.InRange(ink.Bounds.X + ink.Bounds.Width, 200D, 210D);
        Assert.InRange(ink.Bounds.Y + ink.Bounds.Height, 80D, 90D);
        PdfAnnotation saved = Assert.Single(PdfDocument.Load(await SavedBytesAsync(document)).Annotations.GetForEditing());
        Assert.Equal(6D, saved.BorderWidth);
        Assert.Equal(0.5D, saved.Opacity);
        Assert.True(saved.HasNormalAppearance);
    }

    [Fact]
    public async Task Notes_change_and_go_as_undoable_edits() {
        using StudioNativeDocument document = await StudioNativeDocument.OpenAsync(CreatePdf(1), null, default);
        await document.AddAnnotationAsync(new StudioNativeAnnotationRequest("note", 1, Bounds: new StudioNativeRectangle(100, 100, 18, 18),
            Text: "Check this figure", Author: "Ann"), default);
        StudioNativeAnnotation note = Assert.Single((await document.GetAnnotationsAsync(1, default)).Annotations);
        Assert.Equal(("Text", "Check this figure"), (note.Subtype, note.Contents));

        await document.UpdateAnnotationAsync(new StudioNativeAnnotationUpdate(note.ObjectNumber, Contents: "Figure checked"), default);
        StudioNativeAnnotation changed = Assert.Single((await document.GetAnnotationsAsync(1, default)).Annotations);
        Assert.Equal(("Figure checked", "Ann"), (changed.Contents, changed.Author));

        await document.RemoveAnnotationAsync(changed.ObjectNumber, default);
        Assert.Empty((await document.GetAnnotationsAsync(1, default)).Annotations);
        await document.UndoAsync(default);
        Assert.Equal("Figure checked", Assert.Single((await document.GetAnnotationsAsync(1, default)).Annotations).Contents);
        await Assert.ThrowsAsync<ArgumentException>(() => document.RemoveAnnotationAsync(9999, default));
    }

    [Fact]
    public async Task Markup_on_a_rotated_page_lands_where_it_was_drawn_on_screen() {
        using StudioNativeDocument document = await StudioNativeDocument.OpenAsync(CreatePdf(1), null, default);
        await document.RotatePagesAsync([1], 90, default);
        // Displayed, the 300 x 400 page is 400 wide and 300 high.
        var drawn = new StudioNativeRectangle(300, 40, 60, 30);

        await document.AddAnnotationAsync(new StudioNativeAnnotationRequest("rectangle", 1, Bounds: drawn), default);

        StudioNativeRectangle listed = Assert.Single((await document.GetAnnotationsAsync(1, default)).Annotations).Bounds;
        Assert.Equal(drawn.X, listed.X, 1);
        Assert.Equal(drawn.Y, listed.Y, 1);
        Assert.Equal(drawn.Width, listed.Width, 1);
        Assert.Equal(drawn.Height, listed.Height, 1);
    }

    [Fact]
    public async Task Markup_requests_outside_the_contract_are_refused() {
        using StudioNativeDocument document = await StudioNativeDocument.OpenAsync(CreatePdf(1), null, default);

        await Assert.ThrowsAsync<ArgumentException>(() => document.AddAnnotationAsync(new StudioNativeAnnotationRequest("redact", 1,
            Bounds: new StudioNativeRectangle(10, 10, 20, 20)), default));
        await Assert.ThrowsAsync<ArgumentException>(() => document.AddAnnotationAsync(new StudioNativeAnnotationRequest("ink", 1), default));
        await Assert.ThrowsAsync<ArgumentOutOfRangeException>(() => document.AddAnnotationAsync(new StudioNativeAnnotationRequest("rectangle", 2,
            Bounds: new StudioNativeRectangle(10, 10, 20, 20)), default));
        await Assert.ThrowsAsync<ArgumentOutOfRangeException>(() => document.AddAnnotationAsync(new StudioNativeAnnotationRequest("ink", 1,
            Strokes: [[new(1, 1), new(5, 5)]], Opacity: 2), default));
        Assert.False((await document.GetInfoAsync(default)).IsDirty);
    }

    [Fact]
    public unsafe void The_C_interface_adds_and_lists_annotations_as_json() {
        string source = CreatePdf(1);
        delegate* unmanaged<byte*, byte*, nint, nint*, byte**, int> open = &StudioNativeExports.DocumentOpen;
        delegate* unmanaged<nint, byte*, nint, byte**, int> add = &StudioNativeExports.DocumentAddAnnotationJson;
        delegate* unmanaged<nint, int, nint, byte**, byte**, int> list = &StudioNativeExports.DocumentAnnotationsJson;
        delegate* unmanaged<nint, void> close = &StudioNativeExports.DocumentClose;
        delegate* unmanaged<void*, void> free = &StudioNativeExports.Free;

        nint document = 0; byte* error = null;
        fixed (byte* path = Encoding.UTF8.GetBytes(source + "\0")) Assert.Equal(StudioNativeAbi.Ok, open(path, null, 0, &document, &error));
        try {
            const string request = """{"tool":"ink","page":1,"strokes":[[{"x":20,"y":20},{"x":80,"y":40}]],"color":{"red":0,"green":0.4,"blue":1},"strokeWidth":3}""";
            fixed (byte* body = Encoding.UTF8.GetBytes(request + "\0")) Assert.Equal(StudioNativeAbi.Ok, add(document, body, 0, &error));
            fixed (byte* body = Encoding.UTF8.GetBytes("{not json\0")) Assert.Equal(StudioNativeAbi.InvalidArgument, add(document, body, 0, &error));
            free(error);

            byte* json = null;
            Assert.Equal(StudioNativeAbi.Ok, list(document, 1, 0, &json, &error));
            using JsonDocument result = JsonDocument.Parse(Marshal.PtrToStringUTF8((nint)json)!);
            free(json);
            JsonElement ink = Assert.Single(result.RootElement.GetProperty("annotations").EnumerateArray());
            Assert.Equal("Ink", ink.GetProperty("subtype").GetString());
            Assert.True(ink.GetProperty("objectNumber").GetInt32() > 0);
            Assert.Equal(0.4D, ink.GetProperty("color").GetProperty("green").GetDouble(), 3);
        } finally { close(document); }
    }

    [Fact]
    public async Task Page_text_gives_each_element_a_box_on_the_page_as_displayed() {
        string path = Path.Combine(_root, "text.pdf");
        PdfDocument.Create(document => document.Page(page => page.Size(400, 300)
            .Content(content => content.Column(column => column.Item().Paragraph(p => p.Text("Hello world")))))).Save(path);
        using StudioNativeDocument document = await StudioNativeDocument.OpenAsync(path, null, default);

        StudioNativePageText text = await document.GetPageTextAsync(1, default);

        Assert.Equal("Hello world", string.Concat(text.Texts).Trim());
        Assert.Equal(text.Texts.Count * 4, text.Boxes.Count);
        for (int index = 0; index < text.Texts.Count; index++) {
            Assert.InRange(text.Boxes[index * 4], 0D, 400D);
            Assert.InRange(text.Boxes[index * 4 + 1], 0D, 300D);
        }
        // Elements read left to right along the line.
        Assert.True(text.Boxes[4] > text.Boxes[0]);

        await document.RotatePagesAsync([1], 90, default);
        StudioNativePageText rotated = await document.GetPageTextAsync(1, default);
        // On the page turned a quarter, the 300 x 400 page is shown 400 high and the line runs down it.
        int last = (rotated.Texts.Count - 1) * 4;
        Assert.True(Math.Abs(rotated.Boxes[last + 1] - rotated.Boxes[1]) > Math.Abs(rotated.Boxes[last] - rotated.Boxes[0]),
            "The rotated line runs down the displayed page.");
        for (int index = 0; index < rotated.Texts.Count; index++) {
            Assert.InRange(rotated.Boxes[index * 4], 0D, 300D);
            Assert.InRange(rotated.Boxes[index * 4 + 1], 0D, 400D);
        }
        Assert.True(rotated.Revision > text.Revision);
    }

    // A blank page image, as a camera would give, rendered by the engine itself.
    private async Task<string> CreateScanImageAsync(string name, double width, double height) {
        string pdf = Path.Combine(_root, name + ".pdf");
        PdfDocument.Create(document => document.Page(page => page.Size(width, height))).Save(pdf);
        using StudioNativeDocument document = await StudioNativeDocument.OpenAsync(pdf, null, default);
        string image = Path.Combine(_root, name + ".png");
        File.WriteAllBytes(image, (await document.RenderPageAsync(1, 2D, default)).Bytes);
        return image;
    }

    [Fact]
    public async Task A_scan_becomes_one_page_per_image_with_its_recognized_words_searchable() {
        string first = await CreateScanImageAsync("receipt", 300, 400);
        string second = await CreateScanImageAsync("wide", 400, 300);
        string output = Path.Combine(_root, "scan.pdf");

        PdfScanResult result = await PdfScanComposer.ComposeAsync([
            new PdfScanPage(first, [new PdfScanWord("Invoice", 0.1, 0.1, 0.3, 0.05, 0.98), new PdfScanWord("4711", 0.407, 0.1, 0.15, 0.05)]),
            new PdfScanPage(second, [])
        ], output, default);

        Assert.Equal((2, 1, 2), (result.Pages, result.SearchablePages, result.Words));
        using StudioNativeDocument scan = await StudioNativeDocument.OpenAsync(output, null, default);
        StudioNativeDocumentInfo info = await scan.GetInfoAsync(default);
        Assert.Equal(2, info.Pages.Count);
        Assert.True(info.Pages[0].Height > info.Pages[0].Width && info.Pages[1].Width > info.Pages[1].Height);
        StudioNativeSearchHit hit = Assert.Single((await scan.SearchAsync("invoice", default)).Hits);
        Assert.Equal(1, hit.Page);
        // The word sits where the recognizer found it, a tenth of the way in from the top left.
        Assert.InRange(hit.Bounds.X / info.Pages[0].Width, 0.08, 0.14);
        Assert.InRange(hit.Bounds.Y / info.Pages[0].Height, 0.07, 0.14);
        // Recognized words sit a few points apart, as a recognizer's word boxes do; the gap reads as a space.
        StudioNativePageText text = await scan.GetPageTextAsync(1, default);
        Assert.Equal("Invoice4711", string.Concat(text.Texts));
        Assert.Equal(text.Texts.Count, text.Breaks.Count);
        Assert.Equal([2, 0, 0, 0, 0, 0, 0, 1, 0, 0, 0], text.Breaks);
    }

    [Fact]
    public async Task A_scanned_document_becomes_searchable_as_one_undoable_change() {
        string image = await CreateScanImageAsync("paper", 300, 400);
        string scan = Path.Combine(_root, "image-only.pdf");
        await PdfScanComposer.ComposeAsync([new PdfScanPage(image, []), new PdfScanPage(image, [])], scan, default);
        using StudioNativeDocument document = await StudioNativeDocument.OpenAsync(scan, null, default);
        Assert.True((await document.GetInfoAsync(default)).CanMakeSearchable);
        Assert.Empty((await document.SearchAsync("quarterly", default)).Hits);
        var seen = new List<(int Page, bool IsPng)>();

        StudioNativeSearchableResult result = await document.MakeSearchableAsync((page, png, _) => {
            seen.Add((page, png.Length > 8 && png[1] == (byte)'P' && png[2] == (byte)'N' && png[3] == (byte)'G'));
            return page == 2 ? [new PdfScanWord("Quarterly", 0.2, 0.3, 0.4, 0.05, 0.9)] : [];
        }, default);

        Assert.Equal([(1, true), (2, true)], seen);
        Assert.Equal((1, 1), (result.Pages, result.Words));
        StudioNativeSearchHit hit = Assert.Single((await document.SearchAsync("quarterly", default)).Hits);
        Assert.Equal(2, hit.Page);
        await document.UndoAsync(default);
        Assert.Empty((await document.SearchAsync("quarterly", default)).Hits);
    }

    [Theory]
    [InlineData(90)]
    [InlineData(180)]
    [InlineData(270)]
    public async Task Text_recognized_on_a_rotated_page_lands_where_the_recognizer_saw_it(int rotation) {
        string image = await CreateScanImageAsync($"turned-{rotation}", 300, 400);
        string scan = Path.Combine(_root, $"turned-{rotation}.pdf");
        await PdfScanComposer.ComposeAsync([new PdfScanPage(image, [])], scan, default);
        using StudioNativeDocument document = await StudioNativeDocument.OpenAsync(scan, null, default);
        await document.RotatePagesAsync([1], rotation, default);
        StudioNativePageInfo page = (await document.GetInfoAsync(default)).Pages[0];
        double width = rotation % 180 == 0 ? page.Width : page.Height, height = rotation % 180 == 0 ? page.Height : page.Width;

        await document.MakeSearchableAsync((_, _, _) => [new PdfScanWord("Quarterly", 0.2, 0.3, 0.4, 0.05, 0.9)], default);

        StudioNativeSearchHit hit = Assert.Single((await document.SearchAsync("quarterly", default)).Hits);
        // The hit is where the word was on the image the recognizer read: the page as displayed.
        Assert.InRange(hit.Bounds.X / width, 0.17, 0.23);
        Assert.InRange(hit.Bounds.Y / height, 0.27, 0.33);
        Assert.InRange(hit.Bounds.Width / width, 0.35, 0.45);
    }

    [Fact]
    public async Task Making_a_digital_document_searchable_leaves_its_text_alone() {
        string path = Path.Combine(_root, "digital.pdf");
        PdfDocument.Create(document => document.Page(page => page.Size(400, 300)
            .Content(content => content.Column(column => column.Item().Paragraph(p => p.Text("Already searchable")))))).Save(path);
        using StudioNativeDocument document = await StudioNativeDocument.OpenAsync(path, null, default);
        long revision = (await document.GetInfoAsync(default)).Revision;

        // A recognizer that reads the page's own text back finds nothing new to add.
        StudioNativePageText text = await document.GetPageTextAsync(1, default);
        double width = text.Boxes.Where((_, index) => index % 4 == 2).Sum(), x = text.Boxes[0], y = text.Boxes[1], height = text.Boxes[3];
        StudioNativeSearchableResult result = await document.MakeSearchableAsync((_, _, _) =>
            [new PdfScanWord("Already", x / 400, y / 300, 0.2, height / 300)], default);

        Assert.Equal(0, result.Words);
        Assert.Equal(revision, (await document.GetInfoAsync(default)).Revision);
        using var cancel = new CancellationTokenSource();
        await Assert.ThrowsAnyAsync<OperationCanceledException>(() => document.MakeSearchableAsync((_, _, _) => {
            cancel.Cancel();
            return [];
        }, cancel.Token));
        Assert.Equal(revision, (await document.GetInfoAsync(default)).Revision);
    }

    [Fact]
    public unsafe void The_C_interface_makes_a_scan_searchable_through_the_host_recognizer() {
        string image = CreateScanImageAsync("c-paper", 300, 400).GetAwaiter().GetResult();
        string scan = Path.Combine(_root, "c-image-only.pdf");
        PdfScanComposer.ComposeAsync([new PdfScanPage(image, [])], scan, default).GetAwaiter().GetResult();
        delegate* unmanaged<byte*, byte*, nint, nint*, byte**, int> open = &StudioNativeExports.DocumentOpen;
        delegate* unmanaged<nint, delegate* unmanaged<nint, int, int, byte*, nuint, byte**, int>, nint, nint, byte**, byte**, int> searchable =
            &StudioNativeExports.DocumentMakeSearchableJson;
        delegate* unmanaged<nint, void> close = &StudioNativeExports.DocumentClose;
        delegate* unmanaged<void*, void> free = &StudioNativeExports.Free;

        nint document = 0; byte* error = null; byte* json = null;
        fixed (byte* location = Encoding.UTF8.GetBytes(scan + "\0")) Assert.Equal(StudioNativeAbi.Ok, open(location, null, 0, &document, &error));
        try {
            Assert.Equal(StudioNativeAbi.Ok, searchable(document, &RecognizeInvoice, 7, 0, &json, &error));
            using JsonDocument result = JsonDocument.Parse(Marshal.PtrToStringUTF8((nint)json)!);
            free(json);
            Assert.Equal(1, result.RootElement.GetProperty("words").GetInt32());
            Assert.Equal(StudioNativeAbi.Failed, searchable(document, &RecognizeNothing, 0, 0, &json, &error));
            Assert.Contains("page 1", Marshal.PtrToStringUTF8((nint)error));
            free(error);
        } finally { close(document); }
    }

    [UnmanagedCallersOnly]
    private static unsafe int RecognizeInvoice(nint context, int page, int pageCount, byte* png, nuint length, byte** words) {
        if (context != 7 || page != 1 || pageCount != 1 || length == 0) return StudioNativeAbi.Failed;
        byte[] text = Encoding.UTF8.GetBytes("""[{"text":"Invoice","x":0.1,"y":0.1,"width":0.3,"height":0.05,"confidence":0.95}]""" + "\0");
        *words = (byte*)NativeMemory.Alloc((nuint)text.Length);
        text.CopyTo(new Span<byte>(*words, text.Length));
        return StudioNativeAbi.Ok;
    }

    [UnmanagedCallersOnly]
    private static unsafe int RecognizeNothing(nint context, int page, int pageCount, byte* png, nuint length, byte** words) => StudioNativeAbi.Failed;

    [Fact]
    public async Task Word_Excel_and_PowerPoint_files_convert_to_searchable_PDFs() {
        string word = Path.Combine(_root, "report.docx");
        using (var document = OfficeIMO.Word.WordDocument.Create(word)) {
            document.AddParagraph("Quarterly report");
            document.Save();
        }
        string workbook = Path.Combine(_root, "figures.xlsx");
        using (var document = OfficeIMO.Excel.ExcelDocument.Create(workbook)) {
            document.AddWorksheet("Data").CellValue(1, 1, "Quarterly figures");
            document.Save();
        }
        string deck = Path.Combine(_root, "deck.pptx");
        using (var presentation = OfficeIMO.PowerPoint.PowerPointPresentation.Create(deck)) {
            presentation.AddSlide().AddTextBox("Quarterly review");
            presentation.Save();
        }

        foreach ((string input, string route) in new[] { (word, "docx-pdf"), (workbook, "xlsx-pdf"), (deck, "pptx-pdf") }) {
            string output = Path.ChangeExtension(input, ".pdf");
            PdfConversionResult result = await PdfConversion.ConvertAsync(input, output, default);
            Assert.Equal(route, result.Route);
            Assert.Equal(new FileInfo(output).Length, result.OutputBytes);
            using StudioNativeDocument converted = await StudioNativeDocument.OpenAsync(output, null, default);
            Assert.NotEmpty((await converted.SearchAsync("quarterly", default)).Hits);
        }
    }

    [Fact]
    public async Task Conversion_names_its_formats_and_refuses_what_it_cannot_convert() {
        Assert.Contains(PdfConversion.Sources, source => source.Extensions.Contains(".docx"));
        Assert.Equal("txt-pdf", PdfConversion.RouteFor("notes.txt")?.Route);
        Assert.Equal("markdown-pdf", PdfConversion.RouteFor("README.md")?.Route);
        Assert.Null(PdfConversion.RouteFor("photo.heic"));

        string unknown = Path.Combine(_root, "photo.heic");
        File.WriteAllBytes(unknown, [1, 2, 3]);
        await Assert.ThrowsAsync<NotSupportedException>(() => PdfConversion.ConvertAsync(unknown, Path.Combine(_root, "photo.pdf"), default));
        await Assert.ThrowsAsync<FileNotFoundException>(() => PdfConversion.ConvertAsync(Path.Combine(_root, "missing.docx"), Path.Combine(_root, "missing.pdf"), default));
        string broken = Path.Combine(_root, "broken.docx");
        File.WriteAllText(broken, "not a document");
        var failure = await Assert.ThrowsAsync<InvalidOperationException>(() => PdfConversion.ConvertAsync(broken, Path.Combine(_root, "broken.pdf"), default));
        Assert.Contains("broken.docx", failure.Message);
        Assert.False(File.Exists(Path.Combine(_root, "broken.pdf")));
    }

    [Fact]
    public unsafe void The_C_interface_lists_formats_and_converts_a_document() {
        string text = Path.Combine(_root, "notes.txt");
        File.WriteAllText(text, "Quarterly notes");
        delegate* unmanaged<byte**, byte**, int> sources = &StudioNativeExports.ConversionSourcesJson;
        delegate* unmanaged<byte*, byte*, nint, byte**, byte**, int> convert = &StudioNativeExports.ConvertToPdfJson;
        delegate* unmanaged<void*, void> free = &StudioNativeExports.Free;
        byte* json = null; byte* error = null;

        Assert.Equal(StudioNativeAbi.Ok, sources(&json, &error));
        using (JsonDocument list = JsonDocument.Parse(Marshal.PtrToStringUTF8((nint)json)!))
            Assert.Contains(list.RootElement.GetProperty("sources").EnumerateArray(), source => source.GetProperty("route").GetString() == "pptx-pdf");
        free(json);

        string output = Path.Combine(_root, "notes.pdf");
        fixed (byte* input = Encoding.UTF8.GetBytes(text + "\0"))
        fixed (byte* destination = Encoding.UTF8.GetBytes(output + "\0")) {
            Assert.Equal(StudioNativeAbi.Ok, convert(input, destination, 0, &json, &error));
            using JsonDocument result = JsonDocument.Parse(Marshal.PtrToStringUTF8((nint)json)!);
            Assert.Equal("txt-pdf", result.RootElement.GetProperty("route").GetString());
            free(json);
        }
        Assert.True(File.Exists(output));
        fixed (byte* input = Encoding.UTF8.GetBytes(Path.Combine(_root, "photo.heic") + "\0"))
        fixed (byte* destination = Encoding.UTF8.GetBytes(output + "\0")) {
            File.WriteAllBytes(Path.Combine(_root, "photo.heic"), [1]);
            Assert.Equal(StudioNativeAbi.Unsupported, convert(input, destination, 0, &json, &error));
            free(error);
        }
    }

    [Fact]
    public async Task A_scan_refuses_what_it_cannot_make() {
        string image = await CreateScanImageAsync("page", 300, 400);
        await Assert.ThrowsAsync<ArgumentException>(() => PdfScanComposer.ComposeAsync([], Path.Combine(_root, "none.pdf"), default));
        await Assert.ThrowsAsync<ArgumentException>(() => PdfScanComposer.ComposeAsync([new PdfScanPage(image, [])], "relative.pdf", default));
        await Assert.ThrowsAnyAsync<IOException>(() => PdfScanComposer.ComposeAsync([new PdfScanPage(Path.Combine(_root, "missing.png"), [])],
            Path.Combine(_root, "missing.pdf"), default));
        Assert.False(File.Exists(Path.Combine(_root, "missing.pdf")));
    }

    private static async Task<byte[]> SavedBytesAsync(StudioNativeDocument document) {
        string copy = Path.Combine(Path.GetTempPath(), "officeimo-native-annotations-" + Guid.NewGuid().ToString("N") + ".pdf");
        try {
            await document.SaveAsync(copy, default);
            return File.ReadAllBytes(copy);
        } finally { File.Delete(copy); }
    }

    [Fact]
    public async Task Page_sizes_are_the_rendered_size_at_the_page_user_unit() {
        string path = Path.Combine(_root, "user-unit.pdf");
        File.WriteAllBytes(path, Encoding.ASCII.GetBytes("""
            %PDF-1.7
            1 0 obj << /Type /Catalog /Pages 2 0 R >> endobj
            2 0 obj << /Type /Pages /Count 1 /Kids [3 0 R] >> endobj
            3 0 obj << /Type /Page /Parent 2 0 R /MediaBox [0 0 300 400] /CropBox [0 0 200 300] /UserUnit 2 /Resources << >> >> endobj
            trailer << /Root 1 0 R /Size 4 >>
            %%EOF
            """));
        using StudioNativeDocument document = await StudioNativeDocument.OpenAsync(path, null, default);

        StudioNativePageInfo page = (await document.GetInfoAsync(default)).Pages[0];

        Assert.Equal((400D, 600D), (page.Width, page.Height));
        Assert.Equal(400, (await document.RenderPageAsync(1, 1D, default)).PixelWidth);
    }

    [Fact]
    public async Task Closing_discards_the_recovery_copy_of_unsaved_edits() {
        string source = CreatePdf();
        StudioNativeDocument document = await StudioNativeDocument.OpenAsync(source, null, default);
        await document.RotatePagesAsync([1], 90, default);
        Assert.NotEmpty(Directory.EnumerateFiles(Path.Combine(_root, "data", "Recovery")));

        document.Dispose();

        Assert.Empty(Directory.EnumerateFiles(Path.Combine(_root, "data", "Recovery"), "*.pdf"));
        using StudioNativeDocument reopened = await StudioNativeDocument.OpenAsync(source, null, default);
        Assert.False((await reopened.GetInfoAsync(default)).HasRecovery);
    }

    [Fact]
    public async Task Tiles_beyond_the_page_budget_join_into_the_zoomed_page() {
        using StudioNativeDocument document = await StudioNativeDocument.OpenAsync(CreatePdf(1), null, default);
        // 300 x 400 points at 10x is 12 million pixels, beyond the 8.4 million pixel budget, so it renders as tiles.
        const double scale = 10D;
        var tiles = new List<PdfRenderedPixels>();
        foreach ((int x, int y) in new[] { (0, 0), (1500, 0), (0, 2000), (1500, 2000) })
            tiles.Add(await document.RenderRegionAsync(1, scale, x, y, 1500, 2000, default));

        Assert.All(tiles, tile => Assert.Equal(1500 * 2000 * 4, tile.Pixels.Length));
        Assert.All(tiles, tile => Assert.True(tile.Pixels.Where((_, index) => index % 4 == 3).All(alpha => alpha == 255)));
        await Assert.ThrowsAsync<ArgumentOutOfRangeException>(() => document.RenderRegionAsync(1, scale, 0, 0, 3000, 4000, default));
    }

    [Fact]
    public unsafe void The_C_interface_renders_page_rectangles_as_pixels() {
        string source = CreatePdf(1);
        delegate* unmanaged<byte*, byte*, nint, nint*, byte**, int> open = &StudioNativeExports.DocumentOpen;
        delegate* unmanaged<nint, int, double, int, int, int, int, nint, byte**, nuint*, byte**, int> region = &StudioNativeExports.DocumentRenderRegion;
        delegate* unmanaged<nint, void> close = &StudioNativeExports.DocumentClose;
        delegate* unmanaged<void*, void> free = &StudioNativeExports.Free;

        nint document = 0; byte* error = null;
        fixed (byte* path = Encoding.UTF8.GetBytes(source + "\0")) Assert.Equal(StudioNativeAbi.Ok, open(path, null, 0, &document, &error));
        try {
            byte* pixels = null; nuint length = 0;
            Assert.Equal(StudioNativeAbi.Ok, region(document, 1, 1.5D, 0, 0, 450, 600, 0, &pixels, &length, &error));
            Assert.Equal((nuint)(450 * 600 * 4), length);
            Assert.Equal(255, pixels[3]);
            free(pixels);

            // Outside the page image, or more than the budget in one rectangle, is the host's argument to correct.
            Assert.Equal(StudioNativeAbi.InvalidArgument, region(document, 1, 1.5D, 400, 0, 100, 100, 0, &pixels, &length, &error));
            free(error);
            Assert.Equal(StudioNativeAbi.InvalidArgument, region(document, 1, 20D, 0, 0, 6000, 8000, 0, &pixels, &length, &error));
            free(error);
            Assert.True(pixels is null);
        } finally { close(document); }
    }

    [Fact]
    public async Task Rotation_preserves_stream_and_tagged_sources_through_native_workspace_history() {
        foreach (string name in new[] { "latin-0-native.pdf", "openpreserve-pdfa1b-text.pdf" }) {
            string source = Path.Combine(_root, name);
            File.Copy(Path.Combine(AppContext.BaseDirectory, "Fixtures", name), source);
            byte[] original = File.ReadAllBytes(source);
            using StudioNativeDocument document = await StudioNativeDocument.OpenAsync(source, null, default);
            await document.RotatePagesAsync([1], 90, default);
            Assert.Equal(90, (await document.GetInfoAsync(default)).Pages[0].Rotation);
            await document.RotatePagesAsync([1], 90, default);
            Assert.Equal(180, (await document.GetInfoAsync(default)).Pages[0].Rotation);
            await document.UndoAsync(default);
            Assert.Equal(90, (await document.GetInfoAsync(default)).Pages[0].Rotation);
            await document.UndoAsync(default);
            Assert.Equal(0, (await document.GetInfoAsync(default)).Pages[0].Rotation);
            await document.RedoAsync(default);
            string copy = Path.Combine(_root, "rotated-" + name);
            await document.SaveAsync(copy, default);
            byte[] output = File.ReadAllBytes(copy);
            Assert.True(output.AsSpan(0, original.Length).SequenceEqual(original));
            using StudioNativeDocument reopened = await StudioNativeDocument.OpenAsync(copy, null, default);
            Assert.Equal(90, (await reopened.GetInfoAsync(default)).Pages[0].Rotation);
            Assert.Equal(original, File.ReadAllBytes(source));
        }
    }

    [Fact]
    public unsafe void The_C_interface_reports_failures_as_status_and_message_without_throwing() {
        string source = CreatePdf(1);
        delegate* unmanaged<byte*, byte*, nint, nint*, byte**, int> open = &StudioNativeExports.DocumentOpen;
        delegate* unmanaged<nint, int, double, nint, byte**, nuint*, int*, int*, byte**, int> render = &StudioNativeExports.DocumentRenderPage;
        delegate* unmanaged<nint, byte**, byte**, int> info = &StudioNativeExports.DocumentInfoJson;
        delegate* unmanaged<nint, void> close = &StudioNativeExports.DocumentClose;
        delegate* unmanaged<void*, void> free = &StudioNativeExports.Free;
        delegate* unmanaged<nint> createCancellation = &StudioNativeExports.CancellationCreate;
        delegate* unmanaged<nint, void> cancel = &StudioNativeExports.CancellationCancel;
        delegate* unmanaged<nint, void> freeCancellation = &StudioNativeExports.CancellationFree;

        string Take(byte* text) { string value = Marshal.PtrToStringUTF8((nint)text)!; free(text); return value; }
        byte[] Utf8(string value) => Encoding.UTF8.GetBytes(value + "\0");

        nint document = 0; byte* error = null;
        byte[] missing = Utf8(Path.Combine(_root, "missing.pdf"));
        fixed (byte* path = missing) Assert.Equal(StudioNativeAbi.NotFound, open(path, null, 0, &document, &error));
        Assert.Equal(0, document);
        Assert.NotEmpty(Take(error));

        fixed (byte* path = Utf8(source)) Assert.Equal(StudioNativeAbi.Ok, open(path, null, 0, &document, &error));
        Assert.True(error is null);
        try {
            byte* json = null;
            Assert.Equal(StudioNativeAbi.Ok, info(document, &json, &error));
            using (JsonDocument facts = JsonDocument.Parse(Take(json))) {
                Assert.Equal(StudioNativeAbi.Version, facts.RootElement.GetProperty("abiVersion").GetInt32());
                Assert.Equal(1, facts.RootElement.GetProperty("pages").GetArrayLength());
                Assert.Equal(StudioPdfSecurityPolicy.MaximumRasterPixels, facts.RootElement.GetProperty("maxRenderPixels").GetInt64());
            }

            byte* png = null; nuint length = 0; int width = 0, height = 0;
            Assert.Equal(StudioNativeAbi.Ok, render(document, 1, 1D, 0, &png, &length, &width, &height, &error));
            Assert.Equal((nuint)new ReadOnlySpan<byte>(png, (int)length).Length, length);
            Assert.Equal((300, 400), (width, height));
            free(png);

            Assert.Equal(StudioNativeAbi.InvalidArgument, render(document, 5, 1D, 0, &png, &length, null, null, &error));
            Assert.True(png is null);
            Take(error);

            // A page image beyond the pixel budget is the host's argument to correct, not an engine failure.
            Assert.Equal(StudioNativeAbi.InvalidArgument, render(document, 1, 20D, 0, &png, &length, null, null, &error));
            Assert.True(png is null);
            Take(error);

            nint cancellation = createCancellation();
            cancel(cancellation);
            Assert.Equal(StudioNativeAbi.Cancelled, render(document, 1, 1D, cancellation, &png, &length, null, null, null));
            freeCancellation(cancellation);

            // A stale or invented handle is refused rather than dereferenced.
            Assert.Equal(StudioNativeAbi.InvalidArgument, info(0, &json, &error));
            Take(error);
        } finally { close(document); }
    }

    [Theory]
    [InlineData(typeof(OperationCanceledException), StudioNativeAbi.Cancelled)]
    [InlineData(typeof(FileNotFoundException), StudioNativeAbi.NotFound)]
    [InlineData(typeof(ArgumentOutOfRangeException), StudioNativeAbi.InvalidArgument)]
    [InlineData(typeof(NotSupportedException), StudioNativeAbi.Unsupported)]
    [InlineData(typeof(StudioNativeNotInitializedException), StudioNativeAbi.NotInitialized)]
    [InlineData(typeof(IOException), StudioNativeAbi.Failed)]
    public void Engine_errors_map_to_stable_status_codes(Type exception, int status) =>
        Assert.Equal(status, StudioNativeAbi.StatusOf((Exception)Activator.CreateInstance(exception)!));

    [Fact]
    public void Password_errors_ask_the_host_to_prompt() {
        Assert.Equal(StudioNativeAbi.PasswordRequired, StudioNativeAbi.StatusOf(new PdfPasswordRequiredException("Password required.")));
        Assert.Equal(StudioNativeAbi.InvalidPassword, StudioNativeAbi.StatusOf(new PdfInvalidPasswordException("Wrong password.")));
        Assert.Equal(StudioNativeAbi.Unsupported, StudioNativeAbi.StatusOf(new PdfUnsupportedEncryptionException("Unsupported encryption.")));
    }
}
