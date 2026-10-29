using System.Runtime.InteropServices;
using System.Text;
using System.Text.Json;
using OfficeIMO.Pdf;
using OfficeIMO.Studio.Native;

namespace OfficeIMO.Studio.Core.Tests;

/// <summary>Document structure and page production through the engine library the native Apple apps call.</summary>
public sealed class StudioNativeStructureTests : IDisposable {
    private readonly string _root = Path.Combine(Path.GetTempPath(), "officeimo-native-structure-" + Guid.NewGuid().ToString("N"));

    public StudioNativeStructureTests() {
        Directory.CreateDirectory(_root);
        StudioNativeEngine.Initialize(Path.Combine(_root, "data"));
    }

    public void Dispose() => Directory.Delete(_root, recursive: true);

    private string CreateTextPdf(string name, params string[] pages) {
        string path = Path.Combine(_root, name);
        PdfDocument.Create(document => {
            foreach (string text in pages)
                document.Page(page => page.Size(400, 300).Content(content => content.Column(column => column.Item().Paragraph(p => p.Text(text)))));
        }).Save(path);
        return path;
    }

    private static async Task<string> PageTextAsync(StudioNativeDocument document, int page) =>
        string.Concat((await document.GetPageTextAsync(page, default)).Texts);

    [Fact]
    public async Task Properties_change_as_one_undoable_edit_and_keep_what_the_request_leaves_out() {
        using StudioNativeDocument document = await StudioNativeDocument.OpenAsync(CreateTextPdf("properties.pdf", "One"), null, default);
        Assert.True((await document.GetStructureAsync(default)).CanEditProperties);

        await document.EditAsync(new StudioNativeEditRequest(Properties: new StudioNativeProperties("Quarterly report", "Alex Morgan", "Sales", "q3, sales")), default);
        await document.EditAsync(new StudioNativeEditRequest(Properties: new StudioNativeProperties(null, null, null, "")), default);

        StudioNativeProperties properties = (await document.GetStructureAsync(default)).Properties;
        Assert.Equal(("Quarterly report", "Alex Morgan", "Sales"), (properties.Title, properties.Author, properties.Subject));
        Assert.True(string.IsNullOrEmpty(properties.Keywords));
        await document.UndoAsync(default);
        Assert.Equal("q3, sales", (await document.GetStructureAsync(default)).Properties.Keywords);
    }

    [Fact]
    public async Task Bookmarks_are_added_nested_renamed_and_removed_with_one_undo_step_per_request() {
        using StudioNativeDocument document = await StudioNativeDocument.OpenAsync(CreateTextPdf("outline.pdf", "One", "Two", "Three"), null, default);
        Assert.Empty((await document.GetStructureAsync(default)).Outline);

        await document.EditAsync(new StudioNativeEditRequest(Outline: [
            new StudioNativeOutlineEdit("add", Title: "Introduction", Page: 1),
            new StudioNativeOutlineEdit("add", Title: "Results", Page: 2)
        ]), default);
        StudioNativeStructure structure = await document.GetStructureAsync(default);
        Assert.Equal(new[] { "Introduction", "Results" }, structure.Outline.Select(item => item.Title));
        string results = structure.Outline[1].Id!;

        await document.EditAsync(new StudioNativeEditRequest(Outline: [
            new StudioNativeOutlineEdit("add", Title: "Details", Page: 3, ParentId: results),
            new StudioNativeOutlineEdit("rename", Id: results, Title: "Findings")
        ]), default);
        structure = await document.GetStructureAsync(default);
        StudioNativeOutlineItem findings = Assert.Single(structure.Outline, item => item.Title == "Findings");
        Assert.Equal((3, "Details"), (findings.Children.Single().Page!.Value, findings.Children.Single().Title));

        await document.EditAsync(new StudioNativeEditRequest(Outline: [new StudioNativeOutlineEdit("remove", Id: structure.Outline[0].Id)]), default);
        Assert.Equal(new[] { "Findings" }, (await document.GetStructureAsync(default)).Outline.Select(item => item.Title));

        // Each request was one step: undoing twice returns to the two top-level bookmarks without children.
        await document.UndoAsync(default);
        await document.UndoAsync(default);
        structure = await document.GetStructureAsync(default);
        Assert.Equal(new[] { "Introduction", "Results" }, structure.Outline.Select(item => item.Title));
        Assert.All(structure.Outline, item => Assert.Empty(item.Children));
    }

    [Fact]
    public async Task Bookmarks_move_down_into_the_one_before_and_back_out() {
        using StudioNativeDocument document = await StudioNativeDocument.OpenAsync(CreateTextPdf("order.pdf", "One", "Two", "Three"), null, default);
        await document.EditAsync(new StudioNativeEditRequest(Outline: [
            new StudioNativeOutlineEdit("add", Title: "A", Page: 1), new StudioNativeOutlineEdit("add", Title: "B", Page: 2),
            new StudioNativeOutlineEdit("add", Title: "C", Page: 3)
        ]), default);
        StudioNativeStructure structure = await document.GetStructureAsync(default);
        string Id(string title) => Find(structure.Outline, title)!.Id!;

        // Down past the next sibling: removed first, then inserted at the next position.
        await document.EditAsync(new StudioNativeEditRequest(Outline: [new StudioNativeOutlineEdit("move", Id: Id("A"), Index: 1)]), default);
        structure = await document.GetStructureAsync(default);
        Assert.Equal(new[] { "B", "A", "C" }, structure.Outline.Select(item => item.Title));

        // Indent C under A, the bookmark before it, then outdent it to follow A again.
        await document.EditAsync(new StudioNativeEditRequest(Outline: [new StudioNativeOutlineEdit("move", Id: Id("C"), ParentId: Id("A"), Index: -1)]), default);
        structure = await document.GetStructureAsync(default);
        Assert.Equal(new[] { "B", "A" }, structure.Outline.Select(item => item.Title));
        Assert.Equal("C", Assert.Single(Find(structure.Outline, "A")!.Children).Title);
        await document.EditAsync(new StudioNativeEditRequest(Outline: [new StudioNativeOutlineEdit("move", Id: Id("C"), Index: 2)]), default);
        structure = await document.GetStructureAsync(default);
        Assert.Equal(new[] { "B", "A", "C" }, structure.Outline.Select(item => item.Title));

        await Assert.ThrowsAnyAsync<Exception>(() =>
            document.EditAsync(new StudioNativeEditRequest(Outline: [new StudioNativeOutlineEdit("move", Id: Id("A"), ParentId: Id("A"))]), default));

        static StudioNativeOutlineItem? Find(IReadOnlyList<StudioNativeOutlineItem> items, string title) =>
            items.FirstOrDefault(item => item.Title == title) ?? items.Select(item => Find(item.Children, title)).FirstOrDefault(found => found is not null);
    }

    [Fact]
    public async Task Bookmark_changes_outside_the_document_are_refused() {
        using StudioNativeDocument document = await StudioNativeDocument.OpenAsync(CreateTextPdf("refuse.pdf", "One"), null, default);
        await Assert.ThrowsAnyAsync<ArgumentException>(() =>
            document.EditAsync(new StudioNativeEditRequest(Outline: [new StudioNativeOutlineEdit("add", Title: "Gone", Page: 4)]), default));
        await Assert.ThrowsAnyAsync<ArgumentException>(() =>
            document.EditAsync(new StudioNativeEditRequest(Outline: [new StudioNativeOutlineEdit("rename", Id: "bookmark-9", Title: "x")]), default));
        await Assert.ThrowsAnyAsync<ArgumentException>(() =>
            document.EditAsync(new StudioNativeEditRequest(Properties: new StudioNativeProperties("a", null, null, null),
                Outline: [new StudioNativeOutlineEdit("rebuild")]), default));
        InvalidOperationException noHeadings = await Assert.ThrowsAsync<InvalidOperationException>(() =>
            document.EditAsync(new StudioNativeEditRequest(Outline: [new StudioNativeOutlineEdit("rebuild")]), default));
        Assert.Contains("no headings", noHeadings.Message);
        Assert.False((await document.GetInfoAsync(default)).IsDirty);
    }

    [Fact]
    public async Task Cropping_and_resizing_change_the_displayed_page_size() {
        using StudioNativeDocument document = await StudioNativeDocument.OpenAsync(CreateTextPdf("crop.pdf", "One", "Two"), null, default);

        await document.EditAsync(new StudioNativeEditRequest(Crop: new StudioNativeCropEdit([1], Left: 20, Top: 10, Right: 30, Bottom: 40)), default);
        StudioNativePageInfo[] pages = (await document.GetInfoAsync(default)).Pages.ToArray();
        Assert.Equal((350D, 250D), (Math.Round(pages[0].Width), Math.Round(pages[0].Height)));
        Assert.Equal((400D, 300D), (pages[1].Width, pages[1].Height));

        await document.EditAsync(new StudioNativeEditRequest(PageSize: new StudioNativePageSizeEdit([2], 595, 842)), default);
        pages = (await document.GetInfoAsync(default)).Pages.ToArray();
        Assert.Equal((595D, 842D), (Math.Round(pages[1].Width), Math.Round(pages[1].Height)));

        await Assert.ThrowsAnyAsync<ArgumentException>(() =>
            document.EditAsync(new StudioNativeEditRequest(Crop: new StudioNativeCropEdit([2], 300, 0, 300, 0)), default));
    }

    [Fact]
    public async Task Page_numbers_watermarks_and_Bates_numbers_become_page_text() {
        using StudioNativeDocument document = await StudioNativeDocument.OpenAsync(CreateTextPdf("stamps.pdf", "One", "Two"), null, default);
        Assert.True((await document.GetStructureAsync(default)).CanStampPages);

        await document.EditAsync(new StudioNativeEditRequest(HeaderFooter: new StudioNativeHeaderFooterEdit(FooterCenter: "Page {page} of {pages}")), default);
        await document.EditAsync(new StudioNativeEditRequest(Watermark: new StudioNativeWatermarkEdit("DRAFT", Pages: [2])), default);
        await document.EditAsync(new StudioNativeEditRequest(Bates: new StudioNativeBatesEdit(Prefix: "ACME", Start: 7, Digits: 4)), default);

        string first = await PageTextAsync(document, 1);
        string second = await PageTextAsync(document, 2);
        Assert.Contains("Page 1 of 2", first);
        Assert.Contains("Page 2 of 2", second);
        Assert.DoesNotContain("DRAFT", first);
        Assert.Contains("DRAFT", second);
        Assert.Contains("ACME0007", first);
        Assert.Contains("ACME0008", second);
    }

    [Fact]
    public async Task Omitted_crop_and_resize_page_lists_apply_to_every_page() {
        using StudioNativeDocument document = await StudioNativeDocument.OpenAsync(CreateTextPdf("all-pages.pdf", "One", "Two"), null, default);
        await document.EditAsync(new StudioNativeEditRequest(Crop: new StudioNativeCropEdit(null, 10, 10, 10, 10)), default);
        Assert.All((await document.GetInfoAsync(default)).Pages, page => {
            Assert.Equal(380D, page.Width);
            Assert.Equal(280D, page.Height);
        });
        await document.EditAsync(new StudioNativeEditRequest(PageSize: new StudioNativePageSizeEdit(null, 500, 600)), default);
        Assert.All((await document.GetInfoAsync(default)).Pages, page => {
            Assert.Equal(500D, page.Width);
            Assert.Equal(600D, page.Height);
        });
    }

    [Fact]
    public async Task Tagged_document_resize_capability_matches_the_operation() {
        string path = Path.Combine(_root, "tagged.pdf");
        PdfDocument.Create(document => document.Page(page => page.Size(400, 300)
            .Content(content => content.Column(column => column.Item().Paragraph(paragraph => paragraph.Text("Tagged page"))))),
            new PdfCreateOptions().EnableTaggedPdfCatalogMarkers()).Save(path);
        Assert.True(PdfDocument.Load(path).Inspect().HasTaggedContent);
        using StudioNativeDocument document = await StudioNativeDocument.OpenAsync(path, null, default);
        Assert.True((await document.GetInfoAsync(default)).CanMutatePages);
        StudioNativeStructure structure = await document.GetStructureAsync(default);
        Assert.False(structure.CanResize);
        await Assert.ThrowsAsync<NotSupportedException>(() => document.EditAsync(
            new StudioNativeEditRequest(PageSize: new StudioNativePageSizeEdit([1], 500, 600)), default));
    }

    [Fact]
    public async Task Flattening_draws_the_form_into_its_pages() {
        string path = Path.Combine(_root, "form.pdf");
        File.WriteAllBytes(path, PdfDocument.Create(document => document.Content(content => content.Canvas(canvas => {
            canvas.TextField("Name", "Alex", global::OfficeIMO.OfficeLength.Points(72), global::OfficeIMO.OfficeLength.Points(72), global::OfficeIMO.OfficeLength.Points(200), global::OfficeIMO.OfficeLength.Points(24));
            canvas.CheckBox("Subscribe", true, 72, 120, 16, 16);
        }))).ToBytes());
        using StudioNativeDocument document = await StudioNativeDocument.OpenAsync(path, null, default);
        StudioNativeStructure structure = await document.GetStructureAsync(default);
        Assert.True(structure.HasForms && structure.CanFlattenForms);

        await document.EditAsync(new StudioNativeEditRequest(Flatten: new StudioNativeFlattenEdit()), default);

        Assert.Empty((await document.GetFormFieldsAsync(1, default)).Fields);
        Assert.False((await document.GetStructureAsync(default)).HasForms);
        Assert.Contains("Alex", await PageTextAsync(document, 1));
    }

    [Fact]
    public async Task Attachments_are_added_saved_and_removed() {
        using StudioNativeDocument document = await StudioNativeDocument.OpenAsync(CreateTextPdf("attach.pdf", "One"), null, default);
        string data = Path.Combine(_root, "figures.csv");
        File.WriteAllText(data, "quarter,total\nQ3,42\n");

        await document.EditAsync(new StudioNativeEditRequest(AddAttachment: new StudioNativeAttachmentAdd(data)), default);
        StudioNativeAttachment attachment = Assert.Single((await document.GetStructureAsync(default)).Attachments);
        Assert.Equal("figures.csv", attachment.FileName);

        string saved = Path.Combine(_root, "saved.csv");
        StudioNativeWriteResult written = await document.WriteAsync(new StudioNativeWriteRequest(Attachment: new StudioNativeAttachmentSave("figures.csv", saved)), default);
        Assert.Equal([saved], written.Files);
        Assert.Equal(File.ReadAllBytes(data), File.ReadAllBytes(saved));

        await document.EditAsync(new StudioNativeEditRequest(RemoveAttachment: "figures.csv"), default);
        Assert.Empty((await document.GetStructureAsync(default)).Attachments);
        await Assert.ThrowsAnyAsync<ArgumentException>(() => document.EditAsync(new StudioNativeEditRequest(RemoveAttachment: "figures.csv"), default));
    }

    [Fact]
    public async Task Splitting_writes_parts_without_replacing_files_and_leaves_the_document_unchanged() {
        using StudioNativeDocument document = await StudioNativeDocument.OpenAsync(CreateTextPdf("report.pdf", "1", "2", "3", "4", "5"), null, default);
        string folder = Path.Combine(_root, "parts");

        StudioNativeWriteResult every = await document.WriteAsync(new StudioNativeWriteRequest(Split: new StudioNativeSplit(folder, PagesPerPart: 2)), default);
        Assert.Equal(new[] { "report p1-2.pdf", "report p3-4.pdf", "report p5.pdf" }, every.Files.Select(Path.GetFileName));
        var counts = new List<int>();
        foreach (string file in every.Files) {
            using StudioNativeDocument part = await StudioNativeDocument.OpenAsync(file, null, default);
            counts.Add((await part.GetInfoAsync(default)).Pages.Count);
        }
        Assert.Equal(new[] { 2, 2, 1 }, counts);

        StudioNativeWriteResult named = await document.WriteAsync(new StudioNativeWriteRequest(Split: new StudioNativeSplit(folder, Parts: [
            new StudioNativeSplitPart([1, 2, 3], "Chapter: one"), new StudioNativeSplitPart([4, 5], "report p1-2")
        ])), default);
        Assert.Equal(new[] { "Chapter- one.pdf", "report p1-2 2.pdf" }, named.Files.Select(Path.GetFileName));
        Assert.False((await document.GetInfoAsync(default)).IsDirty);

        await Assert.ThrowsAnyAsync<ArgumentException>(() =>
            document.WriteAsync(new StudioNativeWriteRequest(Split: new StudioNativeSplit(folder, PagesPerPart: 5)), default));
    }

    [Fact]
    public async Task The_text_exports_as_text_and_Markdown() {
        using StudioNativeDocument document = await StudioNativeDocument.OpenAsync(CreateTextPdf("export.pdf", "Quarterly totals"), null, default);
        string text = Path.Combine(_root, "export.txt");
        string markdown = Path.Combine(_root, "export.md");

        await document.WriteAsync(new StudioNativeWriteRequest(Export: new StudioNativeExport("text", text)), default);
        await document.WriteAsync(new StudioNativeWriteRequest(Export: new StudioNativeExport("markdown", markdown)), default);

        Assert.Contains("Quarterly totals", File.ReadAllText(text));
        Assert.Contains("Quarterly totals", File.ReadAllText(markdown));
        await Assert.ThrowsAnyAsync<ArgumentException>(() =>
            document.WriteAsync(new StudioNativeWriteRequest(Export: new StudioNativeExport("docx", text)), default));
    }

    [Fact]
    public async Task The_document_converts_to_Word_Excel_PowerPoint_and_HTML_without_changing() {
        using StudioNativeDocument document = await StudioNativeDocument.OpenAsync(CreateTextPdf("convert.pdf", "Quarterly totals"), null, default);
        long revision = (await document.GetInfoAsync(default)).Revision;
        foreach ((string format, string extension, string part) in new[] {
                     ("word", "docx", "word/document.xml"), ("excel", "xlsx", "xl/workbook.xml"), ("powerpoint", "pptx", "ppt/presentation.xml") }) {
            string path = Path.Combine(_root, "convert." + extension);
            StudioNativeWriteResult result = await document.WriteAsync(new StudioNativeWriteRequest(Export: new StudioNativeExport(format, path)), default);
            Assert.Equal([path], result.Files);
            using System.IO.Compression.ZipArchive package = System.IO.Compression.ZipFile.OpenRead(path);
            Assert.NotNull(package.GetEntry(part));
        }
        string html = Path.Combine(_root, "convert.html");
        await document.WriteAsync(new StudioNativeWriteRequest(Export: new StudioNativeExport("html", html)), default);
        Assert.Contains("Quarterly", File.ReadAllText(html));
        Assert.Equal(revision, (await document.GetInfoAsync(default)).Revision);
    }

    [Fact]
    public async Task Embedded_images_are_saved_into_a_folder() {
        // A one-pixel red PNG.
        byte[] png = Convert.FromBase64String("iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAIAAACQd1PeAAAADElEQVR4nGP4z8AAAAMBAQDJ/pLvAAAAAElFTkSuQmCC");
        string path = Path.Combine(_root, "images.pdf");
        File.WriteAllBytes(path, PdfDocument.Create(document => document.Content(content => content.Image(png, 24, 24))).ToBytes());
        using StudioNativeDocument document = await StudioNativeDocument.OpenAsync(path, null, default);
        string folder = Path.Combine(_root, "images");

        StudioNativeWriteResult result = await document.WriteAsync(new StudioNativeWriteRequest(Export: new StudioNativeExport("images", folder)), default);

        string image = Assert.Single(result.Files);
        Assert.Equal(folder, Path.GetDirectoryName(image));
        Assert.True(new FileInfo(image).Length > 0);
    }

    [Fact]
    public async Task Flattening_markup_draws_it_into_the_page_and_keeps_links() {
        using StudioNativeDocument document = await StudioNativeDocument.OpenAsync(CreateTextPdf("markup.pdf", "Reviewed"), null, default);
        Assert.False((await document.GetStructureAsync(default)).CanFlattenMarkup);
        await Assert.ThrowsAnyAsync<ArgumentException>(() =>
            document.EditAsync(new StudioNativeEditRequest(Flatten: new StudioNativeFlattenEdit(Markup: true)), default));
        await document.AddAnnotationAsync(new StudioNativeAnnotationRequest("rectangle", 1, new StudioNativeRectangle(40, 60, 100, 50), StrokeWidth: 2), default);
        await document.AddAnnotationAsync(new StudioNativeAnnotationRequest("link", 1, new StudioNativeRectangle(40, 140, 100, 20), LinkUri: "https://officeimo.com/"), default);
        Assert.True((await document.GetStructureAsync(default)).CanFlattenMarkup);

        await document.EditAsync(new StudioNativeEditRequest(Flatten: new StudioNativeFlattenEdit(Markup: true)), default);

        StudioNativeAnnotation left = Assert.Single((await document.GetAnnotationsAsync(1, default)).Annotations);
        Assert.Equal("Link", left.Subtype);
        Assert.False((await document.GetStructureAsync(default)).CanFlattenMarkup);
        await document.UndoAsync(default);
        Assert.Equal(2, (await document.GetAnnotationsAsync(1, default)).Annotations.Count);
    }

    [Fact]
    public unsafe void The_C_interface_reads_structure_applies_edits_and_refuses_two_edits_at_once() {
        string path = CreateTextPdf("abi.pdf", "One", "Two");
        delegate* unmanaged<byte*, byte*, nint, nint*, byte**, int> open = &StudioNativeExports.DocumentOpen;
        delegate* unmanaged<nint, nint, byte**, byte**, int> structure = &StudioNativeExports.DocumentStructureJson;
        delegate* unmanaged<nint, byte*, nint, byte**, int> edit = &StudioNativeExports.DocumentEditJson;
        delegate* unmanaged<nint, void> close = &StudioNativeExports.DocumentClose;
        delegate* unmanaged<void*, void> free = &StudioNativeExports.Free;

        nint document = 0; byte* error = null; byte* json = null;
        fixed (byte* location = Encoding.UTF8.GetBytes(path + "\0")) Assert.Equal(StudioNativeAbi.Ok, open(location, null, 0, &document, &error));
        try {
            string request = JsonSerializer.Serialize(new { outline = new[] { new { op = "add", title = "Start", page = 2 } } });
            fixed (byte* body = Encoding.UTF8.GetBytes(request + "\0")) Assert.Equal(StudioNativeAbi.Ok, edit(document, body, 0, &error));

            request = JsonSerializer.Serialize(new { properties = new { title = "x" }, removeAttachment = "a.txt" });
            fixed (byte* body = Encoding.UTF8.GetBytes(request + "\0")) Assert.Equal(StudioNativeAbi.InvalidArgument, edit(document, body, 0, &error));
            free(error);

            Assert.Equal(StudioNativeAbi.Ok, structure(document, 0, &json, &error));
            using JsonDocument parsed = JsonDocument.Parse(Marshal.PtrToStringUTF8((nint)json)!);
            free(json);
            JsonElement entry = parsed.RootElement.GetProperty("outline")[0];
            Assert.Equal(("Start", 2), (entry.GetProperty("title").GetString(), entry.GetProperty("page").GetInt32()));
            Assert.True(parsed.RootElement.GetProperty("canEditOutline").GetBoolean());
        } finally { close(document); }
    }
}
