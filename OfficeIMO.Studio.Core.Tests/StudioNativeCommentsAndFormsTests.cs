using OfficeIMO.Pdf;
using OfficeIMO.Studio.Native;

namespace OfficeIMO.Studio.Core.Tests;

/// <summary>Comment threads and form authoring through the engine library the Apple apps call.</summary>
public sealed class StudioNativeCommentsAndFormsTests : IDisposable {
    private readonly string _root = Path.Combine(Path.GetTempPath(), "officeimo-native-comments-" + Guid.NewGuid().ToString("N"));

    public StudioNativeCommentsAndFormsTests() {
        Directory.CreateDirectory(_root);
        StudioNativeEngine.Initialize(Path.Combine(_root, "data"));
    }

    public void Dispose() => Directory.Delete(_root, recursive: true);

    private string CreateTextPdf(string name, params string[] pages) {
        string path = Path.Combine(_root, name);
        PdfDocument.Create(document => {
            foreach (string text in pages)
                document.Page(page => page.Size(global::OfficeIMO.OfficeLength.Points(400), global::OfficeIMO.OfficeLength.Points(300)).Content(content => content.Column(column => column.Item().AddParagraph(p => p.AddText(text)))));
        }).Save(path);
        return path;
    }

    [Fact]
    public async Task Replies_nest_under_their_comment_and_resolving_is_undoable_while_the_thread_keeps_its_id() {
        using StudioNativeDocument document = await StudioNativeDocument.OpenAsync(CreateTextPdf("comments.pdf", "One", "Two"), null, default);
        await document.AddAnnotationAsync(new StudioNativeAnnotationRequest("note", 2, new StudioNativeRectangle(40, 40, 20, 20),
            Text: "Check the totals", Author: "Alex Morgan"), default);

        StudioNativeComments comments = await document.GetCommentsAsync(default);
        StudioNativeCommentThread thread = Assert.Single(comments.Threads);
        Assert.True(comments.CanReply);
        Assert.Equal((2, "none", false), (thread.Page, thread.State, thread.Resolved));
        Assert.Equal(("Alex Morgan", "Check the totals", 0), (thread.Entries[0].Author, thread.Entries[0].Contents, thread.Entries[0].Depth));
        Assert.NotNull(thread.Bounds);
        string id = thread.Id;

        await document.EditAsync(new StudioNativeEditRequest(Reply: new StudioNativeCommentReply(thread.ObjectNumber, "Fixed in the second draft", "Sam Lee")), default);
        thread = Assert.Single((await document.GetCommentsAsync(default)).Threads);
        Assert.Equal(new (string?, int)[] { ("Alex Morgan", 0), ("Sam Lee", 1) }, thread.Entries.Select(entry => (entry.Author, entry.Depth)));
        Assert.Equal("Fixed in the second draft", thread.Entries[1].Contents);
        Assert.Equal(id, thread.Id);

        await document.EditAsync(new StudioNativeEditRequest(ReviewState: new StudioNativeCommentState(thread.ObjectNumber, "completed")), default);
        thread = Assert.Single((await document.GetCommentsAsync(default)).Threads);
        Assert.Equal((id, "completed", true), (thread.Id, thread.State, thread.Resolved));

        await document.UndoAsync(default);
        thread = Assert.Single((await document.GetCommentsAsync(default)).Threads);
        Assert.False(thread.Resolved);
        Assert.Equal(2, thread.Entries.Count);
        Assert.Equal(id, thread.Id);
    }

    [Fact]
    public async Task A_comment_resolves_and_reopens_on_a_plain_pdf_and_on_one_saved_incrementally() {
        string plain = CreateTextPdf("plain-resolve.pdf", "One");
        // An incremental revision makes annotation changes go into an appended revision, as a tagged PDF's do.
        string incremental = CreateTextPdf("incremental-resolve.pdf", "One");
        File.WriteAllBytes(incremental, PdfDocument.Load(incremental).AppendMetadataRevision(title: "Draft").ToBytes());

        foreach (string path in new[] { plain, incremental }) {
            using StudioNativeDocument document = await StudioNativeDocument.OpenAsync(path, null, default);
            await document.AddAnnotationAsync(new StudioNativeAnnotationRequest("note", 1, new StudioNativeRectangle(40, 40, 20, 20),
                Text: "Check this", Author: "Alex"), default);
            int number = Assert.Single((await document.GetCommentsAsync(default)).Threads).ObjectNumber;

            await document.EditAsync(new StudioNativeEditRequest(ReviewState: new StudioNativeCommentState(number, "completed")), default);
            StudioNativeCommentThread thread = Assert.Single((await document.GetCommentsAsync(default)).Threads);
            Assert.True(thread.Resolved, path);

            await document.EditAsync(new StudioNativeEditRequest(ReviewState: new StudioNativeCommentState(thread.ObjectNumber, "none")), default);
            Assert.False(Assert.Single((await document.GetCommentsAsync(default)).Threads).Resolved, path);
        }
    }

    [Fact]
    public async Task Replies_and_states_are_refused_for_missing_comments_empty_text_and_unknown_states() {
        using StudioNativeDocument document = await StudioNativeDocument.OpenAsync(CreateTextPdf("refused.pdf", "One"), null, default);
        await document.AddAnnotationAsync(new StudioNativeAnnotationRequest("note", 1, new StudioNativeRectangle(40, 40, 20, 20), Text: "Note"), default);
        int number = Assert.Single((await document.GetCommentsAsync(default)).Threads).ObjectNumber;

        await Assert.ThrowsAnyAsync<ArgumentException>(() => document.EditAsync(new StudioNativeEditRequest(Reply: new StudioNativeCommentReply(number + 1000, "Hello")), default));
        await Assert.ThrowsAnyAsync<ArgumentException>(() => document.EditAsync(new StudioNativeEditRequest(Reply: new StudioNativeCommentReply(number, "  ")), default));
        await Assert.ThrowsAnyAsync<ArgumentException>(() => document.EditAsync(new StudioNativeEditRequest(ReviewState: new StudioNativeCommentState(number, "done")), default));
        Assert.Single(Assert.Single((await document.GetCommentsAsync(default)).Threads).Entries);
    }

    [Fact]
    public async Task Fields_are_added_where_they_are_drawn_then_renamed_moved_and_removed_one_undoable_step_each() {
        using StudioNativeDocument document = await StudioNativeDocument.OpenAsync(CreateTextPdf("form.pdf", "Application"), null, default);
        Assert.True((await document.GetInfoAsync(default)).CanAuthorForms);

        await document.EditAsync(new StudioNativeEditRequest(AddFormField: new StudioNativeFormFieldAdd("text", "FullName", 1,
            new StudioNativeRectangle(40, 60, 160, 24), Label: "Full name")), default);
        await document.EditAsync(new StudioNativeEditRequest(AddFormField: new StudioNativeFormFieldAdd("checkbox", "Agree", 1,
            new StudioNativeRectangle(40, 100, 16, 16))), default);
        await document.EditAsync(new StudioNativeEditRequest(AddFormField: new StudioNativeFormFieldAdd("choice", "Country", 1,
            new StudioNativeRectangle(40, 130, 160, 24), ["Poland", "Germany", " ", "Poland"])), default);
        await document.EditAsync(new StudioNativeEditRequest(AddFormField: new StudioNativeFormFieldAdd("signature", "Signature", 1,
            new StudioNativeRectangle(40, 200, 180, 40))), default);

        StudioNativeFormFields fields = await document.GetFormFieldsAsync(1, default);
        Assert.Equal(new[] { ("FullName", "text"), ("Agree", "checkbox"), ("Country", "combo"), ("Signature", "signature") },
            fields.Fields.Select(field => (field.Name, field.Kind)));
        StudioNativeRectangle box = fields.Fields[0].Widgets.Single().Bounds;
        Assert.Equal((40D, 60D, 160D, 24D), (Math.Round(box.X), Math.Round(box.Y), Math.Round(box.Width), Math.Round(box.Height)));
        Assert.Equal("Full name", fields.Fields[0].Label);
        Assert.Equal(new[] { "Poland", "Germany" }, fields.Fields[2].Options.Select(option => option.Value));

        await document.EditAsync(new StudioNativeEditRequest(UpdateFormField: new StudioNativeFormFieldUpdate("FullName", NewName: "Name", Required: true)), default);
        await document.EditAsync(new StudioNativeEditRequest(UpdateFormField: new StudioNativeFormFieldUpdate("Name", Bounds: new StudioNativeRectangle(200, 60, 120, 30))), default);
        StudioNativeFormField renamed = Assert.Single((await document.GetFormFieldsAsync(1, default)).Fields, field => field.Name == "Name");
        Assert.True(renamed.Required);
        box = renamed.Widgets.Single().Bounds;
        Assert.Equal((200D, 60D, 120D, 30D), (Math.Round(box.X), Math.Round(box.Y), Math.Round(box.Width), Math.Round(box.Height)));

        await document.EditAsync(new StudioNativeEditRequest(RemoveFormField: "Agree"), default);
        Assert.DoesNotContain((await document.GetFormFieldsAsync(1, default)).Fields, field => field.Name == "Agree");
        await document.UndoAsync(default);
        Assert.Contains((await document.GetFormFieldsAsync(1, default)).Fields, field => field.Name == "Agree");
    }

    [Fact]
    public async Task Radio_groups_stack_a_button_per_option_and_push_buttons_show_their_caption() {
        using StudioNativeDocument document = await StudioNativeDocument.OpenAsync(CreateTextPdf("radio.pdf", "Survey"), null, default);
        await document.EditAsync(new StudioNativeEditRequest(AddFormField: new StudioNativeFormFieldAdd("radio", "Size", 1,
            new StudioNativeRectangle(40, 40, 80, 10), ["Small", "Medium", "Large"])), default);
        await document.EditAsync(new StudioNativeEditRequest(AddFormField: new StudioNativeFormFieldAdd("button", "Submit", 1,
            new StudioNativeRectangle(200, 40, 90, 24), Caption: "Send")), default);

        StudioNativeFormFields fields = await document.GetFormFieldsAsync(1, default);
        StudioNativeFormField radio = Assert.Single(fields.Fields, field => field.Name == "Size");
        Assert.Equal("radio", radio.Kind);
        Assert.Equal(3, radio.Widgets.Count);
        // Three 14-point buttons with 6 points between them need 54 points, more than the 10 drawn.
        double top = radio.Widgets.Min(widget => widget.Bounds.Y), bottom = radio.Widgets.Max(widget => widget.Bounds.Y + widget.Bounds.Height);
        Assert.True(bottom - top >= 53, $"The group spans {bottom - top} points.");
        Assert.Equal("button", Assert.Single(fields.Fields, field => field.Name == "Submit").Kind);
        Assert.Equal(3, radio.WidgetCount);
        Assert.Equal(new[] { "Size", "Submit" }, fields.AllNames.Order(StringComparer.Ordinal));

        await Assert.ThrowsAnyAsync<ArgumentException>(() => document.EditAsync(new StudioNativeEditRequest(UpdateFormField:
            new StudioNativeFormFieldUpdate("Size", Bounds: new StudioNativeRectangle(40, 120, 80, 60))), default));
        await Assert.ThrowsAnyAsync<ArgumentException>(() => document.EditAsync(new StudioNativeEditRequest(AddFormField:
            new StudioNativeFormFieldAdd("radio", "Empty", 1, new StudioNativeRectangle(40, 200, 80, 40))), default));
    }

    [Fact]
    public async Task A_radio_group_drawn_near_the_bottom_moves_up_to_stay_on_the_page_and_names_cover_every_page() {
        using StudioNativeDocument document = await StudioNativeDocument.OpenAsync(CreateTextPdf("radio-bottom.pdf", "One", "Two"), null, default);
        await document.EditAsync(new StudioNativeEditRequest(AddFormField: new StudioNativeFormFieldAdd("text", "Remote", 2,
            new StudioNativeRectangle(40, 40, 120, 24))), default);
        // Ten options need 194 points; drawn 20 points above the bottom of a 300-point page, the group moves up.
        await document.EditAsync(new StudioNativeEditRequest(AddFormField: new StudioNativeFormFieldAdd("radio", "Rating", 1,
            new StudioNativeRectangle(40, 270, 80, 10), Enumerable.Range(1, 10).Select(static n => n.ToString(System.Globalization.CultureInfo.InvariantCulture)).ToArray())), default);

        StudioNativeFormFields fields = await document.GetFormFieldsAsync(1, default);
        StudioNativeFormField radio = Assert.Single(fields.Fields);
        Assert.All(radio.Widgets, widget => Assert.InRange(widget.Bounds.Y + widget.Bounds.Height, 0D, 300.5D));
        Assert.Equal(new[] { "Rating", "Remote" }, fields.AllNames.Order(StringComparer.Ordinal));

        await Assert.ThrowsAnyAsync<ArgumentException>(() => document.EditAsync(new StudioNativeEditRequest(AddFormField:
            new StudioNativeFormFieldAdd("radio", "TooMany", 1, new StudioNativeRectangle(40, 40, 80, 10),
                Enumerable.Range(1, 20).Select(static n => "Option " + n.ToString(System.Globalization.CultureInfo.InvariantCulture)).ToArray())), default));
    }

    [Fact]
    public async Task Field_authoring_refuses_duplicate_names_choices_without_options_and_unknown_kinds() {
        using StudioNativeDocument document = await StudioNativeDocument.OpenAsync(CreateTextPdf("refused-form.pdf", "Form"), null, default);
        var box = new StudioNativeRectangle(40, 60, 160, 24);
        await document.EditAsync(new StudioNativeEditRequest(AddFormField: new StudioNativeFormFieldAdd("text", "Name", 1, box)), default);

        await Assert.ThrowsAnyAsync<ArgumentException>(() => document.EditAsync(new StudioNativeEditRequest(AddFormField: new StudioNativeFormFieldAdd("text", "Name", 1, box with { Y = 120 })), default));
        await Assert.ThrowsAnyAsync<ArgumentException>(() => document.EditAsync(new StudioNativeEditRequest(AddFormField: new StudioNativeFormFieldAdd("choice", "Pick", 1, box with { Y = 120 })), default));
        await Assert.ThrowsAnyAsync<ArgumentException>(() => document.EditAsync(new StudioNativeEditRequest(AddFormField: new StudioNativeFormFieldAdd("slider", "Go", 1, box with { Y = 120 })), default));
        await Assert.ThrowsAnyAsync<ArgumentException>(() => document.EditAsync(new StudioNativeEditRequest(AddFormField: new StudioNativeFormFieldAdd("text", "Tiny", 1, box with { Width = 2 })), default));
        await Assert.ThrowsAnyAsync<ArgumentException>(() => document.EditAsync(new StudioNativeEditRequest(UpdateFormField: new StudioNativeFormFieldUpdate("Name", NewName: "Other", Bounds: box)), default));
        await Assert.ThrowsAnyAsync<ArgumentException>(() => document.EditAsync(new StudioNativeEditRequest(RemoveFormField: "Missing"), default));
        Assert.Single((await document.GetFormFieldsAsync(1, default)).Fields);
    }
}
