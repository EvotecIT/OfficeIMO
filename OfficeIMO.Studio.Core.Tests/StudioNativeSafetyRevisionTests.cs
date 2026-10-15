using OfficeIMO.Pdf;
using OfficeIMO.Studio.Native;
using OfficeIMO.Studio.Features.Workspace;

namespace OfficeIMO.Studio.Core.Tests;

public sealed class StudioNativeSafetyRevisionTests : IDisposable {
    private readonly string _root = Path.Combine(Path.GetTempPath(), "officeimo-native-safety-" + Guid.NewGuid().ToString("N"));

    public StudioNativeSafetyRevisionTests() {
        Directory.CreateDirectory(_root);
        StudioNativeEngine.Initialize(Path.Combine(_root, "data"));
    }

    public void Dispose() => Directory.Delete(_root, true);

    private string Source() {
        string path = Path.Combine(_root, "source.pdf");
        PdfDocument.Create(document => {
            foreach (string text in new[] { "First secret account", "Second secret account" })
                document.Page(page => page.Size(400, 300).Content(content => content.Paragraph(p => p.Text(text))));
        }).Save(path);
        return path;
    }

    [Theory]
    [InlineData("reorder", false)]
    [InlineData("reorder", true)]
    [InlineData("rotate", false)]
    [InlineData("rotate", true)]
    [InlineData("undo", false)]
    [InlineData("undo", true)]
    [InlineData("redo", false)]
    [InlineData("redo", true)]
    public async Task Redaction_rejects_stale_index_and_area_selections_without_mutation(string transition, bool rectangle) {
        using var document = await StudioNativeDocument.OpenAsync(Source(), null, default);
        if (transition is "undo" or "redo") await document.RotatePagesAsync([1], 90, default);
        if (transition == "redo") await document.UndoAsync(default);
        var found = await document.RedactAsync(new(Find: new("secret")), default);
        Assert.Equal(2, found.Marks!.Count);
        switch (transition) {
            case "reorder": await document.ReorderPagesAsync([2, 1], default); break;
            case "rotate": await document.RotatePagesAsync([1], 90, default); break;
            case "undo": await document.UndoAsync(default); break;
            case "redo": await document.RedoAsync(default); break;
        }
        long revision = (await document.GetInfoAsync(default)).Revision;
        Assert.NotEqual(found.Revision, revision);
        string before = Path.Combine(_root, "before.pdf");
        await document.SaveCopyAsync(before, default);
        var apply = rectangle
            ? new StudioNativeRedactionApply(Areas: [found.Marks[0]], Revision: found.Revision)
            : new StudioNativeRedactionApply(Find: new("secret"), Include: [0], Revision: found.Revision);
        await Assert.ThrowsAsync<InvalidOperationException>(() => document.RedactAsync(new(Apply: apply), default));
        Assert.Equal(revision, (await document.GetInfoAsync(default)).Revision);
        string after = Path.Combine(_root, "after.pdf");
        await document.SaveCopyAsync(after, default);
        Assert.Equal(File.ReadAllBytes(before), File.ReadAllBytes(after));
    }

    [Fact]
    public async Task Redaction_requires_a_revision_for_prepared_selections() {
        using var document = await StudioNativeDocument.OpenAsync(Source(), null, default);
        long revision = (await document.GetInfoAsync(default)).Revision;
        await Assert.ThrowsAsync<ArgumentException>(() => document.RedactAsync(new(Apply:
            new(Find: new("secret"), Include: [0])), default));
        Assert.Equal(revision, (await document.GetInfoAsync(default)).Revision);
    }

    [Theory]
    [InlineData("missing", "copy")]
    [InlineData("empty", "copy")]
    [InlineData("equal", "copy")]
    [InlineData("normalized", "copy")]
    [InlineData("truncated", "copy")]
    [InlineData("missing", "print")]
    [InlineData("missing", "change")]
    public async Task Restricted_copy_requires_effectively_distinct_owner_before_publication(string credential, string restriction) {
        string userPassword = credential switch { "normalized" => "A", "truncated" => new string('x', 127) + "reader", _ => "reader" };
        string? ownerPassword = credential switch {
            "missing" => null, "empty" => "", "normalized" => "Ａ", "truncated" => new string('x', 127) + "owner", _ => userPassword
        };
        string source = Source();
        byte[] original = File.ReadAllBytes(source);
        using var document = await StudioNativeDocument.OpenAsync(source, null, default);
        long revision = (await document.GetInfoAsync(default)).Revision;
        string destination = Path.Combine(_root, "restricted.pdf");
        byte[] existing = [1, 2, 3, 4];
        File.WriteAllBytes(destination, existing);
        await Assert.ThrowsAsync<ArgumentException>(() => document.SaveProtectedCopyAsync(
            new(destination, userPassword, ownerPassword, null, AllowPrinting: restriction != "print",
                AllowCopying: restriction != "copy", AllowChanges: restriction != "change"), default));
        Assert.Equal(existing, File.ReadAllBytes(destination));
        Assert.Equal(original, File.ReadAllBytes(source));
        Assert.Equal(revision, (await document.GetInfoAsync(default)).Revision);
    }

    [Theory]
    [InlineData(PdfStandardEncryptionAlgorithm.Aes128)]
    [InlineData(PdfStandardEncryptionAlgorithm.LegacyRc4)]
    public async Task Workspace_rejects_owner_credentials_equal_after_legacy_truncation(PdfStandardEncryptionAlgorithm algorithm) {
        using var workspace = await PdfWorkspace.OpenAsync(Source(), default);
        string destination = Path.Combine(_root, "legacy.pdf");
        var encryption = new PdfStandardEncryptionOptions(new string('x', 32) + "reader") {
            OwnerPassword = new string('x', 32) + "owner", Algorithm = algorithm, AllowedPermissions = PdfStandardPermissions.None
        };
        await Assert.ThrowsAsync<ArgumentException>(() => workspace.SaveProtectedCopyAsync(destination, encryption, null, default));
        Assert.False(File.Exists(destination));
    }

    [Fact]
    public async Task Unrestricted_copy_allows_a_single_password() {
        using var document = await StudioNativeDocument.OpenAsync(Source(), null, default);
        string destination = Path.Combine(_root, "unrestricted.pdf");
        await document.SaveProtectedCopyAsync(new(destination, "reader", null, null), default);
        using var reopened = await StudioNativeDocument.OpenAsync(destination, "reader", default);
        var facts = await reopened.GetInfoAsync(default);
        Assert.True(facts.CanExtractText);
        Assert.True(facts.CanExtractContent);
        Assert.True(facts.CanPrint);
        Assert.True(facts.CanRotatePages);
    }

    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public async Task Protected_copy_copying_switch_controls_text_and_content_for_user_password(bool allowCopying) {
        string source = Source();
        byte[] original = File.ReadAllBytes(source);
        using var document = await StudioNativeDocument.OpenAsync(source, null, default);
        string protectedCopy = Path.Combine(_root, "protected.pdf");
        await document.SaveProtectedCopyAsync(new(protectedCopy, "reader", "owner", null, AllowCopying: allowCopying), default);
        using var user = await StudioNativeDocument.OpenAsync(protectedCopy, "reader", default);
        var facts = await user.GetInfoAsync(default);
        Assert.Equal(allowCopying, facts.CanExtractText);
        Assert.Equal(allowCopying, facts.CanExtractContent);
        Assert.True(facts.CanPrint);
        Assert.True(facts.CanRotatePages);
        string text = Path.Combine(_root, "user.txt");
        if (allowCopying) {
            await user.WriteAsync(new(Export: new("text", text)), default);
            Assert.Contains("secret", File.ReadAllText(text));
        } else {
            await Assert.ThrowsAsync<NotSupportedException>(() => user.WriteAsync(new(Export: new("text", text)), default));
            Assert.False(File.Exists(text));
        }
        using var owner = await StudioNativeDocument.OpenAsync(protectedCopy, "owner", default);
        Assert.True((await owner.GetInfoAsync(default)).CanExtractText);
        await owner.WriteAsync(new(Export: new("text", Path.Combine(_root, "owner.txt"))), default);
        Assert.Contains("secret", File.ReadAllText(Path.Combine(_root, "owner.txt")));
        Assert.Equal(original, File.ReadAllBytes(source));
    }
}
