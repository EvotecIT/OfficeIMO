using System.Security.Cryptography;
using System.Text.Json;
using Avalonia;
using OfficeIMO.Pdf;
using OfficeIMO.Studio.Features.Editor;
using OfficeIMO.Studio.Features.Shell;

namespace OfficeIMO.Studio.Tests;

public sealed class RedactionReviewTests {
    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public async Task SelectedSearchMarksProduceVerifiedCopyAndContentFreeReport(bool sanitize) {
        using var files = new TestFiles();
        byte[] source = CreateSource();
        await File.WriteAllBytesAsync(files.Input, source);
        using var model = new MainWindowViewModel(_ => Task.FromResult<string?>(null),
            pickSavePdf: _ => Task.FromResult<string?>(files.Output),
            pickSaveRedactionReport: _ => Task.FromResult<string?>(files.Report));
        await model.OpenDocumentAsync(files.Input);
        model.ShowProtectModeCommand.Execute(null);
        model.RedactionSearchText = "private account";
        await model.SearchRedactionsCommand.ExecuteAsync(null);
        Assert.Equal(2, model.RedactionMarks.Count);
        model.RedactionMarks[0].Reason = "private reason never exported";
        model.RedactionMarks[1].IsIncluded = false;
        model.SanitizeAfterRedaction = sanitize;

        Assert.False(model.CanApplyReviewedRedactions);
        await model.ReviewRedactionsCommand.ExecuteAsync(null);
        Assert.True(model.CanApplyReviewedRedactions, model.ErrorMessage ?? model.PendingRedactionSummary);
        await model.ApplyPendingRedactionCommand.ExecuteAsync(null);
        Assert.False(model.HasError, model.ErrorMessage);
        Assert.Empty(model.RedactionMarks);
        Assert.True(model.LastRedactionSummary?.IsVerified);
        Assert.Equal(sanitize, model.LastRedactionSummary!.SanitizedItemCount.HasValue);

        await model.SaveVerifiedRedactionCopyCommand.ExecuteAsync(null);
        Assert.False(model.HasError, model.ErrorMessage);
        await model.ExportRedactionEvidenceCommand.ExecuteAsync(null);
        Assert.False(model.HasError, model.ErrorMessage);
        byte[] saved = await File.ReadAllBytesAsync(files.Output);
        PdfDocumentReadResult read = PdfDocument.Load(saved).Read();
        Assert.DoesNotContain("private account", string.Join(" ", read.Pages[0].Elements.OfType<PdfLogicalTextBlock>().Select(block => block.Text)), StringComparison.OrdinalIgnoreCase);
        Assert.Contains("private account", read.Text, StringComparison.OrdinalIgnoreCase);
        Assert.Equal(source, await File.ReadAllBytesAsync(files.Input));
        string report = await File.ReadAllTextAsync(files.Report);
        Assert.DoesNotContain("private", report, StringComparison.OrdinalIgnoreCase);
        using JsonDocument json = JsonDocument.Parse(report);
        Assert.Equal(Convert.ToHexString(SHA256.HashData(saved)), json.RootElement.GetProperty("OutputSha256").GetString(), ignoreCase: true);
        Assert.Equal(1, json.RootElement.GetProperty("AreaCount").GetInt32());
    }

    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public async Task EvidenceExportRejectsChangedCopyOrReplacingTheCopy(bool overwriteCopy) {
        using var files = new TestFiles();
        await File.WriteAllBytesAsync(files.Input, CreateSource());
        using var model = new MainWindowViewModel(_ => Task.FromResult<string?>(null),
            pickSavePdf: _ => Task.FromResult<string?>(files.Output),
            pickSaveRedactionReport: _ => Task.FromResult<string?>(overwriteCopy ? files.Output : files.Report));
        await model.OpenDocumentAsync(files.Input);
        model.RedactionSearchText = "private account";
        await model.SearchRedactionsCommand.ExecuteAsync(null);
        await model.ReviewRedactionsCommand.ExecuteAsync(null);
        await model.ApplyPendingRedactionCommand.ExecuteAsync(null);
        await model.SaveVerifiedRedactionCopyCommand.ExecuteAsync(null);
        Assert.False(model.HasError, model.ErrorMessage);
        if (!overwriteCopy) await File.WriteAllBytesAsync(files.Output, CreateSource());
        byte[] beforeExport = await File.ReadAllBytesAsync(files.Output);
        await model.ExportRedactionEvidenceCommand.ExecuteAsync(null);
        Assert.True(model.HasError);
        Assert.False(File.Exists(files.Report));
        Assert.Equal(beforeExport, await File.ReadAllBytesAsync(files.Output));
    }

    [Fact]
    public async Task ChangingSelectionOrReasonRequiresAnotherReviewAndMutationClearsMarks() {
        using var files = new TestFiles();
        await File.WriteAllBytesAsync(files.Input, CreateSource());
        using var model = new MainWindowViewModel(_ => Task.FromResult<string?>(null));
        await model.OpenDocumentAsync(files.Input);
        model.RedactionSearchText = "private account";
        await model.SearchRedactionsCommand.ExecuteAsync(null);
        await model.ReviewRedactionsCommand.ExecuteAsync(null);
        Assert.True(model.CanApplyReviewedRedactions);
        model.RedactionMarks[0].IsIncluded = false;
        Assert.False(model.CanApplyReviewedRedactions);
        await model.ApplyPendingRedactionCommand.ExecuteAsync(null);
        Assert.Null(model.LastRedactionSummary);
        await model.ReviewRedactionsCommand.ExecuteAsync(null);
        Assert.True(model.CanApplyReviewedRedactions);
        model.RedactionMarks[1].Reason = "Updated review reason";
        Assert.False(model.CanApplyReviewedRedactions);
        model.SetOrganizerSelection([model.OrganizerPages[0]]);
        await model.DuplicateSelectedCommand.ExecuteAsync(null);
        Assert.Empty(model.RedactionMarks);
        Assert.False(model.CanApplyReviewedRedactions);
        Assert.All(model.Pages, page => Assert.Empty(page.PendingRedactionAreas));
    }

    [Fact]
    public async Task DrawnMarksAccumulateAcrossPagesAndRemainVisibleAfterToolChange() {
        using var files = new TestFiles();
        await File.WriteAllBytesAsync(files.Input, CreateSource());
        using var model = new MainWindowViewModel(_ => Task.FromResult<string?>(null));
        await model.OpenDocumentAsync(files.Input);
        model.ShowProtectModeCommand.Execute(null);
        model.BeginRedactionCommand.Execute(null);
        model.Pages[0].CompleteEditorGesture(new PdfEditorGesture(1, 30, 40, 100, 70, []));
        await WaitUntilAsync(() => model.RedactionMarks.Count == 1);
        model.Pages[1].CompleteEditorGesture(new PdfEditorGesture(2, 40, 80, 130, 120, []));
        await WaitUntilAsync(() => model.RedactionMarks.Count == 2);
        Assert.Equal(new Rect(30, 40, 70, 30), Assert.Single(model.Pages[0].PendingRedactionAreas));
        Assert.Equal(new Rect(40, 80, 90, 40), Assert.Single(model.Pages[1].PendingRedactionAreas));
        model.ShowViewModeCommand.Execute(null);
        Assert.Equal(2, model.RedactionMarks.Count);
        model.SelectedRedactionMark = model.RedactionMarks[0];
        model.RemoveRedactionMarkCommand.Execute(null);
        Assert.Empty(model.Pages[0].PendingRedactionAreas);
        Assert.Single(model.Pages[1].PendingRedactionAreas);
    }

    [Fact]
    public async Task SelectedPageSearchAndInvalidPatternDoNotReplaceExistingMarks() {
        using var files = new TestFiles();
        await File.WriteAllBytesAsync(files.Input, CreateSource());
        using var model = new MainWindowViewModel(_ => Task.FromResult<string?>(null));
        await model.OpenDocumentAsync(files.Input);
        model.SetOrganizerSelection([model.OrganizerPages[1]]);
        model.RedactionSearchSelectedPagesOnly = true;
        model.RedactionSearchText = "PRIVATE ACCOUNT";
        model.RedactionSearchRegex = true;
        await model.SearchRedactionsCommand.ExecuteAsync(null);
        Assert.Equal(2, Assert.Single(model.RedactionMarks).PageNumber);
        model.RedactionSearchText = "[";
        await model.SearchRedactionsCommand.ExecuteAsync(null);
        Assert.True(model.HasError);
        Assert.Single(model.RedactionMarks);
    }

    private static byte[] CreateSource() => PdfDocument.Create(compose => {
        compose.Page(page => page.Size(600, 800).Content(content => content.Item(item =>
            item.Paragraph(paragraph => paragraph.Text("First private account 123")))));
        compose.Page(page => page.Size(600, 800).Content(content => content.Item(item =>
            item.Paragraph(paragraph => paragraph.Text("Second private account 456")))));
    }).ToBytes();

    private static async Task WaitUntilAsync(Func<bool> predicate) {
        for (int attempt = 0; attempt < 200 && !predicate(); attempt++) await Task.Delay(20);
        Assert.True(predicate());
    }

    private sealed class TestFiles : IDisposable {
        private readonly string _root = Path.Combine(Path.GetTempPath(), "officeimo-redaction-review-" + Guid.NewGuid().ToString("N"));
        internal TestFiles() => Directory.CreateDirectory(_root);
        internal string Input => Path.Combine(_root, "source.pdf");
        internal string Output => Path.Combine(_root, "redacted.pdf");
        internal string Report => Path.Combine(_root, "evidence.json");
        public void Dispose() => Directory.Delete(_root, recursive: true);
    }
}
