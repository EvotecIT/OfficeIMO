using OfficeIMO.Pdf;
using OfficeIMO.Studio.Features.Shell;

namespace OfficeIMO.Studio.Tests;

public sealed class StudioPageWorkflowIntakeTests {
    [Theory]
    [InlineData("import")]
    [InlineData("split")]
    public async Task PickerRetainsThePageOptionsThatStartedTheOperation(string operation) {
        string root = Path.Combine(Path.GetTempPath(), "officeimo-page-options-" + Guid.NewGuid().ToString("N"));
        Directory.CreateDirectory(root);
        try {
            string source = Path.Combine(root, "source.pdf");
            string other = Path.Combine(root, "other.pdf");
            string folder = Path.Combine(root, "split");
            PdfDocument.Create(document => {
                document.Page(page => page.Size(250, 350));
                document.Page(page => page.Size(300, 400));
            }).Save(source);
            PdfDocument.Create(document => document.Page(page => page.Size(450, 550))).Save(other);
            MainWindowViewModel? model = null;
            using (model = new MainWindowViewModel(_ => Task.FromResult<string?>(null),
                pickImportPdfs: _ => {
                    model!.ClearPageSelectionCommand.Execute(null);
                    return Task.FromResult<IReadOnlyList<string>>(new[] { other });
                },
                pickOutputFolder: _ => {
                    model!.SplitPagesPerDocument = 1;
                    return Task.FromResult<string?>(folder);
                }, reviewPageSplit: _ => Task.FromResult(true))) {
                await model.OpenDocumentAsync(source);
                model.SelectAllPagesCommand.Execute(null);
                model.SplitPagesPerDocument = 2;
                if (operation == "import") {
                    await model.ImportPagesCommand.ExecuteAsync(null);
                    await model.SaveCommand.ExecuteAsync(null);
                    Assert.Equal(450D, PdfDocument.Load(File.ReadAllBytes(source)).Inspect().Pages[0].Width);
                    Assert.Equal(3, model.Pages.Count);
                } else {
                    await model.SplitCommand.ExecuteAsync(null);
                    string output = Assert.Single(Directory.GetFiles(Path.Combine(folder, "Split PDFs"), "*.pdf"));
                    Assert.Equal(2, PdfDocument.Load(File.ReadAllBytes(output)).Inspect().PageCount);
                }
                Assert.Null(model.ErrorMessage);
            }
        } finally { Directory.Delete(root, true); }
    }

    [Theory]
    [InlineData("import", false)]
    [InlineData("extract", false)]
    [InlineData("split", false)]
    [InlineData("import", true)]
    [InlineData("extract", true)]
    [InlineData("split", true)]
    public async Task PickerResultCannotApplyToAChangedOrReplacedDocument(string operation, bool replace) {
        string root = Path.Combine(Path.GetTempPath(), "officeimo-page-intake-" + Guid.NewGuid().ToString("N"));
        Directory.CreateDirectory(root);
        try {
            string source = Path.Combine(root, "source.pdf");
            string other = Path.Combine(root, "other.pdf");
            string output = Path.Combine(root, "output.pdf");
            string folder = Path.Combine(root, "split");
            PdfDocument.Create(document => document.Page(page => page.Size(250, 350))).Save(source);
            File.Copy(source, other);
            var selected = new TaskCompletionSource(TaskCreationOptions.RunContinuationsAsynchronously);
            var resume = new TaskCompletionSource(TaskCreationOptions.RunContinuationsAsynchronously);
            async Task WaitForPicker() { selected.SetResult(); await resume.Task; }
            using var model = new MainWindowViewModel(_ => Task.FromResult<string?>(null),
                pickSavePdf: async _ => { await WaitForPicker(); return output; },
                pickImportPdfs: async _ => { await WaitForPicker(); return new[] { other }; },
                pickOutputFolder: async _ => { await WaitForPicker(); return folder; });
            await model.OpenDocumentAsync(source);
            model.SelectAllPagesCommand.Execute(null);
            Task pending = operation switch {
                "import" => model.ImportPagesCommand.ExecuteAsync(null),
                "extract" => model.ExtractSelectedCommand.ExecuteAsync(null),
                _ => model.SplitCommand.ExecuteAsync(null)
            };
            await selected.Task.WaitAsync(TimeSpan.FromSeconds(10));
            if (replace) await model.OpenDocumentAsync(other);
            else await model.RotateRightCommand.ExecuteAsync(null);
            resume.SetResult();
            await pending;
            Assert.NotNull(model.ErrorMessage);
            Assert.Single(model.Pages);
            Assert.False(File.Exists(output));
            Assert.False(Directory.Exists(folder));
            Assert.Equal(!replace, model.IsDirty);
        } finally { Directory.Delete(root, true); }
    }
}
