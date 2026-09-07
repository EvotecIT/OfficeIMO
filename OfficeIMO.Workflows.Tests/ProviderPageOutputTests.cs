using OfficeIMO.Drawing;
using OfficeIMO.Pdf;

namespace OfficeIMO.Workflows.Tests;

public sealed class ProviderPageOutputTests {
    [Theory]
    [InlineData(0)]
    [InlineData(1)]
    [InlineData(2)]
    [InlineData(3)]
    public async Task PageExportRechecksProviderAfterHostAuthorization(int failureMode) {
        string root = Path.Combine(Path.GetTempPath(), "officeimo-provider-pages-" + Guid.NewGuid().ToString("N"));
        string output = Path.Combine(root, "pages");
        Directory.CreateDirectory(output);
        string marker = Path.Combine(output, "original.txt");
        File.WriteAllText(marker, "existing output");
        using var cancellation = new CancellationTokenSource();
        byte[] bytes = CreatePdf(2);
        bool denied = false;
        int opened = 0;
        int closed = 0;
        try {
            var result = await new OfficeWorkflowRunner().ExportPdfPagesAsync(new() {
                InputPath = "content://provider/opaque-pages",
                InputStream = new("Selected.pdf", _ => {
                    if (denied) throw new UnauthorizedAccessException("Read permission expired.");
                    opened++;
                    return Task.FromResult<Stream>(new ReadStream(bytes, () => closed++));
                }),
                OutputDirectory = output, MaximumDimension = 64, ConflictPolicy = OfficeWorkflowConflictPolicy.Replace,
                PublicationGuard = new CallbackGuard(() => {
                    if (failureMode == 1) bytes = CreatePdf(3);
                    if (failureMode == 2) denied = true;
                    if (failureMode == 3) cancellation.Cancel();
                })
            }, cancellationToken: cancellation.Token);
            Assert.Equal(opened, closed);
            Assert.Equal(failureMode == 0 ? OfficeWorkflowStatus.Completed : failureMode == 3
                ? OfficeWorkflowStatus.Cancelled : OfficeWorkflowStatus.Failed, result.Status);
            if (failureMode == 0) {
                Assert.False(File.Exists(marker));
                Assert.Equal(2, result.Files.Count);
                foreach (var file in result.Files) {
                    Assert.True(OfficeImageReader.TryValidateContent(File.ReadAllBytes(file.Path), file.Path, default, out var image));
                    Assert.Equal(file.Width, image.Width);
                }
            } else {
                Assert.Null(result.OutputDirectory);
                Assert.Equal("existing output", File.ReadAllText(marker));
                Assert.Single(Directory.GetFiles(output));
            }
        } finally { Directory.Delete(root, true); }
    }

    [Fact]
    public async Task ProviderFileCannotExportOverItsLocalSourceDirectory() {
        string root = Path.Combine(Path.GetTempPath(), "officeimo-provider-pages-" + Guid.NewGuid().ToString("N"));
        Directory.CreateDirectory(root);
        string input = Path.Combine(root, "source.pdf");
        byte[] bytes = CreatePdf(1);
        File.WriteAllBytes(input, bytes);
        try {
            var result = await new OfficeWorkflowRunner().ExportPdfPagesAsync(new() {
                InputPath = input, InputStream = new("source.pdf", _ => Task.FromResult<Stream>(File.OpenRead(input))),
                OutputDirectory = root, ConflictPolicy = OfficeWorkflowConflictPolicy.Replace
            });
            Assert.Equal(OfficeWorkflowStatus.Failed, result.Status);
            Assert.Equal(bytes, File.ReadAllBytes(input));
        } finally { Directory.Delete(root, true); }
    }

    [Fact]
    public async Task ProviderPageInputLimitStopsBeforePublishing() {
        string root = Path.Combine(Path.GetTempPath(), "officeimo-provider-pages-" + Guid.NewGuid().ToString("N"));
        try {
            var result = await new OfficeWorkflowRunner().ExportPdfPagesAsync(new() {
                InputPath = "content://provider/selected", InputStream = new("selected.pdf", _ => Task.FromResult<Stream>(new MemoryStream(CreatePdf(1)))),
                OutputDirectory = root, Limits = new() { MaximumInputBytes = 1 }
            });
            Assert.Equal(OfficeWorkflowStatus.Failed, result.Status);
            Assert.False(Directory.Exists(root));
        } finally { if (Directory.Exists(root)) Directory.Delete(root, true); }
    }

    [Fact]
    public void PrintPlanningUsesTheOpenedDocumentForAProviderReference() {
        var document = PdfDocument.Load(CreatePdf(3));
        var plan = PdfPrintPlanner.Create(document, new() {
            InputPath = "content://provider/opaque", Pages = "1,3", PagesPerSheet = 2
        });
        Assert.Equal(new[] { 1, 3 }, plan.SelectedPages);
        Assert.Equal(2, Assert.Single(plan.Sheets).Placements.Count);
        Assert.Equal(3, plan.SourcePageCount);
    }

    [Fact]
    public void OpenedDocumentPlanningPreservesAuthenticatedPrintingRestrictions() {
        byte[] bytes = PdfDocument.Create(compose => compose.Page(page => page.Size(300, 400)),
            new PdfOptions().SetEncryption(new PdfStandardEncryptionOptions("reader") {
                OwnerPassword = "owner", AllowedPermissions = PdfStandardPermissions.None
            })).ToBytes();
        var document = PdfDocument.Load(bytes, new PdfLoadOptions { Password = "reader" });
        var exception = Assert.Throws<PdfPermissionDeniedException>(() => PdfPrintPlanner.Create(document, new() {
            InputPath = "content://provider/protected", PermissionPolicy = PdfPermissionPolicy.IgnoreRestrictions
        }));
        Assert.Equal(PdfStandardPermissions.Print, exception.Permission);
    }

    private static byte[] CreatePdf(int count) => PdfDocument.Create(compose => {
        for (int index = 0; index < count; index++) compose.Page(page => page.Size(300, 400));
    }).ToBytes();

    private sealed class CallbackGuard(Action callback) : IOfficeWorkflowPublicationGuard {
        public ValueTask<bool> CanPublishAsync(string path, bool isDirectory, CancellationToken token) {
            callback();
            return ValueTask.FromResult(true);
        }
    }

    private sealed class ReadStream(byte[] bytes, Action closed) : MemoryStream(bytes) {
        private bool _closed;
        protected override void Dispose(bool disposing) {
            if (!_closed) { _closed = true; closed(); }
            base.Dispose(disposing);
        }
    }
}
