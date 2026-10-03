using Avalonia.Platform.Storage;
using OfficeIMO.Studio.Features.Mobile;
using OfficeIMO.Studio.Features.Editor;
using OfficeIMO.Studio.Infrastructure;
using Xunit;

namespace OfficeIMO.Studio.Tests;

public sealed class StudioMobileDocumentTests {
    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public async Task ProviderImportRecoversEditsAndSharesWorkingCopy(bool relocateContainer) {
        string root = Path.Combine(Path.GetTempPath(), "officeimo-mobile-files-" + Guid.NewGuid().ToString("N"));
        Directory.CreateDirectory(root);
        try {
            using var session = TestAppBuilder.StartSession();
            await session.Dispatch(async () => {
                byte[] original = File.ReadAllBytes(Path.Combine(AppContext.BaseDirectory, "Fixtures", "openpreserve-pdfa1b-text.pdf"));
                var file = new TestStorageFile("content://files/review.pdf", original, "Review.pdf");
                string container = Path.Combine(root, "ContainerA");
                var paths = new StudioDataPaths(Path.Combine(container, "data"));
                string documents = Path.Combine(container, "Documents");
                string? shared = null;
                string workingCopy;
                using (var first = new MobileDocumentController(StudioApplicationServices.Create(paths, new StudioLocalDocumentRoot(documents)),
                           _ => Task.FromResult<IStorageFile?>(file.Item), _ => Task.CompletedTask)) {
                    await first.Document.OpenCommand.ExecuteAsync(null);
                    Assert.True(first.Document.HasDocument, first.Document.ErrorMessage);
                    workingCopy = first.Document.DocumentPath!;
                    Assert.StartsWith(documents + Path.DirectorySeparatorChar, workingCopy);
                    Assert.Equal(original, File.ReadAllBytes(workingCopy));
                    Assert.Equal(1, file.ClosedReads);
                    Assert.Equal(1, file.Disposals);
                    first.Document.EditorText = "Retain this mobile review note";
                    await first.Document.ApplyPageMarkupAsync(PdfEditorTool.Note, new PdfEditorGesture(1, 24, 24, 48, 48, []));
                    Assert.True(first.Document.IsDirty, first.Document.ErrorMessage);
                    first.Suspend();
                }
                if (relocateContainer) {
                    string destination = Path.Combine(root, "ContainerB");
                    Directory.Move(container, destination);
                    workingCopy = Path.Combine(destination, Path.GetRelativePath(container, workingCopy));
                    documents = Path.Combine(destination, "Documents");
                    paths = new StudioDataPaths(Path.Combine(destination, "data"));
                }
                int shares = 0;
                using var restored = new MobileDocumentController(StudioApplicationServices.Create(paths, new StudioLocalDocumentRoot(documents)),
                    _ => Task.FromResult<IStorageFile?>(null), path => {
                        if (++shares == 1) throw new IOException("Share destination unavailable");
                        shared = path;
                        return Task.CompletedTask;
                    });
                await restored.RestoreAsync();
                Assert.True(restored.Document.HasDocument, restored.Document.ErrorMessage);
                Assert.True(restored.Document.IsDirty, restored.Document.ErrorMessage);
                Assert.Equal(workingCopy, restored.Document.DocumentPath);
                var shareError = await Assert.ThrowsAsync<IOException>(restored.ShareAsync);
                restored.Document.ErrorMessage = shareError.Message;
                await restored.ShareAsync();
                Assert.Equal(workingCopy, shared);
                Assert.False(restored.Document.IsDirty, restored.Document.ErrorMessage);
                Assert.False(original.SequenceEqual(File.ReadAllBytes(shared!)));
                Assert.Equal(original, file.Bytes);
                Assert.Equal(0, file.Writes);
                Assert.Equal(2, shares);
            }, CancellationToken.None);
        } finally { Directory.Delete(root, recursive: true); }
    }
}
