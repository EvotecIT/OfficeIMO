namespace OfficeIMO.IWork.Tests;

public sealed class IWorkLifecycleTests {
    [Fact]
    public void NativePagesLifecyclePreservesOverwriteOptInAndReturnsPreservationEvidence() {
        string directory = Path.Combine(Path.GetTempPath(), Guid.NewGuid().ToString("N"));
        Directory.CreateDirectory(directory);
        string path = Path.Combine(directory, "document.pages");
        try {
            using var document = IWorkPagesDocument.Create(path);
            document.AddParagraph("Lifecycle");
            Assert.False(File.Exists(path));
            document.Save();
            Assert.Throws<IOException>(() => document.Save());
            document.Save(overwrite: true);
            using var source = IWorkSourceDocument.Load(path, IWorkDocumentKind.Pages);
            using var edit = source.BeginEdit();
            byte[] bytes = edit.ToBytes(out IWorkNativeSaveReport report);
            Assert.Equal(source.ComputePackageContentHash(), report.SourcePackageContentHash);
            Assert.Equal(report.SourcePackageContentHash, report.OutputPackageContentHash);
            Assert.Empty(report.UpdatedRecords);
            using var reopened = IWorkSourceDocument.Load(bytes);
            Assert.Equal(source.ComputePackageContentHash(), reopened.ComputePackageContentHash());
            Assert.Equal(Path.GetFullPath(path), edit.FilePath);
            using var stream = new MemoryStream();
            edit.Save(stream);
            Assert.True(stream.CanRead);
            Assert.Equal(Path.GetFullPath(path), edit.FilePath);
        } finally {
            Directory.Delete(directory, recursive: true);
        }
    }
}
