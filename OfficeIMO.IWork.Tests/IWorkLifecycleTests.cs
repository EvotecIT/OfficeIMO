namespace OfficeIMO.IWork.Tests;

public sealed class IWorkLifecycleTests {
    [Fact]
    public void KeynoteRebindsOnlyAfterSuccessfulFileSaves() {
        string directory = Path.Combine(Path.GetTempPath(), Guid.NewGuid().ToString("N"));
        Directory.CreateDirectory(directory);
        try {
            string first = Path.Combine(directory, "first.key");
            string second = Path.Combine(directory, "second.key");
            string blocked = Path.Combine(directory, "blocked.key");
            using var document = IWorkKeynoteDocument.Create();
            document.AddSlide();
            Assert.Null(document.FilePath);
            document.Save(first);
            Assert.Equal(first, document.FilePath);
            byte[] original = File.ReadAllBytes(first);
            using var stream = new MemoryStream();
            document.Save(stream);
            Assert.True(stream.CanWrite);
            Assert.Equal(first, document.FilePath);
            document.AddSlide("112233");
            document.Save(second);
            Assert.Equal(second, document.FilePath);
            byte[] secondVersion = File.ReadAllBytes(second);
            File.WriteAllBytes(blocked, original);
            Assert.Throws<IOException>(() => document.Save(blocked));
            Assert.Equal(second, document.FilePath);
            Assert.Equal(original, File.ReadAllBytes(blocked));
            document.AddSlide("445566");
            document.Save(overwrite: true);
            Assert.Equal(second, document.FilePath);
            Assert.Equal(original, File.ReadAllBytes(first));
            Assert.False(secondVersion.SequenceEqual(File.ReadAllBytes(second)));
            using var reopened = IWorkSourceDocument.Load(second, IWorkDocumentKind.Keynote);
            Assert.NotEmpty(reopened.ComputePackageContentHash());
        } finally {
            Directory.Delete(directory, recursive: true);
        }
    }

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
