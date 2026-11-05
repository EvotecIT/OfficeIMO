namespace OfficeIMO.OneNote.Tests;

public sealed class SectionLifecycleTests {
    [Fact]
    public void NativeSectionLifecycleKeepsFilesystemAssociationSeparateFromPackageProvenance() {
        string directory = Path.Combine(Path.GetTempPath(), Guid.NewGuid().ToString("N"));
        Directory.CreateDirectory(directory);
        string path = Path.Combine(directory, "section.one");
        try {
            using var section = OneNoteSection.Create(path);
            section.Name = "Lifecycle";
            section.Pages.Add(new OneNotePage { Title = "First" });
            Assert.False(File.Exists(path));
            section.Save();
            using var loaded = OneNoteSection.Load(path);
            loaded.Pages[0].Title = "Updated";
            loaded.Save();
            Assert.Equal("Updated", OneNoteSection.Load(path).Pages[0].Title);
            Assert.Equal(Path.GetFullPath(path), loaded.FilePath);
            loaded.SourcePath = "notebook-package/entry.one";
            using var stream = new MemoryStream();
            loaded.Save(stream);
            Assert.True(stream.CanRead);
            Assert.Equal(Path.GetFullPath(path), loaded.FilePath);
            stream.Position = 0;
            using var fromStream = OneNoteSection.Load(stream);
            Assert.Null(fromStream.FilePath);
            Assert.Throws<InvalidOperationException>(() => fromStream.Save());
        } finally {
            Directory.Delete(directory, recursive: true);
        }
    }
}
