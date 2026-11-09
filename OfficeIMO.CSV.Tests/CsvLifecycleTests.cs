using Xunit;

namespace OfficeIMO.CSV.Tests;

public sealed class CsvLifecycleTests {
    [Fact]
    public async Task PathLifecycleRebindsOnlyAfterSuccessfulSavesAndLeavesStreamsOpen() {
        string directory = Path.Combine(Path.GetTempPath(), Guid.NewGuid().ToString("N"));
        Directory.CreateDirectory(directory);
        string first = Path.Combine(directory, "first.csv"), second = Path.Combine(directory, "second.csv");
        try {
            using var document = CsvDocument.Create(first).WithHeader("Name").AddRow("First");
            Assert.False(File.Exists(first));
            document.Save();
            using var loaded = await CsvDocument.LoadAsync(first);
            Assert.Equal(Path.GetFullPath(first), loaded.FilePath);
            loaded.AddRow("Second");
            await loaded.SaveAsync();
            using var reread = CsvDocument.Load(first);
            Assert.Equal(2, System.Linq.Enumerable.Count(reread.AsEnumerable()));
            loaded.Save(second);
            Assert.Equal(Path.GetFullPath(second), loaded.FilePath);
            Assert.Throws<IOException>(() => loaded.Save(first, new CsvSaveOptions { NoClobber = true }));
            Assert.Equal(Path.GetFullPath(second), loaded.FilePath);
            using var stream = new MemoryStream();
            await loaded.SaveAsync(stream);
            Assert.True(stream.CanRead);
            Assert.Equal(Path.GetFullPath(second), loaded.FilePath);
            using var fromStream = CsvDocument.Load(stream);
            Assert.Null(fromStream.FilePath);
            Assert.Throws<InvalidOperationException>(() => fromStream.Save());
        } finally {
            Directory.Delete(directory, recursive: true);
        }
    }
}
