namespace OfficeIMO.Publisher.Tests;

public sealed class PublisherApiLifecycleTests {
    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public async Task Native_path_association_changes_only_after_successful_file_operations(bool asynchronous) {
        string directory = Path.Combine(Path.GetTempPath(), "publisher-lifecycle-" + Guid.NewGuid().ToString("N"));
        Directory.CreateDirectory(directory);
        string originalPath = Path.Combine(directory, "original.pub");
        string copyPath = Path.Combine(directory, "copy.pub");
        string occupiedPath = Path.Combine(directory, "occupied.pub");
        byte[] source = File.ReadAllBytes(PublisherNativeTests.Fixture("NativeComparison/Rectangle/native-rectangle.pub"));
        byte[] sentinel = { 9, 8, 7 };
        try {
            File.WriteAllBytes(originalPath, source);
            using PublisherDocument document = asynchronous
                ? await PublisherDocument.LoadAsync(originalPath, PublisherLoadOptions.UntrustedDefaults)
                : PublisherDocument.Load(originalPath, PublisherLoadOptions.UntrustedDefaults);
            PublisherReadReport baseline = document.SourceReadReport;
            Assert.Equal(originalPath, document.FilePath);
            PublisherSaveReport copy = asynchronous ? await document.SaveAsync(copyPath) : document.Save(copyPath);
            Assert.Equal(copyPath, document.FilePath);
            Assert.Equal(copyPath, copy.OutputPath);
            Assert.Equal(source, File.ReadAllBytes(copyPath));
            using var output = new MemoryStream();
            if (asynchronous) await document.SaveAsync(output); else document.Save(output);
            Assert.True(output.CanWrite);
            Assert.Equal(source, output.ToArray());
            Assert.Equal(copyPath, document.FilePath);
            using var input = new MemoryStream(source, writable: false);
            using PublisherDocument fromStream = asynchronous
                ? await PublisherDocument.LoadAsync(input) : PublisherDocument.Load(input);
            Assert.True(input.CanRead);
            Assert.Null(fromStream.FilePath);
            File.WriteAllBytes(occupiedPath, sentinel);
            if (asynchronous) await Assert.ThrowsAsync<IOException>(() => document.SaveAsync(occupiedPath));
            else Assert.Throws<IOException>(() => document.Save(occupiedPath));
            Assert.Equal(sentinel, File.ReadAllBytes(occupiedPath));
            Assert.Equal(copyPath, document.FilePath);
            File.WriteAllBytes(copyPath, sentinel);
            PublisherSaveReport associated = asynchronous ? await document.SaveAsync() : document.Save();
            Assert.Equal(copyPath, associated.OutputPath);
            Assert.Equal(source, File.ReadAllBytes(copyPath));
            Assert.Equal(source, File.ReadAllBytes(originalPath));
            Assert.Same(baseline, document.SourceReadReport);
        } finally {
            Directory.Delete(directory, recursive: true);
        }
    }

    [Fact]
    public void Physical_authoring_units_project_to_native_point_snapshots() {
        using PublisherDocument document = PublisherDocument.Create(new PublisherCreateOptions {
            Width = OfficeLength.Inches(8.5), Height = OfficeLength.Inches(11)
        });
        PublisherPage page = Assert.Single(document.Pages);
        Assert.Equal(612, page.Width);
        Assert.Equal(792, page.Height);
        PublisherTextFrame frame = document.AddTextFrame(page.Id,
            OfficeLength.Centimeters(1), OfficeLength.Inches(1), OfficeLength.Inches(2), OfficeLength.Points(36),
            Array.Empty<OfficeIMO.Drawing.OfficeRichTextParagraph>());
        Assert.Equal(72D / 2.54D, frame.X, 8);
        Assert.Equal(72, frame.Y);
        Assert.Equal(144, frame.Width);
        Assert.Equal(36, frame.Height);
    }
}
