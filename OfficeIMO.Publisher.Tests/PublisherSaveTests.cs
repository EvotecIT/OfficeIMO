using System.Threading.Tasks;

namespace OfficeIMO.Publisher.Tests;

public sealed class PublisherSaveTests {
    [Fact]
    public void Exact_source_copy_owns_its_bytes_and_reports_saving_separately_from_reading() {
        byte[] source = Source();
        PublisherDocument document = PublisherDocument.Load(source);
        var result = document.ToPubBytesResult();
        Assert.Equal(source, result.Value);
        Assert.True(result.Report.IsExactSourceCopy);
        Assert.False(result.Report.HasErrors); Assert.False(result.Report.HasLoss);
        Assert.Equal(source.Length, result.Report.OutputByteCount);
        Assert.Same(document.SourceReadReport, result.Report.SourceReadReport);
        result.Report.RequireNoLoss();
        result.Value[0] ^= 0xFF;
        Assert.Equal(source, document.ToPubBytes());
    }

    [Fact]
    public void Supported_edits_require_explicit_loss_acceptance_and_keep_the_original_source_baseline() {
        PublisherDocument document = EditedRectangle();
        PublisherReadReport original = document.SourceReadReport;
        var refused = Assert.Throws<OfficeConversionException>(() => document.ToPubBytes());
        var blocked = Assert.IsType<PublisherSaveReport>(refused.Report);
        Assert.True(blocked.HasLoss); Assert.False(blocked.HasErrors);
        Assert.False(blocked.IsExactSourceCopy); Assert.True(blocked.OutputByteCount > 0);
        var accepted = document.ToPubBytesResult(new PublisherSaveOptions { LossPolicy = OfficeConversionLossPolicy.Allow });
        PublisherShape shape = Assert.Single(PublisherDocument.Load(accepted.Value).Pages[0].Shapes);
        Assert.Equal((36D, 48D, 120D, 84D), (shape.X, shape.Y, shape.Width, shape.Height));
        Assert.Same(original, document.SourceReadReport);
        Assert.Throws<OfficeConversionException>(() => document.ToPubBytes());
        Assert.Equal(accepted.Value, document.ToPubBytes(new PublisherSaveOptions { LossPolicy = OfficeConversionLossPolicy.Allow }));
    }

    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public void Loss_acceptance_never_enables_source_free_or_unsupported_topology_writing(bool sourceFree) {
        PublisherDocument document = sourceFree ? PublisherDocument.Create(612, 792) : PublisherDocument.Load(Source());
        if (!sourceFree) document.AddPage();
        var error = Assert.Throws<OfficeConversionException>(() => document.ToPubBytes(new PublisherSaveOptions { LossPolicy = OfficeConversionLossPolicy.Allow }));
        var report = Assert.IsType<PublisherSaveReport>(error.Report);
        Assert.True(report.HasErrors); Assert.True(report.HasLoss);
        Assert.Equal(0, report.OutputByteCount);
    }

    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public async Task File_saving_preserves_conflicts_by_default_and_explicitly_replaces_complete_output(bool asynchronous) {
        string path = Path.Combine(Path.GetTempPath(), Guid.NewGuid().ToString("N") + ".PUB");
        byte[] sentinel = { 9, 8, 7 };
        byte[] source = Source();
        PublisherDocument document = PublisherDocument.Load(source);
        try {
            File.WriteAllBytes(path, sentinel);
            await Assert.ThrowsAsync<IOException>(() => SaveFile(document, path, asynchronous));
            Assert.Equal(sentinel, File.ReadAllBytes(path));
            PublisherSaveReport report = await SaveFile(document, path, asynchronous, new PublisherSaveOptions { FileConflictPolicy = OfficeConversionFileConflictPolicy.Replace });
            Assert.True(report.IsExactSourceCopy); Assert.Equal(source, File.ReadAllBytes(path));
            File.Delete(path);
            await SaveFile(document, path, asynchronous);
            Assert.Equal(source, File.ReadAllBytes(path));
        } finally { File.Delete(path); }
    }

    [Theory]
    [InlineData(false, false)]
    [InlineData(false, true)]
    [InlineData(true, false)]
    [InlineData(true, true)]
    public async Task Streams_remain_open_and_seekability_controls_replacement_or_append(bool asynchronous, bool seekable) {
        byte[] source = Source();
        PublisherDocument document = PublisherDocument.Load(source);
        using var storage = new MemoryStream();
        byte[] prefix = { 9, 8, 7 };
        storage.Write(prefix, 0, prefix.Length);
        if (seekable) { storage.SetLength(source.Length + 17); storage.Position = 2; }
        using var append = new AppendStream(storage);
        Stream destination = seekable ? storage : append;
        PublisherSaveReport report = await SaveStream(document, destination, asynchronous);
        Assert.True(destination.CanWrite); Assert.True(report.IsExactSourceCopy);
        Assert.Equal(seekable ? source : prefix.Concat(source).ToArray(), storage.ToArray());
        Assert.Equal(seekable ? 0 : source.Length + prefix.Length, storage.Position);
    }

    [Theory]
    [InlineData("policy", false)]
    [InlineData("policy", true)]
    [InlineData("limit", false)]
    [InlineData("limit", true)]
    [InlineData("cancellation", false)]
    [InlineData("cancellation", true)]
    [InlineData("unsupported", false)]
    [InlineData("unsupported", true)]
    public async Task Preparation_failures_preserve_existing_files_and_stream_bytes_and_position(string failure, bool asynchronous) {
        PublisherDocument document = failure == "policy" ? EditedRectangle() : PublisherDocument.Load(Source());
        if (failure == "unsupported") document.AddPage();
        var settings = new PublisherSaveOptions { FileConflictPolicy = OfficeConversionFileConflictPolicy.Replace,
            MaxOutputBytes = failure == "limit" ? 1 : 64L * 1024 * 1024,
            LossPolicy = failure == "unsupported" ? OfficeConversionLossPolicy.Allow : OfficeConversionLossPolicy.Block };
        CancellationToken token = new CancellationToken(failure == "cancellation");
        Type expected = failure == "limit" ? typeof(InvalidDataException) : failure == "cancellation" ? typeof(OperationCanceledException) : typeof(OfficeConversionException);
        string path = Path.Combine(Path.GetTempPath(), Guid.NewGuid().ToString("N") + ".pub");
        byte[] sentinel = { 9, 8, 7 };
        using var stream = new MemoryStream();
        stream.Write(sentinel, 0, sentinel.Length); stream.Position = 1;
        try {
            File.WriteAllBytes(path, sentinel);
            Assert.IsType(expected, await Record.ExceptionAsync(() => SaveFile(document, path, asynchronous, settings, token)));
            Assert.Equal(sentinel, File.ReadAllBytes(path));
            Assert.IsType(expected, await Record.ExceptionAsync(() => SaveStream(document, stream, asynchronous, settings, token)));
            Assert.Equal(sentinel, stream.ToArray()); Assert.Equal(1, stream.Position); Assert.True(stream.CanWrite);
        } finally { File.Delete(path); }
    }

    [Fact]
    public void Save_settings_are_validated_and_native_destinations_are_explicit() {
        var settings = new PublisherSaveOptions();
        PublisherSaveOptions clone = settings.Clone(); clone.MaxOutputBytes = 1;
        Assert.Equal(64L * 1024 * 1024, settings.MaxOutputBytes);
        settings.MaxOutputBytes = 0; Assert.Throws<ArgumentOutOfRangeException>(() => settings.Clone());
        settings.MaxOutputBytes = (long)int.MaxValue + 1; Assert.Throws<ArgumentOutOfRangeException>(() => settings.Clone());
        settings.MaxOutputBytes = 1; settings.LossPolicy = (OfficeConversionLossPolicy)99;
        Assert.Throws<ArgumentOutOfRangeException>(() => settings.Clone());
        settings.LossPolicy = OfficeConversionLossPolicy.Block; settings.FileConflictPolicy = (OfficeConversionFileConflictPolicy)99;
        Assert.Throws<ArgumentOutOfRangeException>(() => settings.Clone());
        PublisherDocument document = PublisherDocument.Load(Source());
        Assert.Throws<ArgumentException>(() => document.Save("publication.pdf"));
        Assert.Throws<ArgumentNullException>(() => document.Save((string)null!));
        Assert.Throws<ArgumentNullException>(() => document.Save((Stream)null!));
        using var readOnly = new MemoryStream(Source(), writable: false);
        Assert.Throws<ArgumentException>(() => document.Save(readOnly));
    }

    private static byte[] Source() => File.ReadAllBytes(PublisherNativeTests.Fixture("NativeComparison/Rectangle/native-rectangle.pub"));
    private static PublisherDocument EditedRectangle() {
        PublisherDocument document = PublisherDocument.Load(Source());
        document.SetShapeBounds(Assert.Single(document.Pages[0].Shapes).Id, 36, 48, 120, 84);
        return document;
    }
    private static Task<PublisherSaveReport> SaveFile(PublisherDocument document, string path, bool asynchronous,
        PublisherSaveOptions? settings = null, CancellationToken token = default) => asynchronous
        ? document.SaveAsync(path, settings, token) : Task.FromResult(document.Save(path, settings, token));
    private static Task<PublisherSaveReport> SaveStream(PublisherDocument document, Stream stream, bool asynchronous,
        PublisherSaveOptions? settings = null, CancellationToken token = default) => asynchronous
        ? document.SaveAsync(stream, settings, token) : Task.FromResult(document.Save(stream, settings, token));

    private sealed class AppendStream : Stream {
        private readonly Stream _storage;
        internal AppendStream(Stream storage) => _storage = storage;
        public override bool CanWrite => _storage.CanWrite;
        public override bool CanRead => false;
        public override bool CanSeek => false;
        public override long Length => throw new NotSupportedException();
        public override long Position { get => throw new NotSupportedException(); set => throw new NotSupportedException(); }
        public override void Flush() => _storage.Flush();
        public override void Write(byte[] buffer, int offset, int count) => _storage.Write(buffer, offset, count);
        public override int Read(byte[] buffer, int offset, int count) => throw new NotSupportedException();
        public override long Seek(long offset, SeekOrigin origin) => throw new NotSupportedException();
        public override void SetLength(long value) => throw new NotSupportedException();
    }
}
