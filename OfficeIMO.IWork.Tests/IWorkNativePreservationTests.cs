using System.Threading;
using OfficeIMO.PowerPoint;

namespace OfficeIMO.IWork.Tests;

public sealed class IWorkNativePreservationTests {
    public static IEnumerable<object[]> UnchangedSources() => IWorkCorpusTests.Corpus().Concat(new[] {
        new object[] { "native-exports/pages-created-edited-v15.4.pages", IWorkDocumentKind.Pages },
        new object[] { "native-exports/numbers-created-edited-v14.5.numbers", IWorkDocumentKind.Numbers },
        new object[] { "native-exports/keynote-created-images-edited-v15.4.key", IWorkDocumentKind.Keynote }
    });

    [Theory]
    [MemberData(nameof(UnchangedSources))]
    public void Unchanged_save_retains_all_entries_unknown_payloads_native_ids_versions_and_references(string fixture, IWorkDocumentKind kind) {
        IWorkSourceDocument source = IWorkSourceDocument.Open(Fixture(fixture), kind, new() { PreserveSourceRecords = false });
        IWorkNativeSaveResult result = source.BeginEdit().SaveBytesResult();
        IWorkSourceDocument reopened = IWorkSourceDocument.Open(result.GetPackageBytes(), kind);
        Assert.Equal(source.ComputePackageContentHash(), reopened.ComputePackageContentHash());
        Assert.Equal(result.Report.SourcePackageContentHash, result.Report.OutputPackageContentHash);
        Assert.Equal(source.Entries.Select(entry => entry.Path), reopened.Entries.Select(entry => entry.Path));
        foreach (var entry in source.Entries)
            Assert.Equal(entry.GetBytes(), reopened.Entries.Single(candidate => candidate.Path == entry.Path).GetBytes());
        Assert.Empty(result.Report.UpdatedRecords); Assert.Empty(result.Report.UpdatedEntryPaths);
        Assert.All(result.Report.Diagnostics, diagnostic => Assert.Equal(global::OfficeIMO.OfficeConversionLossKind.None, diagnostic.LossKind));
        Assert.Equal(source.Records.Count, reopened.Records.Count);
        for (int index = 0; index < source.Records.Count; index++) {
            var original = source.Records[index]; var output = reopened.Records[index];
            Assert.Equal(original.Identifier, output.Identifier); Assert.Equal(original.MessageType, output.MessageType);
            Assert.Equal(original.PayloadIndex, output.PayloadIndex); Assert.Equal(original.EntryPath, output.EntryPath);
            Assert.Equal(original.Versions, output.Versions); Assert.Equal(original.ObjectReferences, output.ObjectReferences);
            Assert.Equal(original.DataReferences, output.DataReferences); Assert.Equal(original.GetPayload(), output.GetPayload());
        }
        byte[] defensive = result.GetPackageBytes(); defensive[0] = 0;
        Assert.Equal((byte)'P', result.GetPackageBytes()[0]);
    }

    [Fact]
    public void Image_translation_preserves_other_native_records_resources_and_saved_PowerPoint_content() {
        string path = Fixture("native-exports/keynote-created-images-edited-v15.4.key");
        byte[] originalBytes = File.ReadAllBytes(path); IWorkSourceDocument source = IWorkSourceDocument.Open(originalBytes);
        IWorkImageAsset image = Jpeg(source); var session = source.BeginEdit();
        float left = (float)image.Geometry!.LeftPoints + 20, top = (float)image.Geometry.TopPoints + 10;
        session.MoveKeynoteImage(image.SourceIdentity!, left, top);
        Assert.Throws<InvalidOperationException>(() => session.SaveBytes());
        var result = session.SaveBytesResult(AcceptCaches());
        IWorkSourceDocument edited = IWorkSourceDocument.Open(result.GetPackageBytes());
        IWorkImageAsset moved = Jpeg(edited);
        Assert.Equal(left, moved.Geometry!.LeftPoints); Assert.Equal(top, moved.Geometry.TopPoints);
        Assert.Equal(image.Geometry.WidthPoints, moved.Geometry.WidthPoints); Assert.Equal(image.Geometry.HeightPoints, moved.Geometry.HeightPoints);
        Assert.Equal(image.GetBytes(), moved.GetBytes()); Assert.Equal(image.AccessibilityDescription, moved.AccessibilityDescription);
        Assert.Equal(image.Geometry.LeftPoints, Jpeg(source).Geometry!.LeftPoints); // Source remains immutable.
        Assert.Equal(originalBytes, File.ReadAllBytes(path));
        Assert.Equal(image.SourceIdentity!.RecordIdentifier, Assert.Single(result.Report.UpdatedRecords).RecordIdentifier);
        Assert.NotEqual(result.Report.SourcePackageContentHash, result.Report.OutputPackageContentHash);
        Assert.Contains(result.Report.Diagnostics, diagnostic => diagnostic.Code == "IWORK_NATIVE_VISUAL_CACHES_UNASSESSED"
            && diagnostic.LossKind == global::OfficeIMO.OfficeConversionLossKind.Unassessed);
        foreach (var entry in source.Entries.Where(entry => !result.Report.UpdatedEntryPaths.Contains(entry.Path)))
            Assert.Equal(entry.GetBytes(), edited.Entries.Single(candidate => candidate.Path == entry.Path).GetBytes());
        foreach (var record in source.Records) {
            var output = edited.Records.Single(candidate => candidate.Identifier == record.Identifier && candidate.PayloadIndex == record.PayloadIndex);
            Assert.Equal(record.Versions, output.Versions); Assert.Equal(record.ObjectReferences, output.ObjectReferences); Assert.Equal(record.DataReferences, output.DataReferences);
            if (record.Identifier != image.SourceIdentity.RecordIdentifier) Assert.Equal(record.GetPayload(), output.GetPayload());
        }
        using var converted = PowerPointIWorkConverter.ConvertKeynoteToPowerPointResult(new MemoryStream(result.GetPackageBytes()),
            conversionOptions: new() { Mode = IWorkConversionMode.EditableOnly });
        converted.RequireCompleteEditableReconstruction();
        using var outputStream = new MemoryStream(); converted.Value.Save(outputStream); outputStream.Position = 0;
        using var reopened = PowerPointPresentation.Load(outputStream);
        Assert.Empty(reopened.ValidateDocument()); Assert.Equal(source.ReadKeynote().Slides.Count, reopened.Slides.Count);
        session.MoveKeynoteImage(image.SourceIdentity, (float)image.Geometry.LeftPoints, (float)image.Geometry.TopPoints);
        Assert.Equal(source.ComputePackageContentHash(), session.SaveBytesResult().Report.OutputPackageContentHash);
    }

    [Fact]
    public void Nested_index_unchanged_save_preserves_physical_members_without_duplicating_expansion() {
        IWorkSourceDocument original = IWorkSourceDocument.Open(Fixture("native-exports/keynote-created-images-v15.4.key"));
        using var inner = new MemoryStream();
        using (var index = new ZipArchive(inner, ZipArchiveMode.Create, true))
            foreach (var entry in original.Entries.Where(entry => entry.Path.StartsWith("Index/", StringComparison.Ordinal)))
                Write(index, entry.Path.Substring(6), entry.GetBytes());
        using var outer = new MemoryStream();
        using (var archive = new ZipArchive(outer, ZipArchiveMode.Create, true)) {
            foreach (var entry in original.Entries.Where(entry => !entry.Path.StartsWith("Index/", StringComparison.Ordinal)))
                Write(archive, entry.Path, entry.GetBytes());
            Write(archive, "Index.zip", inner.ToArray());
            Write(archive, "Index/outer-only.bin", new byte[] { 9, 8, 7 });
        }
        IWorkSourceDocument source = IWorkSourceDocument.Open(outer.ToArray());
        var session = source.BeginEdit(); IWorkNativeSaveResult saved = session.SaveBytesResult();
        using var package = new ZipArchive(new MemoryStream(saved.GetPackageBytes()), ZipArchiveMode.Read);
        Assert.NotNull(package.GetEntry("Index.zip")); Assert.NotNull(package.GetEntry("Index/outer-only.bin"));
        Assert.Null(package.GetEntry("Index/Document.iwa"));
        using var indexStream = package.GetEntry("Index.zip")!.Open(); using var copy = new MemoryStream(); indexStream.CopyTo(copy);
        Assert.Equal(inner.ToArray(), copy.ToArray());
        Assert.Equal(source.ComputePackageContentHash(), saved.Report.OutputPackageContentHash);
        Assert.Throws<NotSupportedException>(() => session.MoveKeynoteImage(Jpeg(source).SourceIdentity!, 80, 100));
        Assert.Throws<NotSupportedException>(() => session.ResizeKeynoteImage(Jpeg(source).SourceIdentity!, 300, 180));
    }

    [Fact]
    public void Directory_snapshot_outputs_preserve_content_and_reject_source_members() {
        string directory = TemporaryFolder();
        try {
            string bundle = Path.Combine(directory, "source.key"); Directory.CreateDirectory(bundle);
            var zip = IWorkSourceDocument.Open(Fixture("native-exports/keynote-created-images-v15.4.key"));
            foreach (var entry in zip.Entries) {
                string member = Path.Combine(bundle, entry.Path); Directory.CreateDirectory(Path.GetDirectoryName(member)!); File.WriteAllBytes(member, entry.GetBytes());
            }
            IWorkSourceDocument source = IWorkSourceDocument.Open(bundle); var session = source.BeginEdit();
            Assert.Throws<InvalidOperationException>(() => session.Save(Path.Combine(bundle, "copy.key")));
            Assert.Throws<InvalidOperationException>(() => session.Save(Path.Combine(bundle, "Index", "Document.iwa"), overwrite: true));
            if (!OperatingSystem.IsWindows()) {
                string alias = Path.Combine(directory, "alias"); Directory.CreateSymbolicLink(alias, bundle);
                try {
                    Assert.Throws<InvalidOperationException>(() => session.Save(Path.Combine(alias, "copy.key")));
                    Assert.Throws<InvalidOperationException>(() => session.Save(Path.Combine(alias, "Index", "Document.iwa"), overwrite: true));
                } finally { Directory.Delete(alias); }
            }
            string output = Path.Combine(directory, "copy.key"); var report = session.Save(output);
            Assert.Equal(source.ComputePackageContentHash(), IWorkSourceDocument.Open(output).ComputePackageContentHash());
            Assert.Equal(source.ComputePackageContentHash(), report.OutputPackageContentHash);
            Assert.False(File.Exists(Path.Combine(bundle, "copy.key")));
        } finally { Directory.Delete(directory, true); }
    }

    [Theory]
    [InlineData("package")]
    [InlineData("archive")]
    [InlineData("records")]
    [InlineData("caches")]
    public void Save_rejections_preserve_existing_file_and_caller_stream(string boundary) {
        var source = IWorkSourceDocument.Open(Fixture("native-exports/keynote-created-images-v15.4.key")); var session = source.BeginEdit();
        var images = source.ReadKeynote().Slides.SelectMany(slide => slide.Images).Take(2).ToArray();
        foreach (var image in images) session.MoveKeynoteImage(image.SourceIdentity!, (float)image.Geometry!.LeftPoints + 1, (float)image.Geometry.TopPoints);
        var limits = AcceptCaches();
        switch (boundary) {
            case "package": limits.MaximumPackageBytes = 100; break;
            case "archive": limits.MaximumEditedArchiveBytes = 100; break;
            case "records": limits.MaximumEditedRecords = 1; break;
            case "caches": limits.AllowUnassessedVisualCaches = false; break;
        }
        byte[] sentinel = { 9, 8, 7 }; using var stream = new MemoryStream(); stream.Write(sentinel);
        Type expected = boundary == "caches" ? typeof(InvalidOperationException) : typeof(InvalidDataException);
        Assert.Throws(expected, () => session.Save(stream, limits)); Assert.Equal(sentinel, stream.ToArray());
        string directory = TemporaryFolder();
        try {
            string path = Path.Combine(directory, "existing.key"); File.WriteAllBytes(path, sentinel);
            Assert.Throws(expected, () => session.Save(path, overwrite: true, options: limits));
            Assert.Equal(sentinel, File.ReadAllBytes(path)); Assert.Single(Directory.GetFiles(directory));
        } finally { Directory.Delete(directory, true); }
    }

    [Fact]
    public void Rejected_image_edits_leave_the_session_and_source_unchanged() {
        var source = IWorkSourceDocument.Open(Fixture("native-exports/keynote-created-images-v15.4.key"));
        var session = source.BeginEdit(); var image = Jpeg(source); var projection = source.ReadKeynote();
        Assert.Throws<ArgumentOutOfRangeException>(() => session.MoveKeynoteImage(image.SourceIdentity!, float.NaN, 0));
        Assert.Throws<ArgumentOutOfRangeException>(() => session.MoveKeynoteImage(image.SourceIdentity!, -1, 0));
        Assert.Throws<ArgumentOutOfRangeException>(() => session.MoveKeynoteImage(image.SourceIdentity!, (float)projection.SlideSize!.WidthPoints, 0));
        Assert.Throws<ArgumentException>(() => session.MoveKeynoteImage(projection.Slides[0].SourceIdentity!, 0, 0));
        using var cancellation = new CancellationTokenSource(); cancellation.Cancel();
        Assert.Throws<OperationCanceledException>(() => session.MoveKeynoteImage(image.SourceIdentity!, 0, 0, cancellation.Token));
        var saved = session.SaveBytesResult(); Assert.Empty(saved.Report.UpdatedRecords);
        Assert.Equal(source.ComputePackageContentHash(), saved.Report.OutputPackageContentHash);
    }

    [Fact]
    public void Cancellation_after_copying_starts_leaves_the_caller_stream_open_and_reports_no_success() {
        var source = IWorkSourceDocument.Open(Fixture("native-exports/keynote-created-images-edited-v15.4.key"));
        byte[] before = source.Entries.Single(entry => entry.Path == "Index/Document.iwa").GetBytes();
        using var cancellation = new CancellationTokenSource(); using var stream = new CancelAfterWriteStream(cancellation);
        Assert.Throws<OperationCanceledException>(() => source.BeginEdit().Save(stream, cancellationToken: cancellation.Token));
        Assert.True(stream.CanWrite); Assert.InRange(stream.Length, 1, 65_536);
        Assert.Equal(before, source.Entries.Single(entry => entry.Path == "Index/Document.iwa").GetBytes());
    }

    [Fact]
    public void Output_carrier_overhead_is_bounded_after_all_input_entries_fit() {
        var source = IWorkSourceDocument.Open(Fixture("nim-iwork/simple.pages"));
        var limit = new IWorkNativeSaveOptions { MaximumPackageBytes = (int)source.Entries.Sum(entry => (long)entry.Length) };
        using var destination = new MemoryStream(); destination.WriteByte(9);
        Assert.Throws<InvalidDataException>(() => source.BeginEdit().Save(destination, limit));
        Assert.Equal(new byte[] { 9 }, destination.ToArray());
    }

    [Fact]
    public void Native_saved_translation_retains_the_edited_image_dimensions_resources_and_complete_saved_PowerPoint() {
        var original = IWorkSourceDocument.Open(Fixture("native-exports/keynote-created-images-edited-v15.4.key"));
        string path = Fixture("native-exports/keynote-preserved-translated-edited-v15.4.key");
        var native = IWorkSourceDocument.Open(path); var moved = Jpeg(native);
        Assert.Equal(101, moved.Geometry!.LeftPoints); Assert.Equal(130, moved.Geometry.TopPoints);
        Assert.Equal(600, moved.Geometry.WidthPoints); Assert.Equal(240, moved.Geometry.HeightPoints);
        var originalImages = original.ReadKeynote().Slides.SelectMany(slide => slide.Images).ToArray();
        var nativeImages = native.ReadKeynote().Slides.SelectMany(slide => slide.Images).ToArray();
        Assert.Equal(originalImages.Length, nativeImages.Length);
        for (int i = 0; i < originalImages.Length; i++) Assert.Equal(originalImages[i].GetBytes(), nativeImages[i].GetBytes());
        using var result = PowerPointIWorkConverter.ConvertKeynoteToPowerPointResult(path, conversionOptions: new() { Mode = IWorkConversionMode.EditableOnly });
        result.RequireCompleteEditableReconstruction(); using var stream = new MemoryStream(); result.Value.Save(stream); stream.Position = 0;
        using var saved = PowerPointPresentation.Load(stream); Assert.Equal(3, saved.Slides.Count); Assert.Empty(saved.ValidateDocument());
    }

    [Fact]
    public void Stream_position_cancellation_and_conflict_contracts_apply_to_native_preservation() {
        var source = IWorkSourceDocument.Open(Fixture("nim-iwork/simple.pages")); var session = source.BeginEdit();
        using var stream = new MemoryStream(); stream.WriteByte(9); var report = session.Save(stream);
        Assert.True(stream.CanWrite); Assert.Equal(session.SaveBytes(), stream.ToArray().Skip(1));
        Assert.Equal(source.ComputePackageContentHash(), report.OutputPackageContentHash);
        using var canceled = new CancellationTokenSource(); canceled.Cancel(); byte[] before = stream.ToArray();
        Assert.Throws<OperationCanceledException>(() => session.Save(stream, cancellationToken: canceled.Token)); Assert.Equal(before, stream.ToArray());
        string directory = TemporaryFolder();
        try {
            string path = Path.Combine(directory, "copy.pages"); session.Save(path); before = File.ReadAllBytes(path);
            Assert.Throws<IOException>(() => session.Save(path));
            Assert.Throws<OperationCanceledException>(() => session.Save(path, true, cancellationToken: canceled.Token));
            Assert.Equal(before, File.ReadAllBytes(path)); Assert.Single(Directory.GetFiles(directory));
        } finally { Directory.Delete(directory, true); }
    }

    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public void Saving_adapts_carrier_verification_to_valid_tight_read_bounds_without_relaxing_decompression(bool translate) {
        string path = Fixture("native-exports/keynote-created-images-edited-v15.4.key");
        var source = IWorkSourceDocument.Open(path, options: new() {
            MaximumPackageBytes = new FileInfo(path).Length,
            MaximumEntryBytes = 92_000, MaximumTotalEntryBytes = 92_000,
            MaximumIwaBytes = 32_768, MaximumDecompressedIwaBytes = 32_768,
            MaximumSnappyChunkBytes = 32_768, MaximumTotalDecompressedIwaBytes = 50_000,
            MaximumRecordCount = 205
        });
        Assert.Equal(91_998, source.Entries.Sum(entry => entry.Length));
        var session = source.BeginEdit();
        if (translate) session.MoveKeynoteImage(Jpeg(source).SourceIdentity!, 100, 130);
        var result = session.SaveBytesResult(translate ? AcceptCaches() : null);
        var reopened = IWorkSourceDocument.Open(result.GetPackageBytes(), IWorkDocumentKind.Keynote);
        Assert.Equal(205, reopened.Records.Count);
        if (translate) {
            Assert.Equal(93_320, reopened.Entries.Sum(entry => entry.Length));
            Assert.Equal(100, Jpeg(reopened).Geometry!.LeftPoints);
            Assert.Equal(130, Jpeg(reopened).Geometry!.TopPoints);
            Assert.Single(result.Report.UpdatedRecords);
        } else Assert.Equal(source.ComputePackageContentHash(), result.Report.OutputPackageContentHash);
        Assert.Equal(80, Jpeg(source).Geometry!.LeftPoints);
    }

    private static IWorkNativeSaveOptions AcceptCaches() => new() { AllowUnassessedVisualCaches = true };
    private static IWorkImageAsset Jpeg(IWorkSourceDocument source) => source.ReadKeynote().Slides.SelectMany(slide => slide.Images).Single(image => image.MediaType == "image/jpeg");
    private static string Fixture(string path) => Path.Combine(AppContext.BaseDirectory, "Documents", "IWorkCorpus", path);
    private static void Write(ZipArchive archive, string path, byte[] bytes) { using Stream entry = archive.CreateEntry(path).Open(); entry.Write(bytes); }
    private static string TemporaryFolder() { string path = Path.Combine(Path.GetTempPath(), "OfficeIMO-NativePreservation-" + Guid.NewGuid().ToString("N")); Directory.CreateDirectory(path); return path; }
    private sealed class CancelAfterWriteStream(CancellationTokenSource cancellation) : MemoryStream {
        public override void Write(byte[] buffer, int offset, int count) { base.Write(buffer, offset, count); cancellation.Cancel(); }
    }
}
