using System.Text;
using System.IO.Compression;
using System.Threading;
using OfficeIMO.IWork;
using OfficeIMO.IWork.Internal;
using OfficeIMO.PowerPoint;

namespace OfficeIMO.IWork.Tests;

public sealed class IWorkKeynoteWritingTests {
    [Fact]
    public void Apple_saved_native_creation_retains_the_edit_fonts_colors_geometry_and_zero_padding() {
        string path = Path.Combine(AppContext.BaseDirectory, "Documents", "IWorkCorpus", "native-exports",
            "keynote-created-edited-v15.4.key");
        IWorkKeynoteProjection projection = IWorkSourceDocument.Open(path).ReadKeynote();
        Assert.True(projection.HasEditableContent);
        Assert.Empty(projection.SourceDeclarationIssues);
        Assert.Equal(3, projection.Slides.Count);
        var box = Assert.Single(projection.Slides[1].TextBoxes);
        Assert.Equal("Zażółć gęślą jaźń\nA😀B — Edited in Keynote\n", box.Content.PlainText);
        Assert.Equal(60, box.Geometry!.LeftPoints);
        Assert.Equal(840, box.Geometry.WidthPoints);
        Assert.Equal("003366", box.Content.Paragraphs[0].Runs[0].Style.Color!.RgbHex);
        Assert.Equal("Arial", box.Content.Paragraphs[0].Runs[0].Style.FontName);
        Assert.Equal(40, box.Content.Paragraphs[0].Runs[0].Style.FontSizePoints);
        Assert.Equal(0, box.Layout!.LeftInsetPoints);
        Assert.Equal(0, box.Layout.TopInsetPoints);
        Assert.Equal("Times New Roman", projection.Slides[0].TextBoxes[1].Content.Paragraphs[0].Runs[0].Style.FontName);
        using var result = PowerPointIWorkConverter.ConvertKeynoteToPowerPointResult(path,
            conversionOptions: new IWorkConversionOptions { Mode = IWorkConversionMode.EditableOnly });
        result.Report.RequireCompleteEditableReconstruction();
        using var saved = new MemoryStream();
        result.Value.Save(saved); saved.Position = 0;
        using var reopened = PowerPointPresentation.Load(saved);
        Assert.Empty(reopened.ValidateDocument());
        Assert.Contains("Edited in Keynote", reopened.Slides[1].TextBoxes.Single().Text);
        Assert.Equal("003366", reopened.Slides[1].TextBoxes.Single().Color);
    }

    [Fact]
    public void Native_creation_preserves_positioned_multilingual_text_fonts_colors_and_blank_slides() {
        var document = Example();
        byte[] bytes = document.SaveBytes();
        using var input = new MemoryStream(bytes);
        IWorkSourceDocument source = IWorkSourceDocument.Open(input);
        Assert.True(input.CanRead);
        Assert.Equal(IWorkDocumentKind.Keynote, source.Kind);
        IWorkKeynoteProjection deck = source.ReadKeynote();
        Assert.True(deck.HasEditableContent);
        Assert.Empty(deck.SourceDeclarationIssues);
        Assert.Empty(deck.SourceReferenceIssues);
        Assert.Equal(3, deck.Slides.Count);
        Assert.Equal(960, deck.SlideSize!.WidthPoints);
        Assert.Equal("E6F2FF", deck.Slides[0].BackgroundColor!.RgbHex);
        Assert.Equal("FFE6D9", deck.Slides[1].BackgroundColor!.RgbHex);
        Assert.Equal("FFFFFF", deck.Slides[2].BackgroundColor!.RgbHex);
        Assert.Empty(deck.Slides[2].TextBoxes);
        Assert.Equal(2, deck.Slides[0].TextBoxes.Count);
        IWorkTextBox box = Assert.Single(deck.Slides[1].TextBoxes);
        Assert.Equal("Zażółć gęślą jaźń\nA😀B — café\n", box.Content.PlainText);
        Assert.Equal(60, box.Geometry!.LeftPoints);
        Assert.Equal(100, box.Geometry.TopPoints);
        Assert.Equal(840, box.Geometry.WidthPoints);
        Assert.Equal(260, box.Geometry.HeightPoints);
        IWorkTextRun run = box.Content.Paragraphs[0].Runs[0];
        Assert.Equal("Arial", run.Style!.FontName);
        Assert.Equal(40, run.Style.FontSizePoints);
        Assert.Equal("003366", run.Style.Color!.RgbHex);
        Assert.Equal("Times New Roman", deck.Slides[0].TextBoxes[1].Content.Paragraphs[0].Runs[0].Style!.FontName);

        input.Position = 0;
        using var result = PowerPointIWorkConverter.ConvertKeynoteToPowerPointResult(input,
            conversionOptions: new IWorkConversionOptions { Mode = IWorkConversionMode.EditableOnly });
        result.Report.RequireCompleteEditableReconstruction();
        using var saved = new MemoryStream();
        result.Value.Save(saved);
        saved.Position = 0;
        using var reopened = PowerPointPresentation.Load(saved);
        Assert.Equal(3, reopened.Slides.Count);
        Assert.Empty(reopened.ValidateDocument());
        Assert.Contains("Zażółć gęślą jaźń", reopened.Slides[1].TextBoxes.Single().Text);
        Assert.Equal("003366", reopened.Slides[1].TextBoxes.Single().Color);
    }

    [Fact]
    public void Native_slide_and_paragraph_profile_protects_Apple_editing_and_export_regressions() {
        IWorkSourceDocument source = IWorkSourceDocument.Open(new MemoryStream(Example().SaveBytes()));
        var ids = source.Records.Select(record => record.Identifier).ToHashSet();
        foreach (IWorkArchiveRecord record in source.Records) {
            Assert.All(record.ObjectReferences, id => Assert.Contains(id, ids));
            IWorkWireMessage wire = IWorkProtobuf.Parse(record.GetPayload(), new IWorkReadOptions());
            if (record.MessageType == 5) {
                // Apple creates a layout class when field 10 is present; putting it on a show slide breaks editing/export.
                Assert.Equal(wire.HasField(17) ? null : "Blank", wire.GetString(10));
            } else if (record.MessageType == 2001) {
                Assert.True(wire.HasField(6)); // Paragraph data is required for native editing even when there is no list.
                Assert.True(wire.HasField(14));
                Assert.True(wire.HasField(24));
            } else if (record.MessageType is 2021 or 2022) {
                // Native export rendered legacy-only colors black, which self-round-trip could not detect.
                IWorkWireMessage character = IWorkObjectIndex.TryGetMessage(wire, 11)!;
                IWorkWireMessage fill = IWorkObjectIndex.TryGetMessage(character, 46)!;
                Assert.NotNull(fill);
                Assert.Equal(character.GetBytes(7), fill.GetBytes(1));
            }
        }
    }

    [Fact]
    public void Native_packages_are_deterministic_and_content_dependent_across_byte_stream_and_path_APIs() {
        var document = Example();
        byte[] bytes = document.SaveBytes();
        Assert.Equal(bytes, Example().SaveBytes());
        using (var zip = new ZipArchive(new MemoryStream(bytes))) {
            Assert.All(zip.Entries, entry => Assert.Equal(new DateTime(1980, 1, 1), entry.LastWriteTime.DateTime));
        }
        using var stream = new MemoryStream();
        stream.WriteByte(7);
        document.Save(stream);
        Assert.True(stream.CanWrite);
        Assert.Equal(new byte[] { 7 }.Concat(bytes), stream.ToArray());
        string folder = TemporaryFolder();
        try {
            string path = Path.Combine(folder, "created.key");
            document.Save(path);
            Assert.Equal(bytes, File.ReadAllBytes(path));
            Assert.Throws<IOException>(() => document.Save(path));
            Assert.Equal(bytes, File.ReadAllBytes(path));
            document.AddSlide("000000");
            byte[] changed = document.SaveBytes();
            Assert.NotEqual(DocumentId(bytes), DocumentId(changed));
            document.Save(path, overwrite: true);
            Assert.Equal(changed, File.ReadAllBytes(path));
            Assert.Single(Directory.GetFiles(folder));
        } finally { Directory.Delete(folder, recursive: true); }
    }

    [Theory]
    [InlineData("slides")]
    [InlineData("boxes")]
    [InlineData("characters")]
    [InlineData("bytes")]
    public void Native_write_limits_reject_before_destination_mutation(string limit) {
        var options = new IWorkKeynoteWriteOptions();
        switch (limit) {
            case "slides": options.MaximumSlides = 1; break;
            case "boxes": options.MaximumTextBoxes = 1; break;
            case "characters": options.MaximumTextCharacters = 1; break;
            case "bytes": options.MaximumPackageBytes = 128; break;
        }
        var document = Example();
        using var stream = new MemoryStream(new byte[] { 9, 8, 7 });
        Assert.Throws<InvalidDataException>(() => document.Save(stream, options));
        Assert.Equal(new byte[] { 9, 8, 7 }, stream.ToArray());
        string folder = TemporaryFolder();
        try {
            string path = Path.Combine(folder, "existing.key");
            File.WriteAllBytes(path, new byte[] { 9, 8, 7 });
            Assert.Throws<InvalidDataException>(() => document.Save(path, overwrite: true, options: options));
            Assert.Equal(new byte[] { 9, 8, 7 }, File.ReadAllBytes(path));
            Assert.Single(Directory.GetFiles(folder));
        } finally { Directory.Delete(folder, recursive: true); }
    }

    [Fact]
    public void Cancellation_preserves_file_and_stream_before_encoding_and_leaves_caller_stream_open_during_copy() {
        using var cancellation = new CancellationTokenSource();
        cancellation.Cancel();
        var document = Example();
        using var stream = new MemoryStream();
        Assert.Throws<OperationCanceledException>(() => document.Save(stream, cancellationToken: cancellation.Token));
        Assert.Empty(stream.ToArray());
        Assert.True(stream.CanWrite);
        string folder = TemporaryFolder();
        try {
            string path = Path.Combine(folder, "existing.key");
            File.WriteAllBytes(path, new byte[] { 9 });
            Assert.Throws<OperationCanceledException>(() => document.Save(path, overwrite: true, cancellationToken: cancellation.Token));
            Assert.Equal(new byte[] { 9 }, File.ReadAllBytes(path));
            Assert.Single(Directory.GetFiles(folder));
        } finally { Directory.Delete(folder, recursive: true); }

        using var copying = new CancellationTokenSource();
        var large = IWorkKeynoteDocument.Create();
        large.AddSlide().AddText(new string('a', 150_000), 0, 0, 960, 540);
        using var target = new CancelAfterWriteStream(copying);
        Assert.Throws<OperationCanceledException>(() => large.Save(target, cancellationToken: copying.Token));
        Assert.True(target.CanWrite);
        Assert.Equal(65_536, target.Length);
    }

    [Fact]
    public void Multichunk_native_text_is_readable_with_64KiB_Snappy_bounds() {
        var document = IWorkKeynoteDocument.Create();
        string text = string.Concat(Enumerable.Repeat("A😀B café\n", 12_000));
        document.AddSlide().AddText(text, 0, 0, 960, 540);
        byte[] bytes = document.SaveBytes();
        IWorkSourceDocument source = IWorkSourceDocument.Open(new MemoryStream(bytes),
            options: new IWorkReadOptions { MaximumSnappyChunkBytes = 65_536 });
        Assert.Equal(text + "\n", Assert.Single(source.ReadKeynote().Slides[0].TextBoxes).Content.PlainText);
        using var zip = new ZipArchive(new MemoryStream(bytes));
        using Stream entry = zip.GetEntry("Index/Document.iwa")!.Open();
        using var encoded = new MemoryStream(); entry.CopyTo(encoded);
        byte[] chunks = encoded.ToArray(); int count = 0;
        for (int offset = 0; offset < chunks.Length;) {
            Assert.Equal(0, chunks[offset]);
            int length = chunks[offset + 1] | chunks[offset + 2] << 8 | chunks[offset + 3] << 16;
            offset += 4 + length; count++;
        }
        Assert.True(count > 1);
    }

    [Theory]
    [InlineData(float.NaN)]
    [InlineData(float.PositiveInfinity)]
    [InlineData(0)]
    [InlineData(-1)]
    public void Native_creation_rejects_invalid_canvas_and_font_sizes(float value) {
        Assert.Throws<ArgumentOutOfRangeException>(() => IWorkKeynoteDocument.Create(value));
        var slide = IWorkKeynoteDocument.Create().AddSlide();
        Assert.Throws<ArgumentOutOfRangeException>(() => slide.AddText("text", 0, 0, 100, 100, fontSizePoints: value));
        Assert.Empty(slide.TextBoxes);
    }

    [Fact]
    public void Native_creation_normalizes_lines_and_rejects_unencodable_text_and_invalid_frames() {
        var document = IWorkKeynoteDocument.Create();
        Assert.Throws<InvalidOperationException>(() => document.SaveBytes());
        Assert.Throws<ArgumentException>(() => document.AddSlide("#FFFFFF"));
        var slide = document.AddSlide();
        Assert.Equal("one\ntwo\nthree", slide.AddText("one\r\ntwo\rthree", 0, 0, 100, 100).Text);
        Assert.Throws<EncoderFallbackException>(() => slide.AddText("\ud800", 0, 0, 100, 100));
        Assert.Throws<ArgumentException>(() => slide.AddText("tab\ttext", 0, 0, 100, 100));
        Assert.Throws<ArgumentOutOfRangeException>(() => slide.AddText("outside", 900, 0, 100, 100));
        Assert.Throws<ArgumentOutOfRangeException>(() => slide.AddText("outside", float.NaN, 0, 100, 100));
        Assert.Single(slide.TextBoxes);
    }

    private static IWorkKeynoteDocument Example() {
        var document = IWorkKeynoteDocument.Create();
        var first = document.AddSlide("E6F2FF");
        first.AddText("OfficeIMO native Keynote\nCreated without a template", 60, 100, 840, 150, fontSizePoints: 40);
        first.AddText("Times New Roman — 24 pt", 60, 300, 800, 80, "Times New Roman", 24, "003366");
        document.AddSlide("FFE6D9").AddText("Zażółć gęślą jaźń\nA😀B — café", 60, 100, 840, 260,
            fontSizePoints: 40, color: "003366");
        document.AddSlide();
        return document;
    }

    private static string DocumentId(byte[] bytes) {
        using var zip = new ZipArchive(new MemoryStream(bytes));
        using var reader = new StreamReader(zip.GetEntry("Metadata/DocumentIdentifier")!.Open());
        return reader.ReadToEnd();
    }

    private static string TemporaryFolder() {
        string folder = Path.Combine(Path.GetTempPath(), "OfficeIMO-KeynoteWriting-" + Guid.NewGuid().ToString("N"));
        Directory.CreateDirectory(folder);
        return folder;
    }

    private sealed class CancelAfterWriteStream(CancellationTokenSource cancellation) : MemoryStream {
        public override void Write(byte[] buffer, int offset, int count) {
            base.Write(buffer, offset, count);
            cancellation.Cancel();
        }
    }
}
