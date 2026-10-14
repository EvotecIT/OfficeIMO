using System.IO.Compression;
using System.Threading;
using System.Security.Cryptography;
using System.Text.Json;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Validation;
using A = DocumentFormat.OpenXml.Drawing;
using DW = DocumentFormat.OpenXml.Drawing.Wordprocessing;
using W = DocumentFormat.OpenXml.Wordprocessing;

namespace OfficeIMO.IWork.Tests;

public sealed class IWorkPagesImageWritingTests {
    [Fact]
    public void Accepted_example_matches_independent_creation_and_reference_artifacts() {
        string corpus = Path.Combine(AppContext.BaseDirectory, "Documents", "IWorkCorpus", "native-exports");
        using var manifest = JsonDocument.Parse(File.ReadAllText(Path.Combine(corpus, "pages-created-images-v15.4.json")));
        JsonElement artifacts = manifest.RootElement.GetProperty("artifacts");
        foreach (JsonProperty artifact in artifacts.EnumerateObject()) {
            byte[] bytes = File.ReadAllBytes(Path.Combine(corpus, artifact.Value.GetProperty("path").GetString()!));
            Assert.Equal(artifact.Value.GetProperty("sha256").GetString(), Convert.ToHexString(SHA256.HashData(bytes)).ToLowerInvariant());
        }
        byte[] Payload(string name) => File.ReadAllBytes(Path.Combine(corpus, "keynote-created-images-" + name));
        var document = IWorkPagesDocument.Create(); document.AddParagraph("Before 😀");
        document.AddImage(Payload("opaque.png"), 180, 120, "Opaque four-color PNG 😀"); document.AddParagraph("Between");
        document.AddImage(Payload("alpha.png"), 120, 80, "Transparent, half-alpha, opaque purple");
        document.AddTable(IWorkTableData.Create(2, 2).SetText(1, 1, "Grid").SetNumber(2, 2, 42), "Values", 100, 28);
        document.AddImage(Payload("jpeg.jpg"), 160, 80, "Four-color JPEG, stretched frame"); document.AddParagraph("After");
        Assert.Equal(File.ReadAllBytes(Path.Combine(corpus, artifacts.GetProperty("created").GetProperty("path").GetString()!)), document.SaveBytes());
    }

    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public void Apple_accepted_and_resized_inline_images_keep_frames_and_payloads_after_strict_saved_Word(bool edited) {
        string name = edited ? "pages-created-images-edited-v15.4.pages" : "pages-created-images-v15.4.pages";
        var source = IWorkSourceDocument.Open(Path.Combine(AppContext.BaseDirectory, "Documents", "IWorkCorpus", "native-exports", name));
        var projection = source.ReadPages();
        Assert.Equal(3, projection.Images.Count); Assert.Empty(projection.Diagnostics);
        if (edited) Assert.Equal(20.392578125, projection.Images[0].Geometry!.TopPoints);
        using var result = source.ToWordDocumentResult(new() { Mode = IWorkConversionMode.EditableOnly });
        result.RequireCompleteEditableReconstruction();
        Assert.DoesNotContain(result.Report.Diagnostics, diagnostic => diagnostic.Code == "IWORK_PAGES_DOCX_PRECISION");
        using var output = new MemoryStream(); result.Value.Save(output); output.Position = 0;
        using var saved = WordprocessingDocument.Open(output, false); Assert.Empty(new OpenXmlValidator().Validate(saved));
        var drawings = saved.MainDocumentPart!.Document!.Body!.Descendants<DW.Inline>().ToArray();
        Assert.Equal(3, drawings.Length);
        for (int index = 0; index < drawings.Length; index++) {
            IWorkImageAsset image = projection.Images[index];
            Assert.Equal(image.Geometry!.WidthPoints * 12700, drawings[index].Extent!.Cx!.Value);
            Assert.Equal(image.Geometry.HeightPoints * 12700, drawings[index].Extent!.Cy!.Value);
            Assert.Equal(image.AccessibilityDescription, drawings[index].DocProperties!.Description!.Value);
            using Stream resource = saved.MainDocumentPart.GetPartById(drawings[index].Descendants<A.Blip>().Single().Embed!.Value!).GetStream();
            using var bytes = new MemoryStream(); resource.CopyTo(bytes); Assert.Equal(image.GetBytes(), bytes.ToArray());
        }
        Assert.Equal(edited ? 144 : 180, projection.Images[0].Geometry!.WidthPoints);
    }

    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public void Inline_images_keep_payload_frame_description_and_order_around_a_table_in_saved_Word(bool jpeg) {
        byte[] payload = IWorkKeynoteImageWritingTests.ImageBytes(jpeg);
        var document = IWorkPagesDocument.Create(); document.AddParagraph("Before 😀");
        IWorkPagesImage image = document.AddImage(payload, 120, 80, "Alt 😀");
        document.AddTable(IWorkTableData.Create(1, 1).SetText(1, 1, "Grid"));
        document.AddImage(payload, 80, 120); document.AddParagraph("After");
        Assert.Equal(new[] { typeof(IWorkPagesParagraphData), typeof(IWorkPagesImage), typeof(IWorkPagesTable),
            typeof(IWorkPagesImage), typeof(IWorkPagesParagraphData) }, document.BodyBlocks.Select(block => block.GetType()));
        Assert.Equal(24, image.PixelWidth); Assert.Equal(16, image.PixelHeight);
        Assert.Equal(jpeg ? "image/jpeg" : "image/png", image.MediaType);
        byte[] bytes = document.SaveBytes(); Assert.Equal(bytes, document.SaveBytes());
        var source = IWorkSourceDocument.Open(bytes); var projection = source.ReadPages();
        Assert.Empty(projection.Diagnostics); Assert.Empty(projection.SourceDeclarationIssues); Assert.Empty(projection.SourceReferenceIssues);
        Assert.Equal(2, projection.Images.Count);
        for (int index = 0; index < 2; index++) {
            var actual = projection.Images[index]; Assert.Equal(payload, actual.GetBytes()); Assert.False(actual.HasMask);
            Assert.Equal(0, actual.Geometry!.LeftPoints); Assert.Equal(0, actual.Geometry.TopPoints);
            Assert.Equal(index == 0 ? 120 : 80, actual.Geometry.WidthPoints);
            Assert.Equal(index == 0 ? 80 : 120, actual.Geometry.HeightPoints);
        }
        Assert.Equal("Alt 😀", projection.Images[0].AccessibilityDescription);
        using var result = source.ToWordDocumentResult(new() { Mode = IWorkConversionMode.EditableOnly });
        result.RequireCompleteEditableReconstruction();
        using var output = new MemoryStream(); result.Value.Save(output); output.Position = 0;
        using var saved = WordprocessingDocument.Open(output, false); Assert.Empty(new OpenXmlValidator().Validate(saved));
        W.Body body = saved.MainDocumentPart!.Document!.Body!;
        var blocks = body.ChildElements.Where(element => element is W.Table
            || element is W.Paragraph && (element.InnerText.Length > 0 || element.Descendants<W.Drawing>().Any())).ToArray();
        Assert.Equal(new[] { typeof(W.Paragraph), typeof(W.Paragraph), typeof(W.Table), typeof(W.Paragraph), typeof(W.Paragraph) },
            blocks.Select(element => element.GetType()));
        Assert.Equal("Before 😀", body.FirstChild!.InnerText);
        Assert.Equal("After", blocks.Last().InnerText);
        DW.Inline[] drawings = body.Descendants<DW.Inline>().ToArray(); Assert.Equal(2, drawings.Length);
        Assert.Equal("Alt 😀", drawings[0].DocProperties!.Description!.Value);
        Assert.Equal(120 * 12700, drawings[0].Extent!.Cx!.Value); Assert.Equal(80 * 12700, drawings[0].Extent!.Cy!.Value);
        foreach (DW.Inline drawing in drawings) {
            string relationship = drawing.Descendants<A.Blip>().Single().Embed!.Value!;
            using Stream part = saved.MainDocumentPart.GetPartById(relationship).GetStream();
            using var data = new MemoryStream(); part.CopyTo(data); Assert.Equal(payload, data.ToArray());
        }
    }

    [Fact]
    public void An_image_only_body_snapshots_bytes_and_rejected_additions_leave_collections_unchanged() {
        var document = IWorkPagesDocument.Create(); byte[] input = IWorkKeynoteImageWritingTests.ImageBytes(false);
        byte[] expected = (byte[])input.Clone(); IWorkPagesImage image = document.AddImage(input, 120, 80);
        byte[] before = document.SaveBytes(); Array.Clear(input); Array.Clear(image.GetBytes());
        Assert.Equal(expected, image.GetBytes()); Assert.Equal(before, document.SaveBytes());
        Assert.Throws<ArgumentException>(() => document.AddImage(expected.Take(expected.Length - 5).ToArray(), 120, 80));
        Assert.Throws<ArgumentException>(() => document.AddImage([71, 73, 70, 56, 57, 97], 120, 80));
        Assert.Throws<ArgumentOutOfRangeException>(() => document.AddImage(expected, float.NaN, 80));
        Assert.Throws<ArgumentException>(() => document.AddImage(expected, 469, 80));
        Assert.Throws<ArgumentException>(() => document.AddImage(expected, 120, 649));
        Assert.Throws<ArgumentException>(() => document.AddImage(expected, 120, 80, new string('x', 4097)));
        Assert.Equal(before, document.SaveBytes()); Assert.Same(image, Assert.Single(document.BodyBlocks));
        Assert.Empty(document.Paragraphs); Assert.Empty(document.Tables);
        Assert.Throws<NotSupportedException>(() => ((IList<IWorkPagesImage>)document.Images).Clear());
        using var result = IWorkSourceDocument.Open(before).ToWordDocumentResult(new() { Mode = IWorkConversionMode.EditableOnly });
        result.RequireCompleteEditableReconstruction();
        Assert.Equal(expected, Assert.Single(IWorkSourceDocument.Open(before).ReadPages().Images).GetBytes());
    }

    [Theory]
    [InlineData("count")]
    [InlineData("encoded")]
    [InlineData("decoded")]
    [InlineData("characters")]
    [InlineData("paragraphs")]
    [InlineData("runs")]
    [InlineData("package")]
    public void Aggregate_image_limits_preserve_the_caller_stream_and_existing_file(string boundary) {
        byte[] payload = IWorkKeynoteImageWritingTests.ImageBytes(false);
        var document = IWorkPagesDocument.Create(); document.AddImage(payload, 120, 80, "Alt 😀"); document.AddImage(payload, 80, 120);
        var limits = new IWorkPagesWriteOptions();
        switch (boundary) {
            case "count": limits.MaximumImages = 1; break;
            case "encoded": limits.MaximumEncodedImageBytes = payload.Length; break;
            case "decoded": limits.MaximumDecodedImageBytes = 24 * 16 * 4; break;
            case "characters": limits.MaximumTextCharacters = 9; break;
            case "paragraphs": limits.MaximumParagraphs = 1; break;
            case "runs": limits.MaximumTextRuns = 1; break;
            case "package": limits.MaximumPackageBytes = 64; break;
        }
        byte[] sentinel = [9, 8, 7]; using var output = new MemoryStream(); output.Write(sentinel);
        Assert.Throws<InvalidDataException>(() => document.Save(output, limits)); Assert.Equal(sentinel, output.ToArray());
        string folder = Path.Combine(Path.GetTempPath(), "OfficeIMO-PagesImages-" + Guid.NewGuid().ToString("N")); Directory.CreateDirectory(folder);
        try {
            string path = Path.Combine(folder, "existing.pages"); File.WriteAllBytes(path, sentinel);
            Assert.Throws<InvalidDataException>(() => document.Save(path, overwrite: true, options: limits));
            Assert.Equal(sentinel, File.ReadAllBytes(path)); Assert.Single(Directory.GetFiles(folder));
        } finally { Directory.Delete(folder, recursive: true); }
    }

    [Fact]
    public void Exact_image_budgets_count_each_instance_description_and_attachment_marker() {
        byte[] payload = IWorkKeynoteImageWritingTests.ImageBytes(false);
        var document = IWorkPagesDocument.Create(); document.AddImage(payload, 120, 80, "Alt 😀"); document.AddImage(payload, 80, 120);
        byte[] bytes = document.SaveBytes(new() { MaximumImages = 2, MaximumEncodedImageBytes = payload.Length * 2L,
            MaximumDecodedImageBytes = 24 * 16 * 4 * 2, MaximumTextCharacters = 10, MaximumParagraphs = 2, MaximumTextRuns = 2 });
        Assert.Equal(2, IWorkSourceDocument.Open(bytes).ReadPages().Images.Count);
    }

    [Fact]
    public void Cancellation_before_add_or_save_preserves_the_model_and_destination_stream() {
        using var cancel = new CancellationTokenSource(); cancel.Cancel(); var document = IWorkPagesDocument.Create();
        byte[] payload = IWorkKeynoteImageWritingTests.ImageBytes(false);
        Assert.Throws<OperationCanceledException>(() => document.AddImage(payload, 120, 80, cancellationToken: cancel.Token));
        Assert.Empty(document.BodyBlocks); Assert.Empty(document.Images); document.AddImage(payload, 120, 80);
        using var stream = new MemoryStream();
        Assert.Throws<OperationCanceledException>(() => document.Save(stream, cancellationToken: cancel.Token)); Assert.Empty(stream.ToArray());
    }

    [Fact]
    public void Native_identity_includes_image_pixels_frame_description_and_body_order() {
        string Identity(bool changed = false, int width = 120, string description = "Alt", bool before = true) {
            var document = IWorkPagesDocument.Create(); if (before) document.AddParagraph("Label");
            document.AddImage(IWorkKeynoteImageWritingTests.ImageBytes(false, changed), width, 80, description);
            if (!before) document.AddParagraph("Label");
            using var zip = new ZipArchive(new MemoryStream(document.SaveBytes()));
            using var text = new StreamReader(zip.GetEntry("Metadata/DocumentIdentifier")!.Open()); return text.ReadToEnd();
        }
        string original = Identity(); Assert.Equal(original, Identity());
        Assert.NotEqual(original, Identity(changed: true)); Assert.NotEqual(original, Identity(width: 121));
        Assert.NotEqual(original, Identity(description: "Other")); Assert.NotEqual(original, Identity(before: false));
    }
}
