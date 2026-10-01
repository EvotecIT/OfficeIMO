using DocumentFormat.OpenXml.Wordprocessing;
using OfficeIMO.IWork;
using OfficeIMO.Word;
using OfficeIMO.Reader;
using OfficeIMO.Reader.IWork;
using System.Text.Json;
using System.Threading;

namespace OfficeIMO.IWork.Tests;

public sealed partial class IWorkBoundaryTests {
    [Fact]
    public void Native_pages_inline_tables_and_image_keep_their_place_between_body_paragraphs_after_save() {
        using var result = WordIWorkConverter.ConvertPagesToWordResult(CorpusFixture("picodocs/sample-v14.4.pages"),
            conversionOptions: new IWorkConversionOptions { AllowPartialEditableReconstruction = true });
        using var stream = new MemoryStream(); result.Value.Save(stream); stream.Position = 0;
        using var reopened = DocumentFormat.OpenXml.Packaging.WordprocessingDocument.Open(stream, false);
        var blocks = reopened.MainDocumentPart!.Document!.Body!.ChildElements.ToArray();
        int Heading(string text) => Array.FindIndex(blocks, block => block is Paragraph && block.InnerText.StartsWith(text, StringComparison.Ordinal));
        int[] tables = blocks.Select((block, index) => (block, index)).Where(pair => pair.block is Table).Select(pair => pair.index).ToArray();
        Assert.Equal(3, tables.Length);
        Assert.True(tables[0] > Heading("3. Table") && tables[0] < Heading("4. Link and image"));
        Assert.True(tables[1] < Heading("6. Dates and formula"));
        Assert.True(tables[2] > Heading("6. Dates and formula") && tables[2] < Heading("7. End marker"));
        var imageHost = Assert.Single(blocks, block => block.Descendants<DocumentFormat.OpenXml.Drawing.Wordprocessing.Inline>().Any());
        Assert.True(Array.IndexOf(blocks, imageHost) > Heading("4. Link and image"));
        Assert.True(Array.IndexOf(blocks, imageHost) < Heading("Figure 1."));
        Assert.False(result.Projection.Body.HasUnresolvedInlineObjects);
        Assert.True(result.Projection.Body.IsTextComplete);
    }

    [Theory]
    [InlineData("iwork-converter/a.pages")]
    [InlineData("picodocs/sample-v14.4.pages")]
    public void Native_inline_attachment_offsets_and_identities_match_independently_extracted_evidence(string fixture) {
        using var manifest = JsonDocument.Parse(File.ReadAllText(CorpusFixture("pages-inline-anchors.json")));
        JsonElement source = manifest.RootElement.GetProperty("sources").EnumerateArray().Single(entry => entry.GetProperty("path").GetString() == fixture);
        var projection = IWorkSourceDocument.Open(CorpusFixture(fixture)).ReadPages();
        var anchors = projection.Body.Paragraphs.SelectMany(paragraph => paragraph.Runs).Where(run => run.InlineObject != null)
            .Select(run => run.InlineObject!).ToArray();
        var expected = source.GetProperty("anchors").EnumerateArray().ToArray();
        Assert.Equal(expected.Length, anchors.Length);
        for (int index = 0; index < anchors.Length; index++) {
            Assert.Equal(expected[index].GetProperty("storageIdentifier").GetUInt64(), projection.Body.SourceIdentity!.RecordIdentifier);
            Assert.Equal(expected[index].GetProperty("characterOffset").GetInt32(), anchors[index].CharacterOffset);
            Assert.Equal(expected[index].GetProperty("attachmentIdentifier").GetUInt64(), anchors[index].Attachment.RecordIdentifier);
            Assert.Equal(expected[index].GetProperty("drawableIdentifier").GetUInt64(), anchors[index].Drawable.RecordIdentifier);
            Assert.Equal(expected[index].GetProperty("drawableType").GetUInt32(), anchors[index].Drawable.MessageType);
        }
    }

    [Fact]
    public void Reader_keeps_native_inline_table_and_image_blocks_in_body_order_without_duplicates() {
        using var input = File.OpenRead(CorpusFixture("picodocs/sample-v14.4.pages"));
        var result = IWorkReaderAdapter.ReadDocument(input, "sample.pages", new ReaderOptions(), new ReaderIWorkOptions(), CancellationToken.None);
        var blocks = result.Blocks.ToArray();
        int Heading(string text) => Array.FindIndex(blocks, block => block.Text?.StartsWith(text, StringComparison.Ordinal) == true);
        var tables = blocks.Select((block, index) => (block, index)).Where(pair => pair.block.Kind == "table").Select(pair => pair.index).ToArray();
        Assert.Equal(3, tables.Length);
        Assert.True(tables[0] < Heading("4. Link and image"));
        Assert.True(tables[1] < Heading("6. Dates and formula"));
        Assert.True(tables[2] < Heading("7. End marker"));
        int imageIndex = Array.IndexOf(blocks, Assert.Single(blocks, block => block.Kind == "image"));
        Assert.True(imageIndex > Heading("4. Link and image") && imageIndex < Heading("Figure 1."));
    }

    [Fact]
    public void Reader_splits_mixed_inline_paragraphs_in_order_and_reports_the_layout_limit() {
        using var input = InlineImagePackage();
        var result = IWorkReaderAdapter.ReadDocument(input, "sample.pages", new ReaderOptions(), new ReaderIWorkOptions(), CancellationToken.None);
        Assert.Equal(new[] { "😀before ", "image", " middle ", "image", " after" },
            result.Blocks.Select(block => block.Kind == "image" ? "image" : block.Text));
        Assert.Contains(result.Diagnostics, diagnostic => diagnostic.Code == "IWORK_READER_INLINE_PARAGRAPH_SPLIT");
        Assert.Equal(result.Blocks.Where(block => block.Kind == "image").Select(block => block.Id),
            result.Assets.Select(asset => asset.Location.BlockAnchor));
    }

    [Fact]
    public void Inline_images_preserve_UTF16_offsets_and_text_order_with_two_images_in_one_paragraph() {
        using var package = InlineImagePackage();
        using var result = WordIWorkConverter.ConvertPagesToWordResult(package);
        Assert.False(result.IsVisualFallback);
        var anchors = result.Projection.Body.Paragraphs.SelectMany(paragraph => paragraph.Runs)
            .Where(run => run.InlineObject != null).Select(run => run.InlineObject!).ToArray();
        Assert.Equal(new[] { 9, 18 }, anchors.Select(anchor => anchor.CharacterOffset));
        using var stream = new MemoryStream(); result.Value.Save(stream); stream.Position = 0;
        using var reopened = DocumentFormat.OpenXml.Packaging.WordprocessingDocument.Open(stream, false);
        Paragraph paragraph = Assert.Single(reopened.MainDocumentPart!.Document!.Body!.Elements<Paragraph>());
        Assert.Equal(new[] { "😀before ", "image", " middle ", "image", " after" },
            paragraph.Elements<Run>().Where(run => run.InnerText.Length > 0 || run.Descendants<DocumentFormat.OpenXml.Wordprocessing.Drawing>().Any())
                .Select(run => run.Descendants<DocumentFormat.OpenXml.Wordprocessing.Drawing>().Any() ? "image" : run.InnerText));
        Assert.Equal(2, paragraph.Descendants<DocumentFormat.OpenXml.Drawing.Wordprocessing.Inline>().Count());
    }

    [Theory]
    [InlineData("missingAttachment")]
    [InlineData("missingDrawable")]
    [InlineData("offsetInsideSurrogate")]
    [InlineData("duplicateOffset")]
    [InlineData("nonzeroPlacement")]
    [InlineData("unknownEnvelope")]
    public void Unresolved_inline_attachments_keep_explicit_source_evidence_and_do_not_claim_complete_text(string defect) {
        using var package = InlineImagePackage(defect);
        IWorkPagesProjection projection = IWorkSourceDocument.Open(package, IWorkDocumentKind.Pages,
            new IWorkReadOptions { PreserveSourceRecords = false }).ReadPages();
        Assert.True(projection.Body.HasUnresolvedInlineObjects);
        Assert.False(projection.HasEditableContent);
        Assert.Contains(projection.Diagnostics, diagnostic => diagnostic.Code == "IWORK_PAGES_TEXT_UNSUPPORTED");
        if (defect == "missingAttachment") Assert.Contains(projection.SourceReferenceIssues, issue => issue.TargetIdentifier == 999);
        if (defect == "missingDrawable") Assert.Contains(projection.SourceReferenceIssues, issue => issue.TargetIdentifier == 999 && issue.Owner.RecordIdentifier == 20);
        if (defect is "offsetInsideSurrogate" or "duplicateOffset") Assert.NotEmpty(projection.SourceDeclarationIssues);
        if (defect == "nonzeroPlacement") Assert.Contains(projection.SourceDeclarationIssues, issue =>
            issue.Owner.RecordIdentifier == 20 && issue.FieldPath == "3" && issue.Kind == IWorkSourceDeclarationIssueKind.InvalidValue);
    }

    [Fact]
    public void Inline_attachment_tables_obey_the_existing_source_wide_attribute_boundary_limit() {
        using var package = InlineImagePackage();
        IWorkSourceDocument source = IWorkSourceDocument.Open(package, IWorkDocumentKind.Pages,
            new IWorkReadOptions { MaximumProjectedTextItems = 1 });
        Assert.Throws<InvalidDataException>(() => source.ReadPages());
    }

    private static MemoryStream InlineImagePackage(string? defect = null) {
        const string text = "😀before \ufffc middle \ufffc after";
        byte[] Entry(int offset, ulong target) => BytesField(1, Message(VarintField(1, (ulong)offset), ReferenceField(2, target)));
        byte[] attachments = Message(Entry(defect == "offsetInsideSurrogate" ? 1 : 9, defect == "missingAttachment" ? 999ul : 20ul),
            Entry(defect == "duplicateOffset" ? 9 : 18, 21));
        byte[] Attachment(ulong drawable, bool moved) => Message(ReferenceField(1, drawable), VarintField(2, 0),
            FloatField(3, moved ? 1f : 0f), VarintField(4, 0), FloatField(5, 0f), defect == "unknownEnvelope" ? VarintField(6, 1) : Array.Empty<byte>());
        byte[] geometry = Message(BytesField(1, Message(FloatField(1, 0f), FloatField(2, 0f))),
            BytesField(2, Message(FloatField(1, 40f), FloatField(2, 30f))), FloatField(4, 0f));
        byte[] image = Message(BytesField(1, Message(BytesField(1, geometry))), BytesField(11, Message(VarintField(1, 10))));
        byte[] records = Message(
            ArchiveRecord(1, 10000, Message(ReferenceField(4, 2)), new ulong[] { 2 }),
            ArchiveRecord(2, 2001, Message(StringField(3, text), BytesField(9, attachments)), new ulong[] { 20, 21 }),
            ArchiveRecord(20, 2003, Attachment(defect == "missingDrawable" ? 999ul : 30ul, defect == "nonzeroPlacement"), new ulong[] { 30 }),
            ArchiveRecord(21, 2003, Attachment(31, false), new ulong[] { 31 }),
            ArchiveRecord(30, 3005, image), ArchiveRecord(31, 3005, image),
            ArchiveRecord(40, 11006, Message(BytesField(4, Message(VarintField(1, 10), StringField(3, "image.png"), StringField(4, "image.png"))))));
        return CreatePackage(("Index/Document.iwa", FrameIwa(records)), ("Data/image.png", ValidPreviewPng()));
    }
}
