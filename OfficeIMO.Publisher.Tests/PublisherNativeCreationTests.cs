using System.Collections;
using OfficeIMO.Core.Internal;
using OfficeIMO.Drawing;
using OfficeIMO.Publisher.Internal;

namespace OfficeIMO.Publisher.Tests;

public sealed class PublisherNativeCreationTests {
    private const long MaximumBytes = 64L * 1024 * 1024;

    [Fact]
    public void Internal_creation_keeps_the_populated_authored_model_behind_the_public_native_save_gate() {
        var options = new PublisherReadOptions();
        PublisherDocument document = TableDocument(options);
        PublisherNativeCreationPlan plan = PublisherNativeCreationProjection.Tables(document);
        byte[] generated = PublisherNativeCreationCodec.Create(plan, options, MaximumBytes, CancellationToken.None);
        Assert.Equal("Cell", Assert.Single(Assert.Single(PublisherDocument.Load(generated).Pages[0].Tables).Cells).Text);
        Assert.False(document.HasNativeSource);
        var failure = Assert.Throws<OfficeConversionException>(() => document.ToPubBytes(
            new PublisherSaveOptions { LossPolicy = OfficeConversionLossPolicy.Allow }));
        PublisherSaveReport report = Assert.IsType<PublisherSaveReport>(failure.Report);
        Assert.True(report.HasErrors); Assert.Equal(0, report.OutputByteCount);
        Assert.Contains(report.FidelityDiagnostics, finding => finding.Code == "PUB_NATIVE_SOURCE_REQUIRED"
            && finding.LossKind == OfficeConversionLossKind.Failure);
        Assert.False(document.HasNativeSource);
    }

    [Fact]
    public void Reduced_creation_limits_keep_semantic_pages_and_multiple_tabs_without_spending_source_records() {
        var options = new PublisherReadOptions {
            MaximumPages = 2,
            Limits = new OfficeLegacyImportLimits { MaxItems = 8, MaxTextCharacters = 64, MaxRecords = 1 }
        };
        PublisherDocument document = FrameDocument(options);
        byte[] generated = PublisherNativeCreationCodec.Create(PublisherNativeCreationProjection.FrameLayouts(document),
            options, MaximumBytes, CancellationToken.None);
        PublisherDocument decoded = PublisherDocument.Load(generated);
        Assert.Single(decoded.Pages); Assert.Single(decoded.MasterPages);
        OfficeRichTextParagraph paragraph = Assert.Single(Assert.Single(decoded.TextStories).Paragraphs);
        OfficeRichTextRun run = Assert.Single(paragraph.Runs);
        Assert.Equal("A\tB\tC", run.Text);
        Assert.Equal((12D, OfficeColor.Black, "Arial"), (run.FontSize, run.Color, run.FontFamily));
        Assert.Equal(18D, paragraph.LineHeight);
        OfficeTextTabStops tabs = Assert.IsType<OfficeTextTabStops>(paragraph.TabStops);
        Assert.Equal(new[] { 36D, 72D }, tabs.Stops.Select(stop => stop.Position));
        Assert.All(tabs.Stops, stop => Assert.Equal(OfficeTextTabAlignment.Left, stop.Alignment));
        Assert.Equal(36D, tabs.DefaultInterval); Assert.Equal(0D, tabs.Origin);
    }

    [Theory]
    [InlineData("pages")]
    [InlineData("text")]
    public void Creation_applies_stricter_semantic_page_and_text_budgets_to_an_admitted_model(string bound) {
        PublisherNativeCreationPlan plan = PublisherNativeCreationProjection.FrameLayouts(FrameDocument(new PublisherReadOptions()));
        var options = new PublisherReadOptions {
            MaximumPages = 2,
            Limits = new OfficeLegacyImportLimits { MaxTextCharacters = 5 }
        };
        byte[] generated = PublisherNativeCreationCodec.Create(plan, options, MaximumBytes, CancellationToken.None);
        PublisherDocument decoded = PublisherDocument.Load(generated);
        Assert.Single(decoded.Pages); Assert.Single(decoded.MasterPages);
        Assert.Equal("A\tB\tC", Assert.Single(Assert.Single(decoded.TextStories).Paragraphs).Runs.Single().Text);
        PublisherReadOptions smaller = options.Clone();
        if (bound == "pages") smaller.MaximumPages = 1;
        else smaller.Limits.MaxTextCharacters = 4;
        Assert.Throws<InvalidDataException>(() => PublisherNativeCreationCodec.Create(plan, smaller, MaximumBytes, CancellationToken.None));
        Assert.Equal(generated, PublisherNativeCreationCodec.Create(plan, options, MaximumBytes, CancellationToken.None));
    }

    [Theory]
    [InlineData("per-image")]
    [InlineData("total")]
    public void Shared_picture_creation_counts_one_encoded_asset_and_enforces_each_byte_budget(string bound) {
        byte[] image = PublisherPictureEditingFixture.Image(OfficeColor.FromRgb(255, 0, 0));
        PublisherDocument document = AuthoredDocument(new PublisherReadOptions());
        uint pageId = document.Pages[0].Id;
        PublisherPicture first = document.AddPicture(pageId, image, 72, 72, 40, 40);
        PublisherPicture second = document.AddPicture(pageId, first.ImageId!.Value, 144, 72, 40, 40);
        Assert.Equal(first.ImageId, second.ImageId);
        PublisherNativeCreationPlan plan = PublisherNativeCreationProjection.SingleMasterContent(document);
        var options = new PublisherReadOptions { MaximumImageBytes = image.Length, MaximumTotalImageBytes = image.Length };
        byte[] generated = PublisherNativeCreationCodec.Create(plan, options, MaximumBytes, CancellationToken.None);
        PublisherDocument decoded = PublisherDocument.Load(generated);
        PublisherImage asset = Assert.Single(decoded.Images);
        Assert.Equal(image, asset.GetBytes());
        IReadOnlyList<PublisherPicture> pictures = Assert.Single(decoded.Pages).Pictures;
        Assert.Equal(2, pictures.Count); Assert.All(pictures, picture => Assert.Equal(asset.Id, picture.ImageId));
        PublisherReadOptions smaller = options.Clone();
        if (bound == "per-image") smaller.MaximumImageBytes = image.Length - 1;
        else smaller.MaximumTotalImageBytes = image.Length - 1;
        Assert.Throws<InvalidDataException>(() => PublisherNativeCreationCodec.Create(plan, smaller, MaximumBytes, CancellationToken.None));
        Assert.Equal(generated, PublisherNativeCreationCodec.Create(plan, options, MaximumBytes, CancellationToken.None));
    }

    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public void Creation_marks_each_embedded_picture_placement_independently_of_names_and_shared_assets(bool includePictureNames) {
        byte[] image = PublisherPictureEditingFixture.Image(OfficeColor.Red);
        PublisherDocument document = AuthoredDocument(new PublisherReadOptions());
        PublisherPicture first = document.AddPicture(document.Pages[0].Id, image, 72, 144, 96, 96);
        document.AddPicture(document.Pages[0].Id, first.ImageId!.Value, 216, 144, 96, 96);
        PublisherPicture master = document.AddPicture(document.MasterPages[0].Id, first.ImageId.Value, 24, 24, 48, 48);
        var crop = new OfficeImageSourceCrop(.25, .125, .0625, .1875);
        document.SetPictureCrop(master.Id, crop);
        document.AddShape(document.MasterPages[0].Id, OfficeShape.Rectangle(40, 12), 72, 24);
        document.AddTextFrame(document.Pages[0].Id, 72, 72, 120, 24, new[] { Paragraph("Caption") });

        byte[] generated = PublisherNativeCreationCodec.Create(
            PublisherNativeCreationProjection.SingleMasterContent(document, includePictureNames),
            new PublisherReadOptions(), MaximumBytes, CancellationToken.None);
        PublisherDocument decoded = PublisherDocument.Load(generated);
        Assert.False(document.HasNativeSource);
        PublisherImage asset = Assert.Single(decoded.Images);
        Assert.Equal(image, asset.GetBytes());
        Assert.Equal(2, decoded.Pages[0].Pictures.Count);
        Assert.Equal(crop, Assert.Single(decoded.MasterPages[0].Pictures).Crop);
        PublisherPicture[] pictures = decoded.MasterPages.Concat(decoded.Pages).SelectMany(page => page.Pictures).ToArray();
        Assert.Equal(3, pictures.Length);
        Assert.All(pictures, picture => Assert.Equal(asset.Id, picture.ImageId));

        Assert.True(OfficeCompoundFileReader.TryRead(generated, out OfficeCompoundFile? compound, out string? error), error);
        byte[] contentsBytes = compound!.Streams["Contents"];
        var context = new PublisherParseContext(new PublisherReadOptions(), CancellationToken.None);
        PublisherContents contents = new PublisherContentsReader(new PublisherBinaryData(contentsBytes, "Contents"), context).Read();
        var pictureIds = new HashSet<uint>(pictures.Select(picture => picture.Id));
        foreach (PublisherPicture picture in pictures) {
            PublisherSourceShape source = contents.Shapes[picture.Id];
            Assert.Equal(1U, source.Chunk.Kind);
            // Native embedded-picture state is distinct from the managed reader's BLIP resolution.
            PublisherBlock marker = Assert.Single(source.Fields, field => field.Id == 3);
            Assert.Equal((byte)8, marker.Type);
            Assert.Equal(0, marker.Length);
            Assert.Equal(new byte[] { 3, 8 }, new[] { contentsBytes[marker.Offset - 2], contentsBytes[marker.Offset - 1] });
            Assert.Equal(includePictureNames, source.Value(0x3A).HasValue);
        }
        Assert.All(contents.Shapes.Values.Where(shape => !pictureIds.Contains(shape.Chunk.Id)),
            shape => Assert.DoesNotContain(shape.Fields, field => field.Id == 3));
    }

    [Fact]
    public void Table_utility_creation_uses_the_supplied_item_allowance() {
        var options = new PublisherReadOptions { Limits = new OfficeLegacyImportLimits { MaxItems = 7 } };
        PublisherDocument document = TableDocument(options);
        PublisherNativeCreationPlan plan = PublisherNativeCreationProjection.Tables(document);
        byte[] generated = PublisherNativeCreationCodec.Create(plan, options, MaximumBytes, CancellationToken.None);
        PublisherTable table = Assert.Single(PublisherDocument.Load(generated).Pages[0].Tables);
        Assert.Equal((1, 1, 1), (table.RowCount, table.ColumnCount, table.Cells.Count));
        Assert.Equal("Cell", table.Cells[0].Text);
        Assert.Equal(new[] { 120D }, table.ColumnWidths); Assert.Equal(new[] { 42D }, table.RowHeights);
        var smaller = new PublisherReadOptions { Limits = new OfficeLegacyImportLimits { MaxItems = 6 } };
        PublisherNativeCreationPlan bounded = PublisherNativeCreationProjection.Tables(TableDocument(smaller));
        Assert.Throws<NotSupportedException>(() => PublisherNativeCreationCodec.Create(bounded, smaller, MaximumBytes, CancellationToken.None));
        Assert.Equal(generated, PublisherNativeCreationCodec.Create(plan, options, MaximumBytes, CancellationToken.None));
    }

    [Fact]
    public void Creation_accepts_its_exact_compound_output_budget_and_rejects_one_byte_less() {
        var options = new PublisherReadOptions();
        PublisherNativeCreationPlan plan = PublisherNativeCreationProjection.FrameLayouts(FrameDocument(options));
        byte[] generated = PublisherNativeCreationCodec.Create(plan, options, MaximumBytes, CancellationToken.None);
        Assert.Equal("A\tB\tC", Assert.Single(Assert.Single(PublisherDocument.Load(generated).TextStories).Paragraphs).Runs.Single().Text);
        Assert.Equal(generated, PublisherNativeCreationCodec.Create(plan, options, generated.LongLength, CancellationToken.None));
        InvalidDataException failure = Assert.Throws<InvalidDataException>(() =>
            PublisherNativeCreationCodec.Create(plan, options, generated.LongLength - 1, CancellationToken.None));
        Assert.True(OfficeOutputLimit.Is(failure));
        Assert.Equal(generated, PublisherNativeCreationCodec.Create(plan, options, MaximumBytes, CancellationToken.None));
    }

    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public void Creation_observes_cancellation_before_entry_and_after_plan_access(bool afterEntry) {
        var options = new PublisherReadOptions();
        PublisherDocument document = FrameDocument(options);
        document.AddTextFrame(document.Pages[0].Id, 72, 216, 288, 108, new[] { Paragraph("Second frame") });
        PublisherNativeCreationPlan original = PublisherNativeCreationProjection.FrameLayouts(document);
        using var cancellation = new CancellationTokenSource();
        var objects = new CancelingObjects(original.Objects, cancellation);
        PublisherNativeCreationPlan plan = afterEntry ? original with { Objects = objects } : original;
        if (!afterEntry) cancellation.Cancel();
        OperationCanceledException failure = Assert.Throws<OperationCanceledException>(() =>
            PublisherNativeCreationCodec.Create(plan, options, MaximumBytes, cancellation.Token));
        Assert.Equal(cancellation.Token, failure.CancellationToken);
        Assert.Equal(afterEntry, objects.CancellationRequested);
        byte[] retry = PublisherNativeCreationCodec.Create(original, options, MaximumBytes, CancellationToken.None);
        Assert.Equal(new[] { "A\tB\tC", "Second frame" }, PublisherDocument.Load(retry).TextStories
            .Select(story => Assert.Single(story.Paragraphs).Runs.Single().Text));
    }

    private static PublisherDocument FrameDocument(PublisherReadOptions options) {
        PublisherDocument document = AuthoredDocument(options);
        var tabs = new OfficeTextTabStops(new[] { new OfficeTextTabStop(36), new OfficeTextTabStop(72) });
        document.AddTextFrame(document.Pages[0].Id, 72, 72, 288, 108,
            new[] { Paragraph("A\tB\tC").WithTabStops(tabs) });
        return document;
    }

    private static PublisherDocument TableDocument(PublisherReadOptions options) {
        PublisherDocument document = AuthoredDocument(options);
        PublisherTable table = document.AddTable(document.Pages[0].Id, 72, 144, new[] { 120D }, new[] { 42D });
        document.ReplaceTableCellParagraphs(table.Id, 0, 0, new[] { Paragraph("Cell") });
        return document;
    }

    private static PublisherDocument AuthoredDocument(PublisherReadOptions options) {
        PublisherDocument document = PublisherDocument.Create(612, 792, options);
        PublisherPage master = document.AddMasterPage("A");
        document.SetPageMaster(document.Pages[0].Id, master.Id);
        return document;
    }

    private static OfficeRichTextParagraph Paragraph(string text) => new(new[] {
        new OfficeRichTextRun(text, 12, OfficeColor.Black, fontFamily: "Arial")
    }, lineHeight: 18);

    /// <summary>Requests external cancellation during genuine collection access without throwing or changing items.</summary>
    private sealed class CancelingObjects : IReadOnlyList<PublisherNativeCreationObject> {
        private readonly IReadOnlyList<PublisherNativeCreationObject> _items;
        private readonly CancellationTokenSource _cancellation;
        internal CancelingObjects(IReadOnlyList<PublisherNativeCreationObject> items, CancellationTokenSource cancellation) {
            _items = items; _cancellation = cancellation;
        }
        internal bool CancellationRequested { get; private set; }
        public int Count => _items.Count;
        public PublisherNativeCreationObject this[int index] {
            get {
                PublisherNativeCreationObject item = _items[index];
                if (!CancellationRequested) { CancellationRequested = true; _cancellation.Cancel(); }
                return item;
            }
        }
        public IEnumerator<PublisherNativeCreationObject> GetEnumerator() {
            for (int index = 0; index < Count; index++) yield return this[index];
        }
        IEnumerator IEnumerable.GetEnumerator() => GetEnumerator();
    }
}
