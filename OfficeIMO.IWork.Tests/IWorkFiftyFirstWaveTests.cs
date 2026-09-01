using OfficeIMO.IWork;
using OfficeIMO.IWork.Internal;
using OfficeIMO.PowerPoint;
using OfficeIMO.Word;

namespace OfficeIMO.IWork.Tests;

public sealed partial class IWorkBoundaryTests {
    [Fact]
    public void Keynote_nodes_with_multiple_slide_references_are_incomplete() {
        byte[] records = Message(
            ArchiveRecord(1, 1, Message(ReferenceField(2, 2))),
            ArchiveRecord(2, 2, KeynoteShow(Message(ReferenceField(2, 3)))),
            ArchiveRecord(3, 4, Message(ReferenceField(2, 4), ReferenceField(2, 5))),
            ArchiveRecord(4, 5, Message()),
            ArchiveRecord(5, 5, Message()));
        using MemoryStream package = CreatePackage(
            ("Index/Slide.iwa", FrameIwa(records)),
            ("preview.png", ValidPreviewPng()));

        IWorkKeynoteProjection projection = IWorkSourceDocument.Open(package,
            IWorkDocumentKind.Keynote).ReadKeynote();

        Assert.False(projection.HasEditableContent);
        Assert.Contains(projection.Diagnostics, diagnostic =>
            diagnostic.Code == "IWORK_KEYNOTE_SLIDE_MISSING");
    }

    [Fact]
    public void Table_infos_with_multiple_model_references_are_incomplete() {
        byte[] emptyStore = Message(BytesField(3, Message()));
        byte[] model = Message(BytesField(4, emptyStore), VarintField(6, 1),
            VarintField(7, 1), StringField(8, "Table"));
        byte[] records = Message(
            ArchiveRecord(1, 1, Message(ReferenceField(1, 2))),
            ArchiveRecord(2, 2, Message(StringField(1, "Sheet"), ReferenceField(2, 3))),
            ArchiveRecord(3, 6000, Message(ReferenceField(2, 4), ReferenceField(2, 5))),
            ArchiveRecord(4, 6001, model),
            ArchiveRecord(5, 6001, model));
        using MemoryStream package = CreatePackage(
            ("Index/Document.iwa", FrameIwa(records)),
            ("preview.png", ValidPreviewPng()));

        IWorkNumbersProjection projection = IWorkSourceDocument.Open(package,
            IWorkDocumentKind.Numbers).ReadNumbers();

        Assert.False(projection.HasEditableContent);
        Assert.Contains(projection.Diagnostics, diagnostic =>
            diagnostic.Code == "IWORK_TABLE_MODEL_UNSUPPORTED");
    }

    [Theory]
    [InlineData(3)]
    [InlineData(4)]
    public void Formula_range_tracts_reject_repeated_coordinate_ranges(int field) {
        byte[] coordinate = Message(VarintField(1, 0), VarintField(2, 1));
        byte[] tract = Message(
            BytesField(3, coordinate), BytesField(4, coordinate),
            BytesField(field, coordinate));
        byte[] node = Message(VarintField(1, 67), BytesField(40, tract));
        IWorkWireMessage formula = IWorkProtobuf.Parse(
            Message(BytesField(1, Message(BytesField(1, node)))),
            new IWorkReadOptions());

        Assert.False(IWorkFormulaReader.TryReadAbsoluteRange(formula, 10,
            out _, out _, out _, out _));
    }

    [Fact]
    public void Covered_merge_detection_scans_sparse_cells_not_dense_coordinates() {
        var table = new IWorkTable("Sparse", 1_048_576, 16_384,
            new[] { new IWorkTableCell(1, 1, IWorkCellKind.Number, 1d) },
            mergedRanges: new[] {
                new IWorkTableMergeRange(1, 1, 1_048_576, 16_384)
            });

        Assert.False(table.HasPopulatedCoveredMergeCells());
    }

    [Fact]
    public void Geometry_free_keynote_images_fit_inside_the_slide() {
        const string imageName = "large.png";
        byte[] image = Message(BytesField(1, Message()),
            BytesField(11, Message(VarintField(1, 20))));
        byte[] metadata = Message(BytesField(4, Message(VarintField(1, 20),
            StringField(3, imageName), StringField(4, imageName))));
        byte[] records = Message(
            ArchiveRecord(1, 1, Message(ReferenceField(2, 2))),
            ArchiveRecord(2, 2, Message(
                BytesField(3, Message(ReferenceField(2, 3))),
                BytesField(4, Message(FloatField(1, 960f), FloatField(2, 540f))))),
            ArchiveRecord(3, 4, Message(ReferenceField(2, 4))),
            ArchiveRecord(4, 5, Message(ReferenceField(7, 5))),
            ArchiveRecord(5, 3005, image),
            ArchiveRecord(6, 11006, metadata));
        using MemoryStream package = CreatePackage(
            ("Index/Slide.iwa", FrameIwa(records)),
            ($"Data/{imageName}", CreateSizedPreviewPng(2400, 1200)),
            ("preview.png", ValidPreviewPng()));

        using var result = PowerPointPresentation.LoadKeynoteWithReport(package);
        PowerPointPicture picture = Assert.Single(Assert.Single(result.Document.Slides).Pictures);

        Assert.False(result.IsVisualFallback, string.Join("; ",
            result.Projection.Diagnostics.Select(diagnostic =>
                $"{diagnostic.Code}: {diagnostic.Message}")));
        Assert.True(picture.RightPoints <= result.Document.SlideSize.WidthPoints);
        Assert.True(picture.BottomPoints <= result.Document.SlideSize.HeightPoints);
        Assert.Equal(2d, picture.WidthPoints / picture.HeightPoints, 6);
    }

    [Theory]
    [InlineData(true)]
    [InlineData(false)]
    public void Off_page_header_footer_distances_use_visual_fallback(bool header) {
        byte[] document = Message(
            ReferenceField(4, 2),
            FloatField(30, 612f), FloatField(31, 792f),
            FloatField(32, 72f), FloatField(33, 72f),
            FloatField(34, 72f), FloatField(35, 72f),
            FloatField(36, header ? 800f : 36f),
            FloatField(37, header ? 36f : 800f));
        byte[] records = Message(
            ArchiveRecord(1, 10000, document),
            ArchiveRecord(2, 2001, Message(StringField(3, "Body"))));
        using MemoryStream package = CreatePackage(
            ("Index/Document.iwa", FrameIwa(records)),
            ("preview.png", ValidPreviewPng()));

        using var result = WordDocument.LoadPagesWithReport(package);

        Assert.True(result.Projection.HasEditableContent);
        Assert.True(result.IsVisualFallback);
    }
}
