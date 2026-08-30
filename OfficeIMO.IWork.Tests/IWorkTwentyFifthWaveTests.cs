using System.Text;
using OfficeIMO.Excel;
using OfficeIMO.IWork;
using OfficeIMO.IWork.Internal;
using OfficeIMO.PowerPoint;

namespace OfficeIMO.IWork.Tests;

public sealed partial class IWorkBoundaryTests {
    [Fact]
    public void Keynote_owner_preserves_cross_type_drawable_order() {
        using MemoryStream package = CreateKeynotePackageWithImageBeforeTable();

        using var result = PowerPointPresentation.LoadKeynoteWithReport(package);
        IWorkKeynoteSlide sourceSlide = Assert.Single(result.Projection.Slides);
        PowerPointSlide targetSlide = Assert.Single(result.Document.Slides);

        Assert.True(result.Projection.HasEditableContent,
            string.Join("; ", result.Projection.Diagnostics.Select(diagnostic =>
                $"{diagnostic.Code}: {diagnostic.Message}")));
        Assert.False(result.IsVisualFallback);
        Assert.Equal(new[] { IWorkKeynoteDrawableKind.Image, IWorkKeynoteDrawableKind.Table },
            sourceSlide.Drawables.Select(drawable => drawable.Kind));
        Assert.IsType<PowerPointPicture>(targetSlide.Shapes[0]);
        Assert.IsType<PowerPointTable>(targetSlide.Shapes[1]);

        using var saved = new MemoryStream();
        result.Document.Save(saved);
        saved.Position = 0;
        using PowerPointPresentation reopened = PowerPointPresentation.Load(saved);
        Assert.IsType<PowerPointPicture>(Assert.Single(reopened.Slides).Shapes[0]);
        Assert.IsType<PowerPointTable>(Assert.Single(reopened.Slides).Shapes[1]);
    }

    [Fact]
    public void Keynote_placeholder_fallback_scales_to_the_recovered_canvas() {
        using MemoryStream package = CreateKeynotePackageWithCanvasTitle(720f, 540f);

        using var result = PowerPointPresentation.LoadKeynoteWithReport(package);
        PowerPointTextBox title = Assert.Single(Assert.Single(result.Document.Slides).TextBoxes);

        Assert.False(result.IsVisualFallback);
        Assert.Equal(720d, result.Document.SlideSize.WidthPoints, 3);
        Assert.True(title.LeftPoints >= 0);
        Assert.True(title.TopPoints >= 0);
        Assert.True(title.RightPoints <= 720d);
        Assert.True(title.BottomPoints <= 540d);
        Assert.Equal(648d, title.WidthPoints, 3);
    }

    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public void Numbers_error_values_map_to_native_excel_errors(bool cachedFormula) {
        using MemoryStream package = CreateNumbersPackage(new[] {
            new TableSpec("Errors", 1, 1, 0d, hasFormula: cachedFormula, error: true)
        });

        using var result = ExcelDocument.LoadNumbersWithReport(package);
        ExcelSheet sheet = Assert.Single(result.Document.Sheets);

        Assert.False(result.IsVisualFallback);
        Assert.True(sheet.TryGetCellValueSnapshot(1, 1, out ExcelCellValueSnapshot? snapshot));
        Assert.Equal(ExcelCellValueKind.Error, snapshot!.Kind);
        Assert.Equal("#VALUE!", snapshot.RawValue);

        using var saved = new MemoryStream();
        result.Document.Save(saved);
        saved.Position = 0;
        using ExcelDocument reopened = ExcelDocument.Load(saved);
        Assert.True(Assert.Single(reopened.Sheets).TryGetCellValueSnapshot(
            1, 1, out ExcelCellValueSnapshot? persisted));
        Assert.Equal(ExcelCellValueKind.Error, persisted!.Kind);
    }

    [Fact]
    public void Classic_pdf_incremental_xref_chains_are_followed() {
        Assert.True(IWorkPdfInfo.IsComplete(CreateIncrementalPdf()));
    }

    [Fact]
    public void Classic_pdf_incremental_xref_cycles_are_rejected() {
        Assert.False(IWorkPdfInfo.IsComplete(CreateIncrementalPdf(selfReferentialPrevious: true)));
    }

    [Fact]
    public void Classic_pdf_newer_free_entries_do_not_resurrect_older_objects() {
        Assert.False(IWorkPdfInfo.IsComplete(CreateIncrementalPdf(freeRoot: true)));
    }

    private static MemoryStream CreateKeynotePackageWithImageBeforeTable() {
        const ulong documentId = 1;
        const ulong showId = 2;
        const ulong nodeId = 3;
        const ulong slideId = 4;
        const ulong imageId = 10;
        const ulong tableId = 11;
        const ulong modelId = 12;
        const ulong metadataId = 13;
        const ulong dataId = 20;
        const string imageName = "stacked.png";
        byte[] geometry = Message(
            BytesField(1, Message(FloatField(1, 72f), FloatField(2, 72f))),
            BytesField(2, Message(FloatField(1, 144f), FloatField(2, 96f))));
        byte[] image = Message(
            BytesField(1, Message(BytesField(1, geometry))),
            BytesField(11, Message(VarintField(1, dataId))));
        byte[] table = Message(
            BytesField(1, Message(BytesField(1, geometry))),
            ReferenceField(2, modelId));
        byte[] model = Message(BytesField(4, Message(BytesField(3, Message()))),
            VarintField(6, 1), VarintField(7, 1), StringField(8, "Stacked table"));
        byte[] metadata = Message(BytesField(4, Message(VarintField(1, dataId),
            StringField(3, imageName), StringField(4, imageName))));
        byte[] records = Message(
            ArchiveRecord(documentId, 1, Message(ReferenceField(2, showId))),
            ArchiveRecord(showId, 2,
                Message(BytesField(3, Message(ReferenceField(2, nodeId))))),
            ArchiveRecord(nodeId, 4, Message(ReferenceField(2, slideId))),
            ArchiveRecord(slideId, 5,
                Message(ReferenceField(7, imageId), ReferenceField(7, tableId)),
                new[] { imageId, tableId }),
            ArchiveRecord(imageId, 3005, image),
            ArchiveRecord(tableId, 6000, table, new[] { modelId }),
            ArchiveRecord(modelId, 6001, model),
            ArchiveRecord(metadataId, 11006, metadata));
        return CreatePackage(("Index/Slide.iwa", FrameIwa(records)),
            ($"Data/{imageName}", ValidPreviewPng()),
            ("preview.png", ValidPreviewPng()));
    }

    private static MemoryStream CreateKeynotePackageWithCanvasTitle(float width, float height) {
        const ulong documentId = 1;
        const ulong showId = 2;
        const ulong nodeId = 3;
        const ulong slideId = 4;
        const ulong shapeId = 5;
        const ulong storageId = 6;
        byte[] records = Message(
            ArchiveRecord(documentId, 1, Message(ReferenceField(2, showId))),
            ArchiveRecord(showId, 2,
                Message(BytesField(3, Message(ReferenceField(2, nodeId))),
                    BytesField(4, Message(FloatField(1, width), FloatField(2, height))))),
            ArchiveRecord(nodeId, 4, Message(ReferenceField(2, slideId))),
            ArchiveRecord(slideId, 5, Message(ReferenceField(5, shapeId))),
            ArchiveRecord(shapeId, 2011, Message(ReferenceField(2, storageId))),
            ArchiveRecord(storageId, 2001, Message(StringField(3, "Scaled title"))));
        return CreatePackage(("Index/Slide.iwa", FrameIwa(records)));
    }

    private static byte[] CreateIncrementalPdf(
        bool selfReferentialPrevious = false, bool freeRoot = false) {
        const string header = "%PDF-1.4\n";
        const string catalog = "1 0 obj\n<< /Type /Catalog /Pages 2 0 R >>\nendobj\n";
        const string pages = "2 0 obj\n<< /Type /Pages /Count 1 /Kids [3 0 R] >>\nendobj\n";
        const string page = "3 0 obj\n<< /Type /Page /Parent 2 0 R >>\nendobj\n";
        int catalogOffset = Encoding.ASCII.GetByteCount(header);
        int pagesOffset = Encoding.ASCII.GetByteCount(header + catalog);
        int pageOffset = Encoding.ASCII.GetByteCount(header + catalog + pages);
        string baseObjects = header + catalog + pages + page;
        int baseXrefOffset = Encoding.ASCII.GetByteCount(baseObjects);
        string baseRevision = baseObjects + "xref\n0 4\n0000000000 65535 f \n"
            + catalogOffset.ToString("D10", System.Globalization.CultureInfo.InvariantCulture) + " 00000 n \n"
            + pagesOffset.ToString("D10", System.Globalization.CultureInfo.InvariantCulture) + " 00000 n \n"
            + pageOffset.ToString("D10", System.Globalization.CultureInfo.InvariantCulture) + " 00000 n \n"
            + "trailer\n<< /Size 4 /Root 1 0 R >>\nstartxref\n"
            + baseXrefOffset.ToString(System.Globalization.CultureInfo.InvariantCulture)
            + "\n%%EOF\n";
        const string updateObject = "4 0 obj\n<< /Producer (OfficeIMO test) >>\nendobj\n";
        int updateObjectOffset = Encoding.ASCII.GetByteCount(baseRevision);
        int latestXrefOffset = Encoding.ASCII.GetByteCount(baseRevision + updateObject);
        long previous = selfReferentialPrevious ? latestXrefOffset : baseXrefOffset;
        string latestEntry = freeRoot
            ? "1 1\n0000000000 00001 f \n"
            : "4 1\n" + updateObjectOffset.ToString("D10",
                System.Globalization.CultureInfo.InvariantCulture) + " 00000 n \n";
        string latestRevision = "xref\n" + latestEntry
            + "trailer\n<< /Size 5 /Root 1 0 R /Prev "
            + previous.ToString(System.Globalization.CultureInfo.InvariantCulture)
            + " >>\nstartxref\n"
            + latestXrefOffset.ToString(System.Globalization.CultureInfo.InvariantCulture)
            + "\n%%EOF\n";
        return Encoding.ASCII.GetBytes(baseRevision + updateObject + latestRevision);
    }
}
