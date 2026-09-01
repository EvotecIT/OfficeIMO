using System.Text;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using OfficeIMO.Excel;
using OfficeIMO.IWork;
using OfficeIMO.IWork.Internal;
using OfficeIMO.PowerPoint;
using OfficeIMO.Word;
using DrawingWordprocessing = DocumentFormat.OpenXml.Drawing.Wordprocessing;

namespace OfficeIMO.IWork.Tests;

public sealed partial class IWorkBoundaryTests {
    [Fact]
    public void Masked_images_disable_editable_reconstruction() {
        using MemoryStream package = CreatePagesPackageWithImage(masked: true, left: 0, rotation: 0);

        using var result = WordDocument.LoadPagesWithReport(package);

        Assert.True(result.IsVisualFallback);
        Assert.Contains(result.Projection.Diagnostics,
            diagnostic => diagnostic.Code == "IWORK_PAGES_IMAGE_UNSUPPORTED");
    }

    [Theory]
    [InlineData(0f)]
    [InlineData(12f)]
    public void Pages_images_preserve_page_relative_position(float left) {
        using MemoryStream package = CreatePagesPackageWithImage(masked: false, left: left, rotation: 0);

        using var result = WordDocument.LoadPagesWithReport(package);

        Assert.False(result.IsVisualFallback);
        Assert.True(result.Projection.HasEditableContent);
        using var saved = new MemoryStream();
        result.Document.Save(saved);
        saved.Position = 0;
        using WordprocessingDocument document = WordprocessingDocument.Open(saved, false);
        DrawingWordprocessing.Anchor anchor = Assert.Single(document.MainDocumentPart?.Document?
            .Descendants<DrawingWordprocessing.Anchor>()
            ?? throw new InvalidDataException("The reconstructed DOCX has no main document."));
        Assert.Equal(DrawingWordprocessing.HorizontalRelativePositionValues.Page,
            anchor.HorizontalPosition?.RelativeFrom?.Value);
        Assert.Equal(DrawingWordprocessing.VerticalRelativePositionValues.Page,
            anchor.VerticalPosition?.RelativeFrom?.Value);
        Assert.Equal((left * 12700f).ToString(System.Globalization.CultureInfo.InvariantCulture),
            anchor.HorizontalPosition?.PositionOffset?.Text);
        Assert.Equal("0", anchor.VerticalPosition?.PositionOffset?.Text);
    }

    [Theory]
    [InlineData("a.", "lowerLetter")]
    [InlineData("i.", "lowerRoman")]
    [InlineData("iv.", "lowerRoman")]
    public void Alphabetic_and_roman_pages_lists_use_native_word_numbering(
        string label, string expectedFormat) {
        using MemoryStream package = CreatePagesPackageWithListLabel(label);
        using var result = WordDocument.LoadPagesWithReport(package);
        using var saved = new MemoryStream();
        result.Document.Save(saved);
        saved.Position = 0;

        using WordprocessingDocument document = WordprocessingDocument.Open(saved, false);
        Paragraph paragraph = document.MainDocumentPart?.Document?.Body?.Elements<Paragraph>()
            .Single(candidate => candidate.InnerText == "Item")
            ?? throw new InvalidDataException("The reconstructed DOCX has no list paragraph.");
        int numberId = paragraph.ParagraphProperties?.NumberingProperties?.NumberingId?.Val?.Value
            ?? throw new InvalidDataException("The reconstructed paragraph has no numbering identifier.");
        Numbering numbering = document.MainDocumentPart?.NumberingDefinitionsPart?.Numbering
            ?? throw new InvalidDataException("The reconstructed DOCX has no numbering definitions.");
        int abstractId = numbering.Elements<NumberingInstance>()
            .Single(instance => instance.NumberID?.Value == numberId)
            .AbstractNumId?.Val?.Value
            ?? throw new InvalidDataException("The numbering instance has no abstract definition.");
        NumberFormatValues? format = numbering.Elements<AbstractNum>()
            .Single(item => item.AbstractNumberId?.Value == abstractId)
            .Elements<Level>().Single(level => level.LevelIndex?.Value == 0)
            .NumberingFormat?.Val?.Value;

        NumberFormatValues expected = expectedFormat == "lowerLetter"
            ? NumberFormatValues.LowerLetter
            : NumberFormatValues.LowerRoman;
        Assert.Equal(expected, format);
    }

    [Fact]
    public void Transparent_pages_text_uses_visual_fallback() {
        using MemoryStream package = CreatePagesPackageWithTransparentText();

        using var result = WordDocument.LoadPagesWithReport(package);

        Assert.True(result.IsVisualFallback);
        Assert.True(result.Projection.HasEditableContent);
    }

    [Fact]
    public void Transparent_keynote_text_uses_visual_fallback() {
        using MemoryStream package = CreateKeynotePackageWithTransparentText();

        using var result = PowerPointPresentation.LoadKeynoteWithReport(package);

        Assert.True(result.IsVisualFallback);
        Assert.True(result.Projection.HasEditableContent);
    }

    [Fact]
    public void Classic_pdf_page_tree_children_are_validated() {
        Assert.True(IWorkPdfInfo.IsComplete(CreateOnePageClassicPdf(validKids: true)));
        Assert.False(IWorkPdfInfo.IsComplete(CreateOnePageClassicPdf(validKids: false)));
    }

    [Fact]
    public void Unknown_numbers_drawables_disable_editable_reconstruction() {
        byte[] records = Message(
            ArchiveRecord(1, 1, Message(ReferenceField(1, 2))),
            ArchiveRecord(2, 2, Message(StringField(1, "Sheet"), ReferenceField(2, 3))),
            ArchiveRecord(3, 7777, Message()),
            ArchiveRecord(4, 6001, Message()));
        using MemoryStream package = CreatePackage(
            ("Index/Document.iwa", FrameIwa(records)),
            ("preview.png", ValidPreviewPng()));

        using var result = ExcelDocument.LoadNumbersWithReport(package);

        Assert.True(result.IsVisualFallback);
        Assert.Contains(result.Projection.Diagnostics,
            diagnostic => diagnostic.Code == "IWORK_NUMBERS_DRAWABLE_UNSUPPORTED");
    }

    [Fact]
    public void Every_pages_header_row_repeats_in_word() {
        using MemoryStream package = CreatePagesPackageWithHeaderRows(2);
        using var result = WordDocument.LoadPagesWithReport(package);
        using var saved = new MemoryStream();
        result.Document.Save(saved);
        saved.Position = 0;

        using WordprocessingDocument document = WordprocessingDocument.Open(saved, false);
        TableRow[] rows = document.MainDocumentPart?.Document?.Body?
            .Descendants<TableRow>().ToArray()
            ?? throw new InvalidDataException("The reconstructed DOCX has no table rows.");

        Assert.NotNull(rows[0].TableRowProperties?.GetFirstChild<TableHeader>());
        Assert.NotNull(rows[1].TableRowProperties?.GetFirstChild<TableHeader>());
        Assert.Null(rows[2].TableRowProperties?.GetFirstChild<TableHeader>());
    }

    [Fact]
    public void Pages_margins_must_leave_a_positive_content_box() {
        byte[] layout = Message(
            FloatField(30, 100f), FloatField(31, 100f),
            FloatField(32, 60f), FloatField(33, 40f),
            FloatField(34, 10f), FloatField(35, 10f));
        using MemoryStream package = CreatePagesPackage(includeBody: true, textBox: null,
            includePreview: true, documentLayoutFields: layout);

        using var result = WordDocument.LoadPagesWithReport(package);

        Assert.True(result.IsVisualFallback);
        Assert.True(result.Projection.HasEditableContent);
    }

    private static MemoryStream CreatePagesPackageWithImage(bool masked, float left, float rotation,
        float width = 40f, float height = 30f) {
        const ulong documentId = 1;
        const ulong bodyId = 2;
        const ulong imageId = 3;
        const ulong metadataId = 4;
        const ulong dataId = 10;
        const string name = "image.png";
        byte[] geometry = Message(
            BytesField(1, Message(FloatField(1, left), FloatField(2, 0f))),
            BytesField(2, Message(FloatField(1, width), FloatField(2, height))),
            FloatField(4, rotation));
        byte[] image = Message(
            BytesField(1, Message(BytesField(1, geometry))),
            BytesField(11, Message(VarintField(1, dataId))),
            masked ? BytesField(5, Message()) : Array.Empty<byte>());
        byte[] metadataEntry = Message(VarintField(1, dataId),
            StringField(3, name), StringField(4, name));
        byte[] records = Message(
            ArchiveRecord(documentId, 10000, Message(ReferenceField(4, bodyId)),
                new[] { bodyId, imageId }),
            ArchiveRecord(bodyId, 2001, Message(StringField(3, "Body"))),
            ArchiveRecord(imageId, 3005, image),
            ArchiveRecord(metadataId, 11006, Message(BytesField(4, metadataEntry))));
        return CreatePackage(
            ("Index/Document.iwa", FrameIwa(records)),
            ($"Data/{name}", ValidPreviewPng()),
            ("preview.png", ValidPreviewPng()));
    }

    private static MemoryStream CreatePagesPackageWithListLabel(string label,
        bool includePreview = false) {
        const ulong documentId = 1;
        const ulong bodyId = 2;
        const ulong listId = 3;
        byte[] listTable = Message(BytesField(1,
            Message(VarintField(1, 0), ReferenceField(2, listId))));
        byte[] records = Message(
            ArchiveRecord(documentId, 10000, Message(ReferenceField(4, bodyId)), new[] { bodyId }),
            ArchiveRecord(bodyId, 2001,
                Message(StringField(3, "Item"), BytesField(7, listTable)), new[] { listId }),
            ArchiveRecord(listId, 2023, Message(VarintField(11, 1), StringField(16, label))));
        return includePreview
            ? CreatePackage(("Index/Document.iwa", FrameIwa(records)),
                ("preview.png", ValidPreviewPng()))
            : CreatePackage(("Index/Document.iwa", FrameIwa(records)));
    }

    private static MemoryStream CreatePagesPackageWithTransparentText() {
        const ulong documentId = 1;
        const ulong bodyId = 2;
        const ulong styleId = 3;
        byte[] styleTable = Message(BytesField(1,
            Message(VarintField(1, 0), ReferenceField(2, styleId))));
        byte[] color = Message(FloatField(3, 1f), FloatField(4, 0f),
            FloatField(5, 0f), FloatField(6, 0.5f));
        byte[] records = Message(
            ArchiveRecord(documentId, 10000, Message(ReferenceField(4, bodyId)), new[] { bodyId }),
            ArchiveRecord(bodyId, 2001,
                Message(StringField(3, "Color"), BytesField(8, styleTable)), new[] { styleId }),
            ArchiveRecord(styleId, 2021, Message(BytesField(11, Message(BytesField(7, color))))));
        return CreatePackage(
            ("Index/Document.iwa", FrameIwa(records)),
            ("preview.png", ValidPreviewPng()));
    }

    private static MemoryStream CreateKeynotePackageWithTransparentText() {
        const ulong documentId = 1;
        const ulong showId = 2;
        const ulong nodeId = 3;
        const ulong slideId = 4;
        const ulong shapeId = 5;
        const ulong storageId = 6;
        const ulong styleId = 7;
        byte[] styleTable = Message(BytesField(1,
            Message(VarintField(1, 0), ReferenceField(2, styleId))));
        byte[] color = Message(FloatField(3, 0f), FloatField(4, 0f),
            FloatField(5, 1f), FloatField(6, 0.5f));
        byte[] records = Message(
            ArchiveRecord(documentId, 1, Message(ReferenceField(2, showId))),
            ArchiveRecord(showId, 2, Message(BytesField(3, Message(ReferenceField(2, nodeId))))),
            ArchiveRecord(nodeId, 4, Message(ReferenceField(2, slideId))),
            ArchiveRecord(slideId, 5, Message(ReferenceField(5, shapeId))),
            ArchiveRecord(shapeId, 2011, Message(ReferenceField(2, storageId))),
            ArchiveRecord(storageId, 2001,
                Message(StringField(3, "Color"), BytesField(8, styleTable)), new[] { styleId }),
            ArchiveRecord(styleId, 2021, Message(BytesField(11, Message(BytesField(7, color))))));
        return CreatePackage(
            ("Index/Slide.iwa", FrameIwa(records)),
            ("preview.png", ValidPreviewPng()));
    }

    private static byte[] CreateOnePageClassicPdf(bool validKids,
        string pageDictionaryPrefix = "", string trailerDictionaryPrefix = "",
        bool omitCatalogEndObject = false, bool omitPagesEndObject = false,
        bool omitPageEndObject = false, bool trailCatalogDictionary = false,
        bool trailPagesDictionary = false, bool trailPageDictionary = false,
        bool commentCatalogDictionary = false, bool omitMediaBox = false,
        string trailerPrefix = "", string trailerSuffix = "") {
        const string header = "%PDF-1.4\n";
        string catalog = "1 0 obj\n<< /Type /Catalog /Pages 2 0 R >>\n"
            + (trailCatalogDictionary ? "42\n" : string.Empty)
            + (commentCatalogDictionary ? "% valid comment\n" : string.Empty)
            + (omitCatalogEndObject ? string.Empty : "endobj\n");
        string pages = validKids
            ? "2 0 obj\n<< /Type /Pages " + (omitMediaBox ? string.Empty : "/MediaBox [0 0 612 792] ") + "/Count 1 /Kids [3 0 R] >>\n"
            : "2 0 obj\n<< /Type /Pages " + (omitMediaBox ? string.Empty : "/MediaBox [0 0 612 792] ") + "/Count 1 /Kids [] >>\n";
        if (trailPagesDictionary) pages += "42\n";
        if (!omitPagesEndObject) pages += "endobj\n";
        string page = "3 0 obj\n<< /Type /Page " + pageDictionaryPrefix
            + "/Parent 2 0 R >>\n"
            + (trailPageDictionary ? "42\n" : string.Empty)
            + (omitPageEndObject ? string.Empty : "endobj\n");
        int catalogOffset = Encoding.ASCII.GetByteCount(header);
        int pagesOffset = Encoding.ASCII.GetByteCount(header + catalog);
        int pageOffset = Encoding.ASCII.GetByteCount(header + catalog + pages);
        string prefix = header + catalog + pages + page;
        int xrefOffset = Encoding.ASCII.GetByteCount(prefix);
        string suffix = "xref\n0 4\n0000000000 65535 f \n"
            + catalogOffset.ToString("D10", System.Globalization.CultureInfo.InvariantCulture) + " 00000 n \n"
            + pagesOffset.ToString("D10", System.Globalization.CultureInfo.InvariantCulture) + " 00000 n \n"
            + pageOffset.ToString("D10", System.Globalization.CultureInfo.InvariantCulture) + " 00000 n \n"
            + "trailer\n" + trailerPrefix + "<< " + trailerDictionaryPrefix
            + "/Size 4 /Root 1 0 R >>" + trailerSuffix + "\nstartxref\n"
            + xrefOffset.ToString(System.Globalization.CultureInfo.InvariantCulture)
            + "\n%%EOF\n";
        return Encoding.ASCII.GetBytes(prefix + suffix);
    }

    private static MemoryStream CreatePagesPackageWithHeaderRows(int headerRows) {
        const ulong documentId = 1;
        const ulong bodyId = 2;
        const ulong tableId = 3;
        const ulong modelId = 4;
        byte[] model = Message(
            BytesField(4, Message(BytesField(3, Message()))),
            VarintField(6, 3), VarintField(7, 1),
            VarintField(9, checked((ulong)headerRows)), StringField(8, "Headers"));
        byte[] records = Message(
            ArchiveRecord(documentId, 10000, Message(ReferenceField(4, bodyId)),
                new[] { bodyId, tableId }),
            ArchiveRecord(bodyId, 2001, Message(StringField(3, "Body"))),
            ArchiveRecord(tableId, 6000, Message(ReferenceField(2, modelId)), new[] { modelId }),
            ArchiveRecord(modelId, 6001, model));
        return CreatePackage(("Index/Document.iwa", FrameIwa(records)));
    }
}
