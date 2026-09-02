using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using OfficeIMO.Excel;
using OfficeIMO.IWork;
using OfficeIMO.Word;

namespace OfficeIMO.IWork.Tests;

public sealed partial class IWorkBoundaryTests {
    [Fact]
    public void Pages_floating_canvas_rejects_fields_outside_the_page_group_envelope() {
        byte[] pageGroup = Message(BytesField(2,
            Message(ReferenceField(1, 4))));
        byte[] floating = Message(BytesField(1, pageGroup), VarintField(9, 1));
        byte[] records = Message(
            ArchiveRecord(1, 10000,
                Message(ReferenceField(4, 2), ReferenceField(3, 3))),
            ArchiveRecord(2, 2001, Message(StringField(3, "Body"))),
            ArchiveRecord(3, 10020, floating),
            ArchiveRecord(4, 2011, Message(ReferenceField(2, 5))),
            ArchiveRecord(5, 2001, Message(StringField(3, "Shape"))));
        using MemoryStream package = CreatePackage(
            ("Index/Document.iwa", FrameIwa(records)),
            ("preview.png", ValidPreviewPng()));

        using var result = WordIWorkConverter.ConvertPagesToWordResult(package);

        Assert.True(result.IsVisualFallback);
        Assert.Contains(result.Projection.Diagnostics, diagnostic =>
            diagnostic.Code == "IWORK_PAGES_DRAWABLE_UNSUPPORTED");
    }

    [Fact]
    public void Pages_sections_require_one_header_footer_template_reference() {
        byte[] sectionTable = Message(BytesField(1,
            Message(ReferenceField(2, 3))));
        byte[] records = Message(
            ArchiveRecord(1, 10000, Message(ReferenceField(4, 2))),
            ArchiveRecord(2, 2001,
                Message(StringField(3, "Body"), BytesField(17, sectionTable))),
            ArchiveRecord(3, 10011,
                Message(ReferenceField(25, 4), ReferenceField(25, 5))),
            ArchiveRecord(4, 10143, Message()),
            ArchiveRecord(5, 10143, Message()));
        using MemoryStream package = CreatePackage(
            ("Index/Document.iwa", FrameIwa(records)),
            ("preview.png", ValidPreviewPng()));

        using var result = WordIWorkConverter.ConvertPagesToWordResult(package);

        Assert.True(result.IsVisualFallback);
        Assert.Contains(result.Projection.Diagnostics, diagnostic =>
            diagnostic.Code == "IWORK_PAGES_HEADER_FOOTER_UNSUPPORTED");
    }

    [Fact]
    public void Numbers_rows_reject_excess_trailing_offsets_before_scanning_them() {
        using MemoryStream package = CreateNumbersPackage(new[] {
            new TableSpec("Offsets", 1, 1, 42d, trailingEmptyOffsetCount: 100_000)
        }, includePreview: true);

        using var result = ExcelIWorkConverter.ConvertNumbersToExcelResult(package);

        Assert.True(result.IsVisualFallback);
        Assert.Contains(result.Projection.Diagnostics, diagnostic =>
            diagnostic.Code == "IWORK_TABLE_CELL_STORAGE_UNSUPPORTED");
    }

    [Fact]
    public void Pages_images_are_anchored_to_their_source_page_paragraph() {
        const string imageName = "page-one.png";
        byte[] geometry = Message(
            BytesField(1, Message(FloatField(1, 36f), FloatField(2, 36f))),
            BytesField(2, Message(FloatField(1, 40f), FloatField(2, 30f))));
        byte[] image = Message(
            BytesField(1, Message(BytesField(1, geometry))),
            BytesField(11, Message(VarintField(1, 20))));
        byte[] metadata = Message(BytesField(4, Message(VarintField(1, 20),
            StringField(3, imageName), StringField(4, imageName))));
        byte[] firstPage = Message(BytesField(2,
            Message(ReferenceField(1, 4))));
        byte[] floating = Message(BytesField(1, firstPage), BytesField(1, Message()));
        byte[] records = Message(
            ArchiveRecord(1, 10000,
                Message(ReferenceField(4, 2), ReferenceField(3, 3))),
            ArchiveRecord(2, 2001, Message(StringField(3, "First\u000cSecond"))),
            ArchiveRecord(3, 10020, floating),
            ArchiveRecord(4, 3005, image),
            ArchiveRecord(5, 11006, metadata));
        using MemoryStream package = CreatePackage(
            ("Index/Document.iwa", FrameIwa(records)),
            ($"Data/{imageName}", ValidPreviewPng()),
            ("preview.png", ValidPreviewPng()));

        using var result = WordIWorkConverter.ConvertPagesToWordResult(package);
        using var saved = new MemoryStream();
        result.Value.Save(saved);
        saved.Position = 0;
        using WordprocessingDocument document = WordprocessingDocument.Open(saved, false);
        Body body = document.MainDocumentPart?.Document?.Body
            ?? throw new InvalidDataException("The reconstructed DOCX has no body.");
        Paragraph[] paragraphs = body.Elements<Paragraph>().ToArray();
        int imageParagraph = Array.FindIndex(paragraphs,
            paragraph => paragraph.Descendants<DocumentFormat.OpenXml.Drawing.Blip>().Any());
        int pageBreakParagraph = Array.FindIndex(paragraphs,
            paragraph => paragraph.Descendants<Break>()
                .Any(value => value.Type?.Value == BreakValues.Page));

        Assert.False(result.IsVisualFallback);
        Assert.Equal(1, Assert.Single(result.Projection.Drawables).PageIndex);
        Assert.True(imageParagraph >= 0);
        Assert.True(pageBreakParagraph >= 0);
        Assert.True(imageParagraph < pageBreakParagraph);
    }
}
