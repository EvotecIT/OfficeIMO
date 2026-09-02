using System.Text;
using OfficeIMO.Excel;
using OfficeIMO.IWork;
using OfficeIMO.IWork.Internal;
using OfficeIMO.PowerPoint;
using OfficeIMO.Word;

namespace OfficeIMO.IWork.Tests;

public sealed partial class IWorkBoundaryTests {
    [Fact]
    public void Primary_record_messages_are_cached_by_record_identity() {
        using MemoryStream package = CreatePagesPackage(
            includeBody: true, textBox: null, includePreview: false);
        IWorkSourceDocument source = IWorkSourceDocument.Open(
            package, IWorkDocumentKind.Pages);
        IWorkArchiveRecord document = Assert.Single(source.Records,
            record => record.MessageType == 10000);

        IWorkWireMessage first = source.Index.Message(document);
        IWorkWireMessage second = source.Index.Message(document);

        Assert.Same(first, second);
    }

    [Fact]
    public void Repeated_distinct_Keynote_body_placeholders_disable_editable_reconstruction() {
        using MemoryStream package = CreateKeynotePackageWithDistinctBodyPlaceholders();

        using var result = PowerPointIWorkConverter.ConvertKeynoteToPowerPointResult(package);

        Assert.True(result.IsVisualFallback);
        Assert.Contains(result.Projection.Diagnostics,
            diagnostic => diagnostic.Code == "IWORK_KEYNOTE_DRAWABLE_UNSUPPORTED");
    }

    [Fact]
    public void Numbers_numeric_cells_reject_conflicting_value_representations() {
        using MemoryStream package = CreateNumbersPackage(new[] {
            new TableSpec("Conflicting number", 1, 1, 2d,
                conflictingNumberValue: true)
        }, includePreview: true);

        using var result = ExcelIWorkConverter.ConvertNumbersToExcelResult(package);
        IWorkTableCell cell = Assert.Single(Assert.Single(
            Assert.Single(result.Projection.Sheets).Tables).Cells);

        Assert.True(result.IsVisualFallback);
        Assert.Equal(IWorkCellKind.Error, cell.Kind);
    }

    [Fact]
    public void Numbers_dates_reject_sub_tick_precision() {
        using MemoryStream package = CreateNumbersPackage(new[] {
            new TableSpec("Sub-tick date", 1, 1, 0.00000015d, date: true)
        }, includePreview: true);

        using var result = ExcelIWorkConverter.ConvertNumbersToExcelResult(package);
        IWorkTableCell cell = Assert.Single(Assert.Single(
            Assert.Single(result.Projection.Sheets).Tables).Cells);

        Assert.True(result.IsVisualFallback);
        Assert.Equal(IWorkCellKind.Error, cell.Kind);
    }

    [Theory]
    [InlineData("/Contents 4 0 R ")]
    [InlineData("/Contents [4 0 R] ")]
    public void Classic_pdf_pages_reject_missing_content_stream_references(
        string contents) {
        byte[] pdf = CreateOnePageClassicPdf(validKids: true,
            pageDictionaryPrefix: contents);

        Assert.False(IWorkPdfInfo.IsComplete(pdf));
    }

    [Theory]
    [InlineData(false, false)]
    [InlineData(true, false)]
    [InlineData(false, true)]
    [InlineData(true, true)]
    public void Classic_pdf_pages_validate_direct_and_indirect_stream_lengths(
        bool indirectLength, bool contentArray) {
        Assert.True(IWorkPdfInfo.IsComplete(
            CreateClassicPdfWithContentStream(indirectLength, contentArray)));
    }

    [Fact]
    public void Text_colors_reject_conflicting_color_models() {
        using MemoryStream package = CreatePagesPackageWithConflictingColorModels();

        using var result = WordIWorkConverter.ConvertPagesToWordResult(package);

        Assert.True(result.IsVisualFallback);
        Assert.Contains(result.Projection.Diagnostics,
            diagnostic => diagnostic.Code == "IWORK_PAGES_TEXT_UNSUPPORTED");
    }

    private static MemoryStream CreateKeynotePackageWithDistinctBodyPlaceholders() {
        const ulong documentId = 1;
        const ulong showId = 2;
        const ulong nodeId = 3;
        const ulong slideId = 4;
        const ulong firstShapeId = 5;
        const ulong firstStorageId = 6;
        const ulong secondShapeId = 7;
        const ulong secondStorageId = 8;
        byte[] records = Message(
            ArchiveRecord(documentId, 1, Message(ReferenceField(2, showId))),
            ArchiveRecord(showId, 2,
                KeynoteShow(Message(ReferenceField(2, nodeId)))),
            ArchiveRecord(nodeId, 4, Message(ReferenceField(2, slideId))),
            ArchiveRecord(slideId, 5, Message(
                ReferenceField(6, firstShapeId), ReferenceField(6, secondShapeId))),
            ArchiveRecord(firstShapeId, 2011, Message(ReferenceField(2, firstStorageId))),
            ArchiveRecord(firstStorageId, 2001, Message(StringField(3, "First"))),
            ArchiveRecord(secondShapeId, 2011, Message(ReferenceField(2, secondStorageId))),
            ArchiveRecord(secondStorageId, 2001, Message(StringField(3, "Second"))));
        return CreatePackage(
            ("Index/Slide.iwa", FrameIwa(records)),
            ("preview.png", ValidPreviewPng()));
    }

    private static byte[] CreateClassicPdfWithContentStream(bool indirectLength,
        bool contentArray) {
        string contents = contentArray ? "[4 0 R]" : "4 0 R";
        var objects = new List<string> {
            "1 0 obj\n<< /Type /Catalog /Pages 2 0 R >>\nendobj\n",
            "2 0 obj\n<< /Type /Pages /MediaBox [0 0 612 792] /Count 1 /Kids [3 0 R] >>\nendobj\n",
            "3 0 obj\n<< /Type /Page /Parent 2 0 R /Contents "
                + contents + " >>\nendobj\n",
            indirectLength
                ? "4 0 obj\n<< /Length 5 0 R >>\nstream\nq Q\nendstream\nendobj\n"
                : "4 0 obj\n<< /Length 3 >>\nstream\nq Q\nendstream\nendobj\n"
        };
        if (indirectLength) objects.Add("5 0 obj\n3\nendobj\n");

        const string header = "%PDF-1.4\n";
        var prefix = new StringBuilder(header);
        var offsets = new List<int>();
        foreach (string value in objects) {
            offsets.Add(Encoding.ASCII.GetByteCount(prefix.ToString()));
            prefix.Append(value);
        }
        int xrefOffset = Encoding.ASCII.GetByteCount(prefix.ToString());
        var suffix = new StringBuilder();
        suffix.Append("xref\n0 ").Append(objects.Count + 1).Append('\n')
            .Append("0000000000 65535 f \n");
        foreach (int objectOffset in offsets) {
            suffix.Append(objectOffset.ToString("D10",
                    System.Globalization.CultureInfo.InvariantCulture))
                .Append(" 00000 n \n");
        }
        suffix.Append("trailer\n<< /Size ").Append(objects.Count + 1)
            .Append(" /Root 1 0 R >>\nstartxref\n")
            .Append(xrefOffset.ToString(System.Globalization.CultureInfo.InvariantCulture))
            .Append("\n%%EOF\n");
        return Encoding.ASCII.GetBytes(prefix.ToString() + suffix);
    }

    private static MemoryStream CreatePagesPackageWithConflictingColorModels() {
        const ulong documentId = 1;
        const ulong bodyId = 2;
        const ulong styleId = 3;
        byte[] styleTable = Message(BytesField(1,
            Message(VarintField(1, 0), ReferenceField(2, styleId))));
        byte[] color = Message(FloatField(11, 0.5f),
            FloatField(3, 1f), FloatField(4, 0f), FloatField(5, 0f));
        byte[] records = Message(
            ArchiveRecord(documentId, 10000,
                Message(ReferenceField(4, bodyId)), new[] { bodyId }),
            ArchiveRecord(bodyId, 2001,
                Message(StringField(3, "Color"), BytesField(8, styleTable)),
                new[] { styleId }),
            ArchiveRecord(styleId, 2021,
                Message(BytesField(11, Message(BytesField(7, color))))));
        return CreatePackage(
            ("Index/Document.iwa", FrameIwa(records)),
            ("preview.png", ValidPreviewPng()));
    }
}
