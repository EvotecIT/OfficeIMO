using System.Globalization;
using System.Xml.Linq;
using OfficeIMO.IWork;
using OfficeIMO.PowerPoint;
using OfficeIMO.Word;

namespace OfficeIMO.IWork.Tests;

public sealed partial class IWorkBoundaryTests {
    [Fact]
    public void Independent_dimension_fixture_preserves_nonuniform_rows_and_columns() {
        string path = Path.Combine(AppContext.BaseDirectory, "Documents", "IWorkCorpus",
            "numbers-parser", "individual-dimensions.numbers");
        IWorkSourceDocument source = IWorkSourceDocument.Open(path);
        IWorkTable table = Assert.Single(Assert.Single(source.ReadNumbers().Sheets).Tables);
        Assert.Equal(new double?[] { 20.0, 10, 30 }, Enumerable.Range(1, 3).Select(table.GetRowHeight));
        Assert.Equal(new double?[] { 40.0, 20 }, Enumerable.Range(1, 2).Select(table.GetColumnWidth));
        using var result = source.ToExcelDocumentResult();
        Assert.False(result.IsVisualFallback, string.Join("; ", result.Report.Diagnostics.Select(d => d.Code + ": " + d.Message)));
        using var saved = new MemoryStream();
        result.Value.Save(saved);
        saved.Position = 0;
        using var reopened = OfficeIMO.Excel.ExcelDocument.Load(saved);
        saved.Position = 0;
        using var zip = new ZipArchive(saved, ZipArchiveMode.Read, leaveOpen: true);
        using Stream xml = zip.GetEntry("xl/worksheets/sheet1.xml")!.Open();
        XDocument sheet = XDocument.Load(xml);
        Assert.Equal(new[] { 20.0, 10d, 30d }, sheet.Descendants(SpreadsheetNs + "row")
            .Select(row => double.Parse(row.Attribute("ht")!.Value, CultureInfo.InvariantCulture)));
    }

    [Theory]
    [InlineData(IWorkDocumentKind.Numbers)]
    [InlineData(IWorkDocumentKind.Pages)]
    [InlineData(IWorkDocumentKind.Keynote)]
    public void Individual_table_sizes_survive_projection_and_destination_reopen(IWorkDocumentKind kind) {
        using MemoryStream package = DimensionPackage(kind);
        IWorkSourceDocument source = IWorkSourceDocument.Open(package);
        IWorkTable table = kind switch {
            IWorkDocumentKind.Numbers => source.ReadNumbers().Sheets[0].Tables[0],
            IWorkDocumentKind.Pages => source.ReadPages().Tables[0],
            _ => source.ReadKeynote().Slides[0].Tables[0]
        };
        Assert.Equal(20.25d, table.GetRowHeight(1));
        Assert.Equal(10d, table.GetRowHeight(2));
        Assert.Equal(30d, table.GetRowHeight(3));
        Assert.Equal(40.5d, table.GetColumnWidth(1));
        Assert.Equal(20d, table.GetColumnWidth(2));
        Assert.False(table.RowHeights.ContainsKey(2));
        Assert.Throws<NotSupportedException>(() => ((IDictionary<int, double>)table.RowHeights).Add(2, 99));
        Assert.Throws<ArgumentOutOfRangeException>(() => table.GetRowHeight(0));
        Assert.Throws<ArgumentOutOfRangeException>(() => table.GetColumnWidth(3));
        using var saved = new MemoryStream();
        if (kind == IWorkDocumentKind.Numbers) {
            using var result = source.ToExcelDocumentResult();
            Assert.False(result.IsVisualFallback);
            result.Value.Save(saved);
            saved.Position = 0;
            using var reopened = OfficeIMO.Excel.ExcelDocument.Load(saved);
            saved.Position = 0;
            using var zip = new ZipArchive(saved, ZipArchiveMode.Read, leaveOpen: true);
            using Stream xml = zip.GetEntry("xl/worksheets/sheet1.xml")!.Open();
            XDocument sheet = XDocument.Load(xml);
            Assert.Equal(20.25d, double.Parse(sheet.Descendants(SpreadsheetNs + "row")
                .Single(row => row.Attribute("r")!.Value == "1").Attribute("ht")!.Value, CultureInfo.InvariantCulture));
            Assert.Equal((40.5d * 96d / 72d - 5d) / 7d,
                double.Parse(sheet.Descendants(SpreadsheetNs + "col").Single().Attribute("width")!.Value,
                    CultureInfo.InvariantCulture), 10);
        } else if (kind == IWorkDocumentKind.Pages) {
            using var result = source.ToWordDocumentResult();
            Assert.False(result.IsVisualFallback);
            result.Value.Save(saved);
            saved.Position = 0;
            using WordDocument reopened = WordDocument.Load(saved);
            WordTable output = Assert.Single(reopened.Tables);
            Assert.Equal(new[] { 405, 200, 600 }, output.RowHeight);
            Assert.Equal(new[] { 810, 400 }, output.ColumnWidth);
        } else {
            using var result = source.ToPowerPointPresentationResult();
            Assert.False(result.IsVisualFallback);
            result.Value.Save(saved);
            saved.Position = 0;
            using PowerPointPresentation reopened = PowerPointPresentation.Load(saved);
            PowerPointTable output = Assert.Single(Assert.Single(reopened.Slides).Tables);
            Assert.Equal(20.25d, output.GetRowHeightPoints(0), 5);
            Assert.Equal(10d, output.GetRowHeightPoints(1), 5);
            Assert.Equal(30d, output.GetRowHeightPoints(2), 5);
            Assert.Equal(40.5d, output.GetColumnWidthPoints(0), 5);
            Assert.Equal(20d, output.GetColumnWidthPoints(1), 5);
        }
    }

    [Fact]
    public void Keynote_individual_sizes_scale_with_the_drawable_extent_and_report_approximation() {
        using MemoryStream package = DimensionPackage(IWorkDocumentKind.Keynote, scale: 2f);
        using var result = PowerPointIWorkConverter.ConvertKeynoteToPowerPointResult(package);
        PowerPointTable table = Assert.Single(Assert.Single(result.Value.Slides).Tables);
        Assert.False(result.IsVisualFallback);
        Assert.Equal(121d, table.WidthPoints, 5);
        Assert.Equal(120.5d, table.HeightPoints, 5);
        Assert.Equal(81d, table.GetColumnWidthPoints(0), 5);
        Assert.Equal(40.5d, table.GetRowHeightPoints(0), 5);
        Assert.Contains(result.Report.Diagnostics, diagnostic => diagnostic.Code == "IWORK_KEYNOTE_TABLE_SIZING_SCALED"
            && diagnostic.LossKind == OfficeConversionLossKind.Approximation);
    }

    [Theory]
    [InlineData("duplicate")]
    [InlineData("out-of-range")]
    [InlineData("negative")]
    [InlineData("non-finite")]
    [InlineData("wrong-wire")]
    [InlineData("hidden")]
    [InlineData("missing-bucket")]
    public void Unsupported_dimension_records_keep_source_evidence_and_use_visual_fallback(string defect) {
        byte[] header = defect switch {
            "out-of-range" => DimensionHeader(3, 22f),
            "negative" => DimensionHeader(0, -1f),
            "non-finite" => DimensionHeader(0, float.NaN),
            "wrong-wire" => Message(VarintField(1, 0), DoubleField(2, 22), VarintField(3, 0)),
            "hidden" => DimensionHeader(0, 22f, hidingState: 1),
            _ => DimensionHeader(0, 22f)
        };
        using MemoryStream package = DimensionPackage(IWorkDocumentKind.Numbers,
            firstRowHeader: header, duplicate: defect == "duplicate", missingBucket: defect == "missing-bucket");
        using var result = ExcelIWorkConverter.ConvertNumbersToExcelResult(package);
        Assert.True(result.IsVisualFallback);
        Assert.Contains(result.Report.Diagnostics, diagnostic => diagnostic.Code == "IWORK_TABLE_HEADER_DIMENSIONS_UNSUPPORTED");
        Assert.NotEmpty(result.Report.PreservedRecords);
        if (defect == "duplicate") Assert.False(result.Projection.Sheets[0].Tables[0].RowHeights.ContainsKey(1));
        if (defect != "missing-bucket") {
            IWorkSourceDeclarationIssue issue = Assert.Single(result.Report.SourceDeclarationIssues);
            Assert.Equal(defect == "duplicate" ? 13ul : 12ul, issue.Owner.RecordIdentifier);
            Assert.Equal(defect switch {
                "duplicate" => "2[2]/1",
                "out-of-range" => "2[1]/1",
                "hidden" => "2[1]/3",
                _ => "2[1]/2"
            }, issue.FieldPath);
            Assert.Equal(1, issue.DeclaredValueCount);
            Assert.Equal(defect is "duplicate" or "out-of-range" ? IWorkSourceDeclarationIssueKind.InvalidSelectionMetadata
                : IWorkSourceDeclarationIssueKind.InvalidValue, issue.Kind);
        }
    }

    [Fact]
    public void Dimension_entry_budget_counts_all_buckets_and_headers_before_destination_creation() {
        using MemoryStream package = DimensionPackage(IWorkDocumentKind.Numbers, repeatTable: true);
        IWorkSourceDocument source = IWorkSourceDocument.Open(package,
            new IWorkReadOptions { MaximumTableDimensionEntries = 12 });
        Assert.Contains("source-wide", Assert.Throws<InvalidDataException>(() => source.ReadNumbers()).Message);
    }

    [Fact]
    public void Excel_rejects_an_individual_height_outside_its_range_instead_of_clamping() {
        using MemoryStream package = DimensionPackage(IWorkDocumentKind.Numbers,
            firstRowHeader: DimensionHeader(0, 410f));
        using var result = ExcelIWorkConverter.ConvertNumbersToExcelResult(package);
        Assert.True(result.IsVisualFallback);
        Assert.Contains(result.Report.Diagnostics, diagnostic => diagnostic.Code == "IWORK_NUMBERS_EXCEL_DESTINATION_UNSUPPORTED");
    }

    private static readonly XNamespace SpreadsheetNs = "http://schemas.openxmlformats.org/spreadsheetml/2006/main";

    private static byte[] DimensionHeader(ulong index, float size, ulong hidingState = 0) =>
        Message(VarintField(1, index), FloatField(2, size), VarintField(3, hidingState), VarintField(4, 0));

    private static MemoryStream DimensionPackage(IWorkDocumentKind kind, byte[]? firstRowHeader = null,
        bool duplicate = false, bool missingBucket = false, float scale = 1f, bool repeatTable = false,
        byte[]? firstBucketPayload = null, byte[]? secondBucketPayload = null, byte[]? columnBucketPayload = null,
        bool repeatBucket = false, uint firstBucketType = 6006) {
        byte[] store = Message(
            BytesField(1, Message(VarintField(1, 1), ReferenceField(2, missingBucket ? 99UL : 12UL), ReferenceField(2, 13),
                repeatBucket ? ReferenceField(2, 12) : Message())),
            ReferenceField(2, 14), BytesField(3, Message()));
        byte[] model = Message(BytesField(4, store), VarintField(6, 3), VarintField(7, 2),
            StringField(8, "Dimensions"), DoubleField(16, 10), DoubleField(17, 20));
        byte[] firstBucket = Message(VarintField(1, 1), BytesField(2, firstRowHeader ?? DimensionHeader(0, 20.25f)),
            BytesField(2, DimensionHeader(1, 0)));
        byte[] secondBucket = Message(VarintField(1, 1), BytesField(2, DimensionHeader(2, 30)),
            duplicate ? BytesField(2, DimensionHeader(0, 22)) : Array.Empty<byte>());
        byte[] columnBucket = Message(VarintField(1, 1), BytesField(2, DimensionHeader(0, 40.5f)),
            BytesField(2, DimensionHeader(1, 0)));
        firstBucket = firstBucketPayload ?? firstBucket;
        secondBucket = secondBucketPayload ?? secondBucket;
        columnBucket = columnBucketPayload ?? columnBucket;
        var records = new List<byte[]>();
        if (kind == IWorkDocumentKind.Numbers) {
            records.Add(ArchiveRecord(1, 1, Message(ReferenceField(1, 2))));
            records.Add(ArchiveRecord(2, 2, Message(StringField(1, "Sheet"), ReferenceField(2, 10),
                repeatTable ? ReferenceField(2, 20) : Array.Empty<byte>())));
        } else if (kind == IWorkDocumentKind.Pages) {
            records.Add(ArchiveRecord(1, 10000, Message(ReferenceField(4, 2)), new ulong[] { 2, 10 }));
            records.Add(ArchiveRecord(2, 2001, Message(StringField(3, "Body"))));
        } else {
            records.Add(ArchiveRecord(1, 1, Message(ReferenceField(2, 2))));
            records.Add(ArchiveRecord(2, 2, KeynoteShow(Message(ReferenceField(2, 3)))));
            records.Add(ArchiveRecord(3, 4, Message(ReferenceField(2, 4))));
            records.Add(ArchiveRecord(4, 5, Message(ReferenceField(6, 10))));
        }
        records.Add(ArchiveRecord(10, 6000, Message(
            kind == IWorkDocumentKind.Keynote ? BytesField(1, GeometryDrawable(72, 72, 60.5f * scale, 60.25f * scale))
                : Array.Empty<byte>(), ReferenceField(2, 11)), new ulong[] { 11 }));
        records.Add(ArchiveRecord(11, 6001, model, new ulong[] { 12, 13, 14 }));
        if (repeatTable) {
            records.Add(ArchiveRecord(20, 6000, Message(ReferenceField(2, 21)), new ulong[] { 21 }));
            records.Add(ArchiveRecord(21, 6001, model, new ulong[] { 12, 13, 14 }));
        }
        records.Add(ArchiveRecord(12, firstBucketType, firstBucket));
        records.Add(ArchiveRecord(13, 6006, secondBucket));
        records.Add(ArchiveRecord(14, 6006, columnBucket));
        return CreatePackage(("Index/Document.iwa", FrameIwa(Message(records.ToArray()))),
            ("preview.png", ValidPreviewPng()));
    }
}
