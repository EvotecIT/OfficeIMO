using OfficeIMO.IWork;
using OfficeIMO.Word;
using OfficeIMO.PowerPoint;
using System.Globalization;
using System.Xml.Linq;

namespace OfficeIMO.IWork.Tests;

public sealed partial class IWorkBoundaryTests {
    [Theory]
    [InlineData(IWorkDocumentKind.Pages, false)]
    [InlineData(IWorkDocumentKind.Numbers, false)]
    [InlineData(IWorkDocumentKind.Keynote, false)]
    [InlineData(IWorkDocumentKind.Pages, true)]
    [InlineData(IWorkDocumentKind.Numbers, true)]
    [InlineData(IWorkDocumentKind.Keynote, true)]
    public void Unreadable_dimension_bucket_retains_owner_and_invalidates_only_its_axis(IWorkDocumentKind kind, bool columns) {
        using MemoryStream package = DimensionPackage(kind,
            firstBucketPayload: columns ? null : new byte[] { 0x80 },
            columnBucketPayload: columns ? new byte[] { 0x80 } : null);
        IWorkTable table = ReadSelectedRichTable(IWorkSourceDocument.Open(package, kind), kind).Item1;
        if (columns) {
            Assert.Empty(table.ColumnWidths);
            Assert.Equal(20.25d, table.GetRowHeight(1));
            Assert.Equal(30d, table.GetRowHeight(3));
        } else {
            Assert.Empty(table.RowHeights);
            Assert.Equal(40.5d, table.GetColumnWidth(1));
        }
        package.Position = 0;
        IWorkConversionReport report = ConvertUnitReport(package, kind);
        IWorkSourceDeclarationIssue issue = Assert.Single(report.SourceDeclarationIssues);
        Assert.Equal(columns ? 14ul : 12ul, issue.Owner.RecordIdentifier);
        Assert.Equal("$", issue.FieldPath);
        Assert.Null(issue.DeclaredValueCount);
        Assert.Equal(IWorkSourceDeclarationIssueKind.MalformedMessage, issue.Kind);
    }

    [Theory]
    [InlineData(IWorkDocumentKind.Pages, false)]
    [InlineData(IWorkDocumentKind.Numbers, false)]
    [InlineData(IWorkDocumentKind.Keynote, false)]
    [InlineData(IWorkDocumentKind.Pages, true)]
    [InlineData(IWorkDocumentKind.Numbers, true)]
    [InlineData(IWorkDocumentKind.Keynote, true)]
    public void Unreadable_dimension_header_preserves_physical_path_without_trusting_later_bucket_sizes(IWorkDocumentKind kind, bool wrongWire) {
        byte[] first = Message(BytesField(2, DimensionHeader(0, 20.25f)),
            wrongWire ? VarintField(2, 0) : BytesField(2, new byte[] { 0x80 }));
        using MemoryStream package = DimensionPackage(kind, firstBucketPayload: first);
        IWorkTable table = ReadSelectedRichTable(IWorkSourceDocument.Open(package, kind), kind).Item1;
        Assert.Empty(table.RowHeights);
        Assert.Equal(40.5d, table.GetColumnWidth(1));
        package.Position = 0;
        IWorkConversionReport report = ConvertUnitReport(package, kind,
            readOptions: new IWorkReadOptions { PreserveSourceRecords = false });
        Assert.Empty(report.PreservedRecords);
        Assert.Contains(report.FidelityDiagnostics, issue => issue.Code == "IWORK_SOURCE_DECLARATIONS_UNASSESSED"
            && issue.LossKind == OfficeConversionLossKind.Unassessed);
        IWorkSourceDeclarationIssue issue = Assert.Single(report.SourceDeclarationIssues);
        Assert.Equal(12ul, issue.Owner.RecordIdentifier);
        Assert.Equal("2[2]", issue.FieldPath);
        Assert.Equal(1, issue.DeclaredValueCount);
        Assert.Equal(IWorkSourceDeclarationIssueKind.MalformedMessage, issue.Kind);
    }

    [Theory]
    [InlineData(IWorkDocumentKind.Pages)]
    [InlineData(IWorkDocumentKind.Numbers)]
    [InlineData(IWorkDocumentKind.Keynote)]
    public void Invalid_dimension_value_retains_distinct_sibling_sizes_and_physical_evidence(IWorkDocumentKind kind) {
        using MemoryStream package = DimensionPackage(kind, firstRowHeader: DimensionHeader(0, -1));
        IWorkTable table = ReadSelectedRichTable(IWorkSourceDocument.Open(package, kind), kind).Item1;
        Assert.False(table.RowHeights.ContainsKey(1));
        Assert.Equal(30d, table.GetRowHeight(3));
        Assert.Equal(40.5d, table.GetColumnWidth(1));
        package.Position = 0;
        IWorkSourceDeclarationIssue issue = Assert.Single(ConvertUnitReport(package, kind).SourceDeclarationIssues);
        Assert.Equal(12ul, issue.Owner.RecordIdentifier);
        Assert.Equal("2[1]/2", issue.FieldPath);
        Assert.Equal(1, issue.DeclaredValueCount);
        Assert.Equal(IWorkSourceDeclarationIssueKind.InvalidValue, issue.Kind);
    }

    [Theory]
    [InlineData("missing", 0)]
    [InlineData("duplicate", 2)]
    [InlineData("wire", 1)]
    public void Unreadable_dimension_index_does_not_establish_uniqueness(string defect, int count) {
        byte[] index = defect switch {
            "missing" => Message(),
            "duplicate" => Message(VarintField(1, 0), VarintField(1, 1)),
            _ => BytesField(1, new byte[] { 0 })
        };
        using MemoryStream package = DimensionPackage(IWorkDocumentKind.Numbers,
            firstRowHeader: Message(index, FloatField(2, 20), VarintField(3, 0)));
        IWorkNumbersProjection projection = IWorkSourceDocument.Open(package, IWorkDocumentKind.Numbers).ReadNumbers();
        IWorkTable table = Assert.Single(Assert.Single(projection.Sheets).Tables);
        Assert.Empty(table.RowHeights);
        Assert.Equal(40.5d, table.GetColumnWidth(1));
        IWorkSourceDeclarationIssue issue = Assert.Single(projection.SourceDeclarationIssues);
        Assert.Equal("2[1]/1", issue.FieldPath);
        Assert.Equal(count, issue.DeclaredValueCount);
        Assert.Equal(IWorkSourceDeclarationIssueKind.InvalidSelectionMetadata, issue.Kind);
    }

    [Fact]
    public void Third_dimension_duplicate_cannot_restore_an_override_after_an_invalid_value() {
        using MemoryStream package = DimensionPackage(IWorkDocumentKind.Numbers,
            secondBucketPayload: Message(BytesField(2, DimensionHeader(0, -1)), BytesField(2, DimensionHeader(0, 30))));
        IWorkNumbersProjection projection = IWorkSourceDocument.Open(package, IWorkDocumentKind.Numbers).ReadNumbers();
        Assert.False(Assert.Single(Assert.Single(projection.Sheets).Tables).RowHeights.ContainsKey(1));
        Assert.Equal(new[] { "2[1]/1", "2[1]/2", "2[2]/1" }, projection.SourceDeclarationIssues.Select(issue => issue.FieldPath));
        Assert.All(projection.SourceDeclarationIssues, issue => Assert.Equal(13ul, issue.Owner.RecordIdentifier));
    }

    [Theory]
    [InlineData(IWorkDocumentKind.Pages)]
    [InlineData(IWorkDocumentKind.Numbers)]
    [InlineData(IWorkDocumentKind.Keynote)]
    public void Unresolved_dimension_bucket_does_not_leave_other_bucket_overrides_trusted(IWorkDocumentKind kind) {
        using MemoryStream package = DimensionPackage(kind, missingBucket: true);
        IWorkTable table = ReadSelectedRichTable(IWorkSourceDocument.Open(package, kind), kind).Item1;
        Assert.Empty(table.RowHeights);
        Assert.Equal(40.5d, table.GetColumnWidth(1));
        package.Position = 0;
        IWorkConversionReport report = ConvertUnitReport(package, kind);
        Assert.Empty(report.SourceDeclarationIssues);
        AssertMissingReference(Assert.Single(report.SourceReferenceIssues), 11, "4/1/2", 99);
    }

    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public void Configured_dimension_bucket_or_nested_header_field_limit_remains_fatal(bool nested) {
        byte[] fields = Message(Enumerable.Range(0, 9).Select(_ => VarintField(1, 0)).ToArray());
        using MemoryStream package = DimensionPackage(IWorkDocumentKind.Numbers,
            firstBucketPayload: nested ? BytesField(2, fields) : fields);
        IWorkSourceDocument source = IWorkSourceDocument.Open(package, IWorkDocumentKind.Numbers,
            new IWorkReadOptions { MaximumProtobufFieldCount = 8 });
        Assert.Contains("field limit", Assert.Throws<InvalidDataException>(() => source.ReadNumbers()).Message);
    }

    [Fact]
    public void Unreadable_dimension_headers_consume_budget_before_nested_parsing() {
        using MemoryStream package = DimensionPackage(IWorkDocumentKind.Numbers,
            firstBucketPayload: Message(BytesField(2, new byte[] { 0x80 }), VarintField(2, 0), BytesField(2, new byte[] { 0x80 })));
        IWorkSourceDocument source = IWorkSourceDocument.Open(package, IWorkDocumentKind.Numbers,
            new IWorkReadOptions { MaximumTableDimensionEntries = 4 });
        Assert.Contains("dimension entries", Assert.Throws<InvalidDataException>(() => source.ReadNumbers()).Message);
    }

    [Fact]
    public void Dimension_declaration_budget_is_cumulative_and_snapshotted() {
        using MemoryStream package = DimensionPackage(IWorkDocumentKind.Numbers,
            firstBucketPayload: Message(BytesField(2, new byte[] { 0x80 }), BytesField(2, new byte[] { 0x80 })));
        var options = new IWorkReadOptions { MaximumSourceDeclarationIssues = 1 };
        IWorkSourceDocument source = IWorkSourceDocument.Open(package, IWorkDocumentKind.Numbers, options);
        options.MaximumSourceDeclarationIssues = 10;
        Assert.Contains("declaration issues", Assert.Throws<InvalidDataException>(() => source.ReadNumbers()).Message);
    }

    [Fact]
    public void Reused_dimension_bucket_reports_each_physical_path_once() {
        using MemoryStream package = DimensionPackage(IWorkDocumentKind.Numbers, repeatTable: true,
            firstRowHeader: DimensionHeader(0, -1));
        IWorkNumbersProjection projection = IWorkSourceDocument.Open(package, IWorkDocumentKind.Numbers,
            new IWorkReadOptions { MaximumSourceDeclarationIssues = 1 }).ReadNumbers();
        Assert.Equal(2, Assert.Single(projection.Sheets).Tables.Count);
        Assert.Single(projection.SourceDeclarationIssues);
    }

    [Theory]
    [InlineData(0)]
    [InlineData(1)]
    [InlineData(2)]
    public void Rejected_dimension_bucket_selection_retains_evidence_without_trusting_sizes(int repeatCount) {
        bool duplicate = repeatCount > 0;
        using MemoryStream package = DimensionPackage(IWorkDocumentKind.Numbers,
            repeatBucketCount: repeatCount, firstBucketType: duplicate ? 6006u : 2001u);
        IWorkNumbersProjection projection = IWorkSourceDocument.Open(package, IWorkDocumentKind.Numbers).ReadNumbers();
        IWorkTable table = Assert.Single(Assert.Single(projection.Sheets).Tables);
        if (duplicate) {
            Assert.False(table.RowHeights.ContainsKey(1));
            Assert.Equal(30d, table.GetRowHeight(3));
        } else Assert.Empty(table.RowHeights);
        Assert.Equal(40.5d, table.GetColumnWidth(1));
        IWorkSourceDeclarationIssue issue = Assert.Single(projection.SourceDeclarationIssues);
        Assert.Equal(duplicate ? 11ul : 12ul, issue.Owner.RecordIdentifier);
        Assert.Equal(duplicate ? "4/1/2" : "$", issue.FieldPath);
        Assert.Equal(duplicate ? 2 + repeatCount : (int?)null, issue.DeclaredValueCount);
        Assert.Equal(duplicate ? IWorkSourceDeclarationIssueKind.InvalidSelectionMetadata
            : IWorkSourceDeclarationIssueKind.RejectedMessageSet, issue.Kind);
    }

    [Theory]
    [InlineData(IWorkDocumentKind.Pages, false)]
    [InlineData(IWorkDocumentKind.Numbers, false)]
    [InlineData(IWorkDocumentKind.Keynote, false)]
    [InlineData(IWorkDocumentKind.Pages, true)]
    [InlineData(IWorkDocumentKind.Numbers, true)]
    [InlineData(IWorkDocumentKind.Keynote, true)]
    public void Recovered_sibling_dimensions_survive_partial_destination_save_and_reopen(IWorkDocumentKind kind, bool unknownIndex) {
        using MemoryStream package = DimensionPackage(kind, firstRowHeader: DimensionHeader(0, -1),
            firstBucketPayload: unknownIndex ? new byte[] { 0x80 } : null);
        IWorkSourceDocument source = IWorkSourceDocument.Open(package, kind);
        var options = new IWorkConversionOptions { AllowPartialEditableReconstruction = true };
        using var saved = new MemoryStream();
        if (kind == IWorkDocumentKind.Pages) {
            using var automatic = source.ToWordDocumentResult(options:new IWorkConversionOptions { RequireCompleteVisualCoverage = false });
            Assert.True(automatic.IsVisualFallback);
            using var partial = source.ToWordDocumentResult(options);
            Assert.True(partial.Report.IsPartialEditableReconstruction);
            partial.Value.Save(saved); saved.Position = 0;
            using WordDocument reopened = WordDocument.Load(saved);
            WordTable table = Assert.Single(reopened.Tables);
            Assert.Equal(new[] { 200, 200, unknownIndex ? 200 : 600 }, table.RowHeight);
            Assert.Equal(new[] { 810, 400 }, table.ColumnWidth);
        } else if (kind == IWorkDocumentKind.Numbers) {
            using var automatic = source.ToExcelDocumentResult(options:new IWorkConversionOptions { RequireCompleteVisualCoverage = false });
            Assert.True(automatic.IsVisualFallback);
            using var partial = source.ToExcelDocumentResult(options);
            Assert.True(partial.Report.IsPartialEditableReconstruction);
            partial.Value.Save(saved); saved.Position = 0;
            using var reopened = OfficeIMO.Excel.ExcelDocument.Load(saved);
            saved.Position = 0;
            using var zip = new ZipArchive(saved, ZipArchiveMode.Read, leaveOpen: true);
            using Stream xml = zip.GetEntry("xl/worksheets/sheet1.xml")!.Open();
            XDocument sheet = XDocument.Load(xml);
            Assert.Equal(10d, double.Parse(sheet.Descendants(SpreadsheetNs + "sheetFormatPr")
                .Single().Attribute("defaultRowHeight")!.Value, CultureInfo.InvariantCulture));
            Assert.Equal(unknownIndex ? Array.Empty<double>() : new[] { 30d }, sheet.Descendants(SpreadsheetNs + "row")
                .Where(row => row.Attribute("ht") != null)
                .Select(row => double.Parse(row.Attribute("ht")!.Value, CultureInfo.InvariantCulture)));
            Assert.Equal((40.5d * 96d / 72d - 5d) / 7d, double.Parse(sheet.Descendants(SpreadsheetNs + "col")
                .Single().Attribute("width")!.Value, CultureInfo.InvariantCulture), 10);
        } else {
            using var automatic = source.ToPowerPointPresentationResult(options:new IWorkConversionOptions { RequireCompleteVisualCoverage = false });
            Assert.True(automatic.IsVisualFallback);
            using var partial = source.ToPowerPointPresentationResult(options);
            Assert.True(partial.Report.IsPartialEditableReconstruction);
            Assert.Equal(!unknownIndex, partial.Report.Diagnostics.Any(issue => issue.Code == "IWORK_KEYNOTE_TABLE_SIZING_SCALED"));
            partial.Value.Save(saved); saved.Position = 0;
            using PowerPointPresentation reopened = PowerPointPresentation.Load(saved);
            PowerPointTable table = Assert.Single(Assert.Single(reopened.Slides).Tables);
            // Existing PPTX policy scales recovered dimensions to the native drawable extent.
            double expectedHeight = unknownIndex ? 60.25d / 3d : 30d * 60.25d / 50d;
            Assert.InRange(Math.Abs(table.GetRowHeightPoints(2) - expectedHeight), 0d, 1d / 12700d);
            Assert.Equal(40.5d, table.GetColumnWidthPoints(0), 5);
        }
    }
}
