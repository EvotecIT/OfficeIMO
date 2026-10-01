using System.Globalization;
using System.IO.Compression;
using System.Xml.Linq;
using OfficeIMO.IWork;

namespace OfficeIMO.IWork.Tests;

public sealed partial class IWorkBoundaryTests {
    [Theory]
    [InlineData(IWorkDocumentKind.Pages)]
    [InlineData(IWorkDocumentKind.Numbers)]
    [InlineData(IWorkDocumentKind.Keynote)]
    public void Numeric_formats_are_shared_metadata_without_changing_raw_values(IWorkDocumentKind kind) {
        using MemoryStream package = NumberFormatPackage(kind, NumericFormat(258, 2, 3, 1));
        IWorkTableCell cell = Assert.Single(ReadSelectedRichTable(IWorkSourceDocument.Open(package, kind), kind).Item1.Cells);
        Assert.Equal(0.5d, cell.Value);
        Assert.Equal("0.5", cell.DisplayText);
        Assert.False(cell.HasDecodeError);
        Assert.NotNull(cell.NumberFormat);
        Assert.Equal(IWorkNumberFormatKind.Percentage, cell.NumberFormat.Kind);
        Assert.Equal(2, cell.NumberFormat.DecimalPlaces);
        Assert.True(cell.NumberFormat.ThousandsSeparator);
        Assert.Equal(IWorkNegativeNumberStyle.RedAndParentheses, cell.NumberFormat.NegativeStyle);
        package.Position = 0;
        var report = ConvertUnitReport(package, kind, visual: false);
        if (kind != IWorkDocumentKind.Numbers) Assert.Contains(report.Diagnostics, diagnostic =>
            diagnostic.Code == (kind == IWorkDocumentKind.Pages ? "IWORK_PAGES_NUMBER_FORMAT_OMITTED" : "IWORK_KEYNOTE_NUMBER_FORMAT_OMITTED")
            && diagnostic.LossKind == global::OfficeIMO.OfficeConversionLossKind.Omission);
    }

    [Theory]
    [InlineData(256u, 0u, 0u, 0u, "0")]
    [InlineData(256u, 3u, 1u, 1u, "#,##0.000;[Red]#,##0.000")]
    [InlineData(258u, 2u, 2u, 0u, "0.00%;(0.00%)")]
    [InlineData(258u, 0u, 3u, 1u, "#,##0%;[Red](#,##0%)")]
    [InlineData(258u, 253u, 0u, 0u, "0.###############%")]
    public void Excel_numeric_format_survives_save_with_typed_cache(uint type, uint decimals, uint negative, uint grouping, string code) {
        using MemoryStream package = NumberFormatPackage(IWorkDocumentKind.Numbers, NumericFormat(type, decimals, negative, grouping), value: -0.5d);
        using var result = ExcelIWorkConverter.ConvertNumbersToExcelResult(package);
        Assert.False(result.IsVisualFallback);
        using var saved = new MemoryStream(); result.Value.Save(saved); saved.Position = 0;
        using var reopened = global::OfficeIMO.Excel.ExcelDocument.Load(saved);
        Assert.Single(reopened.Sheets);
        saved.Position = 0; using var zip = new ZipArchive(saved, ZipArchiveMode.Read, leaveOpen: true);
        XNamespace ns = "http://schemas.openxmlformats.org/spreadsheetml/2006/main";
        using Stream sheetStream = zip.GetEntry("xl/worksheets/sheet1.xml")!.Open();
        XElement cell = XDocument.Load(sheetStream).Descendants(ns + "c").Single();
        Assert.Equal("-0.5", cell.Element(ns + "v")!.Value);
        Assert.NotEqual("s", cell.Attribute("t")?.Value);
        using Stream styleStream = zip.GetEntry("xl/styles.xml")!.Open();
        XElement styles = XDocument.Load(styleStream).Root!;
        int styleIndex = int.Parse(cell.Attribute("s")!.Value, CultureInfo.InvariantCulture);
        string formatId = styles.Element(ns + "cellXfs")!.Elements().ElementAt(styleIndex).Attribute("numFmtId")!.Value;
        string actualCode = styles.Element(ns + "numFmts")?.Elements().FirstOrDefault(element => element.Attribute("numFmtId")!.Value == formatId)
            ?.Attribute("formatCode")?.Value ?? (formatId == "1" ? "0" : "General");
        Assert.Equal(code, actualCode);
        Assert.Equal(decimals == 253, result.Report.Diagnostics.Any(diagnostic => diagnostic.Code == "IWORK_NUMBERS_AUTOMATIC_DECIMALS_APPROXIMATED"));
    }

    [Theory]
    [InlineData(0)] // Invalid type.
    [InlineData(1)] // Unsupported decimals.
    [InlineData(2)] // Unknown property.
    [InlineData(3)] // Duplicate type.
    [InlineData(4)] // Wrong wire kind.
    [InlineData(5)] // Malformed nested message.
    public void Unsupported_selected_format_retains_numeric_value_and_physical_evidence(int failure) {
        byte[] format = failure switch {
            0 => NumericFormat(257, 2),
            1 => NumericFormat(258, 254),
            2 => Message(NumericFormat(258, 2), VarintField(19, 0)),
            3 => Message(NumericFormat(258, 2), VarintField(1, 258)),
            4 => BytesField(1, Message()),
            _ => new byte[] { 0x80 }
        };
        using MemoryStream package = NumberFormatPackage(IWorkDocumentKind.Numbers, format);
        IWorkTableCell cell = Assert.Single(ReadSelectedRichTable(IWorkSourceDocument.Open(package), IWorkDocumentKind.Numbers).Item1.Cells);
        Assert.Equal(0.5d, cell.Value); Assert.Null(cell.NumberFormat); Assert.False(cell.HasDecodeError);
        package.Position = 0;
        IWorkConversionReport report = ConvertUnitReport(package, IWorkDocumentKind.Numbers, visual: false,
            readOptions: new IWorkReadOptions { PreserveSourceRecords = false });
        Assert.True(report.IsPartialEditableReconstruction);
        IWorkSourceDeclarationIssue issue = Assert.Single(report.SourceDeclarationIssues);
        Assert.Equal(13ul, issue.Owner.RecordIdentifier); Assert.Equal("3[1]/6", issue.FieldPath);
        Assert.Equal(1, issue.DeclaredValueCount);
        Assert.Empty(report.PreservedRecords);
        Assert.Contains(report.Diagnostics, diagnostic => diagnostic.Code == "IWORK_TABLE_NUMBER_FORMAT_UNSUPPORTED"
            && diagnostic.LossKind == global::OfficeIMO.OfficeConversionLossKind.Unassessed);
    }

    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public void Ambiguous_format_keys_never_choose_an_arbitrary_value(bool unknownKey) {
        byte[] healthy = BytesField(3, FormatEntry(NumericFormat(258, 0)));
        byte[] bad = BytesField(3, unknownKey ? Message(BytesField(6, NumericFormat(256, 0))) : FormatEntry(NumericFormat(256, 0)));
        using MemoryStream package = NumberFormatPackage(IWorkDocumentKind.Numbers, catalog: Message(VarintField(1, 2), healthy, bad));
        IWorkTableCell cell = Assert.Single(ReadSelectedRichTable(IWorkSourceDocument.Open(package), IWorkDocumentKind.Numbers).Item1.Cells);
        Assert.Null(cell.NumberFormat); Assert.Equal(0.5d, cell.Value);
        package.Position = 0;
        var issue = Assert.Single(ConvertUnitReport(package, IWorkDocumentKind.Numbers).SourceDeclarationIssues);
        Assert.Equal("3[2]/1", issue.FieldPath);
    }

    [Fact]
    public void Unselected_format_values_are_not_traversed_but_catalog_entries_are_bounded() {
        byte[] catalog = Message(VarintField(1, 2), BytesField(3, FormatEntry(NumericFormat(258, 0))),
            BytesField(3, Message(VarintField(1, 99), BytesField(6, new byte[] { 0x80 }))));
        using MemoryStream package = NumberFormatPackage(IWorkDocumentKind.Numbers, catalog: catalog);
        Assert.NotNull(Assert.Single(ReadSelectedRichTable(IWorkSourceDocument.Open(package), IWorkDocumentKind.Numbers).Item1.Cells).NumberFormat);
        package.Position = 0;
        Assert.Empty(ConvertUnitReport(package, IWorkDocumentKind.Numbers).SourceDeclarationIssues);
        package.Position = 0;
        IWorkSourceDocument limited = IWorkSourceDocument.Open(package, new IWorkReadOptions { MaximumTableCatalogEntries = 1 });
        Assert.Contains("catalog limit", Assert.Throws<InvalidDataException>(() => limited.ReadNumbers()).Message);
    }

    [Fact]
    public void Format_payload_field_limit_remains_fatal() {
        using MemoryStream package = NumberFormatPackage(IWorkDocumentKind.Numbers,
            Message(NumericFormat(258, 0), Enumerable.Range(6, 8).Select(field => VarintField(field, 0)).SelectMany(bytes => bytes).ToArray()));
        IWorkSourceDocument source = IWorkSourceDocument.Open(package, new IWorkReadOptions { MaximumProtobufFieldCount = 8 });
        Assert.Contains("field limit", Assert.Throws<InvalidDataException>(() => source.ReadNumbers()).Message);
    }

    [Theory]
    [InlineData(0)]
    [InlineData(1)]
    [InlineData(2)]
    public void Invalid_BNC_catalog_marker_is_not_accepted(int failure) {
        byte[] marker = failure == 0 ? VarintField(5, 2) : failure == 1 ? BytesField(5, Message()) : Message(VarintField(5, 1), VarintField(5, 0));
        using MemoryStream package = NumberFormatPackage(IWorkDocumentKind.Numbers, catalog: Message(VarintField(1, 2), marker,
            BytesField(3, FormatEntry(NumericFormat(258, 0)))));
        var cell = Assert.Single(ReadSelectedRichTable(IWorkSourceDocument.Open(package), IWorkDocumentKind.Numbers).Item1.Cells);
        Assert.Null(cell.NumberFormat); Assert.Equal(0.5d, cell.Value);
        package.Position = 0; Assert.Single(ConvertUnitReport(package, IWorkDocumentKind.Numbers).SourceDeclarationIssues);
    }

    [Fact]
    public void Currency_selection_does_not_apply_an_inactive_numeric_format() {
        using MemoryStream package = NumberFormatPackage(IWorkDocumentKind.Numbers, NumericFormat(258, 0), currency: true);
        IWorkNumbersProjection projection = IWorkSourceDocument.Open(package).ReadNumbers();
        var cell = Assert.Single(Assert.Single(Assert.Single(projection.Sheets).Tables).Cells);
        Assert.Null(cell.NumberFormat); Assert.Equal(0.5d, cell.Value);
        Assert.Contains(projection.Diagnostics, diagnostic => diagnostic.Code == "IWORK_TABLE_NUMBER_FORMAT_UNSUPPORTED");
    }

    private static byte[] NumericFormat(uint type, uint decimals, uint negative = 0, uint grouping = 0) =>
        Message(VarintField(1, type), VarintField(2, decimals), VarintField(4, negative), VarintField(5, grouping));

    private static byte[] FormatEntry(byte[] format) => Message(VarintField(1, 1), BytesField(6, format));

    private static MemoryStream NumberFormatPackage(IWorkDocumentKind kind, byte[]? format = null, byte[]? catalog = null,
        double value = 0.5d, bool currency = false) {
        byte[] cell = new byte[currency ? 28 : 24]; cell[0] = 5; cell[1] = 2;
        WriteUInt32(cell, 8, (1u << 1) | (1u << 13) | (currency ? 1u << 14 : 0u));
        Buffer.BlockCopy(BitConverter.GetBytes(value), 0, cell, 12, 8); WriteUInt32(cell, 20, 1);
        if (currency) WriteUInt32(cell, 24, 1);
        return TableDependencyPackage(kind, ReferenceField(22, 13), cellPayload: cell,
            additionalRecords: ArchiveRecord(13, 6005, catalog ?? Message(VarintField(1, 2), VarintField(5, 1),
                BytesField(3, FormatEntry(format!)))));
    }
}
