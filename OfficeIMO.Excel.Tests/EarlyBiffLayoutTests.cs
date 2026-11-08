using System;
using System.IO;
using System.IO.Compression;
using System.Linq;
using System.Text;
using System.Xml.Linq;
using OfficeIMO.Excel;
using Xunit;

namespace OfficeIMO.Tests;

public sealed class EarlyBiffLayoutTests {
    private static readonly XNamespace Ns = "http://schemas.openxmlformats.org/spreadsheetml/2006/main";

    [Fact]
    public void NativeConversionRetainsEachUnsupportedArrayCellLocation() {
        string source = Path.Combine(RepositoryTestPaths.Find(), "OfficeIMO.TestAssets", "Documents",
            "LegacyXlsCorpus", "apache-openoffice-testdocs", "formula_import_biff4.xlw");
        string destination = Path.Combine(Path.GetTempPath(), "officeimo-biff-findings-" + Guid.NewGuid().ToString("N") + ".xlsx");
        byte[] original = File.ReadAllBytes(source);
        string[] expectedLocations = {
            "'Array'!B13", "'Array'!C13", "'Array'!B14", "'Array'!C14",
            "'Array'!B22", "'Array'!C22", "'Array'!B23", "'Array'!C23",
            "'Array'!B27", "'Array'!C27", "'Array'!B28", "'Array'!C28"
        };
        var allowed = ExcelDocument.AnalyzeConversion(source, destination,
            new ExcelDocumentConversionOptions { LossPolicy = OfficeConversionLossPolicy.Allow });
        Assert.Equal(expectedLocations, allowed.NativeDiagnostics
            .Where(finding => finding.Code == "early-biff-formula-unsupported"
                && finding.SourceLocation?.StartsWith("'Array'!", StringComparison.Ordinal) == true)
            .Select(finding => finding.SourceLocation));
        Assert.Equal(expectedLocations, allowed.Diagnostics
            .Where(finding => finding.Code == "early-biff-formula-unsupported"
                && finding.Location?.StartsWith("'Array'!", StringComparison.Ordinal) == true)
            .Select(finding => finding.Location));
        var blocked = Assert.Throws<ExcelDocumentConversionException>(() => ExcelDocument.Convert(source, destination));
        Assert.Equal(OfficeConversionFailureReason.DataLossBlocked, blocked.Reason);
        Assert.Equal(expectedLocations, blocked.Result.Report.NativeDiagnostics
            .Where(finding => finding.Code == "early-biff-formula-unsupported"
                && finding.SourceLocation?.StartsWith("'Array'!", StringComparison.Ordinal) == true)
            .Select(finding => finding.SourceLocation));
        Assert.False(File.Exists(destination));
        Assert.Equal(original, File.ReadAllBytes(source));
    }

    [Theory]
    [InlineData(2, true)]
    [InlineData(3, true)]
    [InlineData(4, true)]
    [InlineData(2, false)]
    [InlineData(3, false)]
    [InlineData(4, false)]
    public void OverflowDimensionsAreBoundedAndRequireExplicitLossAcceptance(int generation, bool overflowDefaults) {
        using var source = new MemoryStream(CreateWorksheet(generation, overflowDefaults ? (ushort)8181 : (ushort)300,
            rowHeight: 8182, columnWidth: 65282, standardWidth: overflowDefaults ? (ushort)65281 : (ushort)3072));
        using ExcelDocument document = ExcelDocument.Load(source, null, out OfficeConversionReport report);
        Assert.Equal(42.25d, document.Sheets[0].Cell(1, 1).GetValue<double>());
        Assert.Contains(report.Diagnostics, diagnostic => diagnostic.Code == "early-biff-row-height");
        Assert.Contains(report.Diagnostics, diagnostic => diagnostic.Code == "early-biff-column-width");
        using var destination = new MemoryStream();
        Assert.Throws<NotSupportedException>(() => document.Save(destination));
        Assert.Equal(0, destination.Length);
        XDocument worksheet = SaveWorksheet(document, destination, allowLoss: true);
        XElement format = worksheet.Root!.Element(Ns + "sheetFormatPr")!;
        Assert.Equal(overflowDefaults ? 409d : 15d, (double)format.Attribute("defaultRowHeight")!);
        if (generation == 4) Assert.Equal(overflowDefaults ? 255d : 12d, (double)format.Attribute("defaultColWidth")!);
        Assert.Equal(409d, (double)worksheet.Descendants(Ns + "row").Single(row => (uint?)row.Attribute("r") == 1).Attribute("ht")!);
        Assert.Equal(255d, (double)Assert.Single(worksheet.Descendants(Ns + "col")).Attribute("width")!);
    }

    [Theory]
    [InlineData(2)]
    [InlineData(3)]
    [InlineData(4)]
    public void ExactDestinationLimitsAndRowDefaultHeightFlagDoNotReportLoss(int generation) {
        using var source = new MemoryStream(CreateWorksheet(generation, 8180,
            rowHeight: 0xA000, columnWidth: 65280, standardWidth: 65280));
        using ExcelDocument document = ExcelDocument.Load(source, null, out OfficeConversionReport report);
        Assert.Empty(report.Diagnostics);
        using var destination = new MemoryStream();
        XDocument worksheet = SaveWorksheet(document, destination);
        Assert.Equal(409d, (double)worksheet.Root!.Element(Ns + "sheetFormatPr")!.Attribute("defaultRowHeight")!);
        Assert.Null(worksheet.Descendants(Ns + "row").Single(row => (uint?)row.Attribute("r") == 1).Attribute("ht"));
        Assert.Equal(255d, (double)Assert.Single(worksheet.Descendants(Ns + "col")).Attribute("width")!);
    }

    [Fact]
    public void StandardColumnWidthIsBoundedBeforeEditableProjection() {
        using var source = new MemoryStream(CreateWorksheet(4, 300, standardWidth: 65281));
        using ExcelDocument document = ExcelDocument.Load(source, null, out OfficeConversionReport report);
        Assert.Contains(report.Diagnostics, diagnostic => diagnostic.Code == "early-biff-column-width");
        Assert.Equal(42.25d, document.Sheets[0].Cell(1, 1).GetValue<double>());
        using var destination = new MemoryStream();
        Assert.Throws<NotSupportedException>(() => document.Save(destination));
        Assert.Equal(0, destination.Length);
        XDocument worksheet = SaveWorksheet(document, destination, allowLoss: true);
        Assert.Equal(255d, (double)worksheet.Root!.Element(Ns + "sheetFormatPr")!.Attribute("defaultColWidth")!);
    }

    [Theory]
    [InlineData(2, 0x812C, 15d, false)]
    [InlineData(3, 0x8000, 409d, true)]
    [InlineData(4, 0x8000, 409d, true)]
    public void DefaultHeightUsesTheGenerationSpecificHeightField(int generation, int rawHeight, double expected, bool loss) {
        using var source = new MemoryStream(CreateWorksheet(generation, (ushort)rawHeight));
        using ExcelDocument document = ExcelDocument.Load(source, null, out OfficeConversionReport report);
        Assert.Equal(loss, report.Diagnostics.Any(diagnostic => diagnostic.Code == "early-biff-row-height"));
        using var destination = new MemoryStream();
        XDocument worksheet = SaveWorksheet(document, destination, allowLoss: loss);
        Assert.Equal(expected, (double)worksheet.Root!.Element(Ns + "sheetFormatPr")!.Attribute("defaultRowHeight")!);
    }

    [Theory]
    [InlineData(3, false)]
    [InlineData(3, true)]
    [InlineData(4, false)]
    [InlineData(4, true)]
    public void ZeroDefaultHeightAndHiddenStatePreserveExplicitRowOverrides(int generation, bool hiddenByDefault) {
        using var source = new MemoryStream(CreateWorksheet(generation, 0, hiddenByDefault: hiddenByDefault, rowOverrides: true));
        using ExcelDocument document = ExcelDocument.Load(source, null, out OfficeConversionReport report);
        Assert.Empty(report.Diagnostics);
        ExcelSheet sheet = document.Sheets[0];
        Assert.Equal(hiddenByDefault, sheet.DefaultRowsHidden);
        Assert.Equal(42.25d, sheet.Cell(1, 1).GetValue<double>());
        Assert.Equal(2d, sheet.Cell(2, 1).GetValue<double>());
        Assert.Equal(3d, sheet.Cell(3, 1).GetValue<double>());
        using var destination = new MemoryStream();
        XDocument worksheet = SaveWorksheet(document, destination);
        XElement format = worksheet.Root!.Element(Ns + "sheetFormatPr")!;
        Assert.Equal(0d, (double)format.Attribute("defaultRowHeight")!);
        Assert.Equal(hiddenByDefault, (bool)format.Attribute("zeroHeight")!);
        XElement[] rows = worksheet.Descendants(Ns + "row").ToArray();
        XElement incidental = rows.Single(row => (uint?)row.Attribute("r") == 1);
        Assert.Null(incidental.Attribute("hidden"));
        XElement visible = rows.Single(row => (uint?)row.Attribute("r") == 2);
        XElement hidden = rows.Single(row => (uint?)row.Attribute("r") == 3);
        Assert.Equal(false, (bool?)visible.Attribute("hidden"));
        Assert.Equal(true, (bool?)hidden.Attribute("hidden"));
        Assert.Equal(12d, (double)visible.Attribute("ht")!);
        Assert.Equal(13d, (double)hidden.Attribute("ht")!);
        Assert.Equal(hiddenByDefault, (bool?)incidental.Attribute("hidden") ?? (bool)format.Attribute("zeroHeight")!);
    }

    [Theory]
    [InlineData(2)]
    [InlineData(3)]
    [InlineData(4)]
    public void ExplicitZeroRowHeightOverridesAPositiveDefaultWithoutLoss(int generation) {
        using var source = new MemoryStream(CreateWorksheet(generation, 300, rowHeight: 0));
        using ExcelDocument document = ExcelDocument.Load(source, null, out OfficeConversionReport report);
        Assert.Empty(report.Diagnostics);
        Assert.Equal(42.25d, document.Sheets[0].Cell(1, 1).GetValue<double>());
        using var destination = new MemoryStream();
        XDocument worksheet = SaveWorksheet(document, destination);
        Assert.Equal(15d, (double)worksheet.Root!.Element(Ns + "sheetFormatPr")!.Attribute("defaultRowHeight")!);
        XElement row = Assert.Single(worksheet.Descendants(Ns + "row"));
        Assert.Equal(0d, (double?)row.Attribute("ht"));
        Assert.Equal(true, (bool?)row.Attribute("customHeight"));
        Assert.Equal(false, (bool?)row.Attribute("hidden"));
    }

    [Theory]
    [InlineData(2)]
    [InlineData(3)]
    [InlineData(4)]
    public void AuthoredZeroColumnWidthReportsItsDefaultWidthFallback(int generation) {
        using var source = new MemoryStream(CreateWorksheet(generation, 300, columnWidth: 0, standardWidth: 0, columnHidden: true));
        using ExcelDocument document = ExcelDocument.Load(source, null, out OfficeConversionReport report);
        Assert.Contains(report.Diagnostics, diagnostic => diagnostic.Code == "early-biff-zero-column-width");
        using var destination = new MemoryStream();
        Assert.Throws<NotSupportedException>(() => document.Save(destination));
        XDocument worksheet = SaveWorksheet(document, destination, allowLoss: true);
        XElement column = Assert.Single(worksheet.Descendants(Ns + "col"));
        Assert.Null(column.Attribute("width"));
        Assert.Equal(generation != 2, (bool)column.Attribute("hidden")!);
    }

    [Fact]
    public void GlobalColumnWidthSelectionRetainsItsDefaultSentinelWithoutLoss() {
        using var source = new MemoryStream(CreateWorksheet(4, 300, columnWidth: 3328, standardWidth: 3072, useDefaultWidth: true));
        using ExcelDocument document = ExcelDocument.Load(source, null, out OfficeConversionReport report);
        Assert.Empty(report.Diagnostics);
        using var destination = new MemoryStream();
        XDocument worksheet = SaveWorksheet(document, destination);
        Assert.Equal(12d, (double)worksheet.Root!.Element(Ns + "sheetFormatPr")!.Attribute("defaultColWidth")!);
        Assert.Null(Assert.Single(worksheet.Descendants(Ns + "col")).Attribute("width"));
    }

    [Theory]
    [InlineData(2)]
    [InlineData(3)]
    [InlineData(4)]
    public void UnsupportedBaseColumnWidthRetainsItsInputAdmission(int generation) {
        using var source = new MemoryStream(CreateWorksheet(generation, 300, baseWidth: 256));
        Assert.Throws<InvalidDataException>(() => ExcelDocument.Load(source));
    }

    [Theory]
    [InlineData("Height", false)]
    [InlineData("AutoFitRow", false)]
    [InlineData("AutoFitRows", false)]
    [InlineData("RowLayout", true)]
    [InlineData("RowsLayout", true)]
    [InlineData("Clear", false)]
    [InlineData("RowLayoutClear", true)]
    [InlineData("Show", true)]
    [InlineData("Show", false)]
    public void EditedRowsRetainAuthoredZeroDefaultsAndVisibilityAfterReopening(string edit, bool hiddenByDefault) {
        using var source = new MemoryStream(CreateWorksheet(4, 0, hiddenByDefault: hiddenByDefault, rowOverrides: true));
        using ExcelDocument document = ExcelDocument.Load(source);
        ExcelSheet sheet = document.Sheets[0];
        int editedRow = edit is "Show" or "RowLayout" or "RowsLayout" or "RowLayoutClear" ? 3 : 2;
        ExcelColumnWidth height = ExcelColumnWidth.Length(OfficeLength.Points(24));
        switch (edit) {
            case "Height": sheet.SetRowHeight(editedRow, height); break;
            case "AutoFitRow": sheet.AutoFitRow(editedRow); break;
            case "AutoFitRows": sheet.AutoFitRows(); break;
            case "RowLayout": sheet.SetRowLayout(editedRow, new ExcelRowLayoutOptions { Height = height, Hidden = false }); break;
            case "RowsLayout": sheet.SetRowsLayout(new[] { editedRow }, new ExcelRowLayoutOptions { Height = height, Hidden = false }); break;
            case "Clear": sheet.SetRowHeight(editedRow, default); break;
            case "RowLayoutClear": sheet.SetRowLayout(editedRow, new ExcelRowLayoutOptions { ClearHeight = true, Hidden = false }); break;
            case "Show": sheet.SetRowHidden(editedRow, false); break;
            default: throw new ArgumentOutOfRangeException(nameof(edit));
        }
        using var destination = new MemoryStream();
        XDocument worksheet = SaveWorksheet(document, destination);
        destination.Position = 0;
        using ExcelDocument reopened = ExcelDocument.Load(destination);
        ExcelSheet saved = reopened.Sheets[0];
        Assert.Equal(OfficeLength.Zero, saved.DefaultRowHeight!.Value.ToLength(saved));
        Assert.Equal(hiddenByDefault, saved.DefaultRowsHidden);
        Assert.Equal(42.25d, saved.Cell(1, 1).GetValue<double>());
        Assert.Equal(2d, saved.Cell(2, 1).GetValue<double>());
        Assert.Equal(3d, saved.Cell(3, 1).GetValue<double>());
        XElement format = worksheet.Root!.Element(Ns + "sheetFormatPr")!;
        Assert.Equal(0d, (double)format.Attribute("defaultRowHeight")!);
        Assert.Equal(hiddenByDefault, (bool)format.Attribute("zeroHeight")!);
        XElement row = worksheet.Descendants(Ns + "row").Single(item => (uint?)item.Attribute("r") == editedRow);
        Assert.False((bool?)row.Attribute("hidden") ?? hiddenByDefault);
        if (hiddenByDefault) Assert.Equal(false, (bool?)row.Attribute("hidden"));
        else if (edit == "Show") Assert.Null(row.Attribute("hidden"));
        if (edit is "Clear" or "RowLayoutClear") {
            Assert.Null(row.Attribute("ht"));
            Assert.Null(row.Attribute("customHeight"));
        } else Assert.True((double?)row.Attribute("ht") > 0d);
    }

    [Fact]
    public void RowEditingRetainsTheDefaultForAnUnauthoredSheetHeight() {
        using ExcelDocument document = ExcelDocument.Create();
        ExcelSheet sheet = document.AddSheet("Data");
        sheet.Cell(1, 1).SetValue(42.25d);
        Assert.Null(sheet.DefaultRowHeight);
        sheet.SetRowHeight(1, ExcelColumnWidth.Length(OfficeLength.Points(24)));
        sheet.SetRowHidden(1, false);
        using var destination = new MemoryStream();
        XDocument worksheet = SaveWorksheet(document, destination);
        destination.Position = 0;
        using ExcelDocument reopened = ExcelDocument.Load(destination);
        ExcelSheet saved = reopened.Sheets[0];
        Assert.Equal(OfficeLength.Points(15), saved.DefaultRowHeight!.Value.ToLength(saved));
        Assert.False(saved.DefaultRowsHidden);
        Assert.Equal(42.25d, saved.Cell(1, 1).GetValue<double>());
        Assert.Equal(15d, (double)worksheet.Root!.Element(Ns + "sheetFormatPr")!.Attribute("defaultRowHeight")!);
        Assert.Null(Assert.Single(worksheet.Descendants(Ns + "row")).Attribute("hidden"));
    }

    [Theory]
    [InlineData("absent", 15d, false)]
    [InlineData("zero-visible", 0d, false)]
    [InlineData("zero-hidden", 0d, true)]
    [InlineData("positive-visible", 18d, false)]
    [InlineData("positive-hidden", 18d, true)]
    public void ColumnDefaultsRetainAValidDefaultHeightAndVisibility(string source, double height, bool hidden) {
        using var input = new MemoryStream(source == "absent" ? Array.Empty<byte>()
            : CreateWorksheet(4, (ushort)(height * 20), hiddenByDefault: hidden));
        using ExcelDocument document = source == "absent" ? ExcelDocument.Create() : ExcelDocument.Load(input);
        ExcelSheet sheet = source == "absent" ? document.AddSheet("Data") : document.Sheets[0];
        if (source == "absent") {
            sheet.Cell(1, 1).SetValue(42.25d);
            Assert.Null(sheet.DefaultRowHeight);
        }

        sheet.SetDefaultColumnWidth(ExcelColumnWidth.Characters(13.25d));

        using var destination = new MemoryStream();
        XDocument worksheet = SaveWorksheet(document, destination);
        XElement format = worksheet.Root!.Element(Ns + "sheetFormatPr")!;
        Assert.Equal(height, (double)format.Attribute("defaultRowHeight")!);
        Assert.Equal(hidden, (bool?)format.Attribute("zeroHeight") ?? false);
        Assert.Equal(13.25d, (double)format.Attribute("defaultColWidth")!);
        destination.Position = 0;
        using ExcelDocument reopened = ExcelDocument.Load(destination);
        ExcelSheet saved = reopened.Sheets[0];
        Assert.Equal(OfficeLength.Points(height), saved.DefaultRowHeight!.Value.ToLength(saved));
        Assert.Equal(hidden, saved.DefaultRowsHidden);
        Assert.Equal(42.25d, saved.Cell(1, 1).GetValue<double>());
    }

    private static XDocument SaveWorksheet(ExcelDocument document, MemoryStream destination, bool allowLoss = false) {
        document.Save(destination, new ExcelSaveOptions { AllowLoss = allowLoss, ValidateOpenXml = true });
        destination.Position = 0;
        using var package = new ZipArchive(destination, ZipArchiveMode.Read, leaveOpen: true);
        using Stream worksheet = package.GetEntry("xl/worksheets/sheet1.xml")!.Open();
        return XDocument.Load(worksheet);
    }

    private static byte[] CreateWorksheet(int generation, ushort defaultHeight, ushort? rowHeight = null,
        ushort? columnWidth = null, ushort? standardWidth = null, bool hiddenByDefault = false,
        bool rowOverrides = false, bool columnHidden = false, bool useDefaultWidth = false, ushort baseWidth = 10) {
        using var bytes = new MemoryStream();
        using var writer = new BinaryWriter(bytes);
        WriteRecord(generation == 2 ? (ushort)0x0009 : generation == 3 ? (ushort)0x0209 : (ushort)0x0409, record => {
            record.Write((ushort)0); record.Write((ushort)0x0010);
            if (generation != 2) record.Write((ushort)0);
        });
        WriteRecord(generation == 2 ? (ushort)0x0025 : (ushort)0x0225, record => {
            if (generation != 2) record.Write((ushort)(hiddenByDefault ? 2 : 0));
            record.Write(defaultHeight);
        });
        WriteRecord(0x0055, record => record.Write(baseWidth));
        if (generation == 4 && standardWidth.HasValue) WriteRecord(0x0099, record => record.Write(standardWidth.Value));
        if (columnWidth.HasValue) WriteRecord(generation == 2 ? (ushort)0x0024 : (ushort)0x007D, record => {
            if (generation == 2) { record.Write((byte)0); record.Write((byte)0); record.Write(columnWidth.Value); }
            else { record.Write((ushort)0); record.Write((ushort)0); record.Write(columnWidth.Value); record.Write((ushort)0); record.Write((ushort)(columnHidden ? 1 : 0)); }
        });
        if (useDefaultWidth) WriteRecord(0x00AB, record => {
            record.Write((ushort)32); var mask = new byte[32]; mask[0] = 1; record.Write(mask);
        });
        if (rowHeight.HasValue) WriteRow(0, rowHeight.Value, false);
        if (rowOverrides) { WriteRow(1, 240, false); WriteRow(2, 260, true); }
        WriteNumber(0, 42.25);
        if (rowOverrides) { WriteNumber(1, 2); WriteNumber(2, 3); }
        WriteRecord(0x000A, _ => { });
        return bytes.ToArray();

        void WriteRow(ushort row, ushort height, bool hidden) => WriteRecord(generation == 2 ? (ushort)8 : (ushort)0x0208, record => {
            record.Write(row); record.Write((ushort)0); record.Write((ushort)1); record.Write(height); record.Write((ushort)0);
            if (generation == 2) record.Write((byte)0);
            else { record.Write((ushort)0); record.Write((uint)(hidden ? 32 : 0)); }
        });
        void WriteNumber(ushort row, double value) => WriteRecord(generation == 2 ? (ushort)3 : (ushort)0x0203, record => {
            record.Write(row); record.Write((ushort)0); record.Write((ushort)0);
            if (generation == 2) record.Write((byte)0);
            record.Write(value);
        });
        void WriteRecord(ushort type, Action<BinaryWriter> write) {
            using var payload = new MemoryStream();
            using (var record = new BinaryWriter(payload, Encoding.UTF8, leaveOpen: true)) write(record);
            writer.Write(type); writer.Write((ushort)payload.Length); writer.Write(payload.ToArray());
        }
    }
}
