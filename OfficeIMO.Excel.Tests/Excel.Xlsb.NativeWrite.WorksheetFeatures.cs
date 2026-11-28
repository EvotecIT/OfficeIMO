using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Spreadsheet;
using OfficeIMO.Excel;
using OfficeIMO.Excel.Xlsb.Biff12;
using OfficeIMO.Excel.Xlsb.Write;
using System.IO.Compression;
using System.Xml.Linq;
using Xunit;

namespace OfficeIMO.Tests;

/// <summary>Worksheet content after the cell table that the native XLSB writers encode and the reader projects.</summary>
public sealed class ExcelXlsbWorksheetFeatureTests {
    [Fact]
    public void Native_cell_writer_uses_only_cells_directly_in_the_row() {
        using ExcelDocument document = ExcelDocument.Create();
        ExcelSheet sheet = document.AddSheet("Data");
        sheet.Cell(1, 1).SetValue(3D);
        Row row = sheet.WorksheetPart.Worksheet!.GetFirstChild<SheetData>()!.GetFirstChild<Row>()!;
        const string spreadsheet = "http://schemas.openxmlformats.org/spreadsheetml/2006/main";
        var extensionList = new OpenXmlUnknownElement("x", "extLst", spreadsheet);
        var extension = new OpenXmlUnknownElement("v", "payload", "urn:officeimo:xlsb-test");
        extension.Append(new Cell { CellReference = "B1", CellValue = new CellValue("4") });
        extensionList.Append(extension);
        row.Append(extensionList);

        XlsbWriteCell cell = Assert.Single(XlsbWorksheetCellExtractor.ExtractNew(document, sheet));

        Assert.Equal(1, cell.Row);
        Assert.Equal(1, cell.Column);
        Assert.Equal(3D, cell.Value);
    }

    [Fact]
    public void Native_multiline_text_import_keeps_the_workbook_deferred() {
        const string text = "First\r\nSecond";
        using ExcelDocument seed = ExcelDocument.Create();
        ExcelSheet sheet = seed.AddSheet("Lines");
        sheet.Cell(1, 1).SetValue(text);
        sheet.Cell(1, 2).SetValue(17);
        sheet.SetCellStyleIndex(1, 1, 0, save: false);
        byte[] native = seed.ToBytes(ExcelFileFormat.Xlsb);

        using ExcelDocument document = ExcelDocument.Load(new MemoryStream(native));
        Assert.True(document.IsSdkDeferred);
        Assert.Equal(native, document.ToBytes(ExcelFileFormat.Xlsb));
        Assert.Empty(document.OpenXmlSdkFallbacks);
        Assert.Equal(text, document.Sheets[0].Cell(1, 1).GetValue<string>());
        Assert.Equal(17, document.Sheets[0].Cell(1, 2).GetValue<int>());
    }

    [Fact]
    public void Page_breaks_and_ignored_errors_survive_native_xlsb_save() {
        byte[] xlsx = CreateXlsx(document => {
            ExcelSheet sheet = document.AddSheet("Data");
            sheet.Cell(1, 1).SetValue("12");
            sheet.Cell(40, 3).SetValue(1);
            sheet.AddManualRowPageBreak(20, save: false);
            sheet.AddManualRowPageBreak(35, save: false);
            sheet.AddManualColumnPageBreak(2, save: false);
            sheet.WorksheetPart.Worksheet!.Append(new IgnoredErrors(
                new IgnoredError { SequenceOfReferences = new ListValue<StringValue> { InnerText = "A1:A5 C3" }, NumberStoredAsText = true },
                new IgnoredError { SequenceOfReferences = new ListValue<StringValue> { InnerText = "D4" }, Formula = true, EvalError = true }));
            sheet.EnsureWorksheetElementOrder();
        });

        byte[] xlsb = SaveAsXlsb(xlsx);

        using ExcelDocument reloaded = ExcelDocument.Load(new MemoryStream(xlsb));
        ExcelSheet data = reloaded["Data"];
        Assert.Equal(new[] { 20, 35 }, data.GetManualRowPageBreaks());
        Assert.Equal(new[] { 2 }, data.GetManualColumnPageBreaks());
        IgnoredError[] ignored = data.WorksheetPart.Worksheet!.GetFirstChild<IgnoredErrors>()!.Elements<IgnoredError>().ToArray();
        Assert.Equal(2, ignored.Length);
        Assert.Equal("A1:A5 C3", ignored[0].SequenceOfReferences!.InnerText);
        Assert.True(ignored[0].NumberStoredAsText?.Value);
        Assert.Null(ignored[0].Formula);
        Assert.Equal("D4", ignored[1].SequenceOfReferences!.InnerText);
        Assert.True(ignored[1].Formula?.Value);
        Assert.True(ignored[1].EvalError?.Value);
        Assert.Null(ignored[1].NumberStoredAsText);
    }

    [Fact]
    public void Editing_page_breaks_of_a_loaded_xlsb_keeps_the_other_native_records() {
        byte[] xlsx = CreateXlsx(document => {
            ExcelSheet sheet = document.AddSheet("Data");
            sheet.Cell(1, 1).SetValue("12");
            sheet.AddManualRowPageBreak(10, save: false);
            sheet.WorksheetPart.Worksheet!.Append(new IgnoredErrors(
                new IgnoredError { SequenceOfReferences = new ListValue<StringValue> { InnerText = "A1" }, NumberStoredAsText = true }));
            sheet.EnsureWorksheetElementOrder();
        });
        byte[] source = SaveAsXlsb(xlsx);

        byte[] edited;
        using (ExcelDocument loaded = ExcelDocument.Load(new MemoryStream(source))) {
            loaded["Data"].AddManualRowPageBreak(30, save: false);
            loaded["Data"].AddManualColumnPageBreak(4, save: false);
            edited = loaded.ToBytes(new ExcelSaveOptions { Format = ExcelFileFormat.Xlsb });
        }

        using ExcelDocument reloaded = ExcelDocument.Load(new MemoryStream(edited));
        Assert.Equal(new[] { 10, 30 }, reloaded["Data"].GetManualRowPageBreaks());
        Assert.Equal(new[] { 4 }, reloaded["Data"].GetManualColumnPageBreaks());
        // The unchanged ignored-error collection is copied record for record.
        Assert.Equal(ReadRecords(source, 648, 650), ReadRecords(edited, 648, 650));
    }

    [Fact]
    public void Native_xlsb_rewrite_rejects_unknown_markup_nested_in_a_header_text_element() {
        const string headerText = "&CQuarterly";
        byte[] source = SaveAsXlsb(CreateXlsx(document => {
            ExcelSheet sheet = document.AddSheet("Data");
            sheet.Cell(1, 1).SetValue(1);
            sheet.WorksheetPart.Worksheet!.Append(new HeaderFooter(new OddHeader(headerText)));
            sheet.EnsureWorksheetElementOrder();
        }));

        using ExcelDocument loaded = ExcelDocument.Load(new MemoryStream(source));
        ExcelSheet data = loaded["Data"];
        Worksheet worksheet = data.WorksheetPart.Worksheet!;
        HeaderFooter original = Assert.IsType<HeaderFooter>(worksheet.GetFirstChild<HeaderFooter>());
        XElement markup = XElement.Parse(original.OuterXml);
        markup.Element(markup.Name.Namespace + "oddHeader")!.Add(new XElement(XName.Get("vendor", "urn:xlsb:test")));
        var changed = new HeaderFooter(markup.ToString(SaveOptions.DisableFormatting));
        worksheet.ReplaceChild(changed, original);
        Assert.Equal(headerText, changed.OddHeader!.Text);
        Assert.Contains("urn:xlsb:test", changed.OuterXml, StringComparison.Ordinal);
        data.Cell(1, 1).SetValue(2);

        var exception = Assert.Throws<NotSupportedException>(() => loaded.ToBytes(new ExcelSaveOptions { Format = ExcelFileFormat.Xlsb }));
        Assert.Contains("headers and footers", exception.Message, StringComparison.Ordinal);
    }

    [Fact]
    public void Unchanged_native_save_keeps_an_edit_made_by_the_destination() {
        byte[] source = SaveAsXlsb(CreateXlsx(document => document.AddSheet("Data").Cell(1, 1).SetValue(1)));
        using ExcelDocument loaded = ExcelDocument.Load(new MemoryStream(source));
        var options = new ExcelSaveOptions { Format = ExcelFileFormat.Xlsb };
        using var destination = new CallbackMemoryStream(() => {
            _ = loaded.OpenXmlDocument;
            loaded["Data"].Cell(1, 1).SetValue(2);
        });
        loaded.Save(destination, options);
        Assert.Equal(source, destination.ToArray());

        using ExcelDocument reopened = ExcelDocument.Load(new MemoryStream(loaded.ToBytes(options)));
        Assert.Equal(2, reopened["Data"].Cell(1, 1).GetValue<int>());
    }

    [Fact]
    public void Native_save_limit_applies_to_the_xlsb_output() {
        byte[] source = SaveAsXlsb(CreateXlsx(document => document.AddSheet("Data").Cell(1, 1).SetValue(1)));
        var options = new ExcelSaveOptions { Format = ExcelFileFormat.Xlsb };
        using ExcelDocument unrestricted = ExcelDocument.Load(new MemoryStream(source));
        unrestricted["Data"].Cell(1, 1).SetValue(2);
        byte[] expected = unrestricted.ToBytes(options);
        Assert.True(unrestricted.ToBytes().Length > expected.Length);

        using ExcelDocument bounded = ExcelDocument.Load(new MemoryStream(source));
        Assert.True(bounded.IsSdkDeferred);
        bounded["Data"].Cell(1, 1).SetValue(2);
        options.MaxInMemoryPackageBytes = expected.Length;
        Assert.Equal(expected, bounded.ToBytes(options));
    }

    [Fact]
    public async Task Cancelled_native_save_leaves_the_destination_and_pending_edit_intact() {
        byte[] source = SaveAsXlsb(CreateXlsx(document => document.AddSheet("Data").Cell(1, 1).SetValue(1)));
        using ExcelDocument loaded = ExcelDocument.Load(new MemoryStream(source));
        var options = new ExcelSaveOptions { Format = ExcelFileFormat.Xlsb };
        loaded["Data"].Cell(1, 1).SetValue(2);
        byte[] marker = { 1, 2, 3 };
        using var destination = new MemoryStream();
        destination.Write(marker, 0, marker.Length);
        using var cancellation = new CancellationTokenSource();
        cancellation.Cancel();

        await Assert.ThrowsAnyAsync<OperationCanceledException>(() => loaded.SaveAsync(destination, options, cancellation.Token));
        Assert.Equal(marker, destination.ToArray());
        using ExcelDocument reopened = ExcelDocument.Load(new MemoryStream(loaded.ToBytes(options)));
        Assert.Equal(2, reopened["Data"].Cell(1, 1).GetValue<int>());
    }

    private sealed class CallbackMemoryStream(Action callback) : MemoryStream {
        public override void Write(byte[] buffer, int offset, int count) {
            callback();
            base.Write(buffer, offset, count);
        }
    }

    [Fact]
    public void Collections_added_to_a_loaded_xlsb_are_written_in_worksheet_grammar_order() {
        byte[] source = SaveAsXlsb(CreateXlsx(document => document.AddSheet("Data").Cell(1, 1).SetValue(1)));

        byte[] edited;
        using (ExcelDocument loaded = ExcelDocument.Load(new MemoryStream(source))) {
            loaded["Data"].AddManualRowPageBreak(10, save: false);
            loaded["Data"].Range("B2:B5").Validation().WholeNumber(ExcelDataValidationOperator.GreaterThan, 1);
            edited = loaded.ToBytes(new ExcelSaveOptions { Format = ExcelFileFormat.Xlsb });
        }

        int[] types = ReadAllRecordTypes(edited);
        // MS-XLSB WORKSHEET places DVALS (573) before RWBRK (392).
        Assert.True(Array.IndexOf(types, 573) < Array.IndexOf(types, 392));
        using ExcelDocument reloaded = ExcelDocument.Load(new MemoryStream(edited));
        Assert.Equal(new[] { 10 }, reloaded["Data"].GetManualRowPageBreaks());
        Assert.Single(reloaded["Data"].GetDataValidations());
    }

    [Fact]
    public void Inserted_collections_keep_an_ACUID_prefix_with_the_record_it_belongs_to() {
        // BrtEndSheetData, [ACUID] BrtHLink, BrtEndSheet: ACUID is BrtACBegin, BrtUid, BrtACEnd.
        byte[] part = WriteRecords((146, new byte[0]), (37, new byte[] { 1 }), (3072, new byte[] { 2 }), (38, new byte[0]),
            (494, new byte[] { 3 }), (130, new byte[0]));
        var edits = new[] {
            new XlsbRecordBlockEdit(392, 393, new[] { new XlsbGeneratedRecord(392, new byte[8]), new XlsbGeneratedRecord(393, new byte[0]) }),
            new XlsbRecordBlockEdit(573, 574, new[] { new XlsbGeneratedRecord(573, new byte[18]), new XlsbGeneratedRecord(574, new byte[0]) }),
        };

        byte[] result = XlsbRecordBlockEdit.Apply(part, edits, int.MaxValue);

        using var stream = new MemoryStream(result);
        Assert.Equal(new[] { 146, 573, 574, 37, 3072, 38, 494, 392, 393, 130 }, XlsbRecordReader.ReadAll(stream).Select(record => record.Type));
    }

    private static byte[] WriteRecords(params (int Type, byte[] Payload)[] records) {
        using var stream = new MemoryStream();
        foreach ((int type, byte[] payload) in records) XlsbRecordWriter.Write(stream, type, payload);
        return stream.ToArray();
    }

    private static int[] ReadAllRecordTypes(byte[] xlsb) {
        using var archive = new ZipArchive(new MemoryStream(xlsb), ZipArchiveMode.Read);
        using var part = new MemoryStream();
        using (Stream entry = archive.GetEntry("xl/worksheets/sheet1.bin")!.Open()) entry.CopyTo(part);
        part.Position = 0;
        return XlsbRecordReader.ReadAll(part).Select(record => record.Type).ToArray();
    }

    private static byte[] CreateXlsx(Action<ExcelDocument> build) {
        using ExcelDocument document = ExcelDocument.Create();
        build(document);
        return document.ToBytes();
    }

    private static byte[] SaveAsXlsb(byte[] xlsx) {
        using ExcelDocument loaded = ExcelDocument.Load(new MemoryStream(xlsx));
        return loaded.ToBytes(new ExcelSaveOptions { Format = ExcelFileFormat.Xlsb });
    }

    private static string[] ReadRecords(byte[] xlsb, int beginType, int endType) {
        using var archive = new ZipArchive(new MemoryStream(xlsb), ZipArchiveMode.Read);
        using var part = new MemoryStream();
        using (Stream entry = archive.GetEntry("xl/worksheets/sheet1.bin")!.Open()) entry.CopyTo(part);
        part.Position = 0;
        var records = XlsbRecordReader.ReadAll(part).ToList();
        int begin = records.FindIndex(record => record.Type == beginType);
        int end = records.FindIndex(record => record.Type == endType);
        Assert.True(begin >= 0 && end > begin);
        return records.Skip(begin).Take(end - begin + 1).Select(record => record.Type + ":" + Convert.ToBase64String(record.Data)).ToArray();
    }
}
