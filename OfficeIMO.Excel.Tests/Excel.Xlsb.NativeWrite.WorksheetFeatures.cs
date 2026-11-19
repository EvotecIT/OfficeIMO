using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Spreadsheet;
using OfficeIMO.Excel;
using OfficeIMO.Excel.Xlsb.Biff12;
using System.IO.Compression;
using Xunit;

namespace OfficeIMO.Tests;

/// <summary>Worksheet content after the cell table that the native XLSB writers encode and the reader projects.</summary>
public sealed class ExcelXlsbWorksheetFeatureTests {
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
