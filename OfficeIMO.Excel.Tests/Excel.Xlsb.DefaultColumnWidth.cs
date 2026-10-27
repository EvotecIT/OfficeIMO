using DocumentFormat.OpenXml.Spreadsheet;
using OfficeIMO.Excel;
using OfficeIMO.Excel.Xlsb.Biff12;
using Xunit;

namespace OfficeIMO.Tests;

public partial class Excel {
    [Theory]
    [InlineData(8U)]
    [InlineData(12U)]
    public void Xlsb_Import_Keeps_Implicit_Default_Column_Width_Unset(uint baseWidth) {
        byte[] source = File.ReadAllBytes(GetGeometryExcelGeneratedXlsbFixturePath());
        using var recordsStream = new MemoryStream(ReadZipEntry(source, "xl/worksheets/sheet1.bin"));
        IReadOnlyList<XlsbRecord> records = XlsbRecordReader.ReadAll(recordsStream);
        XlsbRecord format = Assert.Single(records, record => record.Type == 485);
        byte[] data = (byte[])format.Data.Clone();
        WriteXlsbTestUInt32(data, 0, uint.MaxValue);
        data[4] = (byte)baseWidth;
        data[5] = 0;
        byte[] implicitSource = ReplaceWorksheetRecords(source, records, format, data);

        using ExcelDocument document = ExcelDocument.Load(new MemoryStream(implicitSource, writable: false));
        SheetFormatProperties projected = Assert.IsType<SheetFormatProperties>(
            document.Sheets[0].WorksheetPart.Worksheet.GetFirstChild<SheetFormatProperties>());

        Assert.Null(projected.DefaultColumnWidth);
        Assert.Equal(baseWidth, projected.BaseColumnWidth?.Value);
    }

    [Theory]
    [InlineData(8U)]
    [InlineData(12U)]
    public void Xlsb_New_Workbook_Writes_Implicit_Default_Column_Width_Sentinel(uint baseWidth) {
        using ExcelDocument document = ExcelDocument.Create();
        ExcelSheet sheet = document.AddWorksheet("Geometry");
        sheet.CellValue(1, 1, "Retained");
        sheet.SetDefaultRowHeight(ExcelColumnWidth.Length(OfficeIMO.OfficeLength.Points(15D)));
        sheet.WorksheetPart.Worksheet.GetFirstChild<SheetFormatProperties>()!.BaseColumnWidth = baseWidth;

        byte[] package = document.ToBytes(new ExcelSaveOptions { Format = ExcelFileFormat.Xlsb });
        using var stream = new MemoryStream(ReadZipEntry(package, "xl/worksheets/sheet1.bin"));
        XlsbRecord format = Assert.Single(XlsbRecordReader.ReadAll(stream), record => record.Type == 485);

        Assert.Equal(uint.MaxValue, ReadXlsbTestUInt32(format.Data, 0));
        Assert.Equal(baseWidth, (uint)(format.Data[4] | format.Data[5] << 8));
        using ExcelDocument reloaded = ExcelDocument.Load(new MemoryStream(package, writable: false));
        Assert.Null(reloaded.Sheets[0].DefaultColumnWidth);
    }

    [Theory]
    [InlineData(0D)]
    [InlineData(9.253D)]
    [InlineData(255.99609375D)]
    public void Xlsb_Authored_Default_Column_Width_Preserves_Encoded_Value(double width) {
        using ExcelDocument document = ExcelDocument.Create();
        ExcelSheet sheet = document.AddWorksheet("Geometry");
        sheet.CellValue(1, 1, "Retained");
        sheet.SetDefaultRowHeight(ExcelColumnWidth.Length(OfficeIMO.OfficeLength.Points(15D)));
        sheet.WorksheetPart.Worksheet.GetFirstChild<SheetFormatProperties>()!.DefaultColumnWidth = width;

        byte[] package = document.ToBytes(new ExcelSaveOptions { Format = ExcelFileFormat.Xlsb });
        using var stream = new MemoryStream(ReadZipEntry(package, "xl/worksheets/sheet1.bin"));
        XlsbRecord format = Assert.Single(XlsbRecordReader.ReadAll(stream), record => record.Type == 485);
        uint expected = checked((uint)Math.Floor(width * 256D));

        Assert.Equal(expected, ReadXlsbTestUInt32(format.Data, 0));
        using ExcelDocument reloaded = ExcelDocument.Load(new MemoryStream(package, writable: false));
        Assert.Equal(ExcelColumnWidth.Characters(expected / 256D), reloaded.Sheets[0].DefaultColumnWidth);
    }

    [Fact]
    public void Xlsb_Explicit_Default_Column_Width_Ignores_Fallback_Field() {
        byte[] source = File.ReadAllBytes(GetGeometryExcelGeneratedXlsbFixturePath());
        using var stream = new MemoryStream(ReadZipEntry(source, "xl/worksheets/sheet1.bin"));
        IReadOnlyList<XlsbRecord> records = XlsbRecordReader.ReadAll(stream);
        XlsbRecord format = Assert.Single(records, record => record.Type == 485);
        byte[] data = (byte[])format.Data.Clone();
        WriteXlsbTestUInt32(data, 0, 2368U);
        data[4] = data[5] = 255;
        byte[] package = ReplaceWorksheetRecords(source, records, format, data);

        using ExcelDocument document = ExcelDocument.Load(new MemoryStream(package, writable: false));
        SheetFormatProperties projected = Assert.IsType<SheetFormatProperties>(
            document.Sheets[0].WorksheetPart.Worksheet.GetFirstChild<SheetFormatProperties>());

        Assert.Equal(9.25D, projected.DefaultColumnWidth?.Value);
        Assert.Null(projected.BaseColumnWidth);
    }
}
