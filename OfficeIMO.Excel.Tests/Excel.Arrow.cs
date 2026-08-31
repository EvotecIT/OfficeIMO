#if NET8_0_OR_GREATER
using Apache.Arrow;
using Apache.Arrow.Arrays;
using OfficeIMO.Data.Arrow;
using OfficeIMO.Excel;
using System.IO.Compression;
using System.Text;
using Xunit;

namespace OfficeIMO.Tests;

public partial class Excel {
    [Fact]
    public void DataReader_ArrowFastPathPreservesTypedValuesAndNulls() {
        DateTime expectedDate = new(2026, 8, 31, 14, 30, 0, DateTimeKind.Unspecified);
        using var memory = new MemoryStream();
        using (var document = ExcelDocument.Create(
                   memory,
                   new ExcelCreateOptions { PersistenceMode = OfficeIMO.DocumentPersistenceMode.SaveOnDispose })) {
            ExcelSheet sheet = document.AddWorksheet("Data");
            sheet.CellValue(1, 1, "Name");
            sheet.CellValue(1, 2, "Id");
            sheet.CellValue(1, 3, "Amount");
            sheet.CellValue(1, 4, "When");
            sheet.CellValue(2, 1, "Alpha & Beta");
            sheet.CellValue(2, 2, 7);
            sheet.CellValue(2, 3, 12.5d);
            sheet.CellValue(2, 4, expectedDate);
            sheet.CellValue(3, 1, "Gamma");
            sheet.CellValue(3, 2, 8);
        }

        using ExcelWorkbookDataReader reader = ExcelDocument.OpenDataReader(memory.ToArray());
        RecordBatch batch = Assert.Single(reader.ReadArrowBatches(new ArrowReadOptions {
            BatchSize = 16,
            ColumnTypes = [typeof(string), typeof(long), typeof(double), typeof(DateTime)]
        }));
        using (batch) {
            Assert.Equal(2, batch.Length);
            Assert.Equal("Alpha & Beta", Assert.IsType<StringArray>(batch.Column(0)).GetString(0));
            Assert.Equal("Gamma", Assert.IsType<StringArray>(batch.Column(0)).GetString(1));
            Assert.Equal(7L, Assert.IsType<Int64Array>(batch.Column(1)).GetValue(0));
            Assert.Equal(8L, Assert.IsType<Int64Array>(batch.Column(1)).GetValue(1));
            Assert.Equal(12.5d, Assert.IsType<DoubleArray>(batch.Column(2)).GetValue(0));
            Assert.Null(Assert.IsType<DoubleArray>(batch.Column(2)).GetValue(1));
            Assert.Equal(expectedDate, Assert.IsType<TimestampArray>(batch.Column(3)).GetTimestamp(0)?.DateTime);
            Assert.Null(Assert.IsType<TimestampArray>(batch.Column(3)).GetTimestamp(1));
        }
    }

    [Fact]
    public void DataReader_ArrowFastPathNormalizesLiteralXmlLineEndingsLikeGetString() {
        using var memory = new MemoryStream();
        using (var document = ExcelDocument.Create(
                   memory,
                   new ExcelCreateOptions { PersistenceMode = OfficeIMO.DocumentPersistenceMode.SaveOnDispose })) {
            ExcelSheet sheet = document.AddWorksheet("Data");
            sheet.CellValue(1, 1, "Text");
            sheet.CellValue(2, 1, "Placeholder");
        }

        byte[] workbook = ReplaceArrowTestCellXml(
            memory.ToArray(),
            "<c r=\"A2\" t=\"inlineStr\"><is><t>Line1\r\nLine2\rLine3</t></is></c>");
        const string expected = "Line1\nLine2\nLine3";

        using (ExcelWorkbookDataReader reader = ExcelDocument.OpenDataReader(workbook)) {
            Assert.True(reader.Read());
            Assert.Equal(expected, reader.GetString(0));
        }

        using ExcelWorkbookDataReader arrowReader = ExcelDocument.OpenDataReader(workbook);
        RecordBatch batch = Assert.Single(arrowReader.ReadArrowBatches(new ArrowReadOptions {
            ColumnTypes = [typeof(string)]
        }));
        using (batch) {
            Assert.Equal(expected, Assert.IsType<StringArray>(batch.Column(0)).GetString(0));
        }
    }

    private static byte[] ReplaceArrowTestCellXml(byte[] workbook, string replacementCellXml) {
        using var package = new MemoryStream();
        package.Write(workbook, 0, workbook.Length);
        package.Position = 0;
        using (var archive = new ZipArchive(package, ZipArchiveMode.Update, leaveOpen: true)) {
            ZipArchiveEntry entry = archive.GetEntry("xl/worksheets/sheet1.xml")
                ?? throw new InvalidDataException("The generated workbook has no first worksheet part.");
            string xml;
            using (var reader = new StreamReader(entry.Open(), Encoding.UTF8, detectEncodingFromByteOrderMarks: true)) {
                xml = reader.ReadToEnd();
            }

            int reference = xml.IndexOf("r=\"A2\"", StringComparison.Ordinal);
            int cellStart = reference < 0 ? -1 : xml.LastIndexOf("<c", reference, StringComparison.Ordinal);
            int cellEnd = cellStart < 0 ? -1 : xml.IndexOf("</c>", cellStart, StringComparison.Ordinal);
            if (cellStart < 0 || cellEnd < 0) {
                throw new InvalidDataException("The generated workbook has no A2 cell to replace.");
            }

            string changed = xml.Substring(0, cellStart)
                + replacementCellXml
                + xml.Substring(cellEnd + 4);
            entry.Delete();
            ZipArchiveEntry replacement = archive.CreateEntry(
                "xl/worksheets/sheet1.xml",
                CompressionLevel.Optimal);
            using var writer = new StreamWriter(replacement.Open(), new UTF8Encoding(false));
            writer.Write(changed);
        }

        return package.ToArray();
    }
}
#endif
