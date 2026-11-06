using System.IO.Compression;
using OfficeIMO.Excel;

/// <summary>Exercises buffered XLSB compression under NativeAOT.</summary>
internal static class BufferedXlsbSaveSmoke {
    internal static void Run() {
        const int rowCount = 4096;
        using ExcelDocument document = ExcelDocument.Create();
        document.AddSheet("Buffered").WriteData(Enumerable.Range(0, rowCount),
            ("Id", static row => row), ("Label", static row => "_x0041_ & " + row));
        byte[] bytes = document.ToBytes(new ExcelSaveOptions { Format = ExcelFileFormat.Xlsb });
        using var source = new MemoryStream(bytes, writable: false);
        using (var package = new ZipArchive(source, ZipArchiveMode.Read, leaveOpen: true)) {
            if (package.GetEntry("xl/worksheets/sheet1.bin") == null)
                throw new InvalidOperationException("The buffered XLSB smoke did not produce a binary worksheet.");
        }
        source.Position = 0;
        using ExcelWorkbookDataReader reader = ExcelDocument.OpenReader(source);
        if (reader.FieldCount != 2 || reader.GetName(0) != "Id" || reader.GetName(1) != "Label")
            throw new InvalidOperationException("The buffered XLSB smoke lost its schema.");
        for (int row = 0; row < rowCount; row++) {
            if (!reader.Read() || reader.GetInt32(0) != row || reader.GetString(1) != "_x0041_ & " + row)
                throw new InvalidOperationException("The buffered XLSB smoke lost values or their order.");
        }
        if (reader.Read() || reader.NextResult())
            throw new InvalidOperationException("The buffered XLSB smoke added rows or worksheets.");
        Console.WriteLine("PASS | Buffered XLSB NativeAOT save: 4096 ordered rows.");
    }
}
