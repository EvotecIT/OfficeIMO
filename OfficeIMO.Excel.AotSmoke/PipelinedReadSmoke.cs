using OfficeIMO.Data;
using OfficeIMO.Excel;
using System.IO.Compression;
using System.Text;

internal static class PipelinedReadSmoke {
    private const int Count = 50_000;

    internal static void Run() {
        using var output = new MemoryStream();
        using (var writer = ExcelDocument.CreateWriter(output,
            new ExcelWriterOptions { IncludeCellReferences = false, UseSharedStrings = false, CreateTable = false }, leaveOpen: true)) {
            writer.WriteTable("Large", Rows(), ["Name", "Id", "CreatedOn", "Amount"],
                static (rowWriter, row) => rowWriter.Write(row.Name).Write(row.Id).Write(row.CreatedOn).Write(row.Amount));
        }
        byte[] bytes = output.ToArray();
        using (var input = new MemoryStream(bytes))
        using (var reader = ExcelDocument.OpenReader(input)) {
            int count = 0;
            while (reader.Read()) {
                Row expected = Create(count++);
                if (reader.GetString(0) != expected.Name || reader.GetInt32(1) != expected.Id
                    || reader.GetDateTime(2) != expected.CreatedOn || reader.GetDouble(3) != expected.Amount)
                    throw new InvalidOperationException("The large compressed reader lost an indexed value.");
            }
            if (count != Count) throw new InvalidOperationException("The large compressed reader lost rows.");
        }
        using (var input = new MemoryStream(bytes))
        using (var reader = ExcelDocument.OpenReader(input)) {
            int count = 0;
            foreach (Row row in reader.RowsAs<Row>()) {
                if (row != Create(count++)) throw new InvalidOperationException("The large compressed reader lost a typed row.");
            }
            if (count != Count) throw new InvalidOperationException("The large compressed typed reader lost rows.");
        }
        Console.WriteLine("PASS | Large compressed implicit worksheets retain escaped text and every scalar and typed row under NativeAOT");
        RunLoadedPrefixedSnapshot();
    }

    private static void RunLoadedPrefixedSnapshot() {
        using var output = new MemoryStream();
        using (var writer = ExcelDocument.CreateWriter(output,
            new ExcelWriterOptions { IncludeCellReferences = true, UseSharedStrings = false, CreateTable = false }, leaveOpen: true)) {
            writer.WriteTable("Large", Rows(), ["Name", "Id", "CreatedOn", "Amount"],
                static (rowWriter, row) => rowWriter.Write(row.Name).Write(row.Id).Write(row.CreatedOn).Write(row.Amount));
        }
        output.Position = 0;
        using (var archive = new ZipArchive(output, ZipArchiveMode.Update, leaveOpen: true)) {
            var entry = archive.GetEntry("xl/worksheets/sheet1.xml")!;
            string xml;
            using (var text = new StreamReader(entry.Open())) xml = text.ReadToEnd();
            entry.Delete();
            const string spreadsheet = "http://schemas.openxmlformats.org/spreadsheetml/2006/main";
            xml = xml.Replace("<worksheet ", $"<p:worksheet xmlns:p=\"{spreadsheet}\" ");
            if (!xml.Contains("<dimension ", StringComparison.Ordinal))
                xml = xml.Replace("<sheetData>", $"<dimension ref=\"A1:D{Count + 1}\"/><sheetData>");
            foreach (string tag in new[] { "worksheet", "dimension", "sheetData", "row", "c", "v", "is", "t", "f" }) {
                xml = xml.Replace("<" + tag + ">", "<p:" + tag + ">")
                    .Replace("<" + tag + " ", "<p:" + tag + " ")
                    .Replace("</" + tag + ">", "</p:" + tag + ">");
            }
            using var textWriter = new StreamWriter(archive.CreateEntry("xl/worksheets/sheet1.xml").Open(), new UTF8Encoding(false));
            textWriter.Write(xml);
        }
        using var document = ExcelDocument.Load(new MemoryStream(output.ToArray()));
        using var reader = document.CreateDataReader(new ExcelReaderOptions { InferSchema = false });
        int count = 0;
        while (reader.Read()) {
            Row expected = Create(count++);
            if (reader.GetString(0) != expected.Name || reader.GetInt32(1) != expected.Id
                || reader.GetDateTime(2) != expected.CreatedOn || reader.GetDouble(3) != expected.Amount)
                throw new InvalidOperationException("The loaded prefixed snapshot lost a value.");
        }
        if (count != Count) throw new InvalidOperationException("The loaded prefixed snapshot lost rows.");
        Console.WriteLine("PASS | Large prefixed loaded snapshots retain every typed value under NativeAOT");
    }

    private static Row Create(int index) => new("Name <" + index + "> & \"quoted\" 'apostrophe'", index + 1,
        new DateTime(2024, 1, 1).AddDays(index % 365), index * 1.25);

    private static IEnumerable<Row> Rows() {
        for (int index = 0; index < Count; index++) yield return Create(index);
    }

    private sealed record Row(string Name, int Id, DateTime CreatedOn, double Amount);
}
