using System.Text.Json;
using OfficeIMO.Excel;

if (args.Length != 2) throw new ArgumentException("Pass the corpus and output directories.");
string corpus = Path.GetFullPath(args[0]);
string output = Path.GetFullPath(args[1]);
Directory.CreateDirectory(output);
var results = new SortedDictionary<string, Result>(StringComparer.Ordinal);
foreach (string input in Directory.EnumerateFiles(corpus, "*.xlsb", SearchOption.AllDirectories).Order(StringComparer.Ordinal)) {
    string name = Path.GetRelativePath(corpus, input);
    string path = Path.Combine(output, name);
    byte[] bytes = File.ReadAllBytes(input);
    Result result = new(Export(bytes, path + ".xlsx", ExcelFileFormat.Xlsx),
        Export(bytes, path + ".xlsb", ExcelFileFormat.Xlsb),
        Export(bytes, path + ".editedxlsb", ExcelFileFormat.Xlsb, edit: true));
    results.Add(name.Replace(Path.DirectorySeparatorChar, '/'), result);
}
File.WriteAllText(Path.Combine(output, "results.json"), JsonSerializer.Serialize(results,
    new JsonSerializerOptions { WriteIndented = true }));

static string? Export(byte[] bytes, string path, ExcelFileFormat format, bool edit = false) {
    try {
        using var source = new MemoryStream(bytes, writable: false);
        using ExcelDocument document = ExcelDocument.Load(source);
        if (edit) {
            ExcelSheet sheet = document.Sheets.FirstOrDefault()
                ?? throw new NotSupportedException("The workbook has no editable worksheet.");
            sheet.Cell(1024, 27).SetValue(314159);
        }
        using var destination = new MemoryStream();
        document.Save(destination, new ExcelSaveOptions { Format = format });
        Directory.CreateDirectory(Path.GetDirectoryName(path)!);
        File.WriteAllBytes(path, destination.ToArray());
        return null;
    } catch (Exception exception) when (exception is ArgumentException or FormatException or IOException
        or InvalidDataException or InvalidOperationException or NotSupportedException or System.Xml.XmlException) {
        return exception.GetType().FullName + ": " + exception.Message;
    }
}

internal sealed record Result(string? Xlsx, string? Xlsb, string? EditedXlsb);
