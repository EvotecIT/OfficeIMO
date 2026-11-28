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

var creations = new SortedDictionary<string, string>(StringComparer.Ordinal);
Directory.CreateDirectory(Path.Combine(output, "generated"));
foreach (bool sharedStrings in new[] { false, true }) {
    using var plain = ExcelDocument.Create();
    ExcelSheet sheet = plain.AddSheet("Data");
    for (int row = 1; row <= 64; row++) {
        sheet.Cell(row, 1).SetValue(row + 2.5D);
        sheet.Cell(row, 2).SetValue("row " + row);
        sheet.Cell(row, 3).SetValue((row & 1) == 0);
    }
    string name = sharedStrings ? "scalars-shared.xlsb" : "scalars-inline.xlsb";
    string path = Path.Combine("generated", name);
    File.WriteAllBytes(Path.Combine(output, path), plain.ToBytes(new ExcelSaveOptions {
        Format = ExcelFileFormat.Xlsb,
        XlsbUseSharedStrings = sharedStrings
    }));
    creations.Add(name, path.Replace(Path.DirectorySeparatorChar, '/'));
}

using (var styled = ExcelDocument.Create()) {
    ExcelSheet sheet = styled.AddSheet("Styled");
    var cell = sheet.Cell(1, 1);
    cell.SetValue(1.25D);
    cell.SetBold();
    cell.SetItalic();
    string path = Path.Combine("generated", "styled.xlsb");
    File.WriteAllBytes(Path.Combine(output, path), styled.ToBytes(new ExcelSaveOptions { Format = ExcelFileFormat.Xlsb }));
    creations.Add("styled.xlsb", path.Replace(Path.DirectorySeparatorChar, '/'));
}
File.WriteAllText(Path.Combine(output, "creations.json"), JsonSerializer.Serialize(creations,
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
