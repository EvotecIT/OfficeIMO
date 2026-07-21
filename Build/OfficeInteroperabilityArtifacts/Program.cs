using System.Security.Cryptography;
using System.Text;
using System.Text.Json;
using OfficeIMO.Excel;
using OfficeIMO.PowerPoint;
using OfficeIMO.Word;

string outputDirectory = GetOption(args, "--output")
    ?? throw new ArgumentException("--output requires a destination directory.");
outputDirectory = Path.GetFullPath(outputDirectory);
Directory.CreateDirectory(outputDirectory);

var artifacts = new List<object>();
CreateWordArtifact(outputDirectory, artifacts);
CreateExcelArtifacts(outputDirectory, artifacts);
CreatePowerPointArtifact(outputDirectory, artifacts);

var manifest = new {
    schemaVersion = 1,
    producer = "OfficeIMO",
    artifacts
};
string manifestPath = Path.Combine(outputDirectory, "officeimo-artifacts.json");
File.WriteAllText(
    manifestPath,
    JsonSerializer.Serialize(manifest, new JsonSerializerOptions {
        PropertyNamingPolicy = JsonNamingPolicy.CamelCase,
        WriteIndented = true
    }) + Environment.NewLine,
    new UTF8Encoding(encoderShouldEmitUTF8Identifier: false));
Console.WriteLine($"Generated {artifacts.Count} OfficeIMO binary interoperability artifacts in {outputDirectory}.");

static void CreateWordArtifact(string outputDirectory, ICollection<object> artifacts) {
    string source = Path.Combine(outputDirectory, "officeimo-word-source.docx");
    string destination = Path.Combine(outputDirectory, "officeimo-word-generated.doc");
    using (WordDocument document = WordDocument.Create(source)) {
        document.AddParagraph("OfficeIMO Word modern-to-binary interoperability artifact");
        document.Save();
    }
    WordDocument.Convert(source, destination).RequireNoLoss();
    artifacts.Add(CreateArtifact("Word.Doc", "doc", destination));
}

static void CreateExcelArtifacts(string outputDirectory, ICollection<object> artifacts) {
    string source = Path.Combine(outputDirectory, "officeimo-excel-source.xlsx");
    using (ExcelDocument document = ExcelDocument.Create(source)) {
        ExcelSheet sheet = document.AddWorksheet("Compatibility");
        sheet.CellValue(1, 1, "OfficeIMO Excel modern-to-binary interoperability artifact");
        sheet.CellValue(2, 1, 42);
        document.Save();
    }
    foreach ((string formatId, string extension) in new[] {
                 ("Excel.Xls", ".xls"),
                 ("Excel.Xlsb", ".xlsb")
             }) {
        string destination = Path.Combine(outputDirectory, "officeimo-excel-generated" + extension);
        ExcelDocument.Convert(source, destination).RequireNoLoss();
        artifacts.Add(CreateArtifact(formatId, extension.TrimStart('.'), destination));
    }
}

static void CreatePowerPointArtifact(string outputDirectory, ICollection<object> artifacts) {
    string source = Path.Combine(outputDirectory, "officeimo-powerpoint-source.pptx");
    string destination = Path.Combine(outputDirectory, "officeimo-powerpoint-generated.ppt");
    using (PowerPointPresentation presentation = PowerPointPresentation.Create(source)) {
        presentation.AddSlide().AddTextBox("OfficeIMO PowerPoint modern-to-binary interoperability artifact");
        presentation.Save();
    }
    PowerPointPresentation.Convert(source, destination).RequireNoLoss();
    artifacts.Add(CreateArtifact("PowerPoint.Ppt", "ppt", destination));
}

static object CreateArtifact(string formatId, string format, string path) => new {
    formatId,
    format,
    file = Path.GetFileName(path),
    sha256 = Convert.ToHexString(SHA256.HashData(File.ReadAllBytes(path))).ToLowerInvariant()
};

static string? GetOption(string[] values, string name) {
    for (int index = 0; index < values.Length; index++) {
        if (!string.Equals(values[index], name, StringComparison.OrdinalIgnoreCase)) continue;
        if (index + 1 >= values.Length) throw new ArgumentException($"{name} requires a value.");
        return values[index + 1];
    }
    return null;
}
