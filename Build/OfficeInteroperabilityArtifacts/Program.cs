using System.Security.Cryptography;
using System.Text;
using System.Text.Json;
using System.Text.Json.Serialization;
using OfficeIMO.Excel;
using OfficeIMO.PowerPoint;
using OfficeIMO.Word;

string outputDirectory = GetOption(args, "--output")
    ?? throw new ArgumentException("--output requires a destination directory.");
outputDirectory = Path.GetFullPath(outputDirectory);
Directory.CreateDirectory(outputDirectory);

var artifacts = new List<InteroperabilityArtifact>();
CreateWordArtifact(outputDirectory, artifacts);
CreateExcelArtifacts(outputDirectory, artifacts);
CreatePowerPointArtifact(outputDirectory, artifacts);

var manifest = new InteroperabilityManifest(1, "OfficeIMO", artifacts);
string manifestPath = Path.Combine(outputDirectory, "officeimo-artifacts.json");
File.WriteAllText(
    manifestPath,
    JsonSerializer.Serialize(manifest, InteroperabilityJsonSerializerContext.Default.InteroperabilityManifest) + Environment.NewLine,
    new UTF8Encoding(encoderShouldEmitUTF8Identifier: false));
Console.WriteLine($"Generated {artifacts.Count} OfficeIMO binary interoperability artifacts in {outputDirectory}.");

static void CreateWordArtifact(string outputDirectory, ICollection<InteroperabilityArtifact> artifacts) {
    string source = Path.Combine(outputDirectory, "officeimo-word-source.docx");
    string destination = Path.Combine(outputDirectory, "officeimo-word-generated.doc");
    using (WordDocument document = WordDocument.Create(source)) {
        document.AddParagraph("OfficeIMO Word modern-to-binary interoperability artifact");
        document.Save();
    }
    WordDocument.Convert(source, destination).RequireNoLoss();
    artifacts.Add(CreateArtifact("Word.Doc", "doc", destination));
}

static void CreateExcelArtifacts(string outputDirectory, ICollection<InteroperabilityArtifact> artifacts) {
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

static void CreatePowerPointArtifact(string outputDirectory, ICollection<InteroperabilityArtifact> artifacts) {
    string source = Path.Combine(outputDirectory, "officeimo-powerpoint-source.pptx");
    string destination = Path.Combine(outputDirectory, "officeimo-powerpoint-generated.ppt");
    using (PowerPointPresentation presentation = PowerPointPresentation.Create(source)) {
        presentation.AddSlide().AddTextBox("OfficeIMO PowerPoint modern-to-binary interoperability artifact");
        presentation.Save();
    }
    PowerPointPresentation.Convert(source, destination).RequireNoLoss();
    artifacts.Add(CreateArtifact("PowerPoint.Ppt", "ppt", destination));
}

static InteroperabilityArtifact CreateArtifact(string formatId, string format, string path) => new(
    formatId,
    format,
    Path.GetFileName(path),
    Convert.ToHexString(SHA256.HashData(File.ReadAllBytes(path))).ToLowerInvariant());

static string? GetOption(string[] values, string name) {
    for (int index = 0; index < values.Length; index++) {
        if (!string.Equals(values[index], name, StringComparison.OrdinalIgnoreCase)) continue;
        if (index + 1 >= values.Length) throw new ArgumentException($"{name} requires a value.");
        return values[index + 1];
    }
    return null;
}

internal sealed record InteroperabilityManifest(int SchemaVersion, string Producer, IReadOnlyList<InteroperabilityArtifact> Artifacts);
internal sealed record InteroperabilityArtifact(string FormatId, string Format, string File, string Sha256);

[JsonSourceGenerationOptions(PropertyNamingPolicy = JsonKnownNamingPolicy.CamelCase, WriteIndented = true)]
[JsonSerializable(typeof(InteroperabilityManifest))]
internal sealed partial class InteroperabilityJsonSerializerContext : JsonSerializerContext {
}
