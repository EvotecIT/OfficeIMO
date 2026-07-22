using System.Text;
using System.Text.Json;
using System.Text.Json.Serialization;
using OfficeIMO.Drawing;
using OfficeIMO.Excel;
using OfficeIMO.PowerPoint;
using OfficeIMO.Word;

string outputDirectory = GetOption(args, "--output")
    ?? Path.Combine(Directory.GetCurrentDirectory(), "Docs", "Compatibility", "generated");
bool verify = args.Contains("--verify", StringComparer.OrdinalIgnoreCase);

var capabilityCatalogs = new (string Name, OfficeCapabilityCatalog Catalog)[] {
    ("word-legacy-doc", WordCompatibilityCatalog.Current),
    ("excel-legacy-xls", ExcelCompatibilityCatalog.Xls),
    ("excel-xlsb", ExcelCompatibilityCatalog.Xlsb),
    ("powerpoint-legacy-ppt", PowerPointCompatibilityCatalog.Current)
};
var outputs = new SortedDictionary<string, string>(StringComparer.Ordinal) {
    ["office-formats.json"] = SerializeFormats(),
    ["README.md"] = CreateReadme(capabilityCatalogs)
};
foreach ((string name, OfficeCapabilityCatalog catalog) in capabilityCatalogs) {
    outputs[name + ".json"] = EnsureFinalNewline(catalog.ToJson());
    outputs[name + ".md"] = EnsureFinalNewline(catalog.ToMarkdown());
}

if (verify) {
    var stale = new List<string>();
    foreach ((string fileName, string expected) in outputs) {
        string path = Path.Combine(outputDirectory, fileName);
        if (!File.Exists(path) || Normalize(File.ReadAllText(path)) != Normalize(expected)) stale.Add(fileName);
    }
    if (stale.Count > 0) {
        Console.Error.WriteLine("Compatibility catalog outputs are missing or stale: " + string.Join(", ", stale));
        Environment.ExitCode = 1;
        return;
    }
    Console.WriteLine($"Verified {outputs.Count} compatibility catalog artifacts in {Path.GetFullPath(outputDirectory)}.");
    return;
}

Directory.CreateDirectory(outputDirectory);
foreach ((string fileName, string content) in outputs) {
    File.WriteAllText(Path.Combine(outputDirectory, fileName), Normalize(content), new UTF8Encoding(encoderShouldEmitUTF8Identifier: false));
}
Console.WriteLine($"Generated {outputs.Count} compatibility catalog artifacts in {Path.GetFullPath(outputDirectory)}.");

static string SerializeFormats() {
    var model = new FormatCatalogModel(
        1,
        new[] {
            CreateFamily("Word", WordFormatCatalog.All),
            CreateFamily("Excel", ExcelFormatCatalog.All),
            CreateFamily("PowerPoint", PowerPointFormatCatalog.All)
        });
    return EnsureFinalNewline(JsonSerializer.Serialize(model, CompatibilityCatalogJsonSerializerContext.Default.FormatCatalogModel));
}

static FormatFamilyModel CreateFamily(string id, IReadOnlyList<OfficeFormatDescriptor> formats) => new(
    id,
    formats.Select(format => new FormatDescriptorModel(
        format.Id,
        format.Extension,
        format.Family.ToString(),
        format.DocumentKind.ToString(),
        format.Generation.ToString(),
        format.Encoding.ToString(),
        format.IsMacroEnabled)).ToArray());

static string CreateReadme(IEnumerable<(string Name, OfficeCapabilityCatalog Catalog)> catalogs) {
    var markdown = new StringBuilder();
    markdown.AppendLine("# Generated Office compatibility contracts");
    markdown.AppendLine();
    markdown.AppendLine("These files are generated from the public format and capability catalogs. Do not edit them by hand.");
    markdown.AppendLine();
    markdown.AppendLine("Regenerate:");
    markdown.AppendLine();
    markdown.AppendLine("```powershell");
    markdown.AppendLine("dotnet run --project Build/CompatibilityCatalog/OfficeIMO.CompatibilityCatalog.Tool.csproj -- --output Docs/Compatibility/generated");
    markdown.AppendLine("```");
    markdown.AppendLine();
    markdown.AppendLine("Verify:");
    markdown.AppendLine();
    markdown.AppendLine("```powershell");
    markdown.AppendLine("dotnet run --project Build/CompatibilityCatalog/OfficeIMO.CompatibilityCatalog.Tool.csproj -- --output Docs/Compatibility/generated --verify");
    markdown.AppendLine("```");
    markdown.AppendLine();
    markdown.AppendLine("| Contract | Schema | Rows | JSON | Markdown |");
    markdown.AppendLine("| --- | ---: | ---: | --- | --- |");
    foreach ((string name, OfficeCapabilityCatalog catalog) in catalogs) {
        markdown.Append("| ").Append(catalog.Id)
            .Append(" | ").Append(catalog.SchemaVersion)
            .Append(" | ").Append(catalog.Capabilities.Count)
            .Append(" | [JSON](").Append(name).Append(".json)")
            .Append(" | [Markdown](").Append(name).AppendLine(".md) |");
    }
    markdown.AppendLine();
    markdown.AppendLine("`office-formats.json` is the concrete extension, document-kind, encoding, and macro-carrier inventory used by conversion routing.");
    return EnsureFinalNewline(markdown.ToString());
}

static string? GetOption(string[] values, string name) {
    for (int index = 0; index < values.Length; index++) {
        if (!string.Equals(values[index], name, StringComparison.OrdinalIgnoreCase)) continue;
        if (index + 1 >= values.Length) throw new ArgumentException($"{name} requires a value.");
        return values[index + 1];
    }
    return null;
}

static string EnsureFinalNewline(string value) => Normalize(value).TrimEnd('\n') + "\n";
static string Normalize(string value) => value.Replace("\r\n", "\n").Replace("\r", "\n");

internal sealed record FormatCatalogModel(int SchemaVersion, IReadOnlyList<FormatFamilyModel> Families);
internal sealed record FormatFamilyModel(string Id, IReadOnlyList<FormatDescriptorModel> Formats);
internal sealed record FormatDescriptorModel(
    string Id,
    string Extension,
    string Family,
    string DocumentKind,
    string Generation,
    string Encoding,
    bool IsMacroEnabled);

[JsonSourceGenerationOptions(PropertyNamingPolicy = JsonKnownNamingPolicy.CamelCase, WriteIndented = true)]
[JsonSerializable(typeof(FormatCatalogModel))]
internal sealed partial class CompatibilityCatalogJsonSerializerContext : JsonSerializerContext {
}
