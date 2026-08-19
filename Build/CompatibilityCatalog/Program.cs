using System.Text;
using System.Text.Json;
using System.Text.Json.Serialization;
using OfficeIMO;
using OfficeIMO.Drawing;
using OfficeIMO.Excel;
using OfficeIMO.PowerPoint;
using OfficeIMO.Security;
using OfficeIMO.Word;

string outputDirectory = GetOption(args, "--output")
    ?? Path.Combine(Directory.GetCurrentDirectory(), "Docs", "Compatibility", "generated");
string websiteDataPath = GetOption(args, "--website-data")
    ?? Path.Combine(Directory.GetCurrentDirectory(), "Website", "data", "office_conversion_routes.json");
bool verify = args.Contains("--verify", StringComparer.OrdinalIgnoreCase);
string converterSamplePath = GetOption(args, "--converter-sample")
    ?? Path.Combine(
        Directory.GetCurrentDirectory(),
        "Website", "Apps", "OfficeIMO.Web.Converter", "wwwroot", "samples", "conversion-proof.pptx");

var capabilityCatalogs = new (string Name, OfficeCapabilityCatalog Catalog)[] {
    ("word-legacy-doc", WordCompatibilityCatalog.Current),
    ("excel-legacy-xls", ExcelCompatibilityCatalog.Xls),
    ("excel-xlsb", ExcelCompatibilityCatalog.Xlsb),
    ("powerpoint-legacy-ppt", PowerPointCompatibilityCatalog.Current)
};
var outputs = new SortedDictionary<string, string>(StringComparer.Ordinal) {
    ["conversion-routes.json"] = EnsureFinalNewline(OfficeConversionCapabilityCatalog.ToJson()),
    ["conversion-routes.md"] = EnsureFinalNewline(OfficeConversionCapabilityCatalog.ToMarkdown()),
    ["office-formats.json"] = SerializeFormats(),
    ["protected-content.json"] = EnsureFinalNewline(OfficeProtectionCapabilityCatalog.Current.ToJson()),
    ["protected-content.md"] = EnsureFinalNewline(OfficeProtectionCapabilityCatalog.Current.ToMarkdown()),
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
    string expectedWebsiteData = outputs["conversion-routes.json"];
    if (!File.Exists(websiteDataPath) || Normalize(File.ReadAllText(websiteDataPath)) != Normalize(expectedWebsiteData)) {
        Console.Error.WriteLine("Compatibility catalog website data is missing or stale: " + websiteDataPath);
        Environment.ExitCode = 1;
        return;
    }
    if (!VerifyConverterPowerPointSample(converterSamplePath, out string sampleError)) {
        Console.Error.WriteLine("Converter proof sample is missing or stale: " + sampleError);
        Environment.ExitCode = 1;
        return;
    }
    Console.WriteLine($"Verified {outputs.Count} compatibility catalog artifacts, website route data, and the converter proof sample.");
    return;
}

Directory.CreateDirectory(outputDirectory);
foreach ((string fileName, string content) in outputs) {
    File.WriteAllText(Path.Combine(outputDirectory, fileName), Normalize(content), new UTF8Encoding(encoderShouldEmitUTF8Identifier: false));
}
string? websiteDataDirectory = Path.GetDirectoryName(websiteDataPath);
if (!string.IsNullOrEmpty(websiteDataDirectory)) Directory.CreateDirectory(websiteDataDirectory);
File.WriteAllText(websiteDataPath, Normalize(outputs["conversion-routes.json"]), new UTF8Encoding(encoderShouldEmitUTF8Identifier: false));
GenerateConverterPowerPointSample(converterSamplePath);
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
    markdown.AppendLine("dotnet run --framework net8.0 --project Build/CompatibilityCatalog/OfficeIMO.CompatibilityCatalog.Tool.csproj -- --output Docs/Compatibility/generated --converter-sample Website/Apps/OfficeIMO.Web.Converter/wwwroot/samples/conversion-proof.pptx");
    markdown.AppendLine("```");
    markdown.AppendLine();
    markdown.AppendLine("Use the [conversion route catalog](conversion-routes.md) to find the focused package, output model, proven support level, known limits, browser availability, and result type for each route.");
    markdown.AppendLine();
    markdown.AppendLine("Verify:");
    markdown.AppendLine();
    markdown.AppendLine("```powershell");
    markdown.AppendLine("dotnet run --framework net8.0 --project Build/CompatibilityCatalog/OfficeIMO.CompatibilityCatalog.Tool.csproj -- --output Docs/Compatibility/generated --converter-sample Website/Apps/OfficeIMO.Web.Converter/wwwroot/samples/conversion-proof.pptx --verify");
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
    OfficeProtectionCapabilityCatalog protection = OfficeProtectionCapabilityCatalog.Current;
    markdown.Append("| ").Append(protection.Id)
        .Append(" | ").Append(protection.SchemaVersion)
        .Append(" | ").Append(protection.Capabilities.Count)
        .Append(" | [JSON](protected-content.json)")
        .AppendLine(" | [Markdown](protected-content.md) |");
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

static void GenerateConverterPowerPointSample(string outputPath) {
    string? directory = Path.GetDirectoryName(Path.GetFullPath(outputPath));
    if (!string.IsNullOrEmpty(directory)) Directory.CreateDirectory(directory);

    using PowerPointPresentation presentation = PowerPointPresentation.Create(outputPath);
    presentation.SlideSize.SetPreset(PowerPointSlideSizePreset.Screen16x9);

    PowerPointSlide overview = presentation.AddSlide(PowerPointSlideLayoutType.Blank);
    AddText(overview, "OfficeIMO conversion proof", 42, 28, 650, 46, 28, bold: true, color: "17324D");
    AddText(overview, "Editable PowerPoint content rendered locally to PDF", 42, 76, 650, 26, 14, color: "52657A");
    AddMetric(
        overview,
        OfficeConversionCapabilityCatalog.BrowserRoutes.Count.ToString(System.Globalization.CultureInfo.InvariantCulture),
        "browser routes",
        42,
        126,
        "E8F0FF",
        "2563EB");
    AddMetric(overview, "100%", "local processing", 244, 126, "E9F8F0", "07875B");
    AddMetric(overview, "0", "server uploads", 446, 126, "FFF2E8", "C2410C");

    PowerPointTable table = overview.AddTablePoints(4, 3, 42, 252, 646, 190);
    table.GetCell(0, 0).Text = "Workstream";
    table.GetCell(0, 1).Text = "Owner";
    table.GetCell(0, 2).Text = "Status";
    table.GetCell(1, 0).Text = "PDF rendering";
    table.GetCell(1, 1).Text = "Drawing core";
    table.GetCell(1, 2).Text = "Advanced";
    table.GetCell(2, 0).Text = "Editable imports";
    table.GetCell(2, 1).Text = "Format adapters";
    table.GetCell(2, 2).Text = "Targeted";
    table.GetCell(3, 0).Text = "Browser workbench";
    table.GetCell(3, 1).Text = "WebAssembly";
    table.GetCell(3, 2).Text = "Local";

    PowerPointSlide chartSlide = presentation.AddSlide(PowerPointSlideLayoutType.Blank);
    AddText(chartSlide, "Quarterly conversion volume", 42, 28, 650, 46, 28, bold: true, color: "17324D");
    AddText(chartSlide, "Native chart data, title, axes, legend, and series", 42, 76, 650, 26, 14, color: "52657A");
    var chartData = new OfficeChartData(
        new[] { "Q1", "Q2", "Q3", "Q4" },
        new[] {
            new OfficeChartSeries("Documents", new[] { 120D, 185D, 240D, 318D }),
            new OfficeChartSeries("Reports", new[] { 42D, 68D, 96D, 132D })
        });
    PowerPointChart chart = chartSlide.AddChartPoints(
        OfficeChartKind.ColumnClustered,
        chartData,
        66,
        128,
        600,
        310,
        new PowerPointChartAccessibilityOptions {
            Name = "Quarterly conversion volume",
            AlternativeText = "Document and report conversion volume increases from Q1 through Q4.",
            IncludeDataSummaryInAlternativeText = true
        });
    chart.SetTitle("Completed conversions");
    chart.SetLegend(OfficeChartLegendPosition.Bottom);

    presentation.Save();
}

static bool VerifyConverterPowerPointSample(string inputPath, out string error) {
    if (!File.Exists(inputPath)) {
        error = Path.GetFullPath(inputPath);
        return false;
    }

    try {
        using var stream = File.OpenRead(inputPath);
        using PowerPointPresentation presentation = PowerPointPresentation.Load(stream);
        if (presentation.Slides.Count != 2) {
            error = $"expected 2 slides, found {presentation.Slides.Count}";
            return false;
        }
        if (!presentation.Slides.Any(static slide => slide.Shapes.OfType<PowerPointTable>().Any())) {
            error = "the table proof is absent";
            return false;
        }
        if (!presentation.Slides.Any(static slide => slide.Shapes.OfType<PowerPointChart>().Any())) {
            error = "the native chart proof is absent";
            return false;
        }

        string[] text = presentation.Slides
            .SelectMany(static slide => slide.TextBoxes)
            .Select(static textBox => textBox.Text)
            .ToArray();
        string expectedRouteCount = OfficeConversionCapabilityCatalog.BrowserRoutes.Count
            .ToString(System.Globalization.CultureInfo.InvariantCulture);
        string[] requiredText = [
            "OfficeIMO conversion proof",
            expectedRouteCount,
            "browser routes",
            "Quarterly conversion volume"
        ];
        string? missingText = requiredText.FirstOrDefault(required =>
            !text.Contains(required, StringComparer.Ordinal));
        if (missingText is not null) {
            error = $"required text '{missingText}' is absent";
            return false;
        }

        error = string.Empty;
        return true;
    } catch (Exception exception) {
        error = $"{Path.GetFullPath(inputPath)} could not be opened: {exception.Message}";
        return false;
    }
}

static void AddMetric(
    PowerPointSlide slide,
    string value,
    string label,
    double left,
    double top,
    string fill,
    string accent) {
    PowerPointAutoShape card = slide.AddRectanglePoints(left, top, 180, 92, label);
    card.FillColor = fill;
    card.OutlineColor = accent;
    card.OutlineWidthPoints = 1.25;
    AddText(slide, value, left + 14, top + 12, 150, 34, 24, bold: true, color: accent);
    AddText(slide, label, left + 14, top + 52, 150, 24, 12, color: "52657A");
}

static PowerPointTextBox AddText(
    PowerPointSlide slide,
    string text,
    double left,
    double top,
    double width,
    double height,
    int fontSize,
    bool bold = false,
    string color = "17324D") {
    PowerPointTextBox box = slide.AddTextBoxPoints(text, left, top, width, height);
    box.FontName = "Carlito";
    box.FontSize = fontSize;
    box.Bold = bold;
    box.Color = color;
    box.SetTextMarginsPoints(0, 0, 0, 0);
    box.TextAutoFit = PowerPointTextAutoFit.Normal;
    return box;
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
