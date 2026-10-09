using System.Text.Json;
using System.Xml.Linq;
using System.Runtime.CompilerServices;
using OfficeIMO.Drawing;
using OfficeIMO.Pdf;
using OfficeIMO.Web.Converter.Services;
using OfficeIMO.Workflows;
using Xunit;

namespace OfficeIMO.Web.Converter.Tests;

/// <summary>
/// Website/data/browser_tools.json drives the tool pages, the /convert/ directory and the engine wiring.
/// These tests keep it consistent with the engine catalogs, the lazy-load list, and the generated pages.
/// </summary>
public sealed class BrowserToolCatalogTests {
    private static readonly string WebsiteRoot = FindWebsiteRoot();
    private static readonly JsonElement Catalog = JsonDocument.Parse(File.ReadAllText(Path.Combine(WebsiteRoot, "data", "browser_tools.json"))).RootElement;
    private static IEnumerable<JsonElement> Tools => Catalog.GetProperty("tools").EnumerateArray();

    [Fact]
    public void ToolIdsAreUniqueUrlSlugsAndBelongToAGroup() {
        string[] ids = Tools.Select(static tool => tool.GetProperty("id").GetString()!).ToArray();
        Assert.Equal(ids.Length, ids.Distinct(StringComparer.Ordinal).Count());
        Assert.All(ids, id => Assert.Matches("^[a-z0-9]+(-[a-z0-9]+)*$", id));
        string[] groups = Catalog.GetProperty("groups").EnumerateArray().Select(static group => group.GetProperty("id").GetString()!).ToArray();
        Assert.All(Tools, tool => Assert.Contains(tool.GetProperty("group").GetString(), groups));
    }

    [Fact]
    public void EveryBrowserConversionRouteHasExactlyOneTool() {
        string[] targets = Tools.Where(static tool => Engine(tool).GetProperty("kind").GetString() == "convert")
            .Select(static tool => Engine(tool).GetProperty("target").GetString()!)
            .ToArray();
        Assert.Equal(
            ConversionRouteCatalog.All.Select(static route => route.Id).Order(StringComparer.Ordinal),
            targets.Order(StringComparer.Ordinal));
    }

    [Fact]
    public void ConversionToolsAcceptTheRouteExtensions() {
        foreach (JsonElement tool in Tools.Where(static tool => Engine(tool).GetProperty("kind").GetString() == "convert")) {
            var route = ConversionRouteCatalog.All.Single(route => route.Id == Engine(tool).GetProperty("target").GetString());
            Assert.Equal(
                route.Accept.Split(',').Order(StringComparer.Ordinal),
                Accept(tool).Order(StringComparer.Ordinal));
            Assert.Equal(route.InputKind == Models.ConversionInputKind.File ? "file" : "text", tool.GetProperty("input").GetProperty("kind").GetString());
        }
    }

    [Fact]
    public void EveryPdfWorkbenchOperationHasExactlyOneTool() {
        string[] targets = Tools.Where(static tool => Engine(tool).GetProperty("kind").GetString() == "pdf")
            .Select(static tool => Engine(tool).GetProperty("target").GetString()!)
            .ToArray();
        Assert.Equal(
            PdfToolCatalog.All.Select(static tool => tool.Id).Order(StringComparer.Ordinal),
            targets.Order(StringComparer.Ordinal));
    }

    [Fact]
    public void ToolsThatChangePagesOrContentCarryTheEngineConfirmation() {
        foreach (var definition in PdfToolCatalog.All.Where(static tool => tool.RequiresDestructiveConfirmation)) {
            JsonElement tool = Tools.Single(tool => Engine(tool).GetProperty("kind").GetString() == "pdf" && Engine(tool).GetProperty("target").GetString() == definition.Id);
            Assert.True(Engine(tool).TryGetProperty("confirm", out JsonElement confirm) && confirm.GetBoolean(), $"{definition.Id} must send the confirmation flag.");
        }
    }

    [Fact]
    public void FileOriginToolAcceptsExactlyTheQualifiedFormats() {
        JsonElement tool = Tools.Single(static tool => Engine(tool).GetProperty("kind").GetString() == "origin");
        Assert.Equal(OfficeProvenanceWorkflowCatalog.BrowserExtensions.Order(StringComparer.Ordinal), Accept(tool).Order(StringComparer.Ordinal));
    }

    [Fact]
    public void OptionValuesExistInTheEngineCatalogs() {
        foreach (JsonElement tool in Tools) {
            foreach (JsonElement option in tool.GetProperty("options").EnumerateArray()) {
                if (!option.TryGetProperty("choices", out JsonElement choices)) continue;
                string id = option.GetProperty("id").GetString()!;
                string[] values = choices.EnumerateArray().Select(static choice => choice.GetProperty("value").GetString()!).ToArray();
                Assert.Contains(option.GetProperty("default").GetString(), values);
                string kind = Engine(tool).GetProperty("kind").GetString()!;
                foreach (string value in values) {
                    switch (kind, id) {
                        case ("convert", "profile"):
                            Assert.Contains(value, BrowserPdfProfileCatalog.All.Select(static profile => profile.Id));
                            break;
                        case ("convert", "slides"):
                            Assert.Contains(value, BrowserPowerPointImportProfileCatalog.All.Select(static profile => profile.Id));
                            break;
                        case ("convert", "rows"):
                            Assert.Contains(value, new[] { "full", "preview" });
                            break;
                        case ("pdf", "profile"):
                            Assert.True(Enum.TryParse(value, out PdfOptimizationProfile _), $"{value} is not a PdfOptimizationProfile.");
                            break;
                        case ("pdf", "rotation"):
                            Assert.Contains(value, new[] { "90", "180", "270" });
                            break;
                        default:
                            Assert.Fail($"{tool.GetProperty("id").GetString()} has an unchecked choice option '{id}'. Add it to this test.");
                            break;
                    }
                }
            }
        }
    }

    [Fact]
    public void EngineAssembliesAreAllMarkedForLazyLoading() {
        string project = Path.Combine(WebsiteRoot, "Apps", "OfficeIMO.Web.Converter", "OfficeIMO.Web.Converter.csproj");
        string[] lazy = XDocument.Load(project).Descendants("BlazorWebAssemblyLazyLoad")
            .Select(static element => element.Attribute("Include")!.Value)
            .ToArray();
        foreach (JsonElement tool in Tools) {
            foreach (JsonElement assembly in Engine(tool).GetProperty("assemblies").EnumerateArray()) {
                Assert.Contains(assembly.GetString(), lazy);
            }
        }
        // Every lazily loaded engine is used by at least one tool; otherwise it is dead weight in the publish output.
        string[] used = Tools.SelectMany(static tool => Engine(tool).GetProperty("assemblies").EnumerateArray().Select(static assembly => assembly.GetString()!)).Distinct().ToArray();
        Assert.All(lazy, assembly => Assert.Contains(assembly, used));
    }

    [Fact]
    public void NextStepsAndLegacyLinksPointAtRealTools() {
        string[] ids = Tools.Select(static tool => tool.GetProperty("id").GetString()!).ToArray();
        Assert.All(Tools, tool => Assert.All(tool.GetProperty("next").EnumerateArray(), next => Assert.Contains(next.GetString(), ids)));
        string[] legacy = Tools.SelectMany(static tool => tool.GetProperty("legacy").EnumerateArray().Select(static key => key.GetString()!)).ToArray();
        Assert.Equal(legacy.Length, legacy.Distinct(StringComparer.Ordinal).Count());
        foreach (var route in ConversionRouteCatalog.All) Assert.Contains("route=" + route.Id, legacy);
        foreach (var tool in PdfToolCatalog.All) Assert.Contains("workspace=pdf&tool=" + tool.Id, legacy);
        Assert.Contains("workspace=provenance", legacy);
    }

    [Fact]
    public void EveryToolHasAGeneratedPage() {
        string pages = Path.Combine(WebsiteRoot, "content", "browser");
        foreach (JsonElement tool in Tools) {
            string id = tool.GetProperty("id").GetString()!;
            string page = Path.Combine(pages, id + ".md");
            Assert.True(File.Exists(page), $"Missing content/browser/{id}.md. Run Website/scripts/Sync-BrowserToolPages.ps1.");
            string text = File.ReadAllText(page);
            Assert.Contains($"meta.tool: \"{id}\"", text, StringComparison.Ordinal);
            Assert.Contains($"title: \"{tool.GetProperty("title").GetString()}\"", text, StringComparison.Ordinal);
        }
        Assert.Equal(Tools.Count(), Directory.GetFiles(pages, "*.md").Length);
    }

    [Fact]
    public void SamplesExistInTheEngineFolder() {
        string wwwroot = Path.Combine(WebsiteRoot, "Apps", "OfficeIMO.Web.Converter", "wwwroot");
        foreach (JsonElement tool in Tools) {
            foreach (string key in new[] { "sample", "sampleSecond" }) {
                if (!tool.GetProperty("input").TryGetProperty(key, out JsonElement sample)) continue;
                string path = sample.GetString()!;
                bool linkedByProject = path == "samples/showcase-dashboard.pdf"; // linked from Website/static by the project file
                Assert.True(linkedByProject || File.Exists(Path.Combine(wwwroot, path)), $"{tool.GetProperty("id").GetString()} {key} {path} is missing.");
            }
        }
    }

    private static JsonElement Engine(JsonElement tool) => tool.GetProperty("engine");

    private static string[] Accept(JsonElement tool) =>
        tool.GetProperty("input").GetProperty("accept").GetString()!.Split(',', StringSplitOptions.RemoveEmptyEntries | StringSplitOptions.TrimEntries);

    private static string FindWebsiteRoot([CallerFilePath] string sourcePath = "") {
        string? directory = Path.GetDirectoryName(sourcePath);
        while (directory is not null) {
            if (File.Exists(Path.Combine(directory, "data", "browser_tools.json"))) return directory;
            string nested = Path.Combine(directory, "Website");
            if (File.Exists(Path.Combine(nested, "data", "browser_tools.json"))) return nested;
            directory = Path.GetDirectoryName(directory);
        }
        throw new InvalidOperationException("Website/data/browser_tools.json was not found above the test output folder.");
    }
}
