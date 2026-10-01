using System.Text;
using System.Text.Json;
using OfficeIMO.Html;
using OfficeIMO.Tests;

namespace OfficeIMO.Pdf.Benchmarks.Comparisons;

internal static partial class HtmlCorpusEvidenceRunner {
    private const string StaticGapRoot = "OfficeIMO.TestAssets/Documents/Html/Qualification/StaticPdfGaps";

    private static HtmlRenderOptions CreateCorpusRenderOptions(HtmlCorpusEvidenceInput input) {
        HtmlRenderOptions options = input.Scenario.CreateOptions();
        // These immutable offline inputs embed their fonts and images. Keep the
        // resource boundary explicit for screen as well as PDF conversion.
        if (input.IsStaticGap) options.ResourceUrlPolicy = HtmlUrlPolicy.CreateEmbeddedResourceProfile();
        return options;
    }

    /// <summary>Loads frozen equivalent inputs; criteria remain independent of renderer output.</summary>
    private static HtmlCorpusEvidenceInputSet LoadStaticGapCorpus() {
        string root = Path.Combine(FindRepositoryRoot(), StaticGapRoot);
        byte[] manifestBytes = File.ReadAllBytes(Path.Combine(root, "manifest.json"));
        using JsonDocument manifest = JsonDocument.Parse(manifestBytes);
        JsonElement definition = manifest.RootElement;
        if (definition.GetProperty("schema").GetString() != "officeimo.html.static-gap-corpus"
            || definition.GetProperty("version").GetInt32() != 1) {
            throw new InvalidDataException("Unknown static PDF gap corpus contract.");
        }
        var inputs = new List<HtmlCorpusEvidenceInput>();
        var ids = new HashSet<string>(StringComparer.Ordinal);
        foreach (JsonElement item in definition.GetProperty("cases").EnumerateArray()) {
            string id = item.GetProperty("id").GetString() ?? throw new InvalidDataException("Missing case ID.");
            string file = item.GetProperty("path").GetString() ?? throw new InvalidDataException("Missing input path.");
            if (!ids.Add(id) || file != Path.GetFileName(file) || file != id + ".html") {
                throw new InvalidDataException("Duplicate or invalid static PDF case path.");
            }
            byte[] bytes = File.ReadAllBytes(Path.Combine(root, file));
            if (bytes.Length > 1_048_576 || bytes.Length != item.GetProperty("length").GetInt32()
                || Sha256(bytes) != item.GetProperty("sha256").GetString()) {
                throw new InvalidDataException("Frozen static PDF input differs from its manifest: " + file);
            }
            string[] markers = item.GetProperty("textMarkers").EnumerateArray()
                .Select(marker => marker.GetString() ?? throw new InvalidDataException("Null content marker.")).ToArray();
            string group = item.GetProperty("group").GetString() ?? throw new InvalidDataException("Missing capability group.");
            inputs.Add(new HtmlCorpusEvidenceInput(
                new HtmlRenderingCorpusCase(id, HtmlRenderMode.Paged, Encoding.UTF8.GetString(bytes), markers,
                    minimumVisualCount: 1, minimumHeadingCount: 1),
                StaticGapRoot + "/" + file, new[] { group }, bytes, IsStaticGap: true));
        }
        if (inputs.Count == 0) throw new InvalidDataException("Static PDF gap corpus is empty.");
        return new HtmlCorpusEvidenceInputSet("officeimo-static-html-pdf-gaps", StaticGapRoot,
            Sha256(manifestBytes), null, inputs);
    }
}
