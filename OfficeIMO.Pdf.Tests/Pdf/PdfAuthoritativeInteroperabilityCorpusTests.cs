using System.Text.Json;
using OfficeIMO.Pdf;
using Xunit;

namespace OfficeIMO.Tests.Pdf;

public sealed class PdfAuthoritativeInteroperabilityCorpusTests {
    [Fact]
    public void AuthoritativeCorpus_IsPinnedCompleteAndHashVerified() {
        using JsonDocument manifest = LoadManifest();
        JsonElement root = manifest.RootElement;
        Assert.Equal(1, root.GetProperty("version").GetInt32());
        Assert.Equal("https://github.com/pdf-association/pdf-corpora", root.GetProperty("authority").GetString());

        JsonElement[] sources = root.GetProperty("sources").EnumerateArray().ToArray();
        Assert.Equal(2, sources.Length);
        Assert.All(sources, source => {
            Assert.StartsWith("https://github.com/", RequireString(source, "repository"), StringComparison.Ordinal);
            Assert.Matches("^[0-9a-f]{40}$", RequireString(source, "commit"));
            Assert.False(string.IsNullOrWhiteSpace(RequireString(source, "license")));
        });
        var sourceIds = sources.Select(source => RequireString(source, "id")).ToHashSet(StringComparer.Ordinal);

        JsonElement[] cases = root.GetProperty("cases").EnumerateArray().ToArray();
        Assert.Equal(9, cases.Length);
        Assert.Equal(cases.Length, cases.Select(item => RequireString(item, "id")).Distinct(StringComparer.Ordinal).Count());
        Assert.Contains(cases, item => RequireString(item, "source") == "openpreserve-format-corpus");
        Assert.Contains(cases, item => RequireString(item, "source") == "verapdf-corpus");

        var manifestedFiles = new HashSet<string>(StringComparer.Ordinal);
        foreach (JsonElement item in cases) {
            string fileName = RequireString(item, "file");
            Assert.True(manifestedFiles.Add(fileName), "Duplicate corpus file: " + fileName);
            Assert.Contains(RequireString(item, "source"), sourceIds);
            Assert.EndsWith(".pdf", RequireString(item, "sourcePath"), StringComparison.OrdinalIgnoreCase);
            Assert.NotEmpty(item.GetProperty("features").EnumerateArray());

            string path = Path.Combine(FixtureRoot, fileName);
            Assert.True(File.Exists(path), "Missing corpus fixture: " + fileName);
            byte[] bytes = File.ReadAllBytes(path);
            Assert.Equal(item.GetProperty("byteLength").GetInt64(), bytes.LongLength);
            Assert.Equal(
                RequireString(item, "sha256"),
                PdfArtifactFingerprint.ComputeSha256(bytes));
        }

        string[] actualFiles = Directory.GetFiles(FixtureRoot, "*.pdf")
            .Select(Path.GetFileName)
            .Where(fileName => fileName != null)
            .Cast<string>()
            .OrderBy(fileName => fileName, StringComparer.Ordinal)
            .ToArray();
        Assert.Equal(manifestedFiles.OrderBy(fileName => fileName, StringComparer.Ordinal), actualFiles);
    }

    [Fact]
    public void AuthoritativeCorpus_OpensExtractsInspectsRendersAndPlansMutation() {
        using JsonDocument manifest = LoadManifest();
        foreach (JsonElement item in manifest.RootElement.GetProperty("cases").EnumerateArray()) {
            string id = RequireString(item, "id");
            byte[] bytes = File.ReadAllBytes(Path.Combine(FixtureRoot, RequireString(item, "file")));
            PdfReadDocument document = PdfReadDocument.Open(bytes);
            PdfDocumentInfo info = PdfInspector.Inspect(bytes);
            string text = document.ExtractText();
            PdfPageRenderResult render = Assert.Single(PdfPageImageRenderer.RenderPages(bytes, options: new PdfPageRenderOptions {
                Format = PdfPageRenderFormat.Svg,
                ContinueOnError = true,
                MaxPages = 4
            }));
            PdfMutationPlan plan = PdfMutationPlanner.Plan(bytes, PdfMutationOperation.UpdateMetadata);

            Assert.Equal(item.GetProperty("pageCount").GetInt32(), document.Pages.Count);
            Assert.True(
                text.Length >= item.GetProperty("minimumTextCharacters").GetInt32(),
                id + " extracted too little text.");
            Assert.True(
                info.AttachmentCount >= item.GetProperty("minimumAttachments").GetInt32(),
                id + " lost expected attachments.");
            Assert.True(
                info.LinkAnnotationCount >= item.GetProperty("minimumLinks").GetInt32(),
                id + " lost expected links.");
            Assert.True(
                info.AnnotationCount >= item.GetProperty("minimumAnnotations").GetInt32(),
                id + " lost expected annotations.");
            Assert.Equal(
                ReadStringArray(item, "expectedAnnotationActionTypes"),
                info.AnnotationActionTypes);
            Assert.Equal(
                ReadStringArray(item, "expectedRepairCodes"),
                document.RepairReport.Diagnostics.Select(diagnostic => diagnostic.Code).ToArray());
            Assert.Equal(item.GetProperty("expectedRenderSucceeded").GetBoolean(), render.Succeeded);
            Assert.Equal(
                ReadStringArray(item, "expectedRenderDiagnosticCodes"),
                render.CapabilityDiagnostics.Select(diagnostic => diagnostic.Code).Distinct(StringComparer.Ordinal).ToArray());
            Assert.Equal(
                (PdfMutationExecutionMode)Enum.Parse(
                    typeof(PdfMutationExecutionMode),
                    RequireString(item, "expectedMutationMode")),
                plan.ExecutionMode);
        }
    }

    private static JsonDocument LoadManifest() =>
        JsonDocument.Parse(File.ReadAllBytes(Path.Combine(FixtureRoot, "corpus-manifest.json")));

    private static string RequireString(JsonElement element, string propertyName) {
        string? value = element.GetProperty(propertyName).GetString();
        Assert.False(string.IsNullOrWhiteSpace(value));
        return value!;
    }

    private static string[] ReadStringArray(JsonElement element, string propertyName) =>
        element.GetProperty(propertyName)
            .EnumerateArray()
            .Select(value => value.GetString() ?? string.Empty)
            .ToArray();

    private static string FixtureRoot => Path.Combine(AppContext.BaseDirectory, "Pdf", "Fixtures", "Interoperability");
}
