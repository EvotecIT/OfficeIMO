using System.Security.Cryptography;
using System.Text.Json;
using System.Text.RegularExpressions;
using OfficeIMO.Pdf;

namespace OfficeIMO.IWork.Tests;

public sealed class IWorkPagesNineLevelNativeExportTests {
    [Fact]
    public void Creation_and_preserved_output_equal_the_independently_decoded_Apple_accepted_inputs() {
        using var manifest = Manifest();
        var paths = new HashSet<string>(StringComparer.Ordinal);
        foreach (JsonProperty artifact in manifest.RootElement.GetProperty("artifacts").EnumerateObject()) {
            string path = artifact.Value.GetProperty("path").GetString()!; Assert.True(paths.Add(path));
            byte[] bytes = File.ReadAllBytes(Corpus(path));
            Assert.Equal(artifact.Value.GetProperty("bytes").GetInt64(), bytes.LongLength);
            Assert.Equal(artifact.Value.GetProperty("sha256").GetString(), Convert.ToHexString(SHA256.HashData(bytes)).ToLowerInvariant());
        }
        foreach (JsonElement scenario in manifest.RootElement.GetProperty("cases").EnumerateArray())
            foreach (string field in new[] { "source", "nativeDocx", "nativePdf", "officeimoDocx", "officeimoPdf" })
                Assert.Contains(scenario.GetProperty(field).GetString()!, paths);
        Assert.Equal(File.ReadAllBytes(Corpus("pages-nine-level-created.pages")), IWorkPagesNineLevelWritingTests.Example().SaveBytes());
        var original = IWorkSourceDocument.Open(Corpus("pages-nine-level-saved-v15.4.pages"));
        var body = original.ReadPages().Body; var edit = original.BeginEdit();
        foreach (string text in IWorkPagesNineLevelWritingTests.LevelNames) {
            int index = body.Paragraphs.ToList().FindIndex(p => p.Text == text);
            edit.SetPagesBodyParagraphText(body.SourceIdentity!, index + 1, text.ToUpperInvariant());
        }
        Assert.Equal(File.ReadAllBytes(Corpus("pages-nine-level-preserved.pages")),
            edit.SaveBytes(new() { AllowUnassessedVisualCaches = true }));
    }

    [Theory]
    [InlineData(0)] [InlineData(1)] [InlineData(2)] [InlineData(3)]
    public void Native_creation_save_preservation_and_edit_retain_all_nine_levels_in_strict_saved_Word(int caseIndex) {
        using var manifest = Manifest(); var scenario = manifest.RootElement.GetProperty("cases")[caseIndex];
        var source = IWorkSourceDocument.Open(Corpus(scenario.GetProperty("source").GetString()!));
        var projection = source.ReadPages();
        Assert.True(projection.Body.IsComplete); Assert.True(projection.Body.IsFormattingComplete);
        Assert.Empty(projection.Diagnostics); Assert.Empty(projection.SourceDeclarationIssues); Assert.Empty(projection.SourceReferenceIssues);
        var expected = scenario.GetProperty("expectedItems").EnumerateArray().ToArray();
        var actual = projection.Body.Paragraphs.Where(p => p.ListNumber.HasValue).ToArray();
        Assert.Equal(expected.Select(p => p.GetProperty("level").GetInt32()), actual.Select(p => p.ListLevel));
        Assert.Equal(expected.Select(p => p.GetProperty("number").GetInt32()), actual.Select(p => p.ListNumber!.Value));
        Assert.Equal(expected.Select(p => p.GetProperty("text").GetString()), actual.Select(p => p.Text));
        IWorkPagesDecimalTextEditingTests.CheckWord(source);
        if (caseIndex == 3) {
            int last = projection.Body.Paragraphs.ToList().FindIndex(p => p.Text == "Native deep two");
            Assert.Equal(8, projection.Body.Paragraphs[last].ListLevel);
            var edit = source.BeginEdit();
            edit.SetPagesBodyParagraphText(projection.Body.SourceIdentity!, last + 1, "Native deep six");
            var saved = IWorkSourceDocument.Open(edit.SaveBytes(new() { AllowUnassessedVisualCaches = true }));
            Assert.Equal("Native deep six", saved.ReadPages().Body.Paragraphs[last].Text);
            Assert.Equal(2, saved.ReadPages().Body.Paragraphs[last].ListNumber);
            IWorkPagesDecimalTextEditingTests.CheckWord(saved);
            edit.SetPagesBodyParagraphText(projection.Body.SourceIdentity!, last + 1, "Native deep two");
            Assert.Empty(edit.SaveBytesResult().Report.UpdatedRecords);
            Assert.Equal(source.ComputePackageContentHash(), edit.SaveBytesResult().Report.OutputPackageContentHash);
        }
    }

    [Fact]
    public void Every_native_composed_and_receiving_PDF_retains_the_complete_nine_level_document() {
        using var manifest = Manifest();
        foreach (JsonElement scenario in manifest.RootElement.GetProperty("cases").EnumerateArray()) {
            Check(scenario.GetProperty("nativePdf").GetString()!, scenario);
            Check(scenario.GetProperty("officeimoPdf").GetString()!, scenario);
        }
        Check("pages-nine-level-officeimo-docx-v15.4.pdf", manifest.RootElement.GetProperty("cases")[2]);

        static void Check(string path, JsonElement scenario) {
            var pdf = PdfReadDocument.Open(File.ReadAllBytes(Corpus(path))); Assert.Single(pdf.Pages);
            string text = Regex.Replace(pdf.ExtractText(), @"\s+", string.Empty); int previous = -1;
            foreach (JsonElement item in scenario.GetProperty("expectedItems").EnumerateArray()) {
                string wanted = Regex.Replace(item.GetProperty("number").GetInt32().ToString(System.Globalization.CultureInfo.InvariantCulture)
                    + "." + item.GetProperty("text").GetString(), @"\s+", string.Empty);
                int at = text.IndexOf(wanted, previous + 1, StringComparison.Ordinal);
                Assert.True(at > previous, path + ": " + wanted); previous = at;
            }
            foreach (JsonElement plain in scenario.GetProperty("expectedPlainParagraphs").EnumerateArray())
                Assert.Contains(Regex.Replace(plain.GetString()!, @"\s+", string.Empty), text);
        }
    }

    private static JsonDocument Manifest() => JsonDocument.Parse(File.ReadAllBytes(Corpus("pages-nine-level-v15.4.json")));
    private static string Corpus(string name) => Path.Combine(AppContext.BaseDirectory, "Documents", "IWorkCorpus", "native-exports", name);
}
