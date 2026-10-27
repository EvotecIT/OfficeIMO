using System.Security.Cryptography;
using System.Text.Json;
using System.Text.RegularExpressions;
using OfficeIMO.Pdf;

namespace OfficeIMO.IWork.Tests;

public sealed class IWorkPagesDecimalTextNativeExportTests {
    [Fact]
    public void Preserving_edit_matches_the_independently_compared_Apple_accepted_bytes() {
        var source = IWorkSourceDocument.Open(Corpus("pages-third-level-saved-v15.4.pages"));
        var body = source.ReadPages().Body; var edit = source.BeginEdit();
        edit.SetPagesBodyParagraphText(body.SourceIdentity!, 2, "Edited seven");
        edit.SetPagesBodyParagraphText(body.SourceIdentity!, 3, "Child two");
        edit.SetPagesBodyParagraphText(body.SourceIdentity!, 4, "Grandchild six");
        edit.SetPagesBodyParagraphText(body.SourceIdentity!, 12, "Updated grandchild six");
        Assert.Equal(File.ReadAllBytes(Corpus("pages-preserved-decimal-v15.4.pages")),
            edit.SaveBytes(new() { AllowUnassessedVisualCaches = true }));
        using var manifest = Manifest();
        foreach (JsonProperty artifact in manifest.RootElement.GetProperty("artifacts").EnumerateObject()) {
            byte[] bytes = File.ReadAllBytes(Corpus(artifact.Value.GetProperty("path").GetString()!));
            Assert.Equal(artifact.Value.GetProperty("bytes").GetInt64(), bytes.LongLength);
            Assert.Equal(artifact.Value.GetProperty("sha256").GetString(), Convert.ToHexString(SHA256.HashData(bytes)).ToLowerInvariant());
        }
    }

    [Theory]
    [InlineData(0)]
    [InlineData(1)]
    public void Preserved_and_Apple_resaved_items_retain_counters_and_allow_a_further_fixed_length_edit(int caseIndex) {
        using var manifest = Manifest(); var scenario = manifest.RootElement.GetProperty("cases")[caseIndex];
        var source = IWorkSourceDocument.Open(Corpus(scenario.GetProperty("source").GetString()!));
        var projection = source.ReadPages();
        Assert.True(projection.HasEditableContent); Assert.True(projection.Body.IsFormattingComplete);
        Assert.Empty(projection.Diagnostics); Assert.Empty(projection.SourceDeclarationIssues); Assert.Empty(projection.SourceReferenceIssues);
        Assert.Equal(scenario.GetProperty("expectedParagraphs").EnumerateArray().Select(p => p.GetString()), projection.Paragraphs);
        var expected = scenario.GetProperty("expectedItems").EnumerateArray().ToArray();
        var items = projection.Body.Paragraphs.Where(p => p.ListNumber.HasValue).ToArray();
        Assert.Equal(expected.Select(p => p.GetProperty("number").GetInt32()), items.Select(p => p.ListNumber!.Value));
        Assert.Equal(expected.Select(p => p.GetProperty("level").GetInt32()), items.Select(p => p.ListLevel));
        IWorkPagesDecimalTextEditingTests.CheckWord(source);
        var edit = source.BeginEdit(); var id = projection.Body.SourceIdentity!;
        edit.SetPagesBodyParagraphText(id, 2, "Further edit");
        Assert.Equal("Further edit", IWorkSourceDocument.Open(edit.SaveBytes(new() { AllowUnassessedVisualCaches = true })).ReadPages().Paragraphs[1]);
        edit.SetPagesBodyParagraphText(id, 2, projection.Paragraphs[1]);
        Assert.Equal(source.ComputePackageContentHash(), edit.SaveBytesResult().Report.OutputPackageContentHash);
    }

    [Fact]
    public void Native_composed_and_receiving_PDFs_retain_every_edited_item_and_counter() {
        using var manifest = Manifest();
        foreach (JsonElement scenario in manifest.RootElement.GetProperty("cases").EnumerateArray()) {
            var paths = new List<string> { scenario.GetProperty("nativePdf").GetString()!, scenario.GetProperty("officeimoPdf").GetString()! };
            if (scenario.GetProperty("source").GetString() == "pages-preserved-decimal-v15.4.pages")
                paths.Add("pages-preserved-decimal-officeimo-docx-v15.4.pdf");
            foreach (string path in paths) {
                var pdf = PdfDocument.Load(new System.IO.MemoryStream(File.ReadAllBytes(Corpus(path)))); Assert.Single(pdf.Pages);
                string text = Regex.Replace(pdf.ExtractText(), @"\s+", string.Empty); int previous = -1;
                foreach (JsonElement item in scenario.GetProperty("expectedItems").EnumerateArray()) {
                    string expected = Regex.Replace(item.GetProperty("number").GetInt32()
                        .ToString(System.Globalization.CultureInfo.InvariantCulture) + "." + item.GetProperty("text").GetString(), @"\s+", string.Empty);
                    int at = text.IndexOf(expected, previous + 1, StringComparison.Ordinal);
                    Assert.True(at > previous, path + ": " + expected); previous = at;
                }
                foreach (string plain in new[] { "Three decimal levels", "Plain gap", "End third-level profile" })
                    Assert.Contains(Regex.Replace(plain, @"\s+", string.Empty), text);
            }
        }
    }

    private static JsonDocument Manifest() => JsonDocument.Parse(File.ReadAllBytes(Corpus("pages-preserved-decimal-v15.4.json")));
    private static string Corpus(string name) => Path.Combine(AppContext.BaseDirectory, "Documents", "IWorkCorpus", "native-exports", name);
}
