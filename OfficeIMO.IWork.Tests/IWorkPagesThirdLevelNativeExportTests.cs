using System.Security.Cryptography;
using System.Text.Json;
using System.Text.RegularExpressions;
using DocumentFormat.OpenXml.Packaging;
using OfficeIMO.Pdf;
using OfficeIMO.Reader.IWork;
using OfficeIMO.Word;
using W = DocumentFormat.OpenXml.Wordprocessing;

namespace OfficeIMO.IWork.Tests;

public sealed class IWorkPagesThirdLevelNativeExportTests {
    [Fact]
    public void Authored_third_level_matches_the_independently_decoded_Apple_accepted_input() {
        using var manifest = Manifest();
        foreach (JsonProperty artifact in manifest.RootElement.GetProperty("artifacts").EnumerateObject()) {
            byte[] bytes = File.ReadAllBytes(Corpus(artifact.Value.GetProperty("path").GetString()!));
            Assert.Equal(artifact.Value.GetProperty("bytes").GetInt64(), bytes.LongLength);
            Assert.Equal(artifact.Value.GetProperty("sha256").GetString(), Convert.ToHexString(SHA256.HashData(bytes)).ToLowerInvariant());
        }
        Assert.Equal(File.ReadAllBytes(Corpus("pages-third-level-created.pages")),
            IWorkPagesThirdLevelWritingTests.Example().SaveBytes());
    }

    [Theory]
    [InlineData(0)]
    [InlineData(1)]
    [InlineData(2)]
    public void Generated_saved_and_Apple_edited_third_levels_retain_saved_Word_counters_and_indents(int caseIndex) {
        using var manifest = Manifest();
        JsonElement scenario = manifest.RootElement.GetProperty("cases")[caseIndex];
        var source = IWorkSourceDocument.Open(Corpus(scenario.GetProperty("source").GetString()!));
        var projection = source.ReadPages();
        Assert.True(projection.Body.IsFormattingComplete); Assert.Empty(projection.Diagnostics);
        Assert.Empty(projection.SourceDeclarationIssues); Assert.Empty(projection.SourceReferenceIssues);
        var items = projection.Body.Paragraphs.Where(p => p.ListNumber.HasValue).ToArray();
        var expected = scenario.GetProperty("expectedItems").EnumerateArray().ToArray();
        Assert.Equal(expected.Select(p => p.GetProperty("number").GetInt32()), items.Select(p => p.ListNumber!.Value));
        Assert.Equal(expected.Select(p => p.GetProperty("level").GetInt32()), items.Select(p => p.ListLevel));
        Assert.Equal(expected.Select(p => p.GetProperty("text").GetString()), items.Select(p => p.Text));
        foreach (var item in items)
            Assert.Equal(new string(' ', item.ListLevel * 2) + item.ListNumber + ". " + item.Text,
                IWorkReadProjection.RichTextMarkdown(item));
        using var result = source.ToWordDocumentResult(new() { Mode = IWorkConversionMode.EditableOnly });
        result.RequireCompleteEditableReconstruction();
        using var output = new MemoryStream(); result.Value.Save(output);
        using (var reopened = WordDocument.Load(new MemoryStream(output.ToArray()))) {
            Assert.Empty(reopened.ValidateDocument());
            Assert.Equal(items.Select(p => p.ListNumber!.Value),
                WordDocumentTraversal.BuildListIndices(reopened).Values.Select(p => p.Index));
        }
        using var xml = WordprocessingDocument.Open(new MemoryStream(output.ToArray()), false);
        var paragraphs = xml.MainDocumentPart!.Document!.Body!.Elements<W.Paragraph>()
            .Where(p => p.ParagraphProperties?.NumberingProperties != null).ToArray();
        for (int index = 0; index < items.Length; index++) {
            var properties = paragraphs[index].ParagraphProperties!;
            Assert.Equal(items[index].ListLevel, properties.NumberingProperties!.NumberingLevelReference!.Val!.Value);
            Assert.Equal(((items[index].ListLevel + 1) * 360).ToString(System.Globalization.CultureInfo.InvariantCulture), properties.Indentation!.Left!.Value);
            Assert.Equal("360", properties.Indentation.Hanging!.Value);
        }
    }

    [Fact]
    public void Complete_native_and_composed_PDFs_retain_every_third_level_marker_and_plain_paragraph() {
        using var manifest = Manifest();
        foreach (JsonElement scenario in manifest.RootElement.GetProperty("cases").EnumerateArray()) {
            var paths = new List<string> { scenario.GetProperty("officeimoPdf").GetString()! };
            if (scenario.GetProperty("nativePdf").ValueKind == JsonValueKind.String)
                paths.Add(scenario.GetProperty("nativePdf").GetString()!);
            else paths.Add("pages-third-level-officeimo-docx-v15.4.pdf");
            foreach (string path in paths) {
                var pdf = PdfReadDocument.Open(File.ReadAllBytes(Corpus(path))); Assert.Single(pdf.Pages);
                string text = Regex.Replace(pdf.ExtractText(), @"\s+", string.Empty); int previous = -1;
                foreach (JsonElement item in scenario.GetProperty("expectedItems").EnumerateArray()) {
                    string wanted = Regex.Replace(item.GetProperty("number").GetInt32()
                        .ToString(System.Globalization.CultureInfo.InvariantCulture) + "." + item.GetProperty("text").GetString(), @"\s+", string.Empty);
                    int at = text.IndexOf(wanted, previous + 1, StringComparison.Ordinal);
                    Assert.True(at > previous, path + ": " + wanted); previous = at;
                }
                foreach (string plain in new[] { "Three decimal levels", "Plain gap", "End third-level profile" })
                    Assert.Contains(Regex.Replace(plain, @"\s+", string.Empty), text);
            }
        }
    }

    private static JsonDocument Manifest() => JsonDocument.Parse(File.ReadAllBytes(Corpus("pages-third-level-v15.4.json")));
    private static string Corpus(string name) => Path.Combine(AppContext.BaseDirectory, "Documents", "IWorkCorpus", "native-exports", name);
}
