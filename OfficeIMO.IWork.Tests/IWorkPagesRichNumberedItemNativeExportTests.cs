using System.Security.Cryptography;
using System.Text.Json;
using System.Text.RegularExpressions;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Validation;
using OfficeIMO.Pdf;
using OfficeIMO.Word;
using W = DocumentFormat.OpenXml.Wordprocessing;

namespace OfficeIMO.IWork.Tests;

public sealed class IWorkPagesRichNumberedItemNativeExportTests {
    [Fact]
    public void Rich_creation_matches_the_independently_decoded_Apple_accepted_package() {
        using var manifest = Manifest();
        foreach (JsonProperty artifact in manifest.RootElement.GetProperty("artifacts").EnumerateObject()) {
            byte[] bytes = File.ReadAllBytes(Corpus(artifact.Value.GetProperty("path").GetString()!));
            Assert.Equal(artifact.Value.GetProperty("sha256").GetString(), Convert.ToHexString(SHA256.HashData(bytes)).ToLowerInvariant());
        }
        Assert.Equal(File.ReadAllBytes(Artifact(manifest.RootElement, "created")),
            IWorkPagesRichNumberedItemWritingTests.Example().SaveBytes());
    }

    [Theory]
    [InlineData(0)]
    [InlineData(1)]
    public void Native_edit_save_reopen_retains_rich_runs_plain_markers_and_strict_Word(int caseIndex) {
        using var manifest = Manifest();
        JsonElement scenario = manifest.RootElement.GetProperty("cases")[caseIndex];
        var source = IWorkSourceDocument.Open(Artifact(manifest.RootElement, scenario.GetProperty("source").GetString()!));
        var projection = source.ReadPages();
        Assert.Empty(projection.Diagnostics); Assert.Empty(projection.SourceDeclarationIssues); Assert.Empty(projection.SourceReferenceIssues);
        Assert.True(projection.Body.IsFormattingComplete);
        var items = projection.Body.Paragraphs.Where(paragraph => paragraph.ListNumber.HasValue).ToArray();
        var expected = scenario.GetProperty("expectedItems").EnumerateArray().ToArray();
        Assert.Equal(expected.Select(item => item.GetProperty("number").GetInt32()), items.Select(item => item.ListNumber!.Value));
        Assert.Equal(expected.Select(item => item.GetProperty("level").GetInt32()), items.Select(item => item.ListLevel));
        Assert.Equal(expected.Select(item => item.GetProperty("text").GetString()), items.Select(item => item.Text));
        for (int index = 0; index < items.Length; index++) {
            Assert.Equal(12, items[index].Style.TextStyle.FontSizePoints);
            Assert.Equal("000000", items[index].Style.TextStyle.Color!.RgbHex);
            Assert.False(items[index].Style.TextStyle.Bold); Assert.False(items[index].Style.TextStyle.Italic);
            Assert.False(items[index].Style.TextStyle.Underline); Assert.False(items[index].Style.TextStyle.Strikethrough);
            var runs = expected[index].GetProperty("runs").EnumerateArray().ToArray();
            Assert.Equal(runs.Select(run => run.GetProperty("text").GetString()), items[index].Runs.Select(run => run.Text));
            for (int run = 0; run < runs.Length; run++) {
                var style = items[index].Runs[run].Style;
                Assert.Equal(runs[run].GetProperty("bold").GetBoolean(), style.Bold);
                Assert.Equal(runs[run].GetProperty("italic").GetBoolean(), style.Italic);
                Assert.Equal(runs[run].GetProperty("underline").GetBoolean(), style.Underline);
                Assert.Equal(runs[run].GetProperty("strikethrough").GetBoolean(), style.Strikethrough);
                Assert.Equal(runs[run].GetProperty("color").GetString(), style.Color!.RgbHex);
            }
        }
        using var result = source.ToWordDocumentResult(new() { Mode = IWorkConversionMode.EditableOnly });
        result.RequireCompleteEditableReconstruction(); using var output = new MemoryStream(); result.Value.Save(output);
        using (var reopened = WordDocument.Load(new MemoryStream(output.ToArray())))
            Assert.Equal(items.Select(item => item.ListNumber!.Value),
                WordDocumentTraversal.BuildListIndices(reopened).Values.Select(item => item.Index));
        output.Position = 0; using var saved = WordprocessingDocument.Open(output, false);
        Assert.Empty(new OpenXmlValidator().Validate(saved));
        var numbered = saved.MainDocumentPart!.Document!.Body!.Elements<W.Paragraph>()
            .Where(paragraph => paragraph.ParagraphProperties?.NumberingProperties != null).ToArray();
        Assert.Equal(expected.Length, numbered.Length);
        for (int index = 0; index < numbered.Length; index++) {
            var runs = numbered[index].Elements<W.Run>().Where(run => run.InnerText.Length > 0).ToArray();
            var wanted = expected[index].GetProperty("runs").EnumerateArray().ToArray();
            Assert.Equal(wanted.Select(run => run.GetProperty("text").GetString()), runs.Select(run => run.InnerText));
            for (int run = 0; run < runs.Length; run++) {
                var properties = runs[run].RunProperties!;
                Assert.Equal(wanted[run].GetProperty("bold").GetBoolean(), Enabled(properties.Bold));
                Assert.Equal(wanted[run].GetProperty("italic").GetBoolean(), Enabled(properties.Italic));
                Assert.Equal(wanted[run].GetProperty("underline").GetBoolean(),
                    properties.Underline?.Val?.Value == W.UnderlineValues.Single);
                Assert.Equal(wanted[run].GetProperty("strikethrough").GetBoolean(), Enabled(properties.Strike));
                Assert.Equal(wanted[run].GetProperty("color").GetString(), properties.Color!.Val!.Value);
            }
        }
        if (caseIndex == 1) {
            using var native = WordprocessingDocument.Open(Artifact(manifest.RootElement, "nativeEditedDocx"), false);
            Assert.Equal(numbered.Select(paragraph => paragraph.InnerText),
                native.MainDocumentPart!.Document!.Body!.Elements<W.Paragraph>()
                    .Where(paragraph => paragraph.ParagraphProperties?.NumberingProperties != null).Select(paragraph => paragraph.InnerText));
        }
    }

    [Fact]
    public void Native_and_composed_PDFs_retain_all_counters_and_independent_body_emphasis() {
        using var manifest = Manifest();
        foreach (JsonElement scenario in manifest.RootElement.GetProperty("cases").EnumerateArray()) {
            foreach (string producer in new[] { "nativePdf", "officeimoPdf" }) {
                if (!scenario.TryGetProperty(producer, out JsonElement artifact)) continue;
                PdfDocument pdf = PdfDocument.Load(new System.IO.MemoryStream(File.ReadAllBytes(Artifact(manifest.RootElement, artifact.GetString()!))));
                var page = Assert.Single(pdf.Pages); string text = Compact(page.ExtractText()); int previous = -1;
                foreach (JsonElement item in scenario.GetProperty("expectedItems").EnumerateArray()) {
                    string wanted = Compact(item.GetProperty("number").GetInt32().ToString(System.Globalization.CultureInfo.InvariantCulture)
                        + "." + item.GetProperty("text").GetString());
                    int at = text.IndexOf(wanted, previous + 1, StringComparison.Ordinal);
                    Assert.True(at > previous, scenario.GetProperty("case").GetString() + "/" + producer + ": " + wanted); previous = at;
                }
                var spans = page.GetTextSpans();
                Assert.Contains("Boldfirst", Compact(string.Concat(spans.Where(span => span.IsBold && !span.IsItalic).Select(span => span.Text))));
                Assert.Contains("Childitalic", Compact(string.Concat(spans.Where(span => !span.IsBold && span.IsItalic).Select(span => span.Text))));
                Assert.Contains("Combineddecoration", Compact(string.Concat(spans.Where(span => span.IsBold && span.IsItalic).Select(span => span.Text))));
                Assert.Contains(spans, span => span.IsBold && !span.IsItalic && span.Color?.ToString() == "#003366");
                Assert.Contains(spans, span => !span.IsBold && span.IsItalic && span.Color?.ToString() == "#008000");
                Assert.Contains(spans, span => span.IsBold && span.IsItalic && span.Color?.ToString() == "#660066");
                Assert.All(spans.Where(span => Regex.IsMatch(span.Text.Trim(), @"^\d+\.$")), span => {
                    Assert.False(span.IsBold); Assert.False(span.IsItalic);
                });
                Assert.Contains("Plaingap", text);
            }
        }
    }

    private static bool Enabled(W.OnOffType? value) => value != null && value.Val?.Value != false;
    private static string Compact(string text) => Regex.Replace(text, @"\s+", string.Empty);
    private static string Corpus(string name) => Path.Combine(AppContext.BaseDirectory, "Documents", "IWorkCorpus", "native-exports", name);
    private static JsonDocument Manifest() => JsonDocument.Parse(File.ReadAllText(Corpus("pages-created-rich-numbering-v15.4.json")));
    private static string Artifact(JsonElement manifest, string key) => Corpus(manifest.GetProperty("artifacts").GetProperty(key).GetProperty("path").GetString()!);
}
