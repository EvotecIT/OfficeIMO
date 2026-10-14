using System.Security.Cryptography;
using System.Text.Json;
using System.Text.RegularExpressions;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Validation;
using OfficeIMO.Pdf;
using OfficeIMO.Word;
using W = DocumentFormat.OpenXml.Wordprocessing;

namespace OfficeIMO.IWork.Tests;

public sealed class IWorkPagesMixedFontListNativeExportTests {
    [Fact]
    public void Creation_matches_the_independently_decoded_Apple_accepted_package() {
        using var manifest = JsonDocument.Parse(File.ReadAllBytes(Corpus("pages-marker-identity-v15.4.json")));
        foreach (JsonProperty artifact in manifest.RootElement.GetProperty("artifacts").EnumerateObject()) {
            byte[] bytes = File.ReadAllBytes(Corpus(artifact.Name));
            Assert.Equal(artifact.Value.GetProperty("bytes").GetInt64(), bytes.LongLength);
            Assert.Equal(artifact.Value.GetProperty("sha256").GetString(),
                Convert.ToHexString(SHA256.HashData(bytes)).ToLowerInvariant());
        }
        Assert.Equal(File.ReadAllBytes(Corpus("pages-identity-mixed-font-created.pages")),
            IWorkPagesMixedFontListTests.Example().SaveBytes());
    }

    [Theory]
    [InlineData(0)]
    [InlineData(1)]
    [InlineData(2)]
    public void Native_save_edit_reopen_and_strict_Word_retain_families_and_owned_markers(int caseIndex) {
        using var manifest = Manifest();
        JsonElement scenario = manifest.RootElement.GetProperty("cases")[caseIndex];
        var source = IWorkSourceDocument.Open(Corpus(scenario.GetProperty("source").GetString()!));
        var projection = source.ReadPages();
        Assert.True(projection.Body.IsFormattingComplete);
        Assert.Empty(projection.Diagnostics); Assert.Empty(projection.SourceDeclarationIssues); Assert.Empty(projection.SourceReferenceIssues);
        var expected = scenario.GetProperty("expectedItems").EnumerateArray().ToArray();
        var items = projection.Body.Paragraphs.Where(item => item.ListNumber.HasValue).ToArray();
        Assert.Equal(expected.Select(item => item.GetProperty("number").GetInt32()), items.Select(item => item.ListNumber!.Value));
        Assert.Equal(expected.Select(item => item.GetProperty("level").GetInt32()), items.Select(item => item.ListLevel));
        Assert.Equal(expected.Select(item => item.GetProperty("text").GetString()), items.Select(item => item.Text));
        for (int index = 0; index < items.Length; index++) {
            string family = expected[index].GetProperty("markerFont").GetString()!;
            Assert.Equal(family, items[index].ListFontName);
            Assert.Equal(family, items[index].Style.TextStyle.FontName);
            Assert.Equal(12, items[index].Style.TextStyle.FontSizePoints);
            Assert.Equal("000000", items[index].Style.TextStyle.Color!.RgbHex);
            Assert.False(items[index].Style.TextStyle.Bold); Assert.False(items[index].Style.TextStyle.Italic);
            var runs = expected[index].GetProperty("runs").EnumerateArray().ToArray();
            Assert.Equal(runs.Select(run => run.GetProperty("text").GetString()), items[index].Runs.Select(run => run.Text));
            for (int run = 0; run < runs.Length; run++) {
                var style = items[index].Runs[run].Style;
                Assert.Equal(runs[run].GetProperty("family").GetString(), style.FontName);
                Assert.Equal(12, style.FontSizePoints);
                Assert.Equal(runs[run].GetProperty("bold").GetBoolean(), style.Bold);
                Assert.Equal(runs[run].GetProperty("italic").GetBoolean(), style.Italic);
                Assert.Equal(runs[run].GetProperty("color").GetString(), style.Color!.RgbHex);
            }
        }
        using var result = source.ToWordDocumentResult(new() { Mode = IWorkConversionMode.EditableOnly });
        result.RequireCompleteEditableReconstruction(); using var output = new MemoryStream(); result.Value.Save(output);
        using (var reopened = WordDocument.Load(new MemoryStream(output.ToArray()))) {
            Assert.Equal(expected.Select(item => item.GetProperty("number").GetInt32()),
                WordDocumentTraversal.BuildListIndices(reopened).Values.Select(item => item.Index));
            Assert.Equal(expected.Select(item => item.GetProperty("marker").GetString()),
                WordDocumentTraversal.BuildListMarkers(reopened).Values.Select(item => item.Marker));
        }
        output.Position = 0; using var saved = WordprocessingDocument.Open(output, false);
        Assert.Empty(new OpenXmlValidator().Validate(saved));
        var numbered = saved.MainDocumentPart!.Document!.Body!.Elements<W.Paragraph>()
            .Where(item => item.ParagraphProperties?.NumberingProperties != null).ToArray();
        Assert.Equal(expected.Length, numbered.Length);
        for (int index = 0; index < numbered.Length; index++) {
            var runs = numbered[index].Elements<W.Run>().Where(run => run.InnerText.Length > 0).ToArray();
            var wanted = expected[index].GetProperty("runs").EnumerateArray().ToArray();
            Assert.Equal(wanted.Select(run => run.GetProperty("family").GetString()),
                runs.Select(run => run.RunProperties!.RunFonts!.Ascii!.Value));
            Assert.Equal(wanted.Select(run => run.GetProperty("bold").GetBoolean()),
                runs.Select(run => Enabled(run.RunProperties!.Bold)));
            Assert.Equal(wanted.Select(run => run.GetProperty("italic").GetBoolean()),
                runs.Select(run => Enabled(run.RunProperties!.Italic)));
        }
    }

    [Fact]
    public void Every_native_composed_and_receiving_PDF_retains_all_items_and_plain_markers() {
        using var manifest = Manifest();
        foreach (JsonElement scenario in manifest.RootElement.GetProperty("cases").EnumerateArray()) {
            Check(scenario.GetProperty("portablePdf").GetString()!, scenario);
            Check(scenario.GetProperty("callerFontsPdf").GetString()!, scenario);
            if (scenario.GetProperty("nativePdf").ValueKind == JsonValueKind.String)
                Check(scenario.GetProperty("nativePdf").GetString()!, scenario);
        }
        Check(manifest.RootElement.GetProperty("receivingDocxPdf").GetString()!, manifest.RootElement.GetProperty("cases")[0]);

        static void Check(string path, JsonElement scenario) {
            var pdf = PdfReadDocument.Open(File.ReadAllBytes(Corpus(path))); var page = Assert.Single(pdf.Pages);
            string text = Compact(page.ExtractText()); int previous = -1;
            foreach (JsonElement item in scenario.GetProperty("expectedItems").EnumerateArray()) {
                string wanted = Compact(item.GetProperty("marker").GetString()! + item.GetProperty("text").GetString());
                int at = text.IndexOf(wanted, previous + 1, StringComparison.Ordinal);
                Assert.True(at > previous, path + ": " + wanted); previous = at;
            }
            foreach (JsonElement plain in scenario.GetProperty("expectedPlainParagraphs").EnumerateArray())
                Assert.Contains(Compact(plain.GetString()!), text);
            var spans = page.GetTextSpans();
            Assert.Contains(spans, span => span.IsBold && !span.IsItalic && span.Color?.ToString() == "#003366");
            Assert.Contains(spans, span => !span.IsBold && span.IsItalic && span.Color?.ToString() == "#008000");
            Assert.All(spans.Where(span => Regex.IsMatch(span.Text.Trim(), @"^(\d+\.|\(\d+\)|\d+\))$")), span => {
                Assert.False(span.IsBold); Assert.False(span.IsItalic);
            });
        }
    }

    private static bool Enabled(W.OnOffType? value) => value != null && value.Val?.Value != false;
    private static string Compact(string text) => Regex.Replace(text, @"\s+", string.Empty);
    private static JsonDocument Manifest() => JsonDocument.Parse(File.ReadAllBytes(Corpus("pages-mixed-font-lists-v15.4.json")));
    private static string Corpus(string name) => Path.Combine(AppContext.BaseDirectory, "Documents", "IWorkCorpus", "native-exports", name);
}
