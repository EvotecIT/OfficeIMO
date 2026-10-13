using System.Security.Cryptography;
using System.Text.Json;
using System.Text.RegularExpressions;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Validation;
using OfficeIMO.Pdf;
using OfficeIMO.Word;
using W = DocumentFormat.OpenXml.Wordprocessing;

namespace OfficeIMO.IWork.Tests;

public sealed class IWorkPagesNumberedListNativeExportTests {
    [Fact]
    public void Accepted_decimal_example_matches_independently_decoded_native_bytes_and_artifacts() {
        using var manifest = Manifest();
        foreach (JsonProperty artifact in manifest.RootElement.GetProperty("artifacts").EnumerateObject()) {
            byte[] bytes = File.ReadAllBytes(Corpus(artifact.Value.GetProperty("path").GetString()!));
            Assert.Equal(artifact.Value.GetProperty("sha256").GetString(), Convert.ToHexString(SHA256.HashData(bytes)).ToLowerInvariant());
        }
        Assert.Equal(File.ReadAllBytes(Corpus("pages-created-numbering-v15.4.pages")),
            IWorkPagesNumberedListWritingTests.Example().SaveBytes());
    }

    [Theory]
    [InlineData(0)]
    [InlineData(1)]
    public void Apple_accepted_and_edited_lists_keep_counters_indents_and_objects_in_saved_Word(int caseIndex) {
        using var manifest = Manifest();
        JsonElement scenario = manifest.RootElement.GetProperty("cases")[caseIndex];
        var source = IWorkSourceDocument.Open(Artifact(manifest.RootElement, scenario.GetProperty("source").GetString()!));
        var projection = source.ReadPages();
        Assert.Empty(projection.Diagnostics); Assert.Empty(projection.SourceDeclarationIssues); Assert.Empty(projection.SourceReferenceIssues);
        Assert.True(projection.Body.IsFormattingComplete);
        IWorkTextParagraph[] items = projection.Body.Paragraphs.Where(paragraph => paragraph.ListNumber.HasValue).ToArray();
        JsonElement[] expected = scenario.GetProperty("expectedItems").EnumerateArray().ToArray();
        Assert.Equal(expected.Select(item => item.GetProperty("number").GetInt32()), items.Select(item => item.ListNumber!.Value));
        Assert.Equal(expected.Select(item => item.GetProperty("level").GetInt32()), items.Select(item => item.ListLevel));
        Assert.Equal(expected.Select(item => item.GetProperty("text").GetString()), items.Select(item => item.Text));
        Assert.All(items, item => Assert.Equal(12, item.Style.TextStyle.FontSizePoints));
        Assert.Single(projection.Tables); Assert.Single(projection.Images);
        Assert.Null(projection.Body.Paragraphs.Last().ListNumber);
        using var result = source.ToWordDocumentResult(new() { Mode = IWorkConversionMode.EditableOnly });
        result.RequireCompleteEditableReconstruction();
        using var output = new MemoryStream(); result.Value.Save(output); output.Position = 0;
        using (var reopened = WordDocument.Load(new MemoryStream(output.ToArray())))
            Assert.Equal(items.Select(item => item.ListNumber!.Value),
                WordDocumentTraversal.BuildListIndices(reopened).Values.Select(item => item.Index));
        using var saved = WordprocessingDocument.Open(output, false); Assert.Empty(new OpenXmlValidator().Validate(saved));
        var numbered = saved.MainDocumentPart!.Document!.Body!.Elements<W.Paragraph>()
            .Where(paragraph => paragraph.ParagraphProperties?.NumberingProperties != null).ToArray();
        Assert.Equal(items.Length, numbered.Length);
        using var native = WordprocessingDocument.Open(Corpus(caseIndex == 0
            ? "pages-created-numbering-v15.4.docx" : "pages-created-numbering-edited-v15.4.docx"), false);
        W.Level[] nativeLevels = native.MainDocumentPart!.NumberingDefinitionsPart!.Numbering!
            .Elements<W.AbstractNum>().SelectMany(number => number.Elements<W.Level>()).ToArray();
        for (int index = 0; index < numbered.Length; index++) {
            W.Paragraph paragraph = numbered[index]; W.Indentation indent = paragraph.ParagraphProperties!.Indentation!;
            string left = items[index].ListLevel == 0 ? "360" : "720";
            Assert.Equal(left, indent.Left!.Value); Assert.Equal("360", indent.Hanging!.Value);
            Assert.Contains(nativeLevels, level => level.LevelIndex!.Value == items[index].ListLevel
                && level.GetFirstChild<W.PreviousParagraphProperties>()?.Indentation?.Left?.Value == left
                && level.GetFirstChild<W.PreviousParagraphProperties>()?.Indentation?.Hanging?.Value == "360");
        }
        Assert.Single(saved.MainDocumentPart.Document.Body.Elements<W.Table>());
        Assert.Single(saved.MainDocumentPart.Document.Body.Descendants<W.Drawing>());
    }

    [Fact]
    public void Native_and_composed_reference_PDFs_retain_every_original_and_inserted_marker() {
        using var manifest = Manifest();
        foreach (JsonElement scenario in manifest.RootElement.GetProperty("cases").EnumerateArray()) {
            foreach (string producer in new[] { "nativePdf", "officeimoPdf" }) {
                PdfReadDocument pdf = PdfReadDocument.Open(File.ReadAllBytes(Artifact(manifest.RootElement,
                    scenario.GetProperty(producer).GetString()!)));
                Assert.Single(pdf.Pages);
                string text = Regex.Replace(pdf.ExtractText(), @"\s+", string.Empty); int previous = -1;
                foreach (JsonElement item in scenario.GetProperty("expectedItems").EnumerateArray()) {
                    string wanted = Regex.Replace(item.GetProperty("number").GetInt32()
                        .ToString(System.Globalization.CultureInfo.InvariantCulture) + "." + item.GetProperty("text").GetString(), @"\s+", string.Empty);
                    int at = text.IndexOf(wanted, previous + 1, StringComparison.Ordinal);
                    Assert.True(at > previous, scenario.GetProperty("case").GetString() + "/" + producer + ": " + wanted);
                    previous = at;
                }
                Assert.Contains("Betweenlistitems", text);
            }
        }
    }

    [Fact]
    public void A_list_only_body_matches_the_Apple_accepted_profile_without_a_trailing_marker() {
        var document = IWorkPagesDocument.Create(); var list = document.CreateNumberedList();
        list.AddItem("Only parent"); list.AddItem("Only child", 1);
        Assert.Equal(File.ReadAllBytes(Corpus("pages-created-numbering-only-v15.4.pages")), document.SaveBytes());
        var source = IWorkSourceDocument.Open(document.SaveBytes()); var projection = source.ReadPages();
        Assert.True(projection.Body.IsFormattingComplete); Assert.Empty(projection.Diagnostics);
        Assert.Equal(new int?[] { 1, 1, null }, projection.Body.Paragraphs.Select(paragraph => paragraph.ListNumber));
        using var word = source.ToWordDocumentResult(new() { Mode = IWorkConversionMode.EditableOnly });
        word.RequireCompleteEditableReconstruction();
    }

    private static string Corpus(string name) => Path.Combine(AppContext.BaseDirectory, "Documents", "IWorkCorpus", "native-exports", name);
    private static JsonDocument Manifest() => JsonDocument.Parse(File.ReadAllText(Corpus("pages-created-numbering-v15.4.json")));
    private static string Artifact(JsonElement manifest, string key) => Corpus(manifest.GetProperty("artifacts").GetProperty(key).GetProperty("path").GetString()!);
}
