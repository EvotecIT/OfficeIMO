using OfficeIMO.IWork;
using OfficeIMO.Reader;
using OfficeIMO.Reader.IWork;
using OfficeIMO.Word;
using DocumentFormat.OpenXml.Packaging;
using W = DocumentFormat.OpenXml.Wordprocessing;

namespace OfficeIMO.IWork.Tests;

public sealed partial class IWorkEverydayNativeExportTests {
    [Fact]
    public void Native_table_counter_destination_PDFs_retain_all_markers_and_text() {
        using var manifest = System.Text.Json.JsonDocument.Parse(File.ReadAllText(
            Corpus("native-exports/list-counter-mapping-v15.4.json")));
        VerifyArtifactHashes(manifest.RootElement);
        foreach (var scenario in manifest.RootElement.GetProperty("cases").EnumerateArray()) {
            string name = scenario.GetProperty("case").GetString()!;
            foreach (string producer in new[] { "pages", "keynote", "composed" }) {
                string text = OfficeIMO.Pdf.PdfReadDocument.Open(File.ReadAllBytes(
                    Corpus("native-exports/list-counter-" + name + "-v15.4-" + producer + ".pdf"))).ExtractText();
                // Native PDF glyph positioning can add extraction whitespace
                // inside words. Check the complete ordered character sequence.
                text = System.Text.RegularExpressions.Regex.Replace(text, @"\s+", string.Empty);
                int previous = -1;
                foreach (var paragraph in scenario.GetProperty("expectedParagraphs").EnumerateArray()) {
                    var number = paragraph.GetProperty("ListNumber");
                    string marker = number.ValueKind == System.Text.Json.JsonValueKind.Number
                        ? number.GetInt32().ToString(System.Globalization.CultureInfo.InvariantCulture) + ". " : string.Empty;
                    string wanted = System.Text.RegularExpressions.Regex.Replace(
                        marker + paragraph.GetProperty("Text").GetString(), @"\s+", string.Empty);
                    int position = text.IndexOf(wanted, previous + 1, StringComparison.Ordinal);
                    Assert.True(position > previous, name + "/" + producer + ": missing or out-of-order " + wanted + "\n" + text);
                    previous = position;
                }
            }
        }
    }

    [Theory]
    [InlineData("pages-list-continuation-v15.4", false)]
    [InlineData("pages-list-marker-size-roundtrip-v15.4", true)]
    public void Native_decimal_levels_and_continuation_preserve_source_counters_and_DOCX_measurements(
        string fixture, bool reimported) {
        VerifyManifest(fixture + ".json");
        using var result = WordIWorkConverter.ConvertPagesToWordResult(
            Corpus("native-exports/" + fixture + ".pages"),
            conversionOptions: new() { Mode = IWorkConversionMode.EditableOnly });
        result.Report.RequireCompleteEditableReconstruction();
        Assert.True(result.Projection.Body.IsFormattingComplete);
        Assert.Empty(result.Projection.SourceDeclarationIssues);
        var source = result.Projection.Body.Paragraphs;
        Assert.Equal(new int?[] { null, 7, 8, null, 9, 1, 2, 10, 1, 5, 6, null, 7, 11, 3, 4 },
            source.Select(paragraph => paragraph.ListNumber));
        Assert.Equal(reimported
                ? new int?[] { null, 7, null, null, 9, null, null, 10, null, 5, null, null, 7, 11, 3, null }
                : new int?[] { null, 7, null, null, null, null, null, null, null, 5, null, null, null, 11, 3, null },
            source.Select(paragraph => paragraph.ListStartNumber));
        Assert.Equal(new[] { -1, 0, 0, -1, 0, 1, 1, 0, 1, 1, 1, -1, 1, 0, 0, 0 },
            source.Select(paragraph => paragraph.ListLevel));
        using var saved = new MemoryStream(); result.Value.Save(saved); saved.Position = 0;
        using (var counterStream = new MemoryStream(saved.ToArray()))
        using (var reopened = WordDocument.Load(counterStream)) {
            Assert.Equal(source.Where(paragraph => paragraph.ListNumber.HasValue)
                    .Select(paragraph => paragraph.ListNumber!.Value),
                WordDocumentTraversal.BuildListIndices(reopened).Values.Select(value => value.Index));
        }
        using var actual = WordprocessingDocument.Open(saved, false);
        using var native = WordprocessingDocument.Open(Corpus("native-exports/" + fixture + ".docx"), false);
        var paragraphs = actual.MainDocumentPart!.Document!.Body!.Elements<W.Paragraph>().ToArray();
        Assert.Equal(native.MainDocumentPart!.Document!.Body!.Elements<W.Paragraph>().Select(Text), paragraphs.Select(Text));
        var numbering = actual.MainDocumentPart.NumberingDefinitionsPart!.Numbering!;
        var nativeLevels = native.MainDocumentPart.NumberingDefinitionsPart!.Numbering!
            .Elements<W.AbstractNum>().SelectMany(number => number.Elements<W.Level>()).ToArray();
        for (int index = 0; index < paragraphs.Length; index++) {
            if (source[index].ListLevel < 0) continue;
            W.Paragraph paragraph = paragraphs[index];
            W.NumberingProperties selection = paragraph.ParagraphProperties!.NumberingProperties!;
            int level = source[index].ListLevel;
            Assert.Equal(level, selection.NumberingLevelReference!.Val!.Value);
            int abstractId = numbering.Elements<W.NumberingInstance>()
                .Single(number => number.NumberID!.Value == selection.NumberingId!.Val!.Value)
                .AbstractNumId!.Val!.Value;
            W.Level marker = numbering.Elements<W.AbstractNum>().Single(number => number.AbstractNumberId!.Value == abstractId)
                .Elements<W.Level>().Single(item => item.LevelIndex!.Value == level);
            string left = level == 0 ? "330" : "690";
            W.Indentation indent = paragraph.ParagraphProperties.Indentation!;
            Assert.Equal(left, indent.Left!.Value);
            Assert.Equal("330", indent.Hanging!.Value);
            Assert.Null(indent.FirstLine);
            var levelIndent = marker.GetFirstChild<W.PreviousParagraphProperties>()!.Indentation!;
            Assert.Equal(left, levelIndent.Left!.Value);
            Assert.Equal("330", levelIndent.Hanging!.Value);
            W.NumberingSymbolRunProperties font = marker.GetFirstChild<W.NumberingSymbolRunProperties>()!;
            Assert.Equal("Calibri", font.GetFirstChild<W.RunFonts>()!.Ascii!.Value);
            Assert.Equal("22", font.GetFirstChild<W.FontSize>()!.Val!.Value);
            Assert.Contains(nativeLevels, expected => expected.LevelIndex!.Value == level
                && expected.GetFirstChild<W.PreviousParagraphProperties>()?.Indentation?.Left?.Value == left
                && expected.GetFirstChild<W.PreviousParagraphProperties>()?.Indentation?.Hanging?.Value == "330"
                && expected.GetFirstChild<W.NumberingSymbolRunProperties>()?.GetFirstChild<W.FontSize>()?.Val?.Value == "22");
            // Pages uses the leading run and paragraph mark to interpret the relative marker size.
            // Both retain the 12-point body size while the numbering level stays at 11 points.
            Assert.Equal("24", paragraph.ParagraphProperties.ParagraphMarkRunProperties!
                .GetFirstChild<W.FontSize>()!.Val!.Value);
            Assert.All(paragraph.Elements<W.Run>(),
                run => Assert.Equal("24", run.RunProperties!.FontSize!.Val!.Value));
        }
        Assert.Empty(new DocumentFormat.OpenXml.Validation.OpenXmlValidator().Validate(actual));
        Assert.DoesNotContain(result.Report.Diagnostics, diagnostic => diagnostic.Code == "IWORK_PAGES_LIST_LAYOUT_APPROXIMATED");
    }

    [Theory]
    [InlineData(8192)]
    [InlineData(8)]
    public void Reader_retains_nested_and_resumed_native_markers_across_chunks(int maximumChunkCharacters) {
        var source = IWorkSourceDocument.Open(Corpus("native-exports/pages-list-continuation-v15.4.pages"));
        var output = new OfficeDocumentReadResult { Kind = ReaderInputKind.IWork };
        var reader = new IWorkReadProjection(output, "list-continuation.pages", new ReaderOptions {
            MaxChars = maximumChunkCharacters }, new ReaderIWorkOptions(), System.Threading.CancellationToken.None);
        reader.AddPages(source.ReadPages()); reader.Complete(source);
        Assert.Equal(new[] { "7.", "8.", "9.", "1.", "2.", "10.", "1.", "5.", "6.", "7.", "11.", "3.", "4." },
            output.Pages.SelectMany(page => page.Blocks).Where(block => block.Kind == "list-item").Select(block => block.Marker));
        Assert.Contains("9. Continued nine", output.Markdown);
        Assert.Contains("  7. Nested continued seven", output.Markdown);
        Assert.Contains("11. Original parent eleven", output.Markdown);
    }
}
