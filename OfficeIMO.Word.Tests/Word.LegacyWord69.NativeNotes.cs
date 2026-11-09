using System.Text;
using System.Text.Json;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Validation;
using OfficeIMO.Word;
using W = DocumentFormat.OpenXml.Wordprocessing;
using Xunit;

namespace OfficeIMO.Tests;

public partial class Word {
    [Fact]
    public void LegacyWord69_OpenOfficeNotesRetainActualAnchorsTabsFormattingAndEditableDocx() {
        string path = GetFixtureDoc(Path.Combine("LegacyDocCorpus", "word69-notes", "apache-openoffice-4.1.17-notes.doc"));
        using JsonDocument oracle = JsonDocument.Parse(File.ReadAllText(Path.ChangeExtension(path, ".oracle.json")));
        byte[] original = File.ReadAllBytes(path);
        using WordDocument document = WordDocument.Load(path);
        Assert.DoesNotContain(document.LegacyDocUnsupportedFeatures, finding =>
            finding.Code.StartsWith("doc-word69-note", StringComparison.Ordinal)
            || finding.Code is "doc-word69-footnotes" or "doc-word69-endnotes");
        AssertOpenOfficeWord69Notes(document, oracle.RootElement, edited: false);

        document.FootNotes[0].Paragraphs!.Single(run => run.Text == "bold").SetText("editable bold");
        using var output = new MemoryStream();
        document.Save(output, new WordSaveOptions { AllowLoss = true });
        output.Position = 0;
        using WordDocument reopened = WordDocument.Load(output);
        AssertOpenOfficeWord69Notes(reopened, oracle.RootElement, edited: true);
        Assert.Empty(new OpenXmlValidator(FileFormatVersions.Microsoft365).Validate(reopened.OpenXmlDocument));
        Assert.Equal(original, File.ReadAllBytes(path));
    }

    private static void AssertOpenOfficeWord69Notes(WordDocument document, JsonElement oracle, bool edited) {
        Assert.Equal(2, document.FootNotes.Count);
        Assert.Equal(2, document.EndNotes.Count);
        W.Body body = document.OpenXmlDocument.MainDocumentPart!.Document.Body!;
        JsonElement[] expectedNotes = oracle.GetProperty("notes").EnumerateArray().ToArray();
        string[] expectedOrder = oracle.GetProperty("bodyAnchorOrder").EnumerateArray().Select(item => item.GetString()!).ToArray();
        string sourceBody = oracle.GetProperty("bodyRawText").GetString()!;
        var sourcePrefix = new StringBuilder();
        var references = new List<(string Kind, long Id, JsonElement Expected)>();
        foreach (W.Paragraph paragraph in body.Elements<W.Paragraph>()) {
            foreach (OpenXmlElement element in paragraph.Descendants()) {
                if (element is W.Text text) sourcePrefix.Append(text.Text);
                if (element is W.TabChar) sourcePrefix.Append('\t');
                if (element is W.FootnoteReference || element is W.EndnoteReference) {
                    string key = expectedOrder[references.Count];
                    JsonElement expected = Assert.Single(expectedNotes, note => note.GetProperty("key").GetString() == key);
                    string kind = element is W.FootnoteReference ? "footnote" : "endnote";
                    long id = element is W.FootnoteReference footnote ? footnote.Id!.Value : ((W.EndnoteReference)element).Id!.Value;
                    Assert.Equal(expected.GetProperty("kind").GetString(), kind);
                    Assert.Equal(sourceBody.Substring(0, expected.GetProperty("bodyAnchorCp").GetInt32()), sourcePrefix.ToString());
                    references.Add((kind, id, expected));
                    sourcePrefix.Append('\u0002');
                }
            }
            sourcePrefix.Append('\r');
        }
        Assert.Equal(expectedOrder.Length, references.Count);

        WordRunInfo[] inspectedRuns = document.Inspect().Sections.SelectMany(section => section.Elements)
            .OfType<WordParagraphInfo>().SelectMany(paragraph => paragraph.Runs).ToArray();
        Assert.Equal(2, inspectedRuns.Count(run => run.Footnote != null));
        Assert.Equal(2, inspectedRuns.Count(run => run.Endnote != null));
        foreach (var reference in references) {
            bool footnote = reference.Kind == "footnote";
            OpenXmlCompositeElement note = footnote
                ? Assert.Single(document.OpenXmlDocument.MainDocumentPart.FootnotesPart!.Footnotes.Elements<W.Footnote>(), item => item.Id?.Value == reference.Id)
                : Assert.Single(document.OpenXmlDocument.MainDocumentPart.EndnotesPart!.Endnotes.Elements<W.Endnote>(), item => item.Id?.Value == reference.Id);
            string key = reference.Expected.GetProperty("key").GetString()!;
            string[] expectedParagraphs = reference.Expected.GetProperty("binaryCanonicalParagraphs").EnumerateArray()
                .Select(item => edited && key == "F1" ? item.GetString()!.Replace("bold", "editable bold") : item.GetString()!).ToArray();
            W.Paragraph[] paragraphs = note.Elements<W.Paragraph>().ToArray();
            Assert.Equal(expectedParagraphs, paragraphs.Select(ReadWord69NoteParagraphText));
            JsonElement[] paragraphFormats = reference.Expected.GetProperty("paragraphFormatting").EnumerateArray().ToArray();
            for (int index = 0; index < paragraphs.Length; index++) {
                AssertOpenOfficeWord69NoteParagraphFormat(document, paragraphs[index], paragraphFormats[index]);
            }
            IReadOnlyList<WordParagraphInfo> inspectedParagraphs = footnote
                ? Assert.Single(inspectedRuns, run => run.Footnote?.ReferenceId == reference.Id).Footnote!.Paragraphs
                : Assert.Single(inspectedRuns, run => run.Endnote?.ReferenceId == reference.Id).Endnote!.Paragraphs;
            Assert.Equal(expectedParagraphs, inspectedParagraphs.Select(paragraph => paragraph.Text));

            IReadOnlyList<WordParagraph> handles = footnote
                ? Assert.Single(document.FootNotes, item => item.ReferenceId == reference.Id).Paragraphs!
                : Assert.Single(document.EndNotes, item => item.ReferenceId == reference.Id).Paragraphs!;
            foreach (JsonElement span in reference.Expected.GetProperty("selectedFormatting").EnumerateArray()) {
                string selectedText = span.GetProperty("text").GetString()!;
                if (edited && key == "F1" && selectedText == "bold") selectedText = "editable bold";
                int paragraphIndex = span.GetProperty("paragraphIndex").GetInt32();
                WordParagraph selected = Assert.Single(handles, run => run._paragraph == paragraphs[paragraphIndex] && run.Text == selectedText);
                JsonElement expectedFormat = Assert.Single(span.GetProperty("effectivePortions").EnumerateArray());
                var sources = new List<OpenXmlCompositeElement>();
                if (selected._run.RunProperties != null) sources.Add(selected._run.RunProperties);
                sources.AddRange(WordRunFormatting.EnumerateRunStyleProperties(selected));
                W.RunPropertiesBaseStyle? defaults = document.OpenXmlDocument.MainDocumentPart.StyleDefinitionsPart?.Styles?.DocDefaults?.RunPropertiesDefault?.RunPropertiesBaseStyle;
                if (defaults != null) sources.Add(defaults);
                W.Bold? bold = sources.Select(item => item.GetFirstChild<W.Bold>()).FirstOrDefault(value => value != null);
                W.Italic? italic = sources.Select(item => item.GetFirstChild<W.Italic>()).FirstOrDefault(value => value != null);
                Assert.Equal(expectedFormat.GetProperty("weight").GetDouble() >= 150, bold != null && (bold.Val?.Value ?? true));
                Assert.Equal(expectedFormat.GetProperty("posture").GetString() == "ITALIC", italic != null && (italic.Val?.Value ?? true));
                Assert.Equal(expectedFormat.GetProperty("font").GetString(), sources.Select(item => item.GetFirstChild<W.RunFonts>()?.Ascii?.Value).FirstOrDefault(value => value != null));
                string? size = sources.Select(item => item.GetFirstChild<W.FontSize>()?.Val?.Value).FirstOrDefault(value => value != null);
                Assert.NotNull(size);
                Assert.Equal(expectedFormat.GetProperty("sizePoints").GetDouble(), double.Parse(size!, System.Globalization.CultureInfo.InvariantCulture) / 2);
            }
        }
        foreach (WordSection section in document.Sections) {
            Assert.Equal(W.NumberFormatValues.Decimal, section.FootnoteProperties!.NumberingFormat!.Val!.Value);
            Assert.Equal(1, section.FootnoteProperties.NumberingStart!.Val!.Value);
            Assert.Equal(W.RestartNumberValues.Continuous, section.FootnoteProperties.NumberingRestart!.Val!.Value);
            Assert.Equal(W.FootnotePositionValues.PageBottom, section.FootnoteProperties.FootnotePosition!.Val!.Value);
            Assert.Equal(W.NumberFormatValues.Decimal, section.EndnoteProperties!.NumberingFormat!.Val!.Value);
            Assert.Equal(1, section.EndnoteProperties.NumberingStart!.Val!.Value);
            Assert.Equal(W.RestartNumberValues.Continuous, section.EndnoteProperties.NumberingRestart!.Val!.Value);
            Assert.Equal(W.EndnotePositionValues.SectionEnd, section.EndnoteProperties.EndnotePosition!.Val!.Value);
        }
    }

    private static void AssertOpenOfficeWord69NoteParagraphFormat(WordDocument document, W.Paragraph paragraph, JsonElement expected) {
        W.Styles? styles = document.OpenXmlDocument.MainDocumentPart!.StyleDefinitionsPart?.Styles;
        var sources = new List<OpenXmlCompositeElement>();
        if (paragraph.ParagraphProperties != null) sources.Add(paragraph.ParagraphProperties);
        sources.AddRange(WordRunFormatting.EnumerateStyleChain(styles, paragraph.ParagraphProperties?.ParagraphStyleId?.Val?.Value, W.StyleValues.Paragraph)
            .Select(style => style.StyleParagraphProperties).OfType<OpenXmlCompositeElement>());
        if (styles?.DocDefaults?.ParagraphPropertiesDefault?.ParagraphPropertiesBaseStyle is { } defaults) sources.Add(defaults);
        var indents = sources.Select(item => item.GetFirstChild<W.Indentation>()).OfType<W.Indentation>().ToArray();
        int Twips(string? value) => value == null ? 0 : int.Parse(value, System.Globalization.CultureInfo.InvariantCulture);
        int ExpectedTwips(string name) => (int)Math.Round(expected.GetProperty(name).GetInt32() * 1440D / 2540D, MidpointRounding.AwayFromZero);
        Assert.Equal(ExpectedTwips("leftMarginHundredthsMm"), Twips(indents.Select(item => item.Left?.Value).FirstOrDefault(value => value != null)));
        Assert.Equal(ExpectedTwips("rightMarginHundredthsMm"), Twips(indents.Select(item => item.Right?.Value).FirstOrDefault(value => value != null)));
        W.Indentation? firstLine = indents.FirstOrDefault(item => item.FirstLine != null || item.Hanging != null);
        Assert.Equal(ExpectedTwips("firstLineIndentHundredthsMm"), firstLine?.FirstLine != null ? Twips(firstLine.FirstLine.Value) : -Twips(firstLine?.Hanging?.Value));
        var spacings = sources.Select(item => item.GetFirstChild<W.SpacingBetweenLines>()).OfType<W.SpacingBetweenLines>().ToArray();
        Assert.Equal(ExpectedTwips("topMarginHundredthsMm"), Twips(spacings.Select(item => item.Before?.Value).FirstOrDefault(value => value != null)));
        Assert.Equal(ExpectedTwips("bottomMarginHundredthsMm"), Twips(spacings.Select(item => item.After?.Value).FirstOrDefault(value => value != null)));
        Assert.Equal(W.LineSpacingRuleValues.Auto, spacings.Select(item => item.LineRule?.Value).FirstOrDefault(value => value != null) ?? W.LineSpacingRuleValues.Auto);
        // OOXML's absent auto line value means one line (240ths), with zero paragraph spacing.
        Assert.Equal(expected.GetProperty("lineSpacing").GetProperty("Height").GetInt32() * 240 / 100,
            Twips(spacings.Select(item => item.Line?.Value).FirstOrDefault(value => value != null) ?? "240"));
        Assert.Equal(W.JustificationValues.Left, sources.Select(item => item.GetFirstChild<W.Justification>()?.Val?.Value)
            .FirstOrDefault(value => value != null) ?? W.JustificationValues.Left);
    }

    private static string ReadWord69NoteParagraphText(W.Paragraph paragraph) => string.Concat(paragraph.Descendants()
        .Select(element => element switch { W.Text text => text.Text, W.TabChar => "\t", _ => "" }));
}
