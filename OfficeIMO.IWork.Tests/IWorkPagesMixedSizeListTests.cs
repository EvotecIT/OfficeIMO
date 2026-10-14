using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Validation;
using OfficeIMO.Word;
using W = DocumentFormat.OpenXml.Wordprocessing;

namespace OfficeIMO.IWork.Tests;

public sealed class IWorkPagesMixedSizeListTests {
    [Theory]
    [InlineData("orders")]
    [InlineData("saved")]
    [InlineData("edited")]
    public void Apple_mixed_size_reference_keeps_native_marker_sizes_and_counters_in_saved_Word(string scenario) {
        var path = Path.Combine(AppContext.BaseDirectory, "Documents", "IWorkCorpus", "native-exports",
            "pages-mixed-size-" + scenario + "-v15.4.pages");
        var source = IWorkSourceDocument.Open(path);
        Assert.True(source.ReadPages().Body.IsFormattingComplete);
        CheckWord(source);
    }

    [Fact]
    public void Mixed_sizes_keep_first_run_markers_default_indents_and_counters_in_saved_Word() {
        var document = Example();
        byte[] bytes = document.SaveBytes();
        Assert.Equal(bytes, document.SaveBytes());
        var source = IWorkSourceDocument.Open(bytes);
        var projection = source.ReadPages();
        Assert.True(projection.Body.IsFormattingComplete);
        Assert.Empty(projection.Diagnostics);
        Assert.Empty(projection.SourceDeclarationIssues);
        Assert.Empty(projection.SourceReferenceIssues);
        var authored = document.Paragraphs.Where(item => item.NumberedList != null).ToArray();
        var items = projection.Body.Paragraphs.Where(item => item.ListNumber.HasValue).ToArray();
        Assert.Equal(authored.Select(item => item.Text), items.Select(item => item.Text));
        Assert.Equal(authored.Select(item => item.ListLevel), items.Select(item => item.ListLevel));
        Assert.Equal(authored.Select(item => item.ListNumber), items.Select(item => item.ListNumber));
        for (int index = 0; index < items.Length; index++) {
            Assert.Equal(12, items[index].Style.TextStyle.FontSizePoints);
            Assert.Equal(1, items[index].ListLayout!.MarkerScale);
            Assert.Equal(1.5, items[index].ListLayout!.TextIndentEm);
            Assert.Equal(authored[index].Runs.Select(run => (double)run.FontSizePoints),
                items[index].Runs.Select(run => run.Style.FontSizePoints!.Value));
        }
        CheckWord(source);
    }

    [Fact]
    public void Paragraph_default_size_distinguishes_identity_when_run_sizes_match() {
        var identities = new List<string>();
        foreach (float defaultSize in new[] { 12f, 14f }) {
            var document = IWorkPagesDocument.Create();
            document.CreateNumberedList(fontSizePoints: defaultSize).AddRichItem(Runs("Same body", 18, 8, 14));
            byte[] bytes = document.SaveBytes();
            Assert.Equal(bytes, document.SaveBytes());
            var source = IWorkSourceDocument.Open(bytes);
            Assert.Equal(defaultSize, source.ReadPages().Body.Paragraphs[0].Style.TextStyle.FontSizePoints);
            CheckWord(source);
            using var package = new System.IO.Compression.ZipArchive(new MemoryStream(bytes));
            using var identifier = new StreamReader(package.GetEntry("Metadata/DocumentIdentifier")!.Open());
            identities.Add(identifier.ReadToEnd());
        }
        Assert.Equal(2, identities.Distinct(StringComparer.Ordinal).Count());
    }

    internal static void CheckWord(IWorkSourceDocument source) {
        var items = source.ReadPages().Body.Paragraphs.Where(item => item.ListNumber.HasValue).ToArray();
        using var result = source.ToWordDocumentResult(new() { Mode = IWorkConversionMode.EditableOnly });
        result.RequireCompleteEditableReconstruction();
        using var output = new MemoryStream();
        result.Value.Save(output);
        using (var reopened = WordDocument.Load(new MemoryStream(output.ToArray())))
            Assert.Equal(items.Select(item => item.ListNumber!.Value),
                WordDocumentTraversal.BuildListIndices(reopened).Values.Select(item => item.Index));
        output.Position = 0;
        using var saved = WordprocessingDocument.Open(output, false);
        Assert.Empty(new OpenXmlValidator().Validate(saved));
        var numbering = saved.MainDocumentPart!.NumberingDefinitionsPart!.Numbering!;
        var numbered = saved.MainDocumentPart.Document!.Body!.Elements<W.Paragraph>()
            .Where(item => item.ParagraphProperties?.NumberingProperties != null).ToArray();
        Assert.Equal(items.Length, numbered.Length);
        for (int index = 0; index < items.Length; index++) {
            var item = items[index];
            var properties = numbered[index].ParagraphProperties!;
            Assert.Equal((item.Runs[0].Style.FontSizePoints!.Value * 2).ToString(System.Globalization.CultureInfo.InvariantCulture),
                properties.ParagraphMarkRunProperties!.GetFirstChild<W.FontSize>()!.Val!.Value);
            int numberId = properties.NumberingProperties!.NumberingId!.Val!.Value;
            int abstractId = numbering.Elements<W.NumberingInstance>().Single(value => value.NumberID!.Value == numberId)
                .AbstractNumId!.Val!.Value;
            var level = numbering.Elements<W.AbstractNum>().Single(value => value.AbstractNumberId!.Value == abstractId)
                .Elements<W.Level>().Single(value => value.LevelIndex!.Value == item.ListLevel);
            Assert.Equal((item.Runs[0].Style.FontSizePoints!.Value * 2).ToString(System.Globalization.CultureInfo.InvariantCulture),
                level.NumberingSymbolRunProperties!.FontSize!.Val!.Value);
            Assert.Equal(item.ListFontName, level.NumberingSymbolRunProperties.RunFonts!.Ascii!.Value);
            double textOffset = item.Style.TextStyle.FontSizePoints!.Value * item.ListLayout!.TextIndentEm;
            Assert.Equal(((item.ListLayout.MarkerIndentPoints + textOffset) * 20).ToString(System.Globalization.CultureInfo.InvariantCulture),
                properties.Indentation!.Left!.Value);
            Assert.Equal((textOffset * 20).ToString(System.Globalization.CultureInfo.InvariantCulture),
                properties.Indentation.Hanging!.Value);
            var runs = numbered[index].Elements<W.Run>().Where(run => run.InnerText.Length > 0).ToArray();
            Assert.Equal(item.Runs.Select(run => run.Style.FontName), runs.Select(run => run.RunProperties!.RunFonts!.Ascii!.Value));
            Assert.Equal(item.Runs.Select(run => (run.Style.FontSizePoints!.Value * 2).ToString(System.Globalization.CultureInfo.InvariantCulture)),
                runs.Select(run => run.RunProperties!.FontSize!.Val!.Value));
        }
    }

    internal static IWorkPagesDocument Example() {
        var document = IWorkPagesDocument.Create();
        document.AddParagraph("Mixed size decimal items");
        var primary = document.CreateNumberedList("Arial", 12);
        primary.AddRichItem(Runs("Root seven", 18, 8, 14), restartAt: 7);
        for (int level = 1; level <= 8; level++) primary.AddRichItem(Runs("Depth " + level), level: level);
        primary.AddRichItem(Runs("Deep two"), level: 8);
        document.AddParagraph("Plain gap retains counters");
        primary.AddRichItem(Runs("Root eight", 8, 18, 14));
        primary.AddRichItem(Runs("Child one", 8, 10, 14), level: 1);
        var serif = document.CreateNumberedList(IWorkDecimalListMarkerStyle.Parentheses, "Times New Roman", 12);
        serif.AddRichItem(Runs("Serif marker"), restartAt: 3);
        serif.AddRichItem(Runs("Serif child"), level: 1);
        var mono = document.CreateNumberedList(IWorkDecimalListMarkerStyle.ClosingParenthesis, "Courier New", 12);
        mono.AddRichItem(Runs("Mono marker"), restartAt: 5);
        mono.AddRichItem(Runs("Mono child"), level: 1);
        primary.AddRichItem(Runs("Resumed child", 12, 24, 8), level: 1);
        return document;
    }

    private static IWorkTextRunData[] Runs(string text, float first = 12, float middle = 12, float last = 12) => new[] {
        new IWorkTextRunData(text + " serif", "Times New Roman", first, "003366", bold: true),
        new IWorkTextRunData(" mono", "Courier New", middle, "008000", italic: true),
        new IWorkTextRunData(" plain", "Arial", last)
    };
}
