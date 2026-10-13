using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Validation;
using OfficeIMO.Word;
using W = DocumentFormat.OpenXml.Wordprocessing;

namespace OfficeIMO.IWork.Tests;

public sealed class IWorkPagesMixedFontListTests {
    [Fact]
    public void Independent_marker_fonts_distinguish_native_document_identity() {
        var identities = new List<string>();
        foreach (string markerFont in new[] { "Times New Roman", "Arial", "Courier New" }) {
            var document = IWorkPagesDocument.Create();
            document.CreateNumberedList(markerFont, 12).AddRichItem(new[] {
                new IWorkTextRunData("Same body", "Times New Roman", 12, bold: true)
            });
            byte[] bytes = document.SaveBytes(); Assert.Equal(bytes, document.SaveBytes());
            Assert.Equal(markerFont, IWorkSourceDocument.Open(bytes).ReadPages().Body.Paragraphs[0].ListFontName);
            using var package = new System.IO.Compression.ZipArchive(new MemoryStream(bytes));
            using var identifier = new StreamReader(package.GetEntry("Metadata/DocumentIdentifier")!.Open());
            identities.Add(identifier.ReadToEnd());
        }
        Assert.Equal(3, identities.Distinct(StringComparer.Ordinal).Count());
    }

    [Fact]
    public void Mixed_character_families_keep_owned_markers_counters_and_saved_Word_fonts() {
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
            var model = authored[index];
            Assert.Equal(model.NumberedList!.FontName, items[index].ListFontName);
            Assert.Equal(model.NumberedList.FontName, items[index].Style.TextStyle.FontName);
            Assert.Equal(12, items[index].Style.TextStyle.FontSizePoints);
            Assert.Equal("000000", items[index].Style.TextStyle.Color!.RgbHex);
            Assert.False(items[index].Style.TextStyle.Bold);
            Assert.False(items[index].Style.TextStyle.Italic);
            Assert.Equal(model.Runs.Select(run => run.Text), items[index].Runs.Select(run => run.Text));
            Assert.Equal(model.Runs.Select(run => run.FontName), items[index].Runs.Select(run => run.Style.FontName));
            Assert.Equal(model.Runs.Select(run => run.Color.RgbHex), items[index].Runs.Select(run => run.Style.Color!.RgbHex));
        }
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
        var numbered = saved.MainDocumentPart!.Document!.Body!.Elements<W.Paragraph>()
            .Where(item => item.ParagraphProperties?.NumberingProperties != null).ToArray();
        Assert.Equal(items.Length, numbered.Length);
        for (int index = 0; index < numbered.Length; index++) {
            var runs = numbered[index].Elements<W.Run>().Where(run => run.InnerText.Length > 0).ToArray();
            Assert.Equal(authored[index].Runs.Select(run => run.FontName),
                runs.Select(run => run.RunProperties!.RunFonts!.Ascii!.Value));
            Assert.Equal(authored[index].Runs.Select(run => run.Color.RgbHex),
                runs.Select(run => run.RunProperties!.Color!.Val!.Value));
        }
    }

    internal static IWorkPagesDocument Example() {
        var document = IWorkPagesDocument.Create();
        document.AddParagraph("Mixed family decimal items");
        var primary = document.CreateNumberedList("Arial", 12);
        primary.AddRichItem(Runs("Root seven"), restartAt: 7);
        for (int level = 1; level <= 8; level++)
            primary.AddRichItem(Runs("Depth " + level), level: level);
        primary.AddRichItem(Runs("Deep two"), level: 8);
        document.AddParagraph("Plain gap retains counters");
        primary.AddRichItem(Runs("Root eight"));
        primary.AddRichItem(Runs("Child one"), level: 1);
        var serif = document.CreateNumberedList(IWorkDecimalListMarkerStyle.Parentheses, "Times New Roman", 12);
        serif.AddRichItem(Runs("Serif marker"), restartAt: 3);
        serif.AddRichItem(Runs("Serif child"), level: 1);
        var mono = document.CreateNumberedList(IWorkDecimalListMarkerStyle.ClosingParenthesis, "Courier New", 12);
        mono.AddRichItem(Runs("Mono marker"), restartAt: 5);
        mono.AddRichItem(Runs("Mono child"), level: 1);
        primary.AddRichItem(Runs("Resumed child"), level: 1);
        return document;
    }

    private static IWorkTextRunData[] Runs(string text) => new[] {
        new IWorkTextRunData(text + " serif", "Times New Roman", 12, "003366", bold: true),
        new IWorkTextRunData(" mono", "Courier New", 12, "008000", italic: true),
        new IWorkTextRunData(" plain", "Arial", 12)
    };
}
