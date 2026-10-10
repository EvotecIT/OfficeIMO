using OfficeIMO.IWork;
using OfficeIMO.Reader;
using OfficeIMO.Reader.IWork;
using DocumentFormat.OpenXml.Packaging;
using W = DocumentFormat.OpenXml.Wordprocessing;

namespace OfficeIMO.IWork.Tests;

public sealed partial class IWorkEverydayNativeExportTests {
    [Fact]
    public void Native_decimal_levels_and_continuation_preserve_source_counters_and_DOCX_measurements() {
        VerifyManifest("pages-list-continuation-v15.4.json");
        using var result = WordIWorkConverter.ConvertPagesToWordResult(
            Corpus("native-exports/pages-list-continuation-v15.4.pages"),
            conversionOptions: new() { Mode = IWorkConversionMode.EditableOnly });
        result.Report.RequireCompleteEditableReconstruction();
        Assert.True(result.Projection.Body.IsFormattingComplete);
        Assert.Empty(result.Projection.SourceDeclarationIssues);
        var source = result.Projection.Body.Paragraphs;
        Assert.Equal(new int?[] { null, 7, 8, null, 9, 1, 2, 10, 1, 5, 6, null, 7, 11, 3, 4 },
            source.Select(paragraph => paragraph.ListNumber));
        Assert.Equal(new int?[] { null, 7, null, null, null, null, null, null, null, 5, null, null, null, 11, 3, null },
            source.Select(paragraph => paragraph.ListStartNumber));
        Assert.Equal(new[] { -1, 0, 0, -1, 0, 1, 1, 0, 1, 1, 1, -1, 1, 0, 0, 0 },
            source.Select(paragraph => paragraph.ListLevel));
        using var saved = new MemoryStream(); result.Value.Save(saved); saved.Position = 0;
        using var actual = WordprocessingDocument.Open(saved, false);
        using var native = WordprocessingDocument.Open(Corpus("native-exports/pages-list-continuation-v15.4.docx"), false);
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
            Assert.All(paragraph.Elements<W.Run>().Where(run => run.InnerText.Length > 0),
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
