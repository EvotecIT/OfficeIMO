using System.Security.Cryptography;
using System.Text.Json;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using OfficeIMO.Excel;
using OfficeIMO.IWork;
using OfficeIMO.Word;

namespace OfficeIMO.IWork.Tests;

public sealed class IWorkEverydayCorpusTests {
    private static string Corpus(string path) => Path.Combine(AppContext.BaseDirectory, "Documents", "IWorkCorpus", path);

    [Fact]
    public void Simple_native_numbers_converts_without_partial_policy_and_matches_Apple_export_values_and_alignment() {
        using var evidence = JsonDocument.Parse(File.ReadAllText(Corpus("native-exports/numbers-simple-v14.5.json")));
        var manifest = evidence.RootElement;
        string path = Corpus(manifest.GetProperty("source").GetString()!);
        string applePath = Corpus(manifest.GetProperty("export").GetString()!);
        Assert.Equal(manifest.GetProperty("sourceSha256").GetString(), Hash(path));
        Assert.Equal(manifest.GetProperty("exportSha256").GetString(), Hash(applePath));
        using var result = ExcelIWorkConverter.ConvertNumbersToExcelResult(path);
        result.Report.RequireCompleteEditableReconstruction();
        Assert.Empty(result.Report.SourceDeclarationIssues);
        IWorkTable table = Assert.Single(Assert.Single(result.Projection.Sheets).Tables);
        Assert.Equal(1, table.TextStyles.Body!.LineSpacingMultiplier);
        Assert.Empty(table.TextStyles.Body.TabStops!);
        using var saved = new MemoryStream();
        result.Value.Save(saved); saved.Position = 0;
        using var actual = ExcelDocument.Load(saved);
        using var apple = ExcelDocument.Load(applePath);
        ExcelSheet destination = Assert.Single(actual.Sheets);
        ExcelSheet reference = Assert.Single(apple.Sheets);
        for (int row = 1; row <= 3; row++) {
            for (int column = 1; column <= 3; column++) {
                Assert.True(destination.TryGetCellValueSnapshot(row, column, out var cell));
                Assert.True(reference.TryGetCellValueSnapshot(row + 1, column, out var expected));
                Assert.Equal(expected!.Kind, cell!.Kind);
                if (expected.Kind == ExcelCellValueKind.Number) Assert.Equal(expected.RawValue, cell.RawValue);
                else Assert.Equal(reference.CellAt(row + 1, column).GetValue<string>(), destination.CellAt(row, column).GetValue<string>());
                var style = destination.GetCellStyle(row, column);
                var nativeStyle = reference.GetCellStyle(row + 1, column);
                Assert.Equal(nativeStyle.Bold, style.Bold);
                Assert.Equal(nativeStyle.FontSize, style.FontSize);
                Assert.True(string.IsNullOrEmpty(nativeStyle.HorizontalAlignment)
                    || nativeStyle.HorizontalAlignment == "general");
                Assert.Equal("general", style.HorizontalAlignment);
            }
        }
        Assert.Empty(actual.ValidateOpenXml());
        using var convenience = ExcelIWorkConverter.ConvertNumbersToExcel(path);
        Assert.Single(convenience.Sheets);
    }

    [Fact]
    public void Simple_native_pages_matches_Apple_body_spacing_paragraph_mark_and_page_layout() {
        using var evidence = JsonDocument.Parse(File.ReadAllText(Corpus("native-exports/pages-simple-v15.4.json")));
        var manifest = evidence.RootElement;
        string path = Corpus(manifest.GetProperty("source").GetString()!);
        string applePath = Corpus(manifest.GetProperty("export").GetString()!);
        Assert.Equal(manifest.GetProperty("sourceSha256").GetString(), Hash(path));
        Assert.Equal(manifest.GetProperty("exportSha256").GetString(), Hash(applePath));
        using var apple = WordprocessingDocument.Open(applePath, false);
        var reference = apple.MainDocumentPart!.Document!.Body!;
        var expected = reference.Elements<Paragraph>().ToArray();
        Assert.Equal(3, expected.Length);
        using var result = WordIWorkConverter.ConvertPagesToWordResult(path);
        result.Report.RequireCompleteEditableReconstruction();
        Assert.Empty(result.Report.SourceDeclarationIssues);
        Assert.Equal(expected.Select(paragraph => paragraph.InnerText),
            result.Projection.Body.Paragraphs.Select(paragraph => paragraph.Text));
        using var saved = new MemoryStream();
        result.Value.Save(saved); saved.Position = 0;
        using var document = WordDocument.Load(saved);
        Assert.Empty(document.ValidateDocument());
        saved.Position = 0;
        using var package = WordprocessingDocument.Open(saved, false);
        var body = package.MainDocumentPart!.Document!.Body!;
        var actual = body.Elements<Paragraph>().ToArray();
        Assert.Equal(expected.Select(paragraph => paragraph.InnerText), actual.Select(paragraph => paragraph.InnerText));
        for (int index = 0; index < actual.Length; index++) {
            var nativeSpacing = expected[index].ParagraphProperties!.SpacingBetweenLines!;
            var spacing = actual[index].ParagraphProperties!.SpacingBetweenLines!;
            Assert.Equal(nativeSpacing.Line!.Value, spacing.Line!.Value);
            Assert.Equal(nativeSpacing.LineRule!.Value, spacing.LineRule!.Value);
        }
        var nativeStyle = apple.MainDocumentPart.StyleDefinitionsPart!.Styles!.Elements<Style>()
            .Single(style => style.StyleId?.Value == "Default");
        var emptyParagraph = Assert.Single(actual, paragraph => paragraph.InnerText.Length == 0);
        var mark = emptyParagraph.ParagraphProperties!.ParagraphMarkRunProperties;
        Assert.NotNull(mark);
        Assert.Equal(nativeStyle.StyleRunProperties!.FontSize!.Val!.Value, mark.GetFirstChild<FontSize>()!.Val!.Value);
        Assert.Equal(manifest.GetProperty("qualifiedSourceFields").GetProperty("paragraphTextFont").GetString(),
            mark.GetFirstChild<RunFonts>()!.Ascii!.Value);
        var nativeSection = reference.Elements<SectionProperties>().Single();
        var section = body.Elements<SectionProperties>().Single();
        var nativePageSize = nativeSection.GetFirstChild<PageSize>()!;
        var pageSize = section.GetFirstChild<PageSize>()!;
        Assert.Equal(nativePageSize.Width!.Value, pageSize.Width!.Value);
        Assert.Equal(nativePageSize.Height!.Value, pageSize.Height!.Value);
        const string wordNamespace = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";
        foreach (string side in new[] { "top", "right", "bottom", "left", "header", "footer" }) {
            Assert.Equal(nativeSection.GetFirstChild<PageMargin>()!.GetAttribute(side, wordNamespace).Value,
                section.GetFirstChild<PageMargin>()!.GetAttribute(side, wordNamespace).Value);
        }
        Assert.Equal(apple.MainDocumentPart.DocumentSettingsPart!.Settings!.GetFirstChild<DefaultTabStop>()!.Val!.Value,
            package.MainDocumentPart.DocumentSettingsPart!.Settings!.GetFirstChild<DefaultTabStop>()!.Val!.Value);
    }

    private static string Hash(string path) => Convert.ToHexString(SHA256.HashData(File.ReadAllBytes(path))).ToLowerInvariant();
}
