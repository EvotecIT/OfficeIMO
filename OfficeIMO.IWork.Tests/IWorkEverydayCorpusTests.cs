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
    public void Simple_native_pages_converts_complete_editable_body_including_empty_paragraphs() {
        using var result = WordIWorkConverter.ConvertPagesToWordResult(Corpus("nim-iwork/simple.pages"));
        result.Report.RequireCompleteEditableReconstruction();
        Assert.Empty(result.Report.SourceDeclarationIssues);
        string[] expected = { "hello pages", "", "second paragraph with some words\"" };
        Assert.Equal(expected, result.Projection.Body.Paragraphs.Select(paragraph => paragraph.Text));
        using var saved = new MemoryStream();
        result.Value.Save(saved); saved.Position = 0;
        using var document = WordDocument.Load(saved);
        Assert.Empty(document.ValidateDocument());
        saved.Position = 0;
        using var package = WordprocessingDocument.Open(saved, false);
        Assert.Equal(expected, package.MainDocumentPart!.Document.Body!.Elements<Paragraph>()
            .Select(paragraph => paragraph!.InnerText));
    }

    private static string Hash(string path) => Convert.ToHexString(SHA256.HashData(File.ReadAllBytes(path))).ToLowerInvariant();
}
