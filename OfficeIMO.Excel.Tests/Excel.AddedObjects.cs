using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Validation;
using OfficeIMO.Excel;
using Xunit;

namespace OfficeIMO.Tests;

public partial class Excel {
    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public void AddedPageBreak_ReturnsAttachedObjectAndRejectsAReplacement(bool rowBreak) {
        using var document = ExcelDocument.Create();
        var sheet = document.AddSheet("Data");
        ExcelPageBreak pageBreak = rowBreak
            ? sheet.AddManualRowPageBreak(4) : sheet.AddManualColumnPageBreak(4);
        Assert.Same(sheet, pageBreak.Sheet);
        Assert.Equal(rowBreak, pageBreak.IsRowBreak);
        Assert.Equal(4, pageBreak.Index);
        using var output = new MemoryStream();
        document.Save(output);
        using (var package = SpreadsheetDocument.Open(new MemoryStream(output.ToArray()), false))
            Assert.Empty(new OpenXmlValidator().Validate(package));
        using (var loaded = ExcelDocument.Load(new MemoryStream(output.ToArray()))) {
            var rows = rowBreak ? loaded.Sheets[0].GetManualRowPageBreaks() : loaded.Sheets[0].GetManualColumnPageBreaks();
            Assert.Equal(new[] { 4 }, rows);
        }
        pageBreak.Remove();
        ExcelPageBreak replacement = rowBreak
            ? sheet.AddManualRowPageBreak(4) : sheet.AddManualColumnPageBreak(4);
        Assert.Throws<InvalidOperationException>(() => pageBreak.Remove());
        Assert.Throws<InvalidOperationException>(() => pageBreak.Index);
        Assert.Equal(4, replacement.Index);
        replacement.Remove();
        Assert.Empty(rowBreak ? sheet.GetManualRowPageBreaks() : sheet.GetManualColumnPageBreaks());
    }

    [Fact]
    public void AddedMacro_ReturnsTheAttachedOpaqueProjectAndRejectsReplacement() {
        using var document = ExcelDocument.Create();
        document.AddSheet("Data");
        ExcelMacroProject project = document.AddMacro(new byte[] { 1, 2, 3 });
        Assert.Same(document, project.Document);
        Assert.EndsWith("vbaProject.bin", project.PartUri, StringComparison.Ordinal);
        Assert.Equal(new byte[] { 1, 2, 3 }, project.ToBytes());
        using var source = new MemoryStream(new byte[] { 4, 5 });
        ExcelMacroProject replacement = document.AddMacro(source);
        Assert.True(source.CanRead);
        Assert.Throws<InvalidOperationException>(() => project.ToBytes());
        Assert.Throws<InvalidOperationException>(() => project.Remove());
        Assert.Equal(new byte[] { 4, 5 }, replacement.ToBytes());
        replacement.Remove();
        Assert.Empty(document.ExtractMacros());
        Assert.Throws<InvalidOperationException>(() => replacement.ToBytes());
    }

    [Fact]
    public void AddedIgnoredErrorRegionAndBackLinks_ReturnTheirCreatedMetadata() {
        using var document = ExcelDocument.Create();
        var first = document.AddSheet("First");
        document.AddSheet("Second");
        ExcelIgnoredErrorRegionInfo region = first.AddIgnoredErrorRegion(new[] { "a1:b4" }, ExcelIgnoredErrorKind.NumberStoredAsText);
        Assert.Equal(new[] { "A1:B4" }, region.Ranges);
        Assert.Equal(ExcelIgnoredErrorKind.NumberStoredAsText, region.Errors);
        Assert.Equal(region.Ranges, Assert.Single(first.GetIgnoredErrorRegions()).Ranges);
        var toc = document.AddTableOfContents(styled: false);
        var links = document.AddBackLinksToToc(toc.Name);
        Assert.Equal(new[] { "First", "Second" }, links.Keys.Select(sheet => sheet.Name).OrderBy(name => name));
        foreach (var link in links) {
            Assert.False(link.Value.IsExternal);
            Assert.Equal(link.Key.GetHyperlinks()["A2"].Target, link.Value.Target);
            Assert.Contains(toc.Name, link.Value.Target, StringComparison.Ordinal);
        }
        using var loaded = ExcelDocument.Load(new MemoryStream(document.ToBytes()));
        Assert.Equal(2, loaded.Sheets.Where(sheet => sheet.Name != toc.Name).Sum(sheet => sheet.GetHyperlinks().Count));
    }
    [Fact]
    public void CustomProperties_ReturnAttachedValuesAndKeepDictionaryContracts() {
        using var document = ExcelDocument.Create();
        document.AddSheet("Data");
        var properties = document.CustomDocumentProperties;
        var supplied = new ExcelCustomProperty { Value = "Draft", PropertyType = ExcelCustomPropertyType.Text };
        var added = properties.Add(" Workflow ", supplied);
        Assert.Same(supplied, added);
        Assert.Same(added, properties["Workflow"]);
        var pair = new KeyValuePair<string, ExcelCustomProperty>("Reviewer",
            new ExcelCustomProperty { Value = "First", PropertyType = ExcelCustomPropertyType.Text });
        Assert.Same(pair.Value, properties.Add(pair));
        ((IDictionary<string, ExcelCustomProperty>)properties).Add("Dictionary",
            new ExcelCustomProperty { Value = "Present", PropertyType = ExcelCustomPropertyType.Text });
        ((ICollection<KeyValuePair<string, ExcelCustomProperty>>)properties).Add(
            new KeyValuePair<string, ExcelCustomProperty>("Collection",
                new ExcelCustomProperty { Value = "Present", PropertyType = ExcelCustomPropertyType.Text }));
        added.Value = "Reviewed";
        pair.Value.Value = "Second";
        using var loaded = ExcelDocument.Load(new MemoryStream(document.ToBytes()));
        Assert.Equal("Reviewed", loaded.CustomDocumentProperties["Workflow"].Text);
        Assert.Equal("Second", loaded.CustomDocumentProperties["Reviewer"].Text);
        Assert.Equal("Present", loaded.CustomDocumentProperties["Dictionary"].Text);
        Assert.Equal("Present", loaded.CustomDocumentProperties["Collection"].Text);
    }

}
