using System.Text.Json;
using OfficeIMO.Excel;
using Xunit;

namespace OfficeIMO.Tests;

public sealed class ExcelSheetSelectorTests {
    // Expected results qualified against PowerShell WildcardPattern with IgnoreCase.
    [Theory]
    [InlineData("Data?", "Data12", false)]
    [InlineData("DATA?", "Data2", true)]
    [InlineData("Data[!1]", "Data2", false)]
    [InlineData("Data[!1]", "Data1", true)]
    [InlineData("Data[^1]", "Data2", false)]
    [InlineData("Data[^1]", "Data1", true)]
    [InlineData("Da`ta?", "Data2", true)]
    [InlineData("Tick`Name", "Tick`Name", true)]
    [InlineData("Tick`Name*", "TickName1", true)]
    [InlineData("Data[-1]", "Data1", true)]
    [InlineData("Data[1-]", "Data1", true)]
    [InlineData("Data`*", "Data1", false)]
    public void NativeNameMatchingPreservesPowerShellWildcardSemantics(string pattern, string name, bool matches) {
        var selectors = new[] { OfficeSheetSelector.From(pattern) };
        if (matches) Assert.Equal(new[] { name }, OfficeSheetSelector.ResolveNames(new[] { name }, selectors));
        else Assert.Throws<ArgumentException>(() => OfficeSheetSelector.ResolveNames(new[] { name }, selectors));
    }

    [Fact]
    public void LoadedWorkbookAndReaderNamesUseTheSameOrderedSelection() {
        using var document = ExcelDocument.Create();
        document.AddSheet("2");
        var firstData = document.AddSheet("Data1");
        var secondData = document.AddSheet("Data2");
        document.AddSheet("Other");
        Assert.Equal("2", OfficeSheetSelector.From("2").Resolve(document.Sheets, sheet => sheet.Name).Name);
        Assert.Equal("Data1", OfficeSheetSelector.From(2).Resolve(document.Sheets, sheet => sheet.Name).Name);
        OfficeSheetSelector[] selectors = {
            OfficeSheetSelector.From("data[1-3]"), OfficeSheetSelector.From(1), OfficeSheetSelector.From("Data?")
        };
        string[] expected = { "2", "Data1", "Data2" };
        Assert.Equal(expected, OfficeSheetSelector.ResolveMany(document, selectors).Select(sheet => sheet.Name));
        using var reader = ExcelDocument.OpenReader(new MemoryStream(document.ToBytes()));
        string[] names = reader.SheetNames.ToArray();
        Assert.Equal(expected, OfficeSheetSelector.ResolveNames(names, selectors));
        Assert.Equal(names, OfficeSheetSelector.ResolveNames(names, null));
        Assert.Equal(new[] { "Data2" }, OfficeSheetSelector.ResolveMany(document, null, secondData).Select(sheet => sheet.Name));
        Assert.Throws<ArgumentException>(() => OfficeSheetSelector.ResolveNames(names, Array.Empty<OfficeSheetSelector>()));
        Assert.Throws<ArgumentException>(() => OfficeSheetSelector.ResolveNames(names, new[] { OfficeSheetSelector.From("Missing*") }));
        Assert.Throws<ArgumentException>(() => OfficeSheetSelector.ResolveNames(names, new[] { OfficeSheetSelector.From("Data[") }));
        Assert.ThrowsAny<ArgumentException>(() => OfficeSheetSelector.ResolveNames(names, new[] { OfficeSheetSelector.From("Data[]") }));
        Assert.Throws<ArgumentException>(() => OfficeSheetSelector.ResolveNames(names, new[] { OfficeSheetSelector.From(firstData) }));
        Assert.Throws<ArgumentOutOfRangeException>(() => OfficeSheetSelector.From(0));
        Assert.Throws<ArgumentException>(() => new OfficeSheetSelector(name: "Data1", position: 2));
    }

    [Fact]
    public void LiveIdentitySurvivesWrapperRefreshButRejectsForeignAndReplacementSheets() {
        using var document = ExcelDocument.Create();
        var original = document.AddSheet("Same");
        var selector = OfficeSheetSelector.From(original);
        document.AddSheet("Refresh");
        Assert.Equal("Same", OfficeSheetSelector.ResolveMany(document, new[] { selector }).Single().Name);
        using var foreign = ExcelDocument.Create();
        foreign.AddSheet("Same");
        Assert.Throws<ArgumentException>(() => OfficeSheetSelector.ResolveMany(foreign, new[] { selector }));
        document.RemoveWorksheet("Same");
        document.AddSheet("Same");
        Assert.Throws<ArgumentException>(() => OfficeSheetSelector.ResolveMany(document, new[] { selector }));
    }

    [Fact]
    public void SavedSelectorsRoundTripWithoutSerializingLiveWorkbookState() {
        using var document = ExcelDocument.Create();
        var sheet = document.AddSheet("2");
        document.AddSheet("Data2");
        OfficeSheetSelector[] source = { OfficeSheetSelector.From(sheet), OfficeSheetSelector.From(2), OfficeSheetSelector.From("Data*") };
        OfficeSheetSelector[] clone = source.ToArray();
        source[0] = OfficeSheetSelector.From("Other");
        string json = JsonSerializer.Serialize(clone);
        Assert.DoesNotContain("Sheet", json);
        OfficeSheetSelector[] restored = JsonSerializer.Deserialize<OfficeSheetSelector[]>(json)!;
        Assert.Null(restored[0].Sheet);
        Assert.Equal("2", restored[0].Name);
        Assert.Null(restored[0].Position);
        Assert.Equal(2, restored[1].Position);
        Assert.Equal("Data*", restored[2].Name);
        Assert.Equal(new[] { "2", "Data2" }, OfficeSheetSelector.ResolveMany(document, restored).Select(item => item.Name));
    }
}
