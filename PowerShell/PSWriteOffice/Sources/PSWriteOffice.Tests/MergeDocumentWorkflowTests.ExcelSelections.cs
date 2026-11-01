using System.Management.Automation;
using OfficeIMO;
using OfficeIMO.Excel;

namespace PSWriteOffice.Tests;

public sealed partial class MergeDocumentWorkflowTests {
    [Theory]
    [InlineData("wildcard", false)]
    [InlineData("position", true)]
    [InlineData("object", true)]
    [InlineData("none", true)]
    public void ExcelMergeFactorySelectsAppendedSheetsAndNeverTreatsNoMatchesAsAll(string selection, bool liveInput) {
        using var files = new MergeFiles();
        string cover = files.Excel("cover.xlsx", "Cover", "COVER");
        string first = files.Excel("north.xlsx", "Ignore A", "IGNORE A");
        string second = files.Excel("south.xlsx", "Ignore B", "IGNORE B");
        foreach (var item in new[] { (Path: first, Name: "Sales A", Marker: "NORTH"), (Path: second, Name: "Sales B", Marker: "SOUTH") }) {
            using var source = ExcelDocument.Load(item.Path, new ExcelLoadOptions { PersistenceMode = DocumentPersistenceMode.Explicit });
            var sales = source.AddSheet(item.Name); sales.CellValue(1, 1, item.Marker);
            sales.Cell(2, 1, 1.2345, numberFormat: "0.0000");
            source.AddSheet("Secret").CellValue(1, 1, "MUST NOT IMPORT");
            source.Save();
        }
        var before = files.Snapshot();
        using var selectedOwner = ExcelDocument.Load(first, new ExcelLoadOptions {
            AccessMode = DocumentAccessMode.ReadOnly, PersistenceMode = DocumentPersistenceMode.Explicit
        });
        using var target = ExcelDocument.Create(); target.AddSheet("Cover").CellValue(1, 1, "COVER");
        object selector = selection switch {
            "position" => 2,
            "object" => selectedOwner.Sheets[1],
            "none" => "Absent*",
            _ => "Sales *"
        };
        string output = files.PathFor("merged.xlsx");
        using var runspace = OpenRunspace();
        Set(runspace, ("selector", selector), ("target", target), ("inputs", new[] { first, second }),
            ("allInputs", new[] { cover, first, second }), ("output", output), ("liveInput", liveInput));
        var result = Invoke(runspace, """
            $options = New-OfficeExcelMergeOptions -Sheet $selector -NamePrefix 'Imported ' -CopyMode Package
            if ($liveInput) {
                $target | Merge-OfficeDocument -Path $inputs -Options $options -PassThru -ErrorAction Stop
            } else {
                Merge-OfficeDocument -LiteralPath $allInputs -OutputPath $output -Options $options -PassThru -ErrorAction Stop
            }
            """);
        ExcelDocument? reopened = null;
        try {
            if (liveInput) Assert.Same(target, Assert.Single(result).BaseObject);
            else {
                Assert.Equal(output, Assert.IsType<FileInfo>(Assert.Single(result).BaseObject).FullName);
                reopened = ExcelDocument.Load(output);
            }
            ExcelDocument merged = reopened ?? target;
            string[] expected = selection switch {
                "object" => new[] { "Cover", "Imported Sales A" },
                "none" => new[] { "Cover" },
                _ => new[] { "Cover", "Imported Sales A", "Imported Sales B" }
            };
            Assert.Equal(expected, merged.Sheets.Select(sheet => sheet.Name));
            foreach (var sheet in merged.Sheets.Skip(1)) {
                Assert.True(sheet.TryGetCellText(1, 1, out string text));
                Assert.Equal(sheet.Name.EndsWith("A", StringComparison.Ordinal) ? "NORTH" : "SOUTH", text);
                Assert.Equal("0.0000", sheet.CellAt(2, 1).GetStyle().NumberFormatCode);
            }
        } finally { reopened?.Dispose(); }
        Assert.Equal(3, selectedOwner.Sheets.Count);
        Assert.Equal("Sales A", selectedOwner.Sheets[1].Name);
        target.AddSheet("Still open");
        files.AssertSourcesUnchanged(before);
    }

    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public void ExcelMergeResolvesAllTypedSelectionsBeforeChangingALiveTarget(bool foreignObject) {
        using var files = new MergeFiles();
        string first = files.Excel("first.xlsx", "First", "FIRST");
        string second = files.Excel("second.xlsx", "Shared", "SECOND");
        string unrelated = files.Excel("unrelated.xlsx", "Shared", "UNRELATED");
        using (var source = ExcelDocument.Load(first, new ExcelLoadOptions { PersistenceMode = DocumentPersistenceMode.Explicit })) {
            source.AddSheet("Shared").CellValue(1, 1, "FIRST SHARED"); source.Save();
        }
        var before = files.Snapshot();
        using var foreign = ExcelDocument.Load(unrelated, new ExcelLoadOptions {
            AccessMode = DocumentAccessMode.ReadOnly, PersistenceMode = DocumentPersistenceMode.Explicit
        });
        using var target = ExcelDocument.Create(); target.AddSheet("Cover").CellValue(1, 1, "COVER");
        string output = files.PathFor("absent/merged.xlsx");
        using var runspace = OpenRunspace();
        Set(runspace, ("selector", foreignObject ? (object)foreign.Sheets[0] : 2), ("target", target),
            ("inputs", new[] { first, second }), ("output", output));
        string error = Fail(runspace, """
            $options = New-OfficeExcelMergeOptions -Sheet $selector -CopyMode Package
            $target | Merge-OfficeDocument -Path $inputs -OutputPath $output -Options $options -ErrorAction Stop
            """);
        Assert.Contains("sheet", error, StringComparison.OrdinalIgnoreCase);
        Assert.Equal(new[] { "Cover" }, target.Sheets.Select(sheet => sheet.Name));
        Assert.False(Directory.Exists(Path.GetDirectoryName(output)));
        Assert.Equal("Shared", foreign.Sheets[0].Name);
        files.AssertSourcesUnchanged(before);
    }
}
