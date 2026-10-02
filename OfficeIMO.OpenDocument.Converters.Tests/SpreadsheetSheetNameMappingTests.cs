using System;
using System.IO;
using System.Linq;
using OfficeIMO.Excel;
using OfficeIMO.Excel.OpenDocument;
using OfficeIMO.OpenDocument;
using Xunit;

namespace OfficeIMO.OpenDocument.Converters.Tests;

public sealed class SpreadsheetSheetNameMappingTests {
    [Theory]
    [InlineData("Source:2026")]
    [InlineData("Source]2026")]
    [InlineData("[Book]Source")]
    public void SanitizedAndCollidingSheetNamesAreMappedInFormulasRangesAndLinks(string sourceName) {
        OdsDocument source = OdsDocument.Create();
        OdsSheet first = source.AddSheet(sourceName);
        OdsSheet second = source.AddSheet("Source_2026");
        OdsSheet consumer = source.AddSheet("Summary");
        first.Cell(0, 0).SetNumber(10);
        second.Cell(0, 0).SetNumber(20);
        consumer.Cell(0, 0).Formula = "of:=[$'" + sourceName + "'.$A$1]+[$'Source_2026'.$A$1]";
        consumer.Cell(0, 1).Formula = "of:=\"" + sourceName + "\"";
        consumer.Cell(1, 0).SetHyperlink("Link", "#$'" + sourceName + "'.A1");
        source.AddNamedRange("Original", "$'" + sourceName + "'.$A$1");
        using ExcelDocument converted = source.ToExcelDocument();
        using ExcelDocument result = ExcelDocument.Load(new MemoryStream(converted.ToBytes()));
        string firstName = result.Sheets[0].Name, secondName = result.Sheets[1].Name;
        Assert.NotEqual(firstName, secondName);
        ExcelSheet target = result.Sheets[2];
        var snapshot = result.CreateInspectionSnapshot();
        var cells = snapshot.Worksheets[2].Cells;
        string formula = cells.Single(cell => cell.Row == 1 && cell.Column == 1).Formula!;
        Assert.Contains("'" + firstName + "'!$A$1", formula);
        Assert.Contains("'" + secondName + "'!$A$1", formula);
        Assert.DoesNotContain(sourceName, formula);
        Assert.Contains(sourceName, cells.Single(cell => cell.Row == 1 && cell.Column == 2).Formula);
        Assert.Contains(snapshot.NamedRanges, name => name.Name == "Original" && name.ReferenceA1.Contains("'" + firstName + "'!"));
        Assert.Contains(target.GetHyperlinks().Values, link => link.Target?.Contains("'" + firstName + "'!") == true);
    }
}
