using System;
using System.IO;
using System.Linq;
using OfficeIMO.Excel;
using OfficeIMO.Excel.OpenDocument;
using OfficeIMO.OpenDocument;
using Xunit;

namespace OfficeIMO.OpenDocument.Converters.Tests;

public sealed class SpreadsheetSheetNameMappingTests {
    [Fact]
    public void SanitizedAndCollidingSheetNamesAreMappedInFormulasRangesAndLinks() {
        OdsDocument source = OdsDocument.Create();
        OdsSheet first = source.AddSheet("Source:2026");
        OdsSheet second = source.AddSheet("Source_2026");
        OdsSheet consumer = source.AddSheet("Summary");
        first.Cell(0, 0).SetNumber(10);
        second.Cell(0, 0).SetNumber(20);
        consumer.Cell(0, 0).Formula = "of:=[$'Source:2026'.$A$1]+[$'Source_2026'.$A$1]";
        consumer.Cell(0, 1).Formula = "of:=\"Source:2026\"";
        consumer.Cell(1, 0).SetHyperlink("Link", "#$'Source:2026'.A1");
        source.AddNamedRange("Original", "$'Source:2026'.$A$1");
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
        Assert.DoesNotContain("Source:2026", formula);
        Assert.Contains("Source:2026", cells.Single(cell => cell.Row == 1 && cell.Column == 2).Formula);
        Assert.Contains(snapshot.NamedRanges, name => name.Name == "Original" && name.ReferenceA1.Contains("'" + firstName + "'!"));
        Assert.Contains(target.GetHyperlinks().Values, link => link.Target?.Contains("'" + firstName + "'!") == true);
    }
}
