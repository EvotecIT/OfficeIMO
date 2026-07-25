using System.Reflection;
using OfficeIMO.Excel;
using Xunit;

namespace OfficeIMO.Tests;

public class ExcelAllSeverityBatch21Tests {
    [Fact]
    public void EscapedAmpersandBeforeDigitsRemainsLiteralHeaderText() {
        using ExcelDocument document = ExcelDocument.Create(new MemoryStream());
        ExcelSheet sheet = document.AddWorksheet("Report");
        MethodInfo parser = typeof(ExcelSheet).GetMethod(
            "TryResolveHeaderFooterText",
            BindingFlags.Instance | BindingFlags.NonPublic)!;
        object?[] arguments = {
            "Budget &&14 Report",
            1,
            1,
            new DateTime(1970, 1, 1, 0, 0, 0, DateTimeKind.Utc),
            false,
            null
        };

        Assert.True((bool)parser.Invoke(sheet, arguments)!);
        Assert.NotNull(arguments[5]);
        object section = arguments[5]!;
        Assert.Equal(parser.GetParameters()[5].ParameterType.GetElementType(), section.GetType());
        string text = (string)section.GetType()
            .GetProperty("Text", BindingFlags.Instance | BindingFlags.NonPublic | BindingFlags.Public)!
            .GetValue(section)!;
        Assert.Equal("Budget &14 Report", text);
    }

    [Fact]
    public void DirectHeaderStylingHonorsCustomHeaderConversion() {
        using ExcelDocument document = ExcelDocument.Create(new MemoryStream());
        ExcelSheet sheet = document.AddWorksheet("Data");
        sheet.InsertObjects(
            new[] { new HeaderRow("Alpha") },
            ("RawHeader", row => row.Value));
        ExcelReadOptions options = new() {
            CellValueConverter = _ => new ExcelCellValue("ConvertedHeader")
        };

        bool found = sheet.TryGetColumnStyleByHeader(
            "ConvertedHeader",
            includeHeader: false,
            out _,
            out int columnIndex,
            options);

        Assert.True(found);
        Assert.Equal(1, columnIndex);
    }

    [Fact]
    public void PendingPreflightSurvivesDirectCellValuePromotion() {
        using ExcelDocument document = ExcelDocument.Create(new MemoryStream());
        ExcelSheet sheet = document.AddWorksheet("Data");
        sheet.CellValues(Enumerable.Range(1, 160)
            .Select(row => (row, 1, (object)("value-" + row))));
        typeof(ExcelDocument)
            .GetMethod("MarkRequiresSavePreflight", BindingFlags.Instance | BindingFlags.NonPublic)!
            .Invoke(document, null);

        using MemoryStream output = new();
        document.Save(output);

        Assert.NotEqual(ExcelSavePackageWriter.DirectDataSetPackage, document.LastSaveDiagnostics.Writer);
    }

    private sealed record HeaderRow(string Value);
}
