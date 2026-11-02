using OfficeIMO.Excel;
using Xunit;

namespace OfficeIMO.Tests;

public partial class Excel {
    [Theory]
    [InlineData("RANDBETWEEN(1.1,2.9)", 2d)]
    [InlineData("RANDBETWEEN(-2.9,-1.1)", -2d)]
    [InlineData("RANDBETWEEN(-0.9,0.9)", 0d)]
    [InlineData("RANDBETWEEN(1,1.9)", 1d)]
    [InlineData("RANDBETWEEN(-1.9,-1)", -1d)]
    [InlineData("RANDBETWEEN(0,0.9)", 0d)]
    [InlineData("RANDBETWEEN(-0.9,0)", 0d)]
    [InlineData("RANDBETWEEN(2147483647.1,2147483648.9)", 2147483648d)]
    public void RandomBetween_fractional_bounds_with_one_integer_produce_that_integer(string formula, double expected) {
        using var document = ExcelDocument.Create();
        var sheet = document.AddSheet("Bounds");
        sheet.CellValue(1, 1, 42d); sheet.CellFormula(1, 1, formula);
        Assert.Equal(1, document.Calculate());
        Assert.Equal(expected, sheet.CellAt(1, 1).GetValue<double>());
        using var saved = new MemoryStream(); document.Save(saved); saved.Position = 0;
        using var reopened = ExcelDocument.Load(saved);
        Assert.Equal(formula, reopened.Sheets[0].GetFormulaText(1, 1));
        Assert.Equal(expected, reopened.Sheets[0].CellAt(1, 1).GetValue<double>());
        Assert.Empty(reopened.ValidateOpenXml());
    }

    [Fact]
    public void RandomBetween_fractional_references_use_integer_intersection_after_edits() {
        using var document = ExcelDocument.Create();
        var sheet = document.AddSheet("Bounds");
        sheet.CellValue(1, 1, 1.1d); sheet.CellValue(1, 2, 3.9d);
        sheet.CellFormula(1, 3, "RANDBETWEEN(A1,B1)");
        sheet.CellFormula(1, 4, "C1-C1");
        for (int run = 0; run < 16; run++) {
            Assert.Equal(2, document.Calculate());
            double number = sheet.CellAt(1, 3).GetValue<double>();
            Assert.InRange(number, 2d, 3d); Assert.Equal(Math.Truncate(number), number);
            Assert.Equal(0d, sheet.CellAt(1, 4).GetValue<double>());
        }
        sheet.CellValue(1, 1, -2.9d); sheet.CellValue(1, 2, -1.1d);
        Assert.Equal(2, document.Calculate());
        Assert.Equal(-2d, sheet.CellAt(1, 3).GetValue<double>());
        // A later empty integral interval retains its old cache. The existing
        // scalar-reference policy can use that cache in dependent formulas.
        sheet.CellValue(1, 1, 7.1d); sheet.CellValue(1, 2, 7.9d);
        Assert.Equal(1, document.Calculate());
        Assert.Equal(-2d, sheet.CellAt(1, 3).GetValue<double>());
        Assert.Equal(0d, sheet.CellAt(1, 4).GetValue<double>());
    }

    [Fact]
    public void RandomBetween_empty_integral_intervals_preserve_caches_and_reversed_bounds_are_typed_errors() {
        using var document = ExcelDocument.Create();
        var sheet = document.AddSheet("Bounds");
        string[] expressions = { "RANDBETWEEN(1.1,1.9)", "RANDBETWEEN(-1.9,-1.1)",
            "RANDBETWEEN(7.1,7.1)", "RANDBETWEEN(2147483648.1,2147483648.9)" };
        for (int index = 0; index < expressions.Length; index++) {
            sheet.CellValue(index + 1, 1, 42d); sheet.CellFormula(index + 1, 1, expressions[index]);
        }
        sheet.CellFormula(5, 1, "IFERROR(RANDBETWEEN(1.9,1.1),99)");
        sheet.CellFormula(6, 1, "IFERROR(RANDBETWEEN(-1.1,-1.9),99)");
        Assert.Equal(2, document.Calculate());
        for (int row = 1; row <= expressions.Length; row++) Assert.Equal(42d, sheet.CellAt(row, 1).GetValue<double>());
        Assert.Equal(99d, sheet.CellAt(5, 1).GetValue<double>());
        Assert.Equal(99d, sheet.CellAt(6, 1).GetValue<double>());
    }
}
