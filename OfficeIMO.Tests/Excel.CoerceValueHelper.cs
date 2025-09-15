using DocumentFormat.OpenXml.Spreadsheet;
using OfficeIMO.Excel;
using Xunit;

namespace OfficeIMO.Tests;

public class ExcelCoerceValueHelper
{
    [Fact]
    public void Coerce_NumberAndString()
    {
        var (numVal, numType) = CoerceValueHelper.Coerce(123, s => new CellValue(s));
        Assert.Equal(CellValues.Number, numType);
        Assert.Equal("123", numVal.Text);

        var (strVal, strType) = CoerceValueHelper.Coerce("Hello", s => new CellValue("IDX"));
        Assert.Equal(CellValues.SharedString, strType);
        Assert.Equal("IDX", strVal.Text);
    }
}
