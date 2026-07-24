using OfficeIMO.Excel;
using Xunit;

namespace OfficeIMO.Tests;

public partial class Excel {
    [Fact]
    public void Batch20_ConditionalIconKindRetainsPublishedOrdinals() {
        Assert.Equal(0, (int)ExcelConditionalIconKind.GreenUpArrow);
        Assert.Equal(1, (int)ExcelConditionalIconKind.YellowSideArrow);
        Assert.Equal(2, (int)ExcelConditionalIconKind.RedDownArrow);
        Assert.Equal(3, (int)ExcelConditionalIconKind.GreenCheck);
        Assert.Equal(4, (int)ExcelConditionalIconKind.YellowExclamation);
        Assert.Equal(5, (int)ExcelConditionalIconKind.RedCross);
        Assert.Equal(6, (int)ExcelConditionalIconKind.GreenCircle);
        Assert.Equal(7, (int)ExcelConditionalIconKind.YellowCircle);
        Assert.Equal(8, (int)ExcelConditionalIconKind.RedCircle);
        Assert.Equal(9, (int)ExcelConditionalIconKind.YellowUpArrow);
    }
}
