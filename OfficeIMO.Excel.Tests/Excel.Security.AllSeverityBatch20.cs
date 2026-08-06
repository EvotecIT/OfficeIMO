using OfficeIMO.Excel;
using Xunit;

namespace OfficeIMO.Tests;

public partial class Excel {
    [Fact]
    public void Batch20_ConditionalIconKindRetainsPublishedOrdinals() {
        Assert.Equal(0, (int)OfficeConditionalIconKind.GreenUpArrow);
        Assert.Equal(1, (int)OfficeConditionalIconKind.YellowSideArrow);
        Assert.Equal(2, (int)OfficeConditionalIconKind.RedDownArrow);
        Assert.Equal(3, (int)OfficeConditionalIconKind.GreenCheck);
        Assert.Equal(4, (int)OfficeConditionalIconKind.YellowExclamation);
        Assert.Equal(5, (int)OfficeConditionalIconKind.RedCross);
        Assert.Equal(6, (int)OfficeConditionalIconKind.GreenCircle);
        Assert.Equal(7, (int)OfficeConditionalIconKind.YellowCircle);
        Assert.Equal(8, (int)OfficeConditionalIconKind.RedCircle);
        Assert.Equal(9, (int)OfficeConditionalIconKind.YellowUpArrow);
    }

    [Theory]
    [InlineData("[m]")]
    [InlineData("[mm]")]
    public void Batch20_ElapsedMinuteFormatsDoNotRequireDateSystemShifts(string formatCode) {
        Assert.False(ExcelNumberFormatClassifier.LooksLikeDateSystemFormat(formatCode));
    }
}
