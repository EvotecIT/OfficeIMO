using OfficeIMO.Excel.Fluent;
using Xunit;

namespace OfficeIMO.Tests {
    public partial class Excel {
        [Fact]
        public void SheetTheme_Default_ReturnsIndependentInstances() {
            SheetTheme first = SheetTheme.Default;
            SheetTheme second = SheetTheme.Default;

            first.TitleColorHex = "#FFFFFF";

            Assert.NotSame(first, second);
            Assert.Equal("#1F497D", second.TitleColorHex);
        }
    }
}
