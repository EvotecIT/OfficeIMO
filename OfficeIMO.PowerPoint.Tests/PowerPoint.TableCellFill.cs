using DocumentFormat.OpenXml.Drawing;
using OfficeIMO.PowerPoint;
using Xunit;

namespace OfficeIMO.Tests {
    public class PowerPointTableCellFillTests {
        [Fact]
        public void ExplicitCellFillTransitionsPreserveBordersPaddingAndSchema() {
            using var source = PowerPointPresentation.Create();
            var table = source.AddSlide().AddTablePoints(1, 1, 20, 20, 120, 48);
            var cell = table.GetCell(0, 0);
            cell.BorderColor = "008000";
            cell.PaddingLeftPoints = 4;
            cell.Cell.TableCellProperties!.AddChild(new ExtensionList(), true);
            cell.FillColor = "FF0000";
            cell.NoFill = true;
            Assert.Null(cell.FillColor);
            Assert.True(cell.NoFill);
            Assert.Empty(source.ValidateDocument());
            cell.FillColor = "0000FF";
            Assert.False(cell.NoFill);
            Assert.Equal("0000FF", cell.FillColor);
            Assert.Empty(source.ValidateDocument());
            cell.FillColor = null;
            Assert.Null(cell.FillColor);
            Assert.False(cell.NoFill);
            cell.NoFill = true;
            cell.NoFill = false;
            Assert.False(cell.NoFill);
            Assert.Equal("008000", cell.BorderColor);
            Assert.Equal(4, cell.PaddingLeftPoints);
            Assert.NotNull(cell.Cell.TableCellProperties.GetFirstChild<ExtensionList>());
            Assert.Empty(source.ValidateDocument());
        }
    }
}
