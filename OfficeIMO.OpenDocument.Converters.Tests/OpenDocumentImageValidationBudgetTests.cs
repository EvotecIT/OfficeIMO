using OfficeIMO.Drawing;
using OfficeIMO.OpenDocument;
using Xunit;

namespace OfficeIMO.OpenDocument.Converters.Tests;

public sealed class OpenDocumentImageValidationBudgetTests {
    [Fact]
    public void ConversionImageValidationBudgetBoundsRepeatedDecodeWork() {
        byte[] png = OfficePngWriter.Encode(new OfficeRasterImage(1, 1, OfficeColor.White));
        var budget = new OdfImageValidationBudget();

        for (int index = 0; index < 256; index++) {
            Assert.True(budget.TryConsume(png, "pixel.png"));
        }

        Assert.False(budget.TryConsume(png, "pixel.png"));
    }
}
