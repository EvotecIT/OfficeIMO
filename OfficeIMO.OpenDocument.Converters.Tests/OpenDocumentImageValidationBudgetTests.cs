using OfficeIMO.Drawing;
using OfficeIMO.OpenDocument;
using System.Text;
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

    [Fact]
    public void ConversionImageValidationBudgetChargesMalformedAttemptsBeforeIdentification() {
        var budget = new OdfImageValidationBudget();
        byte[] malformed = Encoding.UTF8.GetBytes("<svg xmlns='http://www.w3.org/2000/svg'>");

        for (int index = 0; index < 256; index++) {
            Assert.False(budget.TryConsume(malformed, "broken.svg"));
        }

        byte[] png = OfficePngWriter.Encode(new OfficeRasterImage(1, 1, OfficeColor.White));
        Assert.False(budget.TryConsume(png, "pixel.png"));
    }

    [Fact]
    public void ConversionImageValidationBudgetAcceptsDimensionlessSvg() {
        var budget = new OdfImageValidationBudget();
        byte[] svg = Encoding.UTF8.GetBytes(
            "<svg xmlns='http://www.w3.org/2000/svg'><path d='M0 0h1v1z'/></svg>");

        Assert.True(OdfImagePayloadValidator.TryResolvePreservedFileName(
            svg,
            "vector.svg",
            out string storedFileName,
            budget));
        Assert.Equal("image.svg", storedFileName);
    }
}
