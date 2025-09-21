using OfficeIMO.Pdf;
using Xunit;

namespace OfficeIMO.Tests.Pdf;

public class PdfDocBulletListTests {
    [Fact]
    public void Bullets_WithNullItems_ThrowsArgumentNullException() {
        var doc = PdfDoc.Create();

        var exception = Assert.Throws<ArgumentNullException>(() => doc.Bullets(null!));

        Assert.Equal("items", exception.ParamName);
        Assert.Contains("Parameter 'items' cannot be null.", exception.Message, StringComparison.Ordinal);
    }
}
