using OfficeIMO.Drawing;
using Xunit;

namespace OfficeIMO.Tests {
    public class DrawingPageSizeTests {
        [Fact]
        public void OfficePageSize_ConvertsPhysicalSizeToPixelsAndPoints() {
            OfficePageSize letter = OfficePageSizes.Letter;

            Assert.Equal(816, letter.ToPixelWidth(96D));
            Assert.Equal(1056, letter.ToPixelHeight(96D));
            Assert.Equal(612D, letter.ToPointWidth());
            Assert.Equal(792D, letter.ToPointHeight());
        }
    }
}
