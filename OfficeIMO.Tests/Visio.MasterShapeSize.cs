using System;
using System.IO;
using System.Linq;
using OfficeIMO.Visio;
using Xunit;

namespace OfficeIMO.Tests {
    public class VisioMasterShapeSize {
        private static string AssetsPath => Path.GetFullPath(Path.Combine(AppContext.BaseDirectory, "..", "..", "..", "..", "Assets"));

        [Fact]
        public void ShapesInheritMasterSize() {
            string file = Path.Combine(AssetsPath, "DrawingWithShapes.vsdx");
            VisioDocument doc = VisioDocument.Load(file);
            VisioPage page = doc.Pages[0];

            VisioShape rectangle = page.Shapes.First(s => s.NameU == "Rectangle");
            Assert.NotNull(rectangle.Master);
            Assert.Equal(rectangle.Master.Shape.Width, rectangle.Width);
            Assert.Equal(rectangle.Master.Shape.Height, rectangle.Height);
            Assert.Equal(rectangle.Master.Shape.LocPinX, rectangle.LocPinX);
            Assert.Equal(rectangle.Master.Shape.LocPinY, rectangle.LocPinY);

            VisioShape square = page.Shapes.First(s => s.NameU == "Square");
            Assert.NotNull(square.Master);
            Assert.Equal(square.Master.Shape.Width, square.Width);
            Assert.Equal(square.Master.Shape.Height, square.Height);
            Assert.Equal(square.Master.Shape.LocPinX, square.LocPinX);
            Assert.Equal(square.Master.Shape.LocPinY, square.LocPinY);
        }
    }
}
