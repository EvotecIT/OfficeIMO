using System;
using System.IO;
using OfficeIMO.Visio;
using Xunit;

namespace OfficeIMO.Tests {
    public class VisioLoad {
        [Fact]
        public void CanRoundTripVisioDocument() {
            string filePath = Path.Combine(Path.GetTempPath(), Guid.NewGuid() + ".vsdx");

            VisioDocument document = new();
            VisioPage page = document.AddPage("Page-1");
            VisioShape shape = new("1", 1, 2, 1, 1, "Rectangle");
            VisioMaster master = new("2", "Rectangle", shape);
            shape.Master = master;
            page.Shapes.Add(shape);
            document.Save(filePath);

            VisioDocument loaded = VisioDocument.Load(filePath);
            Assert.Single(loaded.Pages);
            VisioPage loadedPage = loaded.Pages[0];
            Assert.Single(loadedPage.Shapes);
            VisioShape loadedShape = loadedPage.Shapes[0];
            Assert.Equal("1", loadedShape.Id);
            Assert.Equal("Rectangle", loadedShape.NameU);
            Assert.Equal("Rectangle", loadedShape.Text);
            Assert.Equal(1d, loadedShape.PinX);
            Assert.Equal(2d, loadedShape.PinY);
            Assert.Equal(0d, loadedShape.Width);
            Assert.Equal(0d, loadedShape.Height);
            Assert.NotNull(loadedShape.Master);

            string secondPath = Path.Combine(Path.GetTempPath(), Guid.NewGuid() + ".vsdx");
            loaded.Save(secondPath);
            VisioDocument roundTrip = VisioDocument.Load(secondPath);
            VisioShape roundTripShape = roundTrip.Pages[0].Shapes[0];
            Assert.Equal(loadedShape.Text, roundTripShape.Text);
            Assert.Equal(0d, roundTripShape.Width);
        }
    }
}
