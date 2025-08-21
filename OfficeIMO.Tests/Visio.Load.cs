using System;
using System.IO;
using OfficeIMO.Visio;
using Xunit;

namespace OfficeIMO.Tests {
    public class VisioLoad {
        [Fact]
        public void CanRoundTripVisioDocument() {
            string filePath = Path.Combine(Path.GetTempPath(), Guid.NewGuid() + ".vsdx");

            VisioDocument document = VisioDocument.Create(filePath);
            VisioPage page = document.AddPage("Page-1");
            VisioShape masterShape = new("0", 0, 0, 1, 1, string.Empty);
            VisioMaster master = new("2", "Rectangle", masterShape);

            VisioShape shape1 = new("1", 1, 2, 1, 1, "Rectangle") { Master = master };
            VisioShape shape2 = new("3", 3, 4, 2, 3, "Rectangle") { Master = master };
            page.Shapes.Add(shape1);
            page.Shapes.Add(shape2);
            document.Save();

            VisioDocument loaded = VisioDocument.Load(filePath);
            Assert.Single(loaded.Pages);
            VisioPage loadedPage = loaded.Pages[0];
            Assert.Equal(2, loadedPage.Shapes.Count);
            VisioShape loadedShape1 = loadedPage.Shapes[0];
            VisioShape loadedShape2 = loadedPage.Shapes[1];
            Assert.Equal("1", loadedShape1.Id);
            Assert.Equal(1d, loadedShape1.Width);
            Assert.Equal(1d, loadedShape1.Height);
            Assert.Equal(2d, loadedShape2.Width);
            Assert.Equal(3d, loadedShape2.Height);

            string secondPath = Path.Combine(Path.GetTempPath(), Guid.NewGuid() + ".vsdx");
            loaded.Save(secondPath);
            VisioDocument roundTrip = VisioDocument.Load(secondPath);
            VisioShape rtShape1 = roundTrip.Pages[0].Shapes[0];
            VisioShape rtShape2 = roundTrip.Pages[0].Shapes[1];
            Assert.Equal(1d, rtShape1.Width);
            Assert.Equal(1d, rtShape1.Height);
            Assert.Equal(2d, rtShape2.Width);
            Assert.Equal(3d, rtShape2.Height);
        }
    }
}
