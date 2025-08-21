using System;
using System.IO;
using System.Linq;
using OfficeIMO.Visio;
using Xunit;

namespace OfficeIMO.Tests {
    public class VisioMasterShapeSize {
        [Fact]
        public void ShapesInheritMasterSize() {
            string filePath = Path.Combine(Path.GetTempPath(), Guid.NewGuid() + ".vsdx");

            VisioDocument document = VisioDocument.Create(filePath);
            VisioPage page = document.AddPage("Page-1");

            VisioShape masterRectShape = new("0", 0, 0, 2, 1, string.Empty);
            VisioMaster rectMaster = new("2", "Rectangle", masterRectShape);
            VisioShape rect = new("1", 1, 1, 2, 1, "First") { Master = rectMaster };

            VisioShape masterSquareShape = new("0", 0, 0, 2, 2, string.Empty);
            VisioMaster squareMaster = new("3", "Square", masterSquareShape);
            VisioShape square = new("2", 4, 1, 2, 2, "Second") { Master = squareMaster };

            page.Shapes.Add(rect);
            page.Shapes.Add(square);
            document.Save();

            VisioDocument loaded = VisioDocument.Load(filePath);
            VisioPage loadedPage = loaded.Pages[0];

            VisioShape loadedRect = loadedPage.Shapes.First(s => s.Id == "1");
            Assert.NotNull(loadedRect.Master);
            Assert.Equal(loadedRect.Master.Shape.Width, loadedRect.Width);
            Assert.Equal(loadedRect.Master.Shape.Height, loadedRect.Height);
            Assert.Equal(loadedRect.Master.Shape.LocPinX, loadedRect.LocPinX);
            Assert.Equal(loadedRect.Master.Shape.LocPinY, loadedRect.LocPinY);

            VisioShape loadedSquare = loadedPage.Shapes.First(s => s.Id == "2");
            Assert.NotNull(loadedSquare.Master);
            Assert.Equal(loadedSquare.Master.Shape.Width, loadedSquare.Width);
            Assert.Equal(loadedSquare.Master.Shape.Height, loadedSquare.Height);
            Assert.Equal(loadedSquare.Master.Shape.LocPinX, loadedSquare.LocPinX);
            Assert.Equal(loadedSquare.Master.Shape.LocPinY, loadedSquare.LocPinY);
        }
    }
}
