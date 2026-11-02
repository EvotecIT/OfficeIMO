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

            VisioShape masterRectShape = new("0", OfficeLength.Inches(0), OfficeLength.Inches(0), OfficeLength.Inches(2), OfficeLength.Inches(1), string.Empty);
            VisioMaster rectMaster = new("2", "Rectangle", masterRectShape);
            VisioShape rect = new("1", OfficeLength.Inches(1), OfficeLength.Inches(1), OfficeLength.Inches(2), OfficeLength.Inches(1), "First") { Master = rectMaster };

            VisioShape masterSquareShape = new("0", OfficeLength.Inches(0), OfficeLength.Inches(0), OfficeLength.Inches(2), OfficeLength.Inches(2), string.Empty);
            VisioMaster squareMaster = new("3", "Square", masterSquareShape);
            VisioShape square = new("2", OfficeLength.Inches(4), OfficeLength.Inches(1), OfficeLength.Inches(2), OfficeLength.Inches(2), "Second") { Master = squareMaster };

            page.Shapes.Add(rect);
            page.Shapes.Add(square);
            document.Save();

            VisioDocument loaded = VisioDocument.Load(filePath);
            VisioPage loadedPage = loaded.Pages[0];

            VisioShape loadedRect = loadedPage.Shapes.First(s => s.Id == "1");
            Assert.NotNull(loadedRect.Master);
            Assert.Equal(loadedRect.Master.Shape.Width.ToInches(), loadedRect.Width.ToInches());
            Assert.Equal(loadedRect.Master.Shape.Height.ToInches(), loadedRect.Height.ToInches());
            Assert.Equal(loadedRect.Master.Shape.LocPinX.ToInches(), loadedRect.LocPinX.ToInches());
            Assert.Equal(loadedRect.Master.Shape.LocPinY.ToInches(), loadedRect.LocPinY.ToInches());

            VisioShape loadedSquare = loadedPage.Shapes.First(s => s.Id == "2");
            Assert.NotNull(loadedSquare.Master);
            Assert.Equal(loadedSquare.Master.Shape.Width.ToInches(), loadedSquare.Width.ToInches());
            Assert.Equal(loadedSquare.Master.Shape.Height.ToInches(), loadedSquare.Height.ToInches());
            Assert.Equal(loadedSquare.Master.Shape.LocPinX.ToInches(), loadedSquare.LocPinX.ToInches());
            Assert.Equal(loadedSquare.Master.Shape.LocPinY.ToInches(), loadedSquare.LocPinY.ToInches());
        }
    }
}
