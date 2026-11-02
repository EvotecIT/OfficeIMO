using System;
using System.IO;
using OfficeIMO.Visio;
using Xunit;

namespace OfficeIMO.Tests {
    public class VisioShapeProperties {
        [Fact]
        public void RoundTripsShapeProperties() {
            string filePath = Path.Combine(Path.GetTempPath(), Guid.NewGuid() + ".vsdx");

            VisioDocument document = VisioDocument.Create(filePath);
            VisioPage page = document.AddPage("Page-1");
            VisioShape shape = new("1", OfficeLength.Inches(2), OfficeLength.Inches(3), OfficeLength.Inches(4), OfficeLength.Inches(6), string.Empty);
            page.Shapes.Add(shape);
            document.Save();

            VisioDocument roundTrip = VisioDocument.Load(filePath);
            VisioShape loaded = roundTrip.Pages[0].Shapes[0];
            Assert.Equal(0.02, loaded.LineWeight, 5);
            Assert.Equal(1.5, loaded.LocPinX.ToInches(), 5);
            Assert.Equal(2.5, loaded.LocPinY.ToInches(), 5);
            Assert.Equal(0.3, loaded.Angle, 5);
            Assert.Equal(4, loaded.Width.ToInches(), 5);
            Assert.Equal(6, loaded.Height.ToInches(), 5);
        }

        [Fact]
        public void DefaultsWhenMissing() {
            string filePath = Path.Combine(Path.GetTempPath(), Guid.NewGuid() + ".vsdx");

            VisioDocument document = VisioDocument.Create(filePath);
            VisioPage page = document.AddPage("Page-1");
            VisioShape shape = new("1", OfficeLength.Inches(1), OfficeLength.Inches(1), OfficeLength.Inches(2), OfficeLength.Inches(4), string.Empty);
            page.Shapes.Add(shape);
            document.Save();

            VisioDocument loaded = VisioDocument.Load(filePath);
            VisioShape s = loaded.Pages[0].Shapes[0];
            Assert.Equal(1, s.LocPinX.ToInches(), 5);
            Assert.Equal(2, s.LocPinY.ToInches(), 5);
            Assert.Equal(0.0138889, s.LineWeight, 5);
            Assert.Equal(0, s.Angle, 5);
        }

    }
}
