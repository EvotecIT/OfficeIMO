using System;
using System.IO;
using OfficeIMO.Visio;
using Xunit;

namespace OfficeIMO.Tests {
    public class VisioShapeData {
        [Fact]
        public void RoundTripsShapeData() {
            string filePath = Path.Combine(Path.GetTempPath(), Guid.NewGuid() + ".vsdx");

            VisioDocument document = new();
            VisioPage page = document.AddPage("Page-1");
            VisioShape shape = new("1", 2, 3, 4, 5, string.Empty);
            shape.Data["Key"] = "Value";
            page.Shapes.Add(shape);
            document.Save(filePath);

            VisioDocument roundTrip = VisioDocument.Load(filePath);
            VisioShape loaded = roundTrip.Pages[0].Shapes[0];
            Assert.True(loaded.Data.TryGetValue("Key", out string? value));
            Assert.Equal("Value", value);
        }

    }
}
