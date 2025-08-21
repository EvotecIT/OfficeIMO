using System;
using System.IO;
using OfficeIMO.Visio;
using Xunit;

namespace OfficeIMO.Tests {
    public class VisioShapeDataAndText {
        [Fact]
        public void RoundTripsShapeDataAndText() {
            string filePath = Path.Combine(Path.GetTempPath(), Guid.NewGuid() + ".vsdx");

            VisioDocument document = VisioDocument.Create(filePath);
            VisioPage page = document.AddPage("Page-1");
            VisioShape shape = new("1", 2, 3, 4, 5, string.Empty);
            shape.Data["Key"] = "Value";
            shape.Text = "Hello";
            page.Shapes.Add(shape);
            document.Save();

            VisioDocument roundTrip = VisioDocument.Load(filePath);
            VisioShape loaded = roundTrip.Pages[0].Shapes[0];
            Assert.True(loaded.Data.TryGetValue("Key", out string? value));
            Assert.Equal("Value", value);
            Assert.Equal("Hello", loaded.Text);
        }
    }
}
