using System;
using System.IO;
using System.Linq;
using OfficeIMO.Visio;
using OfficeIMO.Visio.Diagrams;
using Xunit;

namespace OfficeIMO.Tests {
    public class VisioGraphDiagramTitleLayoutTests {
        [Theory]
        [InlineData(VisioMeasurementUnit.Inches, 1D)]
        [InlineData(VisioMeasurementUnit.Centimeters, 2.54D)]
        public void MultilineTitleReservesItsNativeFontHeightThroughSave(VisioMeasurementUnit unit, double scale) {
            const string text = "Service delivery\nThree editable services · prepared routes";
            string path = Path.Combine(Path.GetTempPath(), Guid.NewGuid() + ".vsdx");
            var theme = VisioStyleTheme.Technical();
            theme.TitleText.FontFamily = "Arial";
            theme.TitleText.Size = 22D;
            theme.TitleText.TopMargin = 0.04D;
            theme.TitleText.BottomMargin = 0.04D;
            try {
                VisioDocument document = VisioDocument.Create(path)
                    .GraphDiagram("Services", graph => graph
                        .PageSize(6.25D * scale, 3.75D * scale, unit)
                        .Margins(0.4D * scale, 0.16D * scale, 0.4D * scale, 0.4D * scale)
                        .NodeSize(1.65D * scale, 0.82D * scale)
                        .Spacing(1.1D * scale, 0.65D * scale)
                        .Theme(theme)
                        .Title(text, height: 0.45D * scale, gap: 0.08D * scale)
                        .Root("api", "API")
                        .Node("database", "Database")
                        .Edge("api", "database"));

                void Check(VisioPage page) {
                    VisioShape title = page.Shapes.Single(shape => shape.Id == "title");
                    Assert.Equal(text, title.Text);
                    Assert.True(title.Height > 2D * 22D / 72D);
                    Assert.Equal(5.45D, title.Width, 6);
                    Assert.Equal("Arial", title.TextStyle!.FontFamily);
                    Assert.Equal(22D, title.TextStyle.Size!.Value, 6);
                    Assert.True(title.TextStyle.Bold);
                    Assert.Equal(0.04D, title.TextStyle.TopMargin!.Value, 6);
                    Assert.Equal(0.04D, title.TextStyle.BottomMargin!.Value, 6);
                    var bounds = title.GetShapeBounds();
                    Assert.True(bounds.Top <= page.Height);
                    Assert.True(bounds.Bottom >= 0);
                    Assert.All(page.Shapes.Where(shape => shape.Id == "api" || shape.Id == "database"),
                        shape => Assert.True(bounds.Bottom >= shape.GetShapeBounds().Top + 0.08D - 0.000001D));
                }

                Check(document.Pages[0]);
                double height = document.Pages[0].Shapes.Single(shape => shape.Id == "title").Height;
                document.Save();
                Assert.Empty(VisioValidator.Validate(path));
                VisioPage loaded = VisioDocument.Load(path).Pages[0];
                Check(loaded);
                Assert.Equal(height, loaded.Shapes.Single(shape => shape.Id == "title").Height, 6);
            } finally {
                if (File.Exists(path)) File.Delete(path);
            }
        }
    }
}
