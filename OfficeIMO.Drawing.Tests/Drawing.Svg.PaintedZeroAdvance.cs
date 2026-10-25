using System.IO;
using System.Linq;
using System.Xml.Linq;
using OfficeIMO.Drawing;
using OfficeIMO.TestAssets;
using Xunit;

namespace OfficeIMO.Tests;

public partial class DrawingTests {
    [Fact]
    public void SvgPaintedZeroAdvanceKeepsNaturalInkAndLogicalSource() {
        byte[] data = File.ReadAllBytes(Path.Combine(RepositoryTestPaths.Find(),
            "OfficeIMO.Drawing.Tests/TestAssets/SourceSansPro-Regular.otf"));
        string? expectedPath = null;
        foreach (double width in new[] { 10D, 100D }) {
            var drawing = new OfficeDrawing(160D, 80D).AddFont("Zero CFF", data)
                .AddPositionedText("\u0301", 30D, 20D, width, 30D, new OfficeFontInfo("Zero CFF", 20D));
            Assert.Equal(0D, Assert.Single(drawing.Fonts.Faces).Program.Measure("\u0301", 20D));
            drawing.MarkLastTextAsPaintedGlyphs();

            XElement root = XElement.Parse(OfficeDrawingSvgExporter.ToSvg(drawing));
            Assert.Equal("\u0301", Assert.Single(root.Descendants().Attributes("aria-label")).Value);
            XElement path = Assert.Single(root.Descendants().Where(element => element.Name.LocalName == "path"));
            string ink = path.Attribute("d")!.Value;
            Assert.Contains('M', ink);
            if (expectedPath == null) expectedPath = ink;
            else Assert.Equal(expectedPath, ink);
            Assert.DoesNotContain(root.Descendants(), element => element.Name.LocalName == "text");
        }
    }
}
