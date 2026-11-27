using System.IO.Compression;
using System.Xml.Linq;
using OfficeIMO.Drawing;
using OfficeIMO.PowerPoint;
using Xunit;

namespace OfficeIMO.Tests {
    public partial class PowerPointImageExportTests {
        [Fact]
        public void ImageExportResolvesTheDefaultBodyFontWithoutTheSdk() {
            string fixture = Path.Combine(RepositoryTestPaths.Find(), "OfficeIMO.TestAssets", "Documents",
                "Charts", "AutomaticColors", "automatic-colors.pptx");
            using var deck = PowerPointDocument.Load(new MemoryStream(File.ReadAllBytes(fixture)));

            PowerPointSlideVisualInfo snapshot = deck.Slides[0].CreateVisualSnapshot(new PowerPointImageExportOptions {
                IncludeCharts = false
            });

            Assert.Contains("Classic style 2 clustered column", GetSnapshotPlainText(snapshot), StringComparison.Ordinal);
            Assert.False(deck.IsSdkLoaded);
        }

        [Fact]
        public void ImageExportReadsInheritedContentAndStylesWithoutTheSdk() {
            using var source = PowerPointDocument.Create();
            source.SlideSize.SetSize(OfficeLength.Points(320), OfficeLength.Points(180));
            PowerPointSlide slide = source.AddSlide(PowerPointSlideLayoutType.Blank);
            PowerPointSlideLayout layout = slide.Layout!;
            layout.BackgroundColor = OfficeColor.Parse("112233");
            layout.AddTextBox("Layout text", OfficeLength.Points(16), OfficeLength.Points(16),
                OfficeLength.Points(240), OfficeLength.Points(32));
            layout.AddPlaceholder(PowerPointPlaceholderType.Body,
                new PowerPointLayoutBox(OfficeLength.Points(16), OfficeLength.Points(64),
                    OfficeLength.Points(288), OfficeLength.Points(64)), index: 24);
            slide.AddTextBox("Slide text").BindToPlaceholder(PowerPointPlaceholderType.Body, index: 24);
            byte[] picture = OfficePngWriter.Encode(new OfficeRasterImage(2, 2, OfficeColor.CornflowerBlue));
            slide.AddImage(new MemoryStream(picture), OfficeImageFormat.Png, OfficeLength.Points(276),
                OfficeLength.Points(138), OfficeLength.Points(24), OfficeLength.Points(24));

            using var package = new MemoryStream();
            package.Write(source.ToBytes());
            using (var archive = new ZipArchive(package, ZipArchiveMode.Update, leaveOpen: true)) {
                ZipArchiveEntry entry = archive.GetEntry(layout.PartName.TrimStart('/'))!;
                using Stream stream = entry.Open();
                XDocument xml = XDocument.Load(stream);
                XNamespace p = xml.Root!.Name.Namespace;
                XNamespace a = "http://schemas.openxmlformats.org/drawingml/2006/main";
                XNamespace mc = "http://schemas.openxmlformats.org/markup-compatibility/2006";
                XNamespace unknown = "urn:officeimo:test";
                XElement data = xml.Root.Element(p + "cSld")!;
                XElement tree = data.Element(p + "spTree")!;
                XElement placeholder = tree.Elements(p + "sp").Single(shape => shape.Element(p + "nvSpPr")?
                    .Element(p + "nvPr")?.Element(p + "ph")?.Attribute("idx")?.Value == "24");
                XElement list = placeholder.Element(p + "txBody")!.Element(a + "lstStyle")!;
                list.Add(new XElement(a + "lvl1pPr", new XElement(a + "defRPr", new XAttribute("sz", 2300))));
                list.Add(new XElement(a + "extLst", new XElement(a + "ext", new XAttribute("uri", unknown),
                    new XElement(a + "lvl1pPr", new XElement(a + "defRPr", new XAttribute("sz", 9900))))));
                XElement ghost = new(tree.Elements(p + "sp").First(shape => shape.Descendants(a + "t").Any(t => t.Value == "Layout text")));
                ghost.Descendants(a + "t").Single().Value = "Ghost text";
                tree.Add(new XElement(p + "extLst", new XElement(p + "ext", new XAttribute("uri", unknown), ghost)));
                tree.Add(new XElement(mc + "AlternateContent",
                    new XElement(mc + "Choice", new XAttribute(XNamespace.Xmlns + "t", unknown),
                        new XAttribute("Requires", "t"), new XElement(ghost))));
                data.Add(new XElement(p + "extLst", new XElement(p + "ext", new XAttribute("uri", unknown),
                    new XElement(p + "bg", new XElement(p + "bgPr", new XElement(a + "solidFill",
                        new XElement(a + "srgbClr", new XAttribute("val", "FFEEDD"))))))));
                stream.SetLength(0);
                stream.Position = 0;
                xml.Save(stream);
            }

            using var deck = PowerPointDocument.Load(new MemoryStream(package.ToArray()));
            PowerPointSlide saved = deck.Slides[0];
            PowerPointSlideVisualInfo visual = saved.CreateVisualSnapshot();
            OfficeImageExportResult png = saved.ToImage(OfficeImageExportFormat.Png);
            Assert.False(deck.IsSdkLoaded);
            Assert.Contains("Layout text", GetSnapshotPlainText(visual), StringComparison.Ordinal);
            Assert.DoesNotContain("Ghost text", GetSnapshotPlainText(visual), StringComparison.Ordinal);
            Assert.Contains(visual.TextSpans, span => span.Text.Contains("Slide text", StringComparison.Ordinal)
                && Math.Abs(span.FontSize - 23D) < 0.01D);
            Assert.True(OfficePngReader.TryDecode(png.Bytes, out OfficeRasterImage? image));
            Assert.Equal(OfficeColor.Parse("112233"), image!.GetPixel(2, 2));
            Assert.Equal(OfficeColor.CornflowerBlue, image.GetPixel(288, 150));
        }
    }
}
