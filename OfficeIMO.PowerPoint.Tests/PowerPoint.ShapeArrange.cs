using System.Text;
using System.Xml.Linq;
using DocumentFormat.OpenXml.Drawing.Charts;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Presentation;
using OfficeIMO.Core.Internal.OpenXml;
using OfficeIMO.PowerPoint;
using Xunit;

namespace OfficeIMO.Tests {
    public class PowerPointShapeArrangeTests {
        [Fact]
        public void DuplicateShape_OffsetsAndKeepsSize() {
            string filePath = Path.Combine(Path.GetTempPath(), Guid.NewGuid() + ".pptx");

            try {
                using PowerPointDocument presentation = PowerPointDocument.Create(filePath);
                PowerPointSlide slide = presentation.AddSlide();
                PowerPointAutoShape original = slide.AddRectangle(OfficeIMO.OfficeLength.Emus(1000), OfficeIMO.OfficeLength.Emus(2000), OfficeIMO.OfficeLength.Emus(3000), OfficeIMO.OfficeLength.Emus(4000));

                OfficeXmlTree tree = slide.PartRef.Root.Tree!;
                int version = tree.Version;
                Assert.Throws<OverflowException>(() => slide.DuplicateShape(original, offsetY: OfficeLength.Emus(long.MaxValue)));
                Assert.Equal(version, tree.Version);

                PowerPointShape duplicate = slide.DuplicateShape(original, OfficeIMO.OfficeLength.Emus(500), OfficeIMO.OfficeLength.Emus(600));

                Assert.Equal(2, slide.Shapes.Count);
                Assert.Equal(original.Width.ToEmus(), duplicate.Width.ToEmus());
                Assert.Equal(original.Height.ToEmus(), duplicate.Height.ToEmus());
                Assert.Equal(original.Left.ToEmus() + 500, duplicate.Left.ToEmus());
                Assert.Equal(original.Top.ToEmus() + 600, duplicate.Top.ToEmus());
                Assert.NotEqual(original.Name, duplicate.Name);
                Assert.False(presentation.IsSdkLoaded);
            } finally {
                if (File.Exists(filePath)) {
                    File.Delete(filePath);
                }
            }
        }

        [Fact]
        public void BringForwardAndSendBackward_ReordersShapes() {
            string filePath = Path.Combine(Path.GetTempPath(), Guid.NewGuid() + ".pptx");

            try {
                using PowerPointDocument presentation = PowerPointDocument.Create(filePath);
                PowerPointSlide slide = presentation.AddSlide();
                PowerPointAutoShape a = slide.AddRectangle(OfficeIMO.OfficeLength.Emus(0), OfficeIMO.OfficeLength.Emus(0), OfficeIMO.OfficeLength.Emus(1000), OfficeIMO.OfficeLength.Emus(1000));
                PowerPointAutoShape b = slide.AddRectangle(OfficeIMO.OfficeLength.Emus(0), OfficeIMO.OfficeLength.Emus(0), OfficeIMO.OfficeLength.Emus(1000), OfficeIMO.OfficeLength.Emus(1000));
                PowerPointAutoShape c = slide.AddRectangle(OfficeIMO.OfficeLength.Emus(0), OfficeIMO.OfficeLength.Emus(0), OfficeIMO.OfficeLength.Emus(1000), OfficeIMO.OfficeLength.Emus(1000));

                slide.BringForward(a);

                Assert.Equal(new PowerPointShape[] { b, a, c }, slide.Shapes.ToArray());

                slide.SendBackward(c);

                Assert.Equal(new PowerPointShape[] { b, c, a }, slide.Shapes.ToArray());
            } finally {
                if (File.Exists(filePath)) {
                    File.Delete(filePath);
                }
            }
        }

        [Fact]
        public void DuplicateGroupShape_AssignsUniqueNonVisualMetadata() {
            string filePath = Path.Combine(Path.GetTempPath(), Guid.NewGuid() + ".pptx");

            try {
                using PowerPointDocument presentation = PowerPointDocument.Create(filePath);
                PowerPointSlide slide = presentation.AddSlide();
                slide.AddRectangle(OfficeIMO.OfficeLength.Emus(0), OfficeIMO.OfficeLength.Emus(0), OfficeIMO.OfficeLength.Emus(1000), OfficeIMO.OfficeLength.Emus(1000));
                slide.AddRectangle(OfficeIMO.OfficeLength.Emus(1500), OfficeIMO.OfficeLength.Emus(0), OfficeIMO.OfficeLength.Emus(1000), OfficeIMO.OfficeLength.Emus(1000));

                PowerPointGroupShape group = slide.GroupShapes(slide.Shapes);
                PowerPointGroupShape duplicate = Assert.IsType<PowerPointGroupShape>(slide.DuplicateShape(group, OfficeIMO.OfficeLength.Emus(500), OfficeIMO.OfficeLength.Emus(500)));

                Assert.NotEqual(group.Name, duplicate.Name);
                presentation.Save();
            } finally {
                if (File.Exists(filePath)) {
                    using PresentationDocument document = PresentationDocument.Open(filePath, false);
                    var nonVisualProps = document.PresentationPart!.SlideParts.First().Slide
                        .Descendants<NonVisualDrawingProperties>()
                        .ToList();

                    Assert.Equal(nonVisualProps.Count, nonVisualProps.Select(property => property.Id?.Value).Distinct().Count());
                    Assert.Equal(nonVisualProps.Count, nonVisualProps.Select(property => property.Name?.Value).Distinct().Count());
                    File.Delete(filePath);
                }
            }
        }

        [Fact]
        public void DuplicateGroupKeepsMatchingShapeMarkupInsideExtensions() {
            using var source = PowerPointDocument.Create();
            PowerPointSlide slide = source.AddSlide();
            slide.AddShape(OfficePresetShapeType.Rectangle);
            slide.AddShape(OfficePresetShapeType.Ellipse);
            slide.GroupShapes(slide.Shapes);
            XNamespace p = "http://schemas.openxmlformats.org/presentationml/2006/main";
            XNamespace ignored = "urn:officeimo:test:shape-copy";
            using var input = new MemoryStream(source.ToBytes());
            using OfficeOpcEditablePackage package = OfficeOpcEditablePackage.Open(input, leaveOpen: true);
            const string slidePart = "/ppt/slides/slide1.xml";
            XDocument slideMarkup = XDocument.Parse(Encoding.UTF8.GetString(package.ReadPartBytes(slidePart)));
            slideMarkup.Descendants(p + "grpSp").Single().Add(new XElement(p + "extLst",
                new XElement(p + "ext", new XAttribute("uri", "{543D527F-FF81-4FD8-AC94-5833C766C5B9}"),
                    new XElement(ignored + "value", new XElement(p + "sp", new XElement(p + "nvSpPr",
                        new XElement(p + "cNvPr", new XAttribute("id", 2), new XAttribute("name", "Preserved extension"))))))));
            byte[] xml = Encoding.UTF8.GetBytes(slideMarkup.ToString(SaveOptions.DisableFormatting));
            package.ReplacePart(slidePart, xml, xml.Length);
            using var edited = new MemoryStream();
            package.Save(edited);
            edited.Position = 0;

            using var document = PowerPointDocument.Load(edited);
            PowerPointSlide loaded = document.Slides.Single();
            PowerPointGroupShape original = Assert.IsType<PowerPointGroupShape>(Assert.Single(loaded.Shapes));
            var duplicate = Assert.IsType<PowerPointGroupShape>(loaded.DuplicateShape(original, OfficeLength.Emus(10)));
            byte[] saved = document.ToBytes();

            Assert.NotEqual(original.Id, duplicate.Id);
            Assert.False(document.IsSdkLoaded);
            using OfficeOpcEditablePackage result = OfficeOpcEditablePackage.Open(new MemoryStream(saved));
            XDocument savedSlide = XDocument.Parse(Encoding.UTF8.GetString(result.ReadPartBytes(slidePart)));
            XElement[] preserved = savedSlide.Descendants(ignored + "value").Select(element => element.Descendants(p + "cNvPr").Single()).ToArray();
            Assert.Equal(2, preserved.Length);
            Assert.All(preserved, element => {
                Assert.Equal("2", (string?)element.Attribute("id"));
                Assert.Equal("Preserved extension", (string?)element.Attribute("name"));
            });
        }

        [Fact]
        public void DuplicateAfterUngroupKeepsCompatibilityBindingsWhenSaved() {
            using var source = PowerPointDocument.Create();
            PowerPointSlide slide = source.AddSlide();
            slide.AddShape(OfficePresetShapeType.Rectangle, name: "Carried");
            slide.AddShape(OfficePresetShapeType.Ellipse);
            slide.GroupShapes(slide.Shapes);

            XNamespace p = "http://schemas.openxmlformats.org/presentationml/2006/main";
            XNamespace mc = "http://schemas.openxmlformats.org/markup-compatibility/2006";
            XNamespace feature = "urn:officeimo:test:carried-copy";
            using var input = new MemoryStream(source.ToBytes());
            using OfficeOpcEditablePackage package = OfficeOpcEditablePackage.Open(input, leaveOpen: true);
            const string slidePart = "/ppt/slides/slide1.xml";
            XDocument markup = XDocument.Parse(Encoding.UTF8.GetString(package.ReadPartBytes(slidePart)));
            XElement group = markup.Descendants(p + "grpSp").Single();
            group.SetAttributeValue(XNamespace.Xmlns + "feature", feature.NamespaceName);
            group.Elements(p + "sp").First().Add(new XElement(mc + "AlternateContent",
                new XElement(mc + "Choice", new XAttribute("Requires", "feature"), new XElement(p + "extLst")),
                new XElement(mc + "Fallback")));
            byte[] xml = Encoding.UTF8.GetBytes(markup.ToString(SaveOptions.DisableFormatting));
            package.ReplacePart(slidePart, xml, xml.Length);
            using var edited = new MemoryStream();
            package.Save(edited);
            edited.Position = 0;

            using var document = PowerPointDocument.Load(edited);
            PowerPointSlide loaded = document.Slides.Single();
            loaded.UngroupShape(Assert.IsType<PowerPointGroupShape>(Assert.Single(loaded.Shapes)));
            loaded.DuplicateShape(loaded.GetShapeByName("Carried")!);
            byte[] saved = document.ToBytes();

            Assert.False(document.IsSdkLoaded);
            using OfficeOpcEditablePackage result = OfficeOpcEditablePackage.Open(new MemoryStream(saved));
            XDocument after = XDocument.Parse(Encoding.UTF8.GetString(result.ReadPartBytes(slidePart)));
            XElement[] choices = after.Descendants(mc + "Choice").ToArray();
            Assert.Equal(2, choices.Length);
            Assert.All(choices, choice => Assert.Equal(feature, choice.GetNamespaceOfPrefix("feature")));
        }

        [Fact]
        public void DuplicateChart_CreatesIndependentChartPart() {
            string filePath = Path.Combine(Path.GetTempPath(), Guid.NewGuid() + ".pptx");

            try {
                using (PowerPointDocument presentation = PowerPointDocument.Create(filePath)) {
                    PowerPointSlide slide = presentation.AddSlide();
                    PowerPointChart original = slide.AddChart();
                    PowerPointChart duplicate = Assert.IsType<PowerPointChart>(slide.DuplicateShape(original, OfficeIMO.OfficeLength.Emus(250000), OfficeIMO.OfficeLength.Emus(0)));

                    duplicate.UpdateData(new PowerPointChartData(
                        new[] { "Jan", "Feb" },
                        new[] { new PowerPointChartSeries("Duplicate", new[] { 10d, 20d }) }));

                    presentation.Save();
                }

                using (PresentationDocument document = PresentationDocument.Open(filePath, false)) {
                    SlidePart slidePart = document.PresentationPart!.SlideParts.First();
                    GraphicFrame[] chartFrames = slidePart.Slide.CommonSlideData!.ShapeTree!
                        .Elements<GraphicFrame>()
                        .Where(frame => frame.Graphic?.GraphicData?.GetFirstChild<ChartReference>() != null)
                        .ToArray();

                    Assert.Equal(2, chartFrames.Length);

                    string[] relationshipIds = chartFrames
                        .Select(frame => frame.Graphic!.GraphicData!.GetFirstChild<ChartReference>()!.Id!.Value!)
                        .ToArray();
                    Assert.Equal(2, relationshipIds.Distinct(StringComparer.Ordinal).Count());

                    ChartPart originalChartPart = (ChartPart)slidePart.GetPartById(relationshipIds[0]);
                    ChartPart duplicateChartPart = (ChartPart)slidePart.GetPartById(relationshipIds[1]);

                    List<string> originalSeries = GetSeriesNames(originalChartPart);
                    List<string> duplicateSeries = GetSeriesNames(duplicateChartPart);

                    Assert.Equal(new[] { "Series 1", "Series 2", "Series 3" }, originalSeries);
                    Assert.Equal(new[] { "Duplicate" }, duplicateSeries);
                }
            } finally {
                if (File.Exists(filePath)) {
                    File.Delete(filePath);
                }
            }
        }

        private static List<string> GetSeriesNames(ChartPart chartPart) {
            return chartPart.ChartSpace!.Descendants<SeriesText>()
                .Select(series => series.StringReference?.StringCache?.Elements<StringPoint>()
                    .Select(point => point.NumericValue?.Text)
                    .LastOrDefault(text => !string.IsNullOrWhiteSpace(text))
                    ?? series.InnerText)
                .Where(text => !string.IsNullOrWhiteSpace(text))
                .ToList();
        }
    }
}
