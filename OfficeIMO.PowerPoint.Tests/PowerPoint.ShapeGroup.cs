using System.IO.Compression;
using System.Xml;
using System.Xml.Linq;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Presentation;
using DocumentFormat.OpenXml.Validation;
using OfficeIMO.PowerPoint;
using Xunit;

namespace OfficeIMO.Tests {
    public class PowerPointShapeGroupTests {
        [Fact]
        public void GroupShapes_CreatesGroupWithBounds() {
            string filePath = Path.Combine(Path.GetTempPath(), Guid.NewGuid() + ".pptx");
            try {
                using PowerPointDocument presentation = PowerPointDocument.Create(filePath);
                PowerPointSlide slide = presentation.AddSlide();
                PowerPointAutoShape a = slide.AddRectangle(OfficeIMO.OfficeLength.Emus(100), OfficeIMO.OfficeLength.Emus(200), OfficeIMO.OfficeLength.Emus(300), OfficeIMO.OfficeLength.Emus(400));
                PowerPointAutoShape b = slide.AddRectangle(OfficeIMO.OfficeLength.Emus(800), OfficeIMO.OfficeLength.Emus(1000), OfficeIMO.OfficeLength.Emus(200), OfficeIMO.OfficeLength.Emus(200));

                PowerPointGroupShape group = slide.GroupShapes(new PowerPointShape[] { a, b });

                Assert.Single(slide.Shapes);
                Assert.Equal(100, group.Left.ToEmus());
                Assert.Equal(200, group.Top.ToEmus());
                Assert.Equal(900, group.Width.ToEmus());
                Assert.Equal(1000, group.Height.ToEmus());
                Assert.False(presentation.IsSdkLoaded);
            } finally {
                if (File.Exists(filePath)) {
                    File.Delete(filePath);
                }
            }
        }

        [Fact]
        public void UngroupShape_RestoresChildren() {
            string filePath = Path.Combine(Path.GetTempPath(), Guid.NewGuid() + ".pptx");
            try {
                using PowerPointDocument presentation = PowerPointDocument.Create(filePath);
                PowerPointSlide slide = presentation.AddSlide();
                PowerPointAutoShape a = slide.AddRectangle(OfficeIMO.OfficeLength.Emus(0), OfficeIMO.OfficeLength.Emus(0), OfficeIMO.OfficeLength.Emus(100), OfficeIMO.OfficeLength.Emus(100));
                PowerPointAutoShape b = slide.AddRectangle(OfficeIMO.OfficeLength.Emus(200), OfficeIMO.OfficeLength.Emus(300), OfficeIMO.OfficeLength.Emus(150), OfficeIMO.OfficeLength.Emus(120));

                PowerPointGroupShape group = slide.GroupShapes(slide.Shapes);
                var children = slide.UngroupShape(group).ToList();

                Assert.Equal(2, children.Count);
                Assert.Equal(2, slide.Shapes.Count);

                PowerPointShape childA = children[0];
                PowerPointShape childB = children[1];

                Assert.Equal(0, childA.Left.ToEmus());
                Assert.Equal(0, childA.Top.ToEmus());
                Assert.Equal(100, childA.Width.ToEmus());
                Assert.Equal(100, childA.Height.ToEmus());

                Assert.Equal(200, childB.Left.ToEmus());
                Assert.Equal(300, childB.Top.ToEmus());
                Assert.Equal(150, childB.Width.ToEmus());
                Assert.Equal(120, childB.Height.ToEmus());
                Assert.False(presentation.IsSdkLoaded);
            } finally {
                if (File.Exists(filePath)) {
                    File.Delete(filePath);
                }
            }
        }

        [Fact]
        public void GroupedShapeReferencesFollowUngroupingAfterTheSdkOpens() {
            using PowerPointDocument presentation = PowerPointDocument.Create();
            PowerPointSlide slide = presentation.AddSlide();
            PowerPointAutoShape first = slide.AddShape(OfficePresetShapeType.Rectangle, name: "First");
            PowerPointAutoShape second = slide.AddShape(OfficePresetShapeType.Rectangle, name: "Second");
            PowerPointGroupShape group = slide.GroupShapes([first, second]);
            PowerPointShape heldChild = slide.GetGroupChildren(group)[1];
            Assert.False(presentation.IsSdkLoaded);

            _ = presentation.OpenXmlDocument;
            var children = slide.UngroupShape(group);
            first.Name = "Renamed first";
            heldChild.Name = "Renamed second";

            Assert.Equal("Renamed first", children[0].Name);
            Assert.Equal("Renamed second", children[1].Name);
            using var reopened = PowerPointDocument.Load(new MemoryStream(presentation.ToBytes(), writable: false));
            Assert.Equal(new[] { "Renamed first", "Renamed second" },
                Assert.Single(reopened.Slides).Shapes.Select(shape => shape.Name));
        }

        [Theory]
        [InlineData(false)]
        [InlineData(true)]
        public void UngroupRejectsGroupOwnedExtensionBranchesBeforeMovingAnyShape(bool openSdk) {
            byte[] authored;
            using (var presentation = PowerPointDocument.Create()) {
                PowerPointSlide slide = presentation.AddSlide();
                PowerPointAutoShape first = slide.AddRectangle(OfficeLength.Points(0), OfficeLength.Points(0),
                    OfficeLength.Points(10), OfficeLength.Points(10), "First");
                PowerPointAutoShape second = slide.AddRectangle(OfficeLength.Points(10), OfficeLength.Points(0),
                    OfficeLength.Points(10), OfficeLength.Points(10), "Second");
                slide.GroupShapes([first, second]);
                slide.AddRectangle(OfficeLength.Points(20), OfficeLength.Points(0),
                    OfficeLength.Points(10), OfficeLength.Points(10), "Following");

                var sdkSlide = presentation.OpenXmlDocument.PresentationPart!.SlideParts.Single().Slide;
                ShapeTree tree = sdkSlide.CommonSlideData!.ShapeTree!;
                var choice = new AlternateContentChoice(new ExtensionListWithModification()) { Requires = "p14" };
                var wrapper = new AlternateContent(choice, new AlternateContentFallback(new ExtensionListWithModification()));
                wrapper.AddNamespaceDeclaration("p14", "http://schemas.microsoft.com/office/powerpoint/2010/main");
                tree.Elements<GroupShape>().Single().AppendChild(wrapper);
                tree.AppendChild(new ExtensionListWithModification());
                authored = presentation.ToBytes();
            }

            using var original = PresentationDocument.Open(new MemoryStream(authored, writable: false), false,
                new OpenSettings {
                    MarkupCompatibilityProcessSettings = new MarkupCompatibilityProcessSettings(
                        MarkupCompatibilityProcessMode.ProcessAllParts, FileFormatVersions.Office2007)
                });
            Assert.Empty(new OpenXmlValidator().Validate(original.PresentationPart!.SlideParts.Single().Slide));

            using var loaded = PowerPointDocument.Load(new MemoryStream(authored, writable: false));
            PowerPointSlide reopened = Assert.Single(loaded.Slides);
            if (openSdk) _ = loaded.OpenXmlDocument;
            var group = Assert.IsType<PowerPointGroupShape>(reopened.Shapes[0]);

            Assert.Throws<InvalidOperationException>(() => reopened.UngroupShape(group));
            Assert.Equal(openSdk, loaded.IsSdkLoaded);
            Assert.Same(group, reopened.Shapes[0]);
            Assert.Equal("Following", reopened.Shapes[1].Name);
            Assert.Equal(new[] { "First", "Second" }, reopened.GetGroupChildren(group).Select(shape => shape.Name));
            Assert.True(XNode.DeepEquals(ReadDrawingTree(authored), ReadDrawingTree(loaded.ToBytes())));
        }

        private static XElement ReadDrawingTree(byte[] package) {
            using var zip = new ZipArchive(new MemoryStream(package, writable: false));
            using Stream slide = zip.GetEntry("ppt/slides/slide1.xml")!.Open();
            XNamespace p = "http://schemas.openxmlformats.org/presentationml/2006/main";
            return XDocument.Load(slide).Descendants(p + "spTree").Single();
        }

        [Fact]
        public void UngroupAndReorderPreserveNonElementDrawingMarkup() {
            byte[] authored;
            using (PowerPointDocument presentation = PowerPointDocument.Create()) {
                PowerPointSlide slide = presentation.AddSlide();
                PowerPointAutoShape a = slide.AddRectangle(OfficeLength.Points(5), OfficeLength.Points(5),
                    OfficeLength.Points(5), OfficeLength.Points(5), "First");
                PowerPointAutoShape b = slide.AddRectangle(OfficeLength.Points(15), OfficeLength.Points(5),
                    OfficeLength.Points(5), OfficeLength.Points(5), "Second");
                slide.GroupShapes([a, b]);
                GroupShape group = presentation.OpenXmlDocument.PresentationPart!.SlideParts.Single()
                    .Slide.CommonSlideData!.ShapeTree!.Elements<GroupShape>().Single();
                group.InsertAfter(new OpenXmlMiscNode(XmlNodeType.Comment, "<!--drawing marker-->"),
                    group.Elements<Shape>().First());
                authored = presentation.ToBytes();
            }

            using PowerPointDocument loaded = PowerPointDocument.Load(new MemoryStream(authored, writable: false));
            PowerPointSlide loadedSlide = Assert.Single(loaded.Slides);
            var children = loadedSlide.UngroupShape(Assert.IsType<PowerPointGroupShape>(Assert.Single(loadedSlide.Shapes)));
            loadedSlide.BringForward(children[0]);
            Assert.Equal(new[] { "Second", "First" }, loadedSlide.Shapes.Select(shape => shape.Name));

            using var saved = new ZipArchive(new MemoryStream(loaded.ToBytes(), writable: false));
            using Stream markup = saved.GetEntry("ppt/slides/slide1.xml")!.Open();
            XNamespace p = "http://schemas.openxmlformats.org/presentationml/2006/main";
            XElement tree = XDocument.Load(markup).Descendants(p + "spTree").Single();
            XComment marker = Assert.Single(tree.Nodes().OfType<XComment>());
            Assert.Equal("drawing marker", marker.Value);
        }

        [Fact]
        public void AlignGroupChildren_UsesGroupBounds() {
            string filePath = Path.Combine(Path.GetTempPath(), Guid.NewGuid() + ".pptx");
            try {
                using PowerPointDocument presentation = PowerPointDocument.Create(filePath);
                PowerPointSlide slide = presentation.AddSlide();
                PowerPointAutoShape a = slide.AddRectangle(OfficeIMO.OfficeLength.Emus(100), OfficeIMO.OfficeLength.Emus(200), OfficeIMO.OfficeLength.Emus(300), OfficeIMO.OfficeLength.Emus(100));
                PowerPointAutoShape b = slide.AddRectangle(OfficeIMO.OfficeLength.Emus(500), OfficeIMO.OfficeLength.Emus(250), OfficeIMO.OfficeLength.Emus(200), OfficeIMO.OfficeLength.Emus(100));

                PowerPointGroupShape group = slide.GroupShapes(slide.Shapes);
                slide.AlignGroupChildren(group, PowerPointShapeAlignment.Left);

                PowerPointLayoutBox bounds = slide.GetGroupChildBounds(group);
                var children = slide.GetGroupChildren(group);

                Assert.All(children, child => Assert.Equal(bounds.Left, child.Left));
            } finally {
                if (File.Exists(filePath)) {
                    File.Delete(filePath);
                }
            }
        }

        [Fact]
        public void DistributeGroupChildrenWithSpacing_UsesGroupBounds() {
            string filePath = Path.Combine(Path.GetTempPath(), Guid.NewGuid() + ".pptx");
            try {
                using PowerPointDocument presentation = PowerPointDocument.Create(filePath);
                PowerPointSlide slide = presentation.AddSlide();
                slide.AddRectangle(OfficeIMO.OfficeLength.Emus(0), OfficeIMO.OfficeLength.Emus(0), OfficeIMO.OfficeLength.Emus(1000), OfficeIMO.OfficeLength.Emus(500));
                slide.AddRectangle(OfficeIMO.OfficeLength.Emus(2000), OfficeIMO.OfficeLength.Emus(0), OfficeIMO.OfficeLength.Emus(1000), OfficeIMO.OfficeLength.Emus(500));
                slide.AddRectangle(OfficeIMO.OfficeLength.Emus(4000), OfficeIMO.OfficeLength.Emus(0), OfficeIMO.OfficeLength.Emus(1000), OfficeIMO.OfficeLength.Emus(500));

                PowerPointGroupShape group = slide.GroupShapes(slide.Shapes);
                slide.DistributeGroupChildrenWithSpacing(group, PowerPointShapeDistribution.Horizontal, spacing: OfficeIMO.OfficeLength.Emus(500));

                PowerPointLayoutBox bounds = slide.GetGroupChildBounds(group);
                var children = slide.GetGroupChildren(group).OrderBy(s => s.Left.ToEmus()).ToList();

                Assert.Equal(bounds.Left.ToEmus(), children[0].Left.ToEmus());
                Assert.Equal(bounds.Left.ToEmus() + 1500, children[1].Left.ToEmus());
                Assert.Equal(bounds.Left.ToEmus() + 3000, children[2].Left.ToEmus());
            } finally {
                if (File.Exists(filePath)) {
                    File.Delete(filePath);
                }
            }
        }

        [Fact]
        public void ArrangeGroupChildrenInGrid_UsesGroupBounds() {
            string filePath = Path.Combine(Path.GetTempPath(), Guid.NewGuid() + ".pptx");
            try {
                using PowerPointDocument presentation = PowerPointDocument.Create(filePath);
                PowerPointSlide slide = presentation.AddSlide();
                slide.AddRectangle(OfficeIMO.OfficeLength.Emus(0), OfficeIMO.OfficeLength.Emus(0), OfficeIMO.OfficeLength.Emus(100), OfficeIMO.OfficeLength.Emus(100));
                slide.AddRectangle(OfficeIMO.OfficeLength.Emus(200), OfficeIMO.OfficeLength.Emus(0), OfficeIMO.OfficeLength.Emus(100), OfficeIMO.OfficeLength.Emus(100));
                slide.AddRectangle(OfficeIMO.OfficeLength.Emus(0), OfficeIMO.OfficeLength.Emus(200), OfficeIMO.OfficeLength.Emus(100), OfficeIMO.OfficeLength.Emus(100));
                slide.AddRectangle(OfficeIMO.OfficeLength.Emus(200), OfficeIMO.OfficeLength.Emus(200), OfficeIMO.OfficeLength.Emus(100), OfficeIMO.OfficeLength.Emus(100));

                PowerPointGroupShape group = slide.GroupShapes(slide.Shapes);
                slide.ArrangeGroupChildrenInGrid(group, columns: 2, rows: 2);

                PowerPointLayoutBox bounds = slide.GetGroupChildBounds(group);
                long cellWidth = bounds.Width.ToEmus() / 2;
                long cellHeight = bounds.Height.ToEmus() / 2;
                var children = slide.GetGroupChildren(group).ToList();

                Assert.Equal(bounds.Left.ToEmus(), children[0].Left.ToEmus());
                Assert.Equal(bounds.Top.ToEmus(), children[0].Top.ToEmus());

                Assert.Equal(bounds.Left.ToEmus() + cellWidth, children[1].Left.ToEmus());
                Assert.Equal(bounds.Top.ToEmus(), children[1].Top.ToEmus());

                Assert.Equal(bounds.Left.ToEmus(), children[2].Left.ToEmus());
                Assert.Equal(bounds.Top.ToEmus() + cellHeight, children[2].Top.ToEmus());

                Assert.Equal(bounds.Left.ToEmus() + cellWidth, children[3].Left.ToEmus());
                Assert.Equal(bounds.Top.ToEmus() + cellHeight, children[3].Top.ToEmus());
            } finally {
                if (File.Exists(filePath)) {
                    File.Delete(filePath);
                }
            }
        }

        [Fact]
        public void GetGroupTextBoxes_ReturnsOnlyTextBoxes() {
            string filePath = Path.Combine(Path.GetTempPath(), Guid.NewGuid() + ".pptx");
            try {
                using PowerPointDocument presentation = PowerPointDocument.Create(filePath);
                PowerPointSlide slide = presentation.AddSlide();
                slide.AddTextBox("A",global::OfficeIMO.OfficeLength.Emus(0),global::OfficeIMO.OfficeLength.Emus(0),global::OfficeIMO.OfficeLength.Emus(100),global::OfficeIMO.OfficeLength.Emus(100));
                slide.AddTextBox("B",global::OfficeIMO.OfficeLength.Emus(200),global::OfficeIMO.OfficeLength.Emus(0),global::OfficeIMO.OfficeLength.Emus(100),global::OfficeIMO.OfficeLength.Emus(100));
                slide.AddRectangle(OfficeIMO.OfficeLength.Emus(400), OfficeIMO.OfficeLength.Emus(0), OfficeIMO.OfficeLength.Emus(100), OfficeIMO.OfficeLength.Emus(100));

                PowerPointGroupShape group = slide.GroupShapes(slide.Shapes);
                var boxes = slide.GetGroupTextBoxes(group);

                Assert.Equal(2, boxes.Count);
                Assert.All(boxes, box => Assert.False(string.IsNullOrEmpty(box.Text)));
            } finally {
                if (File.Exists(filePath)) {
                    File.Delete(filePath);
                }
            }
        }
    }
}
