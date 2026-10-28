using System;
using System.IO;
using System.Linq;
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
            } finally {
                if (File.Exists(filePath)) {
                    File.Delete(filePath);
                }
            }
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
                slide.DistributeGroupChildrenWithSpacing(group, PowerPointShapeDistribution.Horizontal, spacingEmus: 500);

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
