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
                using PowerPointPresentation presentation = PowerPointPresentation.Create(filePath);
                PowerPointSlide slide = presentation.AddSlide();
                PowerPointAutoShape a = slide.AddRectangle(100, 200, 300, 400);
                PowerPointAutoShape b = slide.AddRectangle(800, 1000, 200, 200);

                PowerPointGroupShape group = slide.GroupShapes(new PowerPointShape[] { a, b });

                Assert.Single(slide.Shapes);
                Assert.Equal(100, group.Left);
                Assert.Equal(200, group.Top);
                Assert.Equal(900, group.Width);
                Assert.Equal(1000, group.Height);
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
                using PowerPointPresentation presentation = PowerPointPresentation.Create(filePath);
                PowerPointSlide slide = presentation.AddSlide();
                PowerPointAutoShape a = slide.AddRectangle(0, 0, 100, 100);
                PowerPointAutoShape b = slide.AddRectangle(200, 300, 150, 120);

                PowerPointGroupShape group = slide.GroupShapes(slide.Shapes);
                var children = slide.UngroupShape(group).ToList();

                Assert.Equal(2, children.Count);
                Assert.Equal(2, slide.Shapes.Count);

                PowerPointShape childA = children[0];
                PowerPointShape childB = children[1];

                Assert.Equal(0, childA.Left);
                Assert.Equal(0, childA.Top);
                Assert.Equal(100, childA.Width);
                Assert.Equal(100, childA.Height);

                Assert.Equal(200, childB.Left);
                Assert.Equal(300, childB.Top);
                Assert.Equal(150, childB.Width);
                Assert.Equal(120, childB.Height);
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
                using PowerPointPresentation presentation = PowerPointPresentation.Create(filePath);
                PowerPointSlide slide = presentation.AddSlide();
                PowerPointAutoShape a = slide.AddRectangle(100, 200, 300, 100);
                PowerPointAutoShape b = slide.AddRectangle(500, 250, 200, 100);

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
                using PowerPointPresentation presentation = PowerPointPresentation.Create(filePath);
                PowerPointSlide slide = presentation.AddSlide();
                slide.AddRectangle(0, 0, 1000, 500);
                slide.AddRectangle(2000, 0, 1000, 500);
                slide.AddRectangle(4000, 0, 1000, 500);

                PowerPointGroupShape group = slide.GroupShapes(slide.Shapes);
                slide.DistributeGroupChildrenWithSpacing(group, PowerPointShapeDistribution.Horizontal, spacingEmus: 500);

                PowerPointLayoutBox bounds = slide.GetGroupChildBounds(group);
                var children = slide.GetGroupChildren(group).OrderBy(s => s.Left).ToList();

                Assert.Equal(bounds.Left, children[0].Left);
                Assert.Equal(bounds.Left + 1500, children[1].Left);
                Assert.Equal(bounds.Left + 3000, children[2].Left);
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
                using PowerPointPresentation presentation = PowerPointPresentation.Create(filePath);
                PowerPointSlide slide = presentation.AddSlide();
                slide.AddRectangle(0, 0, 100, 100);
                slide.AddRectangle(200, 0, 100, 100);
                slide.AddRectangle(0, 200, 100, 100);
                slide.AddRectangle(200, 200, 100, 100);

                PowerPointGroupShape group = slide.GroupShapes(slide.Shapes);
                slide.ArrangeGroupChildrenInGrid(group, columns: 2, rows: 2);

                PowerPointLayoutBox bounds = slide.GetGroupChildBounds(group);
                long cellWidth = bounds.Width / 2;
                long cellHeight = bounds.Height / 2;
                var children = slide.GetGroupChildren(group).ToList();

                Assert.Equal(bounds.Left, children[0].Left);
                Assert.Equal(bounds.Top, children[0].Top);

                Assert.Equal(bounds.Left + cellWidth, children[1].Left);
                Assert.Equal(bounds.Top, children[1].Top);

                Assert.Equal(bounds.Left, children[2].Left);
                Assert.Equal(bounds.Top + cellHeight, children[2].Top);

                Assert.Equal(bounds.Left + cellWidth, children[3].Left);
                Assert.Equal(bounds.Top + cellHeight, children[3].Top);
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
                using PowerPointPresentation presentation = PowerPointPresentation.Create(filePath);
                PowerPointSlide slide = presentation.AddSlide();
                slide.AddTextBox("A", 0, 0, 100, 100);
                slide.AddTextBox("B", 200, 0, 100, 100);
                slide.AddRectangle(400, 0, 100, 100);

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
