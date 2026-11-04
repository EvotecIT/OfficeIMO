using System;
using System.IO;
using System.Linq;
using OfficeIMO.PowerPoint;
using Xunit;

namespace OfficeIMO.Tests {
    public class PowerPointShapeAlignmentTests {
        [Fact]
        public void AlignShapes_Left_UsesSelectionBounds() {
            string filePath = Path.Combine(Path.GetTempPath(), Guid.NewGuid() + ".pptx");
            try {
                using PowerPointDocument presentation = PowerPointDocument.Create(filePath);
                PowerPointSlide slide = presentation.AddSlide();
                PowerPointAutoShape a = slide.AddRectangle(OfficeIMO.OfficeLength.Emus(1000), OfficeIMO.OfficeLength.Emus(0), OfficeIMO.OfficeLength.Emus(1000), OfficeIMO.OfficeLength.Emus(1000));
                PowerPointAutoShape b = slide.AddRectangle(OfficeIMO.OfficeLength.Emus(3000), OfficeIMO.OfficeLength.Emus(0), OfficeIMO.OfficeLength.Emus(1000), OfficeIMO.OfficeLength.Emus(1000));
                PowerPointAutoShape c = slide.AddRectangle(OfficeIMO.OfficeLength.Emus(5000), OfficeIMO.OfficeLength.Emus(0), OfficeIMO.OfficeLength.Emus(1000), OfficeIMO.OfficeLength.Emus(1000));

                slide.AlignShapes(slide.Shapes, PowerPointShapeAlignment.Left);

                Assert.All(new[] { a, b, c }, shape => Assert.Equal(OfficeIMO.OfficeLength.Emus(1000), shape.Left));
            } finally {
                if (File.Exists(filePath)) {
                    File.Delete(filePath);
                }
            }
        }

        [Fact]
        public void AlignShapesToSlide_Center_HonorsSlideBounds() {
            string filePath = Path.Combine(Path.GetTempPath(), Guid.NewGuid() + ".pptx");
            try {
                using PowerPointDocument presentation = PowerPointDocument.Create(filePath);
                PowerPointSlide slide = presentation.AddSlide();
                PowerPointAutoShape shape = slide.AddRectangle(OfficeIMO.OfficeLength.Emus(0), OfficeIMO.OfficeLength.Emus(0), OfficeIMO.OfficeLength.Emus(1000000), OfficeIMO.OfficeLength.Emus(1000000));

                slide.AlignShapesToSlide(new[] { shape }, PowerPointShapeAlignment.Center);

                long expected = (presentation.SlideSize.Width.ToEmus() - shape.Width.ToEmus()) / 2;
                Assert.Equal(expected, shape.Left.ToEmus());
            } finally {
                if (File.Exists(filePath)) {
                    File.Delete(filePath);
                }
            }
        }

        [Fact]
        public void AlignShapesToSlideContent_Center_RespectsMargin() {
            string filePath = Path.Combine(Path.GetTempPath(), Guid.NewGuid() + ".pptx");
            try {
                using PowerPointDocument presentation = PowerPointDocument.Create(filePath);
                PowerPointSlide slide = presentation.AddSlide();
                PowerPointAutoShape shape = slide.AddRectangle(OfficeIMO.OfficeLength.Emus(0), OfficeIMO.OfficeLength.Emus(0), OfficeIMO.OfficeLength.Emus(1000000), OfficeIMO.OfficeLength.Emus(1000000));
                long margin = global::OfficeIMO.OfficeLength.Centimeters(1).ToEmus();

                slide.AlignShapesToSlideContent(new[] { shape }, PowerPointShapeAlignment.Center, OfficeIMO.OfficeLength.Emus(margin));

                PowerPointLayoutBox content = presentation.SlideSize.GetContentBox(OfficeIMO.OfficeLength.Emus(margin));
                long expected = content.Left.ToEmus() + (content.Width.ToEmus() - shape.Width.ToEmus()) / 2;
                Assert.Equal(expected, shape.Left.ToEmus());
            } finally {
                if (File.Exists(filePath)) {
                    File.Delete(filePath);
                }
            }
        }

        [Fact]
        public void DistributeShapes_Horizontal_EvenSpacing() {
            string filePath = Path.Combine(Path.GetTempPath(), Guid.NewGuid() + ".pptx");
            try {
                using PowerPointDocument presentation = PowerPointDocument.Create(filePath);
                PowerPointSlide slide = presentation.AddSlide();
                PowerPointAutoShape a = slide.AddRectangle(OfficeIMO.OfficeLength.Emus(0), OfficeIMO.OfficeLength.Emus(0), OfficeIMO.OfficeLength.Emus(1000), OfficeIMO.OfficeLength.Emus(1000));
                PowerPointAutoShape b = slide.AddRectangle(OfficeIMO.OfficeLength.Emus(4000), OfficeIMO.OfficeLength.Emus(0), OfficeIMO.OfficeLength.Emus(1000), OfficeIMO.OfficeLength.Emus(1000));
                PowerPointAutoShape c = slide.AddRectangle(OfficeIMO.OfficeLength.Emus(9000), OfficeIMO.OfficeLength.Emus(0), OfficeIMO.OfficeLength.Emus(1000), OfficeIMO.OfficeLength.Emus(1000));

                slide.DistributeShapes(slide.Shapes, PowerPointShapeDistribution.Horizontal);

                Assert.Equal(0, a.Left.ToEmus());
                Assert.Equal(4500, b.Left.ToEmus());
                Assert.Equal(9000, c.Left.ToEmus());
            } finally {
                if (File.Exists(filePath)) {
                    File.Delete(filePath);
                }
            }
        }

        [Fact]
        public void DistributeShapes_Horizontal_AnchorsLastShapeToBounds() {
            string filePath = Path.Combine(Path.GetTempPath(), Guid.NewGuid() + ".pptx");
            try {
                using PowerPointDocument presentation = PowerPointDocument.Create(filePath);
                PowerPointSlide slide = presentation.AddSlide();
                PowerPointAutoShape a = slide.AddRectangle(OfficeIMO.OfficeLength.Emus(0), OfficeIMO.OfficeLength.Emus(0), OfficeIMO.OfficeLength.Emus(1111), OfficeIMO.OfficeLength.Emus(1000));
                PowerPointAutoShape b = slide.AddRectangle(OfficeIMO.OfficeLength.Emus(3000), OfficeIMO.OfficeLength.Emus(0), OfficeIMO.OfficeLength.Emus(2222), OfficeIMO.OfficeLength.Emus(1000));
                PowerPointAutoShape c = slide.AddRectangle(OfficeIMO.OfficeLength.Emus(7000), OfficeIMO.OfficeLength.Emus(0), OfficeIMO.OfficeLength.Emus(3333), OfficeIMO.OfficeLength.Emus(1000));
                var bounds = new PowerPointLayoutBox(OfficeIMO.OfficeLength.Emus(100),OfficeIMO.OfficeLength.Emus(0),OfficeIMO.OfficeLength.Emus(10000),OfficeIMO.OfficeLength.Emus(1000));

                slide.DistributeShapes(slide.Shapes, PowerPointShapeDistribution.Horizontal, bounds);

                Assert.Equal(bounds.Left.ToEmus(), a.Left.ToEmus());
                Assert.Equal(bounds.Right.ToEmus(), c.Left.ToEmus() + c.Width.ToEmus());
            } finally {
                if (File.Exists(filePath)) {
                    File.Delete(filePath);
                }
            }
        }

        [Fact]
        public void DistributeShapes_Horizontal_AlignsBottom() {
            string filePath = Path.Combine(Path.GetTempPath(), Guid.NewGuid() + ".pptx");
            try {
                using PowerPointDocument presentation = PowerPointDocument.Create(filePath);
                PowerPointSlide slide = presentation.AddSlide();
                PowerPointAutoShape a = slide.AddRectangle(OfficeIMO.OfficeLength.Emus(0), OfficeIMO.OfficeLength.Emus(0), OfficeIMO.OfficeLength.Emus(1000), OfficeIMO.OfficeLength.Emus(500));
                PowerPointAutoShape b = slide.AddRectangle(OfficeIMO.OfficeLength.Emus(2000), OfficeIMO.OfficeLength.Emus(200), OfficeIMO.OfficeLength.Emus(1000), OfficeIMO.OfficeLength.Emus(700));
                PowerPointAutoShape c = slide.AddRectangle(OfficeIMO.OfficeLength.Emus(4000), OfficeIMO.OfficeLength.Emus(400), OfficeIMO.OfficeLength.Emus(1000), OfficeIMO.OfficeLength.Emus(900));
                var bounds = new PowerPointLayoutBox(OfficeIMO.OfficeLength.Emus(0),OfficeIMO.OfficeLength.Emus(0),OfficeIMO.OfficeLength.Emus(6000),OfficeIMO.OfficeLength.Emus(2000));

                slide.DistributeShapes(slide.Shapes, PowerPointShapeDistribution.Horizontal, bounds,
                    PowerPointShapeAlignment.Bottom);

                Assert.Equal(0, a.Left.ToEmus());
                Assert.Equal(2500, b.Left.ToEmus());
                Assert.Equal(5000, c.Left.ToEmus());

                Assert.Equal(bounds.Bottom.ToEmus() - a.Height.ToEmus(), a.Top.ToEmus());
                Assert.Equal(bounds.Bottom.ToEmus() - b.Height.ToEmus(), b.Top.ToEmus());
                Assert.Equal(bounds.Bottom.ToEmus() - c.Height.ToEmus(), c.Top.ToEmus());
            } finally {
                if (File.Exists(filePath)) {
                    File.Delete(filePath);
                }
            }
        }

        [Fact]
        public void DistributeShapesToSlideContent_Horizontal_UsesMargin() {
            string filePath = Path.Combine(Path.GetTempPath(), Guid.NewGuid() + ".pptx");
            try {
                using PowerPointDocument presentation = PowerPointDocument.Create(filePath);
                PowerPointSlide slide = presentation.AddSlide();
                PowerPointAutoShape a = slide.AddRectangle(OfficeIMO.OfficeLength.Emus(0), OfficeIMO.OfficeLength.Emus(0), OfficeIMO.OfficeLength.Emus(1000), OfficeIMO.OfficeLength.Emus(1000));
                PowerPointAutoShape b = slide.AddRectangle(OfficeIMO.OfficeLength.Emus(2000), OfficeIMO.OfficeLength.Emus(0), OfficeIMO.OfficeLength.Emus(1000), OfficeIMO.OfficeLength.Emus(1000));
                PowerPointAutoShape c = slide.AddRectangle(OfficeIMO.OfficeLength.Emus(4000), OfficeIMO.OfficeLength.Emus(0), OfficeIMO.OfficeLength.Emus(1000), OfficeIMO.OfficeLength.Emus(1000));
                long margin = global::OfficeIMO.OfficeLength.Centimeters(1).ToEmus();

                slide.DistributeShapesToSlideContent(slide.Shapes, PowerPointShapeDistribution.Horizontal, OfficeIMO.OfficeLength.Emus(margin));

                PowerPointLayoutBox content = presentation.SlideSize.GetContentBox(OfficeIMO.OfficeLength.Emus(margin));
                long totalWidth = a.Width.ToEmus() + b.Width.ToEmus() + c.Width.ToEmus();
                double gap = (content.Width.ToEmus() - totalWidth) / 2d;
                double current = content.Left.ToEmus();
                long expectedA = (long)Math.Round(current);
                current += a.Width.ToEmus() + gap;
                long expectedB = (long)Math.Round(current);
                current += b.Width.ToEmus() + gap;
                long expectedC = (long)Math.Round(current);

                Assert.Equal(expectedA, a.Left.ToEmus());
                Assert.Equal(expectedB, b.Left.ToEmus());
                Assert.Equal(expectedC, c.Left.ToEmus());
            } finally {
                if (File.Exists(filePath)) {
                    File.Delete(filePath);
                }
            }
        }

        [Fact]
        public void DistributeShapes_Horizontal_AllowsOverlapToFitBounds() {
            string filePath = Path.Combine(Path.GetTempPath(), Guid.NewGuid() + ".pptx");
            try {
                using PowerPointDocument presentation = PowerPointDocument.Create(filePath);
                PowerPointSlide slide = presentation.AddSlide();
                PowerPointAutoShape a = slide.AddRectangle(OfficeIMO.OfficeLength.Emus(0), OfficeIMO.OfficeLength.Emus(0), OfficeIMO.OfficeLength.Emus(1000), OfficeIMO.OfficeLength.Emus(1000));
                PowerPointAutoShape b = slide.AddRectangle(OfficeIMO.OfficeLength.Emus(2000), OfficeIMO.OfficeLength.Emus(0), OfficeIMO.OfficeLength.Emus(1000), OfficeIMO.OfficeLength.Emus(1000));
                var bounds = new PowerPointLayoutBox(OfficeIMO.OfficeLength.Emus(0),OfficeIMO.OfficeLength.Emus(0),OfficeIMO.OfficeLength.Emus(1500),OfficeIMO.OfficeLength.Emus(1000));

                slide.DistributeShapes(slide.Shapes, PowerPointShapeDistribution.Horizontal, bounds);

                Assert.Equal(0, a.Left.ToEmus());
                Assert.Equal(500, b.Left.ToEmus());
                Assert.Equal(bounds.Right.ToEmus(), b.Left.ToEmus() + b.Width.ToEmus());
            } finally {
                if (File.Exists(filePath)) {
                    File.Delete(filePath);
                }
            }
        }

        [Fact]
        public void DistributeShapes_Vertical_AllowsOverlapToFitBounds() {
            string filePath = Path.Combine(Path.GetTempPath(), Guid.NewGuid() + ".pptx");
            try {
                using PowerPointDocument presentation = PowerPointDocument.Create(filePath);
                PowerPointSlide slide = presentation.AddSlide();
                PowerPointAutoShape a = slide.AddRectangle(OfficeIMO.OfficeLength.Emus(0), OfficeIMO.OfficeLength.Emus(0), OfficeIMO.OfficeLength.Emus(1000), OfficeIMO.OfficeLength.Emus(1000));
                PowerPointAutoShape b = slide.AddRectangle(OfficeIMO.OfficeLength.Emus(0), OfficeIMO.OfficeLength.Emus(2000), OfficeIMO.OfficeLength.Emus(1000), OfficeIMO.OfficeLength.Emus(1000));
                var bounds = new PowerPointLayoutBox(OfficeIMO.OfficeLength.Emus(0),OfficeIMO.OfficeLength.Emus(0),OfficeIMO.OfficeLength.Emus(1000),OfficeIMO.OfficeLength.Emus(1500));

                slide.DistributeShapes(slide.Shapes, PowerPointShapeDistribution.Vertical, bounds);

                Assert.Equal(0, a.Top.ToEmus());
                Assert.Equal(500, b.Top.ToEmus());
                Assert.Equal(bounds.Bottom.ToEmus(), b.Top.ToEmus() + b.Height.ToEmus());
            } finally {
                if (File.Exists(filePath)) {
                    File.Delete(filePath);
                }
            }
        }
        [Fact]
        public void DistributeShapes_RejectsCrossAxisBeforeMoving() {
            using PowerPointDocument presentation = PowerPointDocument.Create();
            PowerPointSlide slide = presentation.AddSlide();
            PowerPointAutoShape first = slide.AddRectangle(OfficeIMO.OfficeLength.Emus(100), OfficeIMO.OfficeLength.Emus(200),
                OfficeIMO.OfficeLength.Emus(1000), OfficeIMO.OfficeLength.Emus(500));
            PowerPointAutoShape second = slide.AddRectangle(OfficeIMO.OfficeLength.Emus(2000), OfficeIMO.OfficeLength.Emus(200),
                OfficeIMO.OfficeLength.Emus(1000), OfficeIMO.OfficeLength.Emus(500));
            var bounds = new PowerPointLayoutBox(OfficeIMO.OfficeLength.Zero, OfficeIMO.OfficeLength.Zero,
                OfficeIMO.OfficeLength.Emus(5000), OfficeIMO.OfficeLength.Emus(2000));

            Assert.Throws<ArgumentOutOfRangeException>(() => slide.DistributeShapes(slide.Shapes,
                PowerPointShapeDistribution.Horizontal, bounds, PowerPointShapeAlignment.Left));
            Assert.Equal(100, first.Left.ToEmus());
            Assert.Equal(2000, second.Left.ToEmus());
            Assert.Equal(200, first.Top.ToEmus());
            Assert.Equal(200, second.Top.ToEmus());
        }

    }
}
