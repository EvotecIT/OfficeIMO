using System;
using System.IO;
using OfficeIMO.PowerPoint;
using Xunit;

namespace OfficeIMO.Tests {
    public class PowerPointShapeStackTests {
        [Fact]
        public void StackShapes_Horizontal_AlignsBottom() {
            string filePath = Path.Combine(Path.GetTempPath(), Guid.NewGuid() + ".pptx");
            try {
                using PowerPointDocument presentation = PowerPointDocument.Create(filePath);
                PowerPointSlide slide = presentation.AddSlide();
                PowerPointAutoShape a = slide.AddRectangle(OfficeIMO.OfficeLength.Emus(0), OfficeIMO.OfficeLength.Emus(0), OfficeIMO.OfficeLength.Emus(1000), OfficeIMO.OfficeLength.Emus(500));
                PowerPointAutoShape b = slide.AddRectangle(OfficeIMO.OfficeLength.Emus(0), OfficeIMO.OfficeLength.Emus(0), OfficeIMO.OfficeLength.Emus(2000), OfficeIMO.OfficeLength.Emus(1000));

                var bounds = new PowerPointLayoutBox(OfficeIMO.OfficeLength.Emus(100),OfficeIMO.OfficeLength.Emus(200),OfficeIMO.OfficeLength.Emus(8000),OfficeIMO.OfficeLength.Emus(3000));
                slide.StackShapes(slide.Shapes, PowerPointShapeStackDirection.Horizontal, bounds,
                    spacing: OfficeIMO.OfficeLength.Emus(300), PowerPointShapeAlignment.Bottom);

                Assert.Equal(100, a.Left.ToEmus());
                Assert.Equal(100 + 1000 + 300, b.Left.ToEmus());
                Assert.Equal(bounds.Bottom.ToEmus() - a.Height.ToEmus(), a.Top.ToEmus());
                Assert.Equal(bounds.Bottom.ToEmus() - b.Height.ToEmus(), b.Top.ToEmus());
            } finally {
                if (File.Exists(filePath)) {
                    File.Delete(filePath);
                }
            }
        }

        [Fact]
        public void StackShapes_Vertical_AlignsCenter() {
            string filePath = Path.Combine(Path.GetTempPath(), Guid.NewGuid() + ".pptx");
            try {
                using PowerPointDocument presentation = PowerPointDocument.Create(filePath);
                PowerPointSlide slide = presentation.AddSlide();
                PowerPointAutoShape a = slide.AddRectangle(OfficeIMO.OfficeLength.Emus(0), OfficeIMO.OfficeLength.Emus(0), OfficeIMO.OfficeLength.Emus(1000), OfficeIMO.OfficeLength.Emus(600));
                PowerPointAutoShape b = slide.AddRectangle(OfficeIMO.OfficeLength.Emus(0), OfficeIMO.OfficeLength.Emus(0), OfficeIMO.OfficeLength.Emus(2000), OfficeIMO.OfficeLength.Emus(800));

                var bounds = new PowerPointLayoutBox(OfficeIMO.OfficeLength.Emus(200),OfficeIMO.OfficeLength.Emus(300),OfficeIMO.OfficeLength.Emus(6000),OfficeIMO.OfficeLength.Emus(4000));
                slide.StackShapes(slide.Shapes, PowerPointShapeStackDirection.Vertical, bounds,
                    spacing: OfficeIMO.OfficeLength.Emus(250), PowerPointShapeAlignment.Center);

                Assert.Equal(300, a.Top.ToEmus());
                Assert.Equal(300 + 600 + 250, b.Top.ToEmus());
                Assert.Equal(bounds.Left.ToEmus() + (bounds.Width.ToEmus() - a.Width.ToEmus()) / 2, a.Left.ToEmus());
                Assert.Equal(bounds.Left.ToEmus() + (bounds.Width.ToEmus() - b.Width.ToEmus()) / 2, b.Left.ToEmus());
            } finally {
                if (File.Exists(filePath)) {
                    File.Delete(filePath);
                }
            }
        }

        [Fact]
        public void StackShapesToSlideContent_UsesMargin() {
            string filePath = Path.Combine(Path.GetTempPath(), Guid.NewGuid() + ".pptx");
            try {
                using PowerPointDocument presentation = PowerPointDocument.Create(filePath);
                PowerPointSlide slide = presentation.AddSlide();
                PowerPointAutoShape a = slide.AddRectangle(OfficeIMO.OfficeLength.Emus(0), OfficeIMO.OfficeLength.Emus(0), OfficeIMO.OfficeLength.Emus(1000), OfficeIMO.OfficeLength.Emus(700));
                PowerPointAutoShape b = slide.AddRectangle(OfficeIMO.OfficeLength.Emus(0), OfficeIMO.OfficeLength.Emus(0), OfficeIMO.OfficeLength.Emus(1000), OfficeIMO.OfficeLength.Emus(900));
                long margin = PowerPointUnits.FromCentimeters(1);

                slide.StackShapesToSlideContent(slide.Shapes, PowerPointShapeStackDirection.Vertical,
                    spacing: OfficeIMO.OfficeLength.Emus(200), margin: OfficeIMO.OfficeLength.Emus(margin));

                PowerPointLayoutBox content = presentation.SlideSize.GetContentBox(OfficeIMO.OfficeLength.Emus(margin));
                Assert.Equal(content.Top.ToEmus(), a.Top.ToEmus());
                Assert.Equal(content.Top.ToEmus() + a.Height.ToEmus() + 200, b.Top.ToEmus());
                Assert.Equal(content.Left.ToEmus(), a.Left.ToEmus());
                Assert.Equal(content.Left.ToEmus(), b.Left.ToEmus());
            } finally {
                if (File.Exists(filePath)) {
                    File.Delete(filePath);
                }
            }
        }

        [Fact]
        public void StackShapes_Horizontal_JustifiesCenter() {
            string filePath = Path.Combine(Path.GetTempPath(), Guid.NewGuid() + ".pptx");
            try {
                using PowerPointDocument presentation = PowerPointDocument.Create(filePath);
                PowerPointSlide slide = presentation.AddSlide();
                PowerPointAutoShape a = slide.AddRectangle(OfficeIMO.OfficeLength.Emus(0), OfficeIMO.OfficeLength.Emus(0), OfficeIMO.OfficeLength.Emus(1000), OfficeIMO.OfficeLength.Emus(500));
                PowerPointAutoShape b = slide.AddRectangle(OfficeIMO.OfficeLength.Emus(0), OfficeIMO.OfficeLength.Emus(0), OfficeIMO.OfficeLength.Emus(2000), OfficeIMO.OfficeLength.Emus(500));

                var bounds = new PowerPointLayoutBox(OfficeIMO.OfficeLength.Emus(0),OfficeIMO.OfficeLength.Emus(0),OfficeIMO.OfficeLength.Emus(8000),OfficeIMO.OfficeLength.Emus(2000));
                slide.StackShapes(slide.Shapes, PowerPointShapeStackDirection.Horizontal, bounds,
                    spacing: OfficeIMO.OfficeLength.Emus(500), PowerPointShapeStackJustify.Center);

                long totalWidth = a.Width.ToEmus() + b.Width.ToEmus() + 500;
                long expectedLeft = (bounds.Width.ToEmus() - totalWidth) / 2;

                Assert.Equal(expectedLeft, a.Left.ToEmus());
                Assert.Equal(expectedLeft + a.Width.ToEmus() + 500, b.Left.ToEmus());
            } finally {
                if (File.Exists(filePath)) {
                    File.Delete(filePath);
                }
            }
        }

        [Fact]
        public void StackShapes_ClampSpacingToBounds() {
            string filePath = Path.Combine(Path.GetTempPath(), Guid.NewGuid() + ".pptx");
            try {
                using PowerPointDocument presentation = PowerPointDocument.Create(filePath);
                PowerPointSlide slide = presentation.AddSlide();
                PowerPointAutoShape a = slide.AddRectangle(OfficeIMO.OfficeLength.Emus(0), OfficeIMO.OfficeLength.Emus(0), OfficeIMO.OfficeLength.Emus(2500), OfficeIMO.OfficeLength.Emus(500));
                PowerPointAutoShape b = slide.AddRectangle(OfficeIMO.OfficeLength.Emus(0), OfficeIMO.OfficeLength.Emus(0), OfficeIMO.OfficeLength.Emus(2500), OfficeIMO.OfficeLength.Emus(500));
                PowerPointAutoShape c = slide.AddRectangle(OfficeIMO.OfficeLength.Emus(0), OfficeIMO.OfficeLength.Emus(0), OfficeIMO.OfficeLength.Emus(2500), OfficeIMO.OfficeLength.Emus(500));

                var bounds = new PowerPointLayoutBox(OfficeIMO.OfficeLength.Emus(0),OfficeIMO.OfficeLength.Emus(0),OfficeIMO.OfficeLength.Emus(8000),OfficeIMO.OfficeLength.Emus(2000));
                slide.StackShapes(slide.Shapes, PowerPointShapeStackDirection.Horizontal, bounds,
                    new PowerPointShapeStackOptions {
                        Spacing = OfficeIMO.OfficeLength.Emus(1000),
                        ClampSpacingToBounds = true
                    });

                Assert.Equal(0, a.Left.ToEmus());
                Assert.Equal(2750, b.Left.ToEmus());
                Assert.Equal(5500, c.Left.ToEmus());
            } finally {
                if (File.Exists(filePath)) {
                    File.Delete(filePath);
                }
            }
        }

        [Fact]
        public void StackShapes_ScalesToFitBounds() {
            string filePath = Path.Combine(Path.GetTempPath(), Guid.NewGuid() + ".pptx");
            try {
                using PowerPointDocument presentation = PowerPointDocument.Create(filePath);
                PowerPointSlide slide = presentation.AddSlide();
                PowerPointAutoShape a = slide.AddRectangle(OfficeIMO.OfficeLength.Emus(0), OfficeIMO.OfficeLength.Emus(0), OfficeIMO.OfficeLength.Emus(3000), OfficeIMO.OfficeLength.Emus(500));
                PowerPointAutoShape b = slide.AddRectangle(OfficeIMO.OfficeLength.Emus(0), OfficeIMO.OfficeLength.Emus(0), OfficeIMO.OfficeLength.Emus(3000), OfficeIMO.OfficeLength.Emus(500));

                var bounds = new PowerPointLayoutBox(OfficeIMO.OfficeLength.Emus(0),OfficeIMO.OfficeLength.Emus(0),OfficeIMO.OfficeLength.Emus(5000),OfficeIMO.OfficeLength.Emus(2000));
                slide.StackShapes(slide.Shapes, PowerPointShapeStackDirection.Horizontal, bounds,
                    new PowerPointShapeStackOptions {
                        Spacing = OfficeIMO.OfficeLength.Emus(1000),
                        ScaleToFitBounds = true
                    });

                Assert.Equal(2000, a.Width.ToEmus());
                Assert.Equal(2000, b.Width.ToEmus());
                Assert.Equal(0, a.Left.ToEmus());
                Assert.Equal(3000, b.Left.ToEmus());
            } finally {
                if (File.Exists(filePath)) {
                    File.Delete(filePath);
                }
            }
        }
        [Fact]
        public void StackShapes_RejectsAlignmentBeforeScalingOrMoving() {
            using PowerPointDocument presentation = PowerPointDocument.Create();
            PowerPointSlide slide = presentation.AddSlide();
            PowerPointAutoShape first = slide.AddRectangle(OfficeIMO.OfficeLength.Emus(100), OfficeIMO.OfficeLength.Emus(200),
                OfficeIMO.OfficeLength.Emus(3000), OfficeIMO.OfficeLength.Emus(500));
            PowerPointAutoShape second = slide.AddRectangle(OfficeIMO.OfficeLength.Emus(4000), OfficeIMO.OfficeLength.Emus(200),
                OfficeIMO.OfficeLength.Emus(3000), OfficeIMO.OfficeLength.Emus(500));
            var bounds = new PowerPointLayoutBox(OfficeIMO.OfficeLength.Zero, OfficeIMO.OfficeLength.Zero,
                OfficeIMO.OfficeLength.Emus(5000), OfficeIMO.OfficeLength.Emus(2000));

            Assert.Throws<ArgumentOutOfRangeException>(() => slide.StackShapes(slide.Shapes,
                PowerPointShapeStackDirection.Horizontal, bounds, new PowerPointShapeStackOptions {
                    Spacing = OfficeIMO.OfficeLength.Emus(1000),
                    ScaleToFitBounds = true,
                    Alignment = PowerPointShapeAlignment.Left
                }));
            Assert.Equal(100, first.Left.ToEmus());
            Assert.Equal(4000, second.Left.ToEmus());
            Assert.Equal(3000, first.Width.ToEmus());
            Assert.Equal(3000, second.Width.ToEmus());
        }

    }
}
