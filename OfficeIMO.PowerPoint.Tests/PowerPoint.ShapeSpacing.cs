using System;
using System.IO;
using OfficeIMO.PowerPoint;
using Xunit;

namespace OfficeIMO.Tests {
    public class PowerPointShapeSpacingTests {
        [Fact]
        public void DistributeShapesWithSpacing_HorizontalUsesFixedGap() {
            string filePath = Path.Combine(Path.GetTempPath(), Guid.NewGuid() + ".pptx");
            try {
                using PowerPointDocument presentation = PowerPointDocument.Create(filePath);
                PowerPointSlide slide = presentation.AddSlide();
                PowerPointAutoShape a = slide.AddRectangle(OfficeIMO.OfficeLength.Emus(0), OfficeIMO.OfficeLength.Emus(0), OfficeIMO.OfficeLength.Emus(1000), OfficeIMO.OfficeLength.Emus(1000));
                PowerPointAutoShape b = slide.AddRectangle(OfficeIMO.OfficeLength.Emus(0), OfficeIMO.OfficeLength.Emus(0), OfficeIMO.OfficeLength.Emus(1000), OfficeIMO.OfficeLength.Emus(1000));
                PowerPointAutoShape c = slide.AddRectangle(OfficeIMO.OfficeLength.Emus(0), OfficeIMO.OfficeLength.Emus(0), OfficeIMO.OfficeLength.Emus(1000), OfficeIMO.OfficeLength.Emus(1000));

                slide.DistributeShapesWithSpacing(slide.Shapes, PowerPointShapeDistribution.Horizontal,
                    new PowerPointLayoutBox(OfficeIMO.OfficeLength.Emus(0),OfficeIMO.OfficeLength.Emus(0),OfficeIMO.OfficeLength.Emus(6000),OfficeIMO.OfficeLength.Emus(2000)), spacing: OfficeIMO.OfficeLength.Emus(500));

                Assert.Equal(0, a.Left.ToEmus());
                Assert.Equal(1500, b.Left.ToEmus());
                Assert.Equal(3000, c.Left.ToEmus());
            } finally {
                if (File.Exists(filePath)) {
                    File.Delete(filePath);
                }
            }
        }

        [Fact]
        public void DistributeShapesWithSpacing_CentersWhenRequested() {
            string filePath = Path.Combine(Path.GetTempPath(), Guid.NewGuid() + ".pptx");
            try {
                using PowerPointDocument presentation = PowerPointDocument.Create(filePath);
                PowerPointSlide slide = presentation.AddSlide();
                PowerPointAutoShape a = slide.AddRectangle(OfficeIMO.OfficeLength.Emus(0), OfficeIMO.OfficeLength.Emus(0), OfficeIMO.OfficeLength.Emus(1000), OfficeIMO.OfficeLength.Emus(1000));
                PowerPointAutoShape b = slide.AddRectangle(OfficeIMO.OfficeLength.Emus(0), OfficeIMO.OfficeLength.Emus(0), OfficeIMO.OfficeLength.Emus(1000), OfficeIMO.OfficeLength.Emus(1000));
                PowerPointAutoShape c = slide.AddRectangle(OfficeIMO.OfficeLength.Emus(0), OfficeIMO.OfficeLength.Emus(0), OfficeIMO.OfficeLength.Emus(1000), OfficeIMO.OfficeLength.Emus(1000));

                slide.DistributeShapesWithSpacing(slide.Shapes, PowerPointShapeDistribution.Horizontal,
                    new PowerPointLayoutBox(OfficeIMO.OfficeLength.Emus(0),OfficeIMO.OfficeLength.Emus(0),OfficeIMO.OfficeLength.Emus(6000),OfficeIMO.OfficeLength.Emus(2000)), spacing: OfficeIMO.OfficeLength.Emus(500), center: true);

                Assert.Equal(1000, a.Left.ToEmus());
                Assert.Equal(2500, b.Left.ToEmus());
                Assert.Equal(4000, c.Left.ToEmus());
            } finally {
                if (File.Exists(filePath)) {
                    File.Delete(filePath);
                }
            }
        }

        [Fact]
        public void DistributeShapesWithSpacing_AlignsRight() {
            string filePath = Path.Combine(Path.GetTempPath(), Guid.NewGuid() + ".pptx");
            try {
                using PowerPointDocument presentation = PowerPointDocument.Create(filePath);
                PowerPointSlide slide = presentation.AddSlide();
                PowerPointAutoShape a = slide.AddRectangle(OfficeIMO.OfficeLength.Emus(0), OfficeIMO.OfficeLength.Emus(0), OfficeIMO.OfficeLength.Emus(1000), OfficeIMO.OfficeLength.Emus(1000));
                PowerPointAutoShape b = slide.AddRectangle(OfficeIMO.OfficeLength.Emus(0), OfficeIMO.OfficeLength.Emus(0), OfficeIMO.OfficeLength.Emus(1000), OfficeIMO.OfficeLength.Emus(1000));
                PowerPointAutoShape c = slide.AddRectangle(OfficeIMO.OfficeLength.Emus(0), OfficeIMO.OfficeLength.Emus(0), OfficeIMO.OfficeLength.Emus(1000), OfficeIMO.OfficeLength.Emus(1000));

                slide.DistributeShapesWithSpacing(slide.Shapes, PowerPointShapeDistribution.Horizontal,
                    new PowerPointLayoutBox(OfficeIMO.OfficeLength.Emus(0),OfficeIMO.OfficeLength.Emus(0),OfficeIMO.OfficeLength.Emus(6000),OfficeIMO.OfficeLength.Emus(2000)), spacing: OfficeIMO.OfficeLength.Emus(500), PowerPointShapeAlignment.Right);

                Assert.Equal(2000, a.Left.ToEmus());
                Assert.Equal(3500, b.Left.ToEmus());
                Assert.Equal(5000, c.Left.ToEmus());
            } finally {
                if (File.Exists(filePath)) {
                    File.Delete(filePath);
                }
            }
        }

        [Fact]
        public void DistributeShapesWithSpacing_AlignsBottom() {
            string filePath = Path.Combine(Path.GetTempPath(), Guid.NewGuid() + ".pptx");
            try {
                using PowerPointDocument presentation = PowerPointDocument.Create(filePath);
                PowerPointSlide slide = presentation.AddSlide();
                PowerPointAutoShape a = slide.AddRectangle(OfficeIMO.OfficeLength.Emus(0), OfficeIMO.OfficeLength.Emus(0), OfficeIMO.OfficeLength.Emus(1000), OfficeIMO.OfficeLength.Emus(1000));
                PowerPointAutoShape b = slide.AddRectangle(OfficeIMO.OfficeLength.Emus(0), OfficeIMO.OfficeLength.Emus(0), OfficeIMO.OfficeLength.Emus(1000), OfficeIMO.OfficeLength.Emus(1000));
                PowerPointAutoShape c = slide.AddRectangle(OfficeIMO.OfficeLength.Emus(0), OfficeIMO.OfficeLength.Emus(0), OfficeIMO.OfficeLength.Emus(1000), OfficeIMO.OfficeLength.Emus(1000));

                slide.DistributeShapesWithSpacing(slide.Shapes, PowerPointShapeDistribution.Vertical,
                    new PowerPointLayoutBox(OfficeIMO.OfficeLength.Emus(0),OfficeIMO.OfficeLength.Emus(0),OfficeIMO.OfficeLength.Emus(2000),OfficeIMO.OfficeLength.Emus(6000)), spacing: OfficeIMO.OfficeLength.Emus(500), PowerPointShapeAlignment.Bottom);

                Assert.Equal(2000, a.Top.ToEmus());
                Assert.Equal(3500, b.Top.ToEmus());
                Assert.Equal(5000, c.Top.ToEmus());
            } finally {
                if (File.Exists(filePath)) {
                    File.Delete(filePath);
                }
            }
        }

        [Fact]
        public void DistributeShapesWithSpacing_AlignsBottomOnCrossAxis() {
            string filePath = Path.Combine(Path.GetTempPath(), Guid.NewGuid() + ".pptx");
            try {
                using PowerPointDocument presentation = PowerPointDocument.Create(filePath);
                PowerPointSlide slide = presentation.AddSlide();
                PowerPointAutoShape a = slide.AddRectangle(OfficeIMO.OfficeLength.Emus(0), OfficeIMO.OfficeLength.Emus(0), OfficeIMO.OfficeLength.Emus(1000), OfficeIMO.OfficeLength.Emus(500));
                PowerPointAutoShape b = slide.AddRectangle(OfficeIMO.OfficeLength.Emus(0), OfficeIMO.OfficeLength.Emus(0), OfficeIMO.OfficeLength.Emus(1000), OfficeIMO.OfficeLength.Emus(800));
                PowerPointAutoShape c = slide.AddRectangle(OfficeIMO.OfficeLength.Emus(0), OfficeIMO.OfficeLength.Emus(0), OfficeIMO.OfficeLength.Emus(1000), OfficeIMO.OfficeLength.Emus(600));

                slide.DistributeShapesWithSpacing(slide.Shapes, PowerPointShapeDistribution.Horizontal,
                    new PowerPointLayoutBox(OfficeIMO.OfficeLength.Emus(0),OfficeIMO.OfficeLength.Emus(0),OfficeIMO.OfficeLength.Emus(6000),OfficeIMO.OfficeLength.Emus(2000)), spacing: OfficeIMO.OfficeLength.Emus(500),
                    PowerPointShapeAlignment.Right, PowerPointShapeAlignment.Bottom);

                Assert.Equal(2000, a.Left.ToEmus());
                Assert.Equal(3500, b.Left.ToEmus());
                Assert.Equal(5000, c.Left.ToEmus());

                var bounds = new PowerPointLayoutBox(OfficeIMO.OfficeLength.Emus(0),OfficeIMO.OfficeLength.Emus(0),OfficeIMO.OfficeLength.Emus(6000),OfficeIMO.OfficeLength.Emus(2000));
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
        public void DistributeShapesWithSpacingToSlideContent_UsesMargin() {
            string filePath = Path.Combine(Path.GetTempPath(), Guid.NewGuid() + ".pptx");
            try {
                using PowerPointDocument presentation = PowerPointDocument.Create(filePath);
                PowerPointSlide slide = presentation.AddSlide();
                PowerPointAutoShape a = slide.AddRectangle(OfficeIMO.OfficeLength.Emus(0), OfficeIMO.OfficeLength.Emus(0), OfficeIMO.OfficeLength.Emus(1000), OfficeIMO.OfficeLength.Emus(1000));
                PowerPointAutoShape b = slide.AddRectangle(OfficeIMO.OfficeLength.Emus(2000), OfficeIMO.OfficeLength.Emus(0), OfficeIMO.OfficeLength.Emus(1000), OfficeIMO.OfficeLength.Emus(1000));
                PowerPointAutoShape c = slide.AddRectangle(OfficeIMO.OfficeLength.Emus(4000), OfficeIMO.OfficeLength.Emus(0), OfficeIMO.OfficeLength.Emus(1000), OfficeIMO.OfficeLength.Emus(1000));
                long margin = global::OfficeIMO.OfficeLength.Centimeters(1).ToEmus();
                const long spacing = 500;

                slide.DistributeShapesWithSpacingToSlideContent(slide.Shapes, PowerPointShapeDistribution.Horizontal,
                    spacing: OfficeIMO.OfficeLength.Emus(spacing), margin: OfficeIMO.OfficeLength.Emus(margin));

                PowerPointLayoutBox content = presentation.SlideSize.GetContentBox(OfficeIMO.OfficeLength.Emus(margin));
                double current = content.Left.ToEmus();
                long expectedA = (long)Math.Round(current);
                current += a.Width.ToEmus() + spacing;
                long expectedB = (long)Math.Round(current);
                current += b.Width.ToEmus() + spacing;
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
        public void DistributeShapesWithSpacing_ClampSpacingToBounds() {
            string filePath = Path.Combine(Path.GetTempPath(), Guid.NewGuid() + ".pptx");
            try {
                using PowerPointDocument presentation = PowerPointDocument.Create(filePath);
                PowerPointSlide slide = presentation.AddSlide();
                PowerPointAutoShape a = slide.AddRectangle(OfficeIMO.OfficeLength.Emus(0), OfficeIMO.OfficeLength.Emus(0), OfficeIMO.OfficeLength.Emus(2500), OfficeIMO.OfficeLength.Emus(500));
                PowerPointAutoShape b = slide.AddRectangle(OfficeIMO.OfficeLength.Emus(0), OfficeIMO.OfficeLength.Emus(0), OfficeIMO.OfficeLength.Emus(2500), OfficeIMO.OfficeLength.Emus(500));
                PowerPointAutoShape c = slide.AddRectangle(OfficeIMO.OfficeLength.Emus(0), OfficeIMO.OfficeLength.Emus(0), OfficeIMO.OfficeLength.Emus(2500), OfficeIMO.OfficeLength.Emus(500));

                slide.DistributeShapesWithSpacing(slide.Shapes, PowerPointShapeDistribution.Horizontal,
                    new PowerPointLayoutBox(OfficeIMO.OfficeLength.Emus(0),OfficeIMO.OfficeLength.Emus(0),OfficeIMO.OfficeLength.Emus(8000),OfficeIMO.OfficeLength.Emus(2000)),
                    new PowerPointShapeSpacingOptions {
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
        public void DistributeShapesWithSpacing_ScalesToFitBounds() {
            string filePath = Path.Combine(Path.GetTempPath(), Guid.NewGuid() + ".pptx");
            try {
                using PowerPointDocument presentation = PowerPointDocument.Create(filePath);
                PowerPointSlide slide = presentation.AddSlide();
                PowerPointAutoShape a = slide.AddRectangle(OfficeIMO.OfficeLength.Emus(0), OfficeIMO.OfficeLength.Emus(0), OfficeIMO.OfficeLength.Emus(3000), OfficeIMO.OfficeLength.Emus(500));
                PowerPointAutoShape b = slide.AddRectangle(OfficeIMO.OfficeLength.Emus(0), OfficeIMO.OfficeLength.Emus(0), OfficeIMO.OfficeLength.Emus(3000), OfficeIMO.OfficeLength.Emus(500));

                slide.DistributeShapesWithSpacing(slide.Shapes, PowerPointShapeDistribution.Horizontal,
                    new PowerPointLayoutBox(OfficeIMO.OfficeLength.Emus(0),OfficeIMO.OfficeLength.Emus(0),OfficeIMO.OfficeLength.Emus(5000),OfficeIMO.OfficeLength.Emus(2000)),
                    new PowerPointShapeSpacingOptions {
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
        public void DistributeShapesWithSpacing_RejectsCrossAxisBeforeScalingOrMoving() {
            using PowerPointDocument presentation = PowerPointDocument.Create();
            PowerPointSlide slide = presentation.AddSlide();
            PowerPointAutoShape first = slide.AddRectangle(OfficeIMO.OfficeLength.Emus(100), OfficeIMO.OfficeLength.Emus(200),
                OfficeIMO.OfficeLength.Emus(3000), OfficeIMO.OfficeLength.Emus(500));
            PowerPointAutoShape second = slide.AddRectangle(OfficeIMO.OfficeLength.Emus(4000), OfficeIMO.OfficeLength.Emus(200),
                OfficeIMO.OfficeLength.Emus(3000), OfficeIMO.OfficeLength.Emus(500));
            var bounds = new PowerPointLayoutBox(OfficeIMO.OfficeLength.Zero, OfficeIMO.OfficeLength.Zero,
                OfficeIMO.OfficeLength.Emus(5000), OfficeIMO.OfficeLength.Emus(2000));

            Assert.Throws<ArgumentOutOfRangeException>(() => slide.DistributeShapesWithSpacing(slide.Shapes,
                PowerPointShapeDistribution.Horizontal, bounds, new PowerPointShapeSpacingOptions {
                    Spacing = OfficeIMO.OfficeLength.Emus(1000),
                    ScaleToFitBounds = true,
                    CrossAxisAlignment = PowerPointShapeAlignment.Left
                }));
            Assert.Throws<ArgumentOutOfRangeException>(() => slide.DistributeShapesWithSpacing(slide.Shapes,
                PowerPointShapeDistribution.Horizontal, bounds, OfficeIMO.OfficeLength.Points(1),
                PowerPointShapeAlignment.Left, PowerPointShapeAlignment.Left));
            Assert.Equal(100, first.Left.ToEmus());
            Assert.Equal(4000, second.Left.ToEmus());
            Assert.Equal(3000, first.Width.ToEmus());
            Assert.Equal(3000, second.Width.ToEmus());
        }

        [Fact]
        public void DistributeShapesWithSpacing_NullOptionsUseZeroGap() {
            using PowerPointDocument presentation = PowerPointDocument.Create();
            PowerPointSlide slide = presentation.AddSlide();
            PowerPointAutoShape first = slide.AddRectangle(OfficeIMO.OfficeLength.Zero, OfficeIMO.OfficeLength.Zero,
                OfficeIMO.OfficeLength.Points(10), OfficeIMO.OfficeLength.Points(10));
            PowerPointAutoShape second = slide.AddRectangle(OfficeIMO.OfficeLength.Points(20), OfficeIMO.OfficeLength.Zero,
                OfficeIMO.OfficeLength.Points(10), OfficeIMO.OfficeLength.Points(10));
            var bounds = new PowerPointLayoutBox(OfficeIMO.OfficeLength.Zero, OfficeIMO.OfficeLength.Zero,
                OfficeIMO.OfficeLength.Inches(1), OfficeIMO.OfficeLength.Inches(1));

            slide.DistributeShapesWithSpacing(slide.Shapes, PowerPointShapeDistribution.Horizontal, bounds,
                (PowerPointShapeSpacingOptions?)null);
            Assert.Equal(OfficeIMO.OfficeLength.Zero, first.Left);
            Assert.Equal(first.Width, second.Left);
        }

    }
}
