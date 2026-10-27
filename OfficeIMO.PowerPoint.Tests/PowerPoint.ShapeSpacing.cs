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
                PowerPointAutoShape a = slide.AddRectangle(0, 0, 1000, 1000);
                PowerPointAutoShape b = slide.AddRectangle(0, 0, 1000, 1000);
                PowerPointAutoShape c = slide.AddRectangle(0, 0, 1000, 1000);

                slide.DistributeShapesWithSpacing(slide.Shapes, PowerPointShapeDistribution.Horizontal,
                    new PowerPointLayoutBox(OfficeIMO.OfficeLength.Emus(0),OfficeIMO.OfficeLength.Emus(0),OfficeIMO.OfficeLength.Emus(6000),OfficeIMO.OfficeLength.Emus(2000)), spacingEmus: 500);

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
                PowerPointAutoShape a = slide.AddRectangle(0, 0, 1000, 1000);
                PowerPointAutoShape b = slide.AddRectangle(0, 0, 1000, 1000);
                PowerPointAutoShape c = slide.AddRectangle(0, 0, 1000, 1000);

                slide.DistributeShapesWithSpacing(slide.Shapes, PowerPointShapeDistribution.Horizontal,
                    new PowerPointLayoutBox(OfficeIMO.OfficeLength.Emus(0),OfficeIMO.OfficeLength.Emus(0),OfficeIMO.OfficeLength.Emus(6000),OfficeIMO.OfficeLength.Emus(2000)), spacingEmus: 500, center: true);

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
                PowerPointAutoShape a = slide.AddRectangle(0, 0, 1000, 1000);
                PowerPointAutoShape b = slide.AddRectangle(0, 0, 1000, 1000);
                PowerPointAutoShape c = slide.AddRectangle(0, 0, 1000, 1000);

                slide.DistributeShapesWithSpacing(slide.Shapes, PowerPointShapeDistribution.Horizontal,
                    new PowerPointLayoutBox(OfficeIMO.OfficeLength.Emus(0),OfficeIMO.OfficeLength.Emus(0),OfficeIMO.OfficeLength.Emus(6000),OfficeIMO.OfficeLength.Emus(2000)), spacingEmus: 500, PowerPointShapeAlignment.Right);

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
                PowerPointAutoShape a = slide.AddRectangle(0, 0, 1000, 1000);
                PowerPointAutoShape b = slide.AddRectangle(0, 0, 1000, 1000);
                PowerPointAutoShape c = slide.AddRectangle(0, 0, 1000, 1000);

                slide.DistributeShapesWithSpacing(slide.Shapes, PowerPointShapeDistribution.Vertical,
                    new PowerPointLayoutBox(OfficeIMO.OfficeLength.Emus(0),OfficeIMO.OfficeLength.Emus(0),OfficeIMO.OfficeLength.Emus(2000),OfficeIMO.OfficeLength.Emus(6000)), spacingEmus: 500, PowerPointShapeAlignment.Bottom);

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
                PowerPointAutoShape a = slide.AddRectangle(0, 0, 1000, 500);
                PowerPointAutoShape b = slide.AddRectangle(0, 0, 1000, 800);
                PowerPointAutoShape c = slide.AddRectangle(0, 0, 1000, 600);

                slide.DistributeShapesWithSpacing(slide.Shapes, PowerPointShapeDistribution.Horizontal,
                    new PowerPointLayoutBox(OfficeIMO.OfficeLength.Emus(0),OfficeIMO.OfficeLength.Emus(0),OfficeIMO.OfficeLength.Emus(6000),OfficeIMO.OfficeLength.Emus(2000)), spacingEmus: 500,
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
                PowerPointAutoShape a = slide.AddRectangle(0, 0, 1000, 1000);
                PowerPointAutoShape b = slide.AddRectangle(2000, 0, 1000, 1000);
                PowerPointAutoShape c = slide.AddRectangle(4000, 0, 1000, 1000);
                long margin = PowerPointUnits.FromCentimeters(1);
                const long spacing = 500;

                slide.DistributeShapesWithSpacingToSlideContent(slide.Shapes, PowerPointShapeDistribution.Horizontal,
                    spacingEmus: spacing, marginEmus: margin);

                PowerPointLayoutBox content = presentation.SlideSize.GetContentBox(margin);
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
                PowerPointAutoShape a = slide.AddRectangle(0, 0, 2500, 500);
                PowerPointAutoShape b = slide.AddRectangle(0, 0, 2500, 500);
                PowerPointAutoShape c = slide.AddRectangle(0, 0, 2500, 500);

                slide.DistributeShapesWithSpacing(slide.Shapes, PowerPointShapeDistribution.Horizontal,
                    new PowerPointLayoutBox(OfficeIMO.OfficeLength.Emus(0),OfficeIMO.OfficeLength.Emus(0),OfficeIMO.OfficeLength.Emus(8000),OfficeIMO.OfficeLength.Emus(2000)),
                    new PowerPointShapeSpacingOptions {
                        SpacingEmus = 1000,
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
                PowerPointAutoShape a = slide.AddRectangle(0, 0, 3000, 500);
                PowerPointAutoShape b = slide.AddRectangle(0, 0, 3000, 500);

                slide.DistributeShapesWithSpacing(slide.Shapes, PowerPointShapeDistribution.Horizontal,
                    new PowerPointLayoutBox(OfficeIMO.OfficeLength.Emus(0),OfficeIMO.OfficeLength.Emus(0),OfficeIMO.OfficeLength.Emus(5000),OfficeIMO.OfficeLength.Emus(2000)),
                    new PowerPointShapeSpacingOptions {
                        SpacingEmus = 1000,
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
    }
}
