using System;
using System.IO;
using OfficeIMO.PowerPoint;
using Xunit;

namespace OfficeIMO.Tests {
    public class PowerPointShapeBoundsTests {
        [Fact]
        public void SlidePlacement_PreservesMixedUnitsAndNativeDefaultsOnReload() {
            using var stream = new MemoryStream();
            using (PowerPointDocument presentation = PowerPointDocument.Create(stream)) {
                PowerPointSlide slide = presentation.AddSlide(PowerPointSlideLayoutType.Blank);
                slide.AddRectangle(
                    OfficeLength.Centimeters(1), OfficeLength.Inches(1),
                    OfficeLength.Points(72), OfficeLength.Millimeters(25.4));
                slide.AddShape(OfficeIMO.Drawing.OfficePresetShapeType.Rectangle,
                    left: null, top: OfficeLength.Points(18), width: null, height: null);
                slide.AddTextShape(OfficeIMO.Drawing.OfficePresetShapeType.Ellipse,
                    "Mixed units", width: OfficeLength.Points(144));
                slide.AddSmartArt(PowerPointSmartArtType.BasicProcess,
                    top: OfficeLength.Points(36));
                slide.AddLinkedAudio(new Uri("https://example.test/audio.wav"), width: null);
                slide.AddLinkedVideo(new Uri("https://example.test/video.mp4"), height: null);
                presentation.Save();
            }

            using var saved = new MemoryStream(stream.ToArray());
            using PowerPointDocument reloaded = PowerPointDocument.Load(saved);
            var shapes = reloaded.Slides[0].Shapes;
            Assert.Equal(6, shapes.Count);
            Assert.Equal(360000L, shapes[0].Left.ToEmus());
            Assert.Equal(914400L, shapes[0].Top.ToEmus());
            Assert.Equal(914400L, shapes[0].Width.ToEmus());
            Assert.Equal(914400L, shapes[0].Height.ToEmus());
            Assert.Equal(0L, shapes[1].Left.ToEmus());
            Assert.Equal(228600L, shapes[1].Top.ToEmus());
            Assert.Equal(914400L, shapes[1].Width.ToEmus());
            Assert.Equal(914400L, shapes[1].Height.ToEmus());
            Assert.Equal(0L, shapes[2].Left.ToEmus());
            Assert.Equal(0L, shapes[2].Top.ToEmus());
            Assert.Equal(1828800L, shapes[2].Width.ToEmus());
            Assert.Equal(914400L, shapes[2].Height.ToEmus());
            Assert.Equal(0L, shapes[3].Left.ToEmus());
            Assert.Equal(457200L, shapes[3].Top.ToEmus());
            Assert.Equal(5486400L, shapes[3].Width.ToEmus());
            Assert.Equal(3200400L, shapes[3].Height.ToEmus());
            Assert.Equal(0L, shapes[4].Left.ToEmus());
            Assert.Equal(0L, shapes[4].Top.ToEmus());
            Assert.Equal(914400L, shapes[4].Width.ToEmus());
            Assert.Equal(914400L, shapes[4].Height.ToEmus());
            Assert.Equal(0L, shapes[5].Left.ToEmus());
            Assert.Equal(0L, shapes[5].Top.ToEmus());
            Assert.Equal(3657600L, shapes[5].Width.ToEmus());
            Assert.Equal(2057400L, shapes[5].Height.ToEmus());
        }

        [Fact]
        public void ShapeBounds_GetSet_Works() {
            string filePath = Path.Combine(Path.GetTempPath(), Guid.NewGuid() + ".pptx");
            try {
                using PowerPointDocument presentation = PowerPointDocument.Create(filePath);
                PowerPointSlide slide = presentation.AddSlide();
                PowerPointAutoShape shape = slide.AddRectangle(OfficeIMO.OfficeLength.Emus(1000), OfficeIMO.OfficeLength.Emus(2000), OfficeIMO.OfficeLength.Emus(3000), OfficeIMO.OfficeLength.Emus(4000));

                PowerPointLayoutBox bounds = shape.Bounds;
                Assert.Equal(1000, bounds.Left.ToEmus());
                Assert.Equal(2000, bounds.Top.ToEmus());
                Assert.Equal(3000, bounds.Width.ToEmus());
                Assert.Equal(4000, bounds.Height.ToEmus());

                shape.Bounds = new PowerPointLayoutBox(OfficeIMO.OfficeLength.Emus(2000),OfficeIMO.OfficeLength.Emus(3000),OfficeIMO.OfficeLength.Emus(4000),OfficeIMO.OfficeLength.Emus(5000));
                Assert.Equal(2000, shape.Left.ToEmus());
                Assert.Equal(3000, shape.Top.ToEmus());
                Assert.Equal(4000, shape.Width.ToEmus());
                Assert.Equal(5000, shape.Height.ToEmus());
            } finally {
                if (File.Exists(filePath)) {
                    File.Delete(filePath);
                }
            }
        }

        [Fact]
        public void ShapeBounds_RightBottomSetters_MoveEdges() {
            string filePath = Path.Combine(Path.GetTempPath(), Guid.NewGuid() + ".pptx");
            try {
                using PowerPointDocument presentation = PowerPointDocument.Create(filePath);
                PowerPointSlide slide = presentation.AddSlide();
                PowerPointAutoShape shape = slide.AddRectangle(OfficeIMO.OfficeLength.Emus(1000), OfficeIMO.OfficeLength.Emus(1000), OfficeIMO.OfficeLength.Emus(500), OfficeIMO.OfficeLength.Emus(600));

                shape.Right= OfficeIMO.OfficeLength.Emus(4000);
                shape.Bottom= OfficeIMO.OfficeLength.Emus(5000);

                Assert.Equal(3500, shape.Left.ToEmus());
                Assert.Equal(4400, shape.Top.ToEmus());
            } finally {
                if (File.Exists(filePath)) {
                    File.Delete(filePath);
                }
            }
        }

        [Fact]
        public void ShapeBounds_CenterSetters_MoveCenter() {
            string filePath = Path.Combine(Path.GetTempPath(), Guid.NewGuid() + ".pptx");
            try {
                using PowerPointDocument presentation = PowerPointDocument.Create(filePath);
                PowerPointSlide slide = presentation.AddSlide();
                PowerPointAutoShape shape = slide.AddRectangle(OfficeIMO.OfficeLength.Emus(0), OfficeIMO.OfficeLength.Emus(0), OfficeIMO.OfficeLength.Emus(1000), OfficeIMO.OfficeLength.Emus(2000));

                shape.CenterX= OfficeIMO.OfficeLength.Emus(3000);
                shape.CenterY= OfficeIMO.OfficeLength.Emus(4000);

                Assert.Equal(2500, shape.Left.ToEmus());
                Assert.Equal(3000, shape.Top.ToEmus());
            } finally {
                if (File.Exists(filePath)) {
                    File.Delete(filePath);
                }
            }
        }

        [Fact]
        public void Slide_GetShapesInBounds_SelectsExpectedShapes() {
            string filePath = Path.Combine(Path.GetTempPath(), Guid.NewGuid() + ".pptx");
            try {
                using PowerPointDocument presentation = PowerPointDocument.Create(filePath);
                PowerPointSlide slide = presentation.AddSlide();
                PowerPointAutoShape a = slide.AddRectangle(OfficeIMO.OfficeLength.Emus(0), OfficeIMO.OfficeLength.Emus(0), OfficeIMO.OfficeLength.Emus(1000), OfficeIMO.OfficeLength.Emus(1000));
                PowerPointAutoShape b = slide.AddRectangle(OfficeIMO.OfficeLength.Emus(2000), OfficeIMO.OfficeLength.Emus(0), OfficeIMO.OfficeLength.Emus(1000), OfficeIMO.OfficeLength.Emus(1000));
                PowerPointAutoShape c = slide.AddRectangle(OfficeIMO.OfficeLength.Emus(4000), OfficeIMO.OfficeLength.Emus(0), OfficeIMO.OfficeLength.Emus(1000), OfficeIMO.OfficeLength.Emus(1000));

                PowerPointLayoutBox bounds = new PowerPointLayoutBox(OfficeIMO.OfficeLength.Emus(0),OfficeIMO.OfficeLength.Emus(0),OfficeIMO.OfficeLength.Emus(2500),OfficeIMO.OfficeLength.Emus(1500));
                var intersects = slide.GetShapesInBounds(bounds, includePartial: true);
                Assert.Contains(a, intersects);
                Assert.Contains(b, intersects);
                Assert.DoesNotContain(c, intersects);

                var contained = slide.GetShapesInBounds(bounds, includePartial: false);
                Assert.Contains(a, contained);
                Assert.DoesNotContain(b, contained);
                Assert.DoesNotContain(c, contained);
            } finally {
                if (File.Exists(filePath)) {
                    File.Delete(filePath);
                }
            }
        }
    }
}
