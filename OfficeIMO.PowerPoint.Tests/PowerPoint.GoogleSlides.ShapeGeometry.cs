using OfficeIMO.Drawing;
using OfficeIMO.PowerPoint;
using OfficeIMO.PowerPoint.GoogleSlides;
using Xunit;

namespace OfficeIMO.Tests;

public sealed partial class GoogleSlidesTests {
        [Fact]
        public void BatchCompiler_PreservesTextBearingPresetGeometry() {
            using PowerPointDocument presentation = PowerPointDocument.Create();
            PowerPointSlide slide = presentation.AddSlide();
            slide.AddTextShapePoints(OfficePresetShapeType.RightArrow, "Next step", 20, 30, 180, 60);
            slide.AddTextBox("Plain text",global::OfficeIMO.OfficeLength.Points(20),global::OfficeIMO.OfficeLength.Points(110),global::OfficeIMO.OfficeLength.Points(180),global::OfficeIMO.OfficeLength.Points(60));

            using MemoryStream saved = new System.IO.MemoryStream(presentation.ToBytes());
            using PowerPointDocument reloaded = PowerPointDocument.Load(saved);
            PowerPointTextBox[] sourceShapes = Assert.Single(reloaded.Slides).Shapes.OfType<PowerPointTextBox>().ToArray();
            Assert.True(Assert.Single(sourceShapes, shape => shape.Text == "Plain text").UsesTextBoxGeometry);
            Assert.False(Assert.Single(sourceShapes, shape => shape.Text == "Next step").UsesTextBoxGeometry);

            GoogleSlidesBatch batch = reloaded.BuildGoogleSlidesBatch();

            GoogleSlidesTextBox[] textShapes = Assert.Single(batch.Slides).Elements.OfType<GoogleSlidesTextBox>().ToArray();
            Assert.Equal("RIGHT_ARROW", Assert.Single(textShapes, shape => shape.Text == "Next step").ShapeType);
            Assert.Equal("TEXT_BOX", Assert.Single(textShapes, shape => shape.Text == "Plain text").ShapeType);
        }

}
