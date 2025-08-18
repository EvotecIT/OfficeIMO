using OfficeIMO.PowerPoint;
using Xunit;

namespace OfficeIMO.Tests {
    public class PowerPointBasicDocument {
        [Fact]
        public void CanCreateBasicPowerPointDocument() {
            PowerPointDocument document = new();
            PowerPointSlide slide = document.AddSlide("Slide1");
            slide.Shapes.Add(new PowerPointShape("Shape1"));

            Assert.Single(document.Slides);
            Assert.Single(slide.Shapes);
        }
    }
}
