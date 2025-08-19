using System;
using System.IO;
using System.Linq;
using OfficeIMO.PowerPoint;
using Xunit;

namespace OfficeIMO.Tests {
    public class PowerPointAdvancedFeatures {
        [Fact]
        public void CanHandleBackgroundFormattingTransitionsAndCharts() {
            string filePath = Path.Combine(Path.GetTempPath(), Guid.NewGuid() + ".pptx");
            string imagePath = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "Images", "BackgroundImage.png");

            using (PowerPointPresentation presentation = PowerPointPresentation.Create(filePath)) {
                PowerPointSlide slide = presentation.AddSlide();
                PowerPointTextBox text = slide.AddTextBox("Test");
                slide.AddPicture(imagePath);
                slide.AddTable(2, 2);
                slide.AddChart();
                slide.Notes.Text = "Notes";

                slide.BackgroundColor = "FF0000";
                text.FillColor = "00FF00";
                slide.Transition = SlideTransition.Fade;

                presentation.Save();
            }

            using (PowerPointPresentation presentation = PowerPointPresentation.Open(filePath)) {
                PowerPointSlide slide = presentation.Slides.Single();
                Assert.Equal("FF0000", slide.BackgroundColor);
                Assert.Equal(SlideTransition.Fade, slide.Transition);
                Assert.Single(slide.TextBoxes);
                Assert.Single(slide.Pictures);
                Assert.Single(slide.Tables);
                Assert.Single(slide.Charts);
                Assert.Equal("00FF00", slide.TextBoxes.First().FillColor);
                Assert.Equal("Notes", slide.Notes.Text);
            }

            File.Delete(filePath);
        }
    }
}
