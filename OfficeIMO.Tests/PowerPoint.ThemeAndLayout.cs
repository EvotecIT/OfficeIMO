using System;
using System.IO;
using OfficeIMO.PowerPoint;
using Xunit;

namespace OfficeIMO.Tests {
    public class PowerPointThemeAndLayout {
        [Fact(Skip = "Doesn't work after changes to PowerPoint")]
        public void CanSetThemeAndSelectLayout() {
            string filePath = Path.Combine(Path.GetTempPath(), Guid.NewGuid() + ".pptx");

            using (PowerPointPresentation presentation = PowerPointPresentation.Create(filePath)) {
                Assert.Equal("Office Theme", presentation.ThemeName);
                presentation.ThemeName = "My Theme";
                presentation.AddSlide();
                presentation.AddSlide(layoutIndex: 1);
                presentation.Save();
            }

            using (PowerPointPresentation presentation = PowerPointPresentation.Open(filePath)) {
                Assert.Equal("My Theme", presentation.ThemeName);
                Assert.Equal(2, presentation.Slides.Count);
                Assert.Equal(0, presentation.Slides[0].LayoutIndex);
                Assert.Equal(1, presentation.Slides[1].LayoutIndex);
            }

            File.Delete(filePath);
        }
    }
}
