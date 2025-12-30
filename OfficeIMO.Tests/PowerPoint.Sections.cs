using System;
using System.IO;
using System.Linq;
using OfficeIMO.PowerPoint;
using Xunit;

namespace OfficeIMO.Tests {
    public class PowerPointSectionsTests {
        [Fact]
        public void CanAddAndRenameSections() {
            string filePath = Path.Combine(Path.GetTempPath(), Guid.NewGuid() + ".pptx");
            try {
                using PowerPointPresentation presentation = PowerPointPresentation.Create(filePath);
                presentation.AddSlide();
                presentation.AddSlide();
                presentation.AddSlide();

                presentation.AddSection("Intro", startSlideIndex: 0);
                presentation.AddSection("Results", startSlideIndex: 2);
                Assert.True(presentation.RenameSection("Results", "Deep Dive"));

                var sections = presentation.GetSections();
                Assert.Contains(sections, s => s.Name == "Intro");
                Assert.Contains(sections, s => s.Name == "Deep Dive");

                PowerPointSectionInfo deepDive = sections.First(s => s.Name == "Deep Dive");
                Assert.Contains(2, deepDive.SlideIndices);
            } finally {
                if (File.Exists(filePath)) {
                    File.Delete(filePath);
                }
            }
        }
    }
}
