using System;
using System.IO;
using OfficeIMO.PowerPoint;
using Xunit;

namespace OfficeIMO.Tests {
    public class PowerPointValidation {
        [Fact]
        public void TemplatePresentationIsValid() {
            string templatePath = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "..", "..", "..", "..", "Assets", "PowerPointTemplates", "PowerPointBlank.pptx");

            using PowerPointPresentation presentation = PowerPointPresentation.Open(templatePath);
            Assert.Empty(presentation.ValidatePresentation());
        }

        [Fact]
        public void NewPresentationIsValid() {
            string filePath = Path.Combine(Path.GetTempPath(), Guid.NewGuid() + ".pptx");

            using (PowerPointPresentation presentation = PowerPointPresentation.Create(filePath)) {
                presentation.AddSlide();
                presentation.Save();
                Assert.Empty(presentation.ValidatePresentation());
            }

            File.Delete(filePath);
        }
    }
}
