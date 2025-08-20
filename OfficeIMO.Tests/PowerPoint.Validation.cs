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
    }
}
