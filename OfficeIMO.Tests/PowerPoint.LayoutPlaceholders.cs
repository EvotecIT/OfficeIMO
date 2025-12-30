using System;
using System.IO;
using DocumentFormat.OpenXml.Presentation;
using OfficeIMO.PowerPoint;
using Xunit;

namespace OfficeIMO.Tests {
    public class PowerPointLayoutPlaceholderTests {
        [Fact]
        public void CanReadLayoutPlaceholders() {
            string filePath = Path.Combine(Path.GetTempPath(), Guid.NewGuid() + ".pptx");
            try {
                using PowerPointPresentation presentation = PowerPointPresentation.Create(filePath);
                PowerPointSlide slide = presentation.AddSlide(SlideLayoutValues.Text);

                var placeholders = slide.GetLayoutPlaceholders();
                Assert.NotEmpty(placeholders);

                PowerPointLayoutPlaceholderInfo? title = slide.GetLayoutPlaceholder(PlaceholderValues.Title);
                Assert.NotNull(title);
                Assert.True(title.Value.Bounds.HasValue);
                Assert.True(title.Value.Bounds.Value.Width > 0);

                PowerPointLayoutBox? bodyBounds = slide.GetLayoutPlaceholderBounds(PlaceholderValues.Body);
                Assert.NotNull(bodyBounds);
                Assert.True(bodyBounds!.Value.Height > 0);
            } finally {
                if (File.Exists(filePath)) {
                    File.Delete(filePath);
                }
            }
        }
    }
}
