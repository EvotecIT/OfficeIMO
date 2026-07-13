using System;
using System.IO;
using System.Linq;
using DocumentFormat.OpenXml.Packaging;
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

        [Fact]
        public void CanCreateNativeLayoutHeaderFooterPlaceholders() {
            string filePath = Path.Combine(Path.GetTempPath(), Guid.NewGuid() + ".pptx");
            try {
                using (PowerPointPresentation presentation = PowerPointPresentation.Create(filePath)) {
                    int layoutIndex = presentation.GetLayoutIndex(SlideLayoutValues.Text);
                    PowerPointSlide slide = presentation.AddSlide(SlideLayoutValues.Text);

                    var placeholders = presentation.EnsureLayoutHeaderFooterPlaceholders(
                        layoutIndex: layoutIndex,
                        footerText: "Confidential",
                        dateTimeText: "2026-05-16",
                        slideNumberText: "#");

                    Assert.Equal(3, placeholders.Count);
                    Assert.NotNull(slide.GetLayoutPlaceholder(PlaceholderValues.Footer, 10U));
                    Assert.NotNull(slide.GetLayoutPlaceholder(PlaceholderValues.DateAndTime, 11U));
                    Assert.NotNull(slide.GetLayoutPlaceholder(PlaceholderValues.SlideNumber, 12U));
                    Assert.Equal("Confidential", placeholders[0].Text);

                    presentation.Save();
                }

                using (PresentationDocument document = PresentationDocument.Open(filePath, false)) {
                    SlidePart slidePart = document.PresentationPart!.SlideParts.First();
                    SlideLayoutPart layoutPart = slidePart.SlideLayoutPart!;
                    HeaderFooter headerFooter = Assert.Single(layoutPart.SlideLayout!.Elements<HeaderFooter>());

                    Assert.True(headerFooter.Footer!.Value);
                    Assert.True(headerFooter.DateTime!.Value);
                    Assert.True(headerFooter.SlideNumber!.Value);

                    var placeholderTypes = layoutPart.SlideLayout.CommonSlideData!.ShapeTree!
                        .Descendants<PlaceholderShape>()
                        .Select(placeholder => placeholder.Type?.Value)
                        .ToArray();

                    Assert.Contains(PlaceholderValues.Footer, placeholderTypes);
                    Assert.Contains(PlaceholderValues.DateAndTime, placeholderTypes);
                    Assert.Contains(PlaceholderValues.SlideNumber, placeholderTypes);
                }

                using (PowerPointPresentation presentation = PowerPointPresentation.Load(filePath)) {
                    PowerPointSlide slide = presentation.Slides[0];

                    Assert.NotNull(slide.GetLayoutPlaceholder(PlaceholderValues.Footer, 10U));
                    Assert.NotNull(slide.GetLayoutPlaceholder(PlaceholderValues.DateAndTime, 11U));
                    Assert.NotNull(slide.GetLayoutPlaceholder(PlaceholderValues.SlideNumber, 12U));
                }
            } finally {
                if (File.Exists(filePath)) {
                    File.Delete(filePath);
                }
            }
        }
    }
}
