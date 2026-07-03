using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Presentation;
using OfficeIMO.PowerPoint;
using Xunit;

namespace OfficeIMO.Tests {
    public class PowerPointThemeAndLayout {
        [Fact]
        public void CanSetThemeAndSelectLayout() {
            string filePath = Path.Combine(Path.GetTempPath(), Guid.NewGuid() + ".pptx");
            List<SlideLayoutValues> expectedLayouts = new() {
                SlideLayoutValues.Title,
                SlideLayoutValues.Text,
                SlideLayoutValues.SectionHeader,
                SlideLayoutValues.TwoColumnText,
                SlideLayoutValues.TwoObjects,
                SlideLayoutValues.TitleOnly,
                SlideLayoutValues.Blank,
                SlideLayoutValues.PictureText,
                SlideLayoutValues.VerticalTitleAndText,
                SlideLayoutValues.VerticalText,
                SlideLayoutValues.TwoObjects
            };

            using (PowerPointPresentation presentation = PowerPointPresentation.Create(filePath)) {
                Assert.Equal("Office Theme", presentation.ThemeName);
                presentation.ThemeName = "My Theme";
                Assert.Single(presentation.Slides);
                Assert.Equal(0, presentation.Slides[0].LayoutIndex);

                // Consume the initial slide so subsequent calls create new slides.
                presentation.AddSlide();

                for (int i = 1; i < expectedLayouts.Count; i++) {
                    PowerPointSlide slide = presentation.AddSlide(layoutIndex: i);
                    Assert.Equal(i, slide.LayoutIndex);
                }

                Assert.Equal(expectedLayouts.Count, presentation.Slides.Count);
                presentation.Save();
            }

            using (PowerPointPresentation presentation = PowerPointPresentation.Open(filePath)) {
                Assert.Equal("My Theme", presentation.ThemeName);
                Assert.Equal(expectedLayouts.Count, presentation.Slides.Count);

                for (int i = 0; i < presentation.Slides.Count; i++) {
                    Assert.Equal(i, presentation.Slides[i].LayoutIndex);
                }
            }

            using (PresentationDocument document = PresentationDocument.Open(filePath, false)) {
                SlideMasterPart masterPart = document.PresentationPart!.SlideMasterParts.First();
                SlideLayoutPart[] layoutParts = masterPart.SlideLayoutParts.ToArray();

                Assert.Equal(expectedLayouts.Count, layoutParts.Length);

                SlideLayoutIdList layoutIdList = masterPart.SlideMaster.SlideLayoutIdList!;
                for (int i = 0; i < layoutParts.Length; i++) {
                    Assert.Equal(expectedLayouts[i], layoutParts[i].SlideLayout.Type?.Value);
                    string relId = masterPart.GetIdOfPart(layoutParts[i]);
                    Assert.Contains(layoutIdList.Elements<SlideLayoutId>(), id => id.RelationshipId == relId);
                }
            }

            File.Delete(filePath);
        }
    }
}
