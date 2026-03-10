using System;
using System.IO;
using System.Linq;
using System.Collections.Generic;
using DocumentFormat.OpenXml.Office2010.PowerPoint;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Presentation;
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

        [Fact]
        public void SectionsStayConsistentAfterDuplicateAndRemove() {
            string filePath = Path.Combine(Path.GetTempPath(), Guid.NewGuid() + ".pptx");
            try {
                using (PowerPointPresentation presentation = PowerPointPresentation.Create(filePath)) {
                    presentation.AddSlide().AddTextBox("Intro 1");
                    presentation.AddSlide().AddTextBox("Intro 2");
                    presentation.AddSlide().AddTextBox("Results 1");
                    presentation.AddSlide().AddTextBox("Results 2");

                    presentation.AddSection("Intro", startSlideIndex: 0);
                    presentation.AddSection("Results", startSlideIndex: 2);

                    presentation.DuplicateSlide(2, insertAt: 3);
                    PowerPointSectionInfo results = presentation.GetSections().First(section => section.Name == "Results");
                    Assert.Contains(2, results.SlideIndices);
                    Assert.Contains(3, results.SlideIndices);

                    presentation.RemoveSlide(4);
                    presentation.Save();
                }

                using PresentationDocument document = PresentationDocument.Open(filePath, false);
                AssertSectionIdsMatchSlides(document);

                using PowerPointPresentation reopened = PowerPointPresentation.Open(filePath);
                PowerPointSectionInfo[] sections = reopened.GetSections().ToArray();
                Assert.Equal(new[] { "Intro", "Results" }, sections.Select(section => section.Name).ToArray());
                Assert.Equal(new[] { 2, 3 }, sections.Single(section => section.Name == "Results").SlideIndices.ToArray());
            } finally {
                if (File.Exists(filePath)) {
                    File.Delete(filePath);
                }
            }
        }

        [Fact]
        public void MovingSlideReordersSectionsByCurrentSlideOrder() {
            string filePath = Path.Combine(Path.GetTempPath(), Guid.NewGuid() + ".pptx");
            try {
                using (PowerPointPresentation presentation = PowerPointPresentation.Create(filePath)) {
                    presentation.AddSlide().AddTextBox("Intro");
                    presentation.AddSlide().AddTextBox("Middle");
                    presentation.AddSlide().AddTextBox("Results");

                    presentation.AddSection("Intro", startSlideIndex: 0);
                    presentation.AddSection("Results", startSlideIndex: 2);

                    presentation.MoveSlide(2, 0);
                    presentation.Save();
                }

                using PresentationDocument document = PresentationDocument.Open(filePath, false);
                AssertSectionIdsMatchSlides(document);

                using PowerPointPresentation reopened = PowerPointPresentation.Open(filePath);
                PowerPointSectionInfo[] sections = reopened.GetSections().ToArray();
                Assert.Equal(new[] { "Results", "Intro" }, sections.Select(section => section.Name).ToArray());
                Assert.Equal(new[] { 0 }, sections[0].SlideIndices.ToArray());
            } finally {
                if (File.Exists(filePath)) {
                    File.Delete(filePath);
                }
            }
        }

        private static void AssertSectionIdsMatchSlides(PresentationDocument document) {
            SlideIdList slideIdList = document.PresentationPart!.Presentation.SlideIdList!;
            HashSet<uint> slideIds = slideIdList.Elements<SlideId>()
                .Select(slideId => slideId.Id?.Value ?? 0U)
                .ToHashSet();

            SectionList sectionList = document.PresentationPart.Presentation
                .GetFirstChild<PresentationExtensionList>()!
                .Elements<PresentationExtension>()
                .First(extension => extension.Uri?.Value == "{521415D9-36F7-43E2-AB2F-B90AF26B5E84}")
                .GetFirstChild<SectionList>()!;

            uint[] sectionSlideIds = sectionList.Elements<Section>()
                .SelectMany(section => section.SectionSlideIdList?.Elements<SectionSlideIdListEntry>()
                    ?? Enumerable.Empty<SectionSlideIdListEntry>())
                .Select(entry => entry.Id?.Value ?? 0U)
                .ToArray();

            Assert.Equal(slideIds.Count, sectionSlideIds.Length);
            Assert.All(sectionSlideIds, id => Assert.Contains(id, slideIds));
            Assert.Equal(sectionSlideIds.Length, sectionSlideIds.Distinct().Count());
        }
    }
}
