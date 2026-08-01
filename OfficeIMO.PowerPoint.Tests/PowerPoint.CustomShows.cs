using System;
using System.IO;
using System.Linq;
using OfficeIMO.PowerPoint;
using Xunit;

namespace OfficeIMO.Tests {
    public sealed class PowerPointCustomShowTests {
        [Fact]
        public void CustomShows_CreateEditAndRoundTripThroughPublicApi() {
            string path = Path.Combine(Path.GetTempPath(),
                "OfficeIMO-CustomShows-" + Guid.NewGuid().ToString("N") + ".pptx");
            try {
                using (PowerPointPresentation presentation =
                       PowerPointPresentation.Create(path)) {
                    PowerPointSlide first = presentation.AddSlide();
                    first.AddTitle("First");
                    PowerPointSlide second = presentation.AddSlide();
                    second.AddTitle("Second");
                    PowerPointSlide third = presentation.AddSlide();
                    third.AddTitle("Third");

                    PowerPointCustomShow show = presentation.AddCustomShow(
                        "Executive path", new[] { third, first });

                    Assert.Equal(1U, show.Id);
                    Assert.Equal(new[] { third, first }, show.Slides);
                    Assert.Contains(presentation.InspectFeatures().EditableFeatures,
                        feature => feature.Name == "Custom shows" && feature.Count == 1);

                    show.InsertSlide(1, second).MoveSlide(2, 0);
                    Assert.True(show.RemoveSlide(second));
                    presentation.RenameCustomShow(show, "Decision path");
                    Assert.Equal(new[] { first, third }, show.Slides);
                    presentation.Save();
                }

                using PowerPointPresentation reopened =
                    PowerPointPresentation.Load(path);
                PowerPointCustomShow saved = Assert.Single(reopened.CustomShows);
                Assert.Equal("Decision path", saved.Name);
                Assert.Equal(new[] { reopened.Slides[0], reopened.Slides[2] },
                    saved.Slides);
                Assert.Equal(saved.Id,
                    reopened.GetCustomShow("decision PATH")!.Id);

                Assert.True(reopened.RemoveCustomShow(saved));
                Assert.Empty(reopened.CustomShows);
                reopened.Save();

                using PowerPointPresentation removed =
                    PowerPointPresentation.Load(path);
                Assert.Empty(removed.CustomShows);
            } finally {
                if (File.Exists(path)) File.Delete(path);
            }
        }

        [Fact]
        public void CustomShows_RejectForeignSlidesShowsDuplicatesAndEmptySequences() {
            using PowerPointPresentation first = PowerPointPresentation.Create();
            PowerPointSlide local = first.AddSlide();
            using PowerPointPresentation second = PowerPointPresentation.Create();
            PowerPointSlide foreign = second.AddSlide();

            PowerPointCustomShow show = first.AddCustomShow("Local",
                new[] { local });
            PowerPointCustomShow foreignShow = second.AddCustomShow("Foreign",
                new[] { foreign });

            Assert.Throws<InvalidOperationException>(() =>
                first.AddCustomShow("Foreign slide", new[] { foreign }));
            Assert.Throws<InvalidOperationException>(() =>
                first.AddCustomShow("LOCAL", new[] { local }));
            Assert.Throws<ArgumentException>(() =>
                first.AddCustomShow("Empty", Array.Empty<PowerPointSlide>()));
            Assert.Throws<InvalidOperationException>(() =>
                show.SetSlides(new[] { foreign }));
            Assert.Throws<InvalidOperationException>(() =>
                first.RenameCustomShow(foreignShow, "Still foreign"));
            Assert.Throws<InvalidOperationException>(() =>
                first.RemoveCustomShow(foreignShow));
        }
    }
}
