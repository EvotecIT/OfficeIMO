using System;
using System.IO;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Presentation;
using OfficeIMO.Core;
using OfficeIMO.PowerPoint;
using Xunit;

namespace OfficeIMO.Tests {
    public class PowerPointLifecycleTests {
        [Fact]
        public void Create_Path_IsDetachedUntilExplicitSave() {
            string path = Path.Combine(Path.GetTempPath(), Guid.NewGuid() + ".pptx");
            try {
                using (PowerPointPresentation presentation = PowerPointPresentation.Create(path)) {
                    presentation.AddSlide().AddTitle("Detached");
                    Assert.False(File.Exists(path));
                }
                Assert.False(File.Exists(path));

                using (PowerPointPresentation presentation = PowerPointPresentation.Create(path)) {
                    presentation.AddSlide().AddTitle("Saved");
                    presentation.Save();
                }

                Assert.True(File.Exists(path));
                using PresentationDocument package = PresentationDocument.Open(path, false);
                Assert.Single(package.PresentationPart!.Presentation.SlideIdList!.Elements<SlideId>());
            } finally {
                if (File.Exists(path)) File.Delete(path);
            }
        }

        [Fact]
        public void Create_Detached_SaveWithoutDestinationFailsClearly() {
            using PowerPointPresentation presentation = PowerPointPresentation.Create();
            presentation.AddSlide();
            InvalidOperationException exception = Assert.Throws<InvalidOperationException>(() => presentation.Save());
            Assert.Contains("no associated destination", exception.Message, StringComparison.OrdinalIgnoreCase);
        }

        [Fact]
        public void Load_Stream_PreservesCallerPositionAndDoesNotWriteByDefault() {
            using var stream = new MemoryStream();
            using (PowerPointPresentation presentation = PowerPointPresentation.Create(stream)) {
                presentation.AddSlide().AddTitle("Original");
                presentation.Save();
            }
            byte[] original = stream.ToArray();
            stream.Position = Math.Min(5, stream.Length);
            long originalPosition = stream.Position;

            using (PowerPointPresentation presentation = PowerPointPresentation.Load(stream,
                       new PowerPointLoadOptions {
                           OpenSettings = new OpenSettings { AutoSave = true }
                       })) {
                presentation.Slides[0].AddTextBox("Unsaved");
                Assert.Equal(originalPosition, stream.Position);
            }

            Assert.Equal(originalPosition, stream.Position);
            Assert.Equal(original, stream.ToArray());
        }

        [Fact]
        public void Load_ReadOnlyRejectsSaveOnDispose() {
            using var stream = new MemoryStream();
            using (PowerPointPresentation presentation = PowerPointPresentation.Create(stream)) {
                presentation.AddSlide();
                presentation.Save();
            }

            Assert.Throws<ArgumentException>(() => PowerPointPresentation.Load(stream,
                new PowerPointLoadOptions {
                    AccessMode = DocumentAccessMode.ReadOnly,
                    PersistenceMode = DocumentPersistenceMode.SaveOnDispose
                }));
        }
    }
}
