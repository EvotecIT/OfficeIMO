using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using OfficeIMO.PowerPoint;
using Xunit;

namespace OfficeIMO.Tests {
    public sealed class PowerPointOpenXmlPackageTypeTests {
        [Theory]
        [InlineData(".potx", PresentationDocumentType.Template, false)]
        [InlineData(".ppsx", PresentationDocumentType.Slideshow, false)]
        [InlineData(".potm", PresentationDocumentType.MacroEnabledTemplate,
            true)]
        [InlineData(".ppsm",
            PresentationDocumentType.MacroEnabledSlideshow, true)]
        public void PathSaveAndCopyPreserveOpenXmlPackageTypeAndVba(
            string extension,
            PresentationDocumentType expectedType,
            bool hasVba) {
            string seedPath = GetTempPath(".pptx");
            string sourcePath = GetTempPath(extension);
            string copyPath = GetTempPath(extension);
            byte[] vbaBytes = { 1, 3, 3, 7 };
            try {
                using (PowerPointPresentation seed =
                       PowerPointPresentation.Create(seedPath)) {
                    seed.AddSlide().AddTitle("Package type");
                    seed.Save();
                }
                File.Copy(seedPath, sourcePath);
                using (PresentationDocument package =
                       PresentationDocument.Open(sourcePath, true)) {
                    package.ChangeDocumentType(expectedType);
                    if (hasVba) {
                        VbaProjectPart part = package.PresentationPart!
                            .AddNewPart<VbaProjectPart>();
                        using var input = new MemoryStream(vbaBytes,
                            writable: false);
                        part.FeedData(input);
                    }
                    package.Save();
                }

                using (PowerPointPresentation presentation =
                       PowerPointPresentation.Load(sourcePath)) {
                    presentation.Slides[0].TextBoxes.Single().Text =
                        "Saved package type";
                    presentation.Save();
                    presentation.SaveCopy(copyPath);
                }

                AssertPackage(sourcePath, expectedType, hasVba, vbaBytes);
                AssertPackage(copyPath, expectedType, hasVba, vbaBytes);
            } finally {
                Delete(seedPath, sourcePath, copyPath);
            }
        }

        private static void AssertPackage(string path,
            PresentationDocumentType expectedType,
            bool hasVba,
            byte[] expectedVbaBytes) {
            using PresentationDocument package =
                PresentationDocument.Open(path, false);
            Assert.Equal(expectedType, package.DocumentType);
            VbaProjectPart? vbaPart = package.PresentationPart!
                .VbaProjectPart;
            if (!hasVba) {
                Assert.Null(vbaPart);
                return;
            }
            Assert.NotNull(vbaPart);
            using Stream input = vbaPart!.GetStream(
                FileMode.Open, FileAccess.Read);
            using var output = new MemoryStream();
            input.CopyTo(output);
            Assert.Equal(expectedVbaBytes, output.ToArray());
        }

        private static string GetTempPath(string extension) =>
            Path.Combine(Path.GetTempPath(),
                Guid.NewGuid().ToString("N") + extension);

        private static void Delete(params string[] paths) {
            foreach (string path in paths) {
                if (File.Exists(path)) File.Delete(path);
            }
        }
    }
}
