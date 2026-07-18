using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using OfficeIMO.PowerPoint;
using System.Threading.Tasks;
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
        [InlineData(".ppam", PresentationDocumentType.AddIn, true)]
        public void PathSaveCopyAndEncryptionPreserveOpenXmlPackageTypeAndVba(
            string extension,
            PresentationDocumentType expectedType,
            bool hasVba) {
            string seedPath = GetTempPath(".pptx");
            string sourcePath = GetTempPath(extension);
            string copyPath = GetTempPath(extension);
            string encryptedPath = GetTempPath(extension);
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
                    presentation.SaveEncrypted(encryptedPath,
                        "OfficeIMO-path-pass");
                }

                AssertPackage(sourcePath, expectedType, hasVba, vbaBytes);
                AssertPackage(copyPath, expectedType, hasVba, vbaBytes);
                using PowerPointPresentation decrypted =
                    PowerPointPresentation.LoadEncrypted(encryptedPath,
                        "OfficeIMO-path-pass");
                Assert.Equal(expectedType,
                    decrypted.OpenXmlDocument.DocumentType);
                VbaProjectPart? encryptedVbaPart =
                    decrypted.OpenXmlDocument.PresentationPart!
                        .VbaProjectPart;
                Assert.Equal(hasVba, encryptedVbaPart != null);
                if (hasVba) {
                    Assert.Equal(vbaBytes,
                        ReadVbaBytes(encryptedVbaPart));
                }
            } finally {
                Delete(seedPath, sourcePath, copyPath, encryptedPath);
            }
        }

        [Theory]
        [InlineData(PresentationDocumentType.MacroEnabledPresentation)]
        [InlineData(PresentationDocumentType.MacroEnabledTemplate)]
        [InlineData(PresentationDocumentType.MacroEnabledSlideshow)]
        public async Task StreamSavesPreserveLoadedPackageTypeAndVba(
            PresentationDocumentType expectedType) {
            byte[] vbaBytes = { 2, 4, 6, 8 };
            using MemoryStream source = CreatePackageStream(expectedType,
                vbaBytes);
            using PowerPointPresentation presentation =
                PowerPointPresentation.Load(source);
            presentation.Slides[0].TextBoxes.Single().Text =
                "Saved to associated stream";

            presentation.Save();
            AssertPackage(source, expectedType, hasVba: true, vbaBytes);

            presentation.Slides[0].AddTextBox("Saved asynchronously");
            await presentation.SaveAsync();
            AssertPackage(source, expectedType, hasVba: true, vbaBytes);

            using var encrypted = new MemoryStream();
            presentation.SaveEncrypted(encrypted, "OfficeIMO-stream-pass");
            using PowerPointPresentation decrypted =
                PowerPointPresentation.LoadEncrypted(encrypted,
                    "OfficeIMO-stream-pass");
            Assert.Equal(expectedType,
                decrypted.OpenXmlDocument.DocumentType);
            Assert.Equal(vbaBytes, ReadVbaBytes(
                decrypted.OpenXmlDocument.PresentationPart!
                    .VbaProjectPart));
        }

        private static MemoryStream CreatePackageStream(
            PresentationDocumentType documentType,
            byte[] vbaBytes) {
            string path = GetTempPath(".pptx");
            try {
                using (PowerPointPresentation seed =
                       PowerPointPresentation.Create(path)) {
                    seed.AddSlide().AddTitle("Package type");
                    seed.Save();
                }
                using (PresentationDocument package =
                       PresentationDocument.Open(path, true)) {
                    package.ChangeDocumentType(documentType);
                    VbaProjectPart part = package.PresentationPart!
                        .AddNewPart<VbaProjectPart>();
                    using var input = new MemoryStream(vbaBytes,
                        writable: false);
                    part.FeedData(input);
                    package.Save();
                }
                byte[] packageBytes = File.ReadAllBytes(path);
                var stream = new MemoryStream(packageBytes.Length + 4096);
                stream.Write(packageBytes, 0, packageBytes.Length);
                stream.Position = 0;
                return stream;
            } finally {
                Delete(path);
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
            Assert.Equal(expectedVbaBytes, ReadVbaBytes(vbaPart));
        }

        private static void AssertPackage(Stream stream,
            PresentationDocumentType expectedType,
            bool hasVba,
            byte[] expectedVbaBytes) {
            stream.Position = 0;
            using PresentationDocument package =
                PresentationDocument.Open(stream, false);
            Assert.Equal(expectedType, package.DocumentType);
            VbaProjectPart? vbaPart = package.PresentationPart!
                .VbaProjectPart;
            Assert.Equal(hasVba, vbaPart != null);
            if (hasVba) {
                Assert.Equal(expectedVbaBytes, ReadVbaBytes(vbaPart));
            }
        }

        private static byte[] ReadVbaBytes(VbaProjectPart? vbaPart) {
            Assert.NotNull(vbaPart);
            using Stream input = vbaPart!.GetStream(
                FileMode.Open, FileAccess.Read);
            using var output = new MemoryStream();
            input.CopyTo(output);
            return output.ToArray();
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
