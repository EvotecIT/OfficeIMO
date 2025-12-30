using System;
using System.IO;
using OfficeIMO.PowerPoint;
using Xunit;

namespace OfficeIMO.Tests {
    public class PowerPointPictureValidationTests {
        [Fact]
        public void AddPictureStream_ThrowsOnNullStream() {
            string filePath = Path.Combine(Path.GetTempPath(), Guid.NewGuid() + ".pptx");
            try {
                using PowerPointPresentation presentation = PowerPointPresentation.Create(filePath);
                PowerPointSlide slide = presentation.AddSlide();

                Assert.Throws<ArgumentNullException>(() => slide.AddPicture(null!, ImagePartType.Png, 0, 0, 1000, 1000));
            } finally {
                if (File.Exists(filePath)) {
                    File.Delete(filePath);
                }
            }
        }

        [Fact]
        public void AddPictureStream_ThrowsOnUnreadableStream() {
            string filePath = Path.Combine(Path.GetTempPath(), Guid.NewGuid() + ".pptx");
            string tempImage = Path.Combine(Path.GetTempPath(), Guid.NewGuid() + ".png");
            try {
                File.WriteAllBytes(tempImage, new byte[] { 0x00, 0x01, 0x02 });
                using PowerPointPresentation presentation = PowerPointPresentation.Create(filePath);
                PowerPointSlide slide = presentation.AddSlide();

                using FileStream stream = new(tempImage, FileMode.Open, FileAccess.Write, FileShare.Read);
                Assert.Throws<ArgumentException>(() => slide.AddPicture(stream, ImagePartType.Png, 0, 0, 1000, 1000));
            } finally {
                if (File.Exists(tempImage)) {
                    File.Delete(tempImage);
                }
                if (File.Exists(filePath)) {
                    File.Delete(filePath);
                }
            }
        }

        [Theory]
        [InlineData(0, 1000)]
        [InlineData(1000, 0)]
        public void AddPictureStream_ThrowsOnInvalidSize(long width, long height) {
            string filePath = Path.Combine(Path.GetTempPath(), Guid.NewGuid() + ".pptx");
            try {
                using PowerPointPresentation presentation = PowerPointPresentation.Create(filePath);
                PowerPointSlide slide = presentation.AddSlide();
                using MemoryStream stream = new(new byte[] { 0x00, 0x01, 0x02 });

                Assert.Throws<ArgumentOutOfRangeException>(() => slide.AddPicture(stream, ImagePartType.Png, 0, 0, width, height));
            } finally {
                if (File.Exists(filePath)) {
                    File.Delete(filePath);
                }
            }
        }

        [Fact]
        public void AddPicturePath_ThrowsOnMissingFile() {
            string filePath = Path.Combine(Path.GetTempPath(), Guid.NewGuid() + ".pptx");
            string missing = Path.Combine(Path.GetTempPath(), Guid.NewGuid() + ".png");
            try {
                using PowerPointPresentation presentation = PowerPointPresentation.Create(filePath);
                PowerPointSlide slide = presentation.AddSlide();

                Assert.Throws<FileNotFoundException>(() => slide.AddPicture(missing, 0, 0, 1000, 1000));
            } finally {
                if (File.Exists(filePath)) {
                    File.Delete(filePath);
                }
            }
        }

        [Theory]
        [InlineData(0, 1000)]
        [InlineData(1000, 0)]
        public void AddPicturePath_ThrowsOnInvalidSize(long width, long height) {
            string filePath = Path.Combine(Path.GetTempPath(), Guid.NewGuid() + ".pptx");
            string imagePath = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "Images", "BackgroundImage.png");
            try {
                using PowerPointPresentation presentation = PowerPointPresentation.Create(filePath);
                PowerPointSlide slide = presentation.AddSlide();

                Assert.Throws<ArgumentOutOfRangeException>(() => slide.AddPicture(imagePath, 0, 0, width, height));
            } finally {
                if (File.Exists(filePath)) {
                    File.Delete(filePath);
                }
            }
        }
    }
}
