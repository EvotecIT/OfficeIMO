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
                using PowerPointDocument presentation = PowerPointDocument.Create(filePath);
                PowerPointSlide slide = presentation.AddSlide();

                Assert.Throws<ArgumentNullException>(() => slide.AddImage(null!, OfficeIMO.Drawing.OfficeImageFormat.Png, global::OfficeIMO.OfficeLength.Emus(0), global::OfficeIMO.OfficeLength.Emus(0), global::OfficeIMO.OfficeLength.Emus(1000), global::OfficeIMO.OfficeLength.Emus(1000)));
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
                using PowerPointDocument presentation = PowerPointDocument.Create(filePath);
                PowerPointSlide slide = presentation.AddSlide();

                using FileStream stream = new(tempImage, FileMode.Open, FileAccess.Write, FileShare.Read);
                Assert.Throws<ArgumentException>(() => slide.AddImage(stream, OfficeIMO.Drawing.OfficeImageFormat.Png, global::OfficeIMO.OfficeLength.Emus(0), global::OfficeIMO.OfficeLength.Emus(0), global::OfficeIMO.OfficeLength.Emus(1000), global::OfficeIMO.OfficeLength.Emus(1000)));
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
                using PowerPointDocument presentation = PowerPointDocument.Create(filePath);
                PowerPointSlide slide = presentation.AddSlide();
                using MemoryStream stream = new(new byte[] { 0x00, 0x01, 0x02 });

                Assert.Throws<ArgumentOutOfRangeException>(() => slide.AddImage(stream, OfficeIMO.Drawing.OfficeImageFormat.Png, global::OfficeIMO.OfficeLength.Emus(0), global::OfficeIMO.OfficeLength.Emus(0), global::OfficeIMO.OfficeLength.Emus(width), global::OfficeIMO.OfficeLength.Emus(height)));
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
                using PowerPointDocument presentation = PowerPointDocument.Create(filePath);
                PowerPointSlide slide = presentation.AddSlide();

                Assert.Throws<FileNotFoundException>(() => slide.AddImage(missing,global::OfficeIMO.OfficeLength.Emus(0),global::OfficeIMO.OfficeLength.Emus(0),global::OfficeIMO.OfficeLength.Emus(1000),global::OfficeIMO.OfficeLength.Emus(1000)));
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
                using PowerPointDocument presentation = PowerPointDocument.Create(filePath);
                PowerPointSlide slide = presentation.AddSlide();

                Assert.Throws<ArgumentOutOfRangeException>(() => slide.AddImage(imagePath,global::OfficeIMO.OfficeLength.Emus(0),global::OfficeIMO.OfficeLength.Emus(0),global::OfficeIMO.OfficeLength.Emus(width),global::OfficeIMO.OfficeLength.Emus(height)));
            } finally {
                if (File.Exists(filePath)) {
                    File.Delete(filePath);
                }
            }
        }
    }
}
