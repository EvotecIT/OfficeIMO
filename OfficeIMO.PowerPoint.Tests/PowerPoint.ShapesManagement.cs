using System;
using System.IO;
using System.Linq;
using DocumentFormat.OpenXml.Presentation;
using OfficeIMO.PowerPoint;
using Xunit;

namespace OfficeIMO.Tests {
    public class PowerPointShapesManagement {
        [Fact]
        public void LoadingSlideRejectsExhaustedDescendantShapeIds() {
            byte[] bytes;
            using (PowerPointPresentation presentation =
                   PowerPointPresentation.Create()) {
                PowerPointSlide slide = presentation.AddSlide();
                PowerPointAutoShape first = slide.AddRectangle(
                    1000, 1000, 2000, 1000);
                PowerPointAutoShape second = slide.AddRectangle(
                    4000, 1000, 2000, 1000);
                PowerPointGroupShape group = slide.GroupShapes(
                    new PowerPointShape[] { first, second });
                NonVisualDrawingProperties descendant = group.GroupShape
                    .Descendants<NonVisualDrawingProperties>().Last();
                descendant.Id = uint.MaxValue;
                bytes = presentation.ToBytes();
            }

            using var stream = new MemoryStream(bytes, writable: false);
            InvalidDataException exception = Assert.Throws<
                InvalidDataException>(() => PowerPointPresentation.Load(
                    stream));
            Assert.Contains("identifier space is exhausted",
                exception.Message, StringComparison.OrdinalIgnoreCase);
        }

        [Fact]
        public void CanGetAndRemoveShapes() {
            string filePath = Path.Combine(Path.GetTempPath(), Guid.NewGuid() + ".pptx");
            string imagePath = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "Images", "BackgroundImage.png");

            using (PowerPointPresentation presentation = PowerPointPresentation.Create(filePath)) {
                PowerPointSlide slide = presentation.AddSlide();
                PowerPointTextBox box1 = slide.AddTextBox("First");
                PowerPointPicture pic1 = slide.AddPicture(imagePath);
                PowerPointTextBox box2 = slide.AddTextBox("Second");
                PowerPointPicture pic2 = slide.AddPicture(imagePath);

                Assert.Equal("TextBox 1", box1.Name);
                Assert.Equal("TextBox 2", box2.Name);
                Assert.Equal("Picture 1", pic1.Name);
                Assert.Equal("Picture 2", pic2.Name);

                Assert.Same(box1, slide.GetTextBox("TextBox 1"));
                Assert.Same(pic2, slide.GetPicture("Picture 2"));
                Assert.Same(box1, slide.GetShape("TextBox 1"));

                slide.RemoveShape(pic1);
                Assert.Null(slide.GetPicture("Picture 1"));
                Assert.Equal(3, slide.Shapes.Count);

                presentation.Save();
            }

            File.Delete(filePath);
        }

        [Fact]
        public void CanLookupAndPersistShapeIdentityMetadata() {
            string filePath = Path.Combine(Path.GetTempPath(), Guid.NewGuid() + ".pptx");

            try {
                using (PowerPointPresentation presentation = PowerPointPresentation.Create(filePath)) {
                    PowerPointSlide slide = presentation.AddSlide();
                    PowerPointTextBox textBox = slide.AddTextBox("Quarterly update");
                    PowerPointAutoShape rectangle = slide.AddRectangle(1000, 1000, 2000, 1000);
                    PowerPointTable table = slide.AddTable(2, 2);

                    textBox.Name = "Hero Title";
                    textBox.AltText = "Main slide title";
                    textBox.Hidden = true;
                    table.Name = "Metrics Table";

                    Assert.NotNull(textBox.Id);
                    Assert.NotNull(rectangle.Id);
                    Assert.NotNull(table.Id);
                    Assert.Equal(PowerPointShapeContentType.TextBox, textBox.ShapeContentType);
                    Assert.Equal(PowerPointShapeContentType.AutoShape, rectangle.ShapeContentType);
                    Assert.Equal(PowerPointShapeContentType.Table, table.ShapeContentType);
                    Assert.Same(textBox, slide.GetShapeByName("Hero Title"));
                    Assert.Same(textBox, slide.GetShapeByName("hero title", ignoreCase: true));
                    Assert.True(slide.TryGetShapeByName("Hero Title", out PowerPointShape? found));
                    Assert.Same(textBox, found);
                    Assert.Same(textBox, slide.GetShapeByName<PowerPointTextBox>("Hero Title"));
                    Assert.Same(table, slide.GetShapeById<PowerPointTable>(table.Id!.Value));
                    Assert.Null(slide.GetShapeByName("Missing"));

                    presentation.Save();
                }

                using (PowerPointPresentation presentation = PowerPointPresentation.Load(filePath)) {
                    PowerPointSlide slide = presentation.Slides[0];
                    PowerPointTextBox? loadedTextBox = slide.GetShapeByName<PowerPointTextBox>("Hero Title");
                    PowerPointTable? loadedTable = slide.GetShapeByName<PowerPointTable>("Metrics Table");

                    Assert.NotNull(loadedTextBox);
                    Assert.NotNull(loadedTable);
                    Assert.Equal("Main slide title", loadedTextBox!.AltText);
                    Assert.True(loadedTextBox.Hidden);
                    Assert.NotNull(slide.GetShapeById(loadedTextBox.Id!.Value));
                    Assert.Equal(PowerPointShapeContentType.TextBox, loadedTextBox.ShapeContentType);
                    Assert.Equal(PowerPointShapeContentType.Table, loadedTable!.ShapeContentType);
                }
            } finally {
                if (File.Exists(filePath)) {
                    File.Delete(filePath);
                }
            }
        }

        [Fact]
        public void ShapeConvenienceDuplicateAndRemoveKeepSlideCollectionInSync() {
            string filePath = Path.Combine(Path.GetTempPath(), Guid.NewGuid() + ".pptx");

            try {
                using (PowerPointPresentation presentation = PowerPointPresentation.Create(filePath)) {
                    PowerPointSlide slide = presentation.AddSlide();
                    PowerPointAutoShape original = slide.AddRectangle(1000, 2000, 3000, 4000);
                    original.Name = "Source Shape";

                    PowerPointShape duplicate = original.Duplicate(offsetX: 500, offsetY: 600);
                    duplicate.Name = "Duplicated Shape";

                    Assert.Equal(2, slide.Shapes.Count);
                    Assert.NotEqual(original.Id, duplicate.Id);
                    Assert.Equal(original.Left + 500, duplicate.Left);
                    Assert.Equal(original.Top + 600, duplicate.Top);
                    Assert.Same(duplicate, slide.GetShapeByName("Duplicated Shape"));

                    original.Remove();

                    Assert.Single(slide.Shapes);
                    Assert.Null(slide.GetShapeByName("Source Shape"));
                    Assert.Same(duplicate, slide.GetShapeByName("Duplicated Shape"));

                    presentation.Save();
                }

                using (PowerPointPresentation presentation = PowerPointPresentation.Load(filePath)) {
                    PowerPointSlide slide = presentation.Slides[0];

                    Assert.Single(slide.Shapes);
                    Assert.Null(slide.GetShapeByName("Source Shape"));
                    Assert.NotNull(slide.GetShapeByName("Duplicated Shape"));
                }
            } finally {
                if (File.Exists(filePath)) {
                    File.Delete(filePath);
                }
            }
        }

        [Fact]
        public void CanAddSvgPictureAndReloadContentType() {
            string filePath = Path.Combine(Path.GetTempPath(), Guid.NewGuid() + ".pptx");
            string svgPath = Path.Combine(Path.GetTempPath(), Guid.NewGuid() + ".svg");

            try {
                File.WriteAllText(svgPath,
                    "<svg xmlns=\"http://www.w3.org/2000/svg\" width=\"120\" height=\"80\" viewBox=\"0 0 120 80\"><rect width=\"120\" height=\"80\" fill=\"#1f4e79\"/><circle cx=\"60\" cy=\"40\" r=\"24\" fill=\"#f4b183\"/></svg>");

                using (PowerPointPresentation presentation = PowerPointPresentation.Create(filePath)) {
                    PowerPointSlide slide = presentation.AddSlide();
                    PowerPointPicture picture = slide.AddPicture(svgPath, left: 1000, top: 1000, width: 2000, height: 1200);
                    picture.Name = "Vector Logo";

                    Assert.Equal("image/svg+xml", picture.ContentType);
                    Assert.Equal(PowerPointShapeContentType.Picture, picture.ShapeContentType);

                    presentation.Save();
                }

                using (PowerPointPresentation presentation = PowerPointPresentation.Load(filePath)) {
                    PowerPointPicture? picture = presentation.Slides[0].GetShapeByName<PowerPointPicture>("Vector Logo");

                    Assert.NotNull(picture);
                    Assert.Equal("image/svg+xml", picture!.ContentType);
                }
            } finally {
                if (File.Exists(filePath)) {
                    File.Delete(filePath);
                }

                if (File.Exists(svgPath)) {
                    File.Delete(svgPath);
                }
            }
        }
    }
}
