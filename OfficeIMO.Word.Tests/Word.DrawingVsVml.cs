using System;
using System.Globalization;
using System.IO;
using System.Linq;
using DocumentFormat.OpenXml.Packaging;
using OfficeIMO.Word;
using Color = OfficeIMO.OfficeColor;
using V = DocumentFormat.OpenXml.Vml;
using Xunit;
using Path = System.IO.Path;

namespace OfficeIMO.Tests {
    public partial class Word {
        [Fact]
        public void VmlImageExplicitPointDimensionsAndHeaderStreamRoundTripUnderCommaDecimalCulture() {
            CultureInfo originalCulture = CultureInfo.CurrentCulture;
            try {
                CultureInfo.CurrentCulture = CultureInfo.GetCultureInfo("pl-PL");
                string source = Path.Combine(AppContext.BaseDirectory, "Images", "EvotecLogo.png");
                byte[] expected = File.ReadAllBytes(source);
                using var document = WordDocument.Create();
                document.AddParagraph().AddImageVml(source, 24.5D, 12.25D);
                document.AddParagraph().AddImageVml(source);
                var header = document.Sections[0].GetOrCreateHeader(WordHeaderFooterType.Default);
                WordImage image;
                using (var input = new MemoryStream(expected, writable: false)) {
                    input.Position = 3;
                    image = header.AddParagraph().InsertImageVml(input, "stream.png", 24.5D, 12.25D, "Stream image");
                    Assert.True(input.CanRead);
                    Assert.Equal(3L, input.Position);
                }
                Assert.Equal(expected, image.ToBytes());

                using var serialized = new MemoryStream(document.ToBytes());
                using var reopened = WordDocument.Load(serialized);
                var main = reopened._wordprocessingDocument.MainDocumentPart!;
                V.Shape[] bodyImages = main.Document.Body!.Descendants<V.Shape>().ToArray();
                Assert.Equal(2, bodyImages.Length);
                Assert.Contains("width:24.5pt", bodyImages[0].Style!.Value);
                Assert.Contains("height:12.25pt", bodyImages[0].Style!.Value);
                Assert.DoesNotContain("width:", bodyImages[1].Style!.Value);
                Assert.DoesNotContain("height:", bodyImages[1].Style!.Value);
                Assert.Equal(2, main.ImageParts.Count());
                HeaderPart headerPart = Assert.Single(main.HeaderParts);
                V.ImageData headerImage = Assert.Single(headerPart.Header.Descendants<V.ImageData>());
                var ownedPart = Assert.IsType<ImagePart>(headerPart.GetPartById(headerImage.RelationshipId!.Value!));
                Assert.Same(ownedPart, Assert.Single(headerPart.ImageParts));
                using var payload = new MemoryStream();
                using (Stream input = ownedPart.GetStream()) input.CopyTo(payload);
                Assert.Equal(expected, payload.ToArray());
                WordImage persisted = reopened.Sections[0].Header.Default!.Paragraphs.Single(paragraph => paragraph.IsImage).Image!;
                Assert.Equal(expected, persisted.ToBytes());
                Assert.Equal("Stream image", persisted.Description);
                Assert.Contains("width:24.5pt", persisted._vmlShape!.Style!.Value);
                Assert.Contains("height:12.25pt", persisted._vmlShape.Style.Value);
            } finally {
                CultureInfo.CurrentCulture = originalCulture;
            }
        }

        [Fact]
        public void VmlImageFileNameFallsBackToEmbeddedPartWhenTitleIsAbsent() {
            string source = Path.Combine(AppContext.BaseDirectory, "Images", "EvotecLogo.png");
            string output = Path.Combine(Path.GetTempPath(), Guid.NewGuid().ToString("N") + ".docx");
            try {
                using (WordDocument document = WordDocument.Create(output)) {
                    document.AddImageVml(source);
                    document.Save();
                }
                string expected;
                using (WordprocessingDocument native = WordprocessingDocument.Open(output, true)) {
                    V.ImageData image = native.MainDocumentPart!.Document!.Descendants<V.ImageData>().Single();
                    expected = Path.GetFileName(native.MainDocumentPart.GetPartById(image.RelationshipId!.Value!).Uri.ToString());
                    image.Title = null;
                    native.MainDocumentPart.Document.Save();
                }
                using WordDocument reopened = WordDocument.Load(output);
                Assert.Equal(expected, Assert.Single(reopened.Images).FileName);
            } finally { if (File.Exists(output)) File.Delete(output); }
        }

        [Fact]
        public void VmlImageFileNameRoundTripsThroughItsPublicProperty() {
            string source = Path.Combine(AppContext.BaseDirectory, "Images", "EvotecLogo.png");
            string output = Path.Combine(Path.GetTempPath(), Guid.NewGuid().ToString("N") + ".docx");
            try {
                using (WordDocument document = WordDocument.Create(output)) {
                    document.AddImageVml(source);
                    WordImage image = Assert.Single(document.Images);
                    Assert.Equal(Path.GetFileName(source), image.FileName);
                    image.FileName = "renamed.png";
                    Assert.Equal("renamed.png", image.FileName);
                    document.Save();
                }
                using WordDocument reopened = WordDocument.Load(output);
                Assert.Equal("renamed.png", Assert.Single(reopened.Images).FileName);
            } finally { if (File.Exists(output)) File.Delete(output); }
        }

        [Fact]
        public void Test_DrawingVsVmlCounts() {
            string img = Path.Combine(AppContext.BaseDirectory, "Images", "EvotecLogo.png");

            string drawingFile = Path.Combine(_directoryWithFiles, "DrawingCounts.docx");
            using (WordDocument doc = WordDocument.Create(drawingFile)) {
                doc.AddParagraph().AddImage(img);
                doc.AddShapeDrawing(WordShapeType.Ellipse, 40, 40);
                doc.AddShapeDrawing(WordShapeType.Rectangle, 60, 30);
                doc.AddShapeDrawing(WordShapeType.RoundedRectangle, 50, 30);
                doc.AddTextBox("Text");
                doc.Save();
            }
            using (WordDocument doc = WordDocument.Load(drawingFile)) {
                Assert.Single(doc.Images);
                Assert.Equal(3, doc.Shapes.Count);
                Assert.Single(doc.TextBoxes);
            }

            string vmlFile = Path.Combine(_directoryWithFiles, "VmlCounts.docx");
            using (WordDocument doc = WordDocument.Create(vmlFile)) {
                doc.AddImageVml(img);
                doc.AddShape(WordShapeType.Ellipse, 40, 40, Color.Red, Color.Blue);
                doc.AddShape(WordShapeType.Rectangle, 60, 30, Color.Green, Color.Black);
                doc.AddShape(WordShapeType.RoundedRectangle, 50, 30, Color.Yellow, Color.Black, 1, arcSize: 0.3);
                doc.AddTextBoxVml("Text");
                doc.Save();
            }
            using (WordDocument doc = WordDocument.Load(vmlFile)) {
                Assert.Single(doc.Images);
                Assert.Equal(3, doc.Shapes.Count);
                Assert.Single(doc.TextBoxes);
            }
        }
    }
}

