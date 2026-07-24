using System;
using System.IO;
using System.Linq;
using System.Text;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Presentation;
using A = DocumentFormat.OpenXml.Drawing;
using C = DocumentFormat.OpenXml.Drawing.Charts;
using S = DocumentFormat.OpenXml.Spreadsheet;
using OfficeIMO.PowerPoint;
using Xunit;

namespace OfficeIMO.Tests {
    public class PowerPointFeatureReportTests {
        [Fact]
        public void PowerPointFeatureReport_DetectsEditableAndPartiallyEditableFeatures() {
            string filePath = Path.Combine(Path.GetTempPath(), Guid.NewGuid() + ".pptx");
            string imagePath = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "Images", "BackgroundImage.png");

            try {
                using (PowerPointPresentation presentation = PowerPointPresentation.Create(filePath)) {
                    PowerPointSlide slide = presentation.AddSlide();
                    PowerPointTextBox textBox = slide.AddTextBox("Quarterly review");
                    textBox.Paragraphs.First().Runs.First().SetHyperlink("https://example.com/review");
                    slide.AddPicture(imagePath);
                    PowerPointTable table = slide.AddTable(2, 2);
                    table.GetCell(0, 0).Text = "Metric";
                    slide.AddChart();
                    slide.Notes.Text = "Talk track";
                    slide.Transition = SlideTransition.Fade;
                    presentation.AddSection("Results", 0);
                    presentation.Save();
                }

                using (PowerPointPresentation presentation = PowerPointPresentation.Load(filePath, new PowerPointLoadOptions { AccessMode = OfficeIMO.Drawing.DocumentAccessMode.ReadOnly })) {
                    PowerPointFeatureReport report = presentation.InspectFeatures();

                    Assert.Contains(report.EditableFeatures, feature => feature.Name == "Slides" && feature.Count == 1);
                    Assert.Contains(report.EditableFeatures, feature => feature.Name == "Text boxes" && feature.Count == 1);
                    Assert.Contains(report.EditableFeatures, feature => feature.Name == "Tables" && feature.Count == 1);
                    Assert.Contains(report.EditableFeatures, feature => feature.Name == "Table style metadata"
                        && feature.Details.Any(detail => detail.Contains("colIds=2", StringComparison.OrdinalIgnoreCase)
                            && detail.Contains("rowIds=2", StringComparison.OrdinalIgnoreCase)));
                    Assert.Contains(report.EditableFeatures, feature => feature.Name == "Speaker notes" && feature.Count == 1);
                    Assert.Contains(report.EditableFeatures, feature => feature.Name == "Slide transitions" && feature.Count == 1);
                    Assert.Contains(report.PartiallyEditableFeatures, feature => feature.Name == "Charts" && feature.Count == 1);
                    Assert.Contains(report.PartiallyEditableFeatures, feature => feature.Name == "Images" && feature.Count == 1);
                    Assert.Contains(report.PartiallyEditableFeatures, feature => feature.Name == "External relationships" && feature.Count == 1);
                    Assert.DoesNotContain(report.PreservedFeatures, feature => feature.Name == "Embedded packages");
                    Assert.Empty(report.UnsupportedFeatures);
                    Assert.Same(report, report.EnsureNoUnsupportedFeatures());
                    Assert.Contains("| Content | Tables |", report.ToMarkdown());
                }
            } finally {
                if (File.Exists(filePath)) {
                    File.Delete(filePath);
                }
            }
        }

        [Fact]
        public void PowerPointFeatureReport_DetectsGroupedPictures() {
            string filePath = Path.Combine(Path.GetTempPath(), Guid.NewGuid() + ".pptx");
            byte[] pixel = Convert.FromBase64String("iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAQAAAC1HAwCAAAAC0lEQVR4nGMAAQAABQABDQottAAAAABJRU5ErkJggg==");

            try {
                using (PowerPointPresentation presentation = PowerPointPresentation.Create(filePath)) {
                    PowerPointSlide slide = presentation.AddSlide();
                    using var image = new MemoryStream(pixel);
                    PowerPointPicture picture = slide.AddPicture(image, OfficeIMO.PowerPoint.ImagePartType.Png);
                    PowerPointAutoShape shape = slide.AddRectangle(914400, 0, 914400, 914400);
                    slide.GroupShapes(new PowerPointShape[] { picture, shape });
                    presentation.Save();
                }

                using (PowerPointPresentation presentation = PowerPointPresentation.Load(filePath, new PowerPointLoadOptions { AccessMode = OfficeIMO.Drawing.DocumentAccessMode.ReadOnly })) {
                    PowerPointFeatureReport report = presentation.InspectFeatures();
                    PowerPointFeatureFinding images = Assert.Single(report.FindFeatures("Images"));

                    Assert.Equal(1, images.Count);
                    Assert.Throws<InvalidOperationException>(() => report.EnsureNoFeatures("Images"));
                }
            } finally {
                if (File.Exists(filePath)) {
                    File.Delete(filePath);
                }
            }
        }

        [Fact]
        public void PowerPointFeatureReport_DetectsBackgroundImages() {
            string filePath = Path.Combine(Path.GetTempPath(), Guid.NewGuid() + ".pptx");
            string imagePath = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "Images", "BackgroundImage.png");

            try {
                using (PowerPointPresentation presentation = PowerPointPresentation.Create(filePath)) {
                    presentation.AddSlide().SetBackgroundImage(imagePath);
                    presentation.Save();
                }

                using (PowerPointPresentation presentation = PowerPointPresentation.Load(filePath, new PowerPointLoadOptions { AccessMode = OfficeIMO.Drawing.DocumentAccessMode.ReadOnly })) {
                    PowerPointFeatureReport report = presentation.InspectFeatures();
                    PowerPointFeatureFinding images = Assert.Single(report.FindFeatures("Images"));

                    Assert.Equal(1, images.Count);
                    Assert.Throws<InvalidOperationException>(() => report.EnsureNoFeatures("Images"));
                }
            } finally {
                if (File.Exists(filePath)) {
                    File.Delete(filePath);
                }
            }
        }

        [Fact]
        public void PowerPointFeatureReport_DetectsUnsupportedBackgroundBlipImages() {
            string filePath = Path.Combine(Path.GetTempPath(), Guid.NewGuid() + ".pptx");
            string imagePath = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "Images", "BackgroundImage.png");

            try {
                using (PowerPointPresentation presentation = PowerPointPresentation.Create(filePath)) {
                    presentation.AddSlide().SetBackgroundImage(imagePath);
                    presentation.Save();
                }

                using (PresentationDocument document = PresentationDocument.Open(filePath, true)) {
                    SlidePart slidePart = document.PresentationPart!.SlideParts.Single();
                    A.BlipFill blipFill = slidePart.Slide.CommonSlideData!.Background!.BackgroundProperties!
                        .GetFirstChild<A.BlipFill>()!;
                    blipFill.RemoveAllChildren<A.Stretch>();
                    blipFill.Append(new A.Tile());
                    slidePart.Slide.Save();
                }

                using (PowerPointPresentation presentation = PowerPointPresentation.Load(filePath, new PowerPointLoadOptions { AccessMode = OfficeIMO.Drawing.DocumentAccessMode.ReadOnly })) {
                    PowerPointFeatureReport report = presentation.InspectFeatures();
                    PowerPointFeatureFinding images = Assert.Single(report.FindFeatures("Images"));

                    Assert.Equal(1, images.Count);
                    Assert.Throws<InvalidOperationException>(() => report.EnsureNoFeatures("Images"));
                }
            } finally {
                if (File.Exists(filePath)) {
                    File.Delete(filePath);
                }
            }
        }

        [Fact]
        public void PowerPointFeatureReport_DoesNotInheritBackgroundImageWhenSlideHasBackgroundReference() {
            string filePath = Path.Combine(Path.GetTempPath(), Guid.NewGuid() + ".pptx");
            string imagePath = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "Images", "BackgroundImage.png");

            try {
                using (PowerPointPresentation presentation = PowerPointPresentation.Create(filePath)) {
                    presentation.AddSlide().AddTextBox("Theme background override");
                    presentation.Save();
                }

                using (PresentationDocument document = PresentationDocument.Open(filePath, true)) {
                    SlidePart slidePart = document.PresentationPart!.SlideParts.Single();
                    SlideMasterPart masterPart = slidePart.SlideLayoutPart!.SlideMasterPart!;
                    ImagePart imagePart = masterPart.AddImagePart(DocumentFormat.OpenXml.Packaging.ImagePartType.Png);
                    using (FileStream stream = File.OpenRead(imagePath)) {
                        imagePart.FeedData(stream);
                    }

                    string relationshipId = masterPart.GetIdOfPart(imagePart);
                    masterPart.SlideMaster!.CommonSlideData!.Background = new Background(
                        new BackgroundProperties(
                            new A.BlipFill(
                                new A.Blip { Embed = relationshipId },
                                new A.Stretch(new A.FillRectangle()))));
                    masterPart.SlideMaster.Save();

                    slidePart.Slide.CommonSlideData!.Background = new Background(new BackgroundStyleReference { Index = 1001U });
                    slidePart.Slide.Save();
                }

                using (PowerPointPresentation presentation = PowerPointPresentation.Load(filePath, new PowerPointLoadOptions { AccessMode = OfficeIMO.Drawing.DocumentAccessMode.ReadOnly })) {
                    PowerPointFeatureReport report = presentation.InspectFeatures();

                    Assert.Empty(report.FindFeatures("Images"));
                    Assert.Same(report, report.EnsureNoFeatures("Images"));
                }
            } finally {
                if (File.Exists(filePath)) {
                    File.Delete(filePath);
                }
            }
        }

        [Fact]
        public void PowerPointFeatureReport_EmptySlideBackgroundPropertiesDoNotHideInheritedImage() {
            string filePath = Path.Combine(Path.GetTempPath(), Guid.NewGuid() + ".pptx");
            string imagePath = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "Images", "BackgroundImage.png");

            try {
                using (PowerPointPresentation presentation = PowerPointPresentation.Create(filePath)) {
                    presentation.AddSlide().AddTextBox("Inherited background image");
                    presentation.Save();
                }

                using (PresentationDocument document = PresentationDocument.Open(filePath, true)) {
                    SlidePart slidePart = document.PresentationPart!.SlideParts.Single();
                    SlideMasterPart masterPart = slidePart.SlideLayoutPart!.SlideMasterPart!;
                    ImagePart imagePart = masterPart.AddImagePart(DocumentFormat.OpenXml.Packaging.ImagePartType.Png);
                    using (FileStream stream = File.OpenRead(imagePath)) {
                        imagePart.FeedData(stream);
                    }

                    string relationshipId = masterPart.GetIdOfPart(imagePart);
                    masterPart.SlideMaster!.CommonSlideData!.Background = new Background(
                        new BackgroundProperties(
                            new A.BlipFill(
                                new A.Blip { Embed = relationshipId },
                                new A.Stretch(new A.FillRectangle()))));
                    masterPart.SlideMaster.Save();

                    slidePart.Slide.CommonSlideData!.Background = new Background(new BackgroundProperties());
                    slidePart.Slide.Save();
                }

                using (PowerPointPresentation presentation = PowerPointPresentation.Load(filePath, new PowerPointLoadOptions { AccessMode = OfficeIMO.Drawing.DocumentAccessMode.ReadOnly })) {
                    PowerPointFeatureReport report = presentation.InspectFeatures();
                    PowerPointFeatureFinding images = Assert.Single(report.FindFeatures("Images"));

                    Assert.Equal(1, images.Count);
                    Assert.Throws<InvalidOperationException>(() => report.EnsureNoFeatures("Images"));
                }
            } finally {
                if (File.Exists(filePath)) {
                    File.Delete(filePath);
                }
            }
        }

        [Fact]
        public void PowerPointFeatureReport_DetectsShapeFillImages() {
            string filePath = Path.Combine(Path.GetTempPath(), Guid.NewGuid() + ".pptx");
            byte[] pixel = Convert.FromBase64String("iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAQAAAC1HAwCAAAAC0lEQVR4nGMAAQAABQABDQottAAAAABJRU5ErkJggg==");

            try {
                using (PowerPointPresentation presentation = PowerPointPresentation.Create(filePath)) {
                    presentation.AddSlide().AddRectangle(0, 0, 914400, 914400);
                    presentation.Save();
                }

                using (PresentationDocument document = PresentationDocument.Open(filePath, true)) {
                    SlidePart slidePart = document.PresentationPart!.SlideParts.Single();
                    ImagePart imagePart = slidePart.AddImagePart(DocumentFormat.OpenXml.Packaging.ImagePartType.Png);
                    using (var image = new MemoryStream(pixel)) {
                        imagePart.FeedData(image);
                    }

                    string relationshipId = slidePart.GetIdOfPart(imagePart);
                    Shape shape = slidePart.Slide.Descendants<Shape>().Single(shape => shape.ShapeProperties != null);
                    shape.ShapeProperties!.Append(
                        new A.BlipFill(
                            new A.Blip { Embed = relationshipId },
                            new A.Stretch(new A.FillRectangle())));
                    slidePart.Slide.Save();
                }

                using (PowerPointPresentation presentation = PowerPointPresentation.Load(filePath, new PowerPointLoadOptions { AccessMode = OfficeIMO.Drawing.DocumentAccessMode.ReadOnly })) {
                    PowerPointFeatureReport report = presentation.InspectFeatures();
                    PowerPointFeatureFinding images = Assert.Single(report.FindFeatures("Images"));

                    Assert.Equal(1, images.Count);
                    Assert.Throws<InvalidOperationException>(() => report.EnsureNoFeatures("Images"));
                }
            } finally {
                if (File.Exists(filePath)) {
                    File.Delete(filePath);
                }
            }
        }

        [Fact]
        public void PowerPointFeatureReport_DetectsGroupedTextBoxes() {
            string filePath = Path.Combine(Path.GetTempPath(), Guid.NewGuid() + ".pptx");

            try {
                using (PowerPointPresentation presentation = PowerPointPresentation.Create(filePath)) {
                    PowerPointSlide slide = presentation.AddSlide();
                    PowerPointTextBox textBox = slide.AddTextBox("Grouped text");
                    PowerPointAutoShape shape = slide.AddRectangle(914400, 0, 914400, 914400);
                    slide.GroupShapes(new PowerPointShape[] { textBox, shape });
                    presentation.Save();
                }

                using (PowerPointPresentation presentation = PowerPointPresentation.Load(filePath, new PowerPointLoadOptions { AccessMode = OfficeIMO.Drawing.DocumentAccessMode.ReadOnly })) {
                    PowerPointFeatureReport report = presentation.InspectFeatures();
                    PowerPointFeatureFinding textBoxes = Assert.Single(report.FindFeatures("Text boxes"));

                    Assert.Equal(1, textBoxes.Count);
                    Assert.Throws<InvalidOperationException>(() => report.EnsureNoFeatures("Text boxes"));
                }
            } finally {
                if (File.Exists(filePath)) {
                    File.Delete(filePath);
                }
            }
        }

        [Fact]
        public void PowerPointFeatureReport_DetectsGroupedTables() {
            string filePath = Path.Combine(Path.GetTempPath(), Guid.NewGuid() + ".pptx");

            try {
                using (PowerPointPresentation presentation = PowerPointPresentation.Create(filePath)) {
                    PowerPointSlide slide = presentation.AddSlide();
                    PowerPointTable table = slide.AddTable(1, 1);
                    table.GetCell(0, 0).Text = "Grouped";
                    PowerPointAutoShape shape = slide.AddRectangle(914400, 0, 914400, 914400);
                    slide.GroupShapes(new PowerPointShape[] { table, shape });
                    presentation.Save();
                }

                using (PowerPointPresentation presentation = PowerPointPresentation.Load(filePath, new PowerPointLoadOptions { AccessMode = OfficeIMO.Drawing.DocumentAccessMode.ReadOnly })) {
                    PowerPointFeatureReport report = presentation.InspectFeatures();
                    PowerPointFeatureFinding tables = Assert.Single(report.FindFeatures("Tables"));

                    Assert.Equal(1, tables.Count);
                    Assert.Throws<InvalidOperationException>(() => report.EnsureNoFeatures("Tables"));
                }
            } finally {
                if (File.Exists(filePath)) {
                    File.Delete(filePath);
                }
            }
        }

        [Fact]
        public void PowerPointFeatureReport_DetectsInheritedTableMetadata() {
            string filePath = Path.Combine(Path.GetTempPath(), Guid.NewGuid() + ".pptx");

            try {
                using (PowerPointPresentation presentation = PowerPointPresentation.Create(filePath)) {
                    PowerPointSlide slide = presentation.AddSlide();
                    PowerPointTable table = slide.AddTable(1, 1);
                    table.GetCell(0, 0).Text = "Inherited";
                    presentation.Save();
                }

                using (PresentationDocument document = PresentationDocument.Open(filePath, true)) {
                    SlidePart slidePart = document.PresentationPart!.SlideParts.Single();
                    SlideLayoutPart layoutPart = slidePart.SlideLayoutPart!;
                    ShapeTree slideTree = slidePart.Slide.CommonSlideData!.ShapeTree!;
                    ShapeTree layoutTree = layoutPart.SlideLayout!.CommonSlideData!.ShapeTree!;
                    GraphicFrame tableFrame = slideTree.Elements<GraphicFrame>()
                        .Single(frame => frame.Graphic?.GraphicData?.GetFirstChild<A.Table>() != null);
                    tableFrame.Remove();
                    layoutTree.Append(tableFrame);
                    slidePart.Slide.Save();
                    layoutPart.SlideLayout.Save();
                }

                using (PowerPointPresentation presentation = PowerPointPresentation.Load(filePath, new PowerPointLoadOptions { AccessMode = OfficeIMO.Drawing.DocumentAccessMode.ReadOnly })) {
                    PowerPointFeatureReport report = presentation.InspectFeatures();
                    PowerPointFeatureFinding tables = Assert.Single(report.FindFeatures("Tables"));
                    PowerPointFeatureFinding metadata = Assert.Single(report.FindFeatures("Table style metadata"));

                    Assert.Equal(1, tables.Count);
                    Assert.Equal(1, metadata.Count);
                    Assert.Throws<InvalidOperationException>(() => report.EnsureNoFeatures("Table style metadata"));
                }
            } finally {
                if (File.Exists(filePath)) {
                    File.Delete(filePath);
                }
            }
        }

        [Fact]
        public void PowerPointFeatureReport_DetectsInheritedLayoutPictures() {
            string filePath = Path.Combine(Path.GetTempPath(), Guid.NewGuid() + ".pptx");
            byte[] pixel = Convert.FromBase64String("iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAQAAAC1HAwCAAAAC0lEQVR4nGMAAQAABQABDQottAAAAABJRU5ErkJggg==");

            try {
                using (PowerPointPresentation presentation = PowerPointPresentation.Create(filePath)) {
                    presentation.AddSlide().AddTextBox("Slide inherits the layout logo");
                    presentation.Save();
                }

                using (PresentationDocument document = PresentationDocument.Open(filePath, true)) {
                    SlidePart slidePart = document.PresentationPart!.SlideParts.Single();
                    SlideLayoutPart layoutPart = slidePart.SlideLayoutPart!;
                    ImagePart imagePart = layoutPart.AddImagePart(DocumentFormat.OpenXml.Packaging.ImagePartType.Png);
                    using (var image = new MemoryStream(pixel)) {
                        imagePart.FeedData(image);
                    }

                    string relationshipId = layoutPart.GetIdOfPart(imagePart);
                    ShapeTree tree = layoutPart.SlideLayout!.CommonSlideData!.ShapeTree!;
                    uint shapeId = tree.Descendants<NonVisualDrawingProperties>()
                        .Select(properties => properties.Id?.Value ?? 0U)
                        .DefaultIfEmpty(0U)
                        .Max() + 1U;
                    tree.Append(new Picture(
                        new NonVisualPictureProperties(
                            new NonVisualDrawingProperties { Id = shapeId, Name = "Layout Logo" },
                            new NonVisualPictureDrawingProperties(new A.PictureLocks { NoChangeAspect = true }),
                            new ApplicationNonVisualDrawingProperties()),
                        new BlipFill(
                            new A.Blip { Embed = relationshipId },
                            new A.Stretch(new A.FillRectangle())),
                        new ShapeProperties(
                            new A.Transform2D(new A.Offset { X = 0, Y = 0 }, new A.Extents { Cx = 914400, Cy = 914400 }),
                            new A.PresetGeometry(new A.AdjustValueList()) { Preset = A.ShapeTypeValues.Rectangle })));
                    layoutPart.SlideLayout.Save();
                }

                using (PowerPointPresentation presentation = PowerPointPresentation.Load(filePath, new PowerPointLoadOptions { AccessMode = OfficeIMO.Drawing.DocumentAccessMode.ReadOnly })) {
                    PowerPointFeatureReport report = presentation.InspectFeatures();
                    PowerPointFeatureFinding images = Assert.Single(report.FindFeatures("Images"));

                    Assert.Equal(1, images.Count);
                    Assert.Throws<InvalidOperationException>(() => report.EnsureNoFeatures("Images"));
                }
            } finally {
                if (File.Exists(filePath)) {
                    File.Delete(filePath);
                }
            }
        }

        [Fact]
        public void PowerPointFeatureReport_DetectsInheritedLayoutTextBoxes() {
            string filePath = Path.Combine(Path.GetTempPath(), Guid.NewGuid() + ".pptx");

            try {
                using (PowerPointPresentation presentation = PowerPointPresentation.Create(filePath)) {
                    presentation.AddSlide().AddTextBox("Layout footer");
                    presentation.Save();
                }

                using (PresentationDocument document = PresentationDocument.Open(filePath, true)) {
                    SlidePart slidePart = document.PresentationPart!.SlideParts.Single();
                    SlideLayoutPart layoutPart = slidePart.SlideLayoutPart!;
                    ShapeTree slideTree = slidePart.Slide.CommonSlideData!.ShapeTree!;
                    ShapeTree layoutTree = layoutPart.SlideLayout!.CommonSlideData!.ShapeTree!;
                    Shape textBox = slideTree.Elements<Shape>()
                        .Single(shape => shape.TextBody != null);
                    textBox.Remove();
                    layoutTree.Append(textBox);
                    slidePart.Slide.Save();
                    layoutPart.SlideLayout.Save();
                }

                using (PowerPointPresentation presentation = PowerPointPresentation.Load(filePath, new PowerPointLoadOptions { AccessMode = OfficeIMO.Drawing.DocumentAccessMode.ReadOnly })) {
                    PowerPointFeatureReport report = presentation.InspectFeatures();
                    PowerPointFeatureFinding textBoxes = Assert.Single(report.FindFeatures("Text boxes"));

                    Assert.Equal(1, textBoxes.Count);
                    Assert.Throws<InvalidOperationException>(() => report.EnsureNoFeatures("Text boxes"));
                }
            } finally {
                if (File.Exists(filePath)) {
                    File.Delete(filePath);
                }
            }
        }

        [Fact]
        public void PowerPointFeatureReport_DetectsInheritedLayoutMedia() {
            string filePath = Path.Combine(Path.GetTempPath(), Guid.NewGuid() + ".pptx");

            try {
                using (PowerPointPresentation presentation = PowerPointPresentation.Create(filePath)) {
                    PowerPointSlide slide = presentation.AddSlide();
                    using var media = new MemoryStream(new byte[] { 1, 2, 3, 4, 5 });
                    slide.AddAudio(media, "audio/mpeg", ".mp3");
                    presentation.Save();
                }

                using (PresentationDocument document = PresentationDocument.Open(filePath, true)) {
                    SlidePart slidePart = document.PresentationPart!.SlideParts.Single();
                    SlideLayoutPart layoutPart = slidePart.SlideLayoutPart!;
                    ShapeTree slideTree = slidePart.Slide.CommonSlideData!.ShapeTree!;
                    ShapeTree layoutTree = layoutPart.SlideLayout!.CommonSlideData!.ShapeTree!;
                    Picture mediaPicture = slideTree.Elements<Picture>()
                        .Single(picture => picture.NonVisualPictureProperties?
                            .ApplicationNonVisualDrawingProperties?
                            .GetFirstChild<A.AudioFromFile>() != null);
                    mediaPicture.Remove();
                    layoutTree.Append(mediaPicture);
                    slidePart.Slide.Save();
                    layoutPart.SlideLayout.Save();
                }

                using (PowerPointPresentation presentation = PowerPointPresentation.Load(filePath, new PowerPointLoadOptions { AccessMode = OfficeIMO.Drawing.DocumentAccessMode.ReadOnly })) {
                    PowerPointFeatureReport report = presentation.InspectFeatures();
                    PowerPointFeatureFinding media = Assert.Single(report.FindFeatures("Audio and video"));

                    Assert.Equal(1, media.Count);
                    Assert.Throws<InvalidOperationException>(() => report.EnsureNoFeatures("Audio and video"));
                    Assert.Same(report, report.EnsureNoAdvancedFeatures());
                }
            } finally {
                if (File.Exists(filePath)) {
                    File.Delete(filePath);
                }
            }
        }

        [Fact]
        public void PowerPointFeatureReport_DetectsRichNotesSlidePictures() {
            string filePath = Path.Combine(Path.GetTempPath(), Guid.NewGuid() + ".pptx");
            byte[] pixel = Convert.FromBase64String("iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAQAAAC1HAwCAAAAC0lEQVR4nGMAAQAABQABDQottAAAAABJRU5ErkJggg==");

            try {
                using (PowerPointPresentation presentation = PowerPointPresentation.Create(filePath)) {
                    PowerPointSlide slide = presentation.AddSlide();
                    slide.AddTextBox("Notes with an image");
                    slide.Notes.Text = "Speaker text";
                    presentation.Save();
                }

                using (PresentationDocument document = PresentationDocument.Open(filePath, true)) {
                    NotesSlidePart notesPart = document.PresentationPart!.SlideParts.Single().NotesSlidePart!;
                    ImagePart imagePart = notesPart.AddImagePart(DocumentFormat.OpenXml.Packaging.ImagePartType.Png);
                    using (var image = new MemoryStream(pixel)) {
                        imagePart.FeedData(image);
                    }

                    string relationshipId = notesPart.GetIdOfPart(imagePart);
                    ShapeTree tree = notesPart.NotesSlide!.CommonSlideData!.ShapeTree!;
                    uint shapeId = tree.Descendants<NonVisualDrawingProperties>()
                        .Select(properties => properties.Id?.Value ?? 0U)
                        .DefaultIfEmpty(0U)
                        .Max() + 1U;
                    tree.Append(new Picture(
                        new NonVisualPictureProperties(
                            new NonVisualDrawingProperties { Id = shapeId, Name = "Notes Picture" },
                            new NonVisualPictureDrawingProperties(new A.PictureLocks { NoChangeAspect = true }),
                            new ApplicationNonVisualDrawingProperties()),
                        new BlipFill(
                            new A.Blip { Embed = relationshipId },
                            new A.Stretch(new A.FillRectangle())),
                        new ShapeProperties(
                            new A.Transform2D(new A.Offset { X = 0, Y = 0 }, new A.Extents { Cx = 914400, Cy = 914400 }),
                            new A.PresetGeometry(new A.AdjustValueList()) { Preset = A.ShapeTypeValues.Rectangle })));
                    notesPart.NotesSlide.Save();
                }

                using (PowerPointPresentation presentation = PowerPointPresentation.Load(filePath, new PowerPointLoadOptions { AccessMode = OfficeIMO.Drawing.DocumentAccessMode.ReadOnly })) {
                    PowerPointFeatureReport report = presentation.InspectFeatures();
                    PowerPointFeatureFinding notes = Assert.Single(report.FindFeatures("Speaker notes"));
                    PowerPointFeatureFinding richNotes = Assert.Single(report.FindFeatures("Rich notes content"));

                    Assert.Equal(1, notes.Count);
                    Assert.Equal(PowerPointFeatureSupportLevel.Preserved, richNotes.SupportLevel);
                    Assert.Equal(1, richNotes.Count);
                    Assert.Contains(richNotes.Details, detail => detail.Contains("picture", StringComparison.OrdinalIgnoreCase));
                    Assert.Throws<InvalidOperationException>(() => report.EnsureNoAdvancedFeatures());
                }
            } finally {
                if (File.Exists(filePath)) {
                    File.Delete(filePath);
                }
            }
        }

        [Fact]
        public void PowerPointFeatureReport_DetectsRichNotesSlideCharts() {
            string filePath = Path.Combine(Path.GetTempPath(), Guid.NewGuid() + ".pptx");

            try {
                using (PowerPointPresentation presentation = PowerPointPresentation.Create(filePath)) {
                    PowerPointSlide slide = presentation.AddSlide();
                    slide.AddTextBox("Notes with a chart");
                    slide.Notes.Text = "Speaker text";
                    presentation.Save();
                }

                using (PresentationDocument document = PresentationDocument.Open(filePath, true)) {
                    NotesSlidePart notesPart = document.PresentationPart!.SlideParts.Single().NotesSlidePart!;
                    ShapeTree tree = notesPart.NotesSlide!.CommonSlideData!.ShapeTree!;
                    uint shapeId = tree.Descendants<NonVisualDrawingProperties>()
                        .Select(properties => properties.Id?.Value ?? 0U)
                        .DefaultIfEmpty(0U)
                        .Max() + 1U;
                    tree.Append(new GraphicFrame(
                        new NonVisualGraphicFrameProperties(
                            new NonVisualDrawingProperties { Id = shapeId, Name = "Notes Chart" },
                            new NonVisualGraphicFrameDrawingProperties(),
                            new ApplicationNonVisualDrawingProperties()),
                        new Transform(new A.Offset { X = 0, Y = 0 }, new A.Extents { Cx = 914400, Cy = 914400 }),
                        new A.Graphic(new A.GraphicData(new C.ChartReference { Id = "rIdChart" }) {
                            Uri = "http://schemas.openxmlformats.org/drawingml/2006/chart"
                        })));
                    notesPart.NotesSlide.Save();
                }

                using (PowerPointPresentation presentation = PowerPointPresentation.Load(filePath, new PowerPointLoadOptions { AccessMode = OfficeIMO.Drawing.DocumentAccessMode.ReadOnly })) {
                    PowerPointFeatureReport report = presentation.InspectFeatures();
                    PowerPointFeatureFinding richNotes = Assert.Single(report.FindFeatures("Rich notes content"));

                    Assert.Equal(PowerPointFeatureSupportLevel.Preserved, richNotes.SupportLevel);
                    Assert.Equal(1, richNotes.Count);
                    Assert.Contains(richNotes.Details, detail => detail.Contains("chart", StringComparison.OrdinalIgnoreCase));
                    Assert.Throws<InvalidOperationException>(() => report.EnsureNoAdvancedFeatures());
                }
            } finally {
                if (File.Exists(filePath)) {
                    File.Delete(filePath);
                }
            }
        }

        [Fact]
        public void PowerPointFeatureReport_DetectsRichNotesSlideExtraTextShapes() {
            string filePath = Path.Combine(Path.GetTempPath(), Guid.NewGuid() + ".pptx");

            try {
                using (PowerPointPresentation presentation = PowerPointPresentation.Create(filePath)) {
                    PowerPointSlide slide = presentation.AddSlide();
                    slide.AddTextBox("Notes with extra text");
                    slide.Notes.Text = "Speaker text";
                    presentation.Save();
                }

                using (PresentationDocument document = PresentationDocument.Open(filePath, true)) {
                    NotesSlidePart notesPart = document.PresentationPart!.SlideParts.Single().NotesSlidePart!;
                    ShapeTree tree = notesPart.NotesSlide!.CommonSlideData!.ShapeTree!;
                    uint shapeId = tree.Descendants<NonVisualDrawingProperties>()
                        .Select(properties => properties.Id?.Value ?? 0U)
                        .DefaultIfEmpty(0U)
                        .Max() + 1U;
                    tree.Append(new Shape(
                        new NonVisualShapeProperties(
                            new NonVisualDrawingProperties { Id = shapeId, Name = "Notes Extra Text" },
                            new NonVisualShapeDrawingProperties(),
                            new ApplicationNonVisualDrawingProperties()),
                        new ShapeProperties(
                            new A.Transform2D(new A.Offset { X = 0, Y = 0 }, new A.Extents { Cx = 914400, Cy = 914400 }),
                            new A.PresetGeometry(new A.AdjustValueList()) { Preset = A.ShapeTypeValues.Rectangle }),
                        new TextBody(
                            new A.BodyProperties(),
                            new A.ListStyle(),
                            new A.Paragraph(new A.Run(new A.Text("Imported notes callout"))))));
                    notesPart.NotesSlide.Save();
                }

                using (PowerPointPresentation presentation = PowerPointPresentation.Load(filePath, new PowerPointLoadOptions { AccessMode = OfficeIMO.Drawing.DocumentAccessMode.ReadOnly })) {
                    PowerPointFeatureReport report = presentation.InspectFeatures();
                    PowerPointFeatureFinding richNotes = Assert.Single(report.FindFeatures("Rich notes content"));

                    Assert.Equal(PowerPointFeatureSupportLevel.Preserved, richNotes.SupportLevel);
                    Assert.Equal(1, richNotes.Count);
                    Assert.Contains(richNotes.Details, detail => detail.Contains("text shape", StringComparison.OrdinalIgnoreCase));
                    Assert.Throws<InvalidOperationException>(() => report.EnsureNoAdvancedFeatures());
                }
            } finally {
                if (File.Exists(filePath)) {
                    File.Delete(filePath);
                }
            }
        }

        [Fact]
        public void PowerPointFeatureReport_DetectsRichNotesMasterExtraTextShapes() {
            string filePath = Path.Combine(Path.GetTempPath(), Guid.NewGuid() + ".pptx");

            try {
                using (PowerPointPresentation presentation = PowerPointPresentation.Create(filePath)) {
                    PowerPointSlide slide = presentation.AddSlide();
                    slide.AddTextBox("Notes master with extra text");
                    slide.Notes.Text = "Speaker text";
                    presentation.Save();
                }

                using (PresentationDocument document = PresentationDocument.Open(filePath, true)) {
                    NotesMasterPart notesMasterPart = document.PresentationPart!.NotesMasterPart!;
                    ShapeTree tree = notesMasterPart.NotesMaster!.CommonSlideData!.ShapeTree!;
                    uint shapeId = tree.Descendants<NonVisualDrawingProperties>()
                        .Select(properties => properties.Id?.Value ?? 0U)
                        .DefaultIfEmpty(0U)
                        .Max() + 1U;
                    tree.Append(new Shape(
                        new NonVisualShapeProperties(
                            new NonVisualDrawingProperties { Id = shapeId, Name = "Notes Master Extra Text" },
                            new NonVisualShapeDrawingProperties(),
                            new ApplicationNonVisualDrawingProperties()),
                        new ShapeProperties(
                            new A.Transform2D(new A.Offset { X = 0, Y = 0 }, new A.Extents { Cx = 914400, Cy = 914400 }),
                            new A.PresetGeometry(new A.AdjustValueList()) { Preset = A.ShapeTypeValues.Rectangle }),
                        new TextBody(
                            new A.BodyProperties(),
                            new A.ListStyle(),
                            new A.Paragraph(new A.Run(new A.Text("Inherited notes callout"))))));
                    notesMasterPart.NotesMaster.Save();
                }

                using (PowerPointPresentation presentation = PowerPointPresentation.Load(filePath, new PowerPointLoadOptions { AccessMode = OfficeIMO.Drawing.DocumentAccessMode.ReadOnly })) {
                    PowerPointFeatureReport report = presentation.InspectFeatures();
                    PowerPointFeatureFinding richNotes = Assert.Single(report.FindFeatures("Rich notes content"));

                    Assert.Equal(PowerPointFeatureSupportLevel.Preserved, richNotes.SupportLevel);
                    Assert.Equal(1, richNotes.Count);
                    Assert.Contains(richNotes.Details, detail => detail.Contains("notes master", StringComparison.OrdinalIgnoreCase)
                        && detail.Contains("text shape", StringComparison.OrdinalIgnoreCase));
                    Assert.Throws<InvalidOperationException>(() => report.EnsureNoAdvancedFeatures());
                }
            } finally {
                if (File.Exists(filePath)) {
                    File.Delete(filePath);
                }
            }
        }

        [Fact]
        public void PowerPointFeatureReport_DoesNotTreatSignatureNamedMediaAsDigitalSignature() {
            string filePath = Path.Combine(Path.GetTempPath(), Guid.NewGuid() + ".pptx");

            try {
                using (PowerPointPresentation presentation = PowerPointPresentation.Create(filePath)) {
                    presentation.AddSlide().AddTextBox("Signature named media");
                    presentation.Save();
                }

                using (PresentationDocument document = PresentationDocument.Open(filePath, true)) {
                    SlidePart slidePart = document.PresentationPart!.SlideParts.Single();
                    AddExtendedPart(
                        slidePart,
                        "http://schemas.openxmlformats.org/officeDocument/2006/relationships/image",
                        "image/png",
                        "signature.png",
                        new byte[] { 137, 80, 78, 71 });
                }

                using (PowerPointPresentation presentation = PowerPointPresentation.Load(filePath, new PowerPointLoadOptions { AccessMode = OfficeIMO.Drawing.DocumentAccessMode.ReadOnly })) {
                    PowerPointFeatureReport report = presentation.InspectFeatures();

                    Assert.Empty(report.FindFeatures("Digital signatures"));
                    Assert.Same(report, report.EnsureNoUnsupportedFeatures());
                }
            } finally {
                if (File.Exists(filePath)) {
                    File.Delete(filePath);
                }
            }
        }

        [Fact]
        public void PowerPointFeatureReport_DoesNotFlagGroupedMediaPlaybackTimingAsAdvanced() {
            string filePath = Path.Combine(Path.GetTempPath(), Guid.NewGuid() + ".pptx");

            try {
                using (PowerPointPresentation presentation = PowerPointPresentation.Create(filePath)) {
                    PowerPointSlide slide = presentation.AddSlide();
                    using var media = new MemoryStream(new byte[] { 1, 2, 3, 4, 5 });
                    PowerPointMedia audio = slide.AddAudio(media, "audio/mpeg", ".mp3");
                    PowerPointAutoShape shape = slide.AddRectangle(914400, 0, 914400, 914400);
                    slide.GroupShapes(new PowerPointShape[] { audio, shape });
                    presentation.Save();
                }

                using (PowerPointPresentation presentation = PowerPointPresentation.Load(filePath, new PowerPointLoadOptions { AccessMode = OfficeIMO.Drawing.DocumentAccessMode.ReadOnly })) {
                    PowerPointFeatureReport report = presentation.InspectFeatures();
                    PowerPointFeatureFinding media = Assert.Single(report.FindFeatures("Audio and video"));

                    Assert.Equal(1, media.Count);
                    Assert.DoesNotContain(report.PreservedFeatures, feature => feature.Name == "Animations and timing");
                    Assert.Same(report, report.EnsureNoAdvancedFeatures());
                }
            } finally {
                if (File.Exists(filePath)) {
                    File.Delete(filePath);
                }
            }
        }

        [Fact]
        public void PowerPointFeatureReport_DetectsGroupedSmartArt() {
            string filePath = Path.Combine(Path.GetTempPath(), Guid.NewGuid() + ".pptx");

            try {
                using (PowerPointPresentation presentation = PowerPointPresentation.Create(filePath)) {
                    PowerPointSlide slide = presentation.AddSlide();
                    PowerPointSmartArt smartArt = slide.AddSmartArt();
                    PowerPointAutoShape shape = slide.AddRectangle(914400, 0, 914400, 914400);
                    slide.GroupShapes(new PowerPointShape[] { smartArt, shape });
                    presentation.Save();
                }

                using (PowerPointPresentation presentation = PowerPointPresentation.Load(filePath, new PowerPointLoadOptions { AccessMode = OfficeIMO.Drawing.DocumentAccessMode.ReadOnly })) {
                    PowerPointFeatureReport report = presentation.InspectFeatures();
                    PowerPointFeatureFinding smartArt = Assert.Single(report.FindFeatures("SmartArt"));

                    Assert.True(smartArt.Count > 0);
                    Assert.Throws<InvalidOperationException>(() => report.EnsureNoFeatures("SmartArt"));
                }
            } finally {
                if (File.Exists(filePath)) {
                    File.Delete(filePath);
                }
            }
        }

        [Fact]
        public void PowerPointFeatureReport_DoesNotBlockZeroCountEditableFeatures() {
            string filePath = Path.Combine(Path.GetTempPath(), Guid.NewGuid() + ".pptx");

            try {
                using (PowerPointPresentation presentation = PowerPointPresentation.Create(filePath)) {
                    presentation.AddSlide().AddTextBox("No tables here");
                    presentation.Save();
                }

                using (PowerPointPresentation presentation = PowerPointPresentation.Load(filePath, new PowerPointLoadOptions { AccessMode = OfficeIMO.Drawing.DocumentAccessMode.ReadOnly })) {
                    PowerPointFeatureReport report = presentation.InspectFeatures();
                    PowerPointFeatureFinding tables = Assert.Single(report.FindFeatures("Tables"));

                    Assert.Equal(PowerPointFeatureSupportLevel.Editable, tables.SupportLevel);
                    Assert.Equal(0, tables.Count);
                    Assert.Same(report, report.EnsureNoFeatures("Tables"));
                }
            } finally {
                if (File.Exists(filePath)) {
                    File.Delete(filePath);
                }
            }
        }

        [Fact]
        public void PowerPointTableCells_IncludeLanguageAwareRunDefaults() {
            string filePath = Path.Combine(Path.GetTempPath(), Guid.NewGuid() + ".pptx");

            try {
                using (PowerPointPresentation presentation = PowerPointPresentation.Create(filePath)) {
                    PowerPointTable table = presentation.AddSlide().AddTable(1, 1);
                    table.GetCell(0, 0).Text = "Header";
                    presentation.Save();
                }

                using (PresentationDocument document = PresentationDocument.Open(filePath, false)) {
                    A.TableCell cell = document.PresentationPart!.SlideParts.First().Slide.Descendants<A.TableCell>().First();
                    A.Paragraph paragraph = cell.TextBody!.Elements<A.Paragraph>().First();
                    A.RunProperties runProperties = paragraph.Elements<A.Run>().First().RunProperties!;
                    A.EndParagraphRunProperties endProperties = paragraph.GetFirstChild<A.EndParagraphRunProperties>()!;

                    Assert.Equal("en-US", runProperties.Language?.Value);
                    Assert.Equal("en-US", endProperties.Language?.Value);
                }
            } finally {
                if (File.Exists(filePath)) {
                    File.Delete(filePath);
                }
            }
        }

        [Fact]
        public void PowerPointTableCellText_ReplacesAllExistingParagraphs() {
            string filePath = Path.Combine(Path.GetTempPath(), Guid.NewGuid() + ".pptx");

            try {
                using (PowerPointPresentation presentation = PowerPointPresentation.Create(filePath)) {
                    PowerPointTable table = presentation.AddSlide().AddTable(1, 1);
                    presentation.Save();
                }

                using (PresentationDocument document = PresentationDocument.Open(filePath, true)) {
                    A.TableCell cell = document.PresentationPart!.SlideParts.First().Slide.Descendants<A.TableCell>().First();
                    cell.TextBody!.Append(
                        new A.Paragraph(new A.Run(new A.Text("Stale one"))),
                        new A.Paragraph(new A.Run(new A.Text("Stale two"))));
                }

                using (PowerPointPresentation presentation = PowerPointPresentation.Load(filePath)) {
                    PowerPointTableCell cell = presentation.Slides.Single().Tables.Single().GetCell(0, 0);
                    cell.Text = "Fresh";
                    presentation.Save();
                }

                using (PowerPointPresentation presentation = PowerPointPresentation.Load(filePath, new PowerPointLoadOptions { AccessMode = OfficeIMO.Drawing.DocumentAccessMode.ReadOnly })) {
                    PowerPointTableCell cell = presentation.Slides.Single().Tables.Single().GetCell(0, 0);
                    Assert.Equal("Fresh", cell.Text);
                }
            } finally {
                if (File.Exists(filePath)) {
                    File.Delete(filePath);
                }
            }
        }

        [Fact]
        public void PowerPointTableCellAddParagraph_ReplacesInitialEmptyPlaceholder() {
            string filePath = Path.Combine(Path.GetTempPath(), Guid.NewGuid() + ".pptx");

            try {
                using (PowerPointPresentation presentation = PowerPointPresentation.Create(filePath)) {
                    PowerPointTableCell cell = presentation.AddSlide().AddTable(1, 1).GetCell(0, 0);
                    PowerPointParagraph paragraph = cell.AddParagraph("Fresh");

                    Assert.Single(cell.Paragraphs);
                    Assert.Same(paragraph.Paragraph, cell.Paragraphs.Single().Paragraph);
                    presentation.Save();
                }

                using (PresentationDocument document = PresentationDocument.Open(filePath, false)) {
                    A.TableCell cell = document.PresentationPart!.SlideParts.First().Slide.Descendants<A.TableCell>().First();
                    A.Paragraph paragraph = Assert.Single(cell.TextBody!.Elements<A.Paragraph>());

                    Assert.Equal("Fresh", paragraph.Elements<A.Run>().Single().Text!.Text);
                }
            } finally {
                if (File.Exists(filePath)) {
                    File.Delete(filePath);
                }
            }
        }

        [Fact]
        public void PowerPointTableCellText_DropsStaleHyperlinkWhenReplacingText() {
            string filePath = Path.Combine(Path.GetTempPath(), Guid.NewGuid() + ".pptx");

            try {
                using (PowerPointPresentation presentation = PowerPointPresentation.Create(filePath)) {
                    PowerPointTable table = presentation.AddSlide().AddTable(1, 1);
                    table.GetCell(0, 0).Text = "Linked";
                    presentation.Save();
                }

                using (PresentationDocument document = PresentationDocument.Open(filePath, true)) {
                    SlidePart slidePart = document.PresentationPart!.SlideParts.Single();
                    HyperlinkRelationship relationship = slidePart.AddHyperlinkRelationship(new Uri("https://example.com/old"), true);
                    A.Run run = slidePart.Slide.Descendants<A.TableCell>().First().TextBody!.Descendants<A.Run>().First();
                    run.RunProperties ??= new A.RunProperties();
                    run.RunProperties.Append(new A.HyperlinkOnClick { Id = relationship.Id });
                    slidePart.Slide.Save();
                }

                using (PowerPointPresentation presentation = PowerPointPresentation.Load(filePath)) {
                    PowerPointTableCell cell = presentation.Slides.Single().Tables.Single().GetCell(0, 0);
                    cell.Text = "Fresh";
                    presentation.Save();
                }

                using (PresentationDocument document = PresentationDocument.Open(filePath, false)) {
                    A.Run run = document.PresentationPart!.SlideParts.Single().Slide.Descendants<A.TableCell>().First()
                        .TextBody!.Descendants<A.Run>().Single();

                    Assert.Equal("Fresh", run.Text!.Text);
                    Assert.Null(run.RunProperties?.GetFirstChild<A.HyperlinkOnClick>());
                }
            } finally {
                if (File.Exists(filePath)) {
                    File.Delete(filePath);
                }
            }
        }

        [Fact]
        public void PowerPointTableCellText_DropsStaleEndParagraphHyperlinkWhenReplacingText() {
            string filePath = Path.Combine(Path.GetTempPath(), Guid.NewGuid() + ".pptx");

            try {
                using (PowerPointPresentation presentation = PowerPointPresentation.Create(filePath)) {
                    PowerPointTable table = presentation.AddSlide().AddTable(1, 1);
                    table.GetCell(0, 0).Text = "Linked";
                    presentation.Save();
                }

                using (PresentationDocument document = PresentationDocument.Open(filePath, true)) {
                    SlidePart slidePart = document.PresentationPart!.SlideParts.Single();
                    HyperlinkRelationship relationship = slidePart.AddHyperlinkRelationship(new Uri("https://example.com/end"), true);
                    A.Paragraph paragraph = slidePart.Slide.Descendants<A.TableCell>().First().TextBody!.Elements<A.Paragraph>().First();
                    A.EndParagraphRunProperties endProperties = paragraph.GetFirstChild<A.EndParagraphRunProperties>()
                        ?? paragraph.AppendChild(new A.EndParagraphRunProperties());
                    endProperties.Append(new A.HyperlinkOnClick { Id = relationship.Id });
                    slidePart.Slide.Save();
                }

                using (PowerPointPresentation presentation = PowerPointPresentation.Load(filePath)) {
                    PowerPointTableCell cell = presentation.Slides.Single().Tables.Single().GetCell(0, 0);
                    cell.Text = "Fresh";
                    presentation.Save();
                }

                using (PresentationDocument document = PresentationDocument.Open(filePath, false)) {
                    A.Paragraph paragraph = document.PresentationPart!.SlideParts.Single().Slide.Descendants<A.TableCell>().First()
                        .TextBody!.Elements<A.Paragraph>().Single();

                    Assert.Null(paragraph.GetFirstChild<A.EndParagraphRunProperties>()?.GetFirstChild<A.HyperlinkOnClick>());
                }

                using (PowerPointPresentation presentation = PowerPointPresentation.Load(filePath, new PowerPointLoadOptions { AccessMode = OfficeIMO.Drawing.DocumentAccessMode.ReadOnly })) {
                    PowerPointFeatureReport report = presentation.InspectFeatures();

                    Assert.Empty(report.FindFeatures("External relationships"));
                    Assert.Same(report, report.EnsureNoFeatures("External relationships"));
                }
            } finally {
                if (File.Exists(filePath)) {
                    File.Delete(filePath);
                }
            }
        }

        [Fact]
        public void PowerPointFeatureReport_IgnoresUnreferencedHyperlinkRelationships() {
            string filePath = Path.Combine(Path.GetTempPath(), Guid.NewGuid() + ".pptx");

            try {
                using (PowerPointPresentation presentation = PowerPointPresentation.Create(filePath)) {
                    presentation.AddSlide().AddTextBox("Stale hyperlink relationship");
                    presentation.Save();
                }

                using (PresentationDocument document = PresentationDocument.Open(filePath, true)) {
                    SlidePart slidePart = document.PresentationPart!.SlideParts.Single();
                    slidePart.AddHyperlinkRelationship(new Uri("https://example.com/stale"), true);
                    slidePart.Slide.Save();
                }

                using (PowerPointPresentation presentation = PowerPointPresentation.Load(filePath, new PowerPointLoadOptions { AccessMode = OfficeIMO.Drawing.DocumentAccessMode.ReadOnly })) {
                    PowerPointFeatureReport report = presentation.InspectFeatures();

                    Assert.Empty(report.FindFeatures("External relationships"));
                    Assert.Same(report, report.EnsureNoFeatures("External relationships"));
                }
            } finally {
                if (File.Exists(filePath)) {
                    File.Delete(filePath);
                }
            }
        }

        [Fact]
        public void PowerPointFeatureReport_DoesNotTreatChartWorkbookAsEmbeddedPackage() {
            string filePath = Path.Combine(Path.GetTempPath(), Guid.NewGuid() + ".pptx");

            try {
                using (PowerPointPresentation presentation = PowerPointPresentation.Create(filePath)) {
                    presentation.AddSlide().AddChart();
                    presentation.Save();
                }

                using (PowerPointPresentation presentation = PowerPointPresentation.Load(filePath, new PowerPointLoadOptions { AccessMode = OfficeIMO.Drawing.DocumentAccessMode.ReadOnly })) {
                    PowerPointFeatureReport report = presentation.InspectFeatures();

                    Assert.Contains(report.PartiallyEditableFeatures, feature => feature.Name == "Charts" && feature.Count == 1
                        && feature.Details.Any(detail => detail.Contains("Microsoft_Excel_Worksheet", StringComparison.OrdinalIgnoreCase)));
                    Assert.DoesNotContain(report.PreservedFeatures, feature => feature.Name == "Embedded packages");
                    Assert.False(report.HasAdvancedFeatures);
                }
            } finally {
                if (File.Exists(filePath)) {
                    File.Delete(filePath);
                }
            }
        }

        [Fact]
        public void PowerPointFeatureReport_TreatsUnsafeChartOwnedPackageAsAdvanced() {
            string filePath = Path.Combine(Path.GetTempPath(), Guid.NewGuid() + ".pptx");

            try {
                using (PowerPointPresentation presentation = PowerPointPresentation.Create(filePath)) {
                    presentation.AddSlide().AddChart();
                    presentation.Save();
                }

                using (PresentationDocument document = PresentationDocument.Open(filePath, true)) {
                    ChartPart chartPart = document.PresentationPart!.SlideParts.Single().ChartParts.Single();
                    if (chartPart.EmbeddedPackagePart is EmbeddedPackagePart existingPackage) {
                        chartPart.DeletePart(existingPackage);
                    }
                    EmbeddedPackagePart package = chartPart.AddEmbeddedPackagePart("application/vnd.ms-excel.sheet.macroEnabled.12");
                    using var content = new MemoryStream(new byte[] { 1, 2, 3, 4 });
                    package.FeedData(content);
                }

                using (PowerPointPresentation presentation = PowerPointPresentation.Load(filePath, new PowerPointLoadOptions { AccessMode = OfficeIMO.Drawing.DocumentAccessMode.ReadOnly })) {
                    PowerPointFeatureReport report = presentation.InspectFeatures();
                    PowerPointFeatureFinding embedded = Assert.Single(report.FindFeatures("Embedded packages"));

                    Assert.Equal(PowerPointFeatureSupportLevel.Preserved, embedded.SupportLevel);
                    Assert.Equal(1, embedded.Count);
                    Assert.Throws<InvalidOperationException>(() => report.EnsureNoAdvancedFeatures());
                }
            } finally {
                if (File.Exists(filePath)) {
                    File.Delete(filePath);
                }
            }
        }

        [Fact]
        public void PowerPointFeatureReport_TreatsUnreferencedChartWorkbookAsAdvanced() {
            string filePath = Path.Combine(Path.GetTempPath(), Guid.NewGuid() + ".pptx");

            try {
                using (PowerPointPresentation presentation = PowerPointPresentation.Create(filePath)) {
                    presentation.AddSlide().AddChart();
                    presentation.Save();
                }

                using (PresentationDocument document = PresentationDocument.Open(filePath, true)) {
                    ChartPart chartPart = document.PresentationPart!.SlideParts.Single().ChartParts.Single();
                    chartPart.ChartSpace!.RemoveAllChildren<C.ExternalData>();
                    chartPart.ChartSpace.Save();
                }

                using (PowerPointPresentation presentation = PowerPointPresentation.Load(filePath, new PowerPointLoadOptions { AccessMode = OfficeIMO.Drawing.DocumentAccessMode.ReadOnly })) {
                    PowerPointFeatureReport report = presentation.InspectFeatures();
                    PowerPointFeatureFinding embedded = Assert.Single(report.FindFeatures("Embedded packages"));

                    Assert.Equal(1, embedded.Count);
                    Assert.Throws<InvalidOperationException>(() => report.EnsureNoAdvancedFeatures());
                }
            } finally {
                if (File.Exists(filePath)) {
                    File.Delete(filePath);
                }
            }
        }

        [Fact]
        public void PowerPointFeatureReport_TreatsMalformedChartWorkbookAsAdvanced() {
            string filePath = Path.Combine(Path.GetTempPath(), Guid.NewGuid() + ".pptx");

            try {
                using (PowerPointPresentation presentation = PowerPointPresentation.Create(filePath)) {
                    presentation.AddSlide().AddChart();
                    presentation.Save();
                }

                using (PresentationDocument document = PresentationDocument.Open(filePath, true)) {
                    ChartPart chartPart = document.PresentationPart!.SlideParts.Single().ChartParts.Single();
                    EmbeddedPackagePart package = Assert.Single(chartPart.GetPartsOfType<EmbeddedPackagePart>());
                    using var content = new MemoryStream(new byte[] { 1, 2, 3, 4 });
                    package.FeedData(content);
                }

                using (PowerPointPresentation presentation = PowerPointPresentation.Load(filePath, new PowerPointLoadOptions { AccessMode = OfficeIMO.Drawing.DocumentAccessMode.ReadOnly })) {
                    PowerPointFeatureReport report = presentation.InspectFeatures();
                    PowerPointFeatureFinding embedded = Assert.Single(report.FindFeatures("Embedded packages"));

                    Assert.Equal(1, embedded.Count);
                    Assert.Throws<InvalidOperationException>(() => report.EnsureNoAdvancedFeatures());
                }
            } finally {
                if (File.Exists(filePath)) {
                    File.Delete(filePath);
                }
            }
        }

        [Fact]
        public void PowerPointFeatureReport_TreatsChartWorkbookWithExternalRelationshipAsAdvanced() {
            string filePath = Path.Combine(Path.GetTempPath(), Guid.NewGuid() + ".pptx");

            try {
                using (PowerPointPresentation presentation = PowerPointPresentation.Create(filePath)) {
                    presentation.AddSlide().AddChart();
                    presentation.Save();
                }

                using (PresentationDocument document = PresentationDocument.Open(filePath, true)) {
                    ChartPart chartPart = document.PresentationPart!.SlideParts.Single().ChartParts.Single();
                    EmbeddedPackagePart package = Assert.Single(chartPart.GetPartsOfType<EmbeddedPackagePart>());
                    using Stream workbookStream = package.GetStream(FileMode.Open, FileAccess.ReadWrite);
                    using SpreadsheetDocument workbook = SpreadsheetDocument.Open(workbookStream, true);
                    workbook.WorkbookPart!.AddExternalRelationship(
                        "urn:officeimo:test-external-workbook",
                        new Uri("https://example.invalid/linked.xlsx"),
                        "rIdUnsafeExternalWorkbook");
                }

                using (PowerPointPresentation presentation = PowerPointPresentation.Load(filePath, new PowerPointLoadOptions { AccessMode = OfficeIMO.Drawing.DocumentAccessMode.ReadOnly })) {
                    PowerPointFeatureReport report = presentation.InspectFeatures();
                    PowerPointFeatureFinding embedded = Assert.Single(report.FindFeatures("Embedded packages"));

                    Assert.Equal(1, embedded.Count);
                    Assert.Throws<InvalidOperationException>(() => report.EnsureNoAdvancedFeatures());
                }
            } finally {
                if (File.Exists(filePath)) {
                    File.Delete(filePath);
                }
            }
        }

        [Fact]
        public void PowerPointFeatureReport_TreatsChartWorkbookWithNestedSharedStringPartAsAdvanced() {
            string filePath = Path.Combine(Path.GetTempPath(), Guid.NewGuid() + ".pptx");

            try {
                using (PowerPointPresentation presentation = PowerPointPresentation.Create(filePath)) {
                    presentation.AddSlide().AddChart();
                    presentation.Save();
                }

                using (PresentationDocument document = PresentationDocument.Open(filePath, true)) {
                    ChartPart chartPart = document.PresentationPart!.SlideParts.Single().ChartParts.Single();
                    EmbeddedPackagePart package = Assert.Single(chartPart.GetPartsOfType<EmbeddedPackagePart>());
                    using Stream workbookStream = package.GetStream(FileMode.Open, FileAccess.ReadWrite);
                    using SpreadsheetDocument workbook = SpreadsheetDocument.Open(workbookStream, true);
                    SharedStringTablePart sharedStrings = Assert.Single(workbook.WorkbookPart!.GetPartsOfType<SharedStringTablePart>());
                    AddExtendedPart(
                        sharedStrings,
                        "urn:officeimo:test-nested-shared-string-part",
                        "application/octet-stream",
                        "bin",
                        new byte[] { 1, 2, 3, 4 });
                }

                using (PowerPointPresentation presentation = PowerPointPresentation.Load(filePath, new PowerPointLoadOptions { AccessMode = OfficeIMO.Drawing.DocumentAccessMode.ReadOnly })) {
                    PowerPointFeatureReport report = presentation.InspectFeatures();
                    PowerPointFeatureFinding embedded = Assert.Single(report.FindFeatures("Embedded packages"));

                    Assert.Equal(1, embedded.Count);
                    Assert.Throws<InvalidOperationException>(() => report.EnsureNoAdvancedFeatures());
                }
            } finally {
                if (File.Exists(filePath)) {
                    File.Delete(filePath);
                }
            }
        }

        [Fact]
        public void PowerPointFeatureReport_TreatsOversizedCompressedChartWorksheetAsAdvanced() {
            string filePath = Path.Combine(Path.GetTempPath(), Guid.NewGuid() + ".pptx");

            try {
                using (PowerPointPresentation presentation = PowerPointPresentation.Create(filePath)) {
                    presentation.AddSlide().AddChart();
                    presentation.Save();
                }

                using (PresentationDocument document = PresentationDocument.Open(filePath, true)) {
                    ChartPart chartPart = document.PresentationPart!.SlideParts.Single().ChartParts.Single();
                    EmbeddedPackagePart package = Assert.Single(chartPart.GetPartsOfType<EmbeddedPackagePart>());
                    using Stream workbookStream = package.GetStream(FileMode.Open, FileAccess.ReadWrite);
                    using SpreadsheetDocument workbook = SpreadsheetDocument.Open(workbookStream, true);
                    WorksheetPart worksheetPart = Assert.Single(workbook.WorkbookPart!.GetPartsOfType<WorksheetPart>());
                    S.SheetData sheetData = worksheetPart.Worksheet.GetFirstChild<S.SheetData>()!;
                    sheetData.Append(new S.Row(
                        new S.Cell {
                            DataType = S.CellValues.String,
                            CellValue = new S.CellValue(new string('A', 3 * 1024 * 1024))
                        }));
                    worksheetPart.Worksheet.Save();
                }

                using (PowerPointPresentation presentation = PowerPointPresentation.Load(filePath, new PowerPointLoadOptions { AccessMode = OfficeIMO.Drawing.DocumentAccessMode.ReadOnly })) {
                    PowerPointFeatureReport report = presentation.InspectFeatures();
                    PowerPointFeatureFinding embedded = Assert.Single(report.FindFeatures("Embedded packages"));

                    Assert.Equal(1, embedded.Count);
                    Assert.Throws<InvalidOperationException>(() => report.EnsureNoAdvancedFeatures());
                }
            } finally {
                if (File.Exists(filePath)) {
                    File.Delete(filePath);
                }
            }
        }

        [Fact]
        public void PowerPointFeatureReport_DoesNotTreatOfficeImoMediaTimingAsAdvancedAnimation() {
            string filePath = Path.Combine(Path.GetTempPath(), Guid.NewGuid() + ".pptx");

            try {
                using (PowerPointPresentation presentation = PowerPointPresentation.Create(filePath)) {
                    using var media = new MemoryStream(new byte[] { 1, 2, 3, 4, 5 });
                    presentation.AddSlide().AddAudio(media, "audio/mpeg", ".mp3");
                    presentation.Save();
                }

                using (PowerPointPresentation presentation = PowerPointPresentation.Load(filePath, new PowerPointLoadOptions { AccessMode = OfficeIMO.Drawing.DocumentAccessMode.ReadOnly })) {
                    PowerPointFeatureReport report = presentation.InspectFeatures();

                    Assert.Contains(report.PartiallyEditableFeatures, feature => feature.Name == "Audio and video" && feature.Count == 1);
                    Assert.Empty(report.FindFeatures("Images"));
                    Assert.Same(report, report.EnsureNoFeatures("Images"));
                    Assert.DoesNotContain(report.PreservedFeatures, feature => feature.Name == "Animations and timing");
                    Assert.False(report.HasAdvancedFeatures);
                }
            } finally {
                if (File.Exists(filePath)) {
                    File.Delete(filePath);
                }
            }
        }

        [Fact]
        public void PowerPointFeatureReport_DetectsAnimationTimingOnMediaShape() {
            string filePath = Path.Combine(Path.GetTempPath(), Guid.NewGuid() + ".pptx");

            try {
                using (PowerPointPresentation presentation = PowerPointPresentation.Create(filePath)) {
                    using var media = new MemoryStream(new byte[] { 1, 2, 3, 4, 5 });
                    presentation.AddSlide().AddAudio(media, "audio/mpeg", ".mp3");
                    presentation.Save();
                }

                using (PresentationDocument document = PresentationDocument.Open(filePath, true)) {
                    SlidePart slidePart = document.PresentationPart!.SlideParts.Single();
                    string shapeId = slidePart.Slide.Descendants<Picture>().Single()
                        .NonVisualPictureProperties!.NonVisualDrawingProperties!.Id!.Value.ToString();
                    ChildTimeNodeList childNodes = slidePart.Slide.Timing!.Descendants<ChildTimeNodeList>().First();
                    childNodes.Append(new Animate(
                        new CommonBehavior(
                            new CommonTimeNode { Id = 900U, Duration = "500" },
                            new TargetElement(new ShapeTarget { ShapeId = shapeId }))));
                    slidePart.Slide.Save();
                }

                using (PowerPointPresentation presentation = PowerPointPresentation.Load(filePath, new PowerPointLoadOptions { AccessMode = OfficeIMO.Drawing.DocumentAccessMode.ReadOnly })) {
                    PowerPointFeatureReport report = presentation.InspectFeatures();
                    PowerPointFeatureFinding timing = Assert.Single(report.FindFeatures("Animations and timing"));

                    Assert.Equal(PowerPointFeatureSupportLevel.Preserved, timing.SupportLevel);
                    Assert.Equal(1, timing.Count);
                    Assert.Throws<InvalidOperationException>(() => report.EnsureNoAdvancedFeatures());
                }
            } finally {
                if (File.Exists(filePath)) {
                    File.Delete(filePath);
                }
            }
        }

        [Fact]
        public void PowerPointFeatureReport_DetectsTargetlessBuildTimingWithMediaPlayback() {
            string filePath = Path.Combine(Path.GetTempPath(), Guid.NewGuid() + ".pptx");

            try {
                using (PowerPointPresentation presentation = PowerPointPresentation.Create(filePath)) {
                    using var media = new MemoryStream(new byte[] { 1, 2, 3, 4, 5 });
                    presentation.AddSlide().AddAudio(media, "audio/mpeg", ".mp3");
                    presentation.Save();
                }

                using (PresentationDocument document = PresentationDocument.Open(filePath, true)) {
                    SlidePart slidePart = document.PresentationPart!.SlideParts.Single();
                    slidePart.Slide.Timing!.Append(new OpenXmlUnknownElement(
                        "p",
                        "bldLst",
                        "http://schemas.openxmlformats.org/presentationml/2006/main"));
                    slidePart.Slide.Save();
                }

                using (PowerPointPresentation presentation = PowerPointPresentation.Load(filePath, new PowerPointLoadOptions { AccessMode = OfficeIMO.Drawing.DocumentAccessMode.ReadOnly })) {
                    PowerPointFeatureReport report = presentation.InspectFeatures();
                    PowerPointFeatureFinding timing = Assert.Single(report.FindFeatures("Animations and timing"));

                    Assert.Equal(PowerPointFeatureSupportLevel.Preserved, timing.SupportLevel);
                    Assert.Equal(1, timing.Count);
                    Assert.Throws<InvalidOperationException>(() => report.EnsureNoAdvancedFeatures());
                }
            } finally {
                if (File.Exists(filePath)) {
                    File.Delete(filePath);
                }
            }
        }

        [Fact]
        public void PowerPointFeatureReport_DetectsLayoutTimingMarkup() {
            string filePath = Path.Combine(Path.GetTempPath(), Guid.NewGuid() + ".pptx");

            try {
                using (PowerPointPresentation presentation = PowerPointPresentation.Create(filePath)) {
                    presentation.AddSlide().AddTextBox("Inherits animated layout");
                    presentation.Save();
                }

                using (PresentationDocument document = PresentationDocument.Open(filePath, true)) {
                    SlideLayoutPart layoutPart = document.PresentationPart!.SlideParts.Single().SlideLayoutPart!;
                    layoutPart.SlideLayout!.Timing = new Timing(
                        new TimeNodeList(
                            new ParallelTimeNode(
                                new CommonTimeNode { Id = 900U, Duration = "500" })));
                    layoutPart.SlideLayout.Save();
                }

                using (PowerPointPresentation presentation = PowerPointPresentation.Load(filePath, new PowerPointLoadOptions { AccessMode = OfficeIMO.Drawing.DocumentAccessMode.ReadOnly })) {
                    PowerPointFeatureReport report = presentation.InspectFeatures();
                    PowerPointFeatureFinding timing = Assert.Single(report.FindFeatures("Animations and timing"));

                    Assert.Equal(PowerPointFeatureSupportLevel.Preserved, timing.SupportLevel);
                    Assert.Equal(1, timing.Count);
                    Assert.Contains(timing.Details, detail => detail.Contains("layout", StringComparison.OrdinalIgnoreCase));
                    Assert.Throws<InvalidOperationException>(() => report.EnsureNoAdvancedFeatures());
                }
            } finally {
                if (File.Exists(filePath)) {
                    File.Delete(filePath);
                }
            }
        }

        [Fact]
        public void PowerPointFeatureReport_DoesNotTreatVbaNamedMediaAsMacros() {
            string filePath = Path.Combine(Path.GetTempPath(), Guid.NewGuid() + ".pptx");

            try {
                using (PowerPointPresentation presentation = PowerPointPresentation.Create(filePath)) {
                    presentation.AddSlide().AddTextBox("Ordinary VBA named asset");
                    presentation.Save();
                }

                using (PresentationDocument document = PresentationDocument.Open(filePath, true)) {
                    SlidePart slidePart = document.PresentationPart!.SlideParts.Single();
                    AddExtendedPart(
                        slidePart,
                        "http://schemas.openxmlformats.org/officeDocument/2006/relationships/image",
                        "image/png",
                        "vbaProject.png",
                        new byte[] { 137, 80, 78, 71 });
                }

                using (PowerPointPresentation presentation = PowerPointPresentation.Load(filePath, new PowerPointLoadOptions { AccessMode = OfficeIMO.Drawing.DocumentAccessMode.ReadOnly })) {
                    PowerPointFeatureReport report = presentation.InspectFeatures();

                    Assert.Empty(report.FindFeatures("VBA macros"));
                    Assert.Same(report, report.EnsureNoAdvancedFeatures());
                }
            } finally {
                if (File.Exists(filePath)) {
                    File.Delete(filePath);
                }
            }
        }

        [Fact]
        public void PowerPointFeatureReport_DoesNotTreatControlNamedMediaAsAdvancedPackageSignals() {
            string filePath = Path.Combine(Path.GetTempPath(), Guid.NewGuid() + ".pptx");

            try {
                using (PowerPointPresentation presentation = PowerPointPresentation.Create(filePath)) {
                    presentation.AddSlide().AddTextBox("Ordinary control named assets");
                    presentation.Save();
                }

                using (PresentationDocument document = PresentationDocument.Open(filePath, true)) {
                    SlidePart slidePart = document.PresentationPart!.SlideParts.Single();
                    AddExtendedPart(
                        slidePart,
                        "http://schemas.openxmlformats.org/officeDocument/2006/relationships/image",
                        "image/png",
                        "activeX.png",
                        new byte[] { 137, 80, 78, 71 });
                    AddExtendedPart(
                        slidePart,
                        "http://schemas.openxmlformats.org/officeDocument/2006/relationships/image",
                        "image/png",
                        "taskpane.png",
                        new byte[] { 137, 80, 78, 71 });
                }

                using (PowerPointPresentation presentation = PowerPointPresentation.Load(filePath, new PowerPointLoadOptions { AccessMode = OfficeIMO.Drawing.DocumentAccessMode.ReadOnly })) {
                    PowerPointFeatureReport report = presentation.InspectFeatures();

                    Assert.Empty(report.FindFeatures("ActiveX controls"));
                    Assert.Empty(report.FindFeatures("Web extensions and task panes"));
                    Assert.Same(report, report.EnsureNoAdvancedFeatures());
                }
            } finally {
                if (File.Exists(filePath)) {
                    File.Delete(filePath);
                }
            }
        }

        [Fact]
        public void PowerPointFeatureReport_DetectsAdvancedPackageSignals() {
            string filePath = Path.Combine(Path.GetTempPath(), Guid.NewGuid() + ".pptx");

            try {
                using (PowerPointPresentation presentation = PowerPointPresentation.Create(filePath)) {
                    presentation.AddSlide().AddTextBox("Advanced package signals");
                    presentation.Save();
                }

                using (PresentationDocument document = PresentationDocument.Open(filePath, true)) {
                    PresentationPart presentationPart = document.PresentationPart!;
                    CustomXmlPart customXmlPart = presentationPart.AddCustomXmlPart(CustomXmlPartType.CustomXml);
                    using (var stream = new MemoryStream(Encoding.UTF8.GetBytes("<root><value>42</value></root>"))) {
                        customXmlPart.FeedData(stream);
                    }

                    AddExtendedPart(presentationPart,
                        "http://schemas.microsoft.com/office/2006/relationships/vbaProject",
                        "application/vnd.ms-office.vbaProject",
                        new byte[] { 1, 2, 3, 4 });
                    AddExtendedPart(presentationPart,
                        "http://schemas.microsoft.com/office/2011/relationships/webextension",
                        "application/vnd.ms-office.webextension+xml",
                        "<we:webextension xmlns:we=\"http://schemas.microsoft.com/office/webextensions/webextension/2010/11\" />");
                    AddExtendedPart(presentationPart,
                        "http://schemas.openxmlformats.org/package/2006/relationships/digital-signature/signature",
                        "application/vnd.openxmlformats-package.digital-signature-xmlsignature+xml",
                        "<Signature xmlns=\"http://www.w3.org/2000/09/xmldsig#\" />");
                    document.AddDigitalSignatureOriginPart();
                    XmlSignaturePart signaturePart = document.DigitalSignatureOriginPart!.AddNewPart<XmlSignaturePart>();
                    using (var stream = new MemoryStream(Encoding.UTF8.GetBytes("<Signature xmlns=\"http://www.w3.org/2000/09/xmldsig#\" />"))) {
                        signaturePart.FeedData(stream);
                    }
                }

                using (PowerPointPresentation presentation = PowerPointPresentation.Load(filePath, new PowerPointLoadOptions { AccessMode = OfficeIMO.Drawing.DocumentAccessMode.ReadOnly })) {
                    PowerPointFeatureReport report = presentation.InspectFeatures();

                    PowerPointFeatureFinding customXml = Assert.Single(report.FindFeatures("Custom XML parts"));
                    PowerPointFeatureFinding macros = Assert.Single(report.FindFeatures("VBA macros"));
                    PowerPointFeatureFinding webExtensions = Assert.Single(report.FindFeatures("Web extensions and task panes"));
                    PowerPointFeatureFinding signatures = Assert.Single(report.FindFeatures("Digital signatures"));

                    Assert.Equal(PowerPointFeatureSupportLevel.Preserved, customXml.SupportLevel);
                    Assert.Equal(PowerPointFeatureSupportLevel.Preserved, macros.SupportLevel);
                    Assert.Equal(PowerPointFeatureSupportLevel.Preserved, webExtensions.SupportLevel);
                    Assert.Equal(PowerPointFeatureSupportLevel.Unsupported, signatures.SupportLevel);
                    Assert.Contains(macros.Details, detail => detail.Contains("vbaProject", StringComparison.OrdinalIgnoreCase));
                    Assert.Contains(webExtensions.Details, detail => detail.Contains("webextension", StringComparison.OrdinalIgnoreCase));
                    Assert.Contains(signatures.Details, detail => detail.Contains("signature", StringComparison.OrdinalIgnoreCase));

                    InvalidOperationException advancedException = Assert.Throws<InvalidOperationException>(() => report.EnsureNoAdvancedFeatures());
                    Assert.Contains("Custom XML parts", advancedException.Message);
                    Assert.Contains("Digital signatures", advancedException.Message);

                    InvalidOperationException unsupportedException = Assert.Throws<InvalidOperationException>(() => report.EnsureNoUnsupportedFeatures());
                    Assert.Contains("Digital signatures", unsupportedException.Message);
                }
            } finally {
                if (File.Exists(filePath)) {
                    File.Delete(filePath);
                }
            }
        }

        [Fact]
        public void PowerPointFeatureReport_DetectsApplicationPropertyDigitalSignature() {
            string filePath = Path.Combine(Path.GetTempPath(), Guid.NewGuid() + ".pptx");

            try {
                using (PowerPointPresentation presentation = PowerPointPresentation.Create(filePath)) {
                    presentation.AddSlide().AddTextBox("Application signature metadata");
                    presentation.Save();
                }

                using (PresentationDocument document = PresentationDocument.Open(filePath, true)) {
                    ExtendedFilePropertiesPart appPart = document.ExtendedFilePropertiesPart
                        ?? document.AddExtendedFilePropertiesPart();
                    appPart.Properties ??= new DocumentFormat.OpenXml.ExtendedProperties.Properties();
                    appPart.Properties.DigitalSignature = new DocumentFormat.OpenXml.ExtendedProperties.DigitalSignature();
                    appPart.Properties.Save();
                }

                using (PowerPointPresentation presentation = PowerPointPresentation.Load(filePath, new PowerPointLoadOptions { AccessMode = OfficeIMO.Drawing.DocumentAccessMode.ReadOnly })) {
                    PowerPointFeatureReport report = presentation.InspectFeatures();
                    PowerPointFeatureFinding signatures = Assert.Single(report.FindFeatures("Digital signatures"));

                    Assert.Equal(PowerPointFeatureSupportLevel.Unsupported, signatures.SupportLevel);
                    Assert.Contains(signatures.Details, detail => detail.Contains("application properties", StringComparison.OrdinalIgnoreCase));
                    Assert.Throws<InvalidOperationException>(() => report.EnsureNoUnsupportedFeatures());
                }
            } finally {
                if (File.Exists(filePath)) {
                    File.Delete(filePath);
                }
            }
        }

        [Fact]
        public void PowerPointFeatureReport_IgnoresPlainTablesWithoutStyleMetadata() {
            string filePath = Path.Combine(Path.GetTempPath(), Guid.NewGuid() + ".pptx");

            try {
                using (PowerPointPresentation presentation = PowerPointPresentation.Create(filePath)) {
                    PowerPointTable table = presentation.AddSlide().AddTable(1, 1);
                    table.GetCell(0, 0).Text = "Plain";
                    presentation.Save();
                }

                using (PresentationDocument document = PresentationDocument.Open(filePath, true)) {
                    A.Table table = document.PresentationPart!.SlideParts.Single().Slide.Descendants<A.Table>().Single();
                    table.TableProperties?.Remove();
                    foreach (OpenXmlElement idElement in table.Descendants()
                        .Where(element => string.Equals(element.LocalName, "colId", StringComparison.OrdinalIgnoreCase)
                            || string.Equals(element.LocalName, "rowId", StringComparison.OrdinalIgnoreCase))
                        .ToArray()) {
                        idElement.Remove();
                    }

                    document.PresentationPart!.SlideParts.Single().Slide.Save();
                }

                using (PowerPointPresentation presentation = PowerPointPresentation.Load(filePath, new PowerPointLoadOptions { AccessMode = OfficeIMO.Drawing.DocumentAccessMode.ReadOnly })) {
                    PowerPointFeatureReport report = presentation.InspectFeatures();
                    PowerPointFeatureFinding tableMetadata = Assert.Single(report.FindFeatures("Table style metadata"));

                    Assert.Equal(0, tableMetadata.Count);
                    Assert.Same(report, report.EnsureNoFeatures("Table style metadata"));
                }
            } finally {
                if (File.Exists(filePath)) {
                    File.Delete(filePath);
                }
            }
        }

        [Fact]
        public void PowerPointFeatureReport_DetectsActiveXControlPackageSignals() {
            string filePath = Path.Combine(Path.GetTempPath(), Guid.NewGuid() + ".pptx");

            try {
                using (PowerPointPresentation presentation = PowerPointPresentation.Create(filePath)) {
                    presentation.AddSlide().AddTextBox("ActiveX package signals");
                    presentation.Save();
                }

                using (PresentationDocument document = PresentationDocument.Open(filePath, true)) {
                    PresentationPart presentationPart = document.PresentationPart!;
                    AddExtendedPart(presentationPart,
                        "http://schemas.openxmlformats.org/officeDocument/2006/relationships/control",
                        "application/vnd.ms-office.activeX+xml",
                        "<ax:ocx xmlns:ax=\"http://schemas.microsoft.com/office/2006/activeX\" />");
                    AddExtendedPart(presentationPart,
                        "http://schemas.openxmlformats.org/officeDocument/2006/relationships/activeXControlBinary",
                        "application/vnd.ms-office.activeX",
                        new byte[] { 1, 2, 3, 4 });
                }

                using (PowerPointPresentation presentation = PowerPointPresentation.Load(filePath, new PowerPointLoadOptions { AccessMode = OfficeIMO.Drawing.DocumentAccessMode.ReadOnly })) {
                    PowerPointFeatureReport report = presentation.InspectFeatures();
                    PowerPointFeatureFinding activeX = Assert.Single(report.FindFeatures("ActiveX controls"));

                    Assert.Equal(PowerPointFeatureSupportLevel.Preserved, activeX.SupportLevel);
                    Assert.Equal(2, activeX.Count);
                    Assert.Contains(activeX.Details, detail => detail.Contains("activeX", StringComparison.OrdinalIgnoreCase));
                    Assert.Throws<InvalidOperationException>(() => report.EnsureNoAdvancedFeatures());
                }
            } finally {
                if (File.Exists(filePath)) {
                    File.Delete(filePath);
                }
            }
        }

        [Fact]
        public void PowerPointFeatureReport_DetectsEmbeddedOleObjectParts() {
            string filePath = Path.Combine(Path.GetTempPath(), Guid.NewGuid() + ".pptx");

            try {
                using (PowerPointPresentation presentation = PowerPointPresentation.Create(filePath)) {
                    presentation.AddSlide().AddTextBox("Embedded object");
                    presentation.Save();
                }

                using (PresentationDocument document = PresentationDocument.Open(filePath, true)) {
                    SlidePart slidePart = document.PresentationPart!.SlideParts.Single();
                    EmbeddedObjectPart embeddedObjectPart =
                        slidePart.AddEmbeddedObjectPart("application/vnd.openxmlformats-officedocument.oleObject");
                    using var stream = new MemoryStream(new byte[] { 1, 2, 3, 4 });
                    embeddedObjectPart.FeedData(stream);
                    slidePart.Slide.Save();
                }

                using (PowerPointPresentation presentation = PowerPointPresentation.Load(filePath, new PowerPointLoadOptions { AccessMode = OfficeIMO.Drawing.DocumentAccessMode.ReadOnly })) {
                    PowerPointFeatureReport report = presentation.InspectFeatures();
                    PowerPointFeatureFinding embedded = Assert.Single(report.FindFeatures("Embedded packages"));

                    Assert.Equal(PowerPointFeatureSupportLevel.Preserved, embedded.SupportLevel);
                    Assert.Equal(1, embedded.Count);
                    Assert.Contains(embedded.Details, detail => detail.Contains("oleObject", StringComparison.OrdinalIgnoreCase));
                    Assert.Throws<InvalidOperationException>(() => report.EnsureNoAdvancedFeatures());
                }
            } finally {
                if (File.Exists(filePath)) {
                    File.Delete(filePath);
                }
            }
        }

        [Fact]
        public void PowerPointFeatureReport_DetectsExternalRelationshipsWithoutHyperlinks() {
            string filePath = Path.Combine(Path.GetTempPath(), Guid.NewGuid() + ".pptx");

            try {
                using (PowerPointPresentation presentation = PowerPointPresentation.Create(filePath)) {
                    presentation.AddSlide().AddTextBox("Linked asset");
                    presentation.Save();
                }

                using (PresentationDocument document = PresentationDocument.Open(filePath, true)) {
                    document.PresentationPart!.AddExternalRelationship(
                        "http://schemas.openxmlformats.org/officeDocument/2006/relationships/image",
                        new Uri("https://example.com/logo.png"));
                }

                using (PowerPointPresentation presentation = PowerPointPresentation.Load(filePath, new PowerPointLoadOptions { AccessMode = OfficeIMO.Drawing.DocumentAccessMode.ReadOnly })) {
                    PowerPointFeatureReport report = presentation.InspectFeatures();
                    PowerPointFeatureFinding external = Assert.Single(report.FindFeatures("External package relationships"));

                    Assert.Equal(PowerPointFeatureSupportLevel.Preserved, external.SupportLevel);
                    Assert.Equal(1, external.Count);
                    Assert.Contains(external.Details, detail => detail.Contains("relationships/image", StringComparison.OrdinalIgnoreCase)
                        && detail.Contains("https://example.com/logo.png", StringComparison.OrdinalIgnoreCase));
                    Assert.Throws<InvalidOperationException>(() => report.EnsureNoAdvancedFeatures());
                }
            } finally {
                if (File.Exists(filePath)) {
                    File.Delete(filePath);
                }
            }
        }

        [Fact]
        public void PowerPointFeatureReport_DoesNotTreatPersonNamedMediaAsComments() {
            string filePath = Path.Combine(Path.GetTempPath(), Guid.NewGuid() + ".pptx");

            try {
                using (PowerPointPresentation presentation = PowerPointPresentation.Create(filePath)) {
                    presentation.AddSlide().AddTextBox("Ordinary person asset");
                    presentation.Save();
                }

                using (PresentationDocument document = PresentationDocument.Open(filePath, true)) {
                    SlidePart slidePart = document.PresentationPart!.SlideParts.Single();
                    AddExtendedPart(
                        slidePart,
                        "http://schemas.openxmlformats.org/officeDocument/2006/relationships/image",
                        "image/png",
                        "person.png",
                        new byte[] { 137, 80, 78, 71 });
                }

                using (PowerPointPresentation presentation = PowerPointPresentation.Load(filePath, new PowerPointLoadOptions { AccessMode = OfficeIMO.Drawing.DocumentAccessMode.ReadOnly })) {
                    PowerPointFeatureReport report = presentation.InspectFeatures();

                    Assert.Empty(report.FindFeatures("Comments"));
                    Assert.Same(report, report.EnsureNoAdvancedFeatures());
                }
            } finally {
                if (File.Exists(filePath)) {
                    File.Delete(filePath);
                }
            }
        }

        [Fact]
        public void PowerPointFeatureReport_DetectsUnsupportedTransitionMarkup() {
            string filePath = Path.Combine(Path.GetTempPath(), Guid.NewGuid() + ".pptx");

            try {
                using (PowerPointPresentation presentation = PowerPointPresentation.Create(filePath)) {
                    presentation.AddSlide().AddTextBox("Unsupported transition");
                    presentation.Save();
                }

                using (PresentationDocument document = PresentationDocument.Open(filePath, true)) {
                    SlidePart slidePart = document.PresentationPart!.SlideParts.Single();
                    slidePart.Slide.Transition = new Transition(
                        new OpenXmlUnknownElement("p14", "doors", "http://schemas.microsoft.com/office/powerpoint/2010/main"));
                    slidePart.Slide.Save();
                }

                using (PowerPointPresentation presentation = PowerPointPresentation.Load(filePath, new PowerPointLoadOptions { AccessMode = OfficeIMO.Drawing.DocumentAccessMode.ReadOnly })) {
                    PowerPointFeatureReport report = presentation.InspectFeatures();
                    PowerPointFeatureFinding unsupported = Assert.Single(report.FindFeatures("Unsupported transition markup"));

                    Assert.Equal(PowerPointFeatureSupportLevel.Preserved, unsupported.SupportLevel);
                    Assert.Equal(1, unsupported.Count);
                    Assert.Contains(unsupported.Details, detail => detail.Contains("doors", StringComparison.OrdinalIgnoreCase));
                    Assert.Throws<InvalidOperationException>(() => report.EnsureNoAdvancedFeatures());
                }
            } finally {
                if (File.Exists(filePath)) {
                    File.Delete(filePath);
                }
            }
        }

        [Fact]
        public void PowerPointFeatureReport_DetectsUnsupportedMetadataOnMappedTransition() {
            string filePath = Path.Combine(Path.GetTempPath(), Guid.NewGuid() + ".pptx");

            try {
                using (PowerPointPresentation presentation = PowerPointPresentation.Create(filePath)) {
                    PowerPointSlide slide = presentation.AddSlide();
                    slide.AddTextBox("Mapped transition with sound metadata");
                    slide.Transition = SlideTransition.Fade;
                    presentation.Save();
                }

                using (PresentationDocument document = PresentationDocument.Open(filePath, true)) {
                    SlidePart slidePart = document.PresentationPart!.SlideParts.Single();
                    slidePart.Slide.Transition!.Append(new OpenXmlUnknownElement(
                        "p",
                        "sndAc",
                        "http://schemas.openxmlformats.org/presentationml/2006/main"));
                    slidePart.Slide.Save();
                }

                using (PowerPointPresentation presentation = PowerPointPresentation.Load(filePath, new PowerPointLoadOptions { AccessMode = OfficeIMO.Drawing.DocumentAccessMode.ReadOnly })) {
                    PowerPointFeatureReport report = presentation.InspectFeatures();
                    PowerPointFeatureFinding unsupported = Assert.Single(report.FindFeatures("Unsupported transition markup"));

                    Assert.Equal(PowerPointFeatureSupportLevel.Preserved, unsupported.SupportLevel);
                    Assert.Equal(1, unsupported.Count);
                    Assert.Contains(unsupported.Details, detail => detail.Contains("sndAc", StringComparison.OrdinalIgnoreCase));
                    Assert.Throws<InvalidOperationException>(() => report.EnsureNoAdvancedFeatures());
                }
            } finally {
                if (File.Exists(filePath)) {
                    File.Delete(filePath);
                }
            }
        }

        [Fact]
        public void PowerPointFeatureReport_DoesNotTreatAuthoredMorphTransitionAsUnsupported() {
            string filePath = Path.Combine(Path.GetTempPath(), Guid.NewGuid() + ".pptx");

            try {
                using (PowerPointPresentation presentation = PowerPointPresentation.Create(filePath)) {
                    PowerPointSlide slide = presentation.AddSlide();
                    slide.AddTextBox("Authored Morph transition");
                    slide.Transition = SlideTransition.Morph;
                    presentation.Save();
                }

                using (PowerPointPresentation presentation = PowerPointPresentation.Load(filePath, new PowerPointLoadOptions { AccessMode = OfficeIMO.Drawing.DocumentAccessMode.ReadOnly })) {
                    PowerPointFeatureReport report = presentation.InspectFeatures();

                    Assert.Contains(report.EditableFeatures, feature => feature.Name == "Slide transitions" && feature.Count == 1);
                    Assert.Empty(report.FindFeatures("Unsupported transition markup"));
                    Assert.Same(report, report.EnsureNoAdvancedFeatures());
                }
            } finally {
                if (File.Exists(filePath)) {
                    File.Delete(filePath);
                }
            }
        }

        [Fact]
        public void PowerPointFeatureReport_DetectsUnsupportedEffectAttributesOnMappedTransition() {
            string filePath = Path.Combine(Path.GetTempPath(), Guid.NewGuid() + ".pptx");

            try {
                using (PowerPointPresentation presentation = PowerPointPresentation.Create(filePath)) {
                    presentation.AddSlide().AddTextBox("Mapped transition with unsupported effect option");
                    presentation.Save();
                }

                using (PresentationDocument document = PresentationDocument.Open(filePath, true)) {
                    WipeTransition wipe = new();
                    wipe.SetAttribute(new OpenXmlAttribute("unsupported", string.Empty, "value"));
                    SlidePart slidePart = document.PresentationPart!.SlideParts.Single();
                    slidePart.Slide.Transition = new Transition(wipe);
                    slidePart.Slide.Save();
                }

                using (PowerPointPresentation presentation = PowerPointPresentation.Load(filePath, new PowerPointLoadOptions { AccessMode = OfficeIMO.Drawing.DocumentAccessMode.ReadOnly })) {
                    PowerPointFeatureReport report = presentation.InspectFeatures();
                    PowerPointFeatureFinding unsupported = Assert.Single(report.FindFeatures("Unsupported transition markup"));

                    Assert.Equal(PowerPointFeatureSupportLevel.Preserved, unsupported.SupportLevel);
                    Assert.Equal(1, unsupported.Count);
                    Assert.Contains(unsupported.Details, detail => detail.Contains("unsupported", StringComparison.OrdinalIgnoreCase));
                    Assert.Throws<InvalidOperationException>(() => report.EnsureNoAdvancedFeatures());
                }
            } finally {
                if (File.Exists(filePath)) {
                    File.Delete(filePath);
                }
            }
        }

        [Fact]
        public void PowerPointFeatureReport_DetectsUnsupportedPresetTransitions() {
            string filePath = Path.Combine(Path.GetTempPath(), Guid.NewGuid() + ".pptx");

            try {
                using (PowerPointPresentation presentation = PowerPointPresentation.Create(filePath)) {
                    presentation.AddSlide().AddTextBox("Unsupported preset transition");
                    presentation.Save();
                }

                using (PresentationDocument document = PresentationDocument.Open(filePath, true)) {
                    OpenXmlUnknownElement presetTransition = new(
                        "p15",
                        "prstTrans",
                        "http://schemas.microsoft.com/office/powerpoint/2012/main");
                    presetTransition.SetAttribute(new OpenXmlAttribute("prst", string.Empty, "gallery"));

                    SlidePart slidePart = document.PresentationPart!.SlideParts.Single();
                    slidePart.Slide.Transition = new Transition(presetTransition);
                    slidePart.Slide.Save();
                }

                using (PowerPointPresentation presentation = PowerPointPresentation.Load(filePath, new PowerPointLoadOptions { AccessMode = OfficeIMO.Drawing.DocumentAccessMode.ReadOnly })) {
                    PowerPointFeatureReport report = presentation.InspectFeatures();
                    PowerPointFeatureFinding unsupported = Assert.Single(report.FindFeatures("Unsupported transition markup"));

                    Assert.Equal(PowerPointFeatureSupportLevel.Preserved, unsupported.SupportLevel);
                    Assert.Equal(1, unsupported.Count);
                    Assert.Contains(unsupported.Details, detail => detail.Contains("gallery", StringComparison.OrdinalIgnoreCase));
                    Assert.Throws<InvalidOperationException>(() => report.EnsureNoAdvancedFeatures());
                }
            } finally {
                if (File.Exists(filePath)) {
                    File.Delete(filePath);
                }
            }
        }

        [Fact]
        public void PowerPointFeatureReport_DetectsUnsupportedAlternateContentTransitionMarkup() {
            string filePath = Path.Combine(Path.GetTempPath(), Guid.NewGuid() + ".pptx");

            try {
                using (PowerPointPresentation presentation = PowerPointPresentation.Create(filePath)) {
                    presentation.AddSlide().AddTextBox("Unsupported alternate transition");
                    presentation.Save();
                }

                using (PresentationDocument document = PresentationDocument.Open(filePath, true)) {
                    SlidePart slidePart = document.PresentationPart!.SlideParts.Single();
                    Transition unsupportedTransition = new(
                        new OpenXmlUnknownElement("p14", "doors", "http://schemas.microsoft.com/office/powerpoint/2010/main"));
                    AlternateContentChoice choice = new() { Requires = "p14" };
                    choice.Append(unsupportedTransition);
                    AlternateContent alternateContent = new();
                    alternateContent.Append(choice);
                    slidePart.Slide.InsertAt(alternateContent, 0);
                    slidePart.Slide.Save();
                }

                using (PowerPointPresentation presentation = PowerPointPresentation.Load(filePath, new PowerPointLoadOptions { AccessMode = OfficeIMO.Drawing.DocumentAccessMode.ReadOnly })) {
                    PowerPointFeatureReport report = presentation.InspectFeatures();
                    PowerPointFeatureFinding unsupported = Assert.Single(report.FindFeatures("Unsupported transition markup"));

                    Assert.Equal(PowerPointFeatureSupportLevel.Preserved, unsupported.SupportLevel);
                    Assert.Equal(1, unsupported.Count);
                    Assert.Contains(unsupported.Details, detail => detail.Contains("doors", StringComparison.OrdinalIgnoreCase));
                    Assert.Throws<InvalidOperationException>(() => report.EnsureNoAdvancedFeatures());
                }
            } finally {
                if (File.Exists(filePath)) {
                    File.Delete(filePath);
                }
            }
        }

        [Fact]
        public void PowerPointFeatureReport_DetectsUnsupportedAlternateContentTransitionFallback() {
            string filePath = Path.Combine(Path.GetTempPath(), Guid.NewGuid() + ".pptx");

            try {
                using (PowerPointPresentation presentation = PowerPointPresentation.Create(filePath)) {
                    presentation.AddSlide().AddTextBox("Unsupported fallback transition");
                    presentation.Save();
                }

                using (PresentationDocument document = PresentationDocument.Open(filePath, true)) {
                    SlidePart slidePart = document.PresentationPart!.SlideParts.Single();
                    AlternateContentChoice choice = new() { Requires = "p14" };
                    choice.Append(new Transition(new FadeTransition()));

                    AlternateContentFallback fallback = new();
                    fallback.Append(new Transition(
                        new OpenXmlUnknownElement("p14", "doors", "http://schemas.microsoft.com/office/powerpoint/2010/main")));

                    AlternateContent alternateContent = new();
                    alternateContent.Append(choice);
                    alternateContent.Append(fallback);
                    slidePart.Slide.InsertAt(alternateContent, 0);
                    slidePart.Slide.Save();
                }

                using (PowerPointPresentation presentation = PowerPointPresentation.Load(filePath, new PowerPointLoadOptions { AccessMode = OfficeIMO.Drawing.DocumentAccessMode.ReadOnly })) {
                    PowerPointFeatureReport report = presentation.InspectFeatures();
                    PowerPointFeatureFinding unsupported = Assert.Single(report.FindFeatures("Unsupported transition markup"));

                    Assert.Equal(PowerPointFeatureSupportLevel.Preserved, unsupported.SupportLevel);
                    Assert.Equal(1, unsupported.Count);
                    Assert.Contains(unsupported.Details, detail => detail.Contains("doors", StringComparison.OrdinalIgnoreCase));
                    Assert.Throws<InvalidOperationException>(() => report.EnsureNoAdvancedFeatures());
                }
            } finally {
                if (File.Exists(filePath)) {
                    File.Delete(filePath);
                }
            }
        }

        private static void AddExtendedPart(OpenXmlPartContainer container, string relationshipType, string contentType, string xml) {
            AddExtendedPart(container, relationshipType, contentType, Encoding.UTF8.GetBytes(xml));
        }

        private static void AddExtendedPart(OpenXmlPartContainer container, string relationshipType, string contentType, byte[] bytes) {
            AddExtendedPart(container, relationshipType, contentType, "xml", bytes);
        }

        private static void AddExtendedPart(OpenXmlPartContainer container, string relationshipType, string contentType, string targetExt, byte[] bytes) {
            ExtendedPart part = container.AddExtendedPart(relationshipType, contentType, targetExt);
            using var stream = new MemoryStream(bytes);
            part.FeedData(stream);
        }
    }
}
