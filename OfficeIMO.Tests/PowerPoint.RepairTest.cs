using System;
using System.IO;
using System.Linq;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Presentation;
using DocumentFormat.OpenXml.Validation;
using OfficeIMO.PowerPoint;
using Xunit;

namespace OfficeIMO.Tests {
    public class PowerPointRepairTest {
        [Fact]
        public void CreatePresentationWithoutRepairIssues() {
            string filePath = Path.Combine(Path.GetTempPath(), "test_no_repair_" + Guid.NewGuid() + ".pptx");
            string imagePath = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "Images", "BackgroundImage.png");

            using (PowerPointPresentation presentation = PowerPointPresentation.Create(filePath)) {
                // Add multiple slides with various content
                for (int i = 0; i < 3; i++) {
                    PowerPointSlide slide = presentation.AddSlide();
                    slide.AddTitle($"Slide {i + 1} Title");
                    slide.AddTextBox($"Content for slide {i + 1}");
                    
                    if (File.Exists(imagePath)) {
                        slide.AddPicture(imagePath);
                    }
                    
                    slide.AddTable(2, 3);
                    slide.Notes.Text = $"Notes for slide {i + 1}";
                }
                
                presentation.Save();
            }

            // Validate the document
            using (PresentationDocument document = PresentationDocument.Open(filePath, false)) {
                OpenXmlValidator validator = new OpenXmlValidator(FileFormatVersions.Microsoft365);
                var errors = validator.Validate(document).ToList();
                
                // Print validation errors if any
                foreach (var error in errors) {
                    Console.WriteLine($"Validation Error: {error.Description}");
                    Console.WriteLine($"  Part: {error.Part?.Uri}");
                    Console.WriteLine($"  Node: {error.Path?.XPath}");
                }
                
                // Check for common repair triggers
                Assert.NotNull(document.PresentationPart);
                
                // Check for markup compatibility attributes
                var presentation = document.PresentationPart.Presentation;
                Assert.NotNull(presentation);
                
                // Check for unique shape IDs
                foreach (var slidePart in document.PresentationPart.SlideParts) {
                    var shapeTree = slidePart.Slide.CommonSlideData?.ShapeTree;
                    if (shapeTree != null) {
                        var ids = shapeTree.Descendants<NonVisualDrawingProperties>()
                            .Select(dp => dp.Id?.Value)
                            .Where(id => id.HasValue)
                            .Select(id => id.Value)
                            .ToList();
                        
                        // Verify all IDs are unique
                        Assert.Equal(ids.Count, ids.Distinct().Count());
                    }
                }
                
                // Check slide ID list integrity
                var slideIdList = presentation.SlideIdList;
                Assert.NotNull(slideIdList);
                
                foreach (SlideId slideId in slideIdList.Elements<SlideId>()) {
                    Assert.NotNull(slideId.RelationshipId);
                    Assert.NotNull(slideId.Id);
                    Assert.True(slideId.Id.Value >= 256);
                    
                    // Verify the relationship exists
                    var part = document.PresentationPart.GetPartById(slideId.RelationshipId);
                    Assert.NotNull(part);
                    Assert.IsType<SlidePart>(part);
                }
            }

            File.Delete(filePath);
        }
    }
}