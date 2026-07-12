using System;
using System.IO;
using System.Linq;
using DocumentFormat.OpenXml.Drawing.Diagrams;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Presentation;
using OfficeIMO.PowerPoint;
using Xunit;

namespace OfficeIMO.Tests {
    public class PowerPointSmartArtTests {
        [Fact]
        public void CanAddSmartArtAndEditNodeText() {
            string filePath = System.IO.Path.Combine(System.IO.Path.GetTempPath(), Guid.NewGuid() + ".pptx");

            try {
                using (PowerPointPresentation presentation = PowerPointPresentation.Create(filePath)) {
                    PowerPointSlide slide = presentation.AddSlide();
                    PowerPointSmartArt smartArt = slide.AddSmartArt();

                    Assert.Equal(PowerPointShapeContentType.SmartArt, smartArt.ShapeContentType);
                    Assert.Equal(1, smartArt.NodeCount);
                    smartArt.SetNodeText(0, "OfficeIMO-native process");
                    Assert.Equal("OfficeIMO-native process", smartArt.GetNodeText(0));
                    presentation.Save();
                }

                using (PresentationDocument document = PresentationDocument.Open(filePath, false)) {
                    SlidePart slidePart = document.PresentationPart!.SlideParts.Single();
                    Assert.Single(slidePart.DiagramDataParts);
                    Assert.Single(slidePart.DiagramLayoutDefinitionParts);
                    Assert.Single(slidePart.DiagramStyleParts);
                    Assert.Single(slidePart.DiagramColorsParts);

                    GraphicFrame frame = slidePart.Slide.Descendants<GraphicFrame>().Single();
                    RelationshipIds relationships = frame.Graphic!.GraphicData!.GetFirstChild<RelationshipIds>()!;
                    Assert.False(string.IsNullOrWhiteSpace(relationships.LayoutPart));
                    Assert.False(string.IsNullOrWhiteSpace(relationships.StylePart));
                    Assert.False(string.IsNullOrWhiteSpace(relationships.ColorPart));
                    Assert.False(string.IsNullOrWhiteSpace(relationships.DataPart));

                    DiagramLayoutDefinitionPart layoutPart =
                        (DiagramLayoutDefinitionPart)slidePart.GetPartById(relationships.LayoutPart!);
                    Assert.Equal("urn:microsoft.com/office/officeart/2005/8/layout/default",
                        layoutPart.LayoutDefinition!.UniqueId!.Value);
                }

                using (PowerPointPresentation reloaded = PowerPointPresentation.Load(filePath)) {
                    PowerPointSmartArt smartArt = Assert.IsType<PowerPointSmartArt>(reloaded.Slides[0].Shapes.Single());
                    Assert.Equal("OfficeIMO-native process", smartArt.GetNodeText(0));
                    Assert.Single(reloaded.Slides[0].SmartArts);
                }
            } finally {
                if (File.Exists(filePath)) {
                    File.Delete(filePath);
                }
            }
        }
    }
}
