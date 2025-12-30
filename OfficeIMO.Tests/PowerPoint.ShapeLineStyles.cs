using System;
using System.IO;
using System.Linq;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Presentation;
using OfficeIMO.PowerPoint;
using Xunit;
using A = DocumentFormat.OpenXml.Drawing;

namespace OfficeIMO.Tests {
    public class PowerPointShapeLineStylesTests {
        [Fact]
        public void CanSetLineDashAndArrowheads() {
            string filePath = Path.Combine(Path.GetTempPath(), Guid.NewGuid() + ".pptx");
            try {
                using (PowerPointPresentation presentation = PowerPointPresentation.Create(filePath)) {
                    PowerPointSlide slide = presentation.AddSlide();
                    PowerPointAutoShape line = slide.AddLine(0, 0, 4000, 0, "ArrowLine");
                    line.OutlineColor = "FF0000";
                    line.OutlineDash = A.PresetLineDashValues.Dash;
                    line.SetLineEnds(A.LineEndValues.Triangle, A.LineEndValues.Stealth, A.LineEndWidthValues.Medium, A.LineEndLengthValues.Medium);
                    presentation.Save();
                }

                using (PresentationDocument document = PresentationDocument.Open(filePath, false)) {
                    SlidePart slidePart = document.PresentationPart!.SlideParts.First();
                    Shape lineShape = slidePart.Slide.CommonSlideData!.ShapeTree!
                        .Elements<Shape>()
                        .First(shape => shape.NonVisualShapeProperties?.NonVisualDrawingProperties?.Name?.Value == "ArrowLine");

                    A.Outline outline = lineShape.ShapeProperties!.GetFirstChild<A.Outline>()!;
                    Assert.Equal(A.PresetLineDashValues.Dash, outline.GetFirstChild<A.PresetDash>()?.Val?.Value);

                    A.HeadEnd? head = outline.GetFirstChild<A.HeadEnd>();
                    A.TailEnd? tail = outline.GetFirstChild<A.TailEnd>();
                    Assert.Equal(A.LineEndValues.Triangle, head?.Type?.Value);
                    Assert.Equal(A.LineEndValues.Stealth, tail?.Type?.Value);
                    Assert.Equal(A.LineEndWidthValues.Medium, head?.Width?.Value);
                    Assert.Equal(A.LineEndLengthValues.Medium, head?.Length?.Value);
                }
            } finally {
                if (File.Exists(filePath)) {
                    File.Delete(filePath);
                }
            }
        }
    }
}
