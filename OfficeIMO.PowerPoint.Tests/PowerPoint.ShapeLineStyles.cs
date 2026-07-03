using System;
using System.IO;
using System.Linq;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Presentation;
using DocumentFormat.OpenXml.Validation;
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

        [Fact]
        public void LineEndsStayAfterExistingLineJoinNodes() {
            string filePath = CreateTempFilePath(".pptx");
            try {
                using (PowerPointPresentation presentation = PowerPointPresentation.Create(filePath)) {
                    PowerPointSlide slide = presentation.AddSlide();
                    PowerPointAutoShape line = slide.AddLine(0, 0, 4000, 0, "JoinedArrowLine");
                    line.OutlineColor = "156082";
                    presentation.Save();
                }

                using (PresentationDocument document = PresentationDocument.Open(filePath, true)) {
                    SlidePart slidePart = document.PresentationPart!.SlideParts.First();
                    Shape lineShape = slidePart.Slide.CommonSlideData!.ShapeTree!
                        .Elements<Shape>()
                        .First(shape => shape.NonVisualShapeProperties?.NonVisualDrawingProperties?.Name?.Value == "JoinedArrowLine");
                    A.Outline outline = lineShape.ShapeProperties!.GetFirstChild<A.Outline>()!;
                    outline.Append(new A.Round());
                    slidePart.Slide.Save();
                }

                using (PowerPointPresentation presentation = PowerPointPresentation.Open(filePath)) {
                    PowerPointShape line = presentation.Slides[0].GetShape("JoinedArrowLine")!;
                    line.SetLineEnds(null, A.LineEndValues.Triangle, A.LineEndWidthValues.Medium, A.LineEndLengthValues.Medium);
                    presentation.Save();
                    Assert.Empty(presentation.ValidateDocument());
                }

                using (PresentationDocument document = PresentationDocument.Open(filePath, false)) {
                    OpenXmlValidator validator = new(FileFormatVersions.Microsoft365);
                    Assert.Empty(validator.Validate(document));

                    SlidePart slidePart = document.PresentationPart!.SlideParts.First();
                    Shape lineShape = slidePart.Slide.CommonSlideData!.ShapeTree!
                        .Elements<Shape>()
                        .First(shape => shape.NonVisualShapeProperties?.NonVisualDrawingProperties?.Name?.Value == "JoinedArrowLine");
                    A.Outline outline = lineShape.ShapeProperties!.GetFirstChild<A.Outline>()!;
                    var children = outline.ChildElements.ToList();
                    int roundIndex = children.FindIndex(child => child is A.Round);
                    int headIndex = children.FindIndex(child => child is A.HeadEnd);

                    Assert.True(roundIndex >= 0);
                    Assert.True(headIndex >= 0);
                    Assert.True(roundIndex < headIndex);
                }
            } finally {
                if (File.Exists(filePath)) {
                    File.Delete(filePath);
                }
            }
        }

        [Fact]
        public void OutlineChildrenStayInSchemaOrderWhenStylingAfterArrowheads() {
            string filePath = CreateTempFilePath(".pptx");
            try {
                using (PowerPointPresentation presentation = PowerPointPresentation.Create(filePath)) {
                    PowerPointSlide slide = presentation.AddSlide();
                    PowerPointAutoShape line = slide.AddLine(0, 0, 4000, 0, "ReverseStyledArrowLine");

                    line.SetLineEnds(A.LineEndValues.Triangle, A.LineEndValues.Stealth, A.LineEndWidthValues.Medium, A.LineEndLengthValues.Medium);
                    line.OutlineColor = "156082";
                    line.OutlineDash = A.PresetLineDashValues.DashDot;

                    presentation.Save();
                    Assert.Empty(presentation.ValidateDocument());
                }

                using (PresentationDocument document = PresentationDocument.Open(filePath, false)) {
                    OpenXmlValidator validator = new(FileFormatVersions.Microsoft365);
                    Assert.Empty(validator.Validate(document));

                    SlidePart slidePart = document.PresentationPart!.SlideParts.First();
                    Shape lineShape = slidePart.Slide.CommonSlideData!.ShapeTree!
                        .Elements<Shape>()
                        .First(shape => shape.NonVisualShapeProperties?.NonVisualDrawingProperties?.Name?.Value == "ReverseStyledArrowLine");

                    A.Outline outline = lineShape.ShapeProperties!.GetFirstChild<A.Outline>()!;
                    var children = outline.ChildElements.ToList();
                    int fillIndex = children.FindIndex(child => child is A.SolidFill);
                    int dashIndex = children.FindIndex(child => child is A.PresetDash);
                    int headIndex = children.FindIndex(child => child is A.HeadEnd);
                    int tailIndex = children.FindIndex(child => child is A.TailEnd);

                    Assert.True(fillIndex >= 0);
                    Assert.True(dashIndex >= 0);
                    Assert.True(headIndex >= 0);
                    Assert.True(tailIndex >= 0);
                    Assert.True(fillIndex < dashIndex);
                    Assert.True(dashIndex < headIndex);
                    Assert.True(headIndex < tailIndex);
                }
            } finally {
                if (File.Exists(filePath)) {
                    File.Delete(filePath);
                }
            }
        }

        private static string CreateTempFilePath(string extension) {
            string path = Path.GetTempFileName();
            File.Delete(path);
            return Path.ChangeExtension(path, extension);
        }
    }
}
