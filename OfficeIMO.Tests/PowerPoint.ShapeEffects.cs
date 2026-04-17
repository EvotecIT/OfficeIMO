using System;
using System.IO;
using System.Linq;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Presentation;
using OfficeIMO.PowerPoint;
using Xunit;
using A = DocumentFormat.OpenXml.Drawing;

namespace OfficeIMO.Tests {
    public class PowerPointShapeEffectsTests {
        [Fact]
        public void CanSetAndClearShapeShadow() {
            string filePath = Path.Combine(Path.GetTempPath(), Guid.NewGuid() + ".pptx");
            try {
                using (PowerPointPresentation presentation = PowerPointPresentation.Create(filePath)) {
                    PowerPointSlide slide = presentation.AddSlide();
                    PowerPointAutoShape shadow = slide.AddRectangle(0, 0, 4000, 2000, "ShadowRect");
                    shadow.SetShadow("000000", blurPoints: 6, distancePoints: 5, angleDegrees: 270, transparencyPercent: 40);

                    PowerPointAutoShape clear = slide.AddRectangle(0, 3000, 4000, 2000, "ClearRect");
                    clear.SetShadow("FF0000", blurPoints: 4, distancePoints: 3, angleDegrees: 45, transparencyPercent: 20);
                    clear.ClearShadow();

                    presentation.Save();
                }

                using (PresentationDocument document = PresentationDocument.Open(filePath, false)) {
                    SlidePart slidePart = document.PresentationPart!.SlideParts.First();
                    Shape shadowShape = slidePart.Slide.CommonSlideData!.ShapeTree!.Elements<Shape>()
                        .First(shape => shape.NonVisualShapeProperties?.NonVisualDrawingProperties?.Name?.Value == "ShadowRect");
                    Shape clearShape = slidePart.Slide.CommonSlideData!.ShapeTree!.Elements<Shape>()
                        .First(shape => shape.NonVisualShapeProperties?.NonVisualDrawingProperties?.Name?.Value == "ClearRect");

                    A.OuterShadow? shadow = shadowShape.ShapeProperties?
                        .GetFirstChild<A.EffectList>()?
                        .GetFirstChild<A.OuterShadow>();
                    Assert.NotNull(shadow);

                    A.RgbColorModelHex? color = shadow?.GetFirstChild<A.RgbColorModelHex>();
                    Assert.Equal("000000", color?.Val?.Value);
                    Assert.Equal(60000, color?.GetFirstChild<A.Alpha>()?.Val?.Value);
                    Assert.Equal(PowerPointUnits.FromPoints(6), shadow?.BlurRadius?.Value);
                    Assert.Equal(PowerPointUnits.FromPoints(5), shadow?.Distance?.Value);
                    Assert.Equal(16200000, shadow?.Direction?.Value);

                    A.OuterShadow? clearedShadow = clearShape.ShapeProperties?
                        .GetFirstChild<A.EffectList>()?
                        .GetFirstChild<A.OuterShadow>();
                    Assert.Null(clearedShadow);
                }
            } finally {
                if (File.Exists(filePath)) {
                    File.Delete(filePath);
                }
            }
        }

        [Fact]
        public void CanSetGlowAndSoftEdges() {
            string filePath = Path.Combine(Path.GetTempPath(), Guid.NewGuid() + ".pptx");
            try {
                using (PowerPointPresentation presentation = PowerPointPresentation.Create(filePath)) {
                    PowerPointSlide slide = presentation.AddSlide();
                    PowerPointAutoShape glowShape = slide.AddRectangle(0, 0, 4000, 2000, "GlowRect");
                    glowShape.SetGlow("FF00FF", radiusPoints: 5, transparencyPercent: 25);
                    glowShape.SetSoftEdges(3);

                    PowerPointAutoShape clearShape = slide.AddRectangle(0, 3000, 4000, 2000, "ClearEffectsRect");
                    clearShape.SetGlow("00FF00", radiusPoints: 4, transparencyPercent: 10);
                    clearShape.SetSoftEdges(2);
                    clearShape.ClearGlow();
                    clearShape.ClearSoftEdges();

                    presentation.Save();
                }

                using (PresentationDocument document = PresentationDocument.Open(filePath, false)) {
                    SlidePart slidePart = document.PresentationPart!.SlideParts.First();
                    Shape glowXml = slidePart.Slide.CommonSlideData!.ShapeTree!.Elements<Shape>()
                        .First(shape => shape.NonVisualShapeProperties?.NonVisualDrawingProperties?.Name?.Value == "GlowRect");
                    Shape clearXml = slidePart.Slide.CommonSlideData!.ShapeTree!.Elements<Shape>()
                        .First(shape => shape.NonVisualShapeProperties?.NonVisualDrawingProperties?.Name?.Value == "ClearEffectsRect");

                    A.EffectList? effects = glowXml.ShapeProperties?.GetFirstChild<A.EffectList>();
                    Assert.NotNull(effects);

                    A.Glow? glow = effects?.GetFirstChild<A.Glow>();
                    Assert.NotNull(glow);
                    Assert.Equal(PowerPointUnits.FromPoints(5), glow?.Radius?.Value);

                    A.RgbColorModelHex? color = glow?.GetFirstChild<A.RgbColorModelHex>();
                    Assert.Equal("FF00FF", color?.Val?.Value);
                    Assert.Equal(75000, color?.GetFirstChild<A.Alpha>()?.Val?.Value);

                    A.SoftEdge? softEdge = effects?.GetFirstChild<A.SoftEdge>();
                    Assert.NotNull(softEdge);
                    Assert.Equal(PowerPointUnits.FromPoints(3), softEdge?.Radius?.Value);

                    A.EffectList? clearedEffects = clearXml.ShapeProperties?.GetFirstChild<A.EffectList>();
                    Assert.Null(clearedEffects?.GetFirstChild<A.Glow>());
                    Assert.Null(clearedEffects?.GetFirstChild<A.SoftEdge>());
                }
            } finally {
                if (File.Exists(filePath)) {
                    File.Delete(filePath);
                }
            }
        }

        [Fact]
        public void CanSetBlurAndReflection() {
            string filePath = Path.Combine(Path.GetTempPath(), Guid.NewGuid() + ".pptx");
            try {
                using (PowerPointPresentation presentation = PowerPointPresentation.Create(filePath)) {
                    PowerPointSlide slide = presentation.AddSlide();
                    PowerPointAutoShape blurShape = slide.AddRectangle(0, 0, 4000, 2000, "BlurRect");
                    blurShape.SetBlur(5, grow: true);

                    PowerPointAutoShape reflectionShape = slide.AddRectangle(0, 3000, 4000, 2000, "ReflectionRect");
                    reflectionShape.SetReflection(blurPoints: 6, distancePoints: 4,
                        directionDegrees: 270, fadeDirectionDegrees: 90,
                        startOpacityPercent: 60, endOpacityPercent: 0,
                        startPositionPercent: 0, endPositionPercent: 100,
                        alignment: A.RectangleAlignmentValues.Bottom,
                        rotateWithShape: true);

                    presentation.Save();
                }

                using (PresentationDocument document = PresentationDocument.Open(filePath, false)) {
                    SlidePart slidePart = document.PresentationPart!.SlideParts.First();
                    Shape blurXml = slidePart.Slide.CommonSlideData!.ShapeTree!.Elements<Shape>()
                        .First(shape => shape.NonVisualShapeProperties?.NonVisualDrawingProperties?.Name?.Value == "BlurRect");
                    Shape reflectionXml = slidePart.Slide.CommonSlideData!.ShapeTree!.Elements<Shape>()
                        .First(shape => shape.NonVisualShapeProperties?.NonVisualDrawingProperties?.Name?.Value == "ReflectionRect");

                    A.Blur? blur = blurXml.ShapeProperties?.GetFirstChild<A.EffectList>()?.GetFirstChild<A.Blur>();
                    Assert.NotNull(blur);
                    Assert.Equal(PowerPointUnits.FromPoints(5), blur?.Radius?.Value);
                    Assert.True(blur?.Grow?.Value == true);

                    A.Reflection? reflection = reflectionXml.ShapeProperties?.GetFirstChild<A.EffectList>()?
                        .GetFirstChild<A.Reflection>();
                    Assert.NotNull(reflection);
                    Assert.Equal(PowerPointUnits.FromPoints(6), reflection?.BlurRadius?.Value);
                    Assert.Equal(PowerPointUnits.FromPoints(4), reflection?.Distance?.Value);
                    Assert.Equal(16200000, reflection?.Direction?.Value);
                    Assert.Equal(5400000, reflection?.FadeDirection?.Value);
                    Assert.Equal(60000, reflection?.StartOpacity?.Value);
                    Assert.Equal(0, reflection?.EndAlpha?.Value);
                    Assert.Equal(0, reflection?.StartPosition?.Value);
                    Assert.Equal(100000, reflection?.EndPosition?.Value);
                    Assert.Equal(A.RectangleAlignmentValues.Bottom, reflection?.Alignment?.Value);
                    Assert.True(reflection?.RotateWithShape?.Value == true);
                }
            } finally {
                if (File.Exists(filePath)) {
                    File.Delete(filePath);
                }
            }
        }

        [Fact]
        public void ShapePropertiesStayInSchemaOrderWhenStylingAfterEffects() {
            string filePath = Path.Combine(Path.GetTempPath(), Guid.NewGuid() + ".pptx");
            try {
                using (PowerPointPresentation presentation = PowerPointPresentation.Create(filePath)) {
                    PowerPointSlide slide = presentation.AddSlide();
                    PowerPointAutoShape shape = slide.AddRectangle(0, 0, 4000, 2000, "StyledAfterEffects");

                    shape.SetSoftEdges(2);
                    shape.SetReflection(blurPoints: 2, distancePoints: 1, startOpacityPercent: 20);
                    shape.SetShadow("000000", blurPoints: 6, distancePoints: 3, transparencyPercent: 55);
                    shape.SetGlow("F26A3D", radiusPoints: 3, transparencyPercent: 30);
                    shape.FillColor = "EFE8DA";
                    shape.FillTransparency = 8;
                    shape.OutlineColor = "6B6EA8";
                    shape.OutlineWidthPoints = 1.2;

                    presentation.Save();
                    Assert.Empty(presentation.ValidateDocument());
                }

                using (PresentationDocument document = PresentationDocument.Open(filePath, false)) {
                    SlidePart slidePart = document.PresentationPart!.SlideParts.First();
                    Shape shape = slidePart.Slide.CommonSlideData!.ShapeTree!.Elements<Shape>()
                        .First(xmlShape => xmlShape.NonVisualShapeProperties?.NonVisualDrawingProperties?.Name?.Value == "StyledAfterEffects");

                    var shapePropertyChildren = shape.ShapeProperties!.ChildElements.ToList();
                    int fillIndex = shapePropertyChildren.FindIndex(child => child is A.SolidFill);
                    int outlineIndex = shapePropertyChildren.FindIndex(child => child is A.Outline);
                    int effectsIndex = shapePropertyChildren.FindIndex(child => child is A.EffectList);

                    Assert.True(fillIndex >= 0);
                    Assert.True(outlineIndex >= 0);
                    Assert.True(effectsIndex >= 0);
                    Assert.True(fillIndex < outlineIndex);
                    Assert.True(outlineIndex < effectsIndex);

                    var effectChildren = shape.ShapeProperties.GetFirstChild<A.EffectList>()!.ChildElements.ToList();
                    int glowIndex = effectChildren.FindIndex(child => child is A.Glow);
                    int shadowIndex = effectChildren.FindIndex(child => child is A.OuterShadow);
                    int reflectionIndex = effectChildren.FindIndex(child => child is A.Reflection);
                    int softEdgeIndex = effectChildren.FindIndex(child => child is A.SoftEdge);

                    Assert.True(glowIndex >= 0);
                    Assert.True(shadowIndex >= 0);
                    Assert.True(reflectionIndex >= 0);
                    Assert.True(softEdgeIndex >= 0);
                    Assert.True(glowIndex < shadowIndex);
                    Assert.True(shadowIndex < reflectionIndex);
                    Assert.True(reflectionIndex < softEdgeIndex);
                }
            } finally {
                if (File.Exists(filePath)) {
                    File.Delete(filePath);
                }
            }
        }
    }
}
