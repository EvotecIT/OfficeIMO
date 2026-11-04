using System.IO;
using System.Linq;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Presentation;
using OfficeIMO.PowerPoint;
using Xunit;
using A = DocumentFormat.OpenXml.Drawing;

namespace OfficeIMO.Tests;

public sealed class PowerPointShapePhysicalLengthTests {
    [Fact]
    public void EffectLengthsAndDuplicateOffsetsPreserveSingleEmus() {
        using var output = new MemoryStream();
        using (var document = PowerPointDocument.Create()) {
            var slide = document.AddSlide();
            var shape = slide.AddRectangle(OfficeLength.Emus(100), OfficeLength.Emus(200),
                OfficeLength.Emus(300), OfficeLength.Emus(400), "Original");
            shape.SetShadow("000000", blur: OfficeLength.Emus(1), distance: OfficeLength.Emus(7));
            shape.SetGlow("445566", radius: OfficeLength.Emus(11));
            shape.SetBlur(OfficeLength.Emus(13));
            shape.SetSoftEdges(OfficeLength.Emus(17));
            shape.SetReflection(blur: OfficeLength.Emus(19), distance: OfficeLength.Emus(23));
            var copy = shape.Duplicate(OfficeLength.Emus(-1), OfficeLength.Emus(1));
            Assert.Equal(OfficeLength.Emus(99), copy.Left);
            Assert.Equal(OfficeLength.Emus(201), copy.Top);
            Assert.Throws<System.ArgumentOutOfRangeException>(() => shape.SetGlow("445566", OfficeLength.Emus(-1)));
            document.Save(output);
        }
        output.Position = 0;
        using var saved = PresentationDocument.Open(output, false);
        var shapes = saved.PresentationPart!.SlideParts.Single().Slide.CommonSlideData!.ShapeTree!.Elements<Shape>().ToArray();
        Assert.Equal(2, shapes.Length);
        foreach (var shape in shapes) {
            var effects = shape.ShapeProperties!.GetFirstChild<A.EffectList>()!;
            Assert.Equal(1L, effects.GetFirstChild<A.OuterShadow>()!.BlurRadius!.Value);
            Assert.Equal(7L, effects.GetFirstChild<A.OuterShadow>()!.Distance!.Value);
            Assert.Equal(11L, effects.GetFirstChild<A.Glow>()!.Radius!.Value);
            Assert.Equal(13L, effects.GetFirstChild<A.Blur>()!.Radius!.Value);
            Assert.Equal(17L, effects.GetFirstChild<A.SoftEdge>()!.Radius!.Value);
            Assert.Equal(19L, effects.GetFirstChild<A.Reflection>()!.BlurRadius!.Value);
            Assert.Equal(23L, effects.GetFirstChild<A.Reflection>()!.Distance!.Value);
        }
    }
}
