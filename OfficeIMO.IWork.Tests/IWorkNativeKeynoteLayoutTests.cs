using DocumentFormat.OpenXml.Packaging;
using OfficeIMO.IWork;
using OfficeIMO.PowerPoint;
using A = DocumentFormat.OpenXml.Drawing;
using P = DocumentFormat.OpenXml.Presentation;

namespace OfficeIMO.IWork.Tests;

public sealed class IWorkNativeKeynoteLayoutTests {
    private static string Corpus(string path) => Path.Combine(AppContext.BaseDirectory, "Documents", "IWorkCorpus", path);

    [Fact]
    public void Native_placeholder_geometry_text_frame_and_bullet_layout_match_Apple_export() {
        using var apple = PresentationDocument.Open(Corpus("native-exports/keynote-simple-v15.4.pptx"), false);
        SlidePart referenceSlide = apple.PresentationPart!.SlideParts.Single(part =>
            part.Slide!.InnerText.Contains("hello keynote"));
        var referenceShapes = referenceSlide.Slide!.CommonSlideData!.ShapeTree!.Elements<P.Shape>().ToArray();
        using var result = PowerPointIWorkConverter.ConvertKeynoteToPowerPointResult(Corpus("nim-iwork/simple.key"));
        result.Report.RequireCompleteEditableReconstruction();
        var sourceSlide = result.Projection.Slides[0];
        Assert.NotNull(sourceSlide.TitleBox!.Geometry);
        Assert.NotNull(sourceSlide.TitleBox.Layout);
        Assert.NotNull(Assert.Single(sourceSlide.TextBoxes).Layout);
        using var saved = new MemoryStream();
        result.Value.Save(saved); saved.Position = 0;
        using var actual = PresentationDocument.Open(saved, false);
        SlidePart destinationSlide = (SlidePart)actual.PresentationPart!.GetPartById(
            actual.PresentationPart.Presentation!.SlideIdList!.Elements<P.SlideId>().First().RelationshipId!);
        var shapes = destinationSlide.Slide!.CommonSlideData!.ShapeTree!.Elements<P.Shape>().ToArray();
        foreach (string text in new[] { "hello keynote", "first bullet" }) {
            var expected = referenceShapes.Single(shape => shape.InnerText.Contains(text));
            var placeholder = expected.NonVisualShapeProperties!.ApplicationNonVisualDrawingProperties!.GetFirstChild<P.PlaceholderShape>()!;
            var layoutShape = referenceSlide.SlideLayoutPart!.SlideLayout!.CommonSlideData!.ShapeTree!.Elements<P.Shape>()
                .Single(shape => {
                    var candidate = shape.NonVisualShapeProperties!.ApplicationNonVisualDrawingProperties!.GetFirstChild<P.PlaceholderShape>();
                    return candidate != null && candidate.Index?.Value == placeholder.Index?.Value;
                });
            var geometry = expected.ShapeProperties!.Transform2D ?? layoutShape.ShapeProperties!.Transform2D!;
            var converted = shapes.Single(shape => shape.InnerText.Contains(text));
            var actualGeometry = converted.ShapeProperties!.Transform2D!;
            Assert.InRange(Math.Abs(geometry.Offset!.X!.Value - actualGeometry.Offset!.X!.Value), 0, 1);
            Assert.InRange(Math.Abs(geometry.Offset.Y!.Value - actualGeometry.Offset.Y!.Value), 0, 1);
            Assert.InRange(Math.Abs(geometry.Extents!.Cx!.Value - actualGeometry.Extents!.Cx!.Value), 0, 2);
            Assert.InRange(Math.Abs(geometry.Extents.Cy!.Value - actualGeometry.Extents.Cy!.Value), 0, 2);
            var body = converted.TextBody!.BodyProperties!;
            var referenceBody = layoutShape.TextBody!.BodyProperties!;
            Assert.Equal(expected.TextBody!.BodyProperties!.Anchor?.Value
                ?? referenceBody.Anchor?.Value ?? A.TextAnchoringTypeValues.Top, body.Anchor!.Value);
            foreach (int inset in new[] { body.LeftInset!.Value, body.TopInset!.Value, body.RightInset!.Value, body.BottomInset!.Value })
                Assert.Equal(4 * 12700, inset);
            Assert.NotNull(body.GetFirstChild<A.NormalAutoFit>());
        }
        var bullet = shapes.Single(shape => shape.InnerText.Contains("first bullet")).TextBody!.Elements<A.Paragraph>().First().ParagraphProperties!;
        Assert.Equal(38 * 12700, bullet.LeftMargin!.Value);
        Assert.Equal(-38 * 12700, bullet.Indent!.Value);
        Assert.Equal(123000, bullet.GetFirstChild<A.BulletSizePercentage>()!.Val!.Value);
        Assert.Empty(result.Value.ValidateDocument());
    }
}
