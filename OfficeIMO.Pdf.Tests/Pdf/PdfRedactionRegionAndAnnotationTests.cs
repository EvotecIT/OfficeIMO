using OfficeIMO.Pdf;
using System.Threading;
using Xunit;

namespace OfficeIMO.Tests.Pdf;

public sealed class PdfRedactionRegionAndAnnotationTests {
    [Fact]
    public void PolygonUsesConservativeBoundsAcrossPlanningAndEvidence() {
        PdfRedactionRegion region = PdfRedactionRegion.Polygon(2, new[] {
            new PdfRedactionPoint(10, 20),
            new PdfRedactionPoint(30, 18),
            new PdfRedactionPoint(42, 35),
            new PdfRedactionPoint(12, 40)
        }, "review-shape");

        PdfRedactionArea area = Assert.Single(region.Areas);
        Assert.Equal(PdfRedactionRegionKind.Polygon, region.Kind);
        Assert.Equal(10, area.X);
        Assert.Equal(18, area.Y);
        Assert.Equal(32, area.Width);
        Assert.Equal(22, area.Height);
        Assert.Equal("review-shape", area.Label);
    }

    [Fact]
    public void FreehandCreatesBoundedSegmentAreas() {
        PdfRedactionRegion region = PdfRedactionRegion.Freehand(1, new[] {
            new PdfRedactionPoint(20, 20),
            new PdfRedactionPoint(40, 30),
            new PdfRedactionPoint(50, 10)
        }, 4);

        Assert.Equal(2, region.Areas.Count);
        Assert.All(region.Areas, area => {
            Assert.True(area.Width >= 4);
            Assert.True(area.Height >= 4);
        });
    }

    [Fact]
    public void StandardRedactAnnotationRoundTripsIntoSourceBoundPlan() {
        PdfDocument source = PdfDocument.Create().Paragraph(paragraph => paragraph.Text("Annotation review target"));
        PdfDocument annotated = source.Redactions.AddAnnotation(new PdfRedactionAnnotationOptions(
            PdfRedactionRegion.Rectangle(1, 30, 700, 160, 30, "review-1")) {
            Contents = "Approved review area",
            Name = "redaction-review-1"
        });

        PdfAnnotation annotation = Assert.Single(annotated.Reader.AnnotationsBySubtype("Redact"));
        Assert.Equal("redaction-review-1", annotation.Name);
        PdfRedactionPlan plan = annotated.Redactions.PlanAnnotations();
        Assert.True(plan.IsReviewable);
        Assert.Equal(annotation.X1, Assert.Single(plan.Areas).X);
    }

    [Fact]
    public void IndependentMultiQuadRedactAnnotationPreservesTheGapBetweenQuads() {
        PdfDocument source = PdfDocument.Create().Paragraph(paragraph => paragraph.Text("Independent producer annotation"));
        PdfDocument annotated = source.Annotations.Add(new PdfAnnotationCreateOptions {
            PageNumber = 1,
            Subtype = "Redact",
            Rectangle = new[] { 20D, 600D, 220D, 700D },
            QuadPoints = new[] {
                20D, 700D, 80D, 700D, 20D, 680D, 80D, 680D,
                160D, 620D, 220D, 620D, 160D, 600D, 220D, 600D
            },
            GenerateAppearance = false
        }).ToDocument();

        PdfRedactionPlan plan = annotated.Redactions.PlanAnnotations();

        Assert.Equal(2, plan.Areas.Count);
        Assert.DoesNotContain(plan.Areas, area => area.X < 120D && area.Right > 120D && area.Y < 650D && area.Top > 650D);
    }

    [Fact]
    public void MultiAreaAnnotationNamesAreUnique() {
        PdfRedactionRegion region = PdfRedactionRegion.Freehand(1, new[] {
            new PdfRedactionPoint(20, 20),
            new PdfRedactionPoint(40, 30),
            new PdfRedactionPoint(50, 10)
        }, 4);
        PdfDocument annotated = PdfDocument.Create().Paragraph(paragraph => paragraph.Text("annotation names"))
            .Redactions.AddAnnotation(new PdfRedactionAnnotationOptions(region) { Name = "review-path" });

        string?[] names = annotated.Reader.AnnotationsBySubtype("Redact").Select(static annotation => annotation.Name).ToArray();

        Assert.Equal(new[] { "review-path:1", "review-path:2" }, names);
    }

    [Fact]
    public void AnnotationAuthoringRejectsExpensiveMultiRewriteRegionsBeforeMutation() {
        PdfRedactionPoint[] points = Enumerable.Range(0, 18)
            .Select(index => new PdfRedactionPoint(20 + index * 4, 20 + index * 2))
            .ToArray();
        PdfRedactionRegion region = PdfRedactionRegion.Freehand(1, points, 4);
        PdfDocument source = PdfDocument.Create().Paragraph(paragraph => paragraph.Text("bounded annotation authoring"));

        InvalidOperationException exception = Assert.Throws<InvalidOperationException>(() =>
            source.Redactions.AddAnnotation(new PdfRedactionAnnotationOptions(region)));

        Assert.Contains("annotation limit", exception.Message, StringComparison.OrdinalIgnoreCase);
        Assert.Empty(source.Reader.AnnotationsBySubtype("Redact"));
    }

    [Fact]
    public void PlanningApplyingAndVerificationHonorPreCancelledTokens() {
        PdfDocument source = PdfDocument.Create().Paragraph(paragraph => paragraph.Text("cancellation target"));
        var area = new PdfRedactionArea(1, 10, 10, 20, 20);
        using var cancellation = new CancellationTokenSource();
        cancellation.Cancel();

        Assert.ThrowsAny<OperationCanceledException>(() => source.Redactions.Plan(new[] { area }, cancellation.Token));
        Assert.ThrowsAny<OperationCanceledException>(() => source.Redactions.Apply(new[] { area }, new PdfRedactionApplyOptions { CancellationToken = cancellation.Token }));
        Assert.ThrowsAny<OperationCanceledException>(() => source.Redactions.Verify(new PdfRedactionVerificationOptions { CancellationToken = cancellation.Token }));
    }
}
