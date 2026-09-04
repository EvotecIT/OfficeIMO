using OfficeIMO.Pdf;
using System.Threading;
using Xunit;

namespace OfficeIMO.Tests.Pdf;

public sealed class PdfRedactionRegionAndAnnotationTests {
    [Fact]
    public void PolygonRetainsExactGeometryInsideItsBoundedArea() {
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
        Assert.True(area.IntersectsRectangle(20, 25, 2, 2));
        Assert.False(area.IntersectsRectangle(39, 19, 2, 2));
    }

    [Fact]
    public void PolygonRejectsSelfIntersectionAndQuadrilateralNormalizesProducerPointOrder() {
        Assert.Throws<ArgumentException>(() => PdfRedactionRegion.Polygon(1, new[] {
            new PdfRedactionPoint(0, 0),
            new PdfRedactionPoint(20, 20),
            new PdfRedactionPoint(0, 20),
            new PdfRedactionPoint(20, 0)
        }));

        PdfRedactionArea quadrilateral = Assert.Single(PdfRedactionRegion.Quadrilateral(1, new[] {
            new PdfRedactionPoint(20, 40),
            new PdfRedactionPoint(80, 40),
            new PdfRedactionPoint(20, 20),
            new PdfRedactionPoint(80, 20)
        }).Areas);

        Assert.True(quadrilateral.ContainsPoint(50, 30));
    }

    [Fact]
    public void PolygonApplicationPreservesTextInsideBoundsButOutsideReviewedShape() {
        PdfDocument source = PdfDocument.Create()
            .Paragraph(paragraph => paragraph.Text("REMOVE-EXACT"))
            .Paragraph(paragraph => paragraph.Text("KEEP-OUTSIDE-SHAPE"));
        PdfTextSpan[] spans = PdfReadDocument.Open(source.ToBytes()).Pages[0].GetTextSpans().ToArray();
        PdfTextSpan remove = Assert.Single(spans, static span => span.Text.Contains("REMOVE-EXACT", StringComparison.Ordinal));
        PdfTextSpan keep = Assert.Single(spans, static span => span.Text.Contains("KEEP-OUTSIDE-SHAPE", StringComparison.Ordinal));
        PdfTextSpanBounds removeBounds = PdfTextSpanGeometry.GetAxisAlignedBounds(remove);
        PdfTextSpanBounds keepBounds = PdfTextSpanGeometry.GetAxisAlignedBounds(keep);
        double tailLeft = Math.Max(removeBounds.Right, keepBounds.Right) + 20D;
        PdfRedactionRegion region = PdfRedactionRegion.Polygon(1, new[] {
            new PdfRedactionPoint(removeBounds.Left - 2, removeBounds.Bottom - 2),
            new PdfRedactionPoint(tailLeft, removeBounds.Bottom - 2),
            new PdfRedactionPoint(tailLeft, keepBounds.Bottom - 2),
            new PdfRedactionPoint(tailLeft + 10, keepBounds.Bottom - 2),
            new PdfRedactionPoint(tailLeft + 10, keepBounds.Top + 2),
            new PdfRedactionPoint(tailLeft, keepBounds.Top + 2),
            new PdfRedactionPoint(tailLeft, removeBounds.Top + 2),
            new PdfRedactionPoint(removeBounds.Left - 2, removeBounds.Top + 2)
        });

        PdfRedactionPlan plan = source.Redactions.Plan(new[] { region });
        PdfDocument redacted = source.Redactions.Apply(plan);
        string text = redacted.Reader.Text();

        Assert.Contains(plan.Matches, static match => match.Text?.Contains("REMOVE-EXACT", StringComparison.Ordinal) == true);
        Assert.DoesNotContain(plan.Matches, static match => match.Text?.Contains("KEEP-OUTSIDE-SHAPE", StringComparison.Ordinal) == true);
        Assert.DoesNotContain("REMOVE-EXACT", text, StringComparison.Ordinal);
        Assert.Contains("KEEP-OUTSIDE-SHAPE", text, StringComparison.Ordinal);
        Assert.Contains(" m", PdfEncoding.Latin1GetString(redacted.ToBytes()), StringComparison.Ordinal);
    }

    [Fact]
    public void ExactPolygonFailsClosedForPartiallyIntersectedVectorPath() {
        PdfDocument source = PdfDocument.Create()
            .Rectangle(180, 80, fillColor: new PdfColor(0D, 0D, 1D))
            .Paragraph(paragraph => paragraph.Text("vector safety"));
        PdfReadPage page = PdfReadDocument.Open(source.ToBytes()).Pages[0];
        PdfPageVisualPrimitive rectangle = Assert.Single(page.GetIdentityVisualPrimitives(), static primitive => primitive.Kind == PdfPageVisualPrimitiveKind.Rectangle);
        PdfVisualBounds bounds = page.TransformVisualBoundsToUser(rectangle.X, rectangle.Y, rectangle.X + rectangle.Width, rectangle.Y + rectangle.Height);
        PdfRedactionRegion region = PdfRedactionRegion.Polygon(1, new[] {
            new PdfRedactionPoint(bounds.Left, bounds.Top),
            new PdfRedactionPoint(bounds.Left + bounds.Width / 2D, bounds.Top),
            new PdfRedactionPoint(bounds.Left, bounds.Top + bounds.Height / 2D)
        });
        PdfRedactionPlan plan = source.Redactions.Plan(new[] { region });

        NotSupportedException exception = Assert.Throws<NotSupportedException>(() => source.Redactions.Apply(plan));

        Assert.Contains("outside the reviewed geometry", exception.Message, StringComparison.Ordinal);
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
        PdfRedactionArea first = region.Areas[0];
        Assert.True(first.ContainsPoint(30, 25));
        Assert.False(first.ContainsPoint(first.X + 0.1D, first.Top - 0.1D));
    }

    [Fact]
    public void FreehandCapsuleDetectsAnEndpointNearTheMiddleOfARectangleEdge() {
        PdfRedactionArea area = Assert.Single(PdfRedactionRegion.Freehand(1, new[] {
            new PdfRedactionPoint(5, -2),
            new PdfRedactionPoint(5, -1)
        }, 2.2D).Areas);

        Assert.True(area.IntersectsRectangle(0, 0, 10, 10));
        Assert.False(area.IntersectsRectangle(0, 1, 10, 10));
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
    public void AuthoredQuadrilateralRedactAnnotationRoundTripsExactGeometry() {
        PdfDocument source = PdfDocument.Create().Paragraph(paragraph => paragraph.Text("quadrilateral annotation"));
        PdfRedactionRegion region = PdfRedactionRegion.Quadrilateral(1, new[] {
            new PdfRedactionPoint(20, 600),
            new PdfRedactionPoint(140, 600),
            new PdfRedactionPoint(100, 640),
            new PdfRedactionPoint(20, 640)
        });

        PdfDocument annotated = source.Redactions.AddAnnotation(new PdfRedactionAnnotationOptions(region));
        PdfAnnotation annotation = Assert.Single(annotated.Reader.AnnotationsBySubtype("Redact"));
        PdfRedactionArea roundTripped = Assert.Single(annotated.Redactions.PlanAnnotations().Areas);

        Assert.Equal(8, annotation.QuadPoints.Count);
        Assert.True(roundTripped.IntersectsRectangle(40, 610, 4, 4));
        Assert.False(roundTripped.IntersectsRectangle(125, 632, 4, 4));
    }

    [Fact]
    public void SignedPdfCanProduceAnExplicitUnsignedDerivative() {
        byte[] source = PdfDocument.Create().Paragraph(paragraph => paragraph.Text("signed derivative content")).ToBytes();
        PdfExternalSignaturePreparation preparation = PdfIncrementalUpdater.PrepareExternalSignature(source, new PdfExternalSignatureOptions {
            FieldName = "OriginalSignature",
            ReservedSignatureContentsBytes = 512
        });
        byte[] signed = PdfIncrementalUpdater.ApplyExternalSignature(preparation, Enumerable.Repeat((byte)0x5A, 128).ToArray());
        Assert.True(PdfDocument.Load(signed).Security.ValidateSignatures().HasSignatures);

        PdfUnsignedDerivativeResult derivative = PdfDocument.Load(signed).Security.CreateUnsignedDerivative();

        Assert.Equal(1, derivative.RemovedSignatureCount);
        Assert.False(derivative.ToDocument().Security.ValidateSignatures().HasSignatures);
        Assert.Contains("signed derivative content", derivative.ToDocument().Reader.Text(), StringComparison.Ordinal);
    }

    [Fact]
    public void UnsignedEncryptedPdfStillProducesAnUnencryptedFullRewriteDerivative() {
        const string ownerPassword = "owner-unsigned-derivative";
        byte[] encrypted = PdfDocument.Create().Paragraph(paragraph => paragraph.Text("unsigned protected content"))
            .Security.Encrypt(new PdfStandardEncryptionOptions("reader-unsigned-derivative") { OwnerPassword = ownerPassword }).Pdf;
        Assert.True(PdfDocument.Load(encrypted, new PdfLoadOptions { Password = ownerPassword }).Inspect().Security.HasEncryption);

        PdfUnsignedDerivativeResult derivative = PdfDocument
            .Load(encrypted, new PdfLoadOptions { Password = ownerPassword })
            .Security.CreateUnsignedDerivative();

        Assert.Equal(0, derivative.RemovedSignatureCount);
        Assert.False(derivative.ToDocument().Inspect().Security.HasEncryption);
        Assert.Contains("unsigned protected content", derivative.ToDocument().Reader.Text(), StringComparison.Ordinal);
        Assert.False(encrypted.AsSpan().SequenceEqual(derivative.Pdf));
    }

    [Fact]
    public void UnsignedDerivativeHonorsPreCancelledOperation() {
        byte[] source = PdfDocument.Create().Paragraph(paragraph => paragraph.Text("cancel derivative")).ToBytes();
        using var cancellation = new CancellationTokenSource();
        cancellation.Cancel();

        Assert.Throws<OperationCanceledException>(() =>
            PdfDocument.Load(source).Security.CreateUnsignedDerivative(cancellation.Token));
    }

    [Fact]
    public void UnsignedDerivativeRemovesSignaturePermissionsAndValidationStore() {
        byte[] signed = PdfRewritePreservationTestSupport.BuildSignedIncrementalProofPdf();

        PdfUnsignedDerivativeResult derivative = PdfDocument.Load(signed).Security.CreateUnsignedDerivative();
        string raw = PdfEncoding.Latin1GetString(derivative.Pdf);

        Assert.True(derivative.RemovedSignatureCount >= 1);
        Assert.False(derivative.ToDocument().Security.ValidateSignatures().HasSignatures);
        Assert.DoesNotContain("/Perms", raw, StringComparison.Ordinal);
        Assert.DoesNotContain("/DSS", raw, StringComparison.Ordinal);
        Assert.DoesNotContain("/SigFlags", raw, StringComparison.Ordinal);
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
        Assert.All(plan.Areas, static area => Assert.True(area.ContainsPoint(area.X + area.Width / 2D, area.Y + area.Height / 2D)));
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
