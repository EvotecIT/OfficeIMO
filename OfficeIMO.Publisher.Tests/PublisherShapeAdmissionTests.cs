using OfficeIMO.Drawing;

namespace OfficeIMO.Publisher.Tests;

public sealed class PublisherShapeAdmissionTests {
    [Fact]
    public void Invalid_creation_bounds_paint_identity_and_cancellation_leave_the_model_and_allocator_unchanged() {
        PublisherDocument document = PublisherShapeEditingFixture.Create();
        var pages = document.Pages; var report = document.ReadReport; string before = document.ToSvg();
        var valid = OfficeShape.Rectangle(30, 20); valid.FillColor = OfficeColor.Black;
        Assert.Throws<ArgumentNullException>(() => document.AddShape(10, null!, OfficeLength.Points(0), OfficeLength.Points(0)));
        Assert.Throws<ArgumentOutOfRangeException>(() => document.AddShape(uint.MaxValue, valid, OfficeLength.Points(0), OfficeLength.Points(0)));
        Assert.Throws<ArgumentOutOfRangeException>(() => OfficeLength.Points(double.NaN));
        Assert.Throws<ArgumentOutOfRangeException>(() => OfficeLength.Points(double.PositiveInfinity));
        Assert.Throws<ArgumentOutOfRangeException>(() => document.AddShape(10, OfficeShape.Rectangle(0, 10), OfficeLength.Points(0), OfficeLength.Points(0)));
        Assert.Throws<OverflowException>(() => OfficeLength.Points(double.MaxValue));
        Assert.Throws<ArgumentOutOfRangeException>(() => document.SetShapeBounds(101, OfficeLength.Points(0), OfficeLength.Points(0), OfficeLength.Points(0), OfficeLength.Points(10)));
        Assert.Throws<ArgumentOutOfRangeException>(() => OfficeLength.Points(double.NaN));
        Assert.Throws<ArgumentOutOfRangeException>(() => document.SetShapeBounds(uint.MaxValue, OfficeLength.Points(0), OfficeLength.Points(0), OfficeLength.Points(10), OfficeLength.Points(10)));
        Assert.Throws<ArgumentOutOfRangeException>(() => document.SetShapePaint(uint.MaxValue, OfficeColor.Black, null));
        Assert.Throws<ArgumentOutOfRangeException>(() => document.SetShapePaint(101, OfficeColor.Black, null, OfficeLength.Points(-1)));
        Assert.Throws<ArgumentOutOfRangeException>(() => OfficeLength.Points(double.NaN));
        Assert.Throws<ArgumentOutOfRangeException>(() => document.SetShapePaint(101, OfficeColor.Black, null, fillOpacity: 1.001));
        Assert.Throws<ArgumentOutOfRangeException>(() => document.SetShapePaint(101, OfficeColor.Black, null, strokeOpacity: -.001));
        Assert.Throws<ArgumentOutOfRangeException>(() => document.SetShapePaint(101, OfficeColor.Black, null, strokeOpacity: double.PositiveInfinity));
        using var cancellation = new CancellationTokenSource(); cancellation.Cancel();
        Assert.Throws<OperationCanceledException>(() => document.AddShape(10, valid, OfficeLength.Points(0), OfficeLength.Points(0), cancellation.Token));
        Assert.Throws<OperationCanceledException>(() => document.SetShapeBounds(101, OfficeLength.Points(0), OfficeLength.Points(0), OfficeLength.Points(10), OfficeLength.Points(10), cancellation.Token));
        Assert.Throws<OperationCanceledException>(() => document.SetShapePaint(101, OfficeColor.Black, null, cancellationToken: cancellation.Token));
        Assert.Same(pages, document.Pages); Assert.Same(report, document.ReadReport); Assert.Equal(before, document.ToSvg());
        PublisherDocument control = PublisherShapeEditingFixture.Create();
        Assert.Equal(control.AddShape(10, valid, OfficeLength.Points(0), OfficeLength.Points(0)).Id, document.AddShape(10, valid, OfficeLength.Points(0), OfficeLength.Points(0)).Id);
    }

    [Fact]
    public void Creation_rejects_unsupported_geometry_effects_and_invalid_mutated_descriptors_without_silent_downgrade() {
        PublisherDocument document = PublisherShapeEditingFixture.Create();
        var pages = document.Pages;
        var transformed = OfficeShape.Rectangle(10, 10); transformed.Transform = OfficeTransform.Identity;
        Assert.Throws<NotSupportedException>(() => document.AddShape(10, transformed, OfficeLength.Points(0), OfficeLength.Points(0)));
        var dashed = OfficeShape.Rectangle(10, 10); dashed.SetStrokeDashArray(new[] { 1D, 2D });
        Assert.Throws<NotSupportedException>(() => document.AddShape(10, dashed, OfficeLength.Points(0), OfficeLength.Points(0)));
        var clipped = OfficeShape.Rectangle(10, 10); clipped.ClipPath = OfficeClipPath.Rectangle(10, 10);
        Assert.Throws<NotSupportedException>(() => document.AddShape(10, clipped, OfficeLength.Points(0), OfficeLength.Points(0)));
        var shadow = OfficeShape.Rectangle(10, 10); shadow.Shadow = new OfficeShadow(OfficeColor.Black, 1, 1, 1, 0);
        Assert.Throws<NotSupportedException>(() => document.AddShape(10, shadow, OfficeLength.Points(0), OfficeLength.Points(0)));
        var gradient = OfficeShape.Rectangle(10, 10); gradient.FillGradient = new OfficeLinearGradient(0, 0, 1, 0,
            new OfficeGradientStop(0, OfficeColor.Black), new OfficeGradientStop(1, OfficeColor.White));
        Assert.Throws<NotSupportedException>(() => document.AddShape(10, gradient, OfficeLength.Points(0), OfficeLength.Points(0)));
        var decorated = OfficeShape.Line(0, 0, 10, 10); decorated.StrokeEndMarker = new OfficeLineMarker(OfficeLineMarkerKind.Triangle, 3, 6);
        Assert.Throws<NotSupportedException>(() => document.AddShape(10, decorated, OfficeLength.Points(0), OfficeLength.Points(0)));
        Assert.Throws<NotSupportedException>(() => document.AddShape(10,
            OfficeShape.Polygon(new OfficePoint(0, 0), new OfficePoint(10, 0), new OfficePoint(0, 10)), OfficeLength.Points(0), OfficeLength.Points(0)));
        var rounded = OfficeShape.RoundedRectangle(10, 10, 2); rounded.CornerRadius = 6;
        Assert.Throws<ArgumentException>(() => document.AddShape(10, rounded, OfficeLength.Points(0), OfficeLength.Points(0)));
        var opacity = OfficeShape.Rectangle(10, 10); opacity.FillOpacity = double.NaN;
        Assert.Throws<ArgumentOutOfRangeException>(() => document.AddShape(10, opacity, OfficeLength.Points(0), OfficeLength.Points(0)));
        var line = OfficeShape.Line(0, 0, 10, 10); line.Width = 11;
        Assert.Throws<ArgumentException>(() => document.AddShape(10, line, OfficeLength.Points(0), OfficeLength.Points(0)));
        Assert.Throws<ArgumentException>(() => document.AddShape(10, new OfficeShape { Kind = OfficeShapeKind.Line }, OfficeLength.Points(0), OfficeLength.Points(0)));
        Assert.Same(pages, document.Pages);
    }

    [Fact]
    public void Solid_paint_accepts_transparent_and_disabled_paint_and_keeps_stroke_width_independent_of_shape_scale() {
        PublisherDocument document = PublisherShapeEditingFixture.Create();
        document.SetShapePaint(101, OfficeColor.White, OfficeColor.Black, OfficeLength.Points(3), 0, 1);
        OfficeShape transparent = document.Pages[0].Shapes.Single(shape => shape.Id == 101).CreateShape();
        Assert.Equal(0, transparent.FillOpacity); Assert.Equal(1, transparent.StrokeOpacity); Assert.Equal(3, transparent.StrokeWidth);
        document.SetShapeBounds(101, OfficeLength.Points(20), OfficeLength.Points(30), OfficeLength.Points(144), OfficeLength.Points(144));
        Assert.Equal(3, document.Pages[0].Shapes.Single(shape => shape.Id == 101).CreateShape().StrokeWidth);
        document.SetShapePaint(101, null, null, OfficeLength.Points(0), 1, 0);
        PublisherShape empty = document.Pages[0].Shapes.Single(shape => shape.Id == 101);
        Assert.True(empty.CanEdit); Assert.Null(empty.CreateShape().FillColor); Assert.Null(empty.CreateShape().StrokeColor);
        Assert.DoesNotContain(PublisherNativeTests.Elements(document.Pages[0].CreateDrawing()), element => element.SourceElementIds?.Contains("publisher-object-101") == true);
        Assert.Equal(document.ReadReport.CurrentObjectCount, document.ReadReport.ProjectedObjectCount);
    }

    [Fact]
    public void Current_object_and_identifier_limits_reject_creation_without_changing_existing_artwork() {
        const int limit = 128;
        PublisherDocument full = PublisherShapeEditingFixture.Create(hiddenObjects: limit - 3,
            options: new PublisherLoadOptions { Limits = new OfficeLegacyImportLimits { MaxItems = limit } });
        Assert.Equal(limit, full.ReadReport.CurrentObjectCount);
        var fullPages = full.Pages;
        var rectangle = OfficeShape.Rectangle(10, 10); rectangle.FillColor = OfficeColor.Black;
        Assert.Throws<InvalidDataException>(() => full.AddShape(10, rectangle, OfficeLength.Points(20), OfficeLength.Points(20))); Assert.Same(fullPages, full.Pages);
        PublisherDocument exhausted = PublisherShapeEditingFixture.Create(exhaustedIds: true);
        var exhaustedPages = exhausted.Pages; string before = exhausted.ToSvg();
        Assert.Throws<InvalidOperationException>(() => exhausted.AddShape(10, rectangle, OfficeLength.Points(20), OfficeLength.Points(20)));
        Assert.Same(exhaustedPages, exhausted.Pages); Assert.Equal(before, exhausted.ToSvg());
    }

    [Fact]
    public void Unqualified_master_registration_rejects_creation_and_edits_without_affecting_document_page_shapes() {
        PublisherDocument document = PublisherShapeEditingFixture.Create("unqualified-master");
        PublisherShape master = Assert.Single(document.MasterPages[0].Shapes);
        Assert.False(master.CanEdit); Assert.True(master.HasPlacement);
        var pages = document.Pages; var masters = document.MasterPages;
        Assert.Throws<NotSupportedException>(() => document.AddShape(20, OfficeShape.Rectangle(10, 10), OfficeLength.Points(10), OfficeLength.Points(10)));
        Assert.Throws<NotSupportedException>(() => document.SetShapeBounds(master.Id, OfficeLength.Points(10), OfficeLength.Points(10), OfficeLength.Points(30), OfficeLength.Points(30)));
        Assert.Same(pages, document.Pages); Assert.Same(masters, document.MasterPages);
        Assert.True(document.Pages[0].Shapes[0].CanEdit);
        Assert.Contains(document.ReadReport.Diagnostics, diagnostic => diagnostic.Code == "PUB_MASTER_COORDINATE_ORIGIN_UNASSESSED");
    }

    [Fact]
    public void Projection_failure_is_atomic_and_does_not_consume_the_next_shape_identity() {
        PublisherDocument document = PublisherShapeEditingFixture.Create(options: new PublisherLoadOptions {
            Limits = new OfficeLegacyImportLimits { MaxItems = 32 }
        });
        var primitive = OfficeShape.Rectangle(10, 10); primitive.FillColor = OfficeColor.Black;
        var successful = new List<PublisherShape>();
        for (int attempt = 0; attempt < 32; attempt++) {
            var pages = document.Pages; var report = document.ReadReport; string before = document.ToSvg();
            try { successful.Add(document.AddShape(20, primitive, OfficeLength.Points(200), OfficeLength.Points(200))); }
            catch (InvalidDataException exception) {
                Assert.Contains("projected drawing element limit", exception.Message);
                Assert.Same(pages, document.Pages); Assert.Same(report, document.ReadReport); Assert.Equal(before, document.ToSvg());
                Assert.NotEmpty(successful);
                uint expected = successful.Last().Id + 1;
                // Removing one repeated master paint releases enough projection
                // work for a page-local shape without deleting its model object.
                document.SetShapePaint(successful[0].Id, null, null, OfficeLength.Points(0));
                PublisherShape created = document.AddShape(10, primitive, OfficeLength.Points(200), OfficeLength.Points(200));
                Assert.Equal(expected, created.Id);
                return;
            }
        }
        Assert.Fail("Repeated master artwork must reach the bounded projection ceiling before the current object ceiling.");
    }
}
