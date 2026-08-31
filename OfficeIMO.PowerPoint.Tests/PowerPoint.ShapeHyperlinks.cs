using DocumentFormat.OpenXml.Packaging;
using OfficeIMO.PowerPoint;
using Xunit;
using A = DocumentFormat.OpenXml.Drawing;
using P = DocumentFormat.OpenXml.Presentation;

namespace OfficeIMO.Tests;

public sealed class PowerPointShapeHyperlinkTests {
    [Fact]
    public void Shape_internal_slide_hyperlink_can_be_assigned_from_its_getter() {
        using PowerPointPresentation presentation = PowerPointPresentation.Create();
        PowerPointSlide source = presentation.AddSlide();
        PowerPointTextBox shape = source.AddTextBox("Open target");
        PowerPointSlide target = presentation.AddSlide();

        shape.SetHyperlink(target, "Target slide");
        Uri fragment = Assert.IsType<Uri>(shape.Hyperlink);
        shape.Hyperlink = fragment;

        Assert.Equal("#slide-2", shape.Hyperlink!.OriginalString);
        Assert.Empty(presentation.ValidateDocument());
    }

    [Fact]
    public void Text_run_internal_slide_hyperlink_can_be_assigned_from_its_getter() {
        using PowerPointPresentation presentation = PowerPointPresentation.Create();
        PowerPointSlide source = presentation.AddSlide();
        PowerPointTextRun run = source.AddTextBox("Open target")
            .Paragraphs.Single().Runs.Single();
        PowerPointSlide target = presentation.AddSlide();

        run.SetHyperlink(target, "Target slide");
        Uri fragment = Assert.IsType<Uri>(run.Hyperlink);
        run.Hyperlink = fragment;

        Assert.Equal("#slide-2", run.Hyperlink!.OriginalString);
        Assert.Empty(presentation.ValidateDocument());
    }

    [Fact]
    public void Replacing_or_clearing_click_link_preserves_relationship_used_by_hover_link() {
        using PowerPointPresentation presentation = PowerPointPresentation.Create();
        PowerPointSlide slide = presentation.AddSlide();
        PowerPointTextBox shape = slide.AddTextBox("Linked");
        shape.SetHyperlink(new Uri("https://example.test/shared"));
        P.NonVisualDrawingProperties properties = ((P.Shape)shape.Element)
            .NonVisualShapeProperties!.NonVisualDrawingProperties!;
        A.HyperlinkOnClick click = properties.GetFirstChild<A.HyperlinkOnClick>()!;
        string sharedRelationshipId = click.Id!.Value!;
        properties.Append(new A.HyperlinkOnHover { Id = sharedRelationshipId });

        shape.SetHyperlink(new Uri("https://example.test/replacement"));

        Assert.Contains(slide.SlidePart.HyperlinkRelationships,
            relationship => relationship.Id == sharedRelationshipId);
        Assert.Equal(sharedRelationshipId,
            properties.GetFirstChild<A.HyperlinkOnHover>()!.Id!.Value);
        string replacementRelationshipId = properties
            .GetFirstChild<A.HyperlinkOnClick>()!.Id!.Value!;
        Assert.NotEqual(sharedRelationshipId, replacementRelationshipId);

        shape.ClearHyperlink();

        Assert.Contains(slide.SlidePart.HyperlinkRelationships,
            relationship => relationship.Id == sharedRelationshipId);
        Assert.DoesNotContain(slide.SlidePart.HyperlinkRelationships,
            relationship => relationship.Id == replacementRelationshipId);
        Assert.Empty(presentation.ValidateDocument());
    }
}
