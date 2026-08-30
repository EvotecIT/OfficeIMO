using DocumentFormat.OpenXml.Packaging;
using OfficeIMO.PowerPoint;
using Xunit;
using A = DocumentFormat.OpenXml.Drawing;
using P = DocumentFormat.OpenXml.Presentation;

namespace OfficeIMO.Tests;

public sealed class PowerPointShapeHyperlinkTests {
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
