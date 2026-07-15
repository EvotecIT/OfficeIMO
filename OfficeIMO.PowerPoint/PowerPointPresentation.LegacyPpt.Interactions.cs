using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Presentation;
using OfficeIMO.PowerPoint.LegacyPpt.Model;
using A = DocumentFormat.OpenXml.Drawing;

namespace OfficeIMO.PowerPoint {
    public sealed partial class PowerPointPresentation {
        private static void ApplyLegacyShapeInteractions(OpenXmlPart ownerPart,
            OpenXmlElement target, LegacyPptShape source) {
            NonVisualDrawingProperties? properties = target switch {
                Shape shape => shape.NonVisualShapeProperties?.NonVisualDrawingProperties,
                ConnectionShape connector => connector.NonVisualConnectionShapeProperties?
                    .NonVisualDrawingProperties,
                Picture picture => picture.NonVisualPictureProperties?.NonVisualDrawingProperties,
                GroupShape group => group.NonVisualGroupShapeProperties?
                    .NonVisualDrawingProperties,
                _ => null
            };
            if (properties == null) return;
            foreach (LegacyPptInteraction interaction in source.Interactions) {
                foreach (OpenXmlElement element in ProjectLegacyInteraction(ownerPart,
                             interaction, shapeLevel: true)) {
                    if (element is A.HyperlinkOnClick) {
                        properties.RemoveAllChildren<A.HyperlinkOnClick>();
                    } else if (element is A.HyperlinkOnHover) {
                        properties.RemoveAllChildren<A.HyperlinkOnHover>();
                    }
                    properties.Append(element);
                }
            }
        }

        private static IReadOnlyList<OpenXmlElement> ProjectLegacyInteraction(
            OpenXmlPart ownerPart, LegacyPptInteraction interaction,
            bool shapeLevel = false) {
            if (interaction.Action == LegacyPptInteractionAction.Hyperlink
                && interaction.Hyperlink?.Uri is Uri uri) {
                HyperlinkRelationship relationship = ownerPart.AddHyperlinkRelationship(uri,
                    isExternal: true);
                return interaction.Trigger == LegacyPptInteractionTrigger.MouseOver
                    ? new OpenXmlElement[] { shapeLevel
                        ? new A.HyperlinkOnHover { Id = relationship.Id }
                        : new A.HyperlinkOnMouseOver { Id = relationship.Id } }
                    : new OpenXmlElement[] { new A.HyperlinkOnClick { Id = relationship.Id } };
            }

            string? action = GetLegacyPowerPointAction(interaction);
            if (action == null) return Array.Empty<OpenXmlElement>();
            return interaction.Trigger == LegacyPptInteractionTrigger.MouseOver
                ? new OpenXmlElement[] { shapeLevel
                    ? new A.HyperlinkOnHover { Id = string.Empty, Action = action }
                    : new A.HyperlinkOnMouseOver { Id = string.Empty, Action = action } }
                : new OpenXmlElement[] {
                    new A.HyperlinkOnClick { Id = string.Empty, Action = action }
                };
        }

        private static string? GetLegacyPowerPointAction(
            LegacyPptInteraction interaction) {
            LegacyPptInteractionJump jump = interaction.Action switch {
                LegacyPptInteractionAction.Jump => interaction.Jump,
                LegacyPptInteractionAction.Hyperlink => interaction.HyperlinkType switch {
                    LegacyPptHyperlinkType.NextSlide => LegacyPptInteractionJump.NextSlide,
                    LegacyPptHyperlinkType.PreviousSlide => LegacyPptInteractionJump.PreviousSlide,
                    LegacyPptHyperlinkType.FirstSlide => LegacyPptInteractionJump.FirstSlide,
                    LegacyPptHyperlinkType.LastSlide => LegacyPptInteractionJump.LastSlide,
                    _ => LegacyPptInteractionJump.None
                },
                _ => LegacyPptInteractionJump.None
            };
            string? value = jump switch {
                LegacyPptInteractionJump.NextSlide => "nextslide",
                LegacyPptInteractionJump.PreviousSlide => "previousslide",
                LegacyPptInteractionJump.FirstSlide => "firstslide",
                LegacyPptInteractionJump.LastSlide => "lastslide",
                LegacyPptInteractionJump.LastViewedSlide => "lastslideviewed",
                LegacyPptInteractionJump.EndShow => "endshow",
                _ => null
            };
            return value == null ? null : "ppaction://hlinkshowjump?jump=" + value;
        }
    }
}
