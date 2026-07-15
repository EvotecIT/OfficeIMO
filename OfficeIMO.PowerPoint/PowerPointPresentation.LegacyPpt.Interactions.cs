using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Presentation;
using OfficeIMO.PowerPoint.LegacyPpt.Model;
using A = DocumentFormat.OpenXml.Drawing;

namespace OfficeIMO.PowerPoint {
    public sealed partial class PowerPointPresentation {
        private static void ApplyLegacyShapeInteractions(OpenXmlPart ownerPart,
            OpenXmlElement target, LegacyPptShape source,
            IReadOnlyDictionary<uint, SlidePart>? slidePartsByLegacyId = null) {
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
                             interaction, shapeLevel: true,
                             slidePartsByLegacyId: slidePartsByLegacyId)) {
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
            bool shapeLevel = false,
            IReadOnlyDictionary<uint, SlidePart>? slidePartsByLegacyId = null) {
            if (interaction.Action == LegacyPptInteractionAction.Macro
                && !string.IsNullOrEmpty(interaction.Name)) {
                return CreateLegacyInteractionElements(interaction, shapeLevel,
                    string.Empty, "ppaction://macro?name=" + interaction.Name);
            }
            if (interaction.Action == LegacyPptInteractionAction.RunProgram
                && !string.IsNullOrEmpty(interaction.Name)
                && Uri.TryCreate(interaction.Name, UriKind.RelativeOrAbsolute,
                    out Uri? programUri)) {
                HyperlinkRelationship programRelationship = ownerPart
                    .AddHyperlinkRelationship(programUri, isExternal: true);
                return CreateLegacyInteractionElements(interaction, shapeLevel,
                    programRelationship.Id, "ppaction://program");
            }
            if (interaction.Action == LegacyPptInteractionAction.Hyperlink
                && interaction.HyperlinkType == LegacyPptHyperlinkType.SlideNumber
                && interaction.Hyperlink?.TargetSlideId is uint targetSlideId) {
                if (slidePartsByLegacyId == null
                    || !slidePartsByLegacyId.TryGetValue(targetSlideId,
                        out SlidePart? targetSlidePart)) {
                    throw new InvalidDataException(
                        $"Internal hyperlink target slide {targetSlideId} cannot be resolved in the projected presentation.");
                }
                if (!ownerPart.Parts.Any(pair => ReferenceEquals(
                        pair.OpenXmlPart, targetSlidePart))) {
                    ownerPart.AddPart(targetSlidePart);
                }
                string relationshipId = ownerPart.GetIdOfPart(targetSlidePart);
                return CreateLegacyInteractionElements(interaction, shapeLevel,
                    relationshipId, "ppaction://hlinksldjump",
                    interaction.Hyperlink.ScreenTip);
            }
            if (interaction.Action == LegacyPptInteractionAction.Hyperlink
                && interaction.HyperlinkType != LegacyPptHyperlinkType.SlideNumber
                && interaction.Hyperlink?.Uri is Uri uri) {
                HyperlinkRelationship relationship = ownerPart.AddHyperlinkRelationship(uri,
                    isExternal: true);
                return CreateLegacyInteractionElements(interaction, shapeLevel,
                    relationship.Id, action: null,
                    interaction.Hyperlink.ScreenTip);
            }

            string? action = GetLegacyPowerPointAction(interaction);
            if (action == null) return Array.Empty<OpenXmlElement>();
            return CreateLegacyInteractionElements(interaction, shapeLevel,
                string.Empty, action);
        }

        private static IReadOnlyList<OpenXmlElement> CreateLegacyInteractionElements(
            LegacyPptInteraction interaction, bool shapeLevel,
            string relationshipId, string? action, string? tooltip = null) {
            A.HyperlinkType hyperlink = interaction.Trigger ==
                LegacyPptInteractionTrigger.MouseOver
                ? shapeLevel
                    ? new A.HyperlinkOnHover()
                    : new A.HyperlinkOnMouseOver()
                : new A.HyperlinkOnClick();
            hyperlink.Id = relationshipId;
            hyperlink.Action = action;
            hyperlink.Tooltip = tooltip;
            if (interaction.IsAnimated) hyperlink.HighlightClick = true;
            if (interaction.StopsSound) hyperlink.EndSound = true;
            return new OpenXmlElement[] { hyperlink };
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
