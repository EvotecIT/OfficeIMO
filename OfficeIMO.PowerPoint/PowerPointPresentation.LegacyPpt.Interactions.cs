using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Presentation;
using OfficeIMO.PowerPoint.LegacyPpt.Model;
using A = DocumentFormat.OpenXml.Drawing;
using P = DocumentFormat.OpenXml.Presentation;

namespace OfficeIMO.PowerPoint {
    public sealed partial class PowerPointPresentation {
        private static void ApplyLegacyShapeInteractions(OpenXmlPart ownerPart,
            OpenXmlElement target, LegacyPptShape source,
            IReadOnlyDictionary<uint, SlidePart>? slidePartsByLegacyId = null,
            LegacyPptSoundProjectionContext? soundContext = null,
            ICollection<LegacyPptDeferredProjection>?
                deferredInteractions = null) {
            NonVisualDrawingProperties? properties = target switch {
                Shape shape => shape.NonVisualShapeProperties?.NonVisualDrawingProperties,
                ConnectionShape connector => connector.NonVisualConnectionShapeProperties?
                    .NonVisualDrawingProperties,
                Picture picture => picture.NonVisualPictureProperties?.NonVisualDrawingProperties,
                GroupShape group => group.NonVisualGroupShapeProperties?
                    .NonVisualDrawingProperties,
                GraphicFrame frame => frame.NonVisualGraphicFrameProperties?
                    .NonVisualDrawingProperties,
                _ => null
            };
            if (properties == null) return;
            foreach (LegacyPptInteraction interaction in source.Interactions) {
                if (slidePartsByLegacyId == null
                    && ShouldDeferLegacyInteraction(interaction)
                    && deferredInteractions != null) {
                    deferredInteractions.Add(new LegacyPptDeferredProjection(
                        projectedSlides => {
                            ApplyLegacyShapeInteraction(ownerPart, properties,
                                interaction, projectedSlides, soundContext);
                            ownerPart.RootElement?.Save();
                        }));
                    continue;
                }
                ApplyLegacyShapeInteraction(ownerPart, properties, interaction,
                    slidePartsByLegacyId, soundContext);
            }
        }

        private static bool ShouldDeferLegacyInteraction(
            LegacyPptInteraction interaction) =>
            interaction.Action == LegacyPptInteractionAction.CustomShow
            || interaction.Action == LegacyPptInteractionAction.Hyperlink
            && interaction.HyperlinkType ==
                LegacyPptHyperlinkType.SlideNumber;

        private static bool ShouldDeferLegacyTextInteractions(
            LegacyPptTextBody textBody,
            IReadOnlyDictionary<uint, SlidePart>? slidePartsByLegacyId,
            ICollection<LegacyPptDeferredProjection>?
                deferredInteractions) =>
            slidePartsByLegacyId == null && deferredInteractions != null
            && textBody.Interactions.Any(item =>
                ShouldDeferLegacyInteraction(item.Interaction));

        private static void ApplyLegacyShapeInteraction(OpenXmlPart ownerPart,
            NonVisualDrawingProperties properties,
            LegacyPptInteraction interaction,
            IReadOnlyDictionary<uint, SlidePart>? slidePartsByLegacyId,
            LegacyPptSoundProjectionContext? soundContext) {
            foreach (OpenXmlElement element in ProjectLegacyInteraction(ownerPart,
                         interaction, shapeLevel: true,
                         slidePartsByLegacyId: slidePartsByLegacyId,
                         soundContext: soundContext)) {
                if (element is A.HyperlinkOnClick) {
                    properties.RemoveAllChildren<A.HyperlinkOnClick>();
                } else if (element is A.HyperlinkOnHover) {
                    properties.RemoveAllChildren<A.HyperlinkOnHover>();
                }
                properties.Append(element);
            }
        }

        internal sealed class LegacyPptDeferredProjection {
            private readonly Action<IReadOnlyDictionary<uint, SlidePart>>
                _apply;

            internal LegacyPptDeferredProjection(
                Action<IReadOnlyDictionary<uint, SlidePart>> apply) {
                _apply = apply ?? throw new ArgumentNullException(nameof(apply));
            }

            internal void Apply(
                IReadOnlyDictionary<uint, SlidePart> slidePartsByLegacyId) {
                _apply(slidePartsByLegacyId);
            }
        }

        private static IReadOnlyList<OpenXmlElement> ProjectLegacyInteraction(
            OpenXmlPart ownerPart, LegacyPptInteraction interaction,
            bool shapeLevel = false,
            IReadOnlyDictionary<uint, SlidePart>? slidePartsByLegacyId = null,
            LegacyPptSoundProjectionContext? soundContext = null) {
            string? customShowName = interaction.Name;
            if (interaction.Action == LegacyPptInteractionAction.None
                && (interaction.SoundIdReference != 0 || interaction.StopsSound)) {
                return CreateLegacyInteractionElements(interaction, shapeLevel,
                    string.Empty, action: null, ownerPart: ownerPart,
                    soundContext: soundContext);
            }
            if (interaction.Action == LegacyPptInteractionAction.CustomShow
                && customShowName != null && customShowName.Length > 0
                && TryResolveLegacyCustomShowId(ownerPart, customShowName,
                    out uint customShowId)) {
                string customShowAction = "ppaction://customshow?id="
                    + customShowId.ToString(
                        System.Globalization.CultureInfo.InvariantCulture);
                if (interaction.ReturnsFromCustomShow) {
                    customShowAction += "&return=true";
                }
                return CreateLegacyInteractionElements(interaction, shapeLevel,
                    string.Empty, customShowAction, ownerPart: ownerPart,
                    soundContext: soundContext);
            }
            if (interaction.Action == LegacyPptInteractionAction.Macro
                && !string.IsNullOrEmpty(interaction.Name)) {
                return CreateLegacyInteractionElements(interaction, shapeLevel,
                    string.Empty, "ppaction://macro?name=" + interaction.Name,
                    ownerPart: ownerPart, soundContext: soundContext);
            }
            if (interaction.Action == LegacyPptInteractionAction.RunProgram
                && !string.IsNullOrEmpty(interaction.Name)
                && Uri.TryCreate(interaction.Name, UriKind.RelativeOrAbsolute,
                    out Uri? programUri)) {
                HyperlinkRelationship programRelationship = ownerPart
                    .AddHyperlinkRelationship(programUri, isExternal: true);
                return CreateLegacyInteractionElements(interaction, shapeLevel,
                    programRelationship.Id, "ppaction://program",
                    ownerPart: ownerPart, soundContext: soundContext);
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
                    interaction.Hyperlink.ScreenTip, ownerPart, soundContext);
            }
            if (interaction.Action == LegacyPptInteractionAction.Hyperlink
                && interaction.HyperlinkType != LegacyPptHyperlinkType.SlideNumber
                && interaction.HyperlinkType != LegacyPptHyperlinkType.CustomShow
                && interaction.Hyperlink?.Uri is Uri uri) {
                HyperlinkRelationship relationship = ownerPart.AddHyperlinkRelationship(uri,
                    isExternal: true);
                return CreateLegacyInteractionElements(interaction, shapeLevel,
                    relationship.Id, action: null,
                    interaction.Hyperlink.ScreenTip, ownerPart, soundContext);
            }

            string? action = GetLegacyPowerPointAction(interaction);
            if (action == null) return Array.Empty<OpenXmlElement>();
            return CreateLegacyInteractionElements(interaction, shapeLevel,
                string.Empty, action, ownerPart: ownerPart,
                soundContext: soundContext);
        }

        private static IReadOnlyList<OpenXmlElement> CreateLegacyInteractionElements(
            LegacyPptInteraction interaction, bool shapeLevel,
            string relationshipId, string? action, string? tooltip = null,
            OpenXmlPart? ownerPart = null,
            LegacyPptSoundProjectionContext? soundContext = null) {
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
            if (interaction.SoundIdReference != 0 && ownerPart != null
                && soundContext?.TryProject(ownerPart,
                    interaction.SoundIdReference, out LegacyPptSound? sound,
                    out string? soundRelationshipId) == true) {
                hyperlink.Append(new A.HyperlinkSound {
                    Embed = soundRelationshipId,
                    Name = sound!.Name,
                    BuiltIn = sound.BuiltInId.HasValue
                });
            }
            return new OpenXmlElement[] { hyperlink };
        }

        private static bool TryResolveLegacyCustomShowId(OpenXmlPart ownerPart,
            string name, out uint customShowId) {
            customShowId = 0;
            PresentationPart? presentationPart = ownerPart.OpenXmlPackage.RootPart
                as PresentationPart;
            P.CustomShow[] matches = presentationPart?.Presentation?.CustomShowList?
                .Elements<P.CustomShow>().Where(show => string.Equals(
                    show.Name?.Value, name, StringComparison.Ordinal)).ToArray()
                ?? Array.Empty<P.CustomShow>();
            uint? id = matches.Length == 1 ? matches[0].Id?.Value : null;
            if (!id.HasValue) {
                return false;
            }
            customShowId = id.Value;
            return true;
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
