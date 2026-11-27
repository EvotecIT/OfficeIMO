using OfficeIMO.OpenXml.Internal;
using OfficeIMO.PowerPoint.LegacyPpt.Internal;
using OfficeIMO.PowerPoint.LegacyPpt.Model;
using OfficeIMO.PowerPoint.OpenXml;

namespace OfficeIMO.PowerPoint {
    public sealed partial class PowerPointDocument {
        private static void ApplyLegacyShapeInteractions(PowerPointPartRef ownerPart,
            OfficeXmlNode target, LegacyPptShape source,
            IReadOnlyDictionary<uint, PowerPointPartRef>? slidePartsByLegacyId = null,
            LegacyPptSoundProjectionContext? soundContext = null,
            ICollection<LegacyPptDeferredProjection>?
                deferredInteractions = null) {
            OfficeXmlNode properties = GetLegacyNonVisualDrawingProperties(target);
            if (properties.IsNone) return;
            foreach (LegacyPptInteraction interaction in source.Interactions) {
                if (slidePartsByLegacyId == null
                    && ShouldDeferLegacyInteraction(interaction)
                    && deferredInteractions != null) {
                    // The shape is found again by id: the element read now may belong to a package tree the SDK
                    // replaced by the time the deferred interaction applies.
                    uint shapeId = properties.UInt("id") ?? 0U;
                    deferredInteractions.Add(new LegacyPptDeferredProjection(
                        projectedSlides => {
                            OfficeXmlNode shape = FindLegacyShapeById(ownerPart.Root, shapeId);
                            OfficeXmlNode current = GetLegacyNonVisualDrawingProperties(shape);
                            if (current.IsNone) return;
                            ApplyLegacyShapeInteraction(ownerPart, current,
                                interaction, projectedSlides, soundContext);
                            ownerPart.MarkEdited();
                        }));
                    continue;
                }
                ApplyLegacyShapeInteraction(ownerPart, properties, interaction,
                    slidePartsByLegacyId, soundContext);
            }
        }

        /// <summary>The <c>p:cNvPr</c> of a projected shape, picture, connector, group or graphic frame.</summary>
        private static OfficeXmlNode GetLegacyNonVisualDrawingProperties(OfficeXmlNode shape) {
            if (shape.IsNone || shape.NamespaceId != LegacyPptXml.P) return default;
            string? nonVisual = shape.LocalName switch {
                "sp" => "nvSpPr",
                "cxnSp" => "nvCxnSpPr",
                "pic" => "nvPicPr",
                "grpSp" => "nvGrpSpPr",
                "graphicFrame" => "nvGraphicFramePr",
                _ => null
            };
            return nonVisual == null ? default : shape.Child(LegacyPptXml.P, nonVisual).Child(LegacyPptXml.P, "cNvPr");
        }

        private static bool ShouldDeferLegacyInteraction(
            LegacyPptInteraction interaction) =>
            interaction.Action == LegacyPptInteractionAction.CustomShow
            || interaction.Action == LegacyPptInteractionAction.Hyperlink
            && interaction.HyperlinkType ==
                LegacyPptHyperlinkType.SlideNumber;

        private static bool ShouldDeferLegacyTextInteractions(
            LegacyPptTextBody textBody,
            IReadOnlyDictionary<uint, PowerPointPartRef>? slidePartsByLegacyId,
            ICollection<LegacyPptDeferredProjection>?
                deferredInteractions) =>
            slidePartsByLegacyId == null && deferredInteractions != null
            && textBody.Interactions.Any(item =>
                ShouldDeferLegacyInteraction(item.Interaction));

        private static void ApplyLegacyShapeInteraction(PowerPointPartRef ownerPart,
            OfficeXmlNode properties,
            LegacyPptInteraction interaction,
            IReadOnlyDictionary<uint, PowerPointPartRef>? slidePartsByLegacyId,
            LegacyPptSoundProjectionContext? soundContext) {
            ProjectLegacyInteraction(ownerPart, interaction, properties,
                shapeLevel: true, slidePartsByLegacyId: slidePartsByLegacyId,
                soundContext: soundContext, replaceExisting: true);
        }

        /// <summary>
        /// A projection step that waits for the projected slides, such as a master hyperlink to a slide or a custom show
        /// that the masters reference before the slides exist.
        /// </summary>
        internal sealed partial class LegacyPptDeferredProjection {
            private readonly Action<IReadOnlyDictionary<uint, PowerPointPartRef>> _apply;

            internal LegacyPptDeferredProjection(
                Action<IReadOnlyDictionary<uint, PowerPointPartRef>> apply) {
                _apply = apply ?? throw new ArgumentNullException(nameof(apply));
            }

            internal void Apply(
                IReadOnlyDictionary<uint, PowerPointPartRef> slidePartsByLegacyId) {
                _apply(slidePartsByLegacyId);
            }
        }

        /// <summary>The interaction projector for text: hyperlinks go into the run properties.</summary>
        private static LegacyPptTextProjection.InteractionProjector CreateLegacyTextInteractionProjector(
            PowerPointPartRef ownerPart,
            IReadOnlyDictionary<uint, PowerPointPartRef>? slidePartsByLegacyId,
            LegacyPptSoundProjectionContext? soundContext) =>
            (interaction, runProperties) => ProjectLegacyInteraction(ownerPart, interaction, runProperties,
                slidePartsByLegacyId: slidePartsByLegacyId, soundContext: soundContext);

        /// <summary>Text runs whose interactions wait for the slides: the run properties stay without hyperlinks.</summary>
        private static readonly LegacyPptTextProjection.InteractionProjector SuppressedLegacyTextInteractions =
            (_, _) => { };

        /// <summary>
        /// Appends the hyperlink element of a binary interaction to <paramref name="container"/> (run properties, or the
        /// non-visual drawing properties of a shape with <paramref name="shapeLevel"/>), adding the relationships it
        /// needs. <paramref name="replaceExisting"/> first removes an element of the same kind.
        /// </summary>
        internal static void ProjectLegacyInteraction(
            PowerPointPartRef ownerPart, LegacyPptInteraction interaction,
            OfficeXmlNode container,
            bool shapeLevel = false,
            IReadOnlyDictionary<uint, PowerPointPartRef>? slidePartsByLegacyId = null,
            LegacyPptSoundProjectionContext? soundContext = null,
            bool replaceExisting = false) {
            string? customShowName = interaction.Name;
            if (interaction.Action == LegacyPptInteractionAction.None
                && (interaction.SoundIdReference != 0 || interaction.StopsSound)) {
                AppendLegacyInteractionElement(container, interaction, shapeLevel,
                    string.Empty, action: null, ownerPart: ownerPart,
                    soundContext: soundContext, replaceExisting: replaceExisting);
                return;
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
                AppendLegacyInteractionElement(container, interaction, shapeLevel,
                    string.Empty, customShowAction, ownerPart: ownerPart,
                    soundContext: soundContext, replaceExisting: replaceExisting);
                return;
            }
            if (interaction.Action == LegacyPptInteractionAction.Macro
                && !string.IsNullOrEmpty(interaction.Name)) {
                AppendLegacyInteractionElement(container, interaction, shapeLevel,
                    string.Empty, "ppaction://macro?name=" + interaction.Name,
                    ownerPart: ownerPart, soundContext: soundContext, replaceExisting: replaceExisting);
                return;
            }
            if (interaction.Action == LegacyPptInteractionAction.RunProgram
                && !string.IsNullOrEmpty(interaction.Name)
                && Uri.TryCreate(interaction.Name, UriKind.RelativeOrAbsolute,
                    out Uri? programUri)) {
                string programRelationship = ownerPart.AddExternalHyperlink(programUri);
                AppendLegacyInteractionElement(container, interaction, shapeLevel,
                    programRelationship, "ppaction://program",
                    ownerPart: ownerPart, soundContext: soundContext, replaceExisting: replaceExisting);
                return;
            }
            if (interaction.Action == LegacyPptInteractionAction.Hyperlink
                && interaction.HyperlinkType == LegacyPptHyperlinkType.SlideNumber
                && interaction.Hyperlink?.TargetSlideId is uint targetSlideId) {
                if (slidePartsByLegacyId == null
                    || !slidePartsByLegacyId.TryGetValue(targetSlideId,
                        out PowerPointPartRef targetSlidePart)) {
                    throw new InvalidDataException(
                        $"Internal hyperlink target slide {targetSlideId} cannot be resolved in the projected presentation.");
                }
                string relationshipId = ownerPart.GetOrAddSlideRelationship(targetSlidePart);
                AppendLegacyInteractionElement(container, interaction, shapeLevel,
                    relationshipId, "ppaction://hlinksldjump",
                    interaction.Hyperlink.ScreenTip, ownerPart, soundContext, replaceExisting);
                return;
            }
            if (interaction.Action == LegacyPptInteractionAction.Hyperlink
                && interaction.HyperlinkType != LegacyPptHyperlinkType.SlideNumber
                && interaction.HyperlinkType != LegacyPptHyperlinkType.CustomShow
                && interaction.Hyperlink?.Uri is Uri uri) {
                string relationship = ownerPart.AddExternalHyperlink(uri);
                AppendLegacyInteractionElement(container, interaction, shapeLevel,
                    relationship, action: null,
                    interaction.Hyperlink.ScreenTip, ownerPart, soundContext, replaceExisting);
                return;
            }

            string? action = GetLegacyPowerPointAction(interaction);
            if (action == null) return;
            AppendLegacyInteractionElement(container, interaction, shapeLevel,
                string.Empty, action, ownerPart: ownerPart,
                soundContext: soundContext, replaceExisting: replaceExisting);
        }

        private static void AppendLegacyInteractionElement(OfficeXmlNode container,
            LegacyPptInteraction interaction, bool shapeLevel,
            string relationshipId, string? action, string? tooltip = null,
            PowerPointPartRef ownerPart = default,
            LegacyPptSoundProjectionContext? soundContext = null,
            bool replaceExisting = false) {
            string name = interaction.Trigger ==
                LegacyPptInteractionTrigger.MouseOver
                ? shapeLevel ? "hlinkHover" : "hlinkMouseOver"
                : "hlinkClick";
            if (replaceExisting) container.RemoveChildren(LegacyPptXml.A, name);
            OfficeXmlNode hyperlink = container.AppendNew(LegacyPptXml.A, name)
                .SetRelationship("id", relationshipId)
                .Set("action", action)
                .Set("tooltip", tooltip);
            if (interaction.IsAnimated) hyperlink.Set("highlightClick", true);
            if (interaction.StopsSound) hyperlink.Set("endSnd", true);
            if (interaction.SoundIdReference != 0 && !ownerPart.IsNone
                && soundContext?.TryProject(ownerPart,
                    interaction.SoundIdReference, out LegacyPptSound? sound,
                    out string? soundRelationshipId) == true) {
                hyperlink.AddA("snd")
                    .SetRelationship("embed", soundRelationshipId)
                    .Set("name", sound!.Name)
                    .Set("builtIn", sound.BuiltInId.HasValue);
            }
        }

        private static bool TryResolveLegacyCustomShowId(PowerPointPartRef ownerPart,
            string name, out uint customShowId) {
            customShowId = 0;
            PowerPointDocument? document = ownerPart.Document;
            if (document == null) return false;
            LegacyPptImportParts parts = LegacyPptImportParts.For(document);
            Dictionary<string, uint?>? index = parts.CustomShowIds;
            if (index == null) {
                OfficeXmlNode list = parts.PresentationPart.Root.Child(LegacyPptXml.P, "custShowLst");
                if (list.IsNone) return false;
                index = CreateLegacyCustomShowIdIndex(list);
                parts.CustomShowIds = index;
            }
            if (!index.TryGetValue(name, out uint? value) || !value.HasValue) return false;
            customShowId = value.Value;
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
