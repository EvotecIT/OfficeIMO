using System.Collections.ObjectModel;
using System.Runtime.CompilerServices;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using OfficeIMO.PowerPoint.LegacyPpt.Model;
using A = DocumentFormat.OpenXml.Drawing;
using P = DocumentFormat.OpenXml.Presentation;

namespace OfficeIMO.PowerPoint.LegacyPpt.Write {
    internal static partial class LegacyPptWriter {
        private const ushort RecordExternalObjectList = 0x0409;
        private const ushort RecordExternalObjectListAtom = 0x040A;
        private const ushort RecordExternalHyperlinkAtom = 0x0FD3;
        private const ushort RecordExternalHyperlink = 0x0FD7;
        private const ushort RecordTextInteractiveInfoAtom = 0x0FDF;
        private const ushort RecordInteractiveInfo = 0x0FF2;
        private const ushort RecordInteractiveInfoAtom = 0x0FF3;

        internal static bool TryReadInteractions(PowerPointPresentation presentation,
            out LegacyPptWriterInteractionCatalog catalog, out string? reason) {
            if (presentation == null) throw new ArgumentNullException(nameof(presentation));
            return TryReadInteractions(presentation.Slides,
                new LegacyPptWriterSoundCatalog(), out catalog, out reason);
        }

        internal static bool TryReadInteractions(IEnumerable<PowerPointSlide> slides,
            out LegacyPptWriterInteractionCatalog catalog, out string? reason) {
            return TryReadInteractions(slides, new LegacyPptWriterSoundCatalog(),
                out catalog, out reason);
        }

        internal static bool TryReadInteractions(IEnumerable<PowerPointSlide> slides,
            LegacyPptWriterSoundCatalog soundCatalog,
            out LegacyPptWriterInteractionCatalog catalog, out string? reason) {
            catalog = new LegacyPptWriterInteractionCatalog(soundCatalog);
            reason = null;
            foreach (PowerPointSlide slide in slides) {
                foreach (PowerPointShape shape in slide.EnumerateShapesDeep(
                             slide.Shapes, includeHidden: true)) {
                    if (!TryReadShapeInteractions(slide.SlidePart, shape, catalog,
                            out LegacyPptWriterShapeInteractions interactions, out reason)) {
                        catalog = new LegacyPptWriterInteractionCatalog(soundCatalog);
                        return false;
                    }
                    if (interactions.HasInteractions) catalog.Add(shape.Element, interactions);
                }
            }
            return true;
        }

        private static bool TryReadShapeInteractions(SlidePart slidePart,
            PowerPointShape shape, LegacyPptWriterInteractionCatalog catalog,
            out LegacyPptWriterShapeInteractions interactions, out string? reason) {
            reason = null;
            var shapeActions = new List<LegacyPptWriterInteraction>();
            var textActions = new List<LegacyPptWriterTextInteraction>();
            P.NonVisualDrawingProperties? drawing = GetNonVisualDrawingProperties(shape.Element);
            if (drawing != null) {
                if (!TryGetSingleHyperlink<A.HyperlinkOnClick>(drawing,
                        out A.HyperlinkOnClick? clickElement, out reason)
                    || !TryGetSingleHyperlink<A.HyperlinkOnHover>(drawing,
                        out A.HyperlinkOnHover? hoverElement, out reason)) {
                    interactions = LegacyPptWriterShapeInteractions.Empty;
                    return false;
                }
                if (shape is PowerPointMedia
                    && IsDefaultMediaActivation(clickElement)) {
                    clickElement = null;
                }
                if (!TryReadHyperlink(slidePart, clickElement,
                        LegacyPptInteractionTrigger.MouseClick, catalog,
                        out LegacyPptWriterInteraction? click, out reason)
                    || !TryReadHyperlink(slidePart, hoverElement,
                        LegacyPptInteractionTrigger.MouseOver, catalog,
                        out LegacyPptWriterInteraction? hover, out reason)) {
                    interactions = LegacyPptWriterShapeInteractions.Empty;
                    return false;
                }
                if (click != null) shapeActions.Add(click);
                if (hover != null) shapeActions.Add(hover);
            }

            if (shape.Element is P.Shape textShape && textShape.TextBody != null
                && !TryReadTextInteractions(slidePart, textShape.TextBody, catalog,
                    textActions, out reason)) {
                interactions = LegacyPptWriterShapeInteractions.Empty;
                return false;
            }
            if (shape is PowerPointTable table) {
                for (int row = 0; row < table.Rows; row++) {
                    for (int column = 0; column < table.Columns; column++) {
                        PowerPointTableCell cell = table.GetCell(row, column);
                        var cellActions = new List<LegacyPptWriterTextInteraction>();
                        if (cell.Cell.TextBody != null
                            && !TryReadTextInteractions(slidePart,
                                cell.Cell.TextBody, catalog, cellActions,
                                out reason)) {
                            interactions = LegacyPptWriterShapeInteractions.Empty;
                            return false;
                        }
                        if (cellActions.Count > 0) {
                            catalog.Add(cell.Cell,
                                new LegacyPptWriterShapeInteractions(
                                    Array.Empty<LegacyPptWriterInteraction>(),
                                    cellActions));
                        }
                    }
                }
            }
            interactions = new LegacyPptWriterShapeInteractions(shapeActions, textActions);
            return true;
        }

        private static bool IsDefaultMediaActivation(
            A.HyperlinkOnClick? hyperlink) => hyperlink != null
            && string.Equals(hyperlink.Action?.Value, "ppaction://media",
                StringComparison.OrdinalIgnoreCase)
            && string.IsNullOrEmpty(hyperlink.Id?.Value)
            && string.IsNullOrEmpty(hyperlink.Tooltip?.Value)
            && hyperlink.HighlightClick?.Value != true
            && hyperlink.EndSound?.Value != true
            && !hyperlink.HasChildren;

        private static bool TryReadTextInteractions(SlidePart slidePart,
            OpenXmlCompositeElement textBody,
            LegacyPptWriterInteractionCatalog catalog,
            ICollection<LegacyPptWriterTextInteraction> result, out string? reason) {
            reason = null;
            A.Paragraph[] paragraphs = textBody.Elements<A.Paragraph>().ToArray();
            int position = 0;
            for (int paragraphIndex = 0; paragraphIndex < paragraphs.Length; paragraphIndex++) {
                foreach (OpenXmlElement child in paragraphs[paragraphIndex].ChildElements) {
                    string text = child switch {
                        A.Run run => run.Text?.Text ?? string.Empty,
                        A.Field => "*",
                        A.Break => "\v",
                        _ => string.Empty
                    };
                    if (child is A.Run textRun && text.Length > 0) {
                        A.RunProperties? properties = textRun.RunProperties;
                        if (!TryGetSingleHyperlink<A.HyperlinkOnClick>(properties,
                                out A.HyperlinkOnClick? clickElement, out reason)
                            || !TryGetSingleHyperlink<A.HyperlinkOnMouseOver>(properties,
                                out A.HyperlinkOnMouseOver? hoverElement, out reason)
                            || !TryReadHyperlink(slidePart, clickElement,
                                LegacyPptInteractionTrigger.MouseClick, catalog,
                                out LegacyPptWriterInteraction? click, out reason)
                            || !TryReadHyperlink(slidePart, hoverElement,
                                LegacyPptInteractionTrigger.MouseOver, catalog,
                                out LegacyPptWriterInteraction? hover, out reason)) {
                            return false;
                        }
                        if (click != null) {
                            result.Add(new LegacyPptWriterTextInteraction(position,
                                checked(position + text.Length), click));
                        }
                        if (hover != null) {
                            result.Add(new LegacyPptWriterTextInteraction(position,
                                checked(position + text.Length), hover));
                        }
                    } else if (child is A.Field field
                               && (field.RunProperties?.GetFirstChild<A.HyperlinkOnClick>() != null
                                   || field.RunProperties?.GetFirstChild<A.HyperlinkOnMouseOver>() != null)) {
                        reason = "Hyperlinks on DrawingML text fields are not encoded by the native binary writer.";
                        return false;
                    }
                    position = checked(position + text.Length);
                }
                if (paragraphIndex + 1 < paragraphs.Length) position++;
            }
            return true;
        }

        private static bool TryGetSingleHyperlink<T>(OpenXmlElement? parent,
            out T? hyperlink, out string? reason) where T : OpenXmlElement {
            hyperlink = null;
            reason = null;
            if (parent == null) return true;
            T[] candidates = parent.Elements<T>().ToArray();
            if (candidates.Length > 1) {
                reason = $"The {parent.LocalName} element contains multiple {candidates[0].LocalName} interactions.";
                return false;
            }
            hyperlink = candidates.SingleOrDefault();
            return true;
        }

        private static bool TryReadHyperlink(SlidePart slidePart,
            OpenXmlElement? hyperlink, LegacyPptInteractionTrigger trigger,
            LegacyPptWriterInteractionCatalog catalog,
            out LegacyPptWriterInteraction? interaction, out string? reason) {
            interaction = null;
            reason = null;
            if (hyperlink == null) return true;
            A.HyperlinkSound[] soundElements = hyperlink
                .Elements<A.HyperlinkSound>().ToArray();
            if (soundElements.Length > 1
                || hyperlink.ChildElements.Any(child => child is not A.HyperlinkSound)
                || hyperlink.GetAttributes().Any(attribute =>
                    !string.Equals(attribute.LocalName, "id", StringComparison.Ordinal)
                    && !string.Equals(attribute.LocalName, "action", StringComparison.Ordinal)
                    && !string.Equals(attribute.LocalName, "tooltip", StringComparison.Ordinal)
                    && !string.Equals(attribute.LocalName, "highlightClick", StringComparison.Ordinal)
                    && !string.Equals(attribute.LocalName, "endSnd", StringComparison.Ordinal))) {
                reason = "Hyperlink target frames, history, duplicate sound children, and extension data are not encoded.";
                return false;
            }
            string? relationshipId;
            string? action;
            string? screenTip;
            if (hyperlink is A.HyperlinkOnClick click) {
                relationshipId = click.Id?.Value;
                action = click.Action?.Value;
                screenTip = click.Tooltip?.Value;
            } else if (hyperlink is A.HyperlinkOnMouseOver hover) {
                relationshipId = hover.Id?.Value;
                action = hover.Action?.Value;
                screenTip = hover.Tooltip?.Value;
            } else if (hyperlink is A.HyperlinkOnHover shapeHover) {
                relationshipId = shapeHover.Id?.Value;
                action = shapeHover.Action?.Value;
                screenTip = shapeHover.Tooltip?.Value;
            } else {
                reason = $"Unsupported DrawingML interaction element {hyperlink.LocalName}.";
                return false;
            }
            if (screenTip?.Length == 0) screenTip = null;
            A.HyperlinkType typedHyperlink = (A.HyperlinkType)hyperlink;
            byte flags = 0;
            if (typedHyperlink.HighlightClick?.Value == true) flags |= 0x01;
            if (typedHyperlink.EndSound?.Value == true) flags |= 0x02;
            uint soundIdReference = 0;
            LegacyPptWriterSound? sound = null;
            if (soundElements.Length == 1
                && !catalog.Sounds.TryGetOrAdd(slidePart, soundElements[0],
                    out sound, out reason)) {
                return false;
            } else if (soundElements.Length == 1) {
                soundIdReference = sound!.Id;
            }

            if (TryParseCustomShowAction(action, out uint customShowId,
                    out bool returnsToSlide)) {
                if (!string.IsNullOrEmpty(relationshipId) || screenTip != null
                    || !TryResolveCustomShowName(slidePart, customShowId,
                        out string? customShowName)) {
                    reason = "Custom-show actions require a valid show id and cannot combine a relationship or screen tip.";
                    return false;
                }
                if (returnsToSlide) flags |= 0x04;
                interaction = new LegacyPptWriterInteraction(trigger,
                    LegacyPptInteractionAction.CustomShow,
                    LegacyPptInteractionJump.None, LegacyPptHyperlinkType.Nil,
                    hyperlinkIdReference: 0, name: customShowName,
                    soundIdReference: soundIdReference, flags: flags);
                return true;
            }

            const string MacroPrefix = "ppaction://macro?name=";
            if (action != null && action.StartsWith(MacroPrefix,
                    StringComparison.OrdinalIgnoreCase)) {
                string name = action.Substring(MacroPrefix.Length);
                if (!string.IsNullOrEmpty(relationshipId) || name.Length == 0
                    || screenTip != null) {
                    reason = "Macro actions require a nonempty name and cannot combine a relationship or screen tip.";
                    return false;
                }
                interaction = new LegacyPptWriterInteraction(trigger,
                    LegacyPptInteractionAction.Macro,
                    LegacyPptInteractionJump.None, LegacyPptHyperlinkType.Nil,
                    hyperlinkIdReference: 0, name: name,
                    soundIdReference: soundIdReference, flags: flags);
                return true;
            }

            if (!string.IsNullOrEmpty(relationshipId)) {
                if (slidePart.TryGetPartById(relationshipId!,
                        out OpenXmlPart? internalPart)) {
                    if (internalPart is not SlidePart targetSlidePart
                        || !string.Equals(action, "ppaction://hlinksldjump",
                            StringComparison.OrdinalIgnoreCase)
                        || !TryResolveInternalSlideTarget(targetSlidePart,
                            out LegacyPptWriterSlideTarget slideTarget)) {
                        reason = $"Internal hyperlink relationship '{relationshipId}' does not identify a supported slide target.";
                        return false;
                    }
                    LegacyPptWriterHyperlink internalTarget = catalog.GetOrAdd(
                        slideTarget, screenTip);
                    interaction = new LegacyPptWriterInteraction(trigger,
                        LegacyPptInteractionAction.Hyperlink,
                        LegacyPptInteractionJump.None,
                        LegacyPptHyperlinkType.SlideNumber, internalTarget.Id,
                        soundIdReference: soundIdReference, flags: flags);
                    return true;
                }
                if (string.Equals(action, "ppaction://program",
                        StringComparison.OrdinalIgnoreCase)) {
                    if (screenTip != null) {
                        reason = "Run-program actions cannot carry a binary screen tip.";
                        return false;
                    }
                    HyperlinkRelationship? programRelationship = slidePart
                        .HyperlinkRelationships.FirstOrDefault(candidate =>
                            string.Equals(candidate.Id, relationshipId,
                                StringComparison.Ordinal));
                    if (programRelationship == null
                        || !programRelationship.IsExternal) {
                        reason = $"Run-program relationship '{relationshipId}' is missing or is not external.";
                        return false;
                    }
                    interaction = new LegacyPptWriterInteraction(trigger,
                        LegacyPptInteractionAction.RunProgram,
                        LegacyPptInteractionJump.None,
                        LegacyPptHyperlinkType.Nil, hyperlinkIdReference: 0,
                        name: programRelationship.Uri.OriginalString,
                        soundIdReference: soundIdReference, flags: flags);
                    return true;
                }
                if (!string.IsNullOrEmpty(action)) {
                    reason = "Hyperlinks that combine a relationship target with a DrawingML action are not encoded yet.";
                    return false;
                }
                HyperlinkRelationship? relationship = slidePart.HyperlinkRelationships
                    .FirstOrDefault(candidate => string.Equals(candidate.Id,
                        relationshipId, StringComparison.Ordinal));
                if (relationship == null || !relationship.IsExternal) {
                    reason = $"Hyperlink relationship '{relationshipId}' is missing or is not external.";
                    return false;
                }
                LegacyPptWriterHyperlink target = catalog.GetOrAdd(relationship.Uri,
                    screenTip);
                interaction = new LegacyPptWriterInteraction(trigger,
                    LegacyPptInteractionAction.Hyperlink, LegacyPptInteractionJump.None,
                    MapHyperlinkType(relationship.Uri), target.Id,
                    soundIdReference: soundIdReference, flags: flags);
                return true;
            }

            if (string.IsNullOrEmpty(action)
                && (soundIdReference != 0 || (flags & 0x02) != 0)) {
                if (screenTip != null) {
                    reason = "A sound-only action cannot carry a binary screen tip.";
                    return false;
                }
                interaction = new LegacyPptWriterInteraction(trigger,
                    LegacyPptInteractionAction.None,
                    LegacyPptInteractionJump.None, LegacyPptHyperlinkType.Nil,
                    hyperlinkIdReference: 0,
                    soundIdReference: soundIdReference, flags: flags);
                return true;
            }
            if (!TryMapShowJump(action, out LegacyPptInteractionJump jump)) {
                reason = string.IsNullOrEmpty(action)
                    ? "A DrawingML hyperlink has neither a relationship target nor a supported action."
                    : $"DrawingML action '{action}' is not representable by the current binary action writer.";
                return false;
            }
            if (screenTip != null) {
                reason = "Screen tips on built-in slide-show jump actions have no binary hyperlink target.";
                return false;
            }
            interaction = new LegacyPptWriterInteraction(trigger,
                LegacyPptInteractionAction.Jump, jump, LegacyPptHyperlinkType.Nil,
                hyperlinkIdReference: 0,
                soundIdReference: soundIdReference, flags: flags);
            return true;
        }

        private static bool TryResolveInternalSlideTarget(SlidePart targetSlidePart,
            out LegacyPptWriterSlideTarget target) {
            target = default;
            PresentationPart? presentationPart = targetSlidePart.OpenXmlPackage.RootPart
                as PresentationPart;
            P.SlideId[] slideIds = presentationPart?.Presentation?.SlideIdList?
                .Elements<P.SlideId>().ToArray() ?? Array.Empty<P.SlideId>();
            for (int index = 0; index < slideIds.Length; index++) {
                string? relationshipId = slideIds[index].RelationshipId?.Value;
                if (string.IsNullOrEmpty(relationshipId)
                    || !presentationPart!.TryGetPartById(relationshipId!,
                        out OpenXmlPart? candidate)
                    || !ReferenceEquals(candidate, targetSlidePart)) continue;
                target = new LegacyPptWriterSlideTarget(
                    targetSlidePart.Uri.ToString(),
                    checked(unchecked((uint)index) + 256U),
                    checked(index + 1),
                    targetSlidePart.Slide?.CommonSlideData?.Name?.Value);
                return true;
            }
            return false;
        }

        private static bool TryParseCustomShowAction(string? action,
            out uint customShowId, out bool returnsToSlide) {
            const string Prefix = "ppaction://customshow?id=";
            const string ReturnSuffix = "&return=true";
            customShowId = 0;
            returnsToSlide = false;
            if (action == null || !action.StartsWith(Prefix,
                    StringComparison.OrdinalIgnoreCase)) return false;
            string value = action.Substring(Prefix.Length);
            if (value.EndsWith(ReturnSuffix, StringComparison.OrdinalIgnoreCase)) {
                returnsToSlide = true;
                value = value.Substring(0, value.Length - ReturnSuffix.Length);
            }
            return value.Length > 0 && value.IndexOf('&') < 0
                && uint.TryParse(value,
                    System.Globalization.NumberStyles.None,
                    System.Globalization.CultureInfo.InvariantCulture,
                    out customShowId);
        }

        private static bool TryResolveCustomShowName(SlidePart slidePart,
            uint customShowId, out string? name) {
            name = null;
            PresentationPart? presentationPart = slidePart.OpenXmlPackage.RootPart
                as PresentationPart;
            P.CustomShow[] matches = presentationPart?.Presentation?.CustomShowList?
                .Elements<P.CustomShow>().Where(show => show.Id?.Value == customShowId)
                .ToArray() ?? Array.Empty<P.CustomShow>();
            if (matches.Length != 1 || string.IsNullOrEmpty(
                    matches[0].Name?.Value)) return false;
            name = matches[0].Name!.Value;
            return true;
        }

        private static bool TryMapShowJump(string? action,
            out LegacyPptInteractionJump jump) {
            const string Prefix = "ppaction://hlinkshowjump?jump=";
            jump = LegacyPptInteractionJump.None;
            if (action == null || !action.StartsWith(Prefix,
                    StringComparison.OrdinalIgnoreCase)) return false;
            string value = action.Substring(Prefix.Length);
            jump = value.ToLowerInvariant() switch {
                "nextslide" => LegacyPptInteractionJump.NextSlide,
                "previousslide" => LegacyPptInteractionJump.PreviousSlide,
                "firstslide" => LegacyPptInteractionJump.FirstSlide,
                "lastslide" => LegacyPptInteractionJump.LastSlide,
                "lastslideviewed" => LegacyPptInteractionJump.LastViewedSlide,
                "endshow" => LegacyPptInteractionJump.EndShow,
                _ => LegacyPptInteractionJump.None
            };
            return jump != LegacyPptInteractionJump.None;
        }

        private static LegacyPptHyperlinkType MapHyperlinkType(Uri uri) {
            if (uri.IsAbsoluteUri && (uri.Scheme == Uri.UriSchemeHttp
                                     || uri.Scheme == Uri.UriSchemeHttps
                                     || uri.Scheme == Uri.UriSchemeMailto
                                     || uri.Scheme == Uri.UriSchemeFtp)) {
                return LegacyPptHyperlinkType.Url;
            }
            return LegacyPptHyperlinkType.OtherFile;
        }

        private static P.NonVisualDrawingProperties? GetNonVisualDrawingProperties(
            OpenXmlElement element) => element switch {
                P.Shape shape => shape.NonVisualShapeProperties?.NonVisualDrawingProperties,
                P.ConnectionShape connector => connector.NonVisualConnectionShapeProperties?
                    .NonVisualDrawingProperties,
                P.Picture picture => picture.NonVisualPictureProperties?
                    .NonVisualDrawingProperties,
                P.GroupShape group => group.NonVisualGroupShapeProperties?
                    .NonVisualDrawingProperties,
                P.GraphicFrame frame => frame.NonVisualGraphicFrameProperties?
                    .NonVisualDrawingProperties,
                _ => null
            };

        internal static byte[] BuildExternalObjectListRecord(
            LegacyPptWriterInteractionCatalog catalog,
            LegacyPptWriterMediaCatalog mediaCatalog,
            LegacyPptWriterOleObjectCatalog oleCatalog) {
            if (catalog == null) throw new ArgumentNullException(nameof(catalog));
            if (mediaCatalog == null) throw new ArgumentNullException(nameof(mediaCatalog));
            if (oleCatalog == null) throw new ArgumentNullException(nameof(oleCatalog));
            if (catalog.Hyperlinks.Count == 0
                && mediaCatalog.Media.Count == 0
                && oleCatalog.Objects.Count == 0) return Array.Empty<byte>();
            var listAtomPayload = new byte[4];
            uint seed = catalog.Hyperlinks.Select(link => link.Id)
                .Concat(mediaCatalog.Media.Select(media => media.Id))
                .Concat(oleCatalog.Objects.Select(ole => ole.Id)).Max();
            WriteUInt32(listAtomPayload, 0, seed);
            var children = new List<byte[]> {
                BuildRecord(version: 0, instance: 0, RecordExternalObjectListAtom,
                    listAtomPayload)
            };
            foreach (LegacyPptWriterHyperlink hyperlink in catalog.Hyperlinks) {
                children.Add(BuildExternalHyperlinkRecord(hyperlink));
            }
            foreach (LegacyPptWriterMedia media in mediaCatalog.Media) {
                children.Add(BuildExternalMediaRecord(media));
            }
            foreach (LegacyPptWriterOleObject ole in oleCatalog.Objects) {
                children.Add(BuildExternalOleObjectRecord(ole));
            }
            return BuildContainer(RecordExternalObjectList, instance: 0, children);
        }

        internal static byte[] BuildExternalHyperlinkRecord(
            LegacyPptWriterHyperlink hyperlink) {
            var atomPayload = new byte[4];
            WriteUInt32(atomPayload, 0, hyperlink.Id);
            ushort stringInstance;
            string value;
            if (hyperlink.IsInternalSlideTarget) {
                if (!hyperlink.TargetSlideId.HasValue
                    || !hyperlink.TargetSlideNumber.HasValue) {
                    throw new InvalidOperationException(
                        "An internal hyperlink has no binary slide destination.");
                }
                stringInstance = 3;
                value = hyperlink.TargetSlideId.Value + ","
                    + hyperlink.TargetSlideNumber.Value + ","
                    + (string.IsNullOrWhiteSpace(hyperlink.TargetSlideName)
                        ? " " : hyperlink.TargetSlideName);
            } else {
                stringInstance = 1;
                value = hyperlink.Target
                    ?? throw new InvalidOperationException(
                        "An external hyperlink has no target URI.");
            }
            return BuildContainer(RecordExternalHyperlink, instance: 0,
                new[] {
                    BuildRecord(version: 0, instance: 0,
                        RecordExternalHyperlinkAtom, atomPayload),
                    BuildRecord(version: 0, instance: stringInstance, RecordCString,
                        System.Text.Encoding.Unicode.GetBytes(value))
                });
        }

        internal static byte[] BuildInteractiveInfoRecord(
            LegacyPptWriterInteraction interaction) {
            var payload = new byte[16];
            WriteUInt32(payload, 0, interaction.SoundIdReference);
            WriteUInt32(payload, 4, interaction.HyperlinkIdReference);
            payload[8] = (byte)interaction.Action;
            payload[9] = interaction.OleVerb;
            payload[10] = (byte)interaction.Jump;
            payload[11] = interaction.Flags;
            payload[12] = (byte)interaction.HyperlinkType;
            byte[] atom = BuildRecord(version: 0, instance: 0,
                RecordInteractiveInfoAtom, payload);
            var children = new List<byte[]> { atom };
            if (!string.IsNullOrEmpty(interaction.Name)) {
                children.Add(BuildRecord(version: 0, instance: 0, RecordCString,
                    System.Text.Encoding.Unicode.GetBytes(interaction.Name)));
            }
            return BuildContainer(RecordInteractiveInfo,
                (ushort)interaction.Trigger, children);
        }

        internal static byte[] BuildTextInteractiveInfoRecord(
            LegacyPptWriterTextInteraction interaction) {
            var payload = new byte[8];
            WriteInt32(payload, 0, interaction.Begin);
            WriteInt32(payload, 4, interaction.End);
            return BuildRecord(version: 0,
                instance: (ushort)interaction.Interaction.Trigger,
                RecordTextInteractiveInfoAtom, payload);
        }

        internal sealed class LegacyPptWriterInteractionCatalog {
            private readonly Dictionary<OpenXmlElement, LegacyPptWriterShapeInteractions> _shapes =
                new(ReferenceComparer.Instance);
            private readonly Dictionary<string, LegacyPptWriterHyperlink> _hyperlinksByTarget =
                new(StringComparer.Ordinal);
            private readonly List<LegacyPptWriterHyperlink> _hyperlinks = new();

            internal LegacyPptWriterInteractionCatalog()
                : this(new LegacyPptWriterSoundCatalog()) { }

            internal LegacyPptWriterInteractionCatalog(
                LegacyPptWriterSoundCatalog sounds) {
                Sounds = sounds ?? throw new ArgumentNullException(nameof(sounds));
            }

            internal LegacyPptWriterSoundCatalog Sounds { get; }

            internal IReadOnlyList<LegacyPptWriterHyperlink> Hyperlinks =>
                new ReadOnlyCollection<LegacyPptWriterHyperlink>(_hyperlinks);

            internal bool HasInteractions => _shapes.Count > 0;

            internal LegacyPptWriterShapeInteractions Get(PowerPointShape shape) =>
                Get(shape.Element);

            internal LegacyPptWriterShapeInteractions Get(OpenXmlElement element) =>
                _shapes.TryGetValue(element, out LegacyPptWriterShapeInteractions? value)
                    ? value
                    : LegacyPptWriterShapeInteractions.Empty;

            internal LegacyPptWriterHyperlink? FindHyperlink(uint id) =>
                _hyperlinks.FirstOrDefault(link => link.Id == id);

            internal LegacyPptWriterHyperlink GetOrAdd(Uri uri, string? screenTip) {
                string target = uri.OriginalString;
                string key = CreateHyperlinkKey(target, screenTip);
                if (_hyperlinksByTarget.TryGetValue(key,
                        out LegacyPptWriterHyperlink? existing)) return existing;
                var created = new LegacyPptWriterHyperlink(
                    checked((uint)_hyperlinks.Count + 1U), target, screenTip);
                _hyperlinks.Add(created);
                _hyperlinksByTarget.Add(key, created);
                return created;
            }

            internal LegacyPptWriterHyperlink GetOrAdd(
                LegacyPptWriterSlideTarget target, string? screenTip) {
                string key = CreateInternalHyperlinkKey(target.PartUri, screenTip);
                if (_hyperlinksByTarget.TryGetValue(key,
                        out LegacyPptWriterHyperlink? existing)) return existing;
                var created = new LegacyPptWriterHyperlink(
                    checked((uint)_hyperlinks.Count + 1U), target.PartUri,
                    target.BinarySlideId, target.SlideNumber, target.Name, screenTip);
                _hyperlinks.Add(created);
                _hyperlinksByTarget.Add(key, created);
                return created;
            }

            internal void Add(OpenXmlElement shape,
                LegacyPptWriterShapeInteractions interactions) =>
                _shapes.Add(shape, interactions);
        }

        internal sealed class LegacyPptWriterShapeInteractions {
            internal static LegacyPptWriterShapeInteractions Empty { get; } = new(
                Array.Empty<LegacyPptWriterInteraction>(),
                Array.Empty<LegacyPptWriterTextInteraction>());

            internal LegacyPptWriterShapeInteractions(
                IReadOnlyList<LegacyPptWriterInteraction> shapeInteractions,
                IReadOnlyList<LegacyPptWriterTextInteraction> textInteractions) {
                ShapeInteractions = shapeInteractions;
                TextInteractions = textInteractions;
            }

            internal IReadOnlyList<LegacyPptWriterInteraction> ShapeInteractions { get; }
            internal IReadOnlyList<LegacyPptWriterTextInteraction> TextInteractions { get; }
            internal bool HasInteractions => ShapeInteractions.Count > 0
                || TextInteractions.Count > 0;
        }

        internal sealed class LegacyPptWriterInteraction {
            internal LegacyPptWriterInteraction(LegacyPptInteractionTrigger trigger,
                LegacyPptInteractionAction action, LegacyPptInteractionJump jump,
                LegacyPptHyperlinkType hyperlinkType, uint hyperlinkIdReference,
                string? name = null, uint soundIdReference = 0,
                byte oleVerb = 0, byte flags = 0) {
                Trigger = trigger;
                Action = action;
                Jump = jump;
                HyperlinkType = hyperlinkType;
                HyperlinkIdReference = hyperlinkIdReference;
                Name = name;
                SoundIdReference = soundIdReference;
                OleVerb = oleVerb;
                Flags = flags;
            }

            internal LegacyPptInteractionTrigger Trigger { get; }
            internal LegacyPptInteractionAction Action { get; }
            internal LegacyPptInteractionJump Jump { get; }
            internal LegacyPptHyperlinkType HyperlinkType { get; }
            internal uint HyperlinkIdReference { get; }
            internal string? Name { get; }
            internal uint SoundIdReference { get; }
            internal byte OleVerb { get; }
            internal byte Flags { get; }
        }

        internal sealed class LegacyPptWriterTextInteraction {
            internal LegacyPptWriterTextInteraction(int begin, int end,
                LegacyPptWriterInteraction interaction) {
                Begin = begin;
                End = end;
                Interaction = interaction;
            }

            internal int Begin { get; }
            internal int End { get; }
            internal LegacyPptWriterInteraction Interaction { get; }
        }

        internal sealed class LegacyPptWriterHyperlink {
            internal LegacyPptWriterHyperlink(uint id, string target,
                string? screenTip = null, uint extensionFlags = 0) {
                Id = id;
                Target = target;
                ScreenTip = screenTip;
                ExtensionFlags = extensionFlags;
            }

            internal LegacyPptWriterHyperlink(uint id, string targetSlidePartUri,
                uint targetSlideId, int targetSlideNumber, string? targetSlideName,
                string? screenTip = null, uint extensionFlags = 0) {
                if (targetSlideNumber <= 0) {
                    throw new ArgumentOutOfRangeException(nameof(targetSlideNumber));
                }
                Id = id;
                TargetSlidePartUri = targetSlidePartUri
                    ?? throw new ArgumentNullException(nameof(targetSlidePartUri));
                TargetSlideId = targetSlideId;
                TargetSlideNumber = targetSlideNumber;
                TargetSlideName = targetSlideName;
                ScreenTip = screenTip;
                ExtensionFlags = extensionFlags;
            }

            internal uint Id { get; }
            internal string? Target { get; }
            internal string? TargetSlidePartUri { get; }
            internal uint? TargetSlideId { get; }
            internal int? TargetSlideNumber { get; }
            internal string? TargetSlideName { get; }
            internal bool IsInternalSlideTarget => TargetSlidePartUri != null;
            internal string? ScreenTip { get; }
            internal uint ExtensionFlags { get; }
        }

        internal static string CreateHyperlinkKey(string target, string? screenTip) =>
            "external\0" + target + "\0"
            + (screenTip == null ? "-" : screenTip.Length + ":" + screenTip);

        internal static string CreateInternalHyperlinkKey(string targetSlidePartUri,
            string? screenTip) => "slide\0" + targetSlidePartUri + "\0"
            + (screenTip == null ? "-" : screenTip.Length + ":" + screenTip);

        internal readonly struct LegacyPptWriterSlideTarget {
            internal LegacyPptWriterSlideTarget(string partUri, uint binarySlideId,
                int slideNumber, string? name) {
                PartUri = partUri ?? throw new ArgumentNullException(nameof(partUri));
                BinarySlideId = binarySlideId;
                SlideNumber = slideNumber;
                Name = name;
            }

            internal string PartUri { get; }
            internal uint BinarySlideId { get; }
            internal int SlideNumber { get; }
            internal string? Name { get; }
        }

        private sealed class ReferenceComparer : IEqualityComparer<OpenXmlElement> {
            internal static ReferenceComparer Instance { get; } = new();
            public bool Equals(OpenXmlElement? x, OpenXmlElement? y) => ReferenceEquals(x, y);
            public int GetHashCode(OpenXmlElement obj) => RuntimeHelpers.GetHashCode(obj);
        }
    }
}
