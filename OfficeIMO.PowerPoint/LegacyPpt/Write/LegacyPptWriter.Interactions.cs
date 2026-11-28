using System.Collections.ObjectModel;
using OfficeIMO.PowerPoint.LegacyPpt.Internal;
using OfficeIMO.PowerPoint.LegacyPpt.Model;
using static OfficeIMO.PowerPoint.LegacyPpt.Write.LegacyPptWriterXml;

namespace OfficeIMO.PowerPoint.LegacyPpt.Write {
    internal static partial class LegacyPptWriter {
        private const ushort RecordExternalObjectList = 0x0409;
        private const ushort RecordExternalObjectListAtom = 0x040A;
        private const ushort RecordExternalHyperlinkAtom = 0x0FD3;
        private const ushort RecordExternalHyperlink = 0x0FD7;
        private const ushort RecordTextInteractiveInfoAtom = 0x0FDF;
        private const ushort RecordInteractiveInfo = 0x0FF2;
        private const ushort RecordInteractiveInfoAtom = 0x0FF3;

        // Caller without a package view: LegacyPptWriter (single-slide records).
        internal static bool TryReadInteractions(IEnumerable<PowerPointSlide> slides,
            out LegacyPptWriterInteractionCatalog catalog, out string? reason) {
            return TryReadInteractions(slides, new LegacyPptWriterSoundCatalog(),
                out catalog, out reason);
        }

        // Caller without a package view: LegacyPptPreservingWriter.ChangePlan.
        internal static bool TryReadInteractions(IEnumerable<PowerPointSlide> slides,
            LegacyPptWriterSoundCatalog soundCatalog,
            out LegacyPptWriterInteractionCatalog catalog, out string? reason) {
            ArgumentNullException.ThrowIfNull(slides);
            PowerPointSlide[] list = slides.ToArray();
            if (list.Length == 0) {
                catalog = new LegacyPptWriterInteractionCatalog(soundCatalog);
                reason = null;
                return true;
            }
            using LegacyPptPackageView view = CaptureView(list[0]);
            return TryReadInteractions(view, list, soundCatalog, out catalog,
                out reason);
        }

        internal static bool TryReadInteractions(LegacyPptPackageView view,
            IEnumerable<PowerPointSlide> slides,
            LegacyPptWriterSoundCatalog soundCatalog,
            out LegacyPptWriterInteractionCatalog catalog, out string? reason) {
            ArgumentNullException.ThrowIfNull(view);
            catalog = new LegacyPptWriterInteractionCatalog(soundCatalog);
            reason = null;
            foreach (PowerPointSlide slide in slides) {
                LegacyPptPartView slidePart = view.Part(slide.PartRef.Name);
                foreach (PowerPointShape shape in slide.EnumerateShapesDeep(
                             slide.Shapes, includeHidden: true)) {
                    if (!TryReadShapeInteractions(view, slidePart, shape, catalog,
                            out LegacyPptWriterShapeInteractions interactions, out reason)) {
                        catalog = new LegacyPptWriterInteractionCatalog(soundCatalog);
                        return false;
                    }
                    if (interactions.HasInteractions) catalog.Add(shape, interactions);
                }
            }
            return true;
        }

        // The shape's element in the view, or its own markup when the view has nothing at its place.
        private static OfficeXmlNode ResolveShape(LegacyPptPackageView view,
            PowerPointShape shape) {
            OfficeXmlNode resolved = view.Resolve(shape.OwnerPartRef.Name,
                shape.XmlElement);
            return resolved.IsNone ? ShapeXml(shape) : resolved;
        }

        private static bool TryReadShapeInteractions(LegacyPptPackageView view,
            LegacyPptPartView slidePart,
            PowerPointShape shape, LegacyPptWriterInteractionCatalog catalog,
            out LegacyPptWriterShapeInteractions interactions, out string? reason) {
            reason = null;
            var shapeActions = new List<LegacyPptWriterInteraction>();
            var textActions = new List<LegacyPptWriterTextInteraction>();
            OfficeXmlNode element = ResolveShape(view, shape);
            OfficeXmlNode drawing = GetNonVisualDrawingProperties(element);
            if (!drawing.IsNone) {
                if (!TryGetSingleHyperlink(drawing, "hlinkClick",
                        out OfficeXmlNode clickElement, out reason)
                    || !TryGetSingleHyperlink(drawing, "hlinkHover",
                        out OfficeXmlNode hoverElement, out reason)) {
                    interactions = LegacyPptWriterShapeInteractions.Empty;
                    return false;
                }
                if (shape is PowerPointMedia
                    && IsDefaultMediaActivation(clickElement)) {
                    clickElement = default;
                }
                if (!TryReadHyperlink(view, slidePart, clickElement,
                        LegacyPptInteractionTrigger.MouseClick, catalog,
                        out LegacyPptWriterInteraction? click, out reason)
                    || !TryReadHyperlink(view, slidePart, hoverElement,
                        LegacyPptInteractionTrigger.MouseOver, catalog,
                        out LegacyPptWriterInteraction? hover, out reason)) {
                    interactions = LegacyPptWriterShapeInteractions.Empty;
                    return false;
                }
                if (click != null) shapeActions.Add(click);
                if (hover != null) shapeActions.Add(hover);
            }

            OfficeXmlNode textBody = element.Is(LegacyPptXml.P, "sp")
                ? element.Child(LegacyPptXml.P, "txBody")
                : default;
            if (!textBody.IsNone
                && !TryReadTextInteractions(view, slidePart, textBody, catalog,
                    textActions, out reason)) {
                interactions = LegacyPptWriterShapeInteractions.Empty;
                return false;
            }
            if (shape is PowerPointTable table) {
                OfficeXmlNode tableElement = element
                    .Child(LegacyPptXml.A, "graphic")
                    .Child(LegacyPptXml.A, "graphicData")
                    .Child(LegacyPptXml.A, "tbl");
                List<OfficeXmlNode> rows = tableElement
                    .Children(LegacyPptXml.A, "tr").ToList();
                for (int row = 0; row < table.RowCount; row++) {
                    List<OfficeXmlNode> cells = row < rows.Count
                        ? rows[row].Children(LegacyPptXml.A, "tc").ToList()
                        : new List<OfficeXmlNode>();
                    for (int column = 0; column < table.ColumnCount; column++) {
                        // The table reads its cells by row and column position; this keeps its checks.
                        _ = table.GetCell(row, column);
                        OfficeXmlNode cellTextBody = column < cells.Count
                            ? cells[column].Child(LegacyPptXml.A, "txBody")
                            : default;
                        var cellActions = new List<LegacyPptWriterTextInteraction>();
                        if (!cellTextBody.IsNone
                            && !TryReadTextInteractions(view, slidePart,
                                cellTextBody, catalog, cellActions,
                                out reason)) {
                            interactions = LegacyPptWriterShapeInteractions.Empty;
                            return false;
                        }
                        if (cellActions.Count > 0) {
                            catalog.Add(table, row, column,
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
            OfficeXmlNode hyperlink) => !hyperlink.IsNone
            && string.Equals(hyperlink.Attribute("action"), "ppaction://media",
                StringComparison.OrdinalIgnoreCase)
            && string.IsNullOrEmpty(hyperlink.Attribute(LegacyPptXml.R, "id"))
            && string.IsNullOrEmpty(hyperlink.Attribute("tooltip"))
            && hyperlink.Bool("highlightClick") != true
            && hyperlink.Bool("endSnd") != true
            && !HasChildren(hyperlink);

        private static bool TryReadTextInteractions(LegacyPptPackageView view,
            LegacyPptPartView slidePart,
            OfficeXmlNode textBody,
            LegacyPptWriterInteractionCatalog catalog,
            ICollection<LegacyPptWriterTextInteraction> result, out string? reason) {
            reason = null;
            List<OfficeXmlNode> paragraphs = textBody
                .Children(LegacyPptXml.A, "p").ToList();
            int position = 0;
            for (int paragraphIndex = 0; paragraphIndex < paragraphs.Count; paragraphIndex++) {
                foreach (OfficeXmlNode child in paragraphs[paragraphIndex].Children()) {
                    bool isRun = child.Is(LegacyPptXml.A, "r");
                    bool isField = child.Is(LegacyPptXml.A, "fld");
                    string text = isRun
                        ? child.Child(LegacyPptXml.A, "t").InnerText
                        : isField ? "*"
                        : child.Is(LegacyPptXml.A, "br") ? "\v"
                        : string.Empty;
                    if (isRun && text.Length > 0) {
                        OfficeXmlNode properties = child.Child(LegacyPptXml.A, "rPr");
                        if (!TryGetSingleHyperlink(properties, "hlinkClick",
                                out OfficeXmlNode clickElement, out reason)
                            || !TryGetSingleHyperlink(properties, "hlinkMouseOver",
                                out OfficeXmlNode hoverElement, out reason)
                            || !TryReadHyperlink(view, slidePart, clickElement,
                                LegacyPptInteractionTrigger.MouseClick, catalog,
                                out LegacyPptWriterInteraction? click, out reason)
                            || !TryReadHyperlink(view, slidePart, hoverElement,
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
                    } else if (isField) {
                        OfficeXmlNode fieldProperties = child.Child(LegacyPptXml.A, "rPr");
                        if (!fieldProperties.Child(LegacyPptXml.A, "hlinkClick").IsNone
                            || !fieldProperties.Child(LegacyPptXml.A, "hlinkMouseOver").IsNone) {
                            reason = "Hyperlinks on DrawingML text fields are not encoded by the native binary writer.";
                            return false;
                        }
                    }
                    position = checked(position + text.Length);
                }
                if (paragraphIndex + 1 < paragraphs.Count) position++;
            }
            return true;
        }

        private static bool TryGetSingleHyperlink(OfficeXmlNode parent,
            string localName, out OfficeXmlNode hyperlink, out string? reason) {
            hyperlink = default;
            reason = null;
            if (parent.IsNone) return true;
            List<OfficeXmlNode> candidates = parent
                .Children(LegacyPptXml.A, localName).ToList();
            if (candidates.Count > 1) {
                reason = $"The {parent.LocalName} element contains multiple {candidates[0].LocalName} interactions.";
                return false;
            }
            hyperlink = candidates.Count == 1 ? candidates[0] : default;
            return true;
        }

        private static readonly string[] HyperlinkAttributes = {
            "id", "action", "tooltip", "highlightClick", "endSnd"
        };

        private static bool TryReadHyperlink(LegacyPptPackageView view,
            LegacyPptPartView slidePart,
            OfficeXmlNode hyperlink, LegacyPptInteractionTrigger trigger,
            LegacyPptWriterInteractionCatalog catalog,
            out LegacyPptWriterInteraction? interaction, out string? reason) {
            interaction = null;
            reason = null;
            if (hyperlink.IsNone) return true;
            if (hyperlink.Children(LegacyPptXml.A, "snd").Count > 1
                || AnyChildOtherThan(hyperlink, LegacyPptXml.A, "snd")
                || HasAttributeOtherThan(hyperlink, HyperlinkAttributes)) {
                reason = "Hyperlink target frames, history, duplicate sound children, and extension data are not encoded.";
                return false;
            }
            if (!IsHyperlinkElement(hyperlink)) {
                reason = $"Unsupported DrawingML interaction element {hyperlink.LocalName}.";
                return false;
            }
            string? relationshipId = hyperlink.Attribute(LegacyPptXml.R, "id");
            string? action = hyperlink.Attribute("action");
            string? screenTip = hyperlink.Attribute("tooltip");
            if (screenTip?.Length == 0) screenTip = null;
            byte flags = 0;
            if (hyperlink.Bool("highlightClick") == true) flags |= 0x01;
            if (hyperlink.Bool("endSnd") == true) flags |= 0x02;

            if (PowerPointCustomShowAction.IsCustomShowAction(action)) {
                if (!TryValidateCustomShowHyperlink(slidePart,
                        hyperlink, catalog.Sounds,
                        out uint customShowId, out bool returnsToSlide,
                        out LegacyPptWriterSound? customShowSound, out reason)
                    || !TryResolveCustomShowName(view, customShowId,
                        out string? customShowName)) {
                    reason ??= "Custom-show actions require a valid show id.";
                    return false;
                }
                if (returnsToSlide) flags |= 0x04;
                interaction = new LegacyPptWriterInteraction(trigger,
                    LegacyPptInteractionAction.CustomShow,
                    LegacyPptInteractionJump.None, LegacyPptHyperlinkType.Nil,
                    hyperlinkIdReference: 0, name: customShowName,
                    soundIdReference: customShowSound?.Id ?? 0U, flags: flags);
                return true;
            }

            if (!TryRegisterHyperlinkSound(slidePart, hyperlink,
                    catalog.Sounds, out LegacyPptWriterSound? sound,
                    out reason)) {
                return false;
            }
            uint soundIdReference = sound?.Id ?? 0U;

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
                if (TryGetRelatedPart(slidePart, relationshipId,
                        out LegacyPptRelationshipView internalRelationship)) {
                    if (!internalRelationship.IsType(
                            LegacyPptPackageView.SlideRelationshipType)
                        || !string.Equals(action, "ppaction://hlinksldjump",
                            StringComparison.OrdinalIgnoreCase)
                        || !TryResolveInternalSlideTarget(view,
                            internalRelationship.TargetPart,
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
                    if (!TryGetHyperlinkRelationship(slidePart, relationshipId!,
                            out LegacyPptRelationshipView programRelationship)
                        || !programRelationship.IsExternal) {
                        reason = $"Run-program relationship '{relationshipId}' is missing or is not external.";
                        return false;
                    }
                    interaction = new LegacyPptWriterInteraction(trigger,
                        LegacyPptInteractionAction.RunProgram,
                        LegacyPptInteractionJump.None,
                        LegacyPptHyperlinkType.Nil, hyperlinkIdReference: 0,
                        name: programRelationship.Target,
                        soundIdReference: soundIdReference, flags: flags);
                    return true;
                }
                if (!string.IsNullOrEmpty(action)) {
                    reason = "Hyperlinks that combine a relationship target with a DrawingML action are not encoded yet.";
                    return false;
                }
                if (!TryGetHyperlinkRelationship(slidePart, relationshipId!,
                        out LegacyPptRelationshipView relationship)
                    || !relationship.IsExternal) {
                    reason = $"Hyperlink relationship '{relationshipId}' is missing or is not external.";
                    return false;
                }
                LegacyPptWriterHyperlink target = catalog.GetOrAdd(
                    relationship.Target, screenTip);
                interaction = new LegacyPptWriterInteraction(trigger,
                    LegacyPptInteractionAction.Hyperlink, LegacyPptInteractionJump.None,
                    MapHyperlinkType(relationship.Target), target.Id,
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

        // The DrawingML hyperlink elements: a:hlinkClick, a:hlinkHover and a:hlinkMouseOver.
        private static bool IsHyperlinkElement(OfficeXmlNode element) =>
            element.Is(LegacyPptXml.A, "hlinkClick")
            || element.Is(LegacyPptXml.A, "hlinkHover")
            || element.Is(LegacyPptXml.A, "hlinkMouseOver");

        /// <summary>
        /// The relationship of a related package part: an internal relationship that is neither a hyperlink nor an audio,
        /// video or media reference, whose target part exists.
        /// </summary>
        private static bool TryGetRelatedPart(LegacyPptPartView owner,
            string? relationshipId, out LegacyPptRelationshipView relationship) =>
            owner.TryGetRelationship(relationshipId, out relationship)
            && !relationship.IsExternal
            && !relationship.TargetPart.IsNone
            && !relationship.IsType(LegacyPptPackageView.HyperlinkRelationshipType)
            && !relationship.IsType(LegacyPptPackageView.AudioRelationshipType)
            && !relationship.IsType(LegacyPptPackageView.VideoRelationshipType)
            && !relationship.IsType(LegacyPptPackageView.MediaRelationshipType);

        private static bool TryGetHyperlinkRelationship(LegacyPptPartView owner,
            string relationshipId, out LegacyPptRelationshipView relationship) =>
            owner.TryGetRelationship(relationshipId, out relationship)
            && relationship.IsType(LegacyPptPackageView.HyperlinkRelationshipType);

        private static bool TryResolveInternalSlideTarget(
            LegacyPptPackageView view, LegacyPptPartView targetSlidePart,
            out LegacyPptWriterSlideTarget target) {
            target = default;
            LegacyPptPartView presentationPart = view.Presentation;
            List<OfficeXmlNode> slideIds = presentationPart.Root
                .Child(LegacyPptXml.P, "sldIdLst")
                .Children(LegacyPptXml.P, "sldId").ToList();
            for (int index = 0; index < slideIds.Count; index++) {
                string? relationshipId = slideIds[index].Attribute(LegacyPptXml.R, "id");
                if (string.IsNullOrEmpty(relationshipId)
                    || !TryGetRelatedPart(presentationPart, relationshipId,
                        out LegacyPptRelationshipView candidate)
                    || !candidate.TargetPart.Equals(targetSlidePart)) continue;
                target = new LegacyPptWriterSlideTarget(
                    targetSlidePart.Name,
                    checked(unchecked((uint)index) + 256U),
                    checked(index + 1),
                    targetSlidePart.Root.Child(LegacyPptXml.P, "cSld")
                        .Attribute("name"));
                return true;
            }
            return false;
        }

        internal static bool TryValidateCustomShowHyperlink(
            LegacyPptPartView ownerPart, OfficeXmlNode hyperlink,
            LegacyPptWriterSoundCatalog soundCatalog,
            out uint customShowId, out bool returnsToSlide,
            out LegacyPptWriterSound? sound, out string? reason) {
            if (ownerPart.IsNone) throw new ArgumentNullException(nameof(ownerPart));
            ArgumentNullException.ThrowIfNull(soundCatalog);
            sound = null;
            if (!TryValidateSupportedCustomShowHyperlink(
                    hyperlink, out customShowId, out returnsToSlide,
                    out reason)) {
                return false;
            }
            return TryRegisterHyperlinkSound(ownerPart, hyperlink,
                soundCatalog, out sound, out reason);
        }

        // The checks PowerPointCustomShowAction.TryValidateSupportedHyperlink makes, on the element's markup.
        private static bool TryValidateSupportedCustomShowHyperlink(
            OfficeXmlNode hyperlink, out uint customShowId,
            out bool returnsToSlide, out string? reason) {
            if (hyperlink.IsNone) throw new ArgumentNullException(nameof(hyperlink));
            reason = null;
            if (!PowerPointCustomShowAction.TryParseSupported(
                    hyperlink.Attribute("action"),
                    out customShowId, out returnsToSlide)) {
                reason = "The custom-show action is malformed.";
                return false;
            }
            if (!string.IsNullOrEmpty(hyperlink.Attribute(LegacyPptXml.R, "id"))
                || !string.IsNullOrEmpty(hyperlink.Attribute("tooltip"))) {
                reason = "Custom-show actions cannot combine a relationship or screen tip.";
                return false;
            }
            if (hyperlink.Children(LegacyPptXml.A, "snd").Count > 1
                || AnyChildOtherThan(hyperlink, LegacyPptXml.A, "snd")
                || HasAttributeOtherThan(hyperlink, HyperlinkAttributes)) {
                reason = "Custom-show actions contain unsupported hyperlink attributes or child markup.";
                return false;
            }
            return true;
        }

        internal static bool TryRegisterHyperlinkSound(
            LegacyPptPartView ownerPart, OfficeXmlNode hyperlink,
            LegacyPptWriterSoundCatalog soundCatalog,
            out LegacyPptWriterSound? sound, out string? reason) {
            if (ownerPart.IsNone) throw new ArgumentNullException(nameof(ownerPart));
            if (hyperlink.IsNone) throw new ArgumentNullException(nameof(hyperlink));
            ArgumentNullException.ThrowIfNull(soundCatalog);
            sound = null;
            reason = null;
            List<OfficeXmlNode> sounds = hyperlink
                .Children(LegacyPptXml.A, "snd").ToList();
            if (sounds.Count == 0) return true;
            if (sounds.Count > 1) {
                reason = "A DrawingML interaction cannot contain multiple action sounds.";
                return false;
            }
            return soundCatalog.TryGetOrAdd(ownerPart, sounds[0],
                out sound, out reason);
        }

        private static bool TryResolveCustomShowName(LegacyPptPackageView view,
            uint customShowId, out string? name) {
            name = null;
            List<OfficeXmlNode> matches = view.Presentation.Root
                .Child(LegacyPptXml.P, "custShowLst")
                .Children(LegacyPptXml.P, "custShow")
                .ToList()
                .Where(show => show.UInt("id") == customShowId)
                .ToList();
            if (matches.Count != 1 || string.IsNullOrEmpty(
                    matches[0].Attribute("name"))) return false;
            name = matches[0].Attribute("name");
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

        private static LegacyPptHyperlinkType MapHyperlinkType(string target) {
            if (Uri.TryCreate(target, UriKind.RelativeOrAbsolute, out Uri? uri)
                && uri.IsAbsoluteUri && (uri.Scheme == Uri.UriSchemeHttp
                                     || uri.Scheme == Uri.UriSchemeHttps
                                     || uri.Scheme == Uri.UriSchemeMailto
                                     || uri.Scheme == Uri.UriSchemeFtp)) {
                return LegacyPptHyperlinkType.Url;
            }
            return LegacyPptHyperlinkType.OtherFile;
        }

        private static OfficeXmlNode GetNonVisualDrawingProperties(
            OfficeXmlNode element) {
            if (element.NamespaceId != LegacyPptXml.P) return default;
            string? container = element.LocalName switch {
                "sp" => "nvSpPr",
                "cxnSp" => "nvCxnSpPr",
                "pic" => "nvPicPr",
                "grpSp" => "nvGrpSpPr",
                "graphicFrame" => "nvGraphicFramePr",
                _ => null
            };
            return container == null
                ? default
                : element.Child(LegacyPptXml.P, container)
                    .Child(LegacyPptXml.P, "cNvPr");
        }

        internal static byte[] BuildExternalObjectListRecord(
            LegacyPptWriterInteractionCatalog catalog,
            LegacyPptWriterMediaCatalog mediaCatalog,
            LegacyPptWriterOleObjectCatalog oleCatalog) {
            ArgumentNullException.ThrowIfNull(catalog);
            ArgumentNullException.ThrowIfNull(mediaCatalog);
            ArgumentNullException.ThrowIfNull(oleCatalog);
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
            // Keyed by the shape's write key, or by its table and cell position for a table cell.
            private readonly Dictionary<object, LegacyPptWriterShapeInteractions> _shapes =
                new();
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
                Get(ShapeKey(shape));

            /// <summary>The text interactions of a table cell.</summary>
            internal LegacyPptWriterShapeInteractions Get(PowerPointTable table,
                int row, int column) =>
                Get(CellKey(table, row, column));

            private LegacyPptWriterShapeInteractions Get(object key) =>
                _shapes.TryGetValue(key, out LegacyPptWriterShapeInteractions? value)
                    ? value
                    : LegacyPptWriterShapeInteractions.Empty;

            private static string CellKey(PowerPointTable table, int row,
                int column) => ShapeKey(table) + "#cell:"
                + row.ToString(System.Globalization.CultureInfo.InvariantCulture)
                + "," + column.ToString(System.Globalization.CultureInfo.InvariantCulture);

            internal LegacyPptWriterHyperlink? FindHyperlink(uint id) =>
                _hyperlinks.FirstOrDefault(link => link.Id == id);

            internal LegacyPptWriterHyperlink GetOrAdd(string target, string? screenTip) {
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

            internal void Add(PowerPointShape shape,
                LegacyPptWriterShapeInteractions interactions) =>
                _shapes.Add(ShapeKey(shape), interactions);

            internal void Add(PowerPointTable table, int row, int column,
                LegacyPptWriterShapeInteractions interactions) =>
                _shapes.Add(CellKey(table, row, column), interactions);
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
                ArgumentOutOfRangeException.ThrowIfNegativeOrZero(targetSlideNumber);
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
    }
}
