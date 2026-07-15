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
            return TryReadInteractions(presentation.Slides, out catalog, out reason);
        }

        internal static bool TryReadInteractions(IEnumerable<PowerPointSlide> slides,
            out LegacyPptWriterInteractionCatalog catalog, out string? reason) {
            catalog = new LegacyPptWriterInteractionCatalog();
            reason = null;
            foreach (PowerPointSlide slide in slides) {
                foreach (PowerPointShape shape in slide.Shapes) {
                    if (!TryReadShapeInteractions(slide.SlidePart, shape, catalog,
                            out LegacyPptWriterShapeInteractions interactions, out reason)) {
                        catalog = new LegacyPptWriterInteractionCatalog();
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
                        out A.HyperlinkOnHover? hoverElement, out reason)
                    || !TryReadHyperlink(slidePart, clickElement,
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
            interactions = new LegacyPptWriterShapeInteractions(shapeActions, textActions);
            return true;
        }

        private static bool TryReadTextInteractions(SlidePart slidePart,
            P.TextBody textBody, LegacyPptWriterInteractionCatalog catalog,
            ICollection<LegacyPptWriterTextInteraction> result, out string? reason) {
            reason = null;
            A.Paragraph[] paragraphs = textBody.Elements<A.Paragraph>().ToArray();
            int position = 0;
            for (int paragraphIndex = 0; paragraphIndex < paragraphs.Length; paragraphIndex++) {
                foreach (OpenXmlElement child in paragraphs[paragraphIndex].ChildElements) {
                    string text = child switch {
                        A.Run run => run.Text?.Text ?? string.Empty,
                        A.Field field => field.Text?.Text ?? string.Empty,
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
            if (hyperlink.ChildElements.Count > 0
                || hyperlink.GetAttributes().Any(attribute =>
                    !string.Equals(attribute.LocalName, "id", StringComparison.Ordinal)
                    && !string.Equals(attribute.LocalName, "action", StringComparison.Ordinal)
                    && !string.Equals(attribute.LocalName, "tooltip", StringComparison.Ordinal))) {
                reason = "Hyperlink target frames, history/highlight flags, sounds, and extension data are not encoded yet.";
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

            if (!string.IsNullOrEmpty(relationshipId)) {
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
                    MapHyperlinkType(relationship.Uri), target.Id);
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
                hyperlinkIdReference: 0);
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
                _ => null
            };

        internal static byte[] BuildExternalObjectListRecord(
            LegacyPptWriterInteractionCatalog catalog) {
            if (catalog == null) throw new ArgumentNullException(nameof(catalog));
            if (catalog.Hyperlinks.Count == 0) return Array.Empty<byte>();
            var listAtomPayload = new byte[4];
            WriteUInt32(listAtomPayload, 0, catalog.Hyperlinks.Max(link => link.Id));
            var children = new List<byte[]> {
                BuildRecord(version: 0, instance: 0, RecordExternalObjectListAtom,
                    listAtomPayload)
            };
            foreach (LegacyPptWriterHyperlink hyperlink in catalog.Hyperlinks) {
                children.Add(BuildExternalHyperlinkRecord(hyperlink.Id,
                    hyperlink.Target));
            }
            return BuildContainer(RecordExternalObjectList, instance: 0, children);
        }

        internal static byte[] BuildExternalHyperlinkRecord(uint id, string target) {
            var atomPayload = new byte[4];
            WriteUInt32(atomPayload, 0, id);
            return BuildContainer(RecordExternalHyperlink, instance: 0,
                new[] {
                    BuildRecord(version: 0, instance: 0,
                        RecordExternalHyperlinkAtom, atomPayload),
                    BuildRecord(version: 0, instance: 1, RecordCString,
                        System.Text.Encoding.Unicode.GetBytes(target))
                });
        }

        internal static byte[] BuildInteractiveInfoRecord(
            LegacyPptWriterInteraction interaction) {
            var payload = new byte[16];
            WriteUInt32(payload, 4, interaction.HyperlinkIdReference);
            payload[8] = (byte)interaction.Action;
            payload[10] = (byte)interaction.Jump;
            payload[12] = (byte)interaction.HyperlinkType;
            byte[] atom = BuildRecord(version: 0, instance: 0,
                RecordInteractiveInfoAtom, payload);
            return BuildContainer(RecordInteractiveInfo,
                (ushort)interaction.Trigger, new[] { atom });
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

            internal IReadOnlyList<LegacyPptWriterHyperlink> Hyperlinks =>
                new ReadOnlyCollection<LegacyPptWriterHyperlink>(_hyperlinks);

            internal bool HasInteractions => _shapes.Count > 0;

            internal LegacyPptWriterShapeInteractions Get(PowerPointShape shape) =>
                _shapes.TryGetValue(shape.Element, out LegacyPptWriterShapeInteractions? value)
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
                LegacyPptHyperlinkType hyperlinkType, uint hyperlinkIdReference) {
                Trigger = trigger;
                Action = action;
                Jump = jump;
                HyperlinkType = hyperlinkType;
                HyperlinkIdReference = hyperlinkIdReference;
            }

            internal LegacyPptInteractionTrigger Trigger { get; }
            internal LegacyPptInteractionAction Action { get; }
            internal LegacyPptInteractionJump Jump { get; }
            internal LegacyPptHyperlinkType HyperlinkType { get; }
            internal uint HyperlinkIdReference { get; }
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

            internal uint Id { get; }
            internal string Target { get; }
            internal string? ScreenTip { get; }
            internal uint ExtensionFlags { get; }
        }

        internal static string CreateHyperlinkKey(string target, string? screenTip) =>
            target + "\0" + (screenTip == null ? "-" : screenTip.Length + ":" + screenTip);

        private sealed class ReferenceComparer : IEqualityComparer<OpenXmlElement> {
            internal static ReferenceComparer Instance { get; } = new();
            public bool Equals(OpenXmlElement? x, OpenXmlElement? y) => ReferenceEquals(x, y);
            public int GetHashCode(OpenXmlElement obj) => RuntimeHelpers.GetHashCode(obj);
        }
    }
}
