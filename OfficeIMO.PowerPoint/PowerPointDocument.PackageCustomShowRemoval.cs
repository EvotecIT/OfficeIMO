using System.Xml;
using OfficeIMO.Core.Internal;
using OfficeIMO.Core.Internal.OpenXml;
using OfficeIMO.OpenXml.Internal;
using OfficeIMO.PowerPoint.OpenXml;

namespace OfficeIMO.PowerPoint {
    public sealed partial class PowerPointDocument {
        private const int LinkDrawingNamespace = OfficeXmlNamespaces.DrawingML;
        private const int LinkPresentationNamespace = OfficeXmlNamespaces.PresentationML;
        private const int LinkRelationshipNamespace = OfficeXmlNamespaces.Relationships;
        private static class CustomShowLinkMarkup {
            internal static readonly string[] Names = { "hlinkClick", "hlinkMouseOver", "hlinkHover" };
        }

        private bool TryRemoveOwnedCustomShow(PowerPointCustomShow show) {
            if (EditablePackage is not { } package) return false;
            PowerPointPartRef presentation = PresentationPartRef.ForMovingXmlNodes();
            if (presentation.IsSdk) return false;

            OfficeXmlNode node = RequireOwnedCustomShow(presentation, show);
            uint? id = PowerPointCustomShow.ReadId(node);
            List<CustomShowLinkEdit>? edits = null;
            if (id is uint showId && !TryPlanOwnedCustomShowLinks(package, presentation, new[] { showId }, out edits)) return false;

            OfficeXmlNode list = node.Parent;
            // Keep the detached show readable after the package that held it has been released.
            node.Tree!.RetainSourceBytes();
            node.Remove();
            if (KnownCustomShowChild(list, "custShow").IsNone) list.Remove();
            ApplyOwnedCustomShowLinkEdits(edits);
            presentation.MarkEdited();
            return true;
        }

        private static void ApplyOwnedCustomShowLinkEdits(List<CustomShowLinkEdit>? edits) {
            if (edits == null) return;
            foreach (CustomShowLinkEdit edit in edits) {
                foreach (OfficeXmlNode action in edit.Actions) action.Remove();
                if (edit.Relationships is { } relationships) {
                    foreach (int relationship in edit.RemovedRelationships!) relationships.Remove(relationship);
                    edit.Owner.Document!.InvalidatePackageRelationships();
                }
                edit.Owner.MarkEdited();
            }
        }

        private sealed record CustomShowLinkEdit(PowerPointPartRef Owner, List<OfficeXmlNode> Actions,
            OfficeXmlTree? Relationships, List<int>? RemovedRelationships);

        private bool TryPlanOwnedCustomShowLinks(OfficeOpcEditablePackage package, PowerPointPartRef presentation,
            ICollection<uint> showIds, out List<CustomShowLinkEdit>? edits) {
            edits = null;
            var xmlParts = new Dictionary<string, bool>(OfficeAsciiIgnoreCaseComparer.Instance) {
                [presentation.Name] = false
            };
            try {
                if (!TryVisitPackageGraph(package, CancellationToken.None, relationship => {
                    string contentType = relationship.ContentType;
                    if (contentType.EndsWith("+xml", StringComparison.OrdinalIgnoreCase) ||
                        contentType.Equals("application/xml", StringComparison.OrdinalIgnoreCase) ||
                        contentType.Equals("text/xml", StringComparison.OrdinalIgnoreCase)) {
                        xmlParts[relationship.Part] = relationship.Type == PowerPointPartRef.SlideRelationshipType &&
                            contentType == SlideContentType;
                    }
                }, root: presentation.Name)) return false;

                foreach (KeyValuePair<string, bool> part in xmlParts) {
                    string partName = part.Key;
                    if (IsOverPartLimit(package, partName)) return false;
                    if (package.TryReadUnparsed(partName, out OfficeXmlPullReader reader, out OfficePooledBytes content,
                        processMarkupCompatibility: false)) {
                        using (content) {
                            if (!MayHaveCustomShowLink(reader, showIds)) continue;
                        }
                    }
                    // An opaque part must keep its original bytes if an existing owned edit is flushed before SDK cleanup.
                    // A newly cached XML tree would otherwise re-encode an unchanged UTF-16 part on legacy targets.
                    if (!part.Value && package.GetParsedXml(partName) == null) return false;
                    OfficeXmlTree tree = package.GetXml(partName);
                    OfficeXmlNode root = new OfficeXmlNode(tree, tree.Root);
                    bool canEdit = part.Value && !tree.HasCommentsOrInstructions();
                    PowerPointPartRef owner = PowerPointPartRef.Package(this, partName);
                    var actions = new List<OfficeXmlNode>();
                    var relationships = new HashSet<string>(StringComparer.Ordinal);
                    foreach (OfficeXmlNode action in FindCustomShowLinks(root, showIds)) {
                        OfficeXmlNode extension = CustomShowLinkExtension(action);
                        if (extension.Exists) {
                            if (IsUnknownCustomShowDecoy(action, extension)) continue;
                            // The SDK can recognize slide shapes and supported extensions inside these wrappers.
                            return false;
                        }
                        if (!canEdit || !IsOwnedSlideCustomShowLink(action) ||
                            action.Child(LinkDrawingNamespace, "snd").Exists) return false;

                        string? relationship = action.Attribute(LinkRelationshipNamespace, "id");
                        if (!string.IsNullOrEmpty(relationship) && !PlanOwnedCustomShowRelationship(owner, presentation, relationship!, relationships)) {
                            return false;
                        }
                        actions.Add(action);
                    }
                    if (actions.Count != 0) {
                        (OfficeXmlTree? relationshipTree, List<int>? remove) = PlanUnusedCustomShowRelationships(package, owner, root, actions, relationships);
                        (edits ??= new List<CustomShowLinkEdit>()).Add(new CustomShowLinkEdit(owner, actions, relationshipTree, remove));
                    }
                }
                return true;
            } catch (Exception exception) when (exception is InvalidDataException or XmlException or IOException or NotSupportedException) {
                return false;
            }
        }

        private static bool MayHaveCustomShowLink(OfficeXmlPullReader reader, ICollection<uint> showIds) {
            // A link in a compatibility wrapper can require the original typed cleanup. Include it in this
            // preflight, then accept owned edits only for known content outside the wrappers.
            while (reader.Read()) {
                if (reader.NodeKind != OfficeXmlNodeKind.StartElement) continue;
                if (reader.NamespaceId == LinkDrawingNamespace &&
                    (reader.LocalNameEquals("hlinkClick") || reader.LocalNameEquals("hlinkMouseOver") || reader.LocalNameEquals("hlinkHover")) &&
                    reader.TryGetAttribute(OfficeXmlNamespaces.None, "action", out string action) &&
                    TryParseCustomShowAction(action, out uint id, out _) && showIds.Contains(id)) return true;
            }
            return false;
        }

        private static IEnumerable<OfficeXmlNode> FindCustomShowLinks(OfficeXmlNode root, ICollection<uint> showIds) {
            foreach (string name in CustomShowLinkMarkup.Names) {
                foreach (OfficeXmlNode link in root.Descendants(LinkDrawingNamespace, name)) {
                    if (TryParseCustomShowAction(link.Attribute("action"), out uint id, out _) && showIds.Contains(id)) yield return link;
                }
            }
        }

        private static OfficeXmlNode CustomShowLinkExtension(OfficeXmlNode link) {
            for (OfficeXmlNode parent = link.Parent; parent.Exists; parent = parent.Parent) {
                if (parent.IsExtensionMarkup()) return parent;
            }
            return default;
        }

        private static bool IsUnknownCustomShowDecoy(OfficeXmlNode link, OfficeXmlNode extension) {
            // A hyperlink put directly in a slide extension or an alternate slide root has no typed parent.
            // An alternate shape, group or run can contain typed links and needs the original cleanup.
            if (!extension.Parent.Is(LinkPresentationNamespace, "sld") || extension.Parent.Parent.Exists) return false;
            OfficeXmlNode parent = link.Parent;
            if (parent.Parent != extension) return false;
            return extension.Is(LinkPresentationNamespace, "extLst")
                ? parent.Is(LinkPresentationNamespace, "ext")
                : parent.Is(OfficeXmlNamespaces.MarkupCompatibility, "Choice") ||
                  parent.Is(OfficeXmlNamespaces.MarkupCompatibility, "Fallback");
        }

        private static bool IsOwnedSlideCustomShowLink(OfficeXmlNode link) {
            OfficeXmlNode owner = link.Parent;
            if (owner.Is(LinkPresentationNamespace, "cNvPr")) {
                OfficeXmlNode shape = owner.Parent.Parent;
                return link.LocalName is "hlinkClick" or "hlinkHover" &&
                    PowerPointShape.NonVisualDrawingElement(shape, create: false).Equals(owner) && IsCustomShowLinkShape(shape);
            }

            if (link.LocalName is not ("hlinkClick" or "hlinkMouseOver")) return false;
            OfficeXmlNode paragraph = owner.Parent;
            if ((owner.Is(LinkDrawingNamespace, "rPr") && paragraph.NamespaceId == LinkDrawingNamespace &&
                paragraph.LocalName is "r" or "br" or "fld") ||
                (owner.Is(LinkDrawingNamespace, "defRPr") && paragraph.Is(LinkDrawingNamespace, "pPr"))) {
                paragraph = paragraph.Parent;
            } else if (!owner.Is(LinkDrawingNamespace, "endParaRPr")) {
                return false;
            }
            OfficeXmlNode body = paragraph.Parent;
            return paragraph.Is(LinkDrawingNamespace, "p") && body.Is(LinkPresentationNamespace, "txBody") &&
                body.Parent.Is(LinkPresentationNamespace, "sp") && IsCustomShowLinkShape(body.Parent);
        }

        private static bool IsCustomShowLinkShape(OfficeXmlNode shape) {
            OfficeXmlNode container = shape.Parent;
            while (container.Is(LinkPresentationNamespace, "grpSp")) container = container.Parent;
            OfficeXmlNode common = container.Parent;
            OfficeXmlNode slide = common.Parent;
            return container.Is(LinkPresentationNamespace, "spTree") && common.Is(LinkPresentationNamespace, "cSld") &&
                slide.Is(LinkPresentationNamespace, "sld") && slide.Parent.IsNone;
        }

        private static (OfficeXmlTree? Tree, List<int>? Remove) PlanUnusedCustomShowRelationships(OfficeOpcEditablePackage package,
            PowerPointPartRef owner, OfficeXmlNode root, List<OfficeXmlNode> actions, HashSet<string> candidateIds) {
            if (candidateIds.Count == 0) return default;

            // Decoding malformed attributes can fail, so resolve usage before any action, show or slide is removed.
            var removedActions = new HashSet<int>(actions.Select(action => action.Node));
            OfficeXmlTree relationships = package.GetRelationships(owner.Name);
            var remove = new List<int>();
            foreach (string id in candidateIds) {
                if (PowerPointPartRef.ReferencesRelationship(root, id, removedActions)) continue;
                int node = FindRelationship(relationships, id);
                if (node != OfficeXmlTree.NoNode) remove.Add(node);
            }
            return remove.Count == 0 ? default : (relationships, remove);
        }

        private static bool PlanOwnedCustomShowRelationship(PowerPointPartRef owner, PowerPointPartRef presentation,
            string id, HashSet<string> removeIfUnused) {
            if (!owner.TryGetRelationship(id, out PowerPointRelationship relationship)) return true;
            string type = OfficeXmlNamespaces.NormalizeRelationshipType(relationship.Type);
            if (relationship.IsExternal || type == PowerPointPartRef.HyperlinkRelationshipType) {
                // An internal hyperlink is a reference, even when its URI points to a real package part.
                if (!relationship.IsExternal && !relationship.TargetPart.IsNone) return false;
                removeIfUnused.Add(id);
                return true;
            }
            if (type != PowerPointPartRef.SlideRelationshipType) return true;
            if (relationship.TargetPart.Kind != PowerPointPartKind.Slide ||
                presentation.RelationshipIdOf(relationship.TargetPart, PowerPointPartRef.SlideRelationshipType) == null) return false;
            removeIfUnused.Add(id);
            return true;
        }
    }
}
