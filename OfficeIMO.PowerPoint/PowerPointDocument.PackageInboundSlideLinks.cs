using System.Xml;
using OfficeIMO.Core.Internal;
using OfficeIMO.Core.Internal.OpenXml;
using OfficeIMO.OpenXml.Internal;
using OfficeIMO.PowerPoint.OpenXml;

namespace OfficeIMO.PowerPoint {
    public sealed partial class PowerPointDocument {
        private sealed record InboundSlideLinkEdit(OfficeXmlTree Relationships, int Relationship, List<OfficeXmlNode> Actions,
            PowerPointPartRef Owner);

        private bool TryPlanPackageInboundSlideLinks(OfficeOpcEditablePackage package, PackageSlideList slides,
            string removedPart, string presentationRelationship, HashSet<string> reachable,
            out List<InboundSlideLinkEdit> edits, out Dictionary<string, string> excluded) {
            edits = new List<InboundSlideLinkEdit>();
            excluded = new Dictionary<string, string>(OfficeAsciiIgnoreCaseComparer.Instance);
            string main = PartIdentity(slides.Main);
            HashSet<string> surviving = Reachable(package, main, main, presentationRelationship, stopAt: removedPart);
            OfficeOpcPackage? loaded = LoadedPackageForCurrentBytes;
            if (!TryReadSlideRemovalRelationships(package, loaded, null, out _)) return false;
            foreach (string owner in reachable) {
                if (!TryReadSlideRemovalRelationships(package, loaded, owner, out IReadOnlyList<OfficeOpcRelationship> edges)) return false;
                if (!surviving.Contains(owner) || OfficeAsciiIgnoreCaseComparer.Instance.Equals(owner, main) ||
                    OfficeAsciiIgnoreCaseComparer.Instance.Equals(owner, removedPart)) continue;

                string? matchId = null;
                foreach (OfficeOpcRelationship edge in edges) {
                    if (edge.IsExternal || edge.TargetPartName is not { } target ||
                        !OfficeAsciiIgnoreCaseComparer.Instance.Equals(PartIdentity(target), removedPart)) continue;
                    // Ambiguous targets and other part kinds retain the model's deletion path.
                    if (matchId != null || edge.Type != PowerPointPartRef.SlideRelationshipType) return false;
                    matchId = edge.Id;
                }
                if (matchId == null) continue;
                if (!slides.Entries.Any(entry => OfficeAsciiIgnoreCaseComparer.Instance.Equals(PartIdentity(entry.Part), owner)) ||
                    IsOverPartLimit(package, owner)) return false;

                OfficeXmlTree relationships = package.GetRelationships(owner);
                int relationship = FindRelationship(relationships, matchId);
                if (relationship == OfficeXmlTree.NoNode) return false;

                OfficeXmlTree tree = package.GetXml(owner);
                if (tree.HasCommentsOrInstructions()) return false;
                var root = new OfficeXmlNode(tree, tree.Root);
                var actions = new List<OfficeXmlNode>();
                if (!root.Is(LinkPresentationNamespace, "sld")) return false;
                foreach (string name in SlideActionMarkup.Names) {
                    foreach (OfficeXmlNode action in root.Descendants(LinkDrawingNamespace, name)) {
                        if (action.Attribute(LinkRelationshipNamespace, "id") != matchId) continue;
                        OfficeXmlNode opaque = LinkOpaqueAncestor(action);
                        if (opaque.Exists) {
                            if (IsUnknownSlideAction(action, opaque)) continue;
                            return false;
                        }
                        if (!IsOwnedSlideLink(action) || action.Child(LinkDrawingNamespace, "snd").Exists) return false;
                        actions.Add(action);
                    }
                }
                edits.Add(new InboundSlideLinkEdit(relationships, relationship, actions, PowerPointPartRef.Package(this, owner)));
                excluded.Add(owner, matchId);
            }

            return edits.Count > 0;
        }

        private static bool TryReadSlideRemovalRelationships(OfficeOpcEditablePackage package, OfficeOpcPackage? loaded,
            string? owner, out IReadOnlyList<OfficeOpcRelationship> edges) {
            edges = Array.Empty<OfficeOpcRelationship>();
            IReadOnlyList<OfficeOpcRelationship>? relationships = loaded?.GetRelationships(owner) ?? ReadEditablePackageRelationships(package, owner);
            if (relationships == null) return false;
            foreach (OfficeOpcRelationship relationship in relationships) {
                if (relationship.Id.Length == 0 || relationship.Type.Length == 0) return false;
                XmlConvert.VerifyNCName(relationship.Id);
                // The SDK does not count hyperlink references to physical parts in its deletion graph.
                if (!relationship.IsExternal && relationship.Type == PowerPointPartRef.HyperlinkRelationshipType &&
                    relationship.TargetPartName is { } target && package.PartExists(target)) return false;
            }
            edges = relationships;
            return true;
        }

        private void ApplyPackageInboundSlideLinks(List<InboundSlideLinkEdit> edits) {
            foreach (InboundSlideLinkEdit edit in edits) {
                // Custom-show cleanup may already have removed a shared action or relationship.
                foreach (OfficeXmlNode action in edit.Actions) {
                    if (action.Parent.Exists) action.Remove();
                }
                if (edit.Relationships.ParentOf(edit.Relationship) != OfficeXmlTree.NoNode) edit.Relationships.Remove(edit.Relationship);
                edit.Owner.MarkEdited();
            }
            InvalidatePackageRelationships();
        }
    }
}
