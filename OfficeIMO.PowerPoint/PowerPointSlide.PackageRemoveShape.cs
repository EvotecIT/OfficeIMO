using System.Runtime.CompilerServices;
using OfficeIMO.Core.Internal.OpenXml;
using OfficeIMO.OpenXml.Internal;
using OfficeIMO.PowerPoint.OpenXml;

namespace OfficeIMO.PowerPoint {
    public partial class PowerPointSlide {
        private ConditionalWeakTable<OfficeXmlTree, HashSet<int>>? _detachedPackageShapes;

        private bool TryRemovePackageShape(PowerPointShape shape, int index) {
            if (!shape.TryEditPackage(out _, out _)) return false;
            PowerPointPartRef part = PartRef.ForMovingXmlNodes();
            if (part.IsSdk) return false;
            OfficeXmlNode element = shape.XmlElement;
            _ = RequireShapeParent(element);
            if (shape is PowerPointOleObject || ShapeNodesToRenumber(element).Any(IsOleFrame)) return false;

            HashSet<string> relationships = GetRemovedShapeRelationships(element);
            foreach (string id in relationships) {
                if (part.TryGetRelationship(id, out PowerPointRelationship relationship) && !CanRemovePackageShapeRelationship(relationship)) {
                    // Dependent package parts still use the existing graph removal path.
                    return false;
                }
            }

            IReadOnlyList<PowerPointClassicAnimation> previous = ReadClassicAnimations();
            if (part.Current.IsSdk) return false;
            if (previous.Count > 0) {
                var removedIds = new HashSet<uint>();
                foreach (OfficeXmlNode child in ShapeNodesToRenumber(element)) {
                    OfficeXmlNode drawing = PowerPointShape.NonVisualDrawingElement(child, create: false);
                    if (drawing.UInt32Attribute("id") is uint id) removedIds.Add(id);
                }
                PowerPointClassicAnimation[] remaining = previous.Where(animation => !removedIds.Contains(animation.ShapeId)).ToArray();
                if (remaining.Length != previous.Count) {
                    OfficeXmlNode timing = part.Root.Child(ClassicPresentationNamespace, "timing");
                    if (timing.Exists && !HasOnlyPackageClassicTiming(timing, previous)) return false;
                    SetClassicAnimations(remaining);
                }
            }

            OfficeXmlTree tree = element.Tree!;
            HashSet<int> detachedRoots = (_detachedPackageShapes ??= new ConditionalWeakTable<OfficeXmlTree, HashSet<int>>())
                .GetValue(tree, static _ => new HashSet<int>());
            if (detachedRoots.Count == 0) tree.RetainSourceBytes();
            element.Remove();
            detachedRoots.Add(element.Node);
            ShapeList.RemoveAt(index);
            foreach (string id in relationships) part.RemoveRelationshipIfUnused(id, part);
            part.MarkEdited();
            return true;
        }

        /// <summary>Whether the node is in a removed shape whose wrapper can still hold and edit its detached XML.</summary>
        internal bool IsInDetachedPackageShape(OfficeXmlNode node) {
            if (node.Tree is not { } tree || _detachedPackageShapes == null
                || !_detachedPackageShapes.TryGetValue(tree, out HashSet<int>? detachedRoots)) return false;
            int root = node.Node;
            for (int parent = tree.ParentOf(root); parent != OfficeXmlTree.NoNode; parent = tree.ParentOf(root)) root = parent;
            return detachedRoots.Contains(root);
        }

        private static HashSet<string> GetRemovedShapeRelationships(OfficeXmlNode element) {
            string relationshipNamespace = OfficeXmlNamespaces.KnownUri(OfficeXmlNamespaces.Relationships)!;
            var relationships = new HashSet<string>(StringComparer.Ordinal);
            var pending = new Stack<OfficeXmlNode>();
            pending.Push(element);
            while (pending.Count > 0) {
                OfficeXmlNode node = pending.Pop();
                foreach (OfficeXmlAttribute attribute in node.Attributes()) {
                    if (attribute.NamespaceUri == relationshipNamespace && attribute.Value.Length != 0) relationships.Add(attribute.Value);
                }
                for (OfficeXmlNode child = node.LastChild; child.Exists; child = child.PreviousSibling) pending.Push(child);
            }
            return relationships;
        }

        private static bool CanRemovePackageShapeRelationship(PowerPointRelationship relationship) {
            const string prefix = OfficeXmlNamespaces.TransitionalRelationshipTypePrefix;
            string type = OfficeXmlNamespaces.NormalizeRelationshipType(relationship.Type);
            if (relationship.IsExternal) return type != prefix + "oleObject";
            bool hyperlink = type == PowerPointPartRef.HyperlinkRelationshipType;
            string? contentType = relationship.TargetPart.ContentType;
            if (contentType == null) return hyperlink;
            bool media = contentType.StartsWith("image/", StringComparison.OrdinalIgnoreCase)
                || contentType.StartsWith("audio/", StringComparison.OrdinalIgnoreCase)
                || contentType.StartsWith("video/", StringComparison.OrdinalIgnoreCase);
            if (hyperlink) return !media;
            if (!media || type is not (prefix + "image" or prefix + "audio" or prefix + "video" or PowerPointMedia.PlaybackMediaRelationshipType)) return false;
            PowerPointPartRef target = relationship.TargetPart;
            return target.Document?.ReadablePackage is { } package
                && !package.PartExists(OfficeOpcPartNames.GetRelationshipsPartName(target.Name));
        }

        private static bool IsOleFrame(OfficeXmlNode element) {
            const int p = OfficeXmlNamespaces.PresentationML;
            const int a = OfficeXmlNamespaces.DrawingML;
            return element.Is(p, "graphicFrame")
                && element.Child(a, "graphic").Child(a, "graphicData").Child(p, "oleObj").Exists;
        }
    }
}
