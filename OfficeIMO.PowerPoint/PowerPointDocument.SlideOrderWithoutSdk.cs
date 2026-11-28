using System.Xml;
using OfficeIMO.Core.Internal;
using OfficeIMO.Core.Internal.OpenXml;
using OfficeIMO.PowerPoint.OpenXml;

namespace OfficeIMO.PowerPoint {
    /// <summary>
    /// Removing and moving slides in a presentation whose SDK package is not open yet edits the slide list on the Open XML
    /// layer, in the package that holds the other edits made without the SDK. Whatever the slide model would do beyond the
    /// slide list, its sections and the removed slide's own parts is left to the model.
    /// </summary>
    public sealed partial class PowerPointDocument {
        /// <summary>The first slide list of the presentation part, with each entry's relationship and slide part.</summary>
        private sealed class PackageSlideList {
            private PackageSlideList(string main, OfficeXmlTree presentation, int list) {
                Main = main;
                Presentation = presentation;
                List = list;
            }

            internal string Main { get; }
            internal OfficeXmlTree Presentation { get; }
            internal int List { get; }
            internal List<(int Node, string Relationship, string Part)> Entries { get; } = new();

            /// <summary>The presentation part and its slide list node, without reading the entries.</summary>
            internal static PackageSlideList? ReadPresentation(OfficeOpcEditablePackage package) {
                string? main = FirstRelated(package, null, OfficeOpcPackage.OfficeDocumentRelationshipType, out _);
                if (main == null) return null;
                OfficeXmlTree presentation = package.GetXml(main);
                return new PackageSlideList(main, presentation, presentation.FindChild(presentation.Root, OfficeXmlNamespaces.PresentationML, "sldIdLst"));
            }

            // Null when an entry's slide is missing, external or listed twice.
            internal static PackageSlideList? Read(OfficeOpcEditablePackage package, bool requireConsistentNamespaces = false) {
                if (ReadPresentation(package) is not { } slides) return null;
                if (slides.List == OfficeXmlTree.NoNode) return slides;
                string main = slides.Main;
                OfficeXmlTree presentation = slides.Presentation;
                int presentationUri = requireConsistentNamespaces ? presentation.UriIdOf(presentation.Root) : 0;
                if (requireConsistentNamespaces && (!presentation.IsElement(presentation.Root, OfficeXmlNamespaces.PresentationML, "presentation") ||
                    presentation.UriIdOf(slides.List) != presentationUri)) return null;
                const string strictRelationships = "http://purl.oclc.org/ooxml/officeDocument/relationships";
                bool strict = requireConsistentNamespaces && OfficeXmlNamespaces.IsStrictUriId(presentationUri);
                int transitionalRelationshipUri = requireConsistentNamespaces && !strict
                    ? OfficeXmlNamespaces.TransitionalUriId(OfficeXmlNamespaces.Relationships)
                    : 0;
                string slideRelationshipType = strict ? strictRelationships + "/slide" : OfficeOpcPackage.SlideRelationshipType;
                // The first relationship with an id is the one it resolves to, and an external one resolves to no part.
                var targets = new Dictionary<string, string?>(StringComparer.Ordinal);
                string mainRelationships = OfficeOpcPartNames.GetRelationshipsPartName(main);
                if (package.PartExists(mainRelationships)) {
                    OfficeXmlTree relationships = package.GetXml(mainRelationships);
                    for (int node = relationships.FindChild(relationships.Root, OfficeXmlNamespaces.PackageRelationships, "Relationship"); node != OfficeXmlTree.NoNode;
                         node = relationships.FindNextSibling(node, OfficeXmlNamespaces.PackageRelationships, "Relationship")) {
                        string? relationshipId = relationships.GetAttribute(node, OfficeXmlNamespaces.None, "Id");
                        if (relationshipId == null || targets.ContainsKey(relationshipId)) continue;
                        bool external = string.Equals(relationships.GetAttribute(node, OfficeXmlNamespaces.None, "TargetMode"), "External", StringComparison.OrdinalIgnoreCase);
                        bool slideRelationship = !requireConsistentNamespaces || relationships.AttributeEquals(node, OfficeXmlNamespaces.None, "Type", slideRelationshipType);
                        targets.Add(relationshipId, external || !slideRelationship
                            ? null : OfficeOpcPartNames.ResolveRelationshipTarget(mainRelationships, relationships.GetAttribute(node, OfficeXmlNamespaces.None, "Target")));
                    }
                }
                var seen = new HashSet<string>(OfficeAsciiIgnoreCaseComparer.Instance);
                for (int entry = presentation.FindChild(slides.List, OfficeXmlNamespaces.PresentationML, "sldId"); entry != OfficeXmlTree.NoNode;
                     entry = presentation.FindNextSibling(entry, OfficeXmlNamespaces.PresentationML, "sldId")) {
                    if (requireConsistentNamespaces && presentation.UriIdOf(entry) != presentationUri) return null;
                    string? id;
                    if (requireConsistentNamespaces) {
                        id = presentation.GetUniqueAttribute(entry, OfficeXmlNamespaces.Relationships, "id", out int uri);
                        if (strict ? !OfficeXmlNamespaces.IsStrictUriId(uri) : uri != transitionalRelationshipUri) return null;
                    } else {
                        id = presentation.GetAttribute(entry, OfficeXmlNamespaces.Relationships, "id");
                    }
                    string? slide = id != null && targets.TryGetValue(id, out string? target) ? target : null;
                    if (slide == null || !package.PartExists(slide) || !seen.Add(PartIdentity(slide))) return null;
                    slides.Entries.Add((entry, id!, slide));
                }
                return slides;
            }

            // The slide ids in list order, or null when one is missing or not a number.
            internal List<uint>? ReadIds() {
                var ids = new List<uint>(Entries.Count);
                foreach ((int node, _, _) in Entries) {
                    if (!PowerPointTreeSections.TryReadId(Presentation, node, out uint id)) return null;
                    ids.Add(id);
                }
                return ids;
            }
        }

        /// <summary>Moves a slide without the SDK, or returns false, having changed nothing, for the model to move it.</summary>
        private bool TryMoveSlideWithoutSdk(int fromIndex, int toIndex) {
            if (GetPackage(edit: true) is not { } package) return false;
            try {
                PackageSlideList? slides = PackageSlideList.Read(package);
                if (slides == null || fromIndex < 0 || toIndex < 0 || fromIndex >= slides.Entries.Count || toIndex >= slides.Entries.Count) return false;
                if (fromIndex == toIndex) return true;
                OfficeXmlTree presentation = slides.Presentation;
                // The model writes the list again from its entries, which drops anything else in it.
                for (int child = presentation.FirstChildOf(slides.List); child != OfficeXmlTree.NoNode; child = presentation.NextSiblingOf(child)) {
                    if (presentation.KindOf(child) == OfficeXmlTreeNodeKind.Element && !presentation.IsElement(child, OfficeXmlNamespaces.PresentationML, "sldId")) return false;
                }
                List<uint>? ids = slides.ReadIds();
                if (ids == null) return false;
                var nodes = slides.Entries.Select(entry => entry.Node).ToList();
                int moving = nodes[fromIndex];
                uint movingId = ids[fromIndex];
                nodes.RemoveAt(fromIndex);
                nodes.Insert(toIndex, moving);
                ids.RemoveAt(fromIndex);
                ids.Insert(toIndex, movingId);
                if (!PowerPointTreeSections.TryPlan(presentation, package.Namespaces.GetId(PowerPointTreeSections.P14Namespace), ids, out PowerPointTreeSections.Plan? sections)) return false;

                while (presentation.FirstChildOf(slides.List) != OfficeXmlTree.NoNode) presentation.Remove(presentation.FirstChildOf(slides.List));
                foreach (int node in nodes) presentation.Append(slides.List, node);
                if (sections != null) PowerPointTreeSections.Apply(sections);
                PackageSlidesChanged();
                _pendingEdits = true;
                return true;
            } catch (Exception exception) when (exception is XmlException or InvalidDataException) {
                return false;
            }
        }

        /// <summary>Removes a slide without the SDK, or returns false, having changed nothing, for the model to remove it.</summary>
        private bool TryRemoveSlideWithoutSdk(int index) {
            if (GetPackage(edit: true) is not { } package) return false;
            try {
                PackageSlideList? slides = PackageSlideList.Read(package, requireConsistentNamespaces: true);
                if (slides == null || index < 0 || index >= slides.Entries.Count) return false;
                OfficeXmlTree presentation = slides.Presentation;
                (int entry, string relationship, string slide) = slides.Entries[index];
                List<uint>? ids = slides.ReadIds();
                if (ids == null) return false;
                ids.RemoveAt(index);

                // The slide's parts that nothing else reaches go with it, as the SDK deletes them.
                string presentationRelationships = OfficeOpcPartNames.GetRelationshipsPartName(slides.Main);
                int presentationRelationship = FindRelationship(package.GetXml(presentationRelationships), relationship);
                if (presentationRelationship == OfficeXmlTree.NoNode) return false;
                string main = ReachablePartIdentity(slides.Main), slidePart = ReachablePartIdentity(slide);
                HashSet<string> kept = Reachable(package, string.Empty, main, relationship);
                List<InboundSlideLinkEdit>? inboundLinks = null;
                if (kept.Contains(slidePart)) {
                    if (!TryPlanPackageInboundSlideLinks(package, slides, slidePart, relationship, kept, out inboundLinks,
                        out Dictionary<string, string> excluded)) return false;
                    kept = Reachable(package, string.Empty, main, relationship, excludedLinks: excluded);
                    if (kept.Contains(slidePart)) return false;
                }
                var removed = new List<string>();
                foreach (string part in Reachable(package, slidePart, null, null)) {
                    if (!kept.Contains(part)) removed.Add(part);
                }
                bool hasClassicComments = false;
                foreach (string part in removed) {
                    if (!CanRemoveWithSlide(package, part, slidePart, ref hasClassicComments)) return false;
                }
                ClassicSlideCommentRemoval? classicComments = null;
                if (hasClassicComments && !TryPlanPackageSlideClassicComments(package, slides, slide, removed, out classicComments)) return false;
                if (!PowerPointTreeSections.TryPlan(presentation, package.Namespaces.GetId(PowerPointTreeSections.P14Namespace), ids, out PowerPointTreeSections.Plan? sections)) return false;

                if (!TryPlanPackageSlideCustomShows(package, slides, relationship, out SlideCustomShowPruning? customShows)) return false;

                presentation.Remove(entry);
                if (customShows != null) ApplyPackageSlideCustomShows(customShows);
                if (inboundLinks != null) ApplyPackageInboundSlideLinks(inboundLinks);
                OfficeXmlTree relationships = package.GetXml(presentationRelationships);
                relationships.Remove(presentationRelationship);
                RelationshipsOf(package, main).RemoveAll(edge => string.Equals(edge.Id, relationship, StringComparison.Ordinal));
                // Text a caller still holds from the slide or its notes stops working before their parts go.
                MarkPackageSlideRemoved(slide);
                InvalidateRemovedPackageObjects(sdkOpening: false);
                foreach (string part in removed) {
                    _pendingRelationships!.Remove(part);
                    string partRelationships = OfficeOpcPartNames.GetRelationshipsPartName(part);
                    if (package.PartExists(partRelationships)) package.DeletePart(partRelationships);
                    package.DeletePart(part);
                }
                if (classicComments != null) ApplyPackageSlideClassicCommentRemoval(package, classicComments);
                if (sections != null) PowerPointTreeSections.Apply(sections);
                _deferredSlideCount = ids.Count;
                // The model searches for new slide relationship ids from the start again after a removal.
                _slideRelationshipSearchStart = 1;
                PackageSlidesChanged(slide, retainValidatedClassicCommentParts: hasClassicComments);
                _pendingEdits = true;
                return true;
            } catch (Exception exception) when (exception is XmlException or InvalidDataException) {
                return false;
            }
        }

        // Classic comment authors are reconciled before removing the slide. Other comments and linked media use the SDK cleanup.
        private static bool CanRemoveWithSlide(OfficeOpcEditablePackage package, string part, string slide, ref bool hasClassicComments) {
            string relationshipsPart = OfficeOpcPartNames.GetRelationshipsPartName(part);
            if (!package.PartExists(relationshipsPart)) return true;
            OfficeXmlTree relationships = package.GetXml(relationshipsPart);
            for (int node = relationships.FindChild(relationships.Root, OfficeXmlNamespaces.PackageRelationships, "Relationship"); node != OfficeXmlTree.NoNode;
                 node = relationships.FindNextSibling(node, OfficeXmlNamespaces.PackageRelationships, "Relationship")) {
                string type = relationships.GetAttribute(node, OfficeXmlNamespaces.None, "Type") ?? string.Empty;
                if (EndsWith(type, "/media") || EndsWith(type, "/audio") || EndsWith(type, "/video")) return false;
                if (EndsWith(type, "/comments") && OfficeAsciiIgnoreCaseComparer.Instance.Equals(part, slide)) {
                    if (type != ClassicCommentsRelationship) return false;
                    hasClassicComments = true;
                }
            }
            return true;
        }

        private static bool EndsWith(string value, string suffix) => value.EndsWith(suffix, StringComparison.OrdinalIgnoreCase);

        private static int FindRelationship(OfficeXmlTree relationships, string id) {
            for (int node = relationships.FindChild(relationships.Root, OfficeXmlNamespaces.PackageRelationships, "Relationship"); node != OfficeXmlTree.NoNode;
                 node = relationships.FindNextSibling(node, OfficeXmlNamespaces.PackageRelationships, "Relationship")) {
                if (string.Equals(relationships.GetAttribute(node, OfficeXmlNamespaces.None, "Id"), id, StringComparison.Ordinal)) return node;
            }
            return OfficeXmlTree.NoNode;
        }

        // Internal relationships of the parts the pending package holds, by source part identity ("" for the package),
        // read once: removing slides walks them for every slide. Only slide removal changes them, and it keeps them current.
        private Dictionary<string, List<(string Id, string Target)>>? _pendingRelationships;

        /// <summary>Drops the cached relationship edges between parts after a relationship was added or removed.</summary>
        internal void InvalidatePackageRelationships() => _pendingRelationships = null;

        private List<(string Id, string Target)> RelationshipsOf(OfficeOpcEditablePackage package, string source) {
            _pendingRelationships ??= new Dictionary<string, List<(string Id, string Target)>>(OfficeAsciiIgnoreCaseComparer.Instance);
            if (_pendingRelationships.TryGetValue(source, out List<(string Id, string Target)>? found)) return found;
            var edges = new List<(string Id, string Target)>();
            string relationshipsPart = OfficeOpcPartNames.GetRelationshipsPartName(source.Length == 0 ? null : source);
            // Relationships nothing has changed are read as they are stored, without building a tree for each part.
            if (package.TryReadUnparsed(relationshipsPart, out OfficeXmlPullReader reader, out OfficePooledBytes content)) {
                using (content) {
                    while (reader.Read()) {
                        if (reader.NodeKind != OfficeXmlNodeKind.StartElement || reader.Depth != 2
                            || !reader.IsElement(OfficeXmlNamespaces.PackageRelationships, "Relationship")) continue;
                        if (reader.TryGetAttribute(OfficeXmlNamespaces.None, "TargetMode", out string mode)
                            && string.Equals(mode, "External", StringComparison.OrdinalIgnoreCase)) continue;
                        string? target = OfficeOpcPartNames.ResolveRelationshipTarget(relationshipsPart,
                            reader.TryGetAttribute(OfficeXmlNamespaces.None, "Target", out string value) ? value : null);
                        if (target == null || !package.PartExists(target)) continue;
                        edges.Add((reader.TryGetAttribute(OfficeXmlNamespaces.None, "Id", out string id) ? id : string.Empty, ReachablePartIdentity(target)));
                    }
                }
            } else if (package.PartExists(relationshipsPart)) {
                OfficeXmlTree relationships = package.GetXml(relationshipsPart);
                for (int node = relationships.FindChild(relationships.Root, OfficeXmlNamespaces.PackageRelationships, "Relationship"); node != OfficeXmlTree.NoNode;
                     node = relationships.FindNextSibling(node, OfficeXmlNamespaces.PackageRelationships, "Relationship")) {
                    if (string.Equals(relationships.GetAttribute(node, OfficeXmlNamespaces.None, "TargetMode"), "External", StringComparison.OrdinalIgnoreCase)) continue;
                    string? target = OfficeOpcPartNames.ResolveRelationshipTarget(relationshipsPart, relationships.GetAttribute(node, OfficeXmlNamespaces.None, "Target"));
                    if (target == null || !package.PartExists(target)) continue;
                    edges.Add((relationships.GetAttribute(node, OfficeXmlNamespaces.None, "Id") ?? string.Empty, ReachablePartIdentity(target)));
                }
            }
            _pendingRelationships.Add(source, edges);
            return edges;
        }

        // Identities of the parts reachable from a part, or from the package when it is "", the part itself included,
        // optionally leaving out planned relationship removals. A stopped part is included, but its edges are not followed.
        private HashSet<string> Reachable(OfficeOpcEditablePackage package, string start, string? skipSource, string? skipId,
            string? stopAt = null, Dictionary<string, string>? excludedLinks = null) {
            var reached = new HashSet<string>(OfficeAsciiIgnoreCaseComparer.Instance);
            var pending = new Stack<string>();
            if (start.Length > 0) reached.Add(start);
            pending.Push(start);
            while (pending.Count > 0) {
                string source = pending.Pop();
                bool skipping = skipSource != null && OfficeAsciiIgnoreCaseComparer.Instance.Equals(source, skipSource);
                string? excluded = null;
                if (excludedLinks != null) excludedLinks.TryGetValue(source, out excluded);
                foreach ((string id, string target) in RelationshipsOf(package, source)) {
                    if (skipping && string.Equals(id, skipId, StringComparison.Ordinal) ||
                        excluded != null && string.Equals(id, excluded, StringComparison.Ordinal)) continue;
                    if (reached.Add(target) && (stopAt == null || !OfficeAsciiIgnoreCaseComparer.Instance.Equals(target, stopAt))) {
                        pending.Push(target);
                    }
                }
            }
            return reached;
        }

        private static string ReachablePartIdentity(string partName) {
            if (partName.IndexOf('%') < 0) return partName;
            string identity = OfficeOpcPackage.GetPartIdentity(partName);
            // A second lookup by this identity would decode an escaped percent again.
            if (identity.IndexOf('%') >= 0) throw new InvalidDataException("The slide relationship target contains an escaped percent.");
            return identity;
        }

        // Part names without percent-encoding are already their identity, which spares a copy per lookup.
        private static string PartIdentity(string partName) => partName.IndexOf('%') < 0 ? partName : OfficeOpcPackage.GetPartIdentity(partName);
    }
}
