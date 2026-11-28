using System.Xml;
using OfficeIMO.Core.Internal;
using OfficeIMO.Core.Internal.OpenXml;
using OfficeIMO.PowerPoint.OpenXml;

namespace OfficeIMO.PowerPoint {
    public sealed partial class PowerPointDocument {
        private bool TryClearPackageSlides(bool requireTransitionalTypes = false) {
            if (EditablePackage is not { } package) return false;
            // ClearSlides has historically left caller-held text usable on the detached SDK nodes.
            if (_packageBindings?.Exists(reference => reference.TryGetTarget(out _)) == true ||
                _packageSlides?.Values.Any(slide => slide.HasMaterializedPackageShapes) == true) return false;
            PackageSlideList? slides;
            OfficeXmlTree? relationships = null;
            int[] removedRelationships = Array.Empty<int>();
            string[] removedParts = Array.Empty<string>();
            try {
                if (package.PartNames.Any(name => name.EndsWith(".piece", StringComparison.OrdinalIgnoreCase) &&
                    OfficeOpcPieceAssembler.HasPieceSuffix(name))) return false;
                slides = PackageSlideList.Read(package);
                if (slides == null) return false;
                var excluded = new HashSet<string>(slides.Entries.Select(entry => entry.Relationship), StringComparer.Ordinal);
                if (!TryPlanClearedSlideParts(package, slides.Main, excluded, requireTransitionalTypes, out removedParts)) return false;
                if (slides.Entries.Count > 0) {
                    var removedIdentities = new HashSet<string>(removedParts.Select(PartIdentity), OfficeAsciiIgnoreCaseComparer.Instance);
                    if (slides.Entries.Any(entry => !removedIdentities.Contains(PartIdentity(entry.Part)))) return false;
                    relationships = package.GetRelationships(slides.Main);
                    removedRelationships = slides.Entries.Select(entry => FindRelationship(relationships, entry.Relationship)).ToArray();
                    if (removedRelationships.Any(node => node == OfficeXmlTree.NoNode)) return false;
                }
            } catch (Exception exception) when (exception is InvalidDataException or XmlException or IOException or ArgumentException
                or NotSupportedException or InvalidOperationException or OverflowException) {
                return false;
            }

            foreach ((int node, _, string part) in slides.Entries) {
                slides.Presentation.Remove(node);
                MarkPackageSlideRemoved(part);
            }
            InvalidateRemovedPackageObjects(sdkOpening: false);
            if (relationships != null) {
                foreach (int node in removedRelationships) relationships.Remove(node);
                if (relationships.FindChild(relationships.Root, OfficeXmlNamespaces.PackageRelationships, "Relationship") == OfficeXmlTree.NoNode) {
                    package.DeletePart(OfficeOpcPartNames.GetRelationshipsPartName(slides.Main));
                }
            }
            foreach (string part in removedParts) {
                string edges = OfficeOpcPartNames.GetRelationshipsPartName(part);
                if (package.PartExists(edges)) package.DeletePart(edges);
                package.DeletePart(part);
                PackageSlidesChanged(part);
            }
            PackageSlidesChanged();
            InvalidatePackageRelationships();
            _deferredSlideCount = 0;
            SyncSectionsWithSlides(PowerPointPartRef.Package(this, slides.Main));
            MarkPackageEdited();
            return true;
        }

        private static bool TryPlanClearedSlideParts(OfficeOpcEditablePackage package, string main,
            HashSet<string> excluded, bool requireTransitionalTypes, out string[] removedParts) {
            removedParts = Array.Empty<string>();
            string mainIdentity = PartIdentity(main);
            var linksByOwner = new Dictionary<string, IReadOnlyList<OfficeOpcRelationship>>(OfficeAsciiIgnoreCaseComparer.Instance);
            IReadOnlyList<OfficeOpcRelationship>? ReadRelationships(string? owner) {
                string key = owner == null ? string.Empty : PartIdentity(owner);
                IReadOnlyList<OfficeOpcRelationship>? links = ReadEditablePackageRelationships(package, owner, strictXml: true,
                    requireTransitionalTypes: requireTransitionalTypes);
                if (links == null) return null;
                foreach (OfficeOpcRelationship edge in links) {
                    if (edge.IsExternal) continue;
                    // Data-part lifetimes and physical aliases reached through reference relationships are handled by the SDK.
                    if (IsReferenceRelationship(edge.Type) && (edge.Type != PowerPointPartRef.HyperlinkRelationshipType
                        || edge.TargetPartName is { } target && package.PartExists(target))) return null;
                    if (OfficeAsciiIgnoreCaseComparer.Instance.Equals(key, mainIdentity) && excluded.Contains(edge.Id)
                        && edge.Type != OfficeOpcPackage.SlideRelationshipType) return null;
                }
                linksByOwner.Add(key, links);
                return links;
            }

            var before = new Dictionary<string, string>(OfficeAsciiIgnoreCaseComparer.Instance);
            if (!TryVisitPackageGraph(ReadRelationships, package.GetPartName, package.GetContentType, default,
                edge => before[PartIdentity(edge.Part)] = edge.Part, validateForSave: true)) return false;
            if (excluded.Count == 0) return true;

            var kept = new HashSet<string>(OfficeAsciiIgnoreCaseComparer.Instance);
            if (!TryVisitPackageGraph(owner => {
                    string key = owner == null ? string.Empty : PartIdentity(owner);
                    IReadOnlyList<OfficeOpcRelationship> links = linksByOwner[key];
                    return OfficeAsciiIgnoreCaseComparer.Instance.Equals(key, mainIdentity)
                        ? links.Where(edge => !excluded.Contains(edge.Id)).ToArray() : links;
                }, package.GetPartName, package.GetContentType, default, edge => kept.Add(PartIdentity(edge.Part)))) return false;

            removedParts = before.Where(part => !kept.Contains(part.Key)).Select(part => part.Value).ToArray();
            return true;
        }
    }
}
