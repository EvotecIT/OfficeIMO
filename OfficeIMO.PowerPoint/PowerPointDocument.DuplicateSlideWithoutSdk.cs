using System.Globalization;
using System.Xml;
using OfficeIMO.Core.Internal;
using OfficeIMO.Core.Internal.OpenXml;
using OfficeIMO.PowerPoint.OpenXml;

namespace OfficeIMO.PowerPoint {
    /// <summary>
    /// Duplicating a slide whose relationships the model shares or copies as they are: its layout, linked slides and images
    /// are shared, and external links are copied with their ids. Anything the model clones with the slide, such as notes,
    /// charts or embedded media, is left to the model.
    /// </summary>
    public sealed partial class PowerPointDocument {
        private const string HyperlinkRelationshipType = OfficeXmlNamespaces.TransitionalRelationshipTypePrefix + "hyperlink";

        private PowerPointSlide? TryDuplicateSlideWithoutSdk(int index, int? insertAt) {
            if (GetPackage(edit: true) is not { } package) return null;
            try {
                PackageSlideList? slides = PackageSlideList.Read(package);
                if (slides == null || index < 0 || index >= slides.Entries.Count) return null;
                int targetIndex = insertAt ?? index + 1;
                if (targetIndex < 0 || targetIndex > slides.Entries.Count) return null;
                string source = slides.Entries[index].Part;
                if (!TryReadSharedRelationships(package, source, out var internalRelationships, out var externalRelationships)) return null;

                List<uint>? ids = slides.ReadIds();
                if (ids == null) return null;
                uint maxId = ids.Count == 0 ? 255U : ids.Max();
                if (maxId >= 2147483647U) return null;
                uint newId = maxId >= 255U ? maxId + 1U : 256U;
                string folder = slides.Main.Substring(0, slides.Main.LastIndexOf('/') + 1) + "slides/";
                // The copied relationship targets are relative to the slide's folder, which the duplicate shares.
                if (!source.StartsWith(folder, StringComparison.OrdinalIgnoreCase) || source.IndexOf('/', folder.Length) >= 0) return null;

                OfficeXmlTree presentation = slides.Presentation;
                int p14 = package.Namespaces.GetId(PowerPointTreeSections.P14Namespace);
                var newIds = new List<uint>(ids);
                newIds.Insert(targetIndex, newId);
                OfficeXmlTree mainRelationships = package.GetRelationships(slides.Main);
                string relationshipId = NextPresentationRelationshipId(mainRelationships);
                string partName = NextPartName(package, SlideContentType, folder + "slide", ".xml");

                int list = new PPresentationView(presentation, presentation.Root).GetOrAddSldIdLst().Node;
                int entry = presentation.CreateElement(OfficeXmlNamespaces.PresentationML, "sldId");
                presentation.SetAttribute(entry, OfficeXmlNamespaces.None, "id", newId.ToString(CultureInfo.InvariantCulture));
                presentation.SetAttribute(entry, OfficeXmlNamespaces.Relationships, "id", relationshipId);
                presentation.InsertBefore(list, entry, targetIndex < slides.Entries.Count ? slides.Entries[targetIndex].Node : OfficeXmlTree.NoNode);
                if (!PowerPointTreeSections.TryPlanAdded(presentation, p14, newIds, targetIndex, out PowerPointTreeSections.Plan? sections, out _)) {
                    presentation.Remove(entry);
                    return null;
                }

                package.CreateBinaryPart(partName, SlideContentType, package.ReadPartBytes(source));
                // The model adds the shared parts first, in the source's order, then the external relationships.
                foreach ((string id, string type, string target, bool external) in internalRelationships.Concat(externalRelationships)) {
                    package.AddRelationship(partName, type, target, external, id);
                }
                package.AddRelationship(slides.Main, OfficeXmlNamespaces.TransitionalRelationshipTypePrefix + "slide",
                    OfficeOpcPartNames.GetRelativeTarget(slides.Main, partName), id: relationshipId);
                if (sections != null) PowerPointTreeSections.Apply(sections);
                _pendingRelationships = null;

                var slide = new PowerPointSlide(this, partName);
                (_packageSlides ??= new Dictionary<string, PowerPointSlide>(OfficeAsciiIgnoreCaseComparer.Instance)).Add(PartIdentity(partName), slide);
                PackageSlidesChanged();
                _deferredSlideCount = newIds.Count;
                _pendingEdits = true;
                return slide;
            } catch (Exception exception) when (exception is XmlException or InvalidDataException) {
                return null;
            }
        }

        // The slide's relationships when the model shares every part it relates to: layouts, notes masters, slides and
        // images, followed by external relationships and then hyperlinks, as the model copies them.
        private static bool TryReadSharedRelationships(OfficeOpcEditablePackage package, string slide,
            out List<(string Id, string Type, string Target, bool External)> shared, out List<(string Id, string Type, string Target, bool External)> references) {
            shared = new List<(string, string, string, bool)>();
            references = new List<(string, string, string, bool)>();
            var hyperlinks = new List<(string, string, string, bool)>();
            var sharedParts = new HashSet<string>(OfficeAsciiIgnoreCaseComparer.Instance);
            string relationshipsPart = OfficeOpcPartNames.GetRelationshipsPartName(slide);
            if (!package.PartExists(relationshipsPart)) return true;
            OfficeXmlTree relationships = package.GetXml(relationshipsPart);
            for (int node = relationships.FindChild(relationships.Root, OfficeXmlNamespaces.PackageRelationships, "Relationship"); node != OfficeXmlTree.NoNode;
                 node = relationships.FindNextSibling(node, OfficeXmlNamespaces.PackageRelationships, "Relationship")) {
                string? id = relationships.GetAttribute(node, OfficeXmlNamespaces.None, "Id");
                string? type = relationships.GetAttribute(node, OfficeXmlNamespaces.None, "Type");
                string? target = relationships.GetAttribute(node, OfficeXmlNamespaces.None, "Target");
                if (id == null || type == null || target == null) return false;
                bool external = string.Equals(relationships.GetAttribute(node, OfficeXmlNamespaces.None, "TargetMode"), "External", StringComparison.OrdinalIgnoreCase);
                string normalized = OfficeXmlNamespaces.NormalizeRelationshipType(type);
                if (normalized == HyperlinkRelationshipType) {
                    hyperlinks.Add((id, type, target, external));
                } else if (external) {
                    references.Add((id, type, target, true));
                } else if (normalized is SlideLayoutRelationshipType or ImageRelationshipType
                    || normalized == OfficeXmlNamespaces.TransitionalRelationshipTypePrefix + "slide"
                    || normalized == OfficeXmlNamespaces.TransitionalRelationshipTypePrefix + "notesMaster") {
                    string? resolved = OfficeOpcPartNames.ResolveRelationshipTarget(relationshipsPart, target);
                    // The SDK refuses to relate the copy to one part under two ids, which the model reports.
                    if (resolved == null || !package.PartExists(resolved) || !sharedParts.Add(PartIdentity(resolved))) return false;
                    shared.Add((id, type, target, false));
                } else {
                    // Notes, charts, embedded objects and media are cloned with the slide by the model.
                    return false;
                }
            }
            references.AddRange(hyperlinks);
            return true;
        }
    }
}
