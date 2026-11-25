using System.Globalization;
using System.Xml;
using OfficeIMO.Core.Internal;
using OfficeIMO.Core.Internal.OpenXml;
using OfficeIMO.PowerPoint.OpenXml;

namespace OfficeIMO.PowerPoint {
    /// <summary>
    /// Adding slides to a presentation whose SDK package is not open writes the slide part, its layout relationship, the
    /// slide list entry and its section in the document's package, as <see cref="AddSlide(int, int)"/> does through the SDK.
    /// </summary>
    public sealed partial class PowerPointDocument {
        private const string SlideMasterRelationshipType = OfficeXmlNamespaces.TransitionalRelationshipTypePrefix + "slideMaster";
        private const string SlideLayoutRelationshipType = OfficeXmlNamespaces.TransitionalRelationshipTypePrefix + "slideLayout";

        // The last number given to a new part of each content type, as the SDK counts them in an open package.
        private Dictionary<string, int>? _packagePartSequence;

        /// <summary>Adds a slide without the SDK, or returns null, having changed nothing, for the model to add it.</summary>
        private PowerPointSlide? TryAddSlideWithoutSdk(int masterIndex, int layoutIndex) {
            if (GetPackage(edit: true) is not { } package) return null;
            try {
                PackageSlideList? slides = PackageSlideList.Read(package);
                if (slides == null) return null;
                // Masters and their layouts in relationship order, as the SDK lists the parts.
                List<string> masters = RelatedParts(package, slides.Main, SlideMasterRelationshipType);
                if (masterIndex < 0 || masterIndex >= masters.Count) return null;
                List<string> layouts = RelatedParts(package, masters[masterIndex], SlideLayoutRelationshipType);
                if (layoutIndex < 0 || layoutIndex >= layouts.Count) return null;
                string layout = layouts[layoutIndex];

                List<uint>? ids = slides.ReadIds();
                if (ids == null) return null;
                uint maxId = ids.Count == 0 ? 255U : ids.Max();
                if (maxId >= 2147483647U) return null;
                uint newId = maxId >= 255U ? maxId + 1U : 256U;

                OfficeXmlTree presentation = slides.Presentation;
                int p14 = package.Namespaces.GetId(PowerPointTreeSections.P14Namespace);
                var newIds = new List<uint>(ids) { newId };
                OfficeXmlTree mainRelationships = package.GetRelationships(slides.Main);
                string relationshipId = NextPresentationRelationshipId(mainRelationships);
                string partName = NextPartName(package, SlideContentType,
                    slides.Main.Substring(0, slides.Main.LastIndexOf('/') + 1) + "slides/slide", ".xml");

                // The slide list entry comes first, so the section plan sees the list as the model does.
                PPresentationView root = new(presentation, presentation.Root);
                int list = root.GetOrAddSldIdLst().Node;
                int entry = presentation.CreateElement(OfficeXmlNamespaces.PresentationML, "sldId");
                presentation.SetAttribute(entry, OfficeXmlNamespaces.None, "id", newId.ToString(CultureInfo.InvariantCulture));
                presentation.SetAttribute(entry, OfficeXmlNamespaces.Relationships, "id", relationshipId);
                presentation.Append(list, entry);
                if (!PowerPointTreeSections.TryPlanAdded(presentation, p14, newIds, newIds.Count - 1, out PowerPointTreeSections.Plan? sections, out _)) {
                    presentation.Remove(entry);
                    return null;
                }

                WriteBlankSlide(package.CreateXmlPart(partName, SlideContentType, OfficeXmlNamespaces.PresentationML, "sld", 24));
                package.AddRelationship(partName, SlideLayoutRelationshipType, OfficeOpcPartNames.GetRelativeTarget(partName, layout), id: "rId1");
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

        // The slide the model writes for a new slide: an empty shape tree and the master's colour mapping.
        private static void WriteBlankSlide(OfficeXmlTree slide) {
            int root = slide.Root;
            slide.DeclareNamespace(root, "a", "http://schemas.openxmlformats.org/drawingml/2006/main");
            slide.DeclareNamespace(root, "r", "http://schemas.openxmlformats.org/officeDocument/2006/relationships");
            slide.DeclareNamespace(root, "p", "http://schemas.openxmlformats.org/presentationml/2006/main");
            int tree = Append(slide, Append(slide, root, OfficeXmlNamespaces.PresentationML, "cSld"), OfficeXmlNamespaces.PresentationML, "spTree");
            int nonVisual = Append(slide, tree, OfficeXmlNamespaces.PresentationML, "nvGrpSpPr");
            int properties = Append(slide, nonVisual, OfficeXmlNamespaces.PresentationML, "cNvPr");
            slide.SetAttribute(properties, OfficeXmlNamespaces.None, "id", "1");
            slide.SetAttribute(properties, OfficeXmlNamespaces.None, "name", string.Empty);
            Append(slide, nonVisual, OfficeXmlNamespaces.PresentationML, "cNvGrpSpPr");
            Append(slide, nonVisual, OfficeXmlNamespaces.PresentationML, "nvPr");
            int transform = Append(slide, Append(slide, tree, OfficeXmlNamespaces.PresentationML, "grpSpPr"), OfficeXmlNamespaces.DrawingML, "xfrm");
            AppendPoint(slide, transform, "off", "x", "y");
            AppendPoint(slide, transform, "ext", "cx", "cy");
            AppendPoint(slide, transform, "chOff", "x", "y");
            AppendPoint(slide, transform, "chExt", "cx", "cy");
            Append(slide, Append(slide, root, OfficeXmlNamespaces.PresentationML, "clrMapOvr"), OfficeXmlNamespaces.DrawingML, "masterClrMapping");
        }

        private static void AppendPoint(OfficeXmlTree tree, int parent, string name, string first, string second) {
            int point = Append(tree, parent, OfficeXmlNamespaces.DrawingML, name);
            tree.SetAttribute(point, OfficeXmlNamespaces.None, first, "0");
            tree.SetAttribute(point, OfficeXmlNamespaces.None, second, "0");
        }

        internal static int Append(OfficeXmlTree tree, int parent, int namespaceId, string localName) {
            int element = tree.CreateElement(namespaceId, localName);
            tree.Append(parent, element);
            return element;
        }

        // Internal relationship targets of a part with a type, in relationship order, each as often as it is listed.
        private static List<string> RelatedParts(OfficeOpcEditablePackage package, string source, string type) {
            var parts = new List<string>();
            string relationshipsPart = OfficeOpcPartNames.GetRelationshipsPartName(source);
            if (!package.PartExists(relationshipsPart)) return parts;
            OfficeXmlTree relationships = package.GetXml(relationshipsPart);
            for (int node = relationships.FindChild(relationships.Root, OfficeXmlNamespaces.PackageRelationships, "Relationship"); node != OfficeXmlTree.NoNode;
                 node = relationships.FindNextSibling(node, OfficeXmlNamespaces.PackageRelationships, "Relationship")) {
                if (string.Equals(relationships.GetAttribute(node, OfficeXmlNamespaces.None, "TargetMode"), "External", StringComparison.OrdinalIgnoreCase)) continue;
                if (OfficeXmlNamespaces.NormalizeRelationshipType(relationships.GetAttribute(node, OfficeXmlNamespaces.None, "Type") ?? string.Empty) != type) continue;
                string? target = OfficeOpcPartNames.ResolveRelationshipTarget(relationshipsPart, relationships.GetAttribute(node, OfficeXmlNamespaces.None, "Target"));
                if (target != null && package.PartExists(target)) parts.Add(target);
            }
            return parts;
        }

        // The next free rIdN of the presentation part, searched from where the last search stopped, as the model does.
        private string NextPresentationRelationshipId(OfficeXmlTree relationships) {
            var used = new HashSet<string>(StringComparer.Ordinal);
            for (int node = relationships.FindChild(relationships.Root, OfficeXmlNamespaces.PackageRelationships, "Relationship"); node != OfficeXmlTree.NoNode;
                 node = relationships.FindNextSibling(node, OfficeXmlNamespaces.PackageRelationships, "Relationship")) {
                if (relationships.GetAttribute(node, OfficeXmlNamespaces.None, "Id") is { } id) used.Add(id);
            }
            long next = _slideRelationshipSearchStart;
            while (true) {
                string candidate = "rId" + next.ToString(CultureInfo.InvariantCulture);
                next = checked(next + 1);
                if (used.Contains(candidate)) continue;
                _slideRelationshipSearchStart = next;
                return candidate;
            }
        }

        // The SDK's name for a new part. It numbers the parts of a content type on from how many the package had when it
        // opened, skipping names in use; a deleted part's name is free again.
        private string NextPartName(OfficeOpcEditablePackage package, string contentType, string prefix, string extension) {
            _packagePartSequence ??= new Dictionary<string, int>(StringComparer.OrdinalIgnoreCase);
            if (!_packagePartSequence.TryGetValue(contentType, out int sequence)) sequence = package.CountParts(contentType);
            string name;
            do {
                sequence++;
                name = prefix + sequence.ToString(CultureInfo.InvariantCulture) + extension;
            } while (package.PartExists(name));
            _packagePartSequence[contentType] = sequence;
            return name;
        }
    }
}
