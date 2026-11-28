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

        private PowerPointSlide? TryAddSlideWithoutSdk(int masterIndex, int layoutIndex) =>
            TryAddPackageSlides(1, masterIndex, layoutIndex, useBatchRelationshipIds: false, CancellationToken.None)?.FirstOrDefault();

        // The completed part of a batch stays in the package if cancellation or unsupported section markup
        // prevents the next slide from being added. The caller can finish an unsupported batch through the SDK.
        private List<PowerPointSlide>? TryAddPackageSlides(int count, int masterIndex, int layoutIndex,
            bool useBatchRelationshipIds, CancellationToken cancellationToken) {
            if (GetPackage(edit: true) is not { } package) return null;

            PackageSlideList? slides;
            string layout;
            List<uint>? ids;
            OfficeXmlTree mainRelationships;
            try {
                slides = PackageSlideList.Read(package);
                if (slides == null) return null;
                List<string> masters = RelatedParts(package, slides.Main, SlideMasterRelationshipType);
                if (masterIndex < 0 || masterIndex >= masters.Count) return null;
                List<string> layouts = RelatedParts(package, masters[masterIndex], SlideLayoutRelationshipType);
                if (layoutIndex < 0 || layoutIndex >= layouts.Count) return null;
                layout = layouts[layoutIndex];
                ids = slides.ReadIds();
                if (ids == null) return null;
                mainRelationships = package.GetRelationships(slides.Main);
            } catch (Exception exception) when (exception is XmlException or InvalidDataException) {
                return null;
            }

            uint maxId = ids.Count == 0 ? 255U : Math.Max(255U, ids.Max());
            if (maxId >= 2147483647U || (ulong)maxId + (uint)count > 2147483647U) return null;
            OfficeXmlTree presentation = slides.Presentation;
            if (useBatchRelationshipIds && presentation.HasCommentsOrInstructions()) return null;
            int p14 = package.Namespaces.GetId(PowerPointTreeSections.P14Namespace);
            HashSet<string> used = ReadPackageRelationshipIds(mainRelationships);
            if (useBatchRelationshipIds) {
                foreach ((_, string id, _) in slides.Entries) used.Add(id);
            }
            long relationshipCursor = useBatchRelationshipIds ? 1 : _slideRelationshipSearchStart;
            string pathPrefix = slides.Main.Substring(0, slides.Main.LastIndexOf('/') + 1) + "slides/slide";
            // The SDK starts with the slide parts present when it opens the package.
            int batchPartSequence = useBatchRelationshipIds ? package.CountParts(SlideContentType, includeDeleted: false) : 0;
            var added = new List<PowerPointSlide>(count);
            try {
                for (int index = 0; index < count; index++) {
                    cancellationToken.ThrowIfCancellationRequested();
                    string partName = useBatchRelationshipIds
                        ? NextFreePartName(package, pathPrefix, ".xml", ref batchPartSequence)
                        : NextPartName(package, SlideContentType, pathPrefix, ".xml");
                    string partRelationships = OfficeOpcPartNames.GetRelationshipsPartName(partName);
                    if (package.PartExists(partRelationships)) {
                        throw new InvalidDataException("The package already has relationships for the new slide '" + partName + "'.");
                    }

                    uint id = maxId + (uint)index + 1U;
                    ids.Add(id);
                    if (!PowerPointTreeSections.TryPlanAdded(presentation, p14, ids, ids.Count - 1, out PowerPointTreeSections.Plan? sections, out int sectionEntry)) {
                        ids.RemoveAt(ids.Count - 1);
                        break;
                    }

                    string relationship = ReserveSlideRelationshipId(used, ref relationshipCursor);
                    string layoutTarget = useBatchRelationshipIds ? layout : OfficeOpcPartNames.GetRelativeTarget(partName, layout);
                    string slideTarget = useBatchRelationshipIds ? partName : OfficeOpcPartNames.GetRelativeTarget(slides.Main, partName);
                    bool createdSlide = false;
                    bool linkedSlide = false;
                    int entry = OfficeXmlTree.NoNode;
                    try {
                        OfficeXmlTree slideMarkup = package.CreateXmlPart(partName, SlideContentType, OfficeXmlNamespaces.PresentationML, "sld", 24);
                        createdSlide = true;
                        WriteBlankSlide(slideMarkup);
                        package.AddRelationship(partName, SlideLayoutRelationshipType, layoutTarget, id: "rId1");
                        package.AddRelationship(slides.Main, PowerPointPartRef.SlideRelationshipType, slideTarget, id: relationship);
                        linkedSlide = true;

                        int list = new PPresentationView(presentation, presentation.Root).GetOrAddSldIdLst().Node;
                        entry = Append(presentation, list, OfficeXmlNamespaces.PresentationML, "sldId");
                        presentation.SetAttribute(entry, OfficeXmlNamespaces.None, "id", id.ToString(CultureInfo.InvariantCulture));
                        presentation.SetAttribute(entry, OfficeXmlNamespaces.Relationships, "id", relationship);
                    } catch {
                        if (entry != OfficeXmlTree.NoNode) presentation.Remove(entry);
                        if (sectionEntry != OfficeXmlTree.NoNode) presentation.Remove(sectionEntry);
                        if (linkedSlide) mainRelationships.Remove(FindRelationship(mainRelationships, relationship));
                        if (createdSlide) {
                            package.DeletePart(partRelationships);
                            package.DeletePart(partName);
                        }
                        throw;
                    }
                    if (sections != null) PowerPointTreeSections.Apply(sections);

                    var slide = new PowerPointSlide(this, partName);
                    (_packageSlides ??= new Dictionary<string, PowerPointSlide>(OfficeAsciiIgnoreCaseComparer.Instance)).Add(PartIdentity(partName), slide);
                    added.Add(slide);
                }
            } finally {
                if (added.Count > 0) {
                    if (!useBatchRelationshipIds) _slideRelationshipSearchStart = relationshipCursor;
                    InvalidatePackageRelationships();
                    PackageSlidesChanged();
                    _deferredSlideCount = slides.Entries.Count + added.Count;
                    MarkPackageEdited();
                }
            }
            return added;
        }

        /// <summary>
        /// The layouts of a master read from the package, in relationship order as the SDK lists them, with each one's
        /// name, type token and relationship id; null when the model has to read them, or reports a missing master.
        /// </summary>
        private List<(string Name, string? Type, string? RelationshipId)>? TryReadPackageLayouts(int masterIndex) {
            if (GetPackage(edit: false) is not { } package) return null;
            try {
                PackageSlideList? slides = PackageSlideList.Read(package);
                if (slides == null) return null;
                List<string> masters = RelatedParts(package, slides.Main, SlideMasterRelationshipType);
                if (masterIndex < 0 || masterIndex >= masters.Count) return null;
                string master = masters[masterIndex];
                string masterRelationships = OfficeOpcPartNames.GetRelationshipsPartName(master);
                var layouts = new List<(string, string?, string?)>();
                foreach (string layout in RelatedParts(package, master, SlideLayoutRelationshipType)) {
                    OfficeXmlTree tree = package.GetXml(layout);
                    if (!tree.IsElement(tree.Root, OfficeXmlNamespaces.PresentationML, "sldLayout")) return null;
                    int data = tree.FindChild(tree.Root, OfficeXmlNamespaces.PresentationML, "cSld");
                    string name = data == OfficeXmlTree.NoNode ? string.Empty : tree.GetAttribute(data, OfficeXmlNamespaces.None, "name") ?? string.Empty;
                    string? type = tree.GetAttribute(tree.Root, OfficeXmlNamespaces.None, "type");
                    layouts.Add((name, type, FirstRelationshipIdTo(package.GetXml(masterRelationships), masterRelationships, layout)));
                }
                return layouts;
            } catch (Exception exception) when (exception is XmlException or InvalidDataException) {
                return null;
            }
        }

        private static string? FirstRelationshipIdTo(OfficeXmlTree relationships, string relationshipsPart, string target) {
            for (int node = relationships.FindChild(relationships.Root, OfficeXmlNamespaces.PackageRelationships, "Relationship"); node != OfficeXmlTree.NoNode;
                 node = relationships.FindNextSibling(node, OfficeXmlNamespaces.PackageRelationships, "Relationship")) {
                if (string.Equals(relationships.GetAttribute(node, OfficeXmlNamespaces.None, "TargetMode"), "External", StringComparison.OrdinalIgnoreCase)) continue;
                string? resolved = OfficeOpcPartNames.ResolveRelationshipTarget(relationshipsPart, relationships.GetAttribute(node, OfficeXmlNamespaces.None, "Target"));
                if (resolved != null && OfficeAsciiIgnoreCaseComparer.Instance.Equals(PartIdentity(resolved), PartIdentity(target))) {
                    return relationships.GetAttribute(node, OfficeXmlNamespaces.None, "Id");
                }
            }
            return null;
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

        // The next free rIdN of the presentation part, searched from where the last search stopped.
        private string NextPresentationRelationshipId(OfficeXmlTree relationships) =>
            ReserveSlideRelationshipId(ReadPackageRelationshipIds(relationships), ref _slideRelationshipSearchStart);

        private static HashSet<string> ReadPackageRelationshipIds(OfficeXmlTree relationships) {
            var used = new HashSet<string>(StringComparer.Ordinal);
            for (int node = relationships.FindChild(relationships.Root, OfficeXmlNamespaces.PackageRelationships, "Relationship"); node != OfficeXmlTree.NoNode;
                 node = relationships.FindNextSibling(node, OfficeXmlNamespaces.PackageRelationships, "Relationship")) {
                if (relationships.GetAttribute(node, OfficeXmlNamespaces.None, "Id") is { } id) used.Add(id);
            }
            return used;
        }

        // The SDK's name for a new part. It numbers the parts of a content type on from how many the package had when it
        // opened, skipping names in use; a deleted part's name is free again.
        private string NextPartName(OfficeOpcEditablePackage package, string contentType, string prefix, string extension) {
            _packagePartSequence ??= new Dictionary<string, int>(StringComparer.OrdinalIgnoreCase);
            if (!_packagePartSequence.TryGetValue(contentType, out int sequence)) sequence = package.CountParts(contentType);
            string name = NextFreePartName(package, prefix, extension, ref sequence);
            _packagePartSequence[contentType] = sequence;
            return name;
        }

        private static string NextFreePartName(OfficeOpcEditablePackage package, string prefix, string extension, ref int sequence) {
            string name;
            do {
                sequence++;
                name = prefix + sequence.ToString(CultureInfo.InvariantCulture) + extension;
            } while (package.PartExists(name));
            return name;
        }

        // The SDK's name for a part created with a requested name: that name while the package has no part of the content
        // type, otherwise the name with the content type's next free number appended. Either way the number moves on.
        private string UniquePartName(OfficeOpcEditablePackage package, string contentType, string requested) {
            _packagePartSequence ??= new Dictionary<string, int>(StringComparer.OrdinalIgnoreCase);
            if (!_packagePartSequence.TryGetValue(contentType, out int sequence)) sequence = package.CountParts(contentType);
            if (sequence == 0 && !package.PartExists(requested)) {
                _packagePartSequence[contentType] = 1;
                return requested;
            }
            int dot = requested.LastIndexOf('.');
            string stem = dot > requested.LastIndexOf('/') ? requested.Substring(0, dot) : requested;
            string extension = dot > requested.LastIndexOf('/') ? requested.Substring(dot) : string.Empty;
            string name;
            do {
                sequence++;
                name = stem + sequence.ToString(CultureInfo.InvariantCulture) + extension;
            } while (package.PartExists(name));
            _packagePartSequence[contentType] = sequence;
            return name;
        }
    }
}
