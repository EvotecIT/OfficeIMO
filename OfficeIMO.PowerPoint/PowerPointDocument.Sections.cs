using OfficeIMO.Core.Internal.OpenXml;
using OfficeIMO.OpenXml.Internal;
using OfficeIMO.PowerPoint.OpenXml;

namespace OfficeIMO.PowerPoint {
    public sealed partial class PowerPointDocument {
        /// <summary>
        ///     Returns the sections defined in the presentation.
        /// </summary>
        public IReadOnlyList<PowerPointSectionInfo> GetSections() {
            ThrowIfDisposed();
            OfficeXmlNode root = ReadablePackage is { } package &&
                FirstRelated(package, null, OfficeOpcPackage.OfficeDocumentRelationshipType, out _) is { } main
                ? PowerPointPartRef.Package(this, main).Root : PresentationPartRef.Root;
            OfficeXmlNode sectionList = GetSectionList(root, create: false);
            if (sectionList.IsNone) {
                return Array.Empty<PowerPointSectionInfo>();
            }

            List<OfficeXmlNode> slideIds = SlidesForSections(root);
            Dictionary<uint, int> slideIndexMap = BuildSlideIndexMap(slideIds);

            return SectionChildren(sectionList, "section").Select(section => BuildSectionInfo(section, slideIndexMap)).ToList();
        }

        /// <summary>
        ///     Adds a new section starting at the specified slide index.
        /// </summary>
        public PowerPointSectionInfo AddSection(string name, int startSlideIndex) {
            ThrowIfDisposed();
            ArgumentException.ThrowIfNullOrWhiteSpace(name);

            PowerPointPartRef presentation = PresentationPartRef.ForMovingXmlNodes();
            OfficeXmlNode slideIdList = presentation.Root.Child(SectionPresentationNamespace, "sldIdLst");
            if (slideIdList.IsNone) {
                throw new InvalidOperationException("Presentation has no slides.");
            }

            List<OfficeXmlNode> slideIds = slideIdList.Children(SectionPresentationNamespace, "sldId").ToList();
            if (slideIds.Count == 0) {
                throw new InvalidOperationException("Presentation has no slides.");
            }
            ArgumentOutOfRangeException.ThrowIfNegative(startSlideIndex);
            ArgumentOutOfRangeException.ThrowIfGreaterThanOrEqual(startSlideIndex, slideIds.Count);

            uint slideIdValue = slideIds[startSlideIndex].UInt32Attribute("id") ?? throw new InvalidOperationException("Slide ID is missing.");
            OfficeXmlNode sectionList = EnsureSectionList(presentation.Root, slideIds);
            presentation.MarkEdited();

            OfficeXmlNode containing = FindSectionBySlideId(sectionList, slideIdValue);
            if (containing.IsNone) {
                OfficeXmlNode fallback = SectionChildren(sectionList, "section").Last();
                OfficeXmlNode members = EnsureSectionSlideIdList(fallback);
                members.Append(CreateSectionEntry(members, slideIdValue));
                return BuildSectionInfo(fallback, BuildSlideIndexMap(slideIds));
            }

            OfficeXmlNode list = EnsureSectionSlideIdList(containing);
            List<OfficeXmlNode> entries = SectionChildren(list, "sldId").ToList();
            int entryIndex = entries.FindIndex(entry => entry.UInt32Attribute("id") == slideIdValue);
            if (entryIndex <= 0) {
                containing.SetAttribute("name", name);
                return BuildSectionInfo(containing, BuildSlideIndexMap(slideIds));
            }

            List<uint> movedIds = entries
                .Skip(entryIndex)
                .Select(entry => entry.UInt32Attribute("id"))
                .Where(id => id != null)
                .Select(id => id!.Value)
                .ToList();
            foreach (OfficeXmlNode entry in entries.Skip(entryIndex)) {
                entry.Remove();
            }

            OfficeXmlNode newSection = CreateSection(sectionList, name, movedIds);
            newSection = sectionList.InsertBefore(newSection, containing.NextNode());
            return BuildSectionInfo(newSection, BuildSlideIndexMap(slideIds));
        }

        /// <summary>
        ///     Renames the first section matching the provided name.
        /// </summary>
        public bool RenameSection(string name, string newName, bool ignoreCase = true) {
            ThrowIfDisposed();
            ArgumentNullException.ThrowIfNull(name);
            ArgumentException.ThrowIfNullOrWhiteSpace(newName);

            PowerPointPartRef presentation = PresentationPartRef;
            OfficeXmlNode sectionList = GetSectionList(presentation.Root, create: false);
            if (sectionList.IsNone) {
                return false;
            }

            StringComparison comparison = ignoreCase ? StringComparison.OrdinalIgnoreCase : StringComparison.Ordinal;
            foreach (OfficeXmlNode section in SectionChildren(sectionList, "section")) {
                string currentName = section.Attribute("name") ?? string.Empty;
                if (string.Equals(currentName, name, comparison)) {
                    presentation.MarkEdited();
                    section.SetAttribute("name", newName);
                    return true;
                }
            }

            return false;
        }

        private const int SectionPresentationNamespace = OfficeXmlNamespaces.PresentationML;

        private static List<OfficeXmlNode> SlidesForSections(OfficeXmlNode presentation) =>
            presentation.Child(SectionPresentationNamespace, "sldIdLst")
                .Children(SectionPresentationNamespace, "sldId").ToList();

        private static OfficeXmlNode GetSectionList(OfficeXmlNode root, bool create) {
            OfficeXmlNode extensions = root.Child(SectionPresentationNamespace, "extLst");
            if (extensions.IsNone && create) {
                extensions = root.Append(root.CreateElement(SectionPresentationNamespace, "extLst"));
            }

            OfficeXmlNode extension = GetSectionExtension(extensions);
            if (extension.IsNone && create) {
                extension = extensions.CreateElement(SectionPresentationNamespace, "ext");
                extension.SetAttribute("uri", PowerPointTreeSections.SectionListUri);
                extension = extensions.Append(extension);
            }

            OfficeXmlNode sections = extension.Child(PowerPointTreeSections.P14Namespace, "sectionLst");
            if (sections.IsNone && create) {
                sections = CreateSectionNode(extension, "sectionLst");
                sections.DeclareNamespace("p14", PowerPointTreeSections.P14Namespace);
                sections = extension.Append(sections);
            }
            return sections;
        }

        private static OfficeXmlNode GetSectionExtension(OfficeXmlNode extensions) {
            foreach (OfficeXmlNode extension in extensions.Children(SectionPresentationNamespace, "ext")) {
                if (string.Equals(extension.Attribute("uri"), PowerPointTreeSections.SectionListUri, StringComparison.Ordinal)) {
                    return extension;
                }
            }

            return default;
        }

        private static IEnumerable<OfficeXmlNode> SectionChildren(OfficeXmlNode parent, string localName) {
            foreach (OfficeXmlNode child in parent.Children()) {
                if (child.Is(PowerPointTreeSections.P14Namespace, localName)) yield return child;
            }
        }

        private static OfficeXmlNode CreateSectionNode(OfficeXmlNode parent, string localName) =>
            parent.CreateElement(PowerPointTreeSections.P14Namespace, localName, "p14");

        private static OfficeXmlNode EnsureSectionList(OfficeXmlNode root, IReadOnlyList<OfficeXmlNode> slideIds) {
            OfficeXmlNode sectionList = GetSectionList(root, create: true);
            if (slideIds.Count > 0 && !SectionChildren(sectionList, "section").Any()) {
                List<uint> ids = slideIds
                    .Select(id => id.UInt32Attribute("id"))
                    .Where(id => id != null)
                    .Select(id => id!.Value)
                    .ToList();
                sectionList.Append(CreateSection(sectionList, DefaultSectionName, ids));
            }

            EnsureSectionCoverage(sectionList, slideIds);
            return sectionList;
        }

        private static void EnsureSectionCoverage(OfficeXmlNode sectionList, IReadOnlyList<OfficeXmlNode> slideIds) {
            Dictionary<uint, int> slideIndexMap = BuildSlideIndexMap(slideIds);
            HashSet<uint> assigned = new();
            List<OfficeXmlNode> orderedSections = new();

            foreach (OfficeXmlNode section in SectionChildren(sectionList, "section").ToList()) {
                OfficeXmlNode list = EnsureSectionSlideIdList(section);
                // An explicitly empty section is a boundary awaiting its first slide. A section
                // whose former members were removed is different and can be pruned below.
                bool wasEmpty = !SectionChildren(list, "sldId").Any();
                List<uint> sectionSlideIds = SectionChildren(list, "sldId")
                    .Select(entry => entry.UInt32Attribute("id"))
                    .Where(id => id != null && slideIndexMap.ContainsKey(id.Value))
                    .Select(id => id!.Value)
                    .OrderBy(id => slideIndexMap[id])
                    .ToList();

                list.RemoveAllChildren();
                foreach (uint slideId in sectionSlideIds) {
                    if (!assigned.Add(slideId)) {
                        continue;
                    }

                    list.Append(CreateSectionEntry(list, slideId));
                }

                if (wasEmpty || SectionChildren(list, "sldId").Any()) {
                    orderedSections.Add(section);
                } else {
                    section.Remove();
                }
            }

            if (orderedSections.Count == 0) {
                if (slideIds.Count == 0) {
                    return;
                }

                OfficeXmlNode defaultSection = CreateSection(sectionList, DefaultSectionName, slideIds
                    .Select(id => id.UInt32Attribute("id"))
                    .Where(id => id != null)
                    .Select(id => id!.Value)
                    .ToList());
                sectionList.RemoveAllChildren();
                sectionList.Append(defaultSection);
                return;
            }

            List<uint> unassigned = slideIds.Select(id => id.UInt32Attribute("id"))
                .Where(id => id != null).Select(id => id!.Value)
                .Where(id => !assigned.Contains(id)).ToList();
            if (unassigned.Count > 0) {
                OfficeXmlNode targetSection = orderedSections.LastOrDefault(section =>
                    SectionChildren(EnsureSectionSlideIdList(section), "sldId").Any());
                if (targetSection.IsNone) orderedSections.Insert(0, CreateSection(sectionList, DefaultSectionName, unassigned));
                else {
                    OfficeXmlNode target = EnsureSectionSlideIdList(targetSection);
                    foreach (uint slideId in unassigned) target.Append(CreateSectionEntry(target, slideId));
                }
            }

            // An unused section stays before the following section, rather than becoming the
            // trailing active boundary again when that following section gains its first slide.
            var sectionOrder = new Dictionary<OfficeXmlNode, int>();
            int followingStart = int.MaxValue;
            for (int index = orderedSections.Count - 1; index >= 0; index--) {
                int start = GetSectionStartIndex(orderedSections[index], slideIndexMap);
                if (start != int.MaxValue) followingStart = start;
                sectionOrder.Add(orderedSections[index], followingStart);
            }
            orderedSections = orderedSections.OrderBy(section => sectionOrder[section]).ToList();

            sectionList.RemoveAllChildren();
            foreach (OfficeXmlNode section in orderedSections) {
                sectionList.Append(section);
            }
        }

        private static OfficeXmlNode CreateSection(OfficeXmlNode parent, string name, IReadOnlyList<uint> slideIds) {
            OfficeXmlNode section = CreateSectionNode(parent, "section");
            section.SetAttribute("id", Guid.NewGuid().ToString("B").ToUpperInvariant());
            section.SetAttribute("name", name);
            OfficeXmlNode members = section.Append(CreateSectionNode(section, "sldIdLst"));
            foreach (uint slideId in slideIds) {
                members.Append(CreateSectionEntry(members, slideId));
            }
            return section;
        }

        private static OfficeXmlNode CreateSectionEntry(OfficeXmlNode parent, uint id) {
            OfficeXmlNode entry = CreateSectionNode(parent, "sldId");
            entry.SetUInt32Attribute("id", id);
            return entry;
        }

        private static OfficeXmlNode EnsureSectionSlideIdList(OfficeXmlNode section) {
            OfficeXmlNode members = section.Child(PowerPointTreeSections.P14Namespace, "sldIdLst");
            return members.IsNone ? section.Append(CreateSectionNode(section, "sldIdLst")) : members;
        }

        private static Dictionary<uint, int> BuildSlideIndexMap(IReadOnlyList<OfficeXmlNode> slideIds) {
            Dictionary<uint, int> map = new();
            for (int i = 0; i < slideIds.Count; i++) {
                uint? id = slideIds[i].UInt32Attribute("id");
                if (id != null) {
                    map[id.Value] = i;
                }
            }
            return map;
        }

        private static OfficeXmlNode FindSectionBySlideId(OfficeXmlNode sectionList, uint slideId) {
            foreach (OfficeXmlNode section in SectionChildren(sectionList, "section")) {
                OfficeXmlNode list = section.Child(PowerPointTreeSections.P14Namespace, "sldIdLst");
                if (list.IsNone) {
                    continue;
                }

                if (SectionChildren(list, "sldId").Any(entry => entry.UInt32Attribute("id") == slideId)) {
                    return section;
                }
            }

            return default;
        }

        private static int GetSectionStartIndex(OfficeXmlNode section, IReadOnlyDictionary<uint, int> slideIndexMap) {
            OfficeXmlNode list = section.Child(PowerPointTreeSections.P14Namespace, "sldIdLst");
            if (list.IsNone) {
                return int.MaxValue;
            }

            return SectionChildren(list, "sldId")
                .Select(entry => entry.UInt32Attribute("id"))
                .Where(id => id != null && slideIndexMap.ContainsKey(id.Value))
                .Select(id => slideIndexMap[id!.Value])
                .DefaultIfEmpty(int.MaxValue)
                .Min();
        }

        private void AssignSlideToNearestSection(uint slideId, int slideIndex) {
            PowerPointPartRef presentation = PresentationPartRef;
            OfficeXmlNode sectionList = GetSectionList(presentation.Root, create: false);
            if (sectionList.IsNone) {
                return;
            }

            List<OfficeXmlNode> slideIds = SlidesForSections(presentation.Root);
            if (slideIds.Count == 0) {
                return;
            }

            presentation.MarkEdited();
            if (FindSectionBySlideId(sectionList, slideId).Exists) {
                EnsureSectionCoverage(sectionList, slideIds);
                return;
            }

            OfficeXmlNode targetSection = default;
            // A boundary declared between slides takes precedence over the preceding slide's section.
            OfficeXmlNode lastSection = SectionChildren(sectionList, "section").LastOrDefault();
            if (slideIndex == slideIds.Count - 1 && lastSection.Exists &&
                !SectionChildren(EnsureSectionSlideIdList(lastSection), "sldId").Any()) {
                targetSection = lastSection;
            }
            if (targetSection.IsNone && slideIndex > 0) {
                uint? previousSlideId = slideIds[slideIndex - 1].UInt32Attribute("id");
                if (previousSlideId != null) {
                    targetSection = FindSectionBySlideId(sectionList, previousSlideId.Value);
                }
            }

            if (targetSection.IsNone && slideIndex + 1 < slideIds.Count) {
                uint? nextSlideId = slideIds[slideIndex + 1].UInt32Attribute("id");
                if (nextSlideId != null) {
                    targetSection = FindSectionBySlideId(sectionList, nextSlideId.Value);
                }
            }

            if (targetSection.IsNone) targetSection = SectionChildren(sectionList, "section").LastOrDefault();
            if (targetSection.IsNone) {
                sectionList.Append(CreateSection(sectionList, DefaultSectionName, new[] { slideId }));
                EnsureSectionCoverage(sectionList, slideIds);
                return;
            }

            OfficeXmlNode members = EnsureSectionSlideIdList(targetSection);
            members.Append(CreateSectionEntry(members, slideId));
            EnsureSectionCoverage(sectionList, slideIds);
        }

        private void SyncSectionsWithSlides() => SyncSectionsWithSlides(PresentationPartRef);

        private static void SyncSectionsWithSlides(PowerPointPartRef presentation) {
            OfficeXmlNode sectionList = GetSectionList(presentation.Root, create: false);
            if (sectionList.IsNone) {
                return;
            }

            presentation.MarkEdited();
            List<OfficeXmlNode> slideIds = SlidesForSections(presentation.Root);
            if (slideIds.Count == 0) {
                EnsureSectionCoverage(sectionList, slideIds);
                if (SectionChildren(sectionList, "section").Any()) return;
                OfficeXmlNode sectionExtension = sectionList.Parent;
                OfficeXmlNode extensionList = sectionExtension.Parent;
                sectionExtension.Remove();
                if (extensionList.Exists && extensionList.Child(SectionPresentationNamespace, "ext").IsNone) {
                    extensionList.Remove();
                }
                return;
            }

            EnsureSectionCoverage(sectionList, slideIds);
        }

        private static PowerPointSectionInfo BuildSectionInfo(OfficeXmlNode section, IReadOnlyDictionary<uint, int> slideIndexMap) {
            List<int> indices = new();
            OfficeXmlNode list = section.Child(PowerPointTreeSections.P14Namespace, "sldIdLst");
            if (list.Exists) {
                foreach (OfficeXmlNode entry in SectionChildren(list, "sldId")) {
                    uint? id = entry.UInt32Attribute("id");
                    if (id != null && slideIndexMap.TryGetValue(id.Value, out int index)) {
                        indices.Add(index);
                    }
                }
            }

            indices.Sort();
            return new PowerPointSectionInfo(section.Attribute("name") ?? string.Empty, section.Attribute("id") ?? string.Empty, indices);
        }

    }
}
