using System;
using System.IO;
using System.Reflection;
using System.Runtime.ExceptionServices;
using System.Xml.Linq;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Presentation;
using DocumentFormat.OpenXml.Validation;
using OfficeIMO.Shared;
using A = DocumentFormat.OpenXml.Drawing;
using P14 = DocumentFormat.OpenXml.Office2010.PowerPoint;

namespace OfficeIMO.PowerPoint {
    public sealed partial class PowerPointPresentation {
        /// <summary>
        ///     Returns the sections defined in the presentation.
        /// </summary>
        public IReadOnlyList<PowerPointSectionInfo> GetSections() {
            ThrowIfDisposed();
            P14.SectionList? sectionList = GetSectionList(create: false);
            if (sectionList == null) {
                return Array.Empty<PowerPointSectionInfo>();
            }

            List<SlideId> slideIds = PresentationRoot?.SlideIdList?
                .Elements<SlideId>()
                .ToList() ?? new List<SlideId>();
            Dictionary<uint, int> slideIndexMap = BuildSlideIndexMap(slideIds);

            List<PowerPointSectionInfo> sections = new();
            foreach (P14.Section section in sectionList.Elements<P14.Section>()) {
                List<int> indices = new();
                P14.SectionSlideIdList? list = section.SectionSlideIdList;
                if (list != null) {
                    foreach (P14.SectionSlideIdListEntry entry in list.Elements<P14.SectionSlideIdListEntry>()) {
                        uint? slideId = entry.Id?.Value;
                        if (slideId != null && slideIndexMap.TryGetValue(slideId.Value, out int index)) {
                            indices.Add(index);
                        }
                    }
                }

                indices.Sort();
                string name = section.Name?.Value ?? string.Empty;
                string id = section.Id?.Value ?? string.Empty;
                sections.Add(new PowerPointSectionInfo(name, id, indices));
            }

            return sections;
        }

        /// <summary>
        ///     Adds a new section starting at the specified slide index.
        /// </summary>
        public PowerPointSectionInfo AddSection(string name, int startSlideIndex) {
            ThrowIfDisposed();
            if (string.IsNullOrWhiteSpace(name)) {
                throw new ArgumentException("Section name cannot be null or empty.", nameof(name));
            }

            SlideIdList? slideIdList = PresentationRoot?.SlideIdList;
            if (slideIdList == null) {
                throw new InvalidOperationException("Presentation has no slides.");
            }

            List<SlideId> slideIds = slideIdList.Elements<SlideId>().ToList();
            if (slideIds.Count == 0) {
                throw new InvalidOperationException("Presentation has no slides.");
            }
            if (startSlideIndex < 0 || startSlideIndex >= slideIds.Count) {
                throw new ArgumentOutOfRangeException(nameof(startSlideIndex));
            }

            P14.SectionList sectionList = EnsureSectionList(slideIds);
            uint slideIdValue = slideIds[startSlideIndex].Id?.Value ?? throw new InvalidOperationException("Slide ID is missing.");

            P14.Section? containing = FindSectionBySlideId(sectionList, slideIdValue);
            if (containing == null) {
                P14.Section fallback = sectionList.Elements<P14.Section>().Last();
                EnsureSectionSlideIdList(fallback)
                    .Append(new P14.SectionSlideIdListEntry { Id = slideIdValue });
                return BuildSectionInfo(fallback, slideIds);
            }

            P14.SectionSlideIdList list = EnsureSectionSlideIdList(containing);
            List<P14.SectionSlideIdListEntry> entries = list.Elements<P14.SectionSlideIdListEntry>().ToList();
            int entryIndex = entries.FindIndex(entry => entry.Id?.Value == slideIdValue);
            if (entryIndex <= 0) {
                containing.Name = name;
                return BuildSectionInfo(containing, slideIds);
            }

            List<uint> movedIds = entries
                .Skip(entryIndex)
                .Select(entry => entry.Id?.Value)
                .Where(id => id != null)
                .Select(id => id!.Value)
                .ToList();
            foreach (P14.SectionSlideIdListEntry entry in entries.Skip(entryIndex)) {
                entry.Remove();
            }

            P14.Section newSection = CreateSection(name, movedIds);
            sectionList.InsertAfter(newSection, containing);
            return BuildSectionInfo(newSection, slideIds);
        }

        /// <summary>
        ///     Renames the first section matching the provided name.
        /// </summary>
        public bool RenameSection(string name, string newName, bool ignoreCase = true) {
            ThrowIfDisposed();
            if (name == null) {
                throw new ArgumentNullException(nameof(name));
            }
            if (string.IsNullOrWhiteSpace(newName)) {
                throw new ArgumentException("Section name cannot be null or empty.", nameof(newName));
            }

            P14.SectionList? sectionList = GetSectionList(create: false);
            if (sectionList == null) {
                return false;
            }

            StringComparison comparison = ignoreCase ? StringComparison.OrdinalIgnoreCase : StringComparison.Ordinal;
            foreach (P14.Section section in sectionList.Elements<P14.Section>()) {
                string currentName = section.Name?.Value ?? string.Empty;
                if (string.Equals(currentName, name, comparison)) {
                    section.Name = newName;
                    return true;
                }
            }

            return false;
        }

        private P14.SectionList? GetSectionList(bool create) {
            Presentation presentation = PresentationRoot ??= new Presentation();
            PresentationExtensionList? extList = presentation.GetFirstChild<PresentationExtensionList>();
            if (extList == null && create) {
                extList = new PresentationExtensionList();
                presentation.Append(extList);
            }

            if (extList == null) {
                return null;
            }

            PresentationExtension? sectionExt = extList.Elements<PresentationExtension>()
                .FirstOrDefault(ext => string.Equals(ext.Uri?.Value, SectionListUri, StringComparison.Ordinal));
            if (sectionExt == null && create) {
                sectionExt = new PresentationExtension { Uri = SectionListUri };
                extList.Append(sectionExt);
            }

            if (sectionExt == null) {
                return null;
            }

            P14.SectionList? sectionList = sectionExt.GetFirstChild<P14.SectionList>();
            if (sectionList == null && create) {
                sectionList = new P14.SectionList();
                sectionList.AddNamespaceDeclaration("p14", P14Namespace);
                sectionExt.Append(sectionList);
            }

            return sectionList;
        }

        private P14.SectionList EnsureSectionList(IReadOnlyList<SlideId> slideIds) {
            P14.SectionList sectionList = GetSectionList(create: true)
                ?? throw new InvalidOperationException("Unable to create a section list.");
            if (!sectionList.Elements<P14.Section>().Any()) {
                List<uint> ids = slideIds
                    .Select(id => id.Id?.Value)
                    .Where(id => id != null)
                    .Select(id => id!.Value)
                    .ToList();
                sectionList.Append(CreateSection(DefaultSectionName, ids));
            }

            EnsureSectionCoverage(sectionList, slideIds);
            return sectionList;
        }

        private static void EnsureSectionCoverage(P14.SectionList sectionList, IReadOnlyList<SlideId> slideIds) {
            Dictionary<uint, int> slideIndexMap = BuildSlideIndexMap(slideIds);
            HashSet<uint> assigned = new();
            List<P14.Section> orderedSections = new();

            foreach (P14.Section section in sectionList.Elements<P14.Section>().ToList()) {
                P14.SectionSlideIdList list = EnsureSectionSlideIdList(section);
                List<uint> sectionSlideIds = list.Elements<P14.SectionSlideIdListEntry>()
                    .Select(entry => entry.Id?.Value)
                    .Where(id => id != null && slideIndexMap.ContainsKey(id.Value))
                    .Select(id => id!.Value)
                    .OrderBy(id => slideIndexMap[id])
                    .ToList();

                list.RemoveAllChildren();
                foreach (uint slideId in sectionSlideIds) {
                    if (!assigned.Add(slideId)) {
                        continue;
                    }

                    list.Append(new P14.SectionSlideIdListEntry { Id = slideId });
                }

                if (list.Elements<P14.SectionSlideIdListEntry>().Any()) {
                    orderedSections.Add(section);
                } else {
                    section.Remove();
                }
            }

            if (orderedSections.Count == 0) {
                if (slideIds.Count == 0) {
                    return;
                }

                P14.Section defaultSection = CreateSection(DefaultSectionName, slideIds
                    .Select(id => id.Id?.Value)
                    .Where(id => id != null)
                    .Select(id => id!.Value)
                    .ToList());
                sectionList.RemoveAllChildren();
                sectionList.Append(defaultSection);
                return;
            }

            P14.SectionSlideIdList target = EnsureSectionSlideIdList(orderedSections.Last());
            foreach (uint slideId in slideIds
                         .Select(id => id.Id?.Value)
                         .Where(id => id != null)
                         .Select(id => id!.Value)
                         .Where(id => !assigned.Contains(id))) {
                target.Append(new P14.SectionSlideIdListEntry { Id = slideId });
            }

            orderedSections = orderedSections
                .OrderBy(section => GetSectionStartIndex(section, slideIndexMap))
                .ToList();

            sectionList.RemoveAllChildren();
            foreach (P14.Section section in orderedSections) {
                sectionList.Append(section);
            }
        }

        private static P14.Section CreateSection(string name, IReadOnlyList<uint> slideIds) {
            P14.Section section = new() {
                Id = CreateSectionId(),
                Name = name
            };
            P14.SectionSlideIdList list = new();
            foreach (uint slideId in slideIds) {
                list.Append(new P14.SectionSlideIdListEntry { Id = slideId });
            }
            section.Append(list);
            return section;
        }

        private static string CreateSectionId() {
            return Guid.NewGuid().ToString("B").ToUpperInvariant();
        }

        private static P14.SectionSlideIdList EnsureSectionSlideIdList(P14.Section section) {
            P14.SectionSlideIdList? list = section.SectionSlideIdList;
            if (list == null) {
                list = new P14.SectionSlideIdList();
                section.Append(list);
            }
            return list;
        }

        private static Dictionary<uint, int> BuildSlideIndexMap(IReadOnlyList<SlideId> slideIds) {
            Dictionary<uint, int> map = new();
            for (int i = 0; i < slideIds.Count; i++) {
                uint? id = slideIds[i].Id?.Value;
                if (id != null) {
                    map[id.Value] = i;
                }
            }
            return map;
        }

        private static P14.Section? FindSectionBySlideId(P14.SectionList sectionList, uint slideId) {
            foreach (P14.Section section in sectionList.Elements<P14.Section>()) {
                P14.SectionSlideIdList? list = section.SectionSlideIdList;
                if (list == null) {
                    continue;
                }

                if (list.Elements<P14.SectionSlideIdListEntry>().Any(entry => entry.Id?.Value == slideId)) {
                    return section;
                }
            }

            return null;
        }

        private static int GetSectionStartIndex(P14.Section section, IReadOnlyDictionary<uint, int> slideIndexMap) {
            P14.SectionSlideIdList? list = section.SectionSlideIdList;
            if (list == null) {
                return int.MaxValue;
            }

            return list.Elements<P14.SectionSlideIdListEntry>()
                .Select(entry => entry.Id?.Value)
                .Where(id => id != null && slideIndexMap.ContainsKey(id.Value))
                .Select(id => slideIndexMap[id!.Value])
                .DefaultIfEmpty(int.MaxValue)
                .Min();
        }

        private void AssignSlideToNearestSection(uint slideId, int slideIndex) {
            P14.SectionList? sectionList = GetSectionList(create: false);
            if (sectionList == null) {
                return;
            }

            List<SlideId> slideIds = PresentationRoot?.SlideIdList?
                .Elements<SlideId>()
                .ToList() ?? new List<SlideId>();
            if (slideIds.Count == 0) {
                return;
            }

            if (FindSectionBySlideId(sectionList, slideId) != null) {
                EnsureSectionCoverage(sectionList, slideIds);
                return;
            }

            P14.Section? targetSection = null;
            if (slideIndex > 0) {
                uint? previousSlideId = slideIds[slideIndex - 1].Id?.Value;
                if (previousSlideId != null) {
                    targetSection = FindSectionBySlideId(sectionList, previousSlideId.Value);
                }
            }

            if (targetSection == null && slideIndex + 1 < slideIds.Count) {
                uint? nextSlideId = slideIds[slideIndex + 1].Id?.Value;
                if (nextSlideId != null) {
                    targetSection = FindSectionBySlideId(sectionList, nextSlideId.Value);
                }
            }

            targetSection ??= sectionList.Elements<P14.Section>().LastOrDefault();
            if (targetSection == null) {
                sectionList.Append(CreateSection(DefaultSectionName, new[] { slideId }));
                EnsureSectionCoverage(sectionList, slideIds);
                return;
            }

            EnsureSectionSlideIdList(targetSection).Append(new P14.SectionSlideIdListEntry { Id = slideId });
            EnsureSectionCoverage(sectionList, slideIds);
        }

        private void SyncSectionsWithSlides() {
            P14.SectionList? sectionList = GetSectionList(create: false);
            if (sectionList == null) {
                return;
            }

            List<SlideId> slideIds = PresentationRoot?.SlideIdList?
                .Elements<SlideId>()
                .ToList() ?? new List<SlideId>();
            if (slideIds.Count == 0) {
                PresentationExtension? sectionExtension = sectionList.Parent as PresentationExtension;
                PresentationExtensionList? extensionList = sectionExtension?.Parent as PresentationExtensionList;
                sectionExtension?.Remove();
                if (extensionList != null && !extensionList.Elements<PresentationExtension>().Any()) {
                    extensionList.Remove();
                }
                return;
            }

            EnsureSectionCoverage(sectionList, slideIds);
        }

        private PowerPointSectionInfo BuildSectionInfo(P14.Section section, IReadOnlyList<SlideId> slideIds) {
            Dictionary<uint, int> slideIndexMap = BuildSlideIndexMap(slideIds);
            List<int> indices = new();
            P14.SectionSlideIdList? list = section.SectionSlideIdList;
            if (list != null) {
                foreach (P14.SectionSlideIdListEntry entry in list.Elements<P14.SectionSlideIdListEntry>()) {
                    uint? id = entry.Id?.Value;
                    if (id != null && slideIndexMap.TryGetValue(id.Value, out int index)) {
                        indices.Add(index);
                    }
                }
            }

            indices.Sort();
            return new PowerPointSectionInfo(section.Name?.Value ?? string.Empty, section.Id?.Value ?? string.Empty, indices);
        }

    }
}
