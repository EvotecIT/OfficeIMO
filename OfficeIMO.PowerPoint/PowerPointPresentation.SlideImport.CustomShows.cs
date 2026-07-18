using System.Globalization;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Presentation;
using A = DocumentFormat.OpenXml.Drawing;

namespace OfficeIMO.PowerPoint {
    public sealed partial class PowerPointPresentation {
        private sealed class ImportedPartRoot {
            internal ImportedPartRoot(OpenXmlPart sourcePart,
                OpenXmlPart targetPart) {
                SourcePart = sourcePart;
                TargetPart = targetPart;
            }

            internal OpenXmlPart SourcePart { get; }

            internal OpenXmlPart TargetPart { get; }
        }

        private sealed class SlideImportPlan {
            internal SlideImportPlan(
                IReadOnlyList<PowerPointSlide> slides,
                IReadOnlyList<CustomShow> customShows) {
                Slides = slides;
                CustomShows = customShows;
            }

            internal IReadOnlyList<PowerPointSlide> Slides { get; }

            internal IReadOnlyList<CustomShow> CustomShows { get; }
        }

        private SlideImportPlan CollectSlideImportPlan(
            PowerPointPresentation sourcePresentation,
            PowerPointSlide requestedSource,
            bool includeLinkedSlides) {
            if (!includeLinkedSlides) {
                return new SlideImportPlan(
                    new[] { requestedSource },
                    Array.Empty<CustomShow>());
            }

            IReadOnlyList<PowerPointSlide> sourceSlides =
                sourcePresentation.Slides;
            var sourceByPart = sourceSlides.ToDictionary(slide =>
                slide.SlidePart);
            var selectedSlides = new HashSet<SlidePart>();
            var selectedShows = new Dictionary<uint, CustomShow>();
            var pending = new Queue<SlidePart>();
            pending.Enqueue(requestedSource.SlidePart);
            while (pending.Count > 0) {
                SlidePart sourcePart = pending.Dequeue();
                if (!selectedSlides.Add(sourcePart)) continue;
                CollectSourcePartImportScope(sourcePresentation,
                    sourceByPart, sourcePart,
                    out OpenXmlPart[] importedPartRoots,
                    out SlidePart[] linkedSlides);
                foreach (SlidePart target in linkedSlides) {
                    EnsureListedSourceSlide(sourceByPart, target);
                    if (!selectedSlides.Contains(target)) {
                        pending.Enqueue(target);
                    }
                }

                foreach (uint customShowId in
                         EnumerateCustomShowActionIds(importedPartRoots)) {
                    if (selectedShows.ContainsKey(customShowId)) continue;
                    if (!TryResolveCustomShow(sourcePresentation,
                            sourceByPart, customShowId,
                            out CustomShow? customShow,
                            out SlidePart[] customShowSlides)) {
                        continue;
                    }
                    selectedShows.Add(customShowId, customShow!);
                    foreach (SlidePart customShowSlide in customShowSlides) {
                        if (!selectedSlides.Contains(customShowSlide)) {
                            pending.Enqueue(customShowSlide);
                        }
                    }
                }
            }

            PowerPointSlide[] slides = new[] { requestedSource }
                .Concat(sourceSlides.Where(slide =>
                    !ReferenceEquals(slide, requestedSource)
                    && selectedSlides.Contains(slide.SlidePart)))
                .ToArray();
            CustomShow[] shows = sourcePresentation.PresentationRoot
                .CustomShowList?.Elements<CustomShow>()
                .Where(show => show.Id?.Value is uint id
                    && selectedShows.ContainsKey(id))
                .ToArray() ?? Array.Empty<CustomShow>();
            return new SlideImportPlan(slides, shows);
        }

        private void CollectSourcePartImportScope(
            PowerPointPresentation sourcePresentation,
            IReadOnlyDictionary<SlidePart, PowerPointSlide> sourceByPart,
            SlidePart sourceSlidePart,
            out OpenXmlPart[] importedPartRoots,
            out SlidePart[] linkedSlides) {
            var roots = new HashSet<OpenXmlPart>();
            var targets = new HashSet<SlidePart>();

            void VisitClonedPart(OpenXmlPart part,
                Func<OpenXmlPart, bool>? shouldSkip = null) {
                if (part is SlidePart targetSlide) {
                    targets.Add(targetSlide);
                    return;
                }
                if (!roots.Add(part)) return;
                foreach (IdPartPair child in part.Parts) {
                    if (shouldSkip?.Invoke(child.OpenXmlPart) == true) {
                        continue;
                    }
                    if (ShouldDiscardCustomShowPartRelationship(
                            sourcePresentation, sourceByPart, part,
                            child.RelationshipId)) {
                        continue;
                    }
                    VisitClonedPart(child.OpenXmlPart);
                }
            }

            roots.Add(sourceSlidePart);
            foreach (IdPartPair child in sourceSlidePart.Parts) {
                if (child.OpenXmlPart is SlideLayoutPart
                    || child.OpenXmlPart is NotesSlidePart) {
                    continue;
                }
                if (ShouldDiscardCustomShowPartRelationship(
                        sourcePresentation, sourceByPart, sourceSlidePart,
                        child.RelationshipId)) {
                    continue;
                }
                VisitClonedPart(child.OpenXmlPart);
            }

            NotesSlidePart? notesPart = sourceSlidePart.NotesSlidePart;
            if (notesPart != null) {
                VisitClonedPart(notesPart,
                    part => part is NotesMasterPart);
            }

            SlideLayoutPart? layoutPart = sourceSlidePart.SlideLayoutPart;
            if (layoutPart != null && FindMatchingLayout(layoutPart) == null) {
                SlideMasterPart? masterPart = layoutPart.SlideMasterPart;
                if (masterPart != null && roots.Add(masterPart)) {
                    foreach (IdPartPair child in masterPart.Parts) {
                        if (ShouldDiscardCustomShowPartRelationship(
                                sourcePresentation, sourceByPart,
                                masterPart, child.RelationshipId)) {
                            continue;
                        }
                        if (child.OpenXmlPart is SlideLayoutPart layout) {
                            VisitClonedPart(layout,
                                part => part is SlideMasterPart);
                        } else {
                            VisitClonedPart(child.OpenXmlPart);
                        }
                    }
                }
            }

            importedPartRoots = roots.ToArray();
            linkedSlides = targets.ToArray();
        }

        private static bool ShouldDiscardCustomShowPartRelationship(
            PowerPointPresentation sourcePresentation,
            IReadOnlyDictionary<SlidePart, PowerPointSlide> sourceByPart,
            OpenXmlPart ownerPart,
            string relationshipId) {
            if (!ownerPart.TryGetPartById(relationshipId,
                    out OpenXmlPart? relatedPart)
                || relatedPart is not SlidePart relatedSlide
                || IsNotesSlideBacklink(ownerPart, relatedSlide)) {
                return false;
            }
            OpenXmlPartRootElement? root = ownerPart.RootElement;
            if (root == null) return false;
            A.HyperlinkType[] discardedLinks = root
                .Descendants<A.HyperlinkType>()
                .Where(link => string.Equals(link.Id?.Value,
                        relationshipId, StringComparison.Ordinal)
                    && IsUnresolvableCustomShowAction(sourcePresentation,
                        sourceByPart, link.Action?.Value))
                .ToArray();
            if (discardedLinks.Length == 0) return false;
            var discarded = new HashSet<A.HyperlinkType>(
                discardedLinks);
            return !new OpenXmlElement[] { root }
                .Concat(root.Descendants())
                .Where(element => element is not A.HyperlinkType link
                    || !discarded.Contains(link))
                .SelectMany(element => element.GetAttributes())
                .Any(attribute => string.Equals(attribute.NamespaceUri,
                        PowerPointUtils.RelationshipIdNamespace,
                        StringComparison.Ordinal)
                    && string.Equals(attribute.Value, relationshipId,
                        StringComparison.Ordinal));
        }

        private static bool IsUnresolvableCustomShowAction(
            PowerPointPresentation sourcePresentation,
            IReadOnlyDictionary<SlidePart, PowerPointSlide> sourceByPart,
            string? action) {
            if (!IsCustomShowActionValue(action)) return false;
            return !TryParseCustomShowAction(action, out uint id, out _)
                || !TryResolveCustomShow(sourcePresentation,
                    sourceByPart, id, out _, out _);
        }

        private static void EnsureListedSourceSlide(
            IReadOnlyDictionary<SlidePart, PowerPointSlide> sourceByPart,
            SlidePart target) {
            if (!sourceByPart.ContainsKey(target)) {
                throw new InvalidDataException(
                    "The source presentation contains an internal link to an unlisted slide.");
            }
        }

        private static bool TryResolveCustomShow(
            PowerPointPresentation sourcePresentation,
            IReadOnlyDictionary<SlidePart, PowerPointSlide> sourceByPart,
            uint customShowId,
            out CustomShow? customShow,
            out SlidePart[] slides) {
            CustomShow[] matches = sourcePresentation.PresentationRoot
                .CustomShowList?.Elements<CustomShow>()
                .Where(show => show.Id?.Value == customShowId)
                .ToArray() ?? Array.Empty<CustomShow>();
            if (matches.Length != 1
                || string.IsNullOrEmpty(matches[0].Name?.Value)) {
                customShow = null;
                slides = Array.Empty<SlidePart>();
                return false;
            }

            var resolved = new List<SlidePart>();
            foreach (SlideListEntry entry in matches[0].SlideList?
                         .Elements<SlideListEntry>()
                     ?? Enumerable.Empty<SlideListEntry>()) {
                string? relationshipId = entry.Id?.Value;
                if (string.IsNullOrEmpty(relationshipId)
                    || !sourcePresentation._presentationPart.TryGetPartById(
                        relationshipId!, out OpenXmlPart? part)
                    || part is not SlidePart slidePart
                    || !sourceByPart.ContainsKey(slidePart)) {
                    customShow = null;
                    slides = Array.Empty<SlidePart>();
                    return false;
                }
                resolved.Add(slidePart);
            }

            customShow = matches[0];
            slides = resolved.ToArray();
            return true;
        }

        private void ImportReferencedCustomShows(
            PowerPointPresentation sourcePresentation,
            IReadOnlyList<CustomShow> sourceShows,
            IReadOnlyDictionary<SlidePart, PowerPointSlide> importedSlides,
            IEnumerable<ImportedPartRoot> importedPartRoots) {
            var idMap = new Dictionary<uint, uint>();
            var usedIds = new HashSet<uint>(PresentationRoot.CustomShowList?
                .Elements<CustomShow>()
                .Where(show => show.Id?.Value != null)
                .Select(show => show.Id!.Value)
                ?? Enumerable.Empty<uint>());
            var usedNames = new HashSet<string>(PresentationRoot
                .CustomShowList?.Elements<CustomShow>()
                .Select(show => show.Name?.Value)
                .Where(name => !string.IsNullOrEmpty(name))
                .Cast<string>()
                ?? Enumerable.Empty<string>(), StringComparer.Ordinal);

            foreach (CustomShow sourceShow in sourceShows) {
                uint sourceId = sourceShow.Id!.Value;
                uint targetId = AllocateCustomShowId(sourceId, usedIds);
                string targetName = AllocateCustomShowName(
                    sourceShow.Name!.Value!, usedNames);
                var slideList = new SlideList();
                foreach (SlideListEntry sourceEntry in sourceShow.SlideList?
                             .Elements<SlideListEntry>()
                         ?? Enumerable.Empty<SlideListEntry>()) {
                    string sourceRelationshipId = sourceEntry.Id!.Value!;
                    SlidePart sourceSlidePart = (SlidePart)
                        sourcePresentation._presentationPart.GetPartById(
                            sourceRelationshipId);
                    SlidePart targetSlidePart =
                        importedSlides[sourceSlidePart].SlidePart;
                    slideList.Append(new SlideListEntry {
                        Id = _presentationPart.GetIdOfPart(targetSlidePart)
                    });
                }

                var targetShow = (CustomShow)sourceShow.CloneNode(true);
                targetShow.Id = targetId;
                targetShow.Name = targetName;
                if (targetShow.SlideList != null) {
                    targetShow.ReplaceChild(slideList,
                        targetShow.SlideList);
                } else {
                    targetShow.PrependChild(slideList);
                }
                PresentationRoot.CustomShowList ??= new CustomShowList();
                PresentationRoot.CustomShowList.Append(targetShow);
                idMap.Add(sourceId, targetId);
            }

            RewriteImportedCustomShowActions(
                importedPartRoots, idMap);
        }

        private static uint AllocateCustomShowId(
            uint preferredId,
            ISet<uint> usedIds) {
            if (usedIds.Add(preferredId)) return preferredId;
            uint candidate = 0;
            while (!usedIds.Add(candidate)) {
                if (candidate == uint.MaxValue) {
                    throw new InvalidOperationException(
                        "The presentation has no available custom-show identifiers.");
                }
                candidate++;
            }
            return candidate;
        }

        private static string AllocateCustomShowName(
            string preferredName,
            ISet<string> usedNames) {
            if (usedNames.Add(preferredName)) return preferredName;
            int suffix = 2;
            while (true) {
                string candidate = preferredName + " (" +
                    suffix.ToString(CultureInfo.InvariantCulture) + ")";
                if (usedNames.Add(candidate)) return candidate;
                suffix = checked(suffix + 1);
            }
        }

        private static void RewriteImportedCustomShowActions(
            IEnumerable<ImportedPartRoot> importedPartRoots,
            IReadOnlyDictionary<uint, uint> idMap) {
            foreach (OpenXmlPart targetPart in importedPartRoots
                         .Select(pair => pair.TargetPart)
                         .Distinct()) {
                OpenXmlPartRootElement? root = targetPart.RootElement;
                if (root == null) continue;
                var discardedActionRelationshipIds = new List<string>();
                var discardedSoundRelationshipIds = new List<string>();
                foreach (A.HyperlinkType hyperlink in root
                             .Descendants<A.HyperlinkType>()
                             .Where(link => IsCustomShowActionValue(
                                 link.Action?.Value))
                             .ToArray()) {
                    if (TryParseCustomShowAction(hyperlink.Action?.Value,
                            out uint sourceId, out string suffix)
                        && idMap.TryGetValue(sourceId,
                            out uint targetId)) {
                        hyperlink.Action =
                            "ppaction://customshow?id=" +
                            targetId.ToString(CultureInfo.InvariantCulture) +
                            suffix;
                        continue;
                    }
                    if (!string.IsNullOrEmpty(hyperlink.Id?.Value)) {
                        discardedActionRelationshipIds.Add(
                            hyperlink.Id!.Value!);
                    }
                    discardedSoundRelationshipIds.AddRange(hyperlink
                        .Elements<A.HyperlinkSound>()
                        .Select(sound => sound.Embed?.Value)
                        .Where(id => !string.IsNullOrEmpty(id))
                        .Cast<string>());
                    hyperlink.Remove();
                }
                root.Save();
                foreach (string relationshipId in
                         discardedActionRelationshipIds.Distinct(
                             StringComparer.Ordinal)) {
                    RemoveActionRelationshipIfUnused(targetPart,
                        relationshipId);
                }
                PowerPointEmbeddedSound.RemoveIfUnused(targetPart,
                    discardedSoundRelationshipIds);
            }
        }

        private static void RemoveActionRelationshipIfUnused(
            OpenXmlPart ownerPart,
            string relationshipId) {
            OpenXmlPartRootElement? root = ownerPart.RootElement;
            if (root != null && new OpenXmlElement[] { root }
                    .Concat(root.Descendants())
                    .SelectMany(element => element.GetAttributes())
                    .Any(attribute => string.Equals(attribute.NamespaceUri,
                            PowerPointUtils.RelationshipIdNamespace,
                            StringComparison.Ordinal)
                        && string.Equals(attribute.Value, relationshipId,
                            StringComparison.Ordinal))) {
                return;
            }

            ExternalRelationship? external = ownerPart
                .ExternalRelationships.FirstOrDefault(relationship =>
                    string.Equals(relationship.Id, relationshipId,
                        StringComparison.Ordinal));
            if (external != null) {
                ownerPart.DeleteReferenceRelationship(external);
                return;
            }
            HyperlinkRelationship? hyperlink = ownerPart
                .HyperlinkRelationships.FirstOrDefault(relationship =>
                    string.Equals(relationship.Id, relationshipId,
                        StringComparison.Ordinal));
            if (hyperlink != null) {
                ownerPart.DeleteReferenceRelationship(hyperlink);
                return;
            }
            if (ownerPart.TryGetPartById(relationshipId,
                    out OpenXmlPart? relatedPart)
                && relatedPart is SlidePart relatedSlide
                && !IsNotesSlideBacklink(ownerPart, relatedSlide)) {
                ownerPart.DeletePart(relationshipId);
            }
        }

        private static bool IsNotesSlideBacklink(
            OpenXmlPart ownerPart,
            SlidePart relatedSlide) =>
            ownerPart is NotesSlidePart notesPart
            && ReferenceEquals(relatedSlide.NotesSlidePart, notesPart);

        private static IEnumerable<uint> EnumerateCustomShowActionIds(
            IEnumerable<OpenXmlPart> importedPartRoots) {
            foreach (OpenXmlPart part in importedPartRoots) {
                OpenXmlPartRootElement? root = part.RootElement;
                if (root == null) continue;
                foreach (string action in root
                             .Descendants<A.HyperlinkType>()
                             .Select(link => link.Action?.Value)
                             .Where(IsCustomShowActionValue)
                             .Cast<string>()) {
                    if (TryParseCustomShowAction(action, out uint id,
                            out _)) {
                        yield return id;
                    }
                }
            }
        }

        private static bool IsCustomShowActionValue(string? action) =>
            action?.StartsWith("ppaction://customshow?id=",
                StringComparison.Ordinal) == true;

        private static bool TryParseCustomShowAction(
            string? action,
            out uint id,
            out string suffix) {
            const string Prefix = "ppaction://customshow?id=";
            id = 0;
            suffix = string.Empty;
            if (!IsCustomShowActionValue(action)) return false;
            int suffixIndex = action!.IndexOf('&', Prefix.Length);
            string idText = suffixIndex < 0
                ? action.Substring(Prefix.Length)
                : action.Substring(Prefix.Length,
                    suffixIndex - Prefix.Length);
            suffix = suffixIndex < 0
                ? string.Empty
                : action.Substring(suffixIndex);
            return uint.TryParse(idText, NumberStyles.None,
                CultureInfo.InvariantCulture, out id);
        }

    }
}
