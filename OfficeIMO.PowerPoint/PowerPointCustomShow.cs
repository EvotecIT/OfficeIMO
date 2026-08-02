using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.Linq;
using DocumentFormat.OpenXml.Packaging;
using P = DocumentFormat.OpenXml.Presentation;

namespace OfficeIMO.PowerPoint {
    /// <summary>
    ///     Represents an editable named slide sequence in a PowerPoint presentation.
    /// </summary>
    public sealed class PowerPointCustomShow {
        private readonly PowerPointPresentation _presentation;
        private readonly P.CustomShow _customShow;

        internal PowerPointCustomShow(PowerPointPresentation presentation,
            P.CustomShow customShow) {
            _presentation = presentation ?? throw new ArgumentNullException(nameof(presentation));
            _customShow = customShow ?? throw new ArgumentNullException(nameof(customShow));
        }

        /// <summary>Gets the stable custom-show identifier stored in the package.</summary>
        public uint Id => _customShow.Id?.Value ?? 0U;

        /// <summary>Gets the custom-show display name.</summary>
        public string Name => _customShow.Name?.Value ?? string.Empty;

        /// <summary>Gets the slides in playback order.</summary>
        public IReadOnlyList<PowerPointSlide> Slides => new ReadOnlyCollection<PowerPointSlide>(
            _presentation.ResolveCustomShowSlides(_customShow).ToList());

        /// <summary>Replaces the custom-show slide sequence.</summary>
        public PowerPointCustomShow SetSlides(IEnumerable<PowerPointSlide> slides) {
            _presentation.SetCustomShowSlides(_customShow, slides);
            return this;
        }

        /// <summary>Adds a slide to the end of the custom show.</summary>
        public PowerPointCustomShow AddSlide(PowerPointSlide slide) {
            _presentation.InsertCustomShowSlide(_customShow,
                Slides.Count, slide);
            return this;
        }

        /// <summary>Inserts a slide at the specified zero-based custom-show position.</summary>
        public PowerPointCustomShow InsertSlide(int index, PowerPointSlide slide) {
            _presentation.InsertCustomShowSlide(_customShow, index, slide);
            return this;
        }

        /// <summary>Removes the first matching slide from the custom show.</summary>
        public bool RemoveSlide(PowerPointSlide slide) {
            return _presentation.RemoveCustomShowSlide(_customShow, slide);
        }

        /// <summary>Moves a slide between zero-based custom-show positions.</summary>
        public PowerPointCustomShow MoveSlide(int sourceIndex, int destinationIndex) {
            _presentation.MoveCustomShowSlide(_customShow, sourceIndex,
                destinationIndex);
            return this;
        }

        internal P.CustomShow OpenXmlElement => _customShow;
    }

    public sealed partial class PowerPointPresentation {
        /// <summary>Gets the custom shows in package order.</summary>
        public IReadOnlyList<PowerPointCustomShow> CustomShows {
            get {
                ThrowIfDisposed();
                P.CustomShow[] shows = PresentationRoot.CustomShowList?
                    .Elements<P.CustomShow>().ToArray() ?? Array.Empty<P.CustomShow>();
                return new ReadOnlyCollection<PowerPointCustomShow>(shows
                    .Select(show => new PowerPointCustomShow(this, show)).ToList());
            }
        }

        /// <summary>Creates a custom show with an explicit, nonempty slide sequence.</summary>
        public PowerPointCustomShow AddCustomShow(string name,
            IEnumerable<PowerPointSlide> slides) {
            ThrowIfDisposed();
            string normalizedName = ValidateCustomShowName(name, except: null);
            PowerPointSlide[] resolvedSlides = ValidateCustomShowSlides(slides);
            uint id = AllocateCustomShowId();
            var customShow = new P.CustomShow {
                Id = id,
                Name = normalizedName,
                SlideList = CreateCustomShowSlideList(resolvedSlides)
            };
            PresentationRoot.CustomShowList ??= new P.CustomShowList();
            PresentationRoot.CustomShowList.Append(customShow);
            return new PowerPointCustomShow(this, customShow);
        }

        /// <summary>Finds a custom show by name using ordinal, case-insensitive matching.</summary>
        public PowerPointCustomShow? GetCustomShow(string name) {
            ThrowIfDisposed();
            if (string.IsNullOrWhiteSpace(name)) return null;
            P.CustomShow? show = PresentationRoot.CustomShowList?
                .Elements<P.CustomShow>()
                .FirstOrDefault(current => string.Equals(current.Name?.Value,
                    name.Trim(), StringComparison.OrdinalIgnoreCase));
            return show == null ? null : new PowerPointCustomShow(this, show);
        }

        /// <summary>Renames a custom show while preserving its identifier and linked actions.</summary>
        public PowerPointCustomShow RenameCustomShow(PowerPointCustomShow customShow,
            string name) {
            ThrowIfDisposed();
            P.CustomShow owned = RequireOwnedCustomShow(customShow);
            owned.Name = ValidateCustomShowName(name, owned);
            return customShow;
        }

        /// <summary>Removes a custom show and actions that target its identifier.</summary>
        public bool RemoveCustomShow(PowerPointCustomShow customShow) {
            ThrowIfDisposed();
            P.CustomShow owned = RequireOwnedCustomShow(customShow);
            uint? id = owned.Id?.Value;
            owned.Remove();
            if (PresentationRoot.CustomShowList?.Elements<P.CustomShow>().Any() == false) {
                PresentationRoot.CustomShowList.Remove();
            }
            if (id.HasValue) RemoveCustomShowLinks(id.Value);
            return true;
        }

        internal IReadOnlyList<PowerPointSlide> ResolveCustomShowSlides(
            P.CustomShow customShow) {
            return ResolveCustomShowEntries(customShow)
                .Select(item => item.Slide).ToList().AsReadOnly();
        }

        private IReadOnlyList<(P.SlideListEntry Entry,
            PowerPointSlide Slide)> ResolveCustomShowEntries(
            P.CustomShow customShow) {
            var entries = new List<(P.SlideListEntry,
                PowerPointSlide)>();
            foreach (P.SlideListEntry entry in customShow.SlideList?
                         .Elements<P.SlideListEntry>()
                     ?? Enumerable.Empty<P.SlideListEntry>()) {
                string? relationshipId = entry.Id?.Value;
                if (string.IsNullOrEmpty(relationshipId)
                    || !_presentationPart.TryGetPartById(relationshipId!, out OpenXmlPart? part)
                    || part is not SlidePart slidePart) {
                    continue;
                }
                PowerPointSlide? slide = _slides.FirstOrDefault(current =>
                    ReferenceEquals(current.SlidePart, slidePart));
                if (slide != null) entries.Add((entry, slide));
            }
            return entries.AsReadOnly();
        }

        internal void InsertCustomShowSlide(P.CustomShow customShow,
            int index, PowerPointSlide slide) {
            ThrowIfDisposed();
            RequireOwnedCustomShow(customShow);
            PowerPointSlide resolvedSlide = ValidateCustomShowSlides(
                new[] { slide })[0];
            IReadOnlyList<(P.SlideListEntry Entry, PowerPointSlide Slide)>
                entries = ResolveCustomShowEntries(customShow);
            if (index < 0 || index > entries.Count) {
                throw new ArgumentOutOfRangeException(nameof(index));
            }
            P.SlideList slideList = customShow.SlideList
                ?? new P.SlideList();
            var entry = new P.SlideListEntry {
                Id = _presentationPart.GetIdOfPart(resolvedSlide.SlidePart)
            };
            InsertCustomShowEntry(slideList, entry, index, entries);
            if (customShow.SlideList == null) {
                customShow.SlideList = slideList;
            }
        }

        internal bool RemoveCustomShowSlide(P.CustomShow customShow,
            PowerPointSlide slide) {
            ThrowIfDisposed();
            RequireOwnedCustomShow(customShow);
            if (slide == null) throw new ArgumentNullException(nameof(slide));
            IReadOnlyList<(P.SlideListEntry Entry, PowerPointSlide Slide)>
                entries = ResolveCustomShowEntries(customShow);
            (P.SlideListEntry Entry, PowerPointSlide Slide) match = entries
                .FirstOrDefault(item => ReferenceEquals(item.Slide, slide));
            if (match.Entry == null) return false;
            if (entries.Count == 1) {
                throw new InvalidOperationException(
                    "A custom show requires at least one slide.");
            }
            match.Entry.Remove();
            return true;
        }

        internal void MoveCustomShowSlide(P.CustomShow customShow,
            int sourceIndex, int destinationIndex) {
            ThrowIfDisposed();
            RequireOwnedCustomShow(customShow);
            IReadOnlyList<(P.SlideListEntry Entry, PowerPointSlide Slide)>
                entries = ResolveCustomShowEntries(customShow);
            if (sourceIndex < 0 || sourceIndex >= entries.Count) {
                throw new ArgumentOutOfRangeException(nameof(sourceIndex));
            }
            if (destinationIndex < 0 || destinationIndex >= entries.Count) {
                throw new ArgumentOutOfRangeException(nameof(destinationIndex));
            }
            if (sourceIndex == destinationIndex) return;
            P.SlideListEntry moved = entries[sourceIndex].Entry;
            moved.Remove();
            var remaining = entries.Where((_, index) => index != sourceIndex)
                .ToList().AsReadOnly();
            InsertCustomShowEntry(customShow.SlideList!, moved,
                destinationIndex, remaining);
        }

        private static void InsertCustomShowEntry(P.SlideList slideList,
            P.SlideListEntry entry, int index,
            IReadOnlyList<(P.SlideListEntry Entry, PowerPointSlide Slide)>
                resolvedEntries) {
            if (index < resolvedEntries.Count) {
                slideList.InsertBefore(entry, resolvedEntries[index].Entry);
                return;
            }
            if (resolvedEntries.Count > 0) {
                slideList.InsertAfter(entry,
                    resolvedEntries[resolvedEntries.Count - 1].Entry);
                return;
            }
            OpenXmlElement? firstExtension = slideList.ChildElements
                .FirstOrDefault(child => child is not P.SlideListEntry);
            if (firstExtension != null) {
                slideList.InsertBefore(entry, firstExtension);
            } else {
                slideList.Append(entry);
            }
        }

        internal void SetCustomShowSlides(P.CustomShow customShow,
            IEnumerable<PowerPointSlide> slides) {
            ThrowIfDisposed();
            RequireOwnedCustomShow(customShow);
            PowerPointSlide[] resolvedSlides = ValidateCustomShowSlides(slides);
            P.SlideList slideList = customShow.SlideList
                ?? new P.SlideList();
            P.SlideListEntry[] existingEntries = slideList
                .Elements<P.SlideListEntry>().ToArray();
            var entriesByRelationship = new Dictionary<string,
                Queue<P.SlideListEntry>>(StringComparer.Ordinal);
            foreach (P.SlideListEntry entry in existingEntries) {
                string? relationshipId = entry.Id?.Value;
                if (string.IsNullOrEmpty(relationshipId)) continue;
                if (!entriesByRelationship.TryGetValue(relationshipId!,
                        out Queue<P.SlideListEntry>? entries)) {
                    entries = new Queue<P.SlideListEntry>();
                    entriesByRelationship.Add(relationshipId!, entries);
                }
                entries.Enqueue(entry);
            }

            var selectedEntries = new List<P.SlideListEntry>(
                resolvedSlides.Length);
            foreach (PowerPointSlide slide in resolvedSlides) {
                string relationshipId = _presentationPart.GetIdOfPart(
                    slide.SlidePart);
                P.SlideListEntry entry = entriesByRelationship.TryGetValue(
                        relationshipId, out Queue<P.SlideListEntry>? entries)
                    && entries.Count > 0
                        ? entries.Dequeue()
                        : new P.SlideListEntry { Id = relationshipId };
                selectedEntries.Add(entry);
            }

            foreach (P.SlideListEntry entry in existingEntries) {
                entry.Remove();
            }
            for (int i = selectedEntries.Count - 1; i >= 0; i--) {
                slideList.PrependChild(selectedEntries[i]);
            }
            if (customShow.SlideList == null) {
                customShow.SlideList = slideList;
            }
        }

        private P.CustomShow RequireOwnedCustomShow(PowerPointCustomShow customShow) {
            if (customShow == null) throw new ArgumentNullException(nameof(customShow));
            return RequireOwnedCustomShow(customShow.OpenXmlElement);
        }

        private P.CustomShow RequireOwnedCustomShow(P.CustomShow customShow) {
            if (PresentationRoot.CustomShowList?.Elements<P.CustomShow>()
                    .Any(current => ReferenceEquals(current, customShow)) != true) {
                throw new InvalidOperationException(
                    "The custom show does not belong to this presentation.");
            }
            return customShow;
        }

        private string ValidateCustomShowName(string name, P.CustomShow? except) {
            if (string.IsNullOrWhiteSpace(name)) {
                throw new ArgumentException("Custom-show name cannot be empty.", nameof(name));
            }
            string normalized = name.Trim();
            PowerPointXmlValueValidator.ValidateCharacters(normalized,
                nameof(name), "Custom-show name");
            bool duplicate = PresentationRoot.CustomShowList?
                .Elements<P.CustomShow>()
                .Any(show => !ReferenceEquals(show, except)
                    && string.Equals(show.Name?.Value, normalized,
                        StringComparison.OrdinalIgnoreCase)) == true;
            if (duplicate) {
                throw new InvalidOperationException(
                    $"A custom show named '{normalized}' already exists.");
            }
            return normalized;
        }

        private PowerPointSlide[] ValidateCustomShowSlides(
            IEnumerable<PowerPointSlide> slides) {
            if (slides == null) throw new ArgumentNullException(nameof(slides));
            PowerPointSlide[] resolved = slides.ToArray();
            if (resolved.Length == 0) {
                throw new ArgumentException(
                    "A custom show requires at least one slide.", nameof(slides));
            }
            if (resolved.Any(slide => slide == null || !_slides.Contains(slide))) {
                throw new InvalidOperationException(
                    "Every custom-show slide must belong to this presentation.");
            }
            return resolved;
        }

        private P.SlideList CreateCustomShowSlideList(
            IEnumerable<PowerPointSlide> slides) => new(slides.Select(slide =>
                new P.SlideListEntry {
                    Id = _presentationPart.GetIdOfPart(slide.SlidePart)
                }));

        private uint AllocateCustomShowId() {
            var usedIds = new HashSet<uint>(PresentationRoot.CustomShowList?
                .Elements<P.CustomShow>()
                .Where(show => show.Id?.Value != null)
                .Select(show => show.Id!.Value)
                ?? Enumerable.Empty<uint>());
            return FindAvailableUInt32Id(usedIds, 1U,
                "custom-show");
        }

        private static uint FindAvailableUInt32Id(
            ISet<uint> usedIds, uint firstCandidate, string identifierKind,
            uint maximumInclusive = uint.MaxValue) {
            uint candidate = firstCandidate <= maximumInclusive
                ? firstCandidate
                : 0U;
            uint initialCandidate = candidate;
            do {
                if (!usedIds.Contains(candidate)) return candidate;
                candidate = candidate == maximumInclusive ? 0U : candidate + 1U;
            } while (candidate != initialCandidate);
            throw new InvalidOperationException(
                $"The presentation has no available {identifierKind} identifiers.");
        }
    }
}
