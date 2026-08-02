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
            List<PowerPointSlide> slides = Slides.ToList();
            slides.Add(slide);
            return SetSlides(slides);
        }

        /// <summary>Inserts a slide at the specified zero-based custom-show position.</summary>
        public PowerPointCustomShow InsertSlide(int index, PowerPointSlide slide) {
            List<PowerPointSlide> slides = Slides.ToList();
            if (index < 0 || index > slides.Count) {
                throw new ArgumentOutOfRangeException(nameof(index));
            }
            slides.Insert(index, slide);
            return SetSlides(slides);
        }

        /// <summary>Removes the first matching slide from the custom show.</summary>
        public bool RemoveSlide(PowerPointSlide slide) {
            if (slide == null) throw new ArgumentNullException(nameof(slide));
            List<PowerPointSlide> slides = Slides.ToList();
            int index = slides.IndexOf(slide);
            if (index < 0) return false;
            slides.RemoveAt(index);
            SetSlides(slides);
            return true;
        }

        /// <summary>Moves a slide between zero-based custom-show positions.</summary>
        public PowerPointCustomShow MoveSlide(int sourceIndex, int destinationIndex) {
            List<PowerPointSlide> slides = Slides.ToList();
            if (sourceIndex < 0 || sourceIndex >= slides.Count) {
                throw new ArgumentOutOfRangeException(nameof(sourceIndex));
            }
            if (destinationIndex < 0 || destinationIndex >= slides.Count) {
                throw new ArgumentOutOfRangeException(nameof(destinationIndex));
            }
            if (sourceIndex == destinationIndex) return this;
            PowerPointSlide slide = slides[sourceIndex];
            slides.RemoveAt(sourceIndex);
            slides.Insert(destinationIndex, slide);
            return SetSlides(slides);
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
            var slides = new List<PowerPointSlide>();
            foreach (P.SlideListEntry entry in customShow.SlideList?
                         .Elements<P.SlideListEntry>() ?? Enumerable.Empty<P.SlideListEntry>()) {
                string? relationshipId = entry.Id?.Value;
                if (string.IsNullOrEmpty(relationshipId)
                    || !_presentationPart.TryGetPartById(relationshipId!, out OpenXmlPart? part)
                    || part is not SlidePart slidePart) {
                    continue;
                }
                PowerPointSlide? slide = _slides.FirstOrDefault(current =>
                    ReferenceEquals(current.SlidePart, slidePart));
                if (slide != null) slides.Add(slide);
            }
            return slides;
        }

        internal void SetCustomShowSlides(P.CustomShow customShow,
            IEnumerable<PowerPointSlide> slides) {
            ThrowIfDisposed();
            RequireOwnedCustomShow(customShow);
            customShow.SlideList = CreateCustomShowSlideList(
                ValidateCustomShowSlides(slides));
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
            uint maximum = PresentationRoot.CustomShowList?
                .Elements<P.CustomShow>()
                .Select(show => show.Id?.Value ?? 0U)
                .DefaultIfEmpty(0U).Max() ?? 0U;
            if (maximum == uint.MaxValue) {
                throw new InvalidOperationException(
                    "The custom-show identifier space is exhausted.");
            }
            return maximum + 1U;
        }
    }
}
