using System;
using System.IO;
using System.Reflection;
using System.Runtime.ExceptionServices;
using System.Xml.Linq;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Presentation;
using DocumentFormat.OpenXml.Validation;
using OfficeIMO.Drawing.Internal;
using A = DocumentFormat.OpenXml.Drawing;
using P14 = DocumentFormat.OpenXml.Office2010.PowerPoint;

namespace OfficeIMO.PowerPoint {
    public sealed partial class PowerPointPresentation {
        /// <summary>
        ///     Adds a new slide using the specified master and layout indexes.
        /// </summary>
        /// <param name="masterIndex">Index of the slide master.</param>
        /// <param name="layoutIndex">Index of the slide layout.</param>
        public PowerPointSlide AddSlide(int masterIndex = 0, int layoutIndex = 0) {
            ThrowIfDisposed();
            string slideRelId = GetNextSlideRelationshipId();
            SlidePart slidePart = _presentationPart.AddNewPart<SlidePart>(slideRelId);
            // Create slide exactly like the working example
            slidePart.Slide = new Slide(
                new CommonSlideData(
                    new ShapeTree(
                        new NonVisualGroupShapeProperties(
                            new NonVisualDrawingProperties() { Id = 1U, Name = "" },
                            new NonVisualGroupShapeDrawingProperties(),
                            new ApplicationNonVisualDrawingProperties()),
                        PowerPointUtils.CreateDefaultGroupShapeProperties())),
                new ColorMapOverride(new A.MasterColorMapping()));

            SlideMasterPart[] masters = _presentationPart.SlideMasterParts.ToArray();
            if (masterIndex < 0 || masterIndex >= masters.Length) {
                throw new ArgumentOutOfRangeException(nameof(masterIndex));
            }

            SlideMasterPart masterPart = masters[masterIndex];

            SlideLayoutPart[] layouts = masterPart.SlideLayoutParts.ToArray();
            if (layoutIndex < 0 || layoutIndex >= layouts.Length) {
                throw new ArgumentOutOfRangeException(nameof(layoutIndex));
            }

            SlideLayoutPart layoutPart = layouts[layoutIndex];

            // Check if this slide part already has a reference to this layout part
            string? existingRelId = null;
            foreach (var partPair in slidePart.Parts) {
                if (partPair.OpenXmlPart == layoutPart) {
                    existingRelId = partPair.RelationshipId;
                    break;
                }
            }

            if (existingRelId == null) {
                // Layout part not yet referenced, add it with a unique relationship ID
                // Check if rId1 is already in use by this slide part
                var slideRelationships = new HashSet<string>(
                    slidePart.Parts.Select(p => p.RelationshipId)
                    .Union(slidePart.ExternalRelationships.Select(r => r.Id))
                    .Union(slidePart.HyperlinkRelationships.Select(r => r.Id))
                    .Where(id => !string.IsNullOrEmpty(id))
                );

                // Find a unique relationship ID for the layout
                string layoutRelId = "rId1";
                if (slideRelationships.Contains(layoutRelId)) {
                    int layoutIdNum = 1;
                    do {
                        layoutRelId = "rId" + layoutIdNum;
                        layoutIdNum++;
                    } while (slideRelationships.Contains(layoutRelId));
                }

                slidePart.AddPart(layoutPart, layoutRelId);
            }
            // If the layout is already referenced, we don't need to add it again

            if (PresentationRoot.SlideIdList == null) {
                PresentationRoot.SlideIdList = new SlideIdList();
            }

            uint newId = GetNextSlideId();
            SlideId slideId = new() { Id = newId };
            PowerPointUtils.SetRelationshipIdValue(slideId, slideRelId);
            PresentationRoot.SlideIdList.Append(slideId);
            AssignSlideToNearestSection(newId, _slides.Count);
            PresentationRoot.Save();

            PowerPointSlide slide = new(slidePart);
            _slides.Add(slide);
            return slide;
        }

        /// <summary>
        ///     Removes the slide at the specified index.
        /// </summary>
        /// <param name="index">Index of the slide to remove.</param>
        public void RemoveSlide(int index) {
            if (index < 0 || index >= _slides.Count) {
                throw new ArgumentOutOfRangeException(nameof(index));
            }

            SlideIdList? slideIdList = PresentationRoot.SlideIdList;
            if (slideIdList == null) {
                throw new InvalidOperationException("Presentation has no slides.");
            }

            SlideId slideId = slideIdList.Elements<SlideId>().ElementAt(index);
            string? relIdValue = PowerPointUtils.GetRelationshipIdValue(slideId);

            _slides.RemoveAt(index);
            slideId.Remove();

            if (!string.IsNullOrWhiteSpace(relIdValue)) {
                string relId = relIdValue!;
                IEnumerable<SlideListEntry> customShowEntries = PresentationRoot
                    .CustomShowList?.Descendants<SlideListEntry>()
                    ?? Enumerable.Empty<SlideListEntry>();
                var emptiedCustomShowIds = new HashSet<uint>();
                foreach (SlideListEntry customShowEntry in customShowEntries.ToArray()) {
                    if (string.Equals(customShowEntry.Id?.Value, relId,
                            StringComparison.Ordinal)) {
                        CustomShow? customShow = customShowEntry
                            .Ancestors<CustomShow>().FirstOrDefault();
                        customShowEntry.Remove();
                        if (customShow?.SlideList?.Elements<SlideListEntry>()
                                .Any() == false) {
                            if (customShow.Id?.Value is uint customShowId) {
                                emptiedCustomShowIds.Add(customShowId);
                            }
                            customShow.Remove();
                        }
                    }
                }
                if (PresentationRoot.CustomShowList?
                        .Elements<CustomShow>().Any() == false) {
                    PresentationRoot.CustomShowList.Remove();
                }
                foreach (uint customShowId in emptiedCustomShowIds) {
                    RemoveCustomShowLinks(customShowId);
                }
                OpenXmlPart part = _presentationPart.GetPartById(relId);
                if (part is SlidePart targetSlidePart) {
                    RemoveInboundSlideLinks(targetSlidePart);
                }
                _presentationPart.DeletePart(part);
            }

            SyncSectionsWithSlides();
            PresentationRoot.Save();
        }

        private void RemoveCustomShowLinks(uint customShowId) {
            string prefix = "ppaction://customshow?id="
                + customShowId.ToString(
                    System.Globalization.CultureInfo.InvariantCulture);
            var visited = new HashSet<OpenXmlPart>();
            var pending = new Stack<OpenXmlPart>();
            pending.Push(_presentationPart);
            while (pending.Count > 0) {
                OpenXmlPart part = pending.Pop();
                if (!visited.Add(part)) continue;
                foreach (IdPartPair child in part.Parts) {
                    pending.Push(child.OpenXmlPart);
                }
                OpenXmlPartRootElement? root = part.RootElement;
                if (root == null) continue;
                A.HyperlinkType[] links = root.Descendants<A.HyperlinkType>()
                    .Where(link => IsCustomShowAction(
                        link.Action?.Value, prefix))
                    .ToArray();
                if (links.Length == 0) continue;
                foreach (A.HyperlinkType link in links) link.Remove();
                root.Save();
            }
        }

        private static bool IsCustomShowAction(string? action,
            string expectedPrefix) => action != null
            && action.StartsWith(expectedPrefix, StringComparison.Ordinal)
            && (action.Length == expectedPrefix.Length
                || action[expectedPrefix.Length] == '&');

        private void RemoveInboundSlideLinks(SlidePart targetSlidePart) {
            foreach (SlidePart sourceSlidePart in _presentationPart.SlideParts
                         .Where(part => !ReferenceEquals(part,
                             targetSlidePart)).ToArray()) {
                Slide? sourceSlide = sourceSlidePart.Slide;
                if (sourceSlide == null) continue;
                string[] relationshipIds = sourceSlidePart.Parts
                    .Where(pair => ReferenceEquals(pair.OpenXmlPart,
                        targetSlidePart))
                    .Select(pair => pair.RelationshipId)
                    .ToArray();
                if (relationshipIds.Length == 0) continue;
                foreach (string relationshipId in relationshipIds) {
                    foreach (A.HyperlinkType hyperlink in sourceSlide
                                 .Descendants<A.HyperlinkType>()
                                 .Where(link => string.Equals(link.Id?.Value,
                                     relationshipId,
                                     StringComparison.Ordinal)).ToArray()) {
                        hyperlink.Remove();
                    }
                    sourceSlidePart.DeletePart(relationshipId);
                }
                sourceSlide.Save();
            }
        }

        private void ValidateSlideIndex(int index) {
            if (index < 0 || index >= _slides.Count) {
                throw new ArgumentOutOfRangeException(nameof(index));
            }
        }

        /// <summary>
        ///     Moves a slide from one index to another.
        /// </summary>
        /// <param name="fromIndex">Current index of the slide.</param>
        /// <param name="toIndex">Destination index of the slide.</param>
        public void MoveSlide(int fromIndex, int toIndex) {
            if (fromIndex < 0 || fromIndex >= _slides.Count) {
                throw new ArgumentOutOfRangeException(nameof(fromIndex));
            }

            if (toIndex < 0 || toIndex >= _slides.Count) {
                throw new ArgumentOutOfRangeException(nameof(toIndex));
            }

            if (fromIndex == toIndex) {
                return;
            }

            SlideIdList? slideIdList = PresentationRoot.SlideIdList;
            if (slideIdList == null) {
                throw new InvalidOperationException("Presentation has no slides.");
            }

            PowerPointSlide slide = _slides[fromIndex];
            _slides.RemoveAt(fromIndex);
            _slides.Insert(toIndex, slide);

            List<SlideId> ids = slideIdList.Elements<SlideId>().ToList();
            SlideId movingId = ids[fromIndex];
            ids.RemoveAt(fromIndex);
            ids.Insert(toIndex, movingId);

            slideIdList.RemoveAllChildren();
            foreach (SlideId id in ids) {
                slideIdList.Append(id);
            }

            SyncSectionsWithSlides();
            PresentationRoot.Save();
        }

        /// <summary>
        ///     Duplicates a slide and inserts it into the presentation.
        /// </summary>
        /// <param name="index">Index of the slide to duplicate.</param>
        /// <param name="insertAt">Index where the duplicate should be inserted. Defaults to index + 1.</param>
        public PowerPointSlide DuplicateSlide(int index, int? insertAt = null) {
            ThrowIfDisposed();
            if (index < 0 || index >= _slides.Count) {
                throw new ArgumentOutOfRangeException(nameof(index));
            }

            int targetIndex = insertAt ?? index + 1;
            if (targetIndex < 0 || targetIndex > _slides.Count) {
                throw new ArgumentOutOfRangeException(nameof(insertAt));
            }

            PowerPointSlide sourceSlide = _slides[index];
            SlidePart sourcePart = sourceSlide.SlidePart;
            Slide sourceSlideRoot = sourcePart.Slide ?? throw new InvalidOperationException("Source slide is missing its slide definition.");

            sourceSlide.Save();

            string slideRelId = GetNextSlideRelationshipId();
            SlidePart slidePart = _presentationPart.AddNewPart<SlidePart>(slideRelId);
            slidePart.Slide = (Slide)sourceSlideRoot.CloneNode(true);

            CloneSlidePartRelationships(sourcePart, slidePart, ShouldSharePart, includeDataParts: true);

            SlideIdList slideIdList = PresentationRoot.SlideIdList ??= new SlideIdList();
            SlideId slideId = new() { Id = GetNextSlideId() };
            PowerPointUtils.SetRelationshipIdValue(slideId, slideRelId);
            InsertSlideId(slideIdList, slideId, targetIndex);
            AssignSlideToNearestSection(slideId.Id?.Value ?? throw new InvalidOperationException("Slide ID is missing."),
                targetIndex);

            PowerPointSlide duplicate = new(slidePart);
            duplicate.Hidden = sourceSlide.Hidden;
            _slides.Insert(targetIndex, duplicate);
            PresentationRoot.Save();
            return duplicate;
        }

        /// <summary>
        ///     Imports a slide from another presentation and inserts it into the current presentation.
        /// </summary>
        /// <param name="sourcePresentation">Presentation to import from.</param>
        /// <param name="sourceIndex">Index of the slide to import.</param>
        /// <param name="insertAt">Index where the imported slide should be inserted. Defaults to end.</param>
        public PowerPointSlide ImportSlide(PowerPointPresentation sourcePresentation, int sourceIndex, int? insertAt = null) {
            ThrowIfDisposed();
            if (sourcePresentation == null) {
                throw new ArgumentNullException(nameof(sourcePresentation));
            }

            if (ReferenceEquals(sourcePresentation, this)) {
                return DuplicateSlide(sourceIndex, insertAt);
            }

            IReadOnlyList<PowerPointSlide> sourceSlides = sourcePresentation.Slides;
            if (sourceIndex < 0 || sourceIndex >= sourceSlides.Count) {
                throw new ArgumentOutOfRangeException(nameof(sourceIndex));
            }

            int targetIndex = insertAt ?? _slides.Count;
            if (targetIndex < 0 || targetIndex > _slides.Count) {
                throw new ArgumentOutOfRangeException(nameof(insertAt));
            }

            PowerPointSlide sourceSlide = sourceSlides[sourceIndex];
            sourceSlide.Save();
            Slide sourceSlideRoot = sourceSlide.SlidePart.Slide ?? throw new InvalidOperationException("Source slide is missing its slide definition.");

            SlideLayoutPart? sourceLayoutPart = sourceSlide.SlidePart.SlideLayoutPart;
            if (sourceLayoutPart == null) {
                throw new InvalidOperationException("Source slide does not have a layout to import.");
            }

            SlideLayoutPart? targetLayoutPart = FindMatchingLayout(sourceLayoutPart);
            if (targetLayoutPart == null) {
                SlideMasterPart sourceMasterPart = sourceLayoutPart.SlideMasterPart
                    ?? throw new InvalidOperationException("Source slide layout does not have a master.");

                Dictionary<SlideLayoutPart, SlideLayoutPart> layoutMap;
                CloneSlideMasterPart(sourceMasterPart, out layoutMap);

                if (!layoutMap.TryGetValue(sourceLayoutPart, out targetLayoutPart)) {
                    throw new InvalidOperationException("Failed to resolve the imported slide layout.");
                }
            }

            string slideRelId = GetNextSlideRelationshipId();
            SlidePart slidePart = _presentationPart.AddNewPart<SlidePart>(slideRelId);
            slidePart.Slide = (Slide)sourceSlideRoot.CloneNode(true);

            Dictionary<DataPart, MediaDataPart> mediaPartMap = new();
            CloneSlidePartRelationships(
                sourceSlide.SlidePart,
                slidePart,
                shouldShare: _ => false,
                includeDataParts: true,
                shouldSkip: part => part is SlideLayoutPart || part is NotesSlidePart,
                dataPartMap: mediaPartMap);

            string? layoutRelId = sourceSlide.SlidePart.GetIdOfPart(sourceLayoutPart);
            if (string.IsNullOrWhiteSpace(layoutRelId)) {
                layoutRelId = GetNextRelationshipId(slidePart);
            }

            slidePart.AddPart(targetLayoutPart, layoutRelId);

            SlideIdList slideIdList = PresentationRoot.SlideIdList ??= new SlideIdList();
            SlideId slideId = new() { Id = GetNextSlideId() };
            PowerPointUtils.SetRelationshipIdValue(slideId, slideRelId);
            InsertSlideId(slideIdList, slideId, targetIndex);
            AssignSlideToNearestSection(slideId.Id?.Value ?? throw new InvalidOperationException("Slide ID is missing."),
                targetIndex);

            PowerPointSlide imported = new(slidePart);
            imported.Hidden = sourceSlide.Hidden;

            if (sourceSlide.SlidePart.NotesSlidePart != null) {
                CloneImportedNotesSlidePart(sourceSlide.SlidePart, slidePart, mediaPartMap);
            }

            _slides.Insert(targetIndex, imported);
            PresentationRoot.Save();

            return imported;
        }

    }
}
