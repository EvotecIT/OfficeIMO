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
        private void InitializeDefaultParts() {
            // IMPORTANT: PowerPoint requires a very specific initialization pattern to avoid the repair dialog.
            // We must create an initial blank slide with relationship ID "rId2" and then create
            // the slide layout, slide master, and theme in a specific order.
            // DO NOT modify this initialization pattern or PowerPoint will show a repair dialog!
            PowerPointUtils.CreatePresentationParts(_document!, _presentationPart);
        }

        private void LoadExistingSlides() {
            if (PresentationRoot.SlideIdList != null) {
                foreach (SlideId slideId in PresentationRoot.SlideIdList.Elements<SlideId>()) {
                    string? relId = PowerPointUtils.GetRelationshipIdValue(slideId);
                    if (!string.IsNullOrEmpty(relId)) {
                        SlidePart slidePart = (SlidePart)_presentationPart.GetPartById(relId!);
                        _slides.Add(new PowerPointSlide(slidePart));
                    }
                }
            }
        }

        private string GetNextSlideRelationshipId() {
            var existingRelationships = new HashSet<string>(
                _presentationPart.Parts
                    .Select(p => p.RelationshipId)
                    .Union(_presentationPart.ExternalRelationships.Select(r => r.Id))
                    .Union(_presentationPart.HyperlinkRelationships.Select(r => r.Id))
                    .Where(id => !string.IsNullOrEmpty(id))
                    .Select(id => id!)
            );

            if (PresentationRoot.SlideIdList != null) {
                foreach (SlideId existingSlideId in PresentationRoot.SlideIdList.Elements<SlideId>()) {
                    string? relId = PowerPointUtils.GetRelationshipIdValue(existingSlideId);
                    if (!string.IsNullOrEmpty(relId)) {
                        existingRelationships.Add(relId!);
                    }
                }
            }

            int nextId = 1;
            string slideRelId;
            do {
                slideRelId = "rId" + nextId;
                nextId++;
            } while (existingRelationships.Contains(slideRelId));

            return slideRelId;
        }

        private uint GetNextSlideId() {
            uint maxId = 255;
            SlideIdList? slideIdList = PresentationRoot.SlideIdList;
            if (slideIdList != null && slideIdList.Elements<SlideId>().Any()) {
                maxId = slideIdList.Elements<SlideId>().Max(s => s.Id?.Value ?? 255);
            }

            return maxId >= 255 ? maxId + 1 : 256;
        }

        private SlideMasterPart GetSlideMasterPart(int masterIndex) {
            SlideMasterPart[] masters = _presentationPart.SlideMasterParts.ToArray();
            if (masterIndex < 0 || masterIndex >= masters.Length) {
                throw new ArgumentOutOfRangeException(nameof(masterIndex));
            }
            return masters[masterIndex];
        }

        private string GetNextSlideMasterRelationshipId() {
            var existingRelationships = new HashSet<string>(
                _presentationPart.Parts
                    .Select(p => p.RelationshipId)
                    .Union(_presentationPart.ExternalRelationships.Select(r => r.Id))
                    .Union(_presentationPart.HyperlinkRelationships.Select(r => r.Id))
                    .Where(id => !string.IsNullOrEmpty(id))
                    .Select(id => id!)
            );

            if (PresentationRoot.SlideMasterIdList != null) {
                foreach (SlideMasterId existingMasterId in PresentationRoot.SlideMasterIdList.Elements<SlideMasterId>()) {
                    string? existingRelId = PowerPointUtils.GetRelationshipIdValue(existingMasterId);
                    if (!string.IsNullOrEmpty(existingRelId)) {
                        existingRelationships.Add(existingRelId!);
                    }
                }
            }

            int nextId = 1;
            string masterRelId;
            do {
                masterRelId = "rId" + nextId;
                nextId++;
            } while (existingRelationships.Contains(masterRelId));

            return masterRelId;
        }

        private uint GetNextSlideMasterId() {
            SlideMasterIdList? slideMasterIdList = PresentationRoot.SlideMasterIdList;
            if (slideMasterIdList != null && slideMasterIdList.Elements<SlideMasterId>().Any()) {
                uint maxId = slideMasterIdList.Elements<SlideMasterId>().Max(s => s.Id?.Value ?? 0U);
                return maxId >= 2147483648U ? maxId + 1U : 2147483648U;
            }

            return 2147483648U;
        }

        private static void InsertSlideId(SlideIdList slideIdList, SlideId slideId, int index) {
            List<SlideId> ids = slideIdList.Elements<SlideId>().ToList();
            if (index < 0 || index > ids.Count) {
                throw new ArgumentOutOfRangeException(nameof(index));
            }

            ids.Insert(index, slideId);
            slideIdList.RemoveAllChildren();
            foreach (SlideId id in ids) {
                slideIdList.Append(id);
            }
        }

        private static string GetNextRelationshipId(OpenXmlPartContainer container) {
            var existingRelationships = new HashSet<string>(
                container.Parts.Select(p => p.RelationshipId)
                    .Concat(container.DataPartReferenceRelationships
                        .Select(r => r.Id))
                    .Concat(container.ExternalRelationships.Select(r => r.Id))
                    .Concat(container.HyperlinkRelationships.Select(r => r.Id))
                    .Where(id => !string.IsNullOrEmpty(id)),
                StringComparer.Ordinal);

            int nextId = 1;
            string relId;
            do {
                relId = "rId" + nextId;
                nextId++;
            } while (!existingRelationships.Add(relId));

            return relId;
        }

        private SlideMasterPart CloneSlideMasterPart(
            SlideMasterPart sourceMasterPart,
            out Dictionary<SlideLayoutPart, SlideLayoutPart> layoutMap,
            Func<SlidePart, SlidePart?>? slideResolver = null,
            bool skipUnresolvedSlideTargets = false,
            Dictionary<DataPart, MediaDataPart>? dataPartMap = null) {
            layoutMap = new Dictionary<SlideLayoutPart, SlideLayoutPart>();

            if (sourceMasterPart.SlideMaster == null) {
                throw new InvalidOperationException("Source slide master is missing.");
            }

            string masterRelId = GetNextSlideMasterRelationshipId();
            SlideMasterPart targetMasterPart = _presentationPart.AddNewPart<SlideMasterPart>(masterRelId);
            targetMasterPart.SlideMaster = (SlideMaster)sourceMasterPart.SlideMaster.CloneNode(true);

            foreach (var partPair in sourceMasterPart.Parts) {
                OpenXmlPart part = partPair.OpenXmlPart;
                string relId = partPair.RelationshipId;

                if (part is SlideLayoutPart sourceLayoutPart) {
                    SlideLayoutPart clonedLayout = CloneSlideLayoutPart(
                        sourceLayoutPart, targetMasterPart, relId,
                        slideResolver, skipUnresolvedSlideTargets,
                        dataPartMap);
                    layoutMap[sourceLayoutPart] = clonedLayout;
                    continue;
                }

                ClonePartRecursive(part, targetMasterPart, relId,
                    _ => false, includeDataParts: true,
                    dataPartMap: dataPartMap,
                    slideResolver: slideResolver,
                    skipUnresolvedSlideTargets:
                        skipUnresolvedSlideTargets);
            }

            CloneReferenceRelationships(sourceMasterPart, targetMasterPart,
                includeDataParts: true, dataPartMap);

            SlideMasterIdList slideMasterIdList = PresentationRoot.SlideMasterIdList ??= new SlideMasterIdList();
            SlideMasterId slideMasterId = new SlideMasterId { Id = GetNextSlideMasterId() };
            PowerPointUtils.SetRelationshipIdValue(slideMasterId, masterRelId);
            slideMasterIdList.Append(slideMasterId);
            PresentationRoot.Save();

            return targetMasterPart;
        }

        private static SlideLayoutPart CloneSlideLayoutPart(
            SlideLayoutPart sourceLayoutPart,
            SlideMasterPart targetMasterPart,
            string relationshipId,
            Func<SlidePart, SlidePart?>? slideResolver = null,
            bool skipUnresolvedSlideTargets = false,
            Dictionary<DataPart, MediaDataPart>? dataPartMap = null) {
            if (sourceLayoutPart.SlideLayout == null) {
                throw new InvalidOperationException("Source slide layout is missing.");
            }

            SlideLayoutPart targetLayoutPart = targetMasterPart.AddNewPart<SlideLayoutPart>(relationshipId);
            targetLayoutPart.SlideLayout = (SlideLayout)sourceLayoutPart.SlideLayout.CloneNode(true);

            CloneChildParts(
                sourceLayoutPart,
                targetLayoutPart,
                shouldSkip: part => part is SlideMasterPart,
                includeDataParts: true,
                dataPartMap: dataPartMap,
                slideResolver: slideResolver,
                skipUnresolvedSlideTargets:
                    skipUnresolvedSlideTargets);

            targetLayoutPart.AddPart(targetMasterPart);
            return targetLayoutPart;
        }

        private static void CloneChildParts(
            OpenXmlPart sourcePart,
            OpenXmlPart targetPart,
            Func<OpenXmlPart, bool> shouldSkip,
            bool includeDataParts,
            Dictionary<DataPart, MediaDataPart>? dataPartMap = null,
            Func<SlidePart, SlidePart?>? slideResolver = null,
            bool skipUnresolvedSlideTargets = false) {
            foreach (var childPair in sourcePart.Parts) {
                if (shouldSkip(childPair.OpenXmlPart)) {
                    continue;
                }

                ClonePartRecursive(childPair.OpenXmlPart, targetPart,
                    childPair.RelationshipId, _ => false, includeDataParts,
                    dataPartMap, slideResolver,
                    skipUnresolvedSlideTargets);
            }

            CloneReferenceRelationships(sourcePart, targetPart, includeDataParts, dataPartMap);
        }

        private static void CloneSlidePartRelationships(
            SlidePart source,
            SlidePart target,
            Func<OpenXmlPart, bool> shouldShare,
            bool includeDataParts,
            Func<OpenXmlPart, bool>? shouldSkip = null,
            Dictionary<DataPart, MediaDataPart>? dataPartMap = null,
            Func<SlidePart, SlidePart?>? slideResolver = null,
            bool skipUnresolvedSlideTargets = false) {
            foreach (var partPair in source.Parts) {
                if (shouldSkip != null && shouldSkip(partPair.OpenXmlPart)) {
                    continue;
                }

                ClonePartRecursive(partPair.OpenXmlPart, target,
                    partPair.RelationshipId, shouldShare, includeDataParts,
                    dataPartMap, slideResolver,
                    skipUnresolvedSlideTargets);
            }

            CloneReferenceRelationships(source, target, includeDataParts, dataPartMap);
        }

        private static void ClonePartRecursive(
            OpenXmlPart sourcePart,
            OpenXmlPartContainer targetContainer,
            string relationshipId,
            Func<OpenXmlPart, bool> shouldShare,
            bool includeDataParts,
            Dictionary<DataPart, MediaDataPart>? dataPartMap = null,
            Func<SlidePart, SlidePart?>? slideResolver = null,
            bool skipUnresolvedSlideTargets = false) {
            if (sourcePart is SlidePart sourceSlide
                && slideResolver != null) {
                SlidePart? targetSlide = slideResolver(sourceSlide);
                if (targetSlide == null) {
                    if (skipUnresolvedSlideTargets) {
                        RemoveInternalSlideLinkMarkup(targetContainer,
                            relationshipId);
                        return;
                    }
                    throw new InvalidDataException(
                        "An imported internal slide target is not present in the import closure.");
                }
                AddExistingPart(targetContainer, targetSlide, relationshipId);
                return;
            }
            if (shouldShare(sourcePart)) {
                AddExistingPart(targetContainer, sourcePart, relationshipId);
                return;
            }

            OpenXmlPart newPart = sourcePart is ExtendedPart extendedPart
                ? targetContainer.AddExtendedPart(extendedPart.RelationshipType, extendedPart.ContentType, relationshipId)
                : AddNewPartWithContentType(targetContainer, sourcePart, relationshipId);

            CopyPartData(sourcePart, newPart);
            CloneReferenceRelationships(sourcePart, newPart, includeDataParts, dataPartMap);

            foreach (var childPair in sourcePart.Parts) {
                ClonePartRecursive(childPair.OpenXmlPart, newPart,
                    childPair.RelationshipId, shouldShare,
                    includeDataParts, dataPartMap, slideResolver,
                    skipUnresolvedSlideTargets);
            }
        }

        private static void RemoveInternalSlideLinkMarkup(
            OpenXmlPartContainer targetContainer,
            string relationshipId) {
            if (targetContainer is not OpenXmlPart targetPart
                || targetPart.RootElement == null) return;
            A.HyperlinkType[] links = targetPart.RootElement
                .Descendants<A.HyperlinkType>()
                .Where(link => string.Equals(link.Id?.Value,
                    relationshipId, StringComparison.Ordinal))
                .ToArray();
            foreach (A.HyperlinkType link in links) link.Remove();
            if (links.Length > 0) targetPart.RootElement.Save();
        }

        private static OpenXmlPart AddNewPartWithContentType(OpenXmlPartContainer container, OpenXmlPart sourcePart, string relationshipId) {
            MethodInfo method = AddNewPartWithContentTypeMethod.MakeGenericMethod(sourcePart.GetType());
            return (OpenXmlPart)method.Invoke(container, new object[] { sourcePart.ContentType, relationshipId })!;
        }

        private static OpenXmlPart AddExistingPart(OpenXmlPartContainer container, OpenXmlPart sourcePart, string relationshipId) {
            MethodInfo method = AddPartWithIdMethod.MakeGenericMethod(sourcePart.GetType());
            return (OpenXmlPart)method.Invoke(container, new object[] { sourcePart, relationshipId })!;
        }

        private static void CopyPartData(OpenXmlPart sourcePart, OpenXmlPart targetPart) {
            using Stream sourceStream = sourcePart.GetStream(FileMode.Open, FileAccess.Read);
            using Stream targetStream = targetPart.GetStream(FileMode.Create, FileAccess.Write);
            sourceStream.CopyTo(targetStream);
        }

        private static void CopyPartData(DataPart sourcePart, DataPart targetPart) {
            using Stream sourceStream = sourcePart.GetStream(FileMode.Open, FileAccess.Read);
            using Stream targetStream = targetPart.GetStream(FileMode.Create, FileAccess.Write);
            sourceStream.CopyTo(targetStream);
        }

        private static void CloneReferenceRelationships(
            OpenXmlPartContainer source,
            OpenXmlPartContainer target,
            bool includeDataParts,
            Dictionary<DataPart, MediaDataPart>? dataPartMap = null) {
            foreach (ExternalRelationship rel in source.ExternalRelationships) {
                target.AddExternalRelationship(rel.RelationshipType, rel.Uri, rel.Id);
            }

            foreach (HyperlinkRelationship rel in source.HyperlinkRelationships) {
                target.AddHyperlinkRelationship(rel.Uri, rel.IsExternal, rel.Id);
            }

            if (includeDataParts) {
                CloneDataPartReferenceRelationships(source, target, dataPartMap);
            }
        }

        private static void CloneDataPartReferenceRelationships(
            OpenXmlPartContainer source,
            OpenXmlPartContainer target,
            Dictionary<DataPart, MediaDataPart>? dataPartMap) {
            OpenXmlPackage? sourcePackage = GetPackage(source);
            OpenXmlPackage? targetPackage = GetPackage(target);
            bool samePackage = sourcePackage != null && targetPackage != null && ReferenceEquals(sourcePackage, targetPackage);

            foreach (DataPartReferenceRelationship rel in source.DataPartReferenceRelationships) {
                if (rel.DataPart is not MediaDataPart mediaPart) {
                    continue;
                }

                MediaDataPart targetMediaPart = mediaPart;
                if (!samePackage) {
                    if (targetPackage == null) {
                        throw new InvalidOperationException("Unable to resolve target package for media import.");
                    }

                    if (dataPartMap != null && dataPartMap.TryGetValue(mediaPart, out MediaDataPart? existing)) {
                        targetMediaPart = existing;
                    } else {
                        targetMediaPart = CreateMediaDataPart(targetPackage, mediaPart.ContentType);
                        CopyPartData(mediaPart, targetMediaPart);
                        dataPartMap?.Add(mediaPart, targetMediaPart);
                    }
                }

                if (rel is AudioReferenceRelationship) {
                    if (TryAddMediaReferenceRelationship(target, "AddAudioReferenceRelationship", targetMediaPart, rel.Id)) {
                        continue;
                    }
                } else if (rel is VideoReferenceRelationship) {
                    if (TryAddMediaReferenceRelationship(target, "AddVideoReferenceRelationship", targetMediaPart, rel.Id)) {
                        continue;
                    }
                } else {
                    if (TryAddMediaReferenceRelationship(target, "AddMediaReferenceRelationship", targetMediaPart, rel.Id)) {
                        continue;
                    }
                }

                if (!samePackage) {
                    throw new InvalidOperationException("Unable to add media reference relationship to the imported slide.");
                }
            }
        }

        private static bool TryAddMediaReferenceRelationship(OpenXmlPartContainer target, string methodName,
            MediaDataPart mediaPart, string relationshipId) {
            MethodInfo? method = target.GetType().GetMethod(methodName,
                new[] { typeof(MediaDataPart), typeof(string) });
            if (method == null) {
                return false;
            }

            method.Invoke(target, new object[] { mediaPart, relationshipId });
            return true;
        }

        private static OpenXmlPackage? GetPackage(OpenXmlPartContainer container) {
            if (container is OpenXmlPackage package) {
                return package;
            }

            if (container is OpenXmlPart part) {
                return part.OpenXmlPackage;
            }

            return null;
        }

        private static MediaDataPart CreateMediaDataPart(OpenXmlPackage targetPackage, string contentType) {
            if (TryInvokeCreateMediaDataPart(targetPackage, new[] { typeof(string) }, new object[] { contentType }, out MediaDataPart? mediaPart) &&
                mediaPart != null) {
                return mediaPart;
            }

            MediaDataPartType? mediaType = TryGetMediaDataPartType(contentType);
            if (mediaType.HasValue &&
                TryInvokeCreateMediaDataPart(targetPackage, new[] { typeof(MediaDataPartType) }, new object[] { mediaType.Value }, out mediaPart) &&
                mediaPart != null) {
                return mediaPart;
            }

            throw new InvalidOperationException($"Unable to create a media data part for content type '{contentType}'.");
        }

        private static bool TryInvokeCreateMediaDataPart(
            OpenXmlPackage targetPackage,
            Type[] parameterTypes,
            object[] args,
            out MediaDataPart? mediaPart) {
            mediaPart = null;
            MethodInfo? method = targetPackage.GetType().GetMethod("CreateMediaDataPart", parameterTypes);
            if (method == null) {
                return false;
            }

            mediaPart = (MediaDataPart?)method.Invoke(targetPackage, args);
            return mediaPart != null;
        }

        private static MediaDataPartType? TryGetMediaDataPartType(string contentType) {
            if (string.IsNullOrWhiteSpace(contentType)) {
                return null;
            }

            return contentType.ToLowerInvariant() switch {
                "audio/aiff" => MediaDataPartType.Aiff,
                "audio/x-aiff" => MediaDataPartType.Aiff,
                "audio/midi" => MediaDataPartType.Midi,
                "audio/x-midi" => MediaDataPartType.Midi,
                "audio/mpeg" => MediaDataPartType.Mp3,
                "audio/mp3" => MediaDataPartType.Mp3,
                "audio/wav" => MediaDataPartType.Wav,
                "audio/x-wav" => MediaDataPartType.Wav,
                "audio/x-ms-wma" => MediaDataPartType.Wma,
                "audio/wma" => MediaDataPartType.Wma,
                "audio/ogg" => MediaDataPartType.OggAudio,
                "application/ogg" => MediaDataPartType.OggAudio,
                "audio/mpegurl" => MediaDataPartType.MpegUrl,
                "application/vnd.ms-asf" => MediaDataPartType.Asx,
                "video/x-msvideo" => MediaDataPartType.Avi,
                "video/avi" => MediaDataPartType.Avi,
                "video/mpeg" => MediaDataPartType.MpegVideo,
                "video/mpg" => MediaDataPartType.Mpg,
                "video/mp4" => MediaDataPartType.MpegVideo,
                "video/quicktime" => MediaDataPartType.Quicktime,
                "video/x-ms-wmv" => MediaDataPartType.Wmv,
                "video/x-ms-wmx" => MediaDataPartType.Wmx,
                "video/x-ms-wvx" => MediaDataPartType.Wvx,
                "video/ogg" => MediaDataPartType.OggVideo,
                "video/vc1" => MediaDataPartType.VC1,
                _ => null
            };
        }

        private static bool ShouldSharePart(OpenXmlPart part) {
            return part is SlideLayoutPart || part is NotesMasterPart
                || part is SlidePart;
        }

        private static void RemapDuplicatedNotesSlideBacklink(
            SlidePart sourceSlidePart, SlidePart targetSlidePart) {
            NotesSlidePart? sourceNotesPart = sourceSlidePart.NotesSlidePart;
            NotesSlidePart? targetNotesPart = targetSlidePart.NotesSlidePart;
            if (sourceNotesPart == null || targetNotesPart == null) return;

            var referencedRelationshipIds = new HashSet<string>(
                (sourceNotesPart.NotesSlide == null
                    ? Enumerable.Empty<OpenXmlElement>()
                    : new OpenXmlElement[] { sourceNotesPart.NotesSlide }
                        .Concat(sourceNotesPart.NotesSlide.Descendants()))
                    .SelectMany(element => element.GetAttributes())
                    .Select(attribute => attribute.Value)
                    .OfType<string>()
                    .Where(value => value.Length > 0),
                StringComparer.Ordinal);
            string? backlinkId = sourceNotesPart.Parts
                .Where(pair => ReferenceEquals(pair.OpenXmlPart,
                    sourceSlidePart))
                .Select(pair => pair.RelationshipId)
                .FirstOrDefault(id => !referencedRelationshipIds.Contains(id));
            if (string.IsNullOrEmpty(backlinkId)
                || !targetNotesPart.TryGetPartById(backlinkId,
                    out OpenXmlPart? clonedBacklink)
                || clonedBacklink is not SlidePart) {
                return;
            }

            targetNotesPart.DeletePart(backlinkId);
            targetNotesPart.AddPart(targetSlidePart, backlinkId);
        }

        private void CloneImportedNotesSlidePart(
            SlidePart sourceSlidePart,
            SlidePart targetSlidePart,
            Dictionary<DataPart, MediaDataPart> mediaPartMap,
            Func<SlidePart, SlidePart?>? slideResolver = null,
            bool skipUnresolvedSlideTargets = false) {
            NotesSlidePart? sourceNotesPart = sourceSlidePart.NotesSlidePart;
            if (sourceNotesPart == null) {
                return;
            }

            NotesSlidePart targetNotesPart = targetSlidePart.AddNewPart<NotesSlidePart>(GetNextRelationshipId(targetSlidePart));
            if (sourceNotesPart.NotesSlide != null) {
                targetNotesPart.NotesSlide = (NotesSlide)sourceNotesPart.NotesSlide.CloneNode(true);
            }

            CloneChildParts(
                sourceNotesPart,
                targetNotesPart,
                shouldSkip: part => part is NotesMasterPart,
                includeDataParts: true,
                dataPartMap: mediaPartMap,
                slideResolver: slideResolver,
                skipUnresolvedSlideTargets:
                    skipUnresolvedSlideTargets);

            NotesMasterPart targetNotesMasterPart = PowerPointUtils.EnsureNotesMasterPart(_presentationPart);
            if (!targetNotesPart.Parts.Any(pair => ReferenceEquals(pair.OpenXmlPart, targetNotesMasterPart))) {
                targetNotesPart.AddPart(targetNotesMasterPart);
            }
        }

    }
}
