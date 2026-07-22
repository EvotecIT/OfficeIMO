using System.Collections.ObjectModel;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using OfficeIMO.PowerPoint.LegacyPpt.Model;
using OfficeIMO.PowerPoint.LegacyPpt.Write;
using A = DocumentFormat.OpenXml.Drawing;
using P = DocumentFormat.OpenXml.Presentation;

namespace OfficeIMO.PowerPoint.LegacyPpt.Internal {
    /// <summary>Links projected Open XML slides and shapes back to their original binary persist records.</summary>
    internal sealed class LegacyPptProjectionMap {
        private readonly IReadOnlyDictionary<string, LegacyPptSlideProjection> _slidesByPartUri;
        private readonly IReadOnlyDictionary<uint, LegacyPptSlideProjection> _slidesByLegacyId;
        private readonly IReadOnlyDictionary<string, uint> _masterIdsByLayoutPartUri;
        private readonly IReadOnlyDictionary<string,
            LegacyPptOrdinaryLayoutProjection> _ordinaryLayoutsByPartUri;
        private readonly IReadOnlyDictionary<string, LegacyPptMasterProjection> _mastersByPartUri;
        private readonly IReadOnlyDictionary<string, LegacyPptMasterProjection>
            _titleMastersByPartUri;
        private readonly IReadOnlyDictionary<string, LegacyPptMasterProjection>
            _specialMastersByPartUri;
        private readonly ISet<string> _masterThemePartUris;
        private readonly ISet<string> _olePartUris;

        private LegacyPptProjectionMap(IReadOnlyList<LegacyPptSlideProjection> slides,
            IReadOnlyDictionary<string, uint> masterIdsByLayoutPartUri,
            IReadOnlyList<LegacyPptOrdinaryLayoutProjection> ordinaryLayouts,
            IReadOnlyList<LegacyPptMasterProjection> masters,
            IReadOnlyList<LegacyPptMasterProjection> titleMasters,
            IReadOnlyList<LegacyPptMasterProjection> specialMasters,
            IReadOnlyList<LegacyPptHyperlink> hyperlinks,
            IReadOnlyList<LegacyPptCustomShow> customShows,
            bool customShowsAreEditable,
            IReadOnlyList<LegacyPptSound> sounds, uint? soundIdSeed,
            LegacyPptPropertySetProjection propertySets,
            LegacyPptVbaProjectProjection vbaProject) {
            Slides = new ReadOnlyCollection<LegacyPptSlideProjection>(slides.ToArray());
            _slidesByPartUri = new ReadOnlyDictionary<string, LegacyPptSlideProjection>(slides.ToDictionary(
                slide => slide.SlidePartUri, StringComparer.Ordinal));
            _slidesByLegacyId = new ReadOnlyDictionary<uint, LegacyPptSlideProjection>(slides.ToDictionary(
                slide => slide.SlideId));
            _masterIdsByLayoutPartUri = new ReadOnlyDictionary<string, uint>(
                masterIdsByLayoutPartUri.ToDictionary(pair => pair.Key, pair => pair.Value,
                    StringComparer.Ordinal));
            _ordinaryLayoutsByPartUri = new ReadOnlyDictionary<string,
                LegacyPptOrdinaryLayoutProjection>(ordinaryLayouts
                    .ToDictionary(layout => layout.PartUri,
                        StringComparer.Ordinal));
            Masters = new ReadOnlyCollection<LegacyPptMasterProjection>(masters.ToArray());
            _mastersByPartUri = new ReadOnlyDictionary<string, LegacyPptMasterProjection>(
                masters.ToDictionary(master => master.MasterPartUri, StringComparer.Ordinal));
            TitleMasters = new ReadOnlyCollection<LegacyPptMasterProjection>(
                titleMasters.ToArray());
            _titleMastersByPartUri = new ReadOnlyDictionary<string,
                LegacyPptMasterProjection>(titleMasters.ToDictionary(
                    master => master.MasterPartUri, StringComparer.Ordinal));
            SpecialMasters = new ReadOnlyCollection<LegacyPptMasterProjection>(
                specialMasters.ToArray());
            _specialMastersByPartUri = new ReadOnlyDictionary<string,
                LegacyPptMasterProjection>(specialMasters.ToDictionary(
                    master => master.MasterPartUri, StringComparer.Ordinal));
            _masterThemePartUris = new HashSet<string>(masters
                .Concat(titleMasters).Concat(specialMasters)
                .Select(master => master.ThemePartUri)
                .Concat(slides.Select(slide => slide.ThemePartUri))
                .Concat(slides.Select(slide => slide.Notes?.ThemePartUri))
                .Where(uri => uri != null).Cast<string>(), StringComparer.Ordinal);
            _olePartUris = new HashSet<string>(slides
                .SelectMany(slide => slide.Shapes)
                .Select(shape => shape.OleObject?.EmbeddedPartUri)
                .Where(uri => uri != null).Cast<string>(),
                StringComparer.Ordinal);
            Hyperlinks = new ReadOnlyCollection<LegacyPptHyperlink>(hyperlinks.ToArray());
            CustomShows = new ReadOnlyCollection<LegacyPptCustomShow>(
                customShows.ToArray());
            CanEditCustomShows = customShowsAreEditable
                && customShows.All(show => show.IsEditable)
                && customShows.Select(show => show.Name)
                    .Distinct(StringComparer.Ordinal).Count() == customShows.Count;
            Sounds = new ReadOnlyCollection<LegacyPptSound>(sounds.ToArray());
            SoundIdSeed = soundIdSeed;
            PropertySets = propertySets
                ?? throw new ArgumentNullException(nameof(propertySets));
            VbaProject = vbaProject
                ?? throw new ArgumentNullException(nameof(vbaProject));
        }

        internal IReadOnlyList<LegacyPptSlideProjection> Slides { get; }

        internal IReadOnlyList<LegacyPptMasterProjection> Masters { get; }

        internal IReadOnlyList<LegacyPptMasterProjection> TitleMasters { get; }

        internal IReadOnlyList<LegacyPptMasterProjection> SpecialMasters { get; }

        internal IReadOnlyList<LegacyPptHyperlink> Hyperlinks { get; }

        internal IReadOnlyList<LegacyPptCustomShow> CustomShows { get; }

        internal bool CanEditCustomShows { get; }

        internal IReadOnlyList<LegacyPptSound> Sounds { get; }

        internal uint? SoundIdSeed { get; }

        internal LegacyPptPropertySetProjection PropertySets { get; }

        internal LegacyPptVbaProjectProjection VbaProject { get; }

        internal bool IsProjectedOlePart(string partUri) =>
            partUri != null && _olePartUris.Contains(partUri);

        internal bool TryGetSlide(PowerPointSlide slide, out LegacyPptSlideProjection? projection) {
            if (slide == null) throw new ArgumentNullException(nameof(slide));
            return _slidesByPartUri.TryGetValue(slide.SlidePart.Uri.ToString(), out projection);
        }

        internal bool TryGetSlide(uint legacySlideId,
            out LegacyPptSlideProjection? projection) =>
            _slidesByLegacyId.TryGetValue(legacySlideId, out projection);

        internal bool TryGetMasterId(PowerPointSlide slide, out uint masterId) {
            if (slide == null) throw new ArgumentNullException(nameof(slide));
            masterId = 0;
            SlideLayoutPart? layoutPart = slide.SlidePart.SlideLayoutPart;
            return layoutPart != null
                && _masterIdsByLayoutPartUri.TryGetValue(layoutPart.Uri.ToString(), out masterId);
        }

        internal bool IsProjectedLayoutPart(string partUri) =>
            partUri != null && _masterIdsByLayoutPartUri.ContainsKey(partUri);

        internal bool IsEditableProjectedOrdinaryLayoutPart(string partUri) =>
            partUri != null
            && _ordinaryLayoutsByPartUri.ContainsKey(partUri);

        internal bool OrdinaryLayoutMatches(SlideLayoutPart layoutPart) {
            if (layoutPart == null) throw new ArgumentNullException(
                nameof(layoutPart));
            return _ordinaryLayoutsByPartUri.TryGetValue(
                    layoutPart.Uri.ToString(), out LegacyPptOrdinaryLayoutProjection?
                        projection)
                && projection.ShapeTreeMatches(layoutPart);
        }

        internal bool OrdinaryLayoutTypeMatches(
            SlideLayoutPart layoutPart) {
            if (layoutPart == null) throw new ArgumentNullException(
                nameof(layoutPart));
            return _ordinaryLayoutsByPartUri.TryGetValue(
                    layoutPart.Uri.ToString(), out LegacyPptOrdinaryLayoutProjection?
                        projection)
                && projection.TypeMatches(layoutPart);
        }

        internal bool TryGetOrdinaryLayoutAddedShapeIds(
            SlideLayoutPart layoutPart,
            out IReadOnlyList<uint> addedShapeIds) {
            if (layoutPart == null) throw new ArgumentNullException(
                nameof(layoutPart));
            addedShapeIds = Array.Empty<uint>();
            return _ordinaryLayoutsByPartUri.TryGetValue(
                    layoutPart.Uri.ToString(), out LegacyPptOrdinaryLayoutProjection?
                        projection)
                && projection.TryGetAddedShapeIds(layoutPart,
                    out addedShapeIds);
        }

        internal bool IsEditableProjectedLayoutBackgroundPart(string partUri) =>
            partUri != null && !_titleMastersByPartUri.ContainsKey(partUri)
            && Slides.Any(slide => !slide.HasExplicitBackground
                && string.Equals(slide.LayoutPartUri, partUri,
                    StringComparison.Ordinal));

        internal bool IsEditableProjectedLayoutThemePart(string partUri) =>
            partUri != null && !_titleMastersByPartUri.ContainsKey(partUri)
            && Slides.Any(slide => slide.ThemePartUri == null
                && string.Equals(slide.LayoutPartUri, partUri,
                    StringComparison.Ordinal));

        internal bool TryGetMaster(SlideMasterPart masterPart,
            out LegacyPptMasterProjection? projection) {
            if (masterPart == null) throw new ArgumentNullException(nameof(masterPart));
            return _mastersByPartUri.TryGetValue(masterPart.Uri.ToString(), out projection);
        }

        internal bool IsProjectedMasterPart(string partUri) =>
            partUri != null && _mastersByPartUri.ContainsKey(partUri);

        internal bool IsProjectedMasterThemePart(string partUri) =>
            partUri != null && _masterThemePartUris.Contains(partUri);

        internal bool TryGetSpecialMaster(OpenXmlPart masterPart,
            out LegacyPptMasterProjection? projection) {
            if (masterPart == null) throw new ArgumentNullException(nameof(masterPart));
            return _specialMastersByPartUri.TryGetValue(
                masterPart.Uri.ToString(), out projection);
        }

        internal bool TryGetTitleMaster(SlideLayoutPart masterPart,
            out LegacyPptMasterProjection? projection) {
            if (masterPart == null) throw new ArgumentNullException(nameof(masterPart));
            return _titleMastersByPartUri.TryGetValue(
                masterPart.Uri.ToString(), out projection);
        }

        internal static LegacyPptProjectionMap Create(PowerPointPresentation presentation,
            LegacyPptPresentation legacy) {
            if (presentation == null) throw new ArgumentNullException(nameof(presentation));
            if (legacy == null) throw new ArgumentNullException(nameof(legacy));
            if (presentation.Slides.Count != legacy.Slides.Count) {
                throw new InvalidDataException("The projected slide count does not match the binary source slide count.");
            }

            var slides = new List<LegacyPptSlideProjection>(legacy.Slides.Count);
            var inheritedBackgrounds = new Dictionary<string, string>(
                StringComparer.Ordinal);
            var inheritedThemes = new Dictionary<string, string>(
                StringComparer.Ordinal);
            var projectableSoundIds = new HashSet<uint>(legacy.Sounds.Where(sound =>
                sound.HasData && sound.ContentType != null).Select(sound => sound.Id));
            for (int slideIndex = 0; slideIndex < legacy.Slides.Count; slideIndex++) {
                PowerPointSlide projectedSlide = presentation.Slides[slideIndex];
                LegacyPptSlide sourceSlide = legacy.Slides[slideIndex];
                PowerPointShape[] projectedShapes = projectedSlide.Shapes.ToArray();
                LegacyPptShape[] sourceShapes = sourceSlide.Shapes
                    .Where(shape => shape.Kind != LegacyPptShapeKind.Unsupported)
                    .ToArray();
                if (projectedShapes.Length != sourceShapes.Length) {
                    throw new InvalidDataException(
                        $"Projected slide {slideIndex + 1} has {projectedShapes.Length} editable shapes, "
                        + $"but the binary source exposed {sourceShapes.Length}.");
                }

                var shapes = new List<LegacyPptShapeProjection>();
                AddShapeProjectionTree(shapes, sourceShapes,
                    projectedShapes, projectableSoundIds,
                    $"slide {slideIndex + 1}", projectedSlide.SlidePart);
                bool hasExplicitBackground = projectedSlide.SlidePart.Slide?
                    .CommonSlideData?.Background != null;
                slides.Add(new LegacyPptSlideProjection(projectedSlide.SlidePart.Uri.ToString(),
                    projectedSlide.SlidePart.SlideLayoutPart?.Uri.ToString(),
                    sourceSlide.PersistId, sourceSlide.SlideId, sourceSlide.MasterId,
                    sourceSlide.LayoutType,
                    sourceSlide.LayoutPlaceholderTypes.Select(value =>
                        unchecked((byte)value)).ToArray(),
                    sourceSlide.Hidden, sourceSlide.FollowsMasterObjects,
                    sourceSlide.HeaderFooter,
                    hasExplicitBackground,
                    CreateSlideBackgroundFingerprint(projectedSlide,
                        hasExplicitBackground, inheritedBackgrounds),
                    projectedSlide.SlidePart.ThemeOverridePart?.Uri.ToString(),
                    CreateSlideThemeFingerprint(projectedSlide,
                        inheritedThemes),
                    LegacyPptSlideProjection.CreateClassicColorFingerprints(
                        projectedSlide),
                    sourceSlide.Transition, sourceSlide.Comments, shapes,
                    sourceSlide.NotesPage == null
                        ? null
                        : CreateNotesProjection(projectedSlide,
                            sourceSlide.NotesPage)));
            }
            IReadOnlyList<LegacyPptMasterProjection> titleMasters =
                CreateTitleMasterProjections(presentation, legacy);
            return new LegacyPptProjectionMap(slides,
                CreateLayoutMasterMap(presentation, legacy),
                LegacyPptOrdinaryLayoutProjection.Create(presentation,
                    titleMasters),
                CreateMasterProjections(presentation, legacy),
                titleMasters,
                CreateSpecialMasterProjections(presentation, legacy),
                legacy.Hyperlinks, legacy.CustomShows,
                legacy.CustomShowsAreEditable, legacy.Sounds,
                legacy.SoundIdSeed,
                LegacyPptPropertySetCodec.CreateProjection(presentation,
                    legacy.Package),
                LegacyPptVbaProjectProjection.Create(legacy.VbaProject));
        }

        private static string CreateSlideBackgroundFingerprint(
            PowerPointSlide slide, bool hasExplicitBackground,
            IDictionary<string, string> inheritedFingerprints) {
            SlideLayoutPart? layoutPart = slide.SlidePart.SlideLayoutPart;
            if (hasExplicitBackground || layoutPart == null) {
                return LegacyPptSlideProjection.CreateBackgroundFingerprint(
                    slide);
            }
            string key = layoutPart.Uri.ToString();
            if (!inheritedFingerprints.TryGetValue(key,
                    out string? fingerprint)) {
                fingerprint = LegacyPptSlideProjection
                    .CreateBackgroundFingerprint(slide);
                inheritedFingerprints.Add(key, fingerprint);
            }
            return fingerprint;
        }

        private static string CreateSlideThemeFingerprint(
            PowerPointSlide slide,
            IDictionary<string, string> inheritedFingerprints) {
            SlideLayoutPart? layoutPart = slide.SlidePart.SlideLayoutPart;
            if (slide.SlidePart.ThemeOverridePart != null
                || layoutPart == null) {
                return LegacyPptSlideProjection.CreateThemeFingerprint(slide);
            }
            string key = layoutPart.Uri.ToString();
            if (!inheritedFingerprints.TryGetValue(key,
                    out string? fingerprint)) {
                fingerprint = LegacyPptSlideProjection
                    .CreateThemeFingerprint(slide);
                inheritedFingerprints.Add(key, fingerprint);
            }
            return fingerprint;
        }

        private static void AddShapeProjectionTree(
            ICollection<LegacyPptShapeProjection> result,
            IReadOnlyList<LegacyPptShape> sourceShapes,
            IReadOnlyList<PowerPointShape> projectedShapes,
            ISet<uint> projectableSoundIds, string ownerName,
            OpenXmlPart ownerPart) {
            if (sourceShapes.Count != projectedShapes.Count) {
                throw new InvalidDataException(
                    $"Projected {ownerName} has {projectedShapes.Count} editable shapes, but the binary source exposed {sourceShapes.Count}.");
            }
            for (int index = 0; index < sourceShapes.Count; index++) {
                LegacyPptShape sourceShape = sourceShapes[index];
                PowerPointShape projectedShape = projectedShapes[index];
                uint? openXmlShapeId = projectedShape.Id;
                if (!openXmlShapeId.HasValue) {
                    throw new InvalidDataException(
                        $"Projected {ownerName}, shape {index + 1} has no Open XML shape id.");
                }
                string? textFormattingFingerprint =
                    (sourceShape.TextBody.HasStyleRecord
                        || sourceShape.TextBody.HasStyle9Record
                        || sourceShape.TextBody.HasFieldRecords
                        || sourceShape.TextBody.HasInteractions
                        || sourceShape.TextBody.HasTextSpecialInfoRecord)
                    && projectedShape.Element is P.Shape projectedTextShape
                        ? LegacyPptTextProjection.CreateFormattingFingerprint(
                            projectedTextShape.TextBody, ownerPart)
                        : null;
                result.Add(new LegacyPptShapeProjection(
                    openXmlShapeId.Value, sourceShape.ShapeId,
                    sourceShape.RecordOffset, sourceShape.Kind,
                    sourceShape.Bounds, sourceShape.Text,
                    sourceShape.Placeholder,
                    textFormattingFingerprint, sourceShape.Interactions,
                    sourceShape.TextBody.Interactions,
                    sourceShape.Animation, projectableSoundIds,
                    sourceShape.Kind == LegacyPptShapeKind.TextBox
                        && !sourceShape.TextBody.IsStyleTruncated
                        && !sourceShape.TextBody.IsStyle9Truncated
                        && !sourceShape.TextBody.IsFieldDataMalformed
                        && !sourceShape.TextBody
                            .HasUnprojectedCharacterFormatting
                        && !sourceShape.TextBody
                            .HasUnprojectedParagraphFormatting
                        && !sourceShape.TextBody
                            .HasUnprojectedTextSpecialInfo
                        && !sourceShape.TextBody
                            .IsTextSpecialInfoTruncated
                        && !sourceShape.TextBody.IsRulerTruncated,
                    sourceShape.Kind == LegacyPptShapeKind.TextBox
                        ? LegacyPptShapeProjection
                            .CreateTextFrameFingerprint(projectedShape)
                        : null,
                    sourceShape.Kind == LegacyPptShapeKind.TextBox
                        && sourceShape.TextFrame
                            .CanRewriteProjectedProperties,
                    LegacyPptShapeProjection
                        .CreateShapeTransformFingerprint(projectedShape),
                    sourceShape.OfficeArtShapeType is 2 or 23
                        ? LegacyPptShapeProjection
                            .CreateShapeGeometryFingerprint(projectedShape)
                        : null,
                    sourceShape.Kind == LegacyPptShapeKind.Group
                        ? LegacyPptShapeProjection
                            .CreateGroupCoordinateFingerprint(projectedShape)
                        : null,
                    sourceShape.Style.CanRewriteProjectedVisualStyle
                        ? LegacyPptShapeProjection
                            .CreateShapeVisualStyleFingerprint(projectedShape)
                        : null,
                    LegacyPptShapeProjection
                        .CreatePictureFormattingFingerprint(projectedShape),
                    sourceShape.Style.CanRewriteHiddenState
                        ? LegacyPptShapeProjection
                            .CreateShapeVisibilityFingerprint(projectedShape)
                        : null,
                    LegacyPptShapeProjection
                        .CreateShapeMetadataFingerprint(projectedShape),
                    sourceShape.Metadata.CanRewrite,
                    sourceShape.OleObject != null
                        && projectedShape is PowerPointOleObject projectedOle
                        ? LegacyPptOleObjectProjection.Create(
                            sourceShape.OleObject, projectedOle)
                        : null));
                if (sourceShape.Kind != LegacyPptShapeKind.Group) continue;
                if (projectedShape is not PowerPointGroupShape group) {
                    throw new InvalidDataException(
                        $"Projected {ownerName}, shape {index + 1} is not the expected group shape.");
                }
                IReadOnlyList<PowerPointShape> projectedChildren =
                    LegacyPptWriter.ReadGroupChildrenForWrite(group,
                        out string? childReason);
                if (childReason != null) {
                    throw new InvalidDataException(childReason);
                }
                LegacyPptShape[] sourceChildren = sourceShape.Children
                    .Where(PowerPointPresentation
                        .CanCreateLegacyOpenXmlShape)
                    .ToArray();
                AddShapeProjectionTree(result, sourceChildren,
                    projectedChildren, projectableSoundIds,
                    $"{ownerName}, group shape {index + 1}", ownerPart);
            }
        }

        private static LegacyPptNotesProjection CreateNotesProjection(
            PowerPointSlide projectedSlide, LegacyPptNotesPage source) {
            NotesSlidePart part = projectedSlide.SlidePart.NotesSlidePart
                ?? throw new InvalidDataException(
                    "The projected binary notes page has no notes-slide part.");
            return new LegacyPptNotesProjection(source.PersistId,
                source.NotesId, source.Text, source.FollowsMasterObjects,
                part.ThemeOverridePart?.Uri.ToString(),
                LegacyPptNotesProjection.CreateThemeFingerprint(part),
                LegacyPptNotesProjection.CreateClassicColorFingerprints(part),
                LegacyPptNotesProjection.CreateBackgroundFingerprint(part));
        }

        private static IReadOnlyList<LegacyPptMasterProjection> CreateMasterProjections(
            PowerPointPresentation presentation, LegacyPptPresentation legacy) {
            SlideMasterPart[] masterParts = presentation.OpenXmlDocument.PresentationPart?
                .SlideMasterParts.ToArray() ?? Array.Empty<SlideMasterPart>();
            LegacyPptMaster[] sourceMasters = legacy.Masters
                .Where(master => master.IsMainMaster).ToArray();
            if (masterParts.Length != sourceMasters.Length) {
                throw new InvalidDataException(
                    "The projected slide-master count does not match the binary source master count.");
            }

            var result = new List<LegacyPptMasterProjection>(masterParts.Length);
            for (int index = 0; index < masterParts.Length; index++) {
                SlideMasterPart masterPart = masterParts[index];
                ThemePart themePart = masterPart.ThemePart
                    ?? throw new InvalidDataException(
                        $"Projected slide master {index + 1} has no theme part.");
                IReadOnlyList<LegacyPptShapeProjection> shapes =
                    CreateMasterShapeProjections(sourceMasters[index].Shapes,
                        LegacyPptWriter.ReadMasterShapesForWrite(masterPart,
                            out _), $"slide master {index + 1}",
                        masterPart);
                result.Add(new LegacyPptMasterProjection(
                    masterPart.Uri.ToString(), themePart.Uri.ToString(),
                    sourceMasters[index].PersistId,
                    LegacyPptMasterProjection.CreateThemeFingerprint(masterPart),
                    LegacyPptMasterProjection.CreateClassicColorFingerprints(masterPart),
                    LegacyPptMasterProjection.CreateBackgroundFingerprint(masterPart),
                    followsMasterObjects: null,
                    shapes,
                    LegacyPptMasterProjection
                        .CreateTextStylesFingerprint(masterPart),
                    CanEditMasterTextStyles(sourceMasters[index])));
            }
            return result;
        }

        private static bool CanEditMasterTextStyles(
            LegacyPptMaster source) {
            LegacyPptTextType[] projectedTypes = {
                LegacyPptTextType.Title,
                LegacyPptTextType.Body,
                LegacyPptTextType.Other
            };
            foreach (LegacyPptTextType type in projectedTypes) {
                LegacyPptTextMasterStyle[] matches = source.TextMasterStyles
                    .Where(style => style.TextType == type).ToArray();
                if (matches.Length != 1 || matches[0].IsTruncated
                    || matches[0].HasUnprojectedFormatting) return false;
            }
            return true;
        }

        private static IReadOnlyList<LegacyPptMasterProjection>
            CreateTitleMasterProjections(PowerPointPresentation presentation,
                LegacyPptPresentation legacy) {
            SlideLayoutPart[] layouts = presentation.OpenXmlDocument
                .PresentationPart?.SlideMasterParts
                .SelectMany(master => master.SlideLayoutParts).ToArray()
                ?? Array.Empty<SlideLayoutPart>();
            var result = new List<LegacyPptMasterProjection>();
            foreach (LegacyPptMaster source in legacy.Masters.Where(master =>
                         !master.IsMainMaster)) {
                string expectedName =
                    $"Binary Title Master {source.MasterId:X8}";
                SlideLayoutPart part = layouts.SingleOrDefault(layout =>
                        string.Equals(layout.SlideLayout?.CommonSlideData?.Name?.Value,
                            expectedName, StringComparison.Ordinal))
                    ?? throw new InvalidDataException(
                        $"The projected title master 0x{source.MasterId:X8} has no exact layout part.");
                result.Add(new LegacyPptMasterProjection(
                    part.Uri.ToString(), part.ThemeOverridePart?.Uri.ToString(),
                    source.PersistId,
                    LegacyPptMasterProjection.CreateThemeFingerprint(part),
                    LegacyPptMasterProjection.CreateClassicColorFingerprints(part),
                    LegacyPptMasterProjection.CreateBackgroundFingerprint(part),
                    source.FollowsMasterObjects,
                    CreateMasterShapeProjections(source.Shapes,
                        LegacyPptWriter.ReadMasterShapesForWrite(part, out _),
                        $"title master 0x{source.MasterId:X8}", part)));
            }
            return result;
        }

        private static IReadOnlyList<LegacyPptMasterProjection>
            CreateSpecialMasterProjections(PowerPointPresentation presentation,
                LegacyPptPresentation legacy) {
            PresentationPart? presentationPart = presentation.OpenXmlDocument
                .PresentationPart;
            var result = new List<LegacyPptMasterProjection>(2);
            if (legacy.NotesMaster != null) {
                NotesMasterPart part = presentationPart?.NotesMasterPart
                    ?? throw new InvalidDataException(
                        "The projected presentation has no notes-master part.");
                ThemePart themePart = part.ThemePart
                    ?? throw new InvalidDataException(
                        "The projected notes master has no theme part.");
                result.Add(new LegacyPptMasterProjection(
                    part.Uri.ToString(), themePart.Uri.ToString(),
                    legacy.NotesMaster.PersistId,
                    LegacyPptMasterProjection.CreateThemeFingerprint(part),
                    LegacyPptMasterProjection.CreateClassicColorFingerprints(part),
                    LegacyPptMasterProjection.CreateBackgroundFingerprint(part),
                    followsMasterObjects: null,
                    CreateMasterShapeProjections(legacy.NotesMaster.Shapes,
                        LegacyPptWriter.ReadMasterShapesForWrite(part, out _),
                        "notes master", part)));
            }
            if (legacy.HandoutMaster != null) {
                HandoutMasterPart part = presentationPart?.HandoutMasterPart
                    ?? throw new InvalidDataException(
                        "The projected presentation has no handout-master part.");
                ThemePart themePart = part.ThemePart
                    ?? throw new InvalidDataException(
                        "The projected handout master has no theme part.");
                result.Add(new LegacyPptMasterProjection(
                    part.Uri.ToString(), themePart.Uri.ToString(),
                    legacy.HandoutMaster.PersistId,
                    LegacyPptMasterProjection.CreateThemeFingerprint(part),
                    LegacyPptMasterProjection.CreateClassicColorFingerprints(part),
                    LegacyPptMasterProjection.CreateBackgroundFingerprint(part),
                    followsMasterObjects: null,
                    CreateMasterShapeProjections(legacy.HandoutMaster.Shapes,
                        LegacyPptWriter.ReadMasterShapesForWrite(part, out _),
                        "handout master", part)));
            }
            return result;
        }

        private static IReadOnlyList<LegacyPptShapeProjection>
            CreateMasterShapeProjections(IReadOnlyList<LegacyPptShape> source,
                IReadOnlyList<PowerPointShape> projected, string ownerName,
                OpenXmlPart ownerPart) {
            LegacyPptShape[] sourceShapes = source.Where(shape =>
                PowerPointPresentation.CanCreateLegacyOpenXmlShape(shape))
                .ToArray();
            var result = new List<LegacyPptShapeProjection>();
            AddShapeProjectionTree(result, sourceShapes, projected,
                new HashSet<uint>(), ownerName, ownerPart);
            return result;
        }

        private static IReadOnlyDictionary<string, uint> CreateLayoutMasterMap(
            PowerPointPresentation presentation, LegacyPptPresentation legacy) {
            SlideMasterPart[] masterParts = presentation.OpenXmlDocument.PresentationPart?
                .SlideMasterParts.ToArray() ?? Array.Empty<SlideMasterPart>();
            var result = new Dictionary<string, uint>(StringComparer.Ordinal);
            var masterIdsByName = legacy.Masters.ToDictionary(master =>
                $"Binary {(master.IsMainMaster ? "Main" : "Title")} Master {master.MasterId:X8}",
                master => master.MasterId, StringComparer.Ordinal);
            foreach (SlideMasterPart masterPart in masterParts) {
                foreach (SlideLayoutPart layoutPart in masterPart.SlideLayoutParts) {
                    string? name = layoutPart.SlideLayout?.CommonSlideData?.Name?.Value;
                    if (name == null) continue;
                    foreach (KeyValuePair<string, uint> candidate in masterIdsByName) {
                        if (string.Equals(name, candidate.Key, StringComparison.Ordinal)
                            || name.StartsWith(candidate.Key + " / ", StringComparison.Ordinal)) {
                            result[layoutPart.Uri.ToString()] = candidate.Value;
                            break;
                        }
                    }
                }
            }
            return result;
        }
    }
}
