using System.Collections.ObjectModel;
using OfficeIMO.PowerPoint.LegacyPpt.Model;
using OfficeIMO.PowerPoint.LegacyPpt.Write;

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
            ArgumentNullException.ThrowIfNull(slide);
            return TryGetSlide(slide.PartRef.Name, out projection);
        }

        internal bool TryGetSlide(string slidePartName, out LegacyPptSlideProjection? projection) {
            projection = null;
            return slidePartName != null
                && _slidesByPartUri.TryGetValue(slidePartName, out projection);
        }

        internal bool TryGetSlide(uint legacySlideId,
            out LegacyPptSlideProjection? projection) =>
            _slidesByLegacyId.TryGetValue(legacySlideId, out projection);

        internal bool TryGetMasterId(PowerPointSlide slide, out uint masterId) {
            ArgumentNullException.ThrowIfNull(slide);
            masterId = 0;
            OpenXml.PowerPointPartRef layoutPart = slide.PartRef.RelatedPart(
                LegacyPptPackageView.SlideLayoutRelationshipType);
            return !layoutPart.IsNone
                && _masterIdsByLayoutPartUri.TryGetValue(layoutPart.Name, out masterId);
        }

        internal bool IsProjectedLayoutPart(string partUri) =>
            partUri != null && _masterIdsByLayoutPartUri.ContainsKey(partUri);

        internal bool IsEditableProjectedOrdinaryLayoutPart(string partUri) =>
            partUri != null
            && _ordinaryLayoutsByPartUri.ContainsKey(partUri);

        internal bool OrdinaryLayoutMatches(LegacyPptPartView layoutPart) =>
            !layoutPart.IsNone
            && _ordinaryLayoutsByPartUri.TryGetValue(
                layoutPart.Name, out LegacyPptOrdinaryLayoutProjection?
                    projection)
            && projection.ShapeTreeMatches(layoutPart);

        internal bool OrdinaryLayoutTypeMatches(LegacyPptPartView layoutPart) =>
            !layoutPart.IsNone
            && _ordinaryLayoutsByPartUri.TryGetValue(
                layoutPart.Name, out LegacyPptOrdinaryLayoutProjection?
                    projection)
            && projection.TypeMatches(layoutPart);

        internal bool TryGetOrdinaryLayoutAddedShapeIds(
            LegacyPptPartView layoutPart,
            out IReadOnlyList<uint> addedShapeIds) {
            addedShapeIds = Array.Empty<uint>();
            return !layoutPart.IsNone
                && _ordinaryLayoutsByPartUri.TryGetValue(
                    layoutPart.Name, out LegacyPptOrdinaryLayoutProjection?
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

        internal bool TryGetMaster(string masterPartName,
            out LegacyPptMasterProjection? projection) {
            projection = null;
            return masterPartName != null
                && _mastersByPartUri.TryGetValue(masterPartName, out projection);
        }

        internal bool IsProjectedMasterPart(string partUri) =>
            partUri != null && _mastersByPartUri.ContainsKey(partUri);

        internal bool IsProjectedMasterThemePart(string partUri) =>
            partUri != null && _masterThemePartUris.Contains(partUri);

        internal bool TryGetSpecialMaster(string masterPartName,
            out LegacyPptMasterProjection? projection) {
            projection = null;
            return masterPartName != null
                && _specialMastersByPartUri.TryGetValue(masterPartName, out projection);
        }

        internal bool TryGetTitleMaster(string layoutPartName,
            out LegacyPptMasterProjection? projection) {
            projection = null;
            return layoutPartName != null
                && _titleMastersByPartUri.TryGetValue(layoutPartName, out projection);
        }

        /// <summary>
        /// Builds the map from a freshly projected presentation. Its markup is read from a copy of its package; the
        /// writer's shape checks read engine shapes, which come from a separate reader of that copy while the
        /// presentation is still held in its package, so building the map leaves it there.
        /// </summary>
        internal static LegacyPptProjectionMap Create(PowerPointDocument presentation,
            LegacyPptDocument legacy) {
            ArgumentNullException.ThrowIfNull(presentation);
            ArgumentNullException.ThrowIfNull(legacy);
            byte[] packageBytes = presentation.CapturePackageBytes();
            using LegacyPptPackageView view = LegacyPptPackageView.Open(packageBytes);
            if (presentation.IsSdkLoaded) return Create(presentation, view, legacy);
            using PowerPointDocument reader = PowerPointDocument.OpenLegacyPptProjectionReader(packageBytes);
            return Create(reader, view, legacy);
        }

        private static LegacyPptProjectionMap Create(PowerPointDocument shapes,
            LegacyPptPackageView view, LegacyPptDocument legacy) {
            if (shapes.Slides.Count != legacy.Slides.Count) {
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
                PowerPointSlide projectedSlide = shapes.Slides[slideIndex];
                LegacyPptPartView slidePart = view.Part(projectedSlide.PartRef.Name);
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

                var shapeProjections = new List<LegacyPptShapeProjection>();
                AddShapeProjectionTree(shapeProjections, sourceShapes,
                    projectedShapes, projectableSoundIds,
                    $"slide {slideIndex + 1}", view);
                bool hasExplicitBackground = !slidePart.Root
                    .Child(LegacyPptXml.P, "cSld").Child(LegacyPptXml.P, "bg").IsNone;
                LegacyPptPartView layoutPart = slidePart.SlideLayout;
                LegacyPptPartView themeOverride = slidePart.ThemeOverride;
                slides.Add(new LegacyPptSlideProjection(slidePart.Name,
                    layoutPart.IsNone ? null : layoutPart.Name,
                    sourceSlide.PersistId, sourceSlide.SlideId, sourceSlide.MasterId,
                    sourceSlide.LayoutType,
                    sourceSlide.LayoutPlaceholderTypes.Select(value =>
                        unchecked((byte)value)).ToArray(),
                    sourceSlide.Hidden, sourceSlide.FollowsMasterObjects,
                    sourceSlide.HeaderFooter,
                    hasExplicitBackground,
                    CreateSlideBackgroundFingerprint(slidePart,
                        hasExplicitBackground, inheritedBackgrounds),
                    themeOverride.IsNone ? null : themeOverride.Name,
                    CreateSlideThemeFingerprint(slidePart,
                        inheritedThemes),
                    LegacyPptSlideProjection.CreateClassicColorFingerprints(
                        slidePart),
                    sourceSlide.Transition, sourceSlide.Comments, shapeProjections,
                    sourceSlide.NotesPage == null
                        ? null
                        : CreateNotesProjection(slidePart,
                            sourceSlide.NotesPage)));
            }
            IReadOnlyList<LegacyPptMasterProjection> titleMasters =
                CreateTitleMasterProjections(shapes, view, legacy);
            return new LegacyPptProjectionMap(slides,
                CreateLayoutMasterMap(view, legacy),
                LegacyPptOrdinaryLayoutProjection.Create(view,
                    titleMasters),
                CreateMasterProjections(shapes, view, legacy),
                titleMasters,
                CreateSpecialMasterProjections(shapes, view, legacy),
                legacy.Hyperlinks, legacy.CustomShows,
                legacy.CustomShowsAreEditable, legacy.Sounds,
                legacy.SoundIdSeed,
                LegacyPptPropertySetCodec.CreateProjection(view,
                    legacy.Package),
                LegacyPptVbaProjectProjection.Create(legacy.VbaProject));
        }

        private static string CreateSlideBackgroundFingerprint(
            LegacyPptPartView slidePart, bool hasExplicitBackground,
            IDictionary<string, string> inheritedFingerprints) {
            LegacyPptPartView layoutPart = slidePart.SlideLayout;
            if (hasExplicitBackground || layoutPart.IsNone) {
                return LegacyPptSlideProjection.CreateBackgroundFingerprint(
                    slidePart);
            }
            string key = layoutPart.Name;
            if (!inheritedFingerprints.TryGetValue(key,
                    out string? fingerprint)) {
                fingerprint = LegacyPptSlideProjection
                    .CreateBackgroundFingerprint(slidePart);
                inheritedFingerprints.Add(key, fingerprint);
            }
            return fingerprint;
        }

        private static string CreateSlideThemeFingerprint(
            LegacyPptPartView slidePart,
            IDictionary<string, string> inheritedFingerprints) {
            LegacyPptPartView layoutPart = slidePart.SlideLayout;
            if (!slidePart.ThemeOverride.IsNone || layoutPart.IsNone) {
                return LegacyPptSlideProjection.CreateThemeFingerprint(slidePart);
            }
            string key = layoutPart.Name;
            if (!inheritedFingerprints.TryGetValue(key,
                    out string? fingerprint)) {
                fingerprint = LegacyPptSlideProjection
                    .CreateThemeFingerprint(slidePart);
                inheritedFingerprints.Add(key, fingerprint);
            }
            return fingerprint;
        }

        private static void AddShapeProjectionTree(
            ICollection<LegacyPptShapeProjection> result,
            IReadOnlyList<LegacyPptShape> sourceShapes,
            IReadOnlyList<PowerPointShape> projectedShapes,
            ISet<uint> projectableSoundIds, string ownerName,
            LegacyPptPackageView view) {
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
                LegacyPptProjectedShape projected = LegacyPptProjectedShape.Resolve(view, projectedShape);
                string? textFormattingFingerprint =
                    (sourceShape.TextBody.HasStyleRecord
                        || sourceShape.TextBody.HasStyle9Record
                        || sourceShape.TextBody.HasFieldRecords
                        || sourceShape.TextBody.HasInteractions
                        || sourceShape.TextBody.HasTextSpecialInfoRecord)
                    && projected.Element.Is(LegacyPptXml.P, "sp")
                        ? LegacyPptShapeProjection.CreateTextFormattingFingerprint(projected)
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
                            .CreateTextFrameFingerprint(projected)
                        : null,
                    sourceShape.Kind == LegacyPptShapeKind.TextBox
                        && sourceShape.TextFrame
                            .CanRewriteProjectedProperties,
                    LegacyPptShapeProjection
                        .CreateShapeTransformFingerprint(projected),
                    sourceShape.OfficeArtShapeType is 2 or 23
                        ? LegacyPptShapeProjection
                            .CreateShapeGeometryFingerprint(projected)
                        : null,
                    sourceShape.Kind == LegacyPptShapeKind.Group
                        ? LegacyPptShapeProjection
                            .CreateGroupCoordinateFingerprint(projected)
                        : null,
                    sourceShape.Style.CanRewriteProjectedVisualStyle
                        ? LegacyPptShapeProjection
                            .CreateShapeVisualStyleFingerprint(projected)
                        : null,
                    LegacyPptShapeProjection
                        .CreatePictureFormattingFingerprint(projected),
                    sourceShape.Style.CanRewriteHiddenState
                        ? LegacyPptShapeProjection
                            .CreateShapeVisibilityFingerprint(projectedShape)
                        : null,
                    LegacyPptShapeProjection
                        .CreateShapeMetadataFingerprint(projectedShape),
                    sourceShape.Metadata.CanRewrite,
                    sourceShape.OleObject != null
                        && projectedShape is PowerPointOleObject
                        ? LegacyPptOleObjectProjection.Create(
                            sourceShape.OleObject, projected.Part, projected.Element)
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
                    .Where(PowerPointDocument
                        .CanCreateLegacyOpenXmlShape)
                    .ToArray();
                AddShapeProjectionTree(result, sourceChildren,
                    projectedChildren, projectableSoundIds,
                    $"{ownerName}, group shape {index + 1}", view);
            }
        }

        private static LegacyPptNotesProjection CreateNotesProjection(
            LegacyPptPartView slidePart, LegacyPptNotesPage source) {
            LegacyPptPartView part = slidePart.NotesSlide;
            if (part.IsNone) {
                throw new InvalidDataException(
                    "The projected binary notes page has no notes-slide part.");
            }
            LegacyPptPartView themeOverride = part.ThemeOverride;
            return new LegacyPptNotesProjection(source.PersistId,
                source.NotesId, source.Text, source.FollowsMasterObjects,
                themeOverride.IsNone ? null : themeOverride.Name,
                LegacyPptNotesProjection.CreateThemeFingerprint(part),
                LegacyPptNotesProjection.CreateClassicColorFingerprints(part),
                LegacyPptNotesProjection.CreateBackgroundFingerprint(part));
        }

        private static IReadOnlyList<LegacyPptMasterProjection> CreateMasterProjections(
            PowerPointDocument shapes, LegacyPptPackageView view, LegacyPptDocument legacy) {
            IReadOnlyList<LegacyPptPartView> masterParts = LegacyPptPackageFingerprint.SlideMasterParts(view);
            LegacyPptMaster[] sourceMasters = legacy.Masters
                .Where(master => master.IsMainMaster).ToArray();
            if (masterParts.Count != sourceMasters.Length) {
                throw new InvalidDataException(
                    "The projected slide-master count does not match the binary source master count.");
            }

            var result = new List<LegacyPptMasterProjection>(masterParts.Count);
            for (int index = 0; index < masterParts.Count; index++) {
                LegacyPptPartView masterPart = masterParts[index];
                LegacyPptPartView themePart = masterPart.Theme;
                if (themePart.IsNone) {
                    throw new InvalidDataException(
                        $"Projected slide master {index + 1} has no theme part.");
                }
                IReadOnlyList<LegacyPptShapeProjection> shapeProjections =
                    CreateMasterShapeProjections(sourceMasters[index].Shapes,
                        LegacyPptPreservingWriter.ReadMasterShapesForProjection(
                            shapes, masterPart.Name), $"slide master {index + 1}",
                        view);
                result.Add(new LegacyPptMasterProjection(
                    masterPart.Name, themePart.Name,
                    sourceMasters[index].PersistId,
                    LegacyPptMasterProjection.CreateThemeFingerprint(masterPart),
                    LegacyPptMasterProjection.CreateClassicColorFingerprints(masterPart),
                    LegacyPptMasterProjection.CreateBackgroundFingerprint(masterPart),
                    followsMasterObjects: null,
                    shapeProjections,
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

        private static string? LayoutName(LegacyPptPartView layout) =>
            layout.Root.Child(LegacyPptXml.P, "cSld").Attribute("name");

        private static IReadOnlyList<LegacyPptMasterProjection>
            CreateTitleMasterProjections(PowerPointDocument shapes,
                LegacyPptPackageView view, LegacyPptDocument legacy) {
            LegacyPptPartView[] layouts = LegacyPptPackageFingerprint.SlideMasterParts(view)
                .SelectMany(LegacyPptPackageFingerprint.SlideLayoutParts).ToArray();
            var result = new List<LegacyPptMasterProjection>();
            foreach (LegacyPptMaster source in legacy.Masters.Where(master =>
                         !master.IsMainMaster)) {
                string expectedName =
                    $"Binary Title Master {source.MasterId:X8}";
                LegacyPptPartView part = layouts.SingleOrDefault(layout =>
                        string.Equals(LayoutName(layout),
                            expectedName, StringComparison.Ordinal));
                if (part.IsNone) {
                    throw new InvalidDataException(
                        $"The projected title master 0x{source.MasterId:X8} has no exact layout part.");
                }
                LegacyPptPartView themeOverride = part.ThemeOverride;
                result.Add(new LegacyPptMasterProjection(
                    part.Name, themeOverride.IsNone ? null : themeOverride.Name,
                    source.PersistId,
                    LegacyPptMasterProjection.CreateThemeFingerprint(part),
                    LegacyPptMasterProjection.CreateClassicColorFingerprints(part),
                    LegacyPptMasterProjection.CreateBackgroundFingerprint(part),
                    source.FollowsMasterObjects,
                    CreateMasterShapeProjections(source.Shapes,
                        LegacyPptPreservingWriter.ReadMasterShapesForProjection(
                            shapes, part.Name),
                        $"title master 0x{source.MasterId:X8}", view)));
            }
            return result;
        }

        private static IReadOnlyList<LegacyPptMasterProjection>
            CreateSpecialMasterProjections(PowerPointDocument shapes,
                LegacyPptPackageView view, LegacyPptDocument legacy) {
            var result = new List<LegacyPptMasterProjection>(2);
            if (legacy.NotesMaster != null) {
                result.Add(CreateSpecialMasterProjection(shapes, view,
                    view.NotesMaster, "notes master", legacy.NotesMaster.PersistId,
                    legacy.NotesMaster.Shapes));
            }
            if (legacy.HandoutMaster != null) {
                result.Add(CreateSpecialMasterProjection(shapes, view,
                    view.HandoutMaster, "handout master", legacy.HandoutMaster.PersistId,
                    legacy.HandoutMaster.Shapes));
            }
            return result;
        }

        private static LegacyPptMasterProjection CreateSpecialMasterProjection(
            PowerPointDocument shapes, LegacyPptPackageView view,
            LegacyPptPartView part, string ownerName, uint persistId,
            IReadOnlyList<LegacyPptShape> sourceShapes) {
            if (part.IsNone) {
                throw new InvalidDataException(
                    $"The projected presentation has no {ownerName.Replace(" ", "-")} part.");
            }
            LegacyPptPartView themePart = part.Theme;
            if (themePart.IsNone) {
                throw new InvalidDataException(
                    $"The projected {ownerName} has no theme part.");
            }
            return new LegacyPptMasterProjection(
                part.Name, themePart.Name,
                persistId,
                LegacyPptMasterProjection.CreateThemeFingerprint(part),
                LegacyPptMasterProjection.CreateClassicColorFingerprints(part),
                LegacyPptMasterProjection.CreateBackgroundFingerprint(part),
                followsMasterObjects: null,
                CreateMasterShapeProjections(sourceShapes,
                    LegacyPptPreservingWriter.ReadMasterShapesForProjection(
                        shapes, part.Name),
                    ownerName, view));
        }

        private static IReadOnlyList<LegacyPptShapeProjection>
            CreateMasterShapeProjections(IReadOnlyList<LegacyPptShape> source,
                IReadOnlyList<PowerPointShape> projected, string ownerName,
                LegacyPptPackageView view) {
            LegacyPptShape[] sourceShapes = source.Where(shape =>
                PowerPointDocument.CanCreateLegacyOpenXmlShape(shape))
                .ToArray();
            var result = new List<LegacyPptShapeProjection>();
            AddShapeProjectionTree(result, sourceShapes, projected,
                new HashSet<uint>(), ownerName, view);
            return result;
        }

        private static IReadOnlyDictionary<string, uint> CreateLayoutMasterMap(
            LegacyPptPackageView view, LegacyPptDocument legacy) {
            var result = new Dictionary<string, uint>(StringComparer.Ordinal);
            var masterIdsByName = legacy.Masters.ToDictionary(master =>
                $"Binary {(master.IsMainMaster ? "Main" : "Title")} Master {master.MasterId:X8}",
                master => master.MasterId, StringComparer.Ordinal);
            foreach (LegacyPptPartView masterPart in LegacyPptPackageFingerprint.SlideMasterParts(view)) {
                foreach (LegacyPptPartView layoutPart in LegacyPptPackageFingerprint.SlideLayoutParts(masterPart)) {
                    string? name = LayoutName(layoutPart);
                    if (name == null) continue;
                    foreach (KeyValuePair<string, uint> candidate in masterIdsByName) {
                        if (string.Equals(name, candidate.Key, StringComparison.Ordinal)
                            || name.StartsWith(candidate.Key + " / ", StringComparison.Ordinal)) {
                            result[layoutPart.Name] = candidate.Value;
                            break;
                        }
                    }
                }
            }
            return result;
        }
    }
}
