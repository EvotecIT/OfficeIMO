using System.Collections.ObjectModel;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using OfficeIMO.PowerPoint.LegacyPpt.Model;
using OfficeIMO.PowerPoint.LegacyPpt.Write;
using A = DocumentFormat.OpenXml.Drawing;

namespace OfficeIMO.PowerPoint.LegacyPpt.Internal {
    /// <summary>Links projected Open XML slides and shapes back to their original binary persist records.</summary>
    internal sealed class LegacyPptProjectionMap {
        private readonly IReadOnlyDictionary<string, LegacyPptSlideProjection> _slidesByPartUri;
        private readonly IReadOnlyDictionary<uint, LegacyPptSlideProjection> _slidesByLegacyId;
        private readonly IReadOnlyDictionary<string, uint> _masterIdsByLayoutPartUri;
        private readonly IReadOnlyDictionary<string, LegacyPptMasterProjection> _mastersByPartUri;
        private readonly IReadOnlyDictionary<string, LegacyPptMasterProjection>
            _titleMastersByPartUri;
        private readonly IReadOnlyDictionary<string, LegacyPptMasterProjection>
            _specialMastersByPartUri;
        private readonly ISet<string> _masterThemePartUris;

        private LegacyPptProjectionMap(IReadOnlyList<LegacyPptSlideProjection> slides,
            IReadOnlyDictionary<string, uint> masterIdsByLayoutPartUri,
            IReadOnlyList<LegacyPptMasterProjection> masters,
            IReadOnlyList<LegacyPptMasterProjection> titleMasters,
            IReadOnlyList<LegacyPptMasterProjection> specialMasters,
            IReadOnlyList<LegacyPptHyperlink> hyperlinks,
            IReadOnlyList<LegacyPptCustomShow> customShows,
            bool customShowsAreEditable,
            IReadOnlyList<LegacyPptSound> sounds, uint? soundIdSeed) {
            Slides = new ReadOnlyCollection<LegacyPptSlideProjection>(slides.ToArray());
            _slidesByPartUri = new ReadOnlyDictionary<string, LegacyPptSlideProjection>(slides.ToDictionary(
                slide => slide.SlidePartUri, StringComparer.Ordinal));
            _slidesByLegacyId = new ReadOnlyDictionary<uint, LegacyPptSlideProjection>(slides.ToDictionary(
                slide => slide.SlideId));
            _masterIdsByLayoutPartUri = new ReadOnlyDictionary<string, uint>(
                masterIdsByLayoutPartUri.ToDictionary(pair => pair.Key, pair => pair.Value,
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
            Hyperlinks = new ReadOnlyCollection<LegacyPptHyperlink>(hyperlinks.ToArray());
            CustomShows = new ReadOnlyCollection<LegacyPptCustomShow>(
                customShows.ToArray());
            CanEditCustomShows = customShowsAreEditable
                && customShows.All(show => show.IsEditable)
                && customShows.Select(show => show.Name)
                    .Distinct(StringComparer.Ordinal).Count() == customShows.Count;
            Sounds = new ReadOnlyCollection<LegacyPptSound>(sounds.ToArray());
            SoundIdSeed = soundIdSeed;
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

                var shapes = new List<LegacyPptShapeProjection>(projectedShapes.Length);
                for (int shapeIndex = 0; shapeIndex < projectedShapes.Length; shapeIndex++) {
                    uint? openXmlShapeId = projectedShapes[shapeIndex].Id;
                    if (!openXmlShapeId.HasValue) {
                        throw new InvalidDataException(
                            $"Projected slide {slideIndex + 1}, shape {shapeIndex + 1} has no Open XML shape id.");
                    }
                    LegacyPptShape sourceShape = sourceShapes[shapeIndex];
                    string? textFormattingFingerprint = (sourceShape.TextBody.HasStyleRecord
                        || sourceShape.TextBody.HasInteractions)
                        && projectedShapes[shapeIndex].Element is DocumentFormat.OpenXml.Presentation.Shape projectedTextShape
                        ? LegacyPptTextProjection.CreateFormattingFingerprint(projectedTextShape.TextBody)
                        : null;
                    shapes.Add(new LegacyPptShapeProjection(openXmlShapeId.Value, sourceShape.ShapeId,
                        sourceShape.RecordOffset, sourceShape.Kind, sourceShape.Bounds, sourceShape.Text,
                        textFormattingFingerprint, sourceShape.Interactions,
                        sourceShape.TextBody.Interactions, sourceShape.Animation,
                        projectableSoundIds));
                }
                slides.Add(new LegacyPptSlideProjection(projectedSlide.SlidePart.Uri.ToString(),
                    projectedSlide.SlidePart.SlideLayoutPart?.Uri.ToString(),
                    sourceSlide.PersistId, sourceSlide.SlideId, sourceSlide.MasterId,
                    sourceSlide.Hidden, sourceSlide.FollowsMasterObjects,
                    sourceSlide.HeaderFooter,
                    projectedSlide.SlidePart.Slide?.CommonSlideData?
                        .Background != null,
                    LegacyPptSlideProjection.CreateBackgroundFingerprint(
                        projectedSlide),
                    projectedSlide.SlidePart.ThemeOverridePart?.Uri.ToString(),
                    LegacyPptSlideProjection.CreateThemeFingerprint(
                        projectedSlide),
                    LegacyPptSlideProjection.CreateClassicColorFingerprints(
                        projectedSlide),
                    sourceSlide.Transition, sourceSlide.Comments, shapes,
                    sourceSlide.NotesPage == null
                        ? null
                        : CreateNotesProjection(projectedSlide,
                            sourceSlide.NotesPage)));
            }
            return new LegacyPptProjectionMap(slides, CreateLayoutMasterMap(presentation, legacy),
                CreateMasterProjections(presentation, legacy),
                CreateTitleMasterProjections(presentation, legacy),
                CreateSpecialMasterProjections(presentation, legacy),
                legacy.Hyperlinks, legacy.CustomShows,
                legacy.CustomShowsAreEditable, legacy.Sounds,
                legacy.SoundIdSeed);
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
                            out _), $"slide master {index + 1}");
                result.Add(new LegacyPptMasterProjection(
                    masterPart.Uri.ToString(), themePart.Uri.ToString(),
                    sourceMasters[index].PersistId,
                    LegacyPptMasterProjection.CreateThemeFingerprint(masterPart),
                    LegacyPptMasterProjection.CreateClassicColorFingerprints(masterPart),
                    LegacyPptMasterProjection.CreateBackgroundFingerprint(masterPart),
                    followsMasterObjects: null,
                    shapes));
            }
            return result;
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
                        $"title master 0x{source.MasterId:X8}")));
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
                        "notes master")));
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
                        "handout master")));
            }
            return result;
        }

        private static IReadOnlyList<LegacyPptShapeProjection>
            CreateMasterShapeProjections(IReadOnlyList<LegacyPptShape> source,
                IReadOnlyList<PowerPointShape> projected, string ownerName) {
            LegacyPptShape[] sourceShapes = source.Where(shape =>
                shape.Kind != LegacyPptShapeKind.Unsupported).ToArray();
            if (projected.Count != sourceShapes.Length) {
                throw new InvalidDataException(
                    $"Projected {ownerName} has {projected.Count} editable shapes, "
                    + $"but the binary source exposed {sourceShapes.Length}.");
            }
            var result = new List<LegacyPptShapeProjection>(projected.Count);
            for (int index = 0; index < projected.Count; index++) {
                PowerPointShape projectedShape = projected[index];
                uint? openXmlShapeId = projectedShape.Id;
                if (!openXmlShapeId.HasValue) {
                    throw new InvalidDataException(
                        $"Projected {ownerName}, shape {index + 1} has no Open XML shape id.");
                }
                LegacyPptShape sourceShape = sourceShapes[index];
                string? textFormattingFingerprint = (sourceShape.TextBody.HasStyleRecord
                        || sourceShape.TextBody.HasInteractions)
                    && projectedShape.Element is DocumentFormat.OpenXml.Presentation.Shape projectedTextShape
                    ? LegacyPptTextProjection.CreateFormattingFingerprint(
                        projectedTextShape.TextBody)
                    : null;
                result.Add(new LegacyPptShapeProjection(openXmlShapeId.Value,
                    sourceShape.ShapeId, sourceShape.RecordOffset,
                    sourceShape.Kind, sourceShape.Bounds, sourceShape.Text,
                    textFormattingFingerprint, sourceShape.Interactions,
                    sourceShape.TextBody.Interactions, sourceShape.Animation,
                    new HashSet<uint>()));
            }
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

    /// <summary>Maps one projected slide master and its editable theme back to a binary persist object.</summary>
    internal sealed class LegacyPptMasterProjection {
        private readonly IReadOnlyDictionary<uint, LegacyPptShapeProjection>
            _shapesByOpenXmlId;

        internal LegacyPptMasterProjection(string masterPartUri, string? themePartUri,
            uint persistId, string themeFingerprint,
            IReadOnlyList<string> classicColorFingerprints,
            string backgroundFingerprint, bool? followsMasterObjects,
            IReadOnlyList<LegacyPptShapeProjection> shapes) {
            MasterPartUri = masterPartUri ?? throw new ArgumentNullException(nameof(masterPartUri));
            ThemePartUri = themePartUri;
            PersistId = persistId;
            ThemeFingerprint = themeFingerprint ?? throw new ArgumentNullException(nameof(themeFingerprint));
            BackgroundFingerprint = backgroundFingerprint
                ?? throw new ArgumentNullException(nameof(backgroundFingerprint));
            FollowsMasterObjects = followsMasterObjects;
            ClassicColorFingerprints = new ReadOnlyCollection<string>(
                (classicColorFingerprints
                    ?? throw new ArgumentNullException(nameof(classicColorFingerprints)))
                .ToArray());
            if (ClassicColorFingerprints.Count != 8) {
                throw new ArgumentException(
                    "A projected classic color scheme requires eight fingerprints.",
                    nameof(classicColorFingerprints));
            }
            Shapes = new ReadOnlyCollection<LegacyPptShapeProjection>(
                (shapes ?? throw new ArgumentNullException(nameof(shapes))).ToArray());
            _shapesByOpenXmlId = new ReadOnlyDictionary<uint, LegacyPptShapeProjection>(
                shapes.ToDictionary(shape => shape.OpenXmlShapeId));
        }

        internal string MasterPartUri { get; }

        internal string? ThemePartUri { get; }

        internal uint PersistId { get; }

        internal string ThemeFingerprint { get; }

        internal IReadOnlyList<string> ClassicColorFingerprints { get; }

        internal string BackgroundFingerprint { get; }

        internal bool? FollowsMasterObjects { get; }

        internal IReadOnlyList<LegacyPptShapeProjection> Shapes { get; }

        internal bool TryGetShape(uint openXmlShapeId,
            out LegacyPptShapeProjection? projection) =>
            _shapesByOpenXmlId.TryGetValue(openXmlShapeId, out projection);

        internal bool ThemeMatches(SlideMasterPart masterPart) => string.Equals(
            ThemeFingerprint, CreateThemeFingerprint(masterPart), StringComparison.Ordinal);

        internal bool ThemeMatches(NotesMasterPart masterPart) => string.Equals(
            ThemeFingerprint, CreateThemeFingerprint(masterPart), StringComparison.Ordinal);

        internal bool ThemeMatches(HandoutMasterPart masterPart) => string.Equals(
            ThemeFingerprint, CreateThemeFingerprint(masterPart), StringComparison.Ordinal);

        internal bool ThemeMatches(SlideLayoutPart masterPart) => string.Equals(
            ThemeFingerprint, CreateThemeFingerprint(masterPart), StringComparison.Ordinal);

        internal bool BackgroundMatches(SlideMasterPart masterPart) => string.Equals(
            BackgroundFingerprint, CreateBackgroundFingerprint(masterPart),
            StringComparison.Ordinal);

        internal bool BackgroundMatches(NotesMasterPart masterPart) => string.Equals(
            BackgroundFingerprint, CreateBackgroundFingerprint(masterPart),
            StringComparison.Ordinal);

        internal bool BackgroundMatches(HandoutMasterPart masterPart) => string.Equals(
            BackgroundFingerprint, CreateBackgroundFingerprint(masterPart),
            StringComparison.Ordinal);

        internal bool BackgroundMatches(SlideLayoutPart masterPart) => string.Equals(
            BackgroundFingerprint, CreateBackgroundFingerprint(masterPart),
            StringComparison.Ordinal);

        internal bool MasterObjectsMatch(SlideLayoutPart masterPart) =>
            !FollowsMasterObjects.HasValue
            || FollowsMasterObjects.Value
                == (masterPart.SlideLayout?.ShowMasterShapes?.Value != false);

        internal IReadOnlyList<int> GetChangedClassicColorSlots(
            SlideMasterPart masterPart) {
            IReadOnlyList<string> current = CreateClassicColorFingerprints(masterPart);
            return GetChangedClassicColorSlots(current);
        }

        internal IReadOnlyList<int> GetChangedClassicColorSlots(
            NotesMasterPart masterPart) {
            IReadOnlyList<string> current = CreateClassicColorFingerprints(masterPart);
            return GetChangedClassicColorSlots(current);
        }

        internal IReadOnlyList<int> GetChangedClassicColorSlots(
            HandoutMasterPart masterPart) {
            IReadOnlyList<string> current = CreateClassicColorFingerprints(masterPart);
            return GetChangedClassicColorSlots(current);
        }

        internal IReadOnlyList<int> GetChangedClassicColorSlots(
            SlideLayoutPart masterPart) {
            IReadOnlyList<string> current = CreateClassicColorFingerprints(masterPart);
            return GetChangedClassicColorSlots(current);
        }

        private IReadOnlyList<int> GetChangedClassicColorSlots(
            IReadOnlyList<string> current) {
            return Enumerable.Range(0, ClassicColorFingerprints.Count)
                .Where(index => !string.Equals(ClassicColorFingerprints[index],
                    current[index], StringComparison.Ordinal))
                .ToArray();
        }

        internal static string CreateThemeFingerprint(SlideMasterPart masterPart) {
            if (masterPart == null) throw new ArgumentNullException(nameof(masterPart));
            return CreateThemeFingerprint(masterPart.ThemePart,
                masterPart.SlideMaster?.ColorMap);
        }

        internal static string CreateThemeFingerprint(NotesMasterPart masterPart) {
            if (masterPart == null) throw new ArgumentNullException(nameof(masterPart));
            return CreateThemeFingerprint(masterPart.ThemePart,
                masterPart.NotesMaster?.ColorMap);
        }

        internal static string CreateThemeFingerprint(HandoutMasterPart masterPart) {
            if (masterPart == null) throw new ArgumentNullException(nameof(masterPart));
            return CreateThemeFingerprint(masterPart.ThemePart,
                masterPart.HandoutMaster?.ColorMap);
        }

        internal static string CreateThemeFingerprint(SlideLayoutPart masterPart) {
            if (masterPart == null) throw new ArgumentNullException(nameof(masterPart));
            string theme = masterPart.ThemeOverridePart?.ThemeOverride?.OuterXml
                ?? string.Empty;
            string colorMap = masterPart.SlideLayout?.ColorMapOverride?.OuterXml
                ?? string.Empty;
            return theme + "\n" + colorMap;
        }

        private static string CreateThemeFingerprint(ThemePart? themePart,
            OpenXmlElement? colorMap) {
            string theme = themePart?.Theme?.OuterXml ?? string.Empty;
            return theme + "\n" + (colorMap?.OuterXml ?? string.Empty);
        }

        internal static IReadOnlyList<string> CreateClassicColorFingerprints(
            SlideMasterPart masterPart) {
            if (masterPart == null) throw new ArgumentNullException(nameof(masterPart));
            return CreateClassicColorFingerprints(masterPart.ThemePart);
        }

        internal static IReadOnlyList<string> CreateClassicColorFingerprints(
            NotesMasterPart masterPart) {
            if (masterPart == null) throw new ArgumentNullException(nameof(masterPart));
            return CreateClassicColorFingerprints(masterPart.ThemePart);
        }

        internal static IReadOnlyList<string> CreateClassicColorFingerprints(
            HandoutMasterPart masterPart) {
            if (masterPart == null) throw new ArgumentNullException(nameof(masterPart));
            return CreateClassicColorFingerprints(masterPart.ThemePart);
        }

        internal static IReadOnlyList<string> CreateClassicColorFingerprints(
            SlideLayoutPart masterPart) {
            if (masterPart == null) throw new ArgumentNullException(nameof(masterPart));
            DocumentFormat.OpenXml.Drawing.ColorScheme? colors = masterPart
                .ThemeOverridePart?.ThemeOverride?.ColorScheme
                ?? masterPart.SlideMasterPart?.ThemePart?.Theme?
                    .ThemeElements?.ColorScheme;
            return CreateClassicColorFingerprints(colors);
        }

        private static IReadOnlyList<string> CreateClassicColorFingerprints(
            ThemePart? themePart) {
            DocumentFormat.OpenXml.Drawing.ColorScheme? colors = themePart?
                .Theme?.ThemeElements?.ColorScheme;
            return CreateClassicColorFingerprints(colors);
        }

        private static IReadOnlyList<string> CreateClassicColorFingerprints(
            DocumentFormat.OpenXml.Drawing.ColorScheme? colors) {
            OpenXmlElement?[] slots = {
                colors?.Light1Color,
                colors?.Dark1Color,
                colors?.Accent4Color,
                colors?.Dark2Color,
                colors?.Light2Color,
                colors?.Accent1Color,
                colors?.Accent2Color,
                colors?.Accent3Color
            };
            return slots.Select(slot => slot?.OuterXml ?? string.Empty).ToArray();
        }

        internal static string CreateBackgroundFingerprint(
            SlideMasterPart masterPart) {
            if (masterPart == null) throw new ArgumentNullException(nameof(masterPart));
            return masterPart.SlideMaster?.CommonSlideData?.Background?.OuterXml
                ?? string.Empty;
        }

        internal static string CreateBackgroundFingerprint(
            NotesMasterPart masterPart) {
            if (masterPart == null) throw new ArgumentNullException(nameof(masterPart));
            return masterPart.NotesMaster?.CommonSlideData?.Background?.OuterXml
                ?? string.Empty;
        }

        internal static string CreateBackgroundFingerprint(
            HandoutMasterPart masterPart) {
            if (masterPart == null) throw new ArgumentNullException(nameof(masterPart));
            return masterPart.HandoutMaster?.CommonSlideData?.Background?.OuterXml
                ?? string.Empty;
        }

        internal static string CreateBackgroundFingerprint(
            SlideLayoutPart masterPart) {
            if (masterPart == null) throw new ArgumentNullException(nameof(masterPart));
            return masterPart.SlideLayout?.CommonSlideData?.Background?.OuterXml
                ?? string.Empty;
        }
    }

    /// <summary>Maps one projected slide to its binary persist object.</summary>
    internal sealed class LegacyPptSlideProjection {
        private readonly IReadOnlyDictionary<uint, LegacyPptShapeProjection> _shapesByOpenXmlId;

        internal LegacyPptSlideProjection(string slidePartUri,
            string? layoutPartUri, uint persistId, uint slideId, uint masterId,
            bool hidden, bool followsMasterObjects,
            LegacyPptHeaderFooterSettings? headerFooter,
            bool hasExplicitBackground,
            string backgroundFingerprint, string? themePartUri,
            string themeFingerprint,
            IReadOnlyList<string> classicColorFingerprints,
            LegacyPptTransition? transition,
            IReadOnlyList<LegacyPptComment> comments,
            IReadOnlyList<LegacyPptShapeProjection> shapes, LegacyPptNotesProjection? notes) {
            SlidePartUri = slidePartUri ?? throw new ArgumentNullException(nameof(slidePartUri));
            LayoutPartUri = layoutPartUri;
            PersistId = persistId;
            SlideId = slideId;
            MasterId = masterId;
            Hidden = hidden;
            FollowsMasterObjects = followsMasterObjects;
            HeaderFooter = headerFooter;
            HasExplicitBackground = hasExplicitBackground;
            BackgroundFingerprint = backgroundFingerprint
                ?? throw new ArgumentNullException(nameof(backgroundFingerprint));
            ThemePartUri = themePartUri;
            ThemeFingerprint = themeFingerprint
                ?? throw new ArgumentNullException(nameof(themeFingerprint));
            ClassicColorFingerprints = new ReadOnlyCollection<string>(
                (classicColorFingerprints
                    ?? throw new ArgumentNullException(
                        nameof(classicColorFingerprints))).ToArray());
            if (ClassicColorFingerprints.Count != 8) {
                throw new ArgumentException(
                    "A projected classic color scheme requires eight fingerprints.",
                    nameof(classicColorFingerprints));
            }
            Transition = transition;
            Comments = new ReadOnlyCollection<LegacyPptComment>(comments.ToArray());
            Notes = notes;
            Shapes = new ReadOnlyCollection<LegacyPptShapeProjection>(shapes.ToArray());
            _shapesByOpenXmlId = new ReadOnlyDictionary<uint, LegacyPptShapeProjection>(shapes.ToDictionary(
                shape => shape.OpenXmlShapeId));
        }

        internal string SlidePartUri { get; }

        internal string? LayoutPartUri { get; }

        internal uint PersistId { get; }

        internal uint SlideId { get; }

        internal uint MasterId { get; }

        internal bool Hidden { get; }

        internal bool FollowsMasterObjects { get; }

        internal LegacyPptHeaderFooterSettings? HeaderFooter { get; }

        internal bool HasExplicitBackground { get; }

        internal string BackgroundFingerprint { get; }

        internal string? ThemePartUri { get; }

        internal string ThemeFingerprint { get; }

        internal IReadOnlyList<string> ClassicColorFingerprints { get; }

        internal bool BackgroundMatches(PowerPointSlide slide) => string.Equals(
            BackgroundFingerprint, CreateBackgroundFingerprint(slide),
            StringComparison.Ordinal);

        internal bool MasterObjectsMatch(PowerPointSlide slide) =>
            FollowsMasterObjects
                == (slide.SlidePart.Slide?.ShowMasterShapes?.Value != false);

        internal static string CreateBackgroundFingerprint(
            PowerPointSlide slide) {
            if (slide == null) throw new ArgumentNullException(nameof(slide));
            return slide.SlidePart.Slide?.CommonSlideData?.Background?.OuterXml
                ?? slide.SlidePart.SlideLayoutPart?.SlideLayout?
                    .CommonSlideData?.Background?.OuterXml
                ?? string.Empty;
        }

        internal bool ThemeMatches(PowerPointSlide slide) => string.Equals(
            ThemeFingerprint, CreateThemeFingerprint(slide),
            StringComparison.Ordinal);

        internal IReadOnlyList<int> GetChangedClassicColorSlots(
            PowerPointSlide slide) {
            IReadOnlyList<string> current =
                CreateClassicColorFingerprints(slide);
            return Enumerable.Range(0, ClassicColorFingerprints.Count)
                .Where(index => !string.Equals(
                    ClassicColorFingerprints[index], current[index],
                    StringComparison.Ordinal))
                .ToArray();
        }

        internal static string CreateThemeFingerprint(
            PowerPointSlide slide) {
            if (slide == null) throw new ArgumentNullException(nameof(slide));
            string theme = (slide.SlidePart.ThemeOverridePart
                    ?? slide.SlidePart.SlideLayoutPart?.ThemeOverridePart)?
                .ThemeOverride?.OuterXml ?? string.Empty;
            string colorMap = (slide.SlidePart.Slide?.ColorMapOverride
                    ?? slide.SlidePart.SlideLayoutPart?.SlideLayout?
                        .ColorMapOverride)?.OuterXml
                ?? string.Empty;
            return theme + "\n" + colorMap;
        }

        internal static IReadOnlyList<string>
            CreateClassicColorFingerprints(PowerPointSlide slide) {
            if (slide == null) throw new ArgumentNullException(nameof(slide));
            A.ColorScheme? colors = (slide.SlidePart.ThemeOverridePart
                    ?? slide.SlidePart.SlideLayoutPart?.ThemeOverridePart)?
                .ThemeOverride?.ColorScheme
                ?? slide.SlidePart.SlideLayoutPart?.SlideMasterPart?.ThemePart?
                    .Theme?.ThemeElements?.ColorScheme;
            OpenXmlElement?[] slots = {
                colors?.Light1Color,
                colors?.Dark1Color,
                colors?.Accent4Color,
                colors?.Dark2Color,
                colors?.Light2Color,
                colors?.Accent1Color,
                colors?.Accent2Color,
                colors?.Accent3Color
            };
            return slots.Select(slot => slot?.OuterXml ?? string.Empty)
                .ToArray();
        }

        internal LegacyPptTransition? Transition { get; }

        internal IReadOnlyList<LegacyPptComment> Comments { get; }

        internal LegacyPptNotesProjection? Notes { get; }

        internal IReadOnlyList<LegacyPptShapeProjection> Shapes { get; }

        internal bool TryGetShape(uint openXmlShapeId, out LegacyPptShapeProjection? projection) =>
            _shapesByOpenXmlId.TryGetValue(openXmlShapeId, out projection);
    }

    /// <summary>Maps projected speaker-note text to its binary NotesContainer.</summary>
    internal sealed class LegacyPptNotesProjection {
        internal LegacyPptNotesProjection(uint persistId, uint notesId,
            string text, bool followsMasterObjects, string? themePartUri,
            string themeFingerprint,
            IReadOnlyList<string> classicColorFingerprints,
            string backgroundFingerprint) {
            PersistId = persistId;
            NotesId = notesId;
            Text = text ?? string.Empty;
            FollowsMasterObjects = followsMasterObjects;
            ThemePartUri = themePartUri;
            ThemeFingerprint = themeFingerprint
                ?? throw new ArgumentNullException(nameof(themeFingerprint));
            ClassicColorFingerprints = new ReadOnlyCollection<string>(
                (classicColorFingerprints
                    ?? throw new ArgumentNullException(
                        nameof(classicColorFingerprints))).ToArray());
            if (ClassicColorFingerprints.Count != 8) {
                throw new ArgumentException(
                    "A projected classic color scheme requires eight fingerprints.",
                    nameof(classicColorFingerprints));
            }
            BackgroundFingerprint = backgroundFingerprint
                ?? throw new ArgumentNullException(nameof(backgroundFingerprint));
        }

        internal uint PersistId { get; }

        internal uint NotesId { get; }

        internal string Text { get; }

        internal bool FollowsMasterObjects { get; }

        internal string? ThemePartUri { get; }

        internal string ThemeFingerprint { get; }

        internal IReadOnlyList<string> ClassicColorFingerprints { get; }

        internal string BackgroundFingerprint { get; }

        internal bool ThemeMatches(NotesSlidePart part) => string.Equals(
            ThemeFingerprint, CreateThemeFingerprint(part),
            StringComparison.Ordinal);

        internal bool BackgroundMatches(NotesSlidePart part) => string.Equals(
            BackgroundFingerprint, CreateBackgroundFingerprint(part),
            StringComparison.Ordinal);

        internal bool MasterObjectsMatch(NotesSlidePart part) =>
            FollowsMasterObjects
                == (part.NotesSlide?.ShowMasterShapes?.Value != false);

        internal IReadOnlyList<int> GetChangedClassicColorSlots(
            NotesSlidePart part) {
            IReadOnlyList<string> current =
                CreateClassicColorFingerprints(part);
            return Enumerable.Range(0, ClassicColorFingerprints.Count)
                .Where(index => !string.Equals(
                    ClassicColorFingerprints[index], current[index],
                    StringComparison.Ordinal))
                .ToArray();
        }

        internal static string CreateThemeFingerprint(NotesSlidePart part) {
            if (part == null) throw new ArgumentNullException(nameof(part));
            string theme = part.ThemeOverridePart?.ThemeOverride?.OuterXml
                ?? string.Empty;
            string colorMap = part.NotesSlide?.ColorMapOverride?.OuterXml
                ?? string.Empty;
            return theme + "\n" + colorMap;
        }

        internal static IReadOnlyList<string>
            CreateClassicColorFingerprints(NotesSlidePart part) {
            if (part == null) throw new ArgumentNullException(nameof(part));
            A.ColorScheme? colors = part.ThemeOverridePart?.ThemeOverride?
                .ColorScheme
                ?? part.NotesMasterPart?.ThemePart?.Theme?.ThemeElements?
                    .ColorScheme;
            OpenXmlElement?[] slots = {
                colors?.Light1Color,
                colors?.Dark1Color,
                colors?.Accent4Color,
                colors?.Dark2Color,
                colors?.Light2Color,
                colors?.Accent1Color,
                colors?.Accent2Color,
                colors?.Accent3Color
            };
            return slots.Select(slot => slot?.OuterXml ?? string.Empty)
                .ToArray();
        }

        internal static string CreateBackgroundFingerprint(
            NotesSlidePart part) {
            if (part == null) throw new ArgumentNullException(nameof(part));
            return part.NotesSlide?.CommonSlideData?.Background?.OuterXml
                ?? string.Empty;
        }
    }

    /// <summary>Maps one projected Open XML shape to its OfficeArt shape container.</summary>
    internal sealed class LegacyPptShapeProjection {
        internal LegacyPptShapeProjection(uint openXmlShapeId, uint officeArtShapeId, long recordOffset,
            LegacyPptShapeKind kind, LegacyPptBounds bounds, string text,
            string? textFormattingFingerprint,
            IReadOnlyList<LegacyPptInteraction> shapeInteractions,
            IReadOnlyList<LegacyPptTextInteraction> textInteractions,
            LegacyPptAnimation? animation,
            ISet<uint> projectableSoundIds) {
            OpenXmlShapeId = openXmlShapeId;
            OfficeArtShapeId = officeArtShapeId;
            RecordOffset = recordOffset;
            Kind = kind;
            Bounds = bounds;
            Text = text ?? string.Empty;
            TextFormattingFingerprint = textFormattingFingerprint;
            ShapeInteractions = new ReadOnlyCollection<LegacyPptInteraction>(
                shapeInteractions.ToArray());
            TextInteractions = new ReadOnlyCollection<LegacyPptTextInteraction>(
                textInteractions.ToArray());
            Animation = animation;
            CanEditInteractions = ShapeInteractions.All(interaction =>
                    IsEditableInteraction(interaction, projectableSoundIds))
                && TextInteractions.All(item => IsEditableInteraction(
                    item.Interaction, projectableSoundIds))
                && ShapeInteractions.GroupBy(item => item.Trigger)
                    .All(group => group.Count() == 1)
                && !HasOverlappingTextTriggers(TextInteractions);
            CanEditAnimation = animation == null || IsEditableAnimation(
                animation, projectableSoundIds);
        }

        internal uint OpenXmlShapeId { get; }

        internal uint OfficeArtShapeId { get; }

        internal long RecordOffset { get; }

        internal LegacyPptShapeKind Kind { get; }

        internal LegacyPptBounds Bounds { get; }

        internal string Text { get; }

        internal string? TextFormattingFingerprint { get; }

        internal IReadOnlyList<LegacyPptInteraction> ShapeInteractions { get; }

        internal IReadOnlyList<LegacyPptTextInteraction> TextInteractions { get; }

        internal LegacyPptAnimation? Animation { get; }

        internal bool CanEditInteractions { get; }

        internal bool CanEditAnimation { get; }

        private static bool IsEditableAnimation(LegacyPptAnimation animation,
            ISet<uint> projectableSoundIds) {
            const uint editableFlags = 0x00004055U;
            if ((animation.RawFlags & ~editableFlags) != 0
                || animation.OleVerb != 0
                || animation.RawUnused != 0
                || animation.HasSoundOverride
                || animation.SlideCount != ushort.MaxValue
                || animation.Automatic && animation.DelayMilliseconds < 0
                || !animation.Automatic && animation.DelayMilliseconds != 0
                || animation.PlaysOnShapeClick
                || animation.Synchronous
                || animation.HiddenWhileNotPlaying) return false;
            return !animation.PlaysSound
                || projectableSoundIds.Contains(animation.SoundIdReference);
        }

        private static bool IsEditableInteraction(LegacyPptInteraction interaction,
            ISet<uint> projectableSoundIds) {
            byte allowedFlags = interaction.Action ==
                LegacyPptInteractionAction.CustomShow ? (byte)0x07 : (byte)0x03;
            if (interaction.OleVerb != 0
                || (interaction.Flags & ~allowedFlags) != 0) return false;
            if (interaction.SoundIdReference != 0
                && !projectableSoundIds.Contains(interaction.SoundIdReference)) {
                return false;
            }
            if (interaction.Action == LegacyPptInteractionAction.Macro) {
                return !string.IsNullOrEmpty(interaction.Name)
                    && interaction.Jump == LegacyPptInteractionJump.None
                    && interaction.HyperlinkType == LegacyPptHyperlinkType.Nil
                    && interaction.HyperlinkIdReference == 0;
            }
            if (interaction.Action == LegacyPptInteractionAction.RunProgram) {
                return !string.IsNullOrEmpty(interaction.Name)
                    && interaction.Jump == LegacyPptInteractionJump.None
                    && interaction.HyperlinkType == LegacyPptHyperlinkType.Nil
                    && interaction.HyperlinkIdReference == 0
                    && Uri.TryCreate(interaction.Name, UriKind.RelativeOrAbsolute,
                        out _);
            }
            if (interaction.Action == LegacyPptInteractionAction.CustomShow) {
                return interaction.CustomShow?.IsEditable == true
                    && interaction.Jump == LegacyPptInteractionJump.None
                    && interaction.HyperlinkType == LegacyPptHyperlinkType.Nil
                    && interaction.HyperlinkIdReference == 0;
            }
            if (interaction.Action == LegacyPptInteractionAction.Jump) {
                return interaction.Jump != LegacyPptInteractionJump.None
                    && interaction.HyperlinkType == LegacyPptHyperlinkType.Nil
                    && interaction.HyperlinkIdReference == 0
                    && string.IsNullOrEmpty(interaction.Name);
            }
            if (interaction.Action != LegacyPptInteractionAction.Hyperlink) return false;
            if (interaction.Jump != LegacyPptInteractionJump.None
                || !string.IsNullOrEmpty(interaction.Name)
                || interaction.HyperlinkType == LegacyPptHyperlinkType.CustomShow
                || interaction.Hyperlink != null
                && interaction.Hyperlink.ExtensionFlags != 0) return false;
            return (interaction.HyperlinkType != LegacyPptHyperlinkType.SlideNumber
                    && interaction.Hyperlink?.Uri != null)
                || (interaction.HyperlinkType == LegacyPptHyperlinkType.SlideNumber
                    && interaction.Hyperlink?.IsInternalSlideTarget == true)
                || interaction.HyperlinkType == LegacyPptHyperlinkType.NextSlide
                || interaction.HyperlinkType == LegacyPptHyperlinkType.PreviousSlide
                || interaction.HyperlinkType == LegacyPptHyperlinkType.FirstSlide
                || interaction.HyperlinkType == LegacyPptHyperlinkType.LastSlide;
        }

        private static bool HasOverlappingTextTriggers(
            IReadOnlyList<LegacyPptTextInteraction> interactions) {
            foreach (IGrouping<LegacyPptInteractionTrigger, LegacyPptTextInteraction> group
                     in interactions.GroupBy(item => item.Interaction.Trigger)) {
                int previousEnd = -1;
                foreach (LegacyPptTextInteraction item in group.OrderBy(item => item.Start)) {
                    if (item.Start < previousEnd) return true;
                    previousEnd = Math.Max(previousEnd,
                        checked(item.Start + item.Length));
                }
            }
            return false;
        }
    }
}
