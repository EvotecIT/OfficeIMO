using System.Collections.ObjectModel;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using OfficeIMO.PowerPoint.LegacyPpt.Model;
using OfficeIMO.PowerPoint.LegacyPpt.Write;

namespace OfficeIMO.PowerPoint.LegacyPpt.Internal {
    /// <summary>Links projected Open XML slides and shapes back to their original binary persist records.</summary>
    internal sealed class LegacyPptProjectionMap {
        private readonly IReadOnlyDictionary<string, LegacyPptSlideProjection> _slidesByPartUri;
        private readonly IReadOnlyDictionary<uint, LegacyPptSlideProjection> _slidesByLegacyId;
        private readonly IReadOnlyDictionary<string, uint> _masterIdsByLayoutPartUri;
        private readonly IReadOnlyDictionary<string, LegacyPptMasterProjection> _mastersByPartUri;
        private readonly ISet<string> _masterThemePartUris;

        private LegacyPptProjectionMap(IReadOnlyList<LegacyPptSlideProjection> slides,
            IReadOnlyDictionary<string, uint> masterIdsByLayoutPartUri,
            IReadOnlyList<LegacyPptMasterProjection> masters,
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
            _masterThemePartUris = new HashSet<string>(masters
                .Select(master => master.ThemePartUri), StringComparer.Ordinal);
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

        internal bool TryGetMaster(SlideMasterPart masterPart,
            out LegacyPptMasterProjection? projection) {
            if (masterPart == null) throw new ArgumentNullException(nameof(masterPart));
            return _mastersByPartUri.TryGetValue(masterPart.Uri.ToString(), out projection);
        }

        internal bool IsProjectedMasterPart(string partUri) =>
            partUri != null && _mastersByPartUri.ContainsKey(partUri);

        internal bool IsProjectedMasterThemePart(string partUri) =>
            partUri != null && _masterThemePartUris.Contains(partUri);

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
                    sourceSlide.PersistId, sourceSlide.SlideId, sourceSlide.MasterId,
                    sourceSlide.Hidden, sourceSlide.HeaderFooter,
                    sourceSlide.Transition, sourceSlide.Comments, shapes,
                    sourceSlide.NotesPage == null
                        ? null
                        : new LegacyPptNotesProjection(sourceSlide.NotesPage.PersistId,
                            sourceSlide.NotesPage.NotesId, sourceSlide.NotesPage.Text)));
            }
            return new LegacyPptProjectionMap(slides, CreateLayoutMasterMap(presentation, legacy),
                CreateMasterProjections(presentation, legacy),
                legacy.Hyperlinks, legacy.CustomShows,
                legacy.CustomShowsAreEditable, legacy.Sounds,
                legacy.SoundIdSeed);
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
                PowerPointShape[] projectedShapes = LegacyPptWriter
                    .ReadMasterShapesForWrite(masterPart, out _).ToArray();
                LegacyPptShape[] sourceShapes = sourceMasters[index].Shapes
                    .Where(shape => shape.Kind != LegacyPptShapeKind.Unsupported)
                    .ToArray();
                if (projectedShapes.Length != sourceShapes.Length) {
                    throw new InvalidDataException(
                        $"Projected slide master {index + 1} has {projectedShapes.Length} editable shapes, "
                        + $"but the binary source exposed {sourceShapes.Length}.");
                }
                var shapes = new List<LegacyPptShapeProjection>(projectedShapes.Length);
                for (int shapeIndex = 0; shapeIndex < projectedShapes.Length; shapeIndex++) {
                    PowerPointShape projectedShape = projectedShapes[shapeIndex];
                    uint? openXmlShapeId = projectedShape.Id;
                    if (!openXmlShapeId.HasValue) {
                        throw new InvalidDataException(
                            $"Projected slide master {index + 1}, shape {shapeIndex + 1} has no Open XML shape id.");
                    }
                    LegacyPptShape sourceShape = sourceShapes[shapeIndex];
                    string? textFormattingFingerprint = (sourceShape.TextBody.HasStyleRecord
                            || sourceShape.TextBody.HasInteractions)
                        && projectedShape.Element is DocumentFormat.OpenXml.Presentation.Shape projectedTextShape
                        ? LegacyPptTextProjection.CreateFormattingFingerprint(
                            projectedTextShape.TextBody)
                        : null;
                    shapes.Add(new LegacyPptShapeProjection(openXmlShapeId.Value,
                        sourceShape.ShapeId, sourceShape.RecordOffset,
                        sourceShape.Kind, sourceShape.Bounds, sourceShape.Text,
                        textFormattingFingerprint, sourceShape.Interactions,
                        sourceShape.TextBody.Interactions, sourceShape.Animation,
                        new HashSet<uint>()));
                }
                result.Add(new LegacyPptMasterProjection(
                    masterPart.Uri.ToString(), themePart.Uri.ToString(),
                    sourceMasters[index].PersistId,
                    LegacyPptMasterProjection.CreateThemeFingerprint(masterPart),
                    LegacyPptMasterProjection.CreateClassicColorFingerprints(masterPart),
                    shapes));
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

        internal LegacyPptMasterProjection(string masterPartUri, string themePartUri,
            uint persistId, string themeFingerprint,
            IReadOnlyList<string> classicColorFingerprints,
            IReadOnlyList<LegacyPptShapeProjection> shapes) {
            MasterPartUri = masterPartUri ?? throw new ArgumentNullException(nameof(masterPartUri));
            ThemePartUri = themePartUri ?? throw new ArgumentNullException(nameof(themePartUri));
            PersistId = persistId;
            ThemeFingerprint = themeFingerprint ?? throw new ArgumentNullException(nameof(themeFingerprint));
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

        internal string ThemePartUri { get; }

        internal uint PersistId { get; }

        internal string ThemeFingerprint { get; }

        internal IReadOnlyList<string> ClassicColorFingerprints { get; }

        internal IReadOnlyList<LegacyPptShapeProjection> Shapes { get; }

        internal bool TryGetShape(uint openXmlShapeId,
            out LegacyPptShapeProjection? projection) =>
            _shapesByOpenXmlId.TryGetValue(openXmlShapeId, out projection);

        internal bool ThemeMatches(SlideMasterPart masterPart) => string.Equals(
            ThemeFingerprint, CreateThemeFingerprint(masterPart), StringComparison.Ordinal);

        internal IReadOnlyList<int> GetChangedClassicColorSlots(
            SlideMasterPart masterPart) {
            IReadOnlyList<string> current = CreateClassicColorFingerprints(masterPart);
            return Enumerable.Range(0, ClassicColorFingerprints.Count)
                .Where(index => !string.Equals(ClassicColorFingerprints[index],
                    current[index], StringComparison.Ordinal))
                .ToArray();
        }

        internal static string CreateThemeFingerprint(SlideMasterPart masterPart) {
            if (masterPart == null) throw new ArgumentNullException(nameof(masterPart));
            string theme = masterPart.ThemePart?.Theme?.OuterXml ?? string.Empty;
            string colorMap = masterPart.SlideMaster?.ColorMap?.OuterXml ?? string.Empty;
            return theme + "\n" + colorMap;
        }

        internal static IReadOnlyList<string> CreateClassicColorFingerprints(
            SlideMasterPart masterPart) {
            if (masterPart == null) throw new ArgumentNullException(nameof(masterPart));
            DocumentFormat.OpenXml.Drawing.ColorScheme? colors = masterPart.ThemePart?
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
            return slots.Select(slot => slot?.OuterXml ?? string.Empty).ToArray();
        }
    }

    /// <summary>Maps one projected slide to its binary persist object.</summary>
    internal sealed class LegacyPptSlideProjection {
        private readonly IReadOnlyDictionary<uint, LegacyPptShapeProjection> _shapesByOpenXmlId;

        internal LegacyPptSlideProjection(string slidePartUri, uint persistId, uint slideId, uint masterId,
            bool hidden, LegacyPptHeaderFooterSettings? headerFooter,
            LegacyPptTransition? transition,
            IReadOnlyList<LegacyPptComment> comments,
            IReadOnlyList<LegacyPptShapeProjection> shapes, LegacyPptNotesProjection? notes) {
            SlidePartUri = slidePartUri ?? throw new ArgumentNullException(nameof(slidePartUri));
            PersistId = persistId;
            SlideId = slideId;
            MasterId = masterId;
            Hidden = hidden;
            HeaderFooter = headerFooter;
            Transition = transition;
            Comments = new ReadOnlyCollection<LegacyPptComment>(comments.ToArray());
            Notes = notes;
            Shapes = new ReadOnlyCollection<LegacyPptShapeProjection>(shapes.ToArray());
            _shapesByOpenXmlId = new ReadOnlyDictionary<uint, LegacyPptShapeProjection>(shapes.ToDictionary(
                shape => shape.OpenXmlShapeId));
        }

        internal string SlidePartUri { get; }

        internal uint PersistId { get; }

        internal uint SlideId { get; }

        internal uint MasterId { get; }

        internal bool Hidden { get; }

        internal LegacyPptHeaderFooterSettings? HeaderFooter { get; }

        internal LegacyPptTransition? Transition { get; }

        internal IReadOnlyList<LegacyPptComment> Comments { get; }

        internal LegacyPptNotesProjection? Notes { get; }

        internal IReadOnlyList<LegacyPptShapeProjection> Shapes { get; }

        internal bool TryGetShape(uint openXmlShapeId, out LegacyPptShapeProjection? projection) =>
            _shapesByOpenXmlId.TryGetValue(openXmlShapeId, out projection);
    }

    /// <summary>Maps projected speaker-note text to its binary NotesContainer.</summary>
    internal sealed class LegacyPptNotesProjection {
        internal LegacyPptNotesProjection(uint persistId, uint notesId, string text) {
            PersistId = persistId;
            NotesId = notesId;
            Text = text ?? string.Empty;
        }

        internal uint PersistId { get; }

        internal uint NotesId { get; }

        internal string Text { get; }
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
