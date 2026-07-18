using System.Collections.ObjectModel;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using OfficeIMO.PowerPoint.LegacyPpt.Model;
using OfficeIMO.PowerPoint.LegacyPpt.Write;
using A = DocumentFormat.OpenXml.Drawing;
using P = DocumentFormat.OpenXml.Presentation;

namespace OfficeIMO.PowerPoint.LegacyPpt.Internal {
    /// <summary>Maps one projected slide to its binary persist object.</summary>
    internal sealed class LegacyPptSlideProjection {
        private readonly IReadOnlyDictionary<uint, LegacyPptShapeProjection> _shapesByOpenXmlId;

        internal LegacyPptSlideProjection(string slidePartUri,
            string? layoutPartUri, uint persistId, uint slideId, uint masterId,
            uint layoutType, IReadOnlyList<byte> layoutPlaceholderTypes,
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
            LayoutType = layoutType;
            LayoutPlaceholderTypes = new ReadOnlyCollection<byte>(
                (layoutPlaceholderTypes ?? throw new ArgumentNullException(
                    nameof(layoutPlaceholderTypes))).ToArray());
            if (LayoutPlaceholderTypes.Count != 8) {
                throw new ArgumentException(
                    "A binary slide layout signature requires eight placeholder slots.",
                    nameof(layoutPlaceholderTypes));
            }
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

        internal uint LayoutType { get; }

        internal IReadOnlyList<byte> LayoutPlaceholderTypes { get; }

        internal bool LayoutContractMatches(uint layoutType,
            IReadOnlyList<byte> placeholderTypes) =>
            LayoutType == layoutType
            && placeholderTypes != null
            && LayoutPlaceholderTypes.SequenceEqual(placeholderTypes);

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
            P.Background? background = slide.SlidePart.Slide?.CommonSlideData?
                .Background;
            if (background != null) {
                return LegacyPptBackgroundProjectionFingerprint.Create(
                    slide.SlidePart, background);
            }
            SlideLayoutPart? layoutPart = slide.SlidePart.SlideLayoutPart;
            return layoutPart == null
                ? string.Empty
                : LegacyPptBackgroundProjectionFingerprint.Create(layoutPart,
                    layoutPart.SlideLayout?.CommonSlideData?.Background);
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
}
