using System.Collections.ObjectModel;
using OfficeIMO.OpenXml.Internal;
using OfficeIMO.PowerPoint.LegacyPpt.Model;

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

        internal bool BackgroundMatches(LegacyPptPartView slidePart) => string.Equals(
            BackgroundFingerprint, CreateBackgroundFingerprint(slidePart),
            StringComparison.Ordinal);

        internal bool MasterObjectsMatch(LegacyPptPartView slidePart) =>
            FollowsMasterObjects
                == (slidePart.Root.Bool("showMasterSp") != false);

        /// <summary>
        /// The background fingerprint of a slide: its own background, or the background its layout supplies when it
        /// has none.
        /// </summary>
        internal static string CreateBackgroundFingerprint(
            LegacyPptPartView slidePart) {
            if (slidePart.IsNone) throw new ArgumentNullException(nameof(slidePart));
            OfficeXmlNode background = slidePart.Root.Child(LegacyPptXml.P, "cSld").Child(LegacyPptXml.P, "bg");
            if (!background.IsNone) {
                return LegacyPptBackgroundProjectionFingerprint.Create(
                    slidePart, background);
            }
            LegacyPptPartView layoutPart = slidePart.SlideLayout;
            return layoutPart.IsNone
                ? string.Empty
                : LegacyPptBackgroundProjectionFingerprint.Create(layoutPart);
        }

        internal bool ThemeMatches(LegacyPptPartView slidePart) => string.Equals(
            ThemeFingerprint, CreateThemeFingerprint(slidePart),
            StringComparison.Ordinal);

        internal IReadOnlyList<int> GetChangedClassicColorSlots(
            LegacyPptPartView slidePart) {
            IReadOnlyList<string> current =
                CreateClassicColorFingerprints(slidePart);
            return Enumerable.Range(0, ClassicColorFingerprints.Count)
                .Where(index => !string.Equals(
                    ClassicColorFingerprints[index], current[index],
                    StringComparison.Ordinal))
                .ToArray();
        }

        /// <summary>The theme override a slide uses: its own, or its layout's.</summary>
        internal static LegacyPptPartView EffectiveThemeOverride(LegacyPptPartView slidePart) {
            LegacyPptPartView themeOverride = slidePart.ThemeOverride;
            return themeOverride.IsNone ? slidePart.SlideLayout.ThemeOverride : themeOverride;
        }

        internal static string CreateThemeFingerprint(
            LegacyPptPartView slidePart) {
            if (slidePart.IsNone) throw new ArgumentNullException(nameof(slidePart));
            string theme = LegacyPptMasterProjection.Markup(EffectiveThemeOverride(slidePart).Root);
            OfficeXmlNode colorMap = slidePart.Root.Child(LegacyPptXml.P, "clrMapOvr");
            if (colorMap.IsNone) colorMap = slidePart.SlideLayout.Root.Child(LegacyPptXml.P, "clrMapOvr");
            return LegacyPptProjectionDigest.Create(theme, LegacyPptMasterProjection.Markup(colorMap));
        }

        internal static IReadOnlyList<string>
            CreateClassicColorFingerprints(LegacyPptPartView slidePart) {
            if (slidePart.IsNone) throw new ArgumentNullException(nameof(slidePart));
            OfficeXmlNode colors = EffectiveThemeOverride(slidePart).Root.Child(LegacyPptXml.A, "clrScheme");
            if (colors.IsNone) {
                colors = LegacyPptMasterProjection.ThemeColorScheme(slidePart.SlideLayout.SlideMaster.Theme);
            }
            return LegacyPptMasterProjection.CreateClassicColorFingerprints(colors);
        }

        internal LegacyPptTransition? Transition { get; }

        internal IReadOnlyList<LegacyPptComment> Comments { get; }

        internal LegacyPptNotesProjection? Notes { get; }

        internal IReadOnlyList<LegacyPptShapeProjection> Shapes { get; }

        internal bool TryGetShape(uint openXmlShapeId, out LegacyPptShapeProjection? projection) =>
            _shapesByOpenXmlId.TryGetValue(openXmlShapeId, out projection);
    }
}
