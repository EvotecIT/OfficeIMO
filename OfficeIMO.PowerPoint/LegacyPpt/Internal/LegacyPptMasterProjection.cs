using System.Collections.ObjectModel;
using OfficeIMO.OpenXml.Internal;

namespace OfficeIMO.PowerPoint.LegacyPpt.Internal {
    /// <summary>Maps one projected slide master and its editable theme back to a binary persist object.</summary>
    internal sealed class LegacyPptMasterProjection {
        private readonly IReadOnlyDictionary<uint, LegacyPptShapeProjection>
            _shapesByOpenXmlId;

        internal LegacyPptMasterProjection(string masterPartUri, string? themePartUri,
            uint persistId, string themeFingerprint,
            IReadOnlyList<string> classicColorFingerprints,
            string backgroundFingerprint, bool? followsMasterObjects,
            IReadOnlyList<LegacyPptShapeProjection> shapes,
            string? textStylesFingerprint = null,
            bool canEditTextStyles = false) {
            MasterPartUri = masterPartUri ?? throw new ArgumentNullException(nameof(masterPartUri));
            ThemePartUri = themePartUri;
            PersistId = persistId;
            ThemeFingerprint = themeFingerprint ?? throw new ArgumentNullException(nameof(themeFingerprint));
            BackgroundFingerprint = backgroundFingerprint
                ?? throw new ArgumentNullException(nameof(backgroundFingerprint));
            FollowsMasterObjects = followsMasterObjects;
            TextStylesFingerprint = textStylesFingerprint;
            CanEditTextStyles = canEditTextStyles;
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

        internal string? TextStylesFingerprint { get; }

        internal bool CanEditTextStyles { get; }

        internal IReadOnlyList<LegacyPptShapeProjection> Shapes { get; }

        internal bool TryGetShape(uint openXmlShapeId,
            out LegacyPptShapeProjection? projection) =>
            _shapesByOpenXmlId.TryGetValue(openXmlShapeId, out projection);

        internal bool ThemeMatches(LegacyPptPartView masterPart) => string.Equals(
            ThemeFingerprint, CreateThemeFingerprint(masterPart), StringComparison.Ordinal);

        internal bool BackgroundMatches(LegacyPptPartView masterPart) => string.Equals(
            BackgroundFingerprint, CreateBackgroundFingerprint(masterPart),
            StringComparison.Ordinal);

        internal bool TextStylesMatch(LegacyPptPartView masterPart) =>
            TextStylesFingerprint == null || string.Equals(
                TextStylesFingerprint, CreateTextStylesFingerprint(masterPart),
                StringComparison.Ordinal);

        internal bool MasterObjectsMatch(LegacyPptPartView masterPart) =>
            !FollowsMasterObjects.HasValue
            || FollowsMasterObjects.Value
                == (masterPart.Root.Bool("showMasterSp") != false);

        internal IReadOnlyList<int> GetChangedClassicColorSlots(
            LegacyPptPartView masterPart) {
            IReadOnlyList<string> current = CreateClassicColorFingerprints(masterPart);
            return Enumerable.Range(0, ClassicColorFingerprints.Count)
                .Where(index => !string.Equals(ClassicColorFingerprints[index],
                    current[index], StringComparison.Ordinal))
                .ToArray();
        }

        /// <summary>
        /// The theme fingerprint of a slide, notes or handout master (its theme and color map), or of a title-master
        /// layout (its theme override and color-map override).
        /// </summary>
        internal static string CreateThemeFingerprint(LegacyPptPartView masterPart) {
            if (masterPart.IsNone) throw new ArgumentNullException(nameof(masterPart));
            if (masterPart.IsSlideLayout) {
                return LegacyPptProjectionDigest.Create(
                    Markup(masterPart.ThemeOverride.Root),
                    Markup(masterPart.Root.Child(LegacyPptXml.P, "clrMapOvr")));
            }
            return LegacyPptProjectionDigest.Create(Markup(masterPart.Theme.Root),
                Markup(masterPart.Root.Child(LegacyPptXml.P, "clrMap")));
        }

        internal static IReadOnlyList<string> CreateClassicColorFingerprints(
            LegacyPptPartView masterPart) {
            if (masterPart.IsNone) throw new ArgumentNullException(nameof(masterPart));
            if (masterPart.IsSlideLayout) {
                OfficeXmlNode colors = masterPart.ThemeOverride.Root.Child(LegacyPptXml.A, "clrScheme");
                if (colors.IsNone) colors = ThemeColorScheme(masterPart.SlideMaster.Theme);
                return CreateClassicColorFingerprints(colors);
            }
            return CreateClassicColorFingerprints(ThemeColorScheme(masterPart.Theme));
        }

        /// <summary>The <c>a:clrScheme</c> of a theme part's theme elements.</summary>
        internal static OfficeXmlNode ThemeColorScheme(LegacyPptPartView themePart) =>
            themePart.Root.Path(LegacyPptXml.A, "themeElements", "clrScheme");

        /// <summary>The eight classic color-scheme slots of a DrawingML color scheme, in binary order.</summary>
        internal static IReadOnlyList<string> CreateClassicColorFingerprints(
            OfficeXmlNode colors) {
            string[] slots = {
                "lt1", "dk1", "accent4", "dk2", "lt2", "accent1", "accent2", "accent3"
            };
            return slots.Select(slot => Markup(colors.Child(LegacyPptXml.A, slot))).ToArray();
        }

        internal static string CreateBackgroundFingerprint(
            LegacyPptPartView masterPart) {
            if (masterPart.IsNone) throw new ArgumentNullException(nameof(masterPart));
            return LegacyPptBackgroundProjectionFingerprint.Create(masterPart);
        }

        internal static string CreateTextStylesFingerprint(
            LegacyPptPartView masterPart) {
            if (masterPart.IsNone) throw new ArgumentNullException(nameof(masterPart));
            OfficeXmlNode textStyles = masterPart.Root.Child(LegacyPptXml.P, "txStyles");
            return Markup(textStyles)
                + LegacyPptShapeProjection.CreatePictureBulletImageFingerprint(
                    textStyles, masterPart);
        }

        /// <summary>The canonical digest of an element, or empty for none.</summary>
        internal static string Markup(OfficeXmlNode element) =>
            element.IsNone ? string.Empty : LegacyPptXml.Digest(element);
    }
}
