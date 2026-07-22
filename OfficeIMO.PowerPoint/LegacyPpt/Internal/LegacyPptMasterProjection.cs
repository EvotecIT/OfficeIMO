using System.Collections.ObjectModel;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using OfficeIMO.PowerPoint.LegacyPpt.Model;
using OfficeIMO.PowerPoint.LegacyPpt.Write;
using A = DocumentFormat.OpenXml.Drawing;
using P = DocumentFormat.OpenXml.Presentation;

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

        internal bool TextStylesMatch(SlideMasterPart masterPart) =>
            TextStylesFingerprint == null || string.Equals(
                TextStylesFingerprint, CreateTextStylesFingerprint(masterPart),
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
            return LegacyPptProjectionDigest.Create(theme, colorMap);
        }

        private static string CreateThemeFingerprint(ThemePart? themePart,
            OpenXmlElement? colorMap) {
            string theme = themePart?.Theme?.OuterXml ?? string.Empty;
            return LegacyPptProjectionDigest.Create(theme,
                colorMap?.OuterXml);
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
            return LegacyPptBackgroundProjectionFingerprint.Create(masterPart,
                masterPart.SlideMaster?.CommonSlideData?.Background);
        }

        internal static string CreateTextStylesFingerprint(
            SlideMasterPart masterPart) {
            if (masterPart == null) {
                throw new ArgumentNullException(nameof(masterPart));
            }
            return (masterPart.SlideMaster?.TextStyles?.OuterXml
                    ?? string.Empty)
                + LegacyPptTextProjection
                    .CreatePictureBulletImageFingerprint(masterPart
                        .SlideMaster?.TextStyles, masterPart);
        }

        internal static string CreateBackgroundFingerprint(
            NotesMasterPart masterPart) {
            if (masterPart == null) throw new ArgumentNullException(nameof(masterPart));
            return LegacyPptBackgroundProjectionFingerprint.Create(masterPart,
                masterPart.NotesMaster?.CommonSlideData?.Background);
        }

        internal static string CreateBackgroundFingerprint(
            HandoutMasterPart masterPart) {
            if (masterPart == null) throw new ArgumentNullException(nameof(masterPart));
            return LegacyPptBackgroundProjectionFingerprint.Create(masterPart,
                masterPart.HandoutMaster?.CommonSlideData?.Background);
        }

        internal static string CreateBackgroundFingerprint(
            SlideLayoutPart masterPart) {
            if (masterPart == null) throw new ArgumentNullException(nameof(masterPart));
            return LegacyPptBackgroundProjectionFingerprint.Create(masterPart,
                masterPart.SlideLayout?.CommonSlideData?.Background);
        }
    }
}
