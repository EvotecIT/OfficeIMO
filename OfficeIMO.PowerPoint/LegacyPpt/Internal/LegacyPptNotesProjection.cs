using System.Collections.ObjectModel;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using OfficeIMO.PowerPoint.LegacyPpt.Model;
using OfficeIMO.PowerPoint.LegacyPpt.Write;
using A = DocumentFormat.OpenXml.Drawing;
using P = DocumentFormat.OpenXml.Presentation;

namespace OfficeIMO.PowerPoint.LegacyPpt.Internal {
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
            return LegacyPptBackgroundProjectionFingerprint.Create(part,
                part.NotesSlide?.CommonSlideData?.Background);
        }
    }
}
