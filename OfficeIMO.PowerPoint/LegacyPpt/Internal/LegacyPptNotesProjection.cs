using System.Collections.ObjectModel;
using OfficeIMO.OpenXml.Internal;

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

        internal bool ThemeMatches(LegacyPptPartView part) => string.Equals(
            ThemeFingerprint, CreateThemeFingerprint(part),
            StringComparison.Ordinal);

        internal bool BackgroundMatches(LegacyPptPartView part) => string.Equals(
            BackgroundFingerprint, CreateBackgroundFingerprint(part),
            StringComparison.Ordinal);

        internal bool MasterObjectsMatch(LegacyPptPartView part) =>
            FollowsMasterObjects
                == (part.Root.Bool("showMasterSp") != false);

        internal IReadOnlyList<int> GetChangedClassicColorSlots(
            LegacyPptPartView part) {
            IReadOnlyList<string> current =
                CreateClassicColorFingerprints(part);
            return Enumerable.Range(0, ClassicColorFingerprints.Count)
                .Where(index => !string.Equals(
                    ClassicColorFingerprints[index], current[index],
                    StringComparison.Ordinal))
                .ToArray();
        }

        /// <summary>The theme fingerprint of a notes slide: its theme override and color-map override.</summary>
        internal static string CreateThemeFingerprint(LegacyPptPartView part) {
            if (part.IsNone) throw new ArgumentNullException(nameof(part));
            return LegacyPptProjectionDigest.Create(
                LegacyPptMasterProjection.Markup(part.ThemeOverride.Root),
                LegacyPptMasterProjection.Markup(part.Root.Child(LegacyPptXml.P, "clrMapOvr")));
        }

        internal static IReadOnlyList<string>
            CreateClassicColorFingerprints(LegacyPptPartView part) {
            if (part.IsNone) throw new ArgumentNullException(nameof(part));
            OfficeXmlNode colors = part.ThemeOverride.Root.Child(LegacyPptXml.A, "clrScheme");
            if (colors.IsNone) {
                colors = LegacyPptMasterProjection.ThemeColorScheme(part
                    .RelatedPart(LegacyPptPackageView.NotesMasterRelationshipType).Theme);
            }
            return LegacyPptMasterProjection.CreateClassicColorFingerprints(colors);
        }

        internal static string CreateBackgroundFingerprint(
            LegacyPptPartView part) {
            if (part.IsNone) throw new ArgumentNullException(nameof(part));
            return LegacyPptBackgroundProjectionFingerprint.Create(part);
        }
    }
}
