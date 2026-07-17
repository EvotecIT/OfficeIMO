using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Presentation;
using OfficeIMO.PowerPoint.LegacyPpt.Model;
using A = DocumentFormat.OpenXml.Drawing;

namespace OfficeIMO.PowerPoint {
    public sealed partial class PowerPointPresentation {
        private static void ProjectLegacyNotesPage(PowerPointSlide slide,
            LegacyPptNotesPage source,
            IReadOnlyDictionary<uint, SlidePart> slidePartsByLegacyId,
            LegacyPptSoundProjectionContext soundContext) {
            slide.Notes.Text = source.Text;
            NotesSlidePart part = slide.SlidePart.NotesSlidePart
                ?? throw new InvalidDataException("The projected slide has no notes part.");
            NotesSlide target = part.NotesSlide
                ?? throw new InvalidDataException("The projected notes part has no NotesSlide root.");

            if (source.Shapes.Count > 0) {
                target.CommonSlideData = new CommonSlideData(CreateLegacyShapeTree(part,
                    source.Shapes, source.ConnectorRules,
                    slidePartsByLegacyId, soundContext)) {
                    Name = "Binary Notes Page"
                };
            }
            ApplyLegacyRoundTripTheme(part, source.RoundTripTheme);
            target.ShowMasterShapes = source.FollowsMasterObjects;
            if (!source.FollowsMasterColorScheme && source.ColorScheme != null
                && source.RoundTripTheme?.ThemeXml == null) {
                ApplyLegacyColorScheme(part, source.ColorScheme);
            }
            if (!source.FollowsMasterBackground && target.CommonSlideData != null) {
                if (source.Background != null) {
                    ApplyLegacyBackground(part, target.CommonSlideData, source.Background);
                } else if (source.ColorScheme != null) {
                    ApplyLegacyBackground(target.CommonSlideData,
                        source.ColorScheme.Background);
                }
            }
            target.Save();
        }

        private static void ApplyLegacyColorScheme(NotesSlidePart notesPart,
            LegacyPptColorScheme source) {
            A.ColorScheme? masterScheme = notesPart.NotesMasterPart?.ThemePart?
                .Theme?.ThemeElements?.ColorScheme;
            A.ColorScheme target = masterScheme?.CloneNode(true) as A.ColorScheme
                ?? new A.ColorScheme { Name = "Binary PowerPoint" };
            SetLegacyThemeColors(target, source);
            ThemeOverridePart overridePart = notesPart.ThemeOverridePart
                ?? notesPart.AddNewPart<ThemeOverridePart>();
            overridePart.ThemeOverride = new A.ThemeOverride(target);
            overridePart.ThemeOverride.Save();
        }
    }
}
