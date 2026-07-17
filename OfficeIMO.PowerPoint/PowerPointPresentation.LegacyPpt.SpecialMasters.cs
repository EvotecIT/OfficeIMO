using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Presentation;
using OfficeIMO.PowerPoint.LegacyPpt;
using OfficeIMO.PowerPoint.LegacyPpt.Model;
using A = DocumentFormat.OpenXml.Drawing;

namespace OfficeIMO.PowerPoint {
    public sealed partial class PowerPointPresentation {
        private static void ProjectLegacySpecialMasters(PowerPointPresentation presentation,
            LegacyPptPresentation legacy,
            LegacyPptSoundProjectionContext soundContext,
            ICollection<LegacyPptDeferredProjection>
                deferredInteractions) {
            if (legacy.NotesMaster != null) {
                ProjectLegacyNotesMaster(presentation, legacy.NotesMaster,
                    legacy.NotesHeaderFooterDefaults, soundContext,
                    deferredInteractions);
            }
            if (legacy.HandoutMaster != null) {
                ProjectLegacyHandoutMaster(presentation, legacy.HandoutMaster,
                    legacy.NotesHeaderFooterDefaults, soundContext,
                    deferredInteractions);
            }
        }

        private static void ProjectLegacyNotesMaster(PowerPointPresentation presentation,
            LegacyPptSpecialMaster source,
            LegacyPptHeaderFooterSettings? headerFooter,
            LegacyPptSoundProjectionContext soundContext,
            ICollection<LegacyPptDeferredProjection>
                deferredInteractions) {
            NotesMasterPart part = PowerPointUtils.EnsureNotesMasterPart(
                presentation._presentationPart);
            NotesMaster target = part.NotesMaster
                ?? throw new InvalidDataException("The native PowerPoint scaffold has no notes master.");
            target.CommonSlideData = CreateLegacySpecialMasterCommonSlideData(part, source,
                "Binary Notes Master", soundContext, deferredInteractions);
            ApplyLegacyRoundTripTheme(part, source.RoundTripTheme);
            if (source.RoundTripTheme?.ThemeXml == null) {
                ApplyLegacySpecialMasterTheme(part.ThemePart,
                    source.ColorScheme);
            }
            ApplyLegacyHeaderFooter(target, target.CommonSlideData,
                headerFooter, allowHeader: true);
            target.Save();
        }

        private static void ProjectLegacyHandoutMaster(PowerPointPresentation presentation,
            LegacyPptSpecialMaster source,
            LegacyPptHeaderFooterSettings? headerFooter,
            LegacyPptSoundProjectionContext soundContext,
            ICollection<LegacyPptDeferredProjection>
                deferredInteractions) {
            PresentationPart presentationPart = presentation._presentationPart;
            HandoutMasterPart part = presentationPart.HandoutMasterPart
                ?? presentationPart.AddNewPart<HandoutMasterPart>();
            EnsureLegacySpecialMasterTheme(part, presentationPart);

            CommonSlideData commonSlideData = CreateLegacySpecialMasterCommonSlideData(part,
                source, "Binary Handout Master", soundContext,
                deferredInteractions);
            ColorMap colorMap = CloneLegacyColorMap(presentationPart);
            part.HandoutMaster = new HandoutMaster(commonSlideData, colorMap);
            ApplyLegacyRoundTripTheme(part, source.RoundTripTheme);
            if (source.RoundTripTheme?.ThemeXml == null) {
                ApplyLegacySpecialMasterTheme(part.ThemePart,
                    source.ColorScheme);
            }
            ApplyLegacyHeaderFooter(part.HandoutMaster, commonSlideData,
                headerFooter, allowHeader: true);
            part.HandoutMaster.Save();

            Presentation root = presentationPart.Presentation ??= new Presentation();
            HandoutMasterIdList list = root.HandoutMasterIdList ??= new HandoutMasterIdList();
            string relationshipId = presentationPart.GetIdOfPart(part);
            if (!list.Elements<HandoutMasterId>().Any(item =>
                    PowerPointUtils.GetRelationshipIdValue(item) == relationshipId)) {
                var id = new HandoutMasterId();
                PowerPointUtils.SetRelationshipIdValue(id, relationshipId);
                list.Append(id);
            }
        }

        private static CommonSlideData CreateLegacySpecialMasterCommonSlideData(
            OpenXmlPart ownerPart, LegacyPptSpecialMaster source, string name,
            LegacyPptSoundProjectionContext soundContext,
            ICollection<LegacyPptDeferredProjection>
                deferredInteractions) {
            var result = new CommonSlideData(CreateLegacyShapeTree(ownerPart, source.Shapes,
                source.ConnectorRules, soundContext: soundContext,
                deferredInteractions: deferredInteractions)) { Name = name };
            if (source.Background != null) {
                ApplyLegacyBackground(ownerPart, result, source.Background);
            } else if (source.ColorScheme != null) {
                ApplyLegacyBackground(result, source.ColorScheme.Background);
            }
            return result;
        }

        private static void EnsureLegacySpecialMasterTheme(HandoutMasterPart target,
            PresentationPart presentationPart) {
            if (target.ThemePart != null) return;
            ThemePart targetTheme = target.AddNewPart<ThemePart>();
            A.Theme? source = presentationPart.SlideMasterParts.FirstOrDefault()?
                .ThemePart?.Theme;
            targetTheme.Theme = source?.CloneNode(true) as A.Theme
                ?? new A.Theme {
                    Name = "Binary PowerPoint",
                    ThemeElements = new A.ThemeElements()
                };
        }

        private static void ApplyLegacySpecialMasterTheme(ThemePart? themePart,
            LegacyPptColorScheme? source) {
            if (themePart == null || source == null) return;
            A.Theme theme = themePart.Theme ??= new A.Theme {
                Name = "Binary PowerPoint",
                ThemeElements = new A.ThemeElements()
            };
            theme.ThemeElements ??= new A.ThemeElements();
            A.ColorScheme target = theme.ThemeElements.ColorScheme
                ??= new A.ColorScheme { Name = "Binary PowerPoint" };
            SetLegacyThemeColors(target, source);
            theme.Save();
        }

        private static ColorMap CloneLegacyColorMap(PresentationPart presentationPart) =>
            presentationPart.SlideMasterParts.FirstOrDefault()?.SlideMaster?.ColorMap?
                .CloneNode(true) as ColorMap
            ?? new ColorMap {
                Background1 = A.ColorSchemeIndexValues.Light1,
                Text1 = A.ColorSchemeIndexValues.Dark1,
                Background2 = A.ColorSchemeIndexValues.Light2,
                Text2 = A.ColorSchemeIndexValues.Dark2,
                Accent1 = A.ColorSchemeIndexValues.Accent1,
                Accent2 = A.ColorSchemeIndexValues.Accent2,
                Accent3 = A.ColorSchemeIndexValues.Accent3,
                Accent4 = A.ColorSchemeIndexValues.Accent4,
                Accent5 = A.ColorSchemeIndexValues.Accent5,
                Accent6 = A.ColorSchemeIndexValues.Accent6,
                Hyperlink = A.ColorSchemeIndexValues.Hyperlink,
                FollowedHyperlink = A.ColorSchemeIndexValues.FollowedHyperlink
            };
    }
}
