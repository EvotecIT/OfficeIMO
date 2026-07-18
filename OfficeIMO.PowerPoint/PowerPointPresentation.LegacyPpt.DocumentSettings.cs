using DocumentFormat.OpenXml.Presentation;
using OfficeIMO.PowerPoint.LegacyPpt;
using OfficeIMO.PowerPoint.LegacyPpt.Model;

namespace OfficeIMO.PowerPoint {
    public sealed partial class PowerPointPresentation {
        private static void ApplyLegacyDocumentSettings(PowerPointPresentation target,
            LegacyPptPresentation source) {
            LegacyPptDocumentSettings? settings = source.DocumentSettings;
            SlideSizeValues sizeType = settings?.SlideSizeType is LegacyPptSlideSizeType type
                ? MapLegacySlideSizeType(type, source.SlideWidth, source.SlideHeight)
                : SlideSizeValues.Custom;
            target.SlideSize.SetSizeEmus(ToEmus(source.SlideWidth),
                ToEmus(source.SlideHeight), sizeType);
            if (settings == null) return;

            Presentation root = target.PresentationRoot;
            if (settings.ServerZoomNumerator > 0 && settings.ServerZoomDenominator > 0) {
                long scaledZoom = ((long)settings.ServerZoomNumerator * 100000L
                    + settings.ServerZoomDenominator / 2L) / settings.ServerZoomDenominator;
                root.ServerZoom = checked((int)Math.Min(int.MaxValue,
                    Math.Max(1L, scaledZoom)));
            }
            if (settings.NotesWidth > 0 && settings.NotesHeight > 0) {
                root.NotesSize = new NotesSize {
                    Cx = checked((int)ToEmus(settings.NotesWidth)),
                    Cy = checked((int)ToEmus(settings.NotesHeight))
                };
            }
            root.FirstSlideNum = settings.FirstSlideNumber;
            root.ShowSpecialPlaceholderOnTitleSlide = !settings.OmitTitlePlaceholders;
            root.RightToLeft = settings.RightToLeft;
            root.EmbedTrueTypeFonts = settings.SaveWithFonts;
            root.SaveSubsetFonts = settings.SaveWithFonts;
            ViewProperties? view = target._presentationPart.ViewPropertiesPart?.ViewProperties;
            if (view != null) view.ShowComments = settings.ShowComments;
        }

        private static SlideSizeValues MapLegacySlideSizeType(LegacyPptSlideSizeType type,
            int width, int height) => type switch {
                LegacyPptSlideSizeType.Screen => width * 9L == height * 16L
                    ? SlideSizeValues.Screen16x9
                    : SlideSizeValues.Screen4x3,
                LegacyPptSlideSizeType.LetterPaper => SlideSizeValues.Letter,
                LegacyPptSlideSizeType.A4Paper => SlideSizeValues.A4,
                LegacyPptSlideSizeType.Film35Mm => SlideSizeValues.Film35mm,
                LegacyPptSlideSizeType.Overhead => SlideSizeValues.Overhead,
                LegacyPptSlideSizeType.Banner => SlideSizeValues.Banner,
                _ => SlideSizeValues.Custom
            };
    }
}
