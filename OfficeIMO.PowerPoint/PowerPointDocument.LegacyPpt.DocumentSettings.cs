using OfficeIMO.OpenXml.Internal;
using OfficeIMO.PowerPoint.LegacyPpt;
using OfficeIMO.PowerPoint.LegacyPpt.Internal;
using OfficeIMO.PowerPoint.LegacyPpt.Model;
using OfficeIMO.PowerPoint.OpenXml;

namespace OfficeIMO.PowerPoint {
    public sealed partial class PowerPointDocument {
        private static void ApplyLegacyDocumentSettings(PowerPointDocument target,
            LegacyPptImportParts parts, LegacyPptDocument source) {
            LegacyPptDocumentSettings? settings = source.DocumentSettings;
            PowerPointSlideSizeType sizeType = settings?.SlideSizeType is LegacyPptSlideSizeType type
                ? MapLegacySlideSizeType(type, source.SlideWidth, source.SlideHeight)
                : PowerPointSlideSizeType.Custom;
            target.SlideSize.SetSize(OfficeIMO.OfficeLength.Emus(ToEmus(source.SlideWidth)),                OfficeIMO.OfficeLength.Emus(ToEmus(source.SlideHeight)),sizeType);
            if (settings == null) return;

            PowerPointPartRef presentationPart = parts.PresentationPart;
            OfficeXmlNode root = presentationPart.Root;
            if (settings.ServerZoomNumerator > 0 && settings.ServerZoomDenominator > 0) {
                long scaledZoom = ((long)settings.ServerZoomNumerator * 100000L
                    + settings.ServerZoomDenominator / 2L) / settings.ServerZoomDenominator;
                root.Set("serverZoom", checked((int)Math.Min(int.MaxValue,
                    Math.Max(1L, scaledZoom))));
            }
            if (settings.NotesWidth > 0 && settings.NotesHeight > 0) {
                root.SetChild(LegacyPptXml.P, "notesSz", PPresentationView.Order)
                    .Set("cx", checked((int)ToEmus(settings.NotesWidth)))
                    .Set("cy", checked((int)ToEmus(settings.NotesHeight)));
            }
            root.Set("firstSlideNum", settings.FirstSlideNumber);
            root.Set("showSpecialPlsOnTitleSld", !settings.OmitTitlePlaceholders);
            root.Set("rtl", settings.RightToLeft);
            root.Set("embedTrueTypeFonts", settings.SaveWithFonts);
            root.Set("saveSubsetFonts", settings.SaveWithFonts);
            presentationPart.MarkEdited();
            PowerPointPartRef viewPart = presentationPart.RelatedPart(LegacyViewPropertiesRelationshipType);
            OfficeXmlNode view = viewPart.IsNone ? default : viewPart.Root;
            if (!view.IsNone && view.Is(LegacyPptXml.P, "viewPr")) {
                view.Set("showComments", settings.ShowComments);
                viewPart.MarkEdited();
            }
        }

        private static PowerPointSlideSizeType MapLegacySlideSizeType(LegacyPptSlideSizeType type,
            int width, int height) => type switch {
                LegacyPptSlideSizeType.Screen => width * 9L == height * 16L
                    ? PowerPointSlideSizeType.Screen16x9
                    : PowerPointSlideSizeType.Screen4x3,
                LegacyPptSlideSizeType.LetterPaper => PowerPointSlideSizeType.Letter,
                LegacyPptSlideSizeType.A4Paper => PowerPointSlideSizeType.A4,
                LegacyPptSlideSizeType.Film35Mm => PowerPointSlideSizeType.Film35mm,
                LegacyPptSlideSizeType.Overhead => PowerPointSlideSizeType.Overhead,
                LegacyPptSlideSizeType.Banner => PowerPointSlideSizeType.Banner,
                _ => PowerPointSlideSizeType.Custom
            };
    }
}
