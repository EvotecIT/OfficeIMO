using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Presentation;
using OfficeIMO.PowerPoint.LegacyPpt;
using OfficeIMO.PowerPoint.LegacyPpt.Model;
using A = DocumentFormat.OpenXml.Drawing;

namespace OfficeIMO.PowerPoint {
    public sealed partial class PowerPointPresentation {
        private static void ApplyLegacyHeaderFooter(OpenXmlCompositeElement owner,
            CommonSlideData? commonSlideData, LegacyPptHeaderFooterSettings? source,
            bool allowHeader) {
            if (source == null) return;
            owner.RemoveAllChildren<HeaderFooter>();
            var headerFooter = new HeaderFooter {
                DateTime = source.ShowDate,
                Footer = source.ShowFooter,
                Header = allowHeader && source.ShowHeader,
                SlideNumber = source.ShowSlideNumber
            };
            OpenXmlElement? successor = owner switch {
                SlideMaster master => (OpenXmlElement?)master.TextStyles
                    ?? (OpenXmlElement?)master.SlideMasterExtensionList,
                SlideLayout layout => (OpenXmlElement?)layout.Timing
                    ?? (OpenXmlElement?)layout.Transition
                    ?? (OpenXmlElement?)layout.SlideLayoutExtensionList,
                NotesMaster notes => (OpenXmlElement?)notes.NotesStyle
                    ?? (OpenXmlElement?)notes.NotesMasterExtensionList,
                HandoutMaster handout => (OpenXmlElement?)handout.HandoutMasterExtensionList,
                _ => null
            };
            if (successor != null) owner.InsertBefore(headerFooter, successor);
            else owner.Append(headerFooter);

            if (commonSlideData == null) return;
            if (source.UseUserDate && source.UserDateText.Length > 0) {
                SetLegacyPlaceholderText(commonSlideData, PlaceholderValues.DateAndTime,
                    source.UserDateText);
            }
            if (allowHeader && source.HeaderText.Length > 0) {
                SetLegacyPlaceholderText(commonSlideData, PlaceholderValues.Header,
                    source.HeaderText);
            }
            if (source.FooterText.Length > 0) {
                SetLegacyPlaceholderText(commonSlideData, PlaceholderValues.Footer,
                    source.FooterText);
            }
        }

        private static void SetLegacyPlaceholderText(CommonSlideData commonSlideData,
            PlaceholderValues type, string text) {
            Shape? shape = commonSlideData.ShapeTree?.Elements<Shape>()
                .FirstOrDefault(candidate => candidate.NonVisualShapeProperties?
                    .ApplicationNonVisualDrawingProperties?.PlaceholderShape?.Type?.Value == type);
            if (shape?.TextBody == null) return;
            shape.TextBody.RemoveAllChildren<A.Paragraph>();
            shape.TextBody.Append(new A.Paragraph(
                new A.Run(new A.Text(text)),
                new A.EndParagraphRunProperties()));
        }

        private static LegacyPptHeaderFooterSettings? GetEffectiveLegacySlideHeaderFooter(
            LegacyPptPresentation legacy, LegacyPptMaster master) =>
            master.HeaderFooter ?? legacy.SlideHeaderFooterDefaults;
    }
}
