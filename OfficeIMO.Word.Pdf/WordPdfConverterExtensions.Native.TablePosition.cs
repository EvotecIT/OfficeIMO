using W = DocumentFormat.OpenXml.Wordprocessing;
using PdfCore = OfficeIMO.Pdf;

namespace OfficeIMO.Word.Pdf {
    public static partial class WordPdfConverterExtensions {
        private static PdfCore.PdfTablePosition CreateNativeTablePosition(W.TablePositionProperties position) {
            PdfCore.PdfTableAnchor horizontalAnchor = position.HorizontalAnchor?.Value == W.HorizontalAnchorValues.Page
                ? PdfCore.PdfTableAnchor.Page : PdfCore.PdfTableAnchor.Margin;
            PdfCore.PdfTableAnchor verticalAnchor = position.VerticalAnchor?.Value == W.VerticalAnchorValues.Page
                ? PdfCore.PdfTableAnchor.Page : position.VerticalAnchor?.Value == W.VerticalAnchorValues.Margin
                    ? PdfCore.PdfTableAnchor.Margin : PdfCore.PdfTableAnchor.Flow;
            var horizontal = position.TablePositionXAlignment?.Value;
            PdfCore.PdfAlign alignment = horizontal == W.HorizontalAlignmentValues.Center ? PdfCore.PdfAlign.Center
                : horizontal == W.HorizontalAlignmentValues.Right || horizontal == W.HorizontalAlignmentValues.Outside
                    ? PdfCore.PdfAlign.Right : PdfCore.PdfAlign.Left;
            var vertical = position.TablePositionYAlignment?.Value;
            PdfCore.PdfTableVerticalAlignment verticalAlignment = vertical == W.VerticalAlignmentValues.Center
                ? PdfCore.PdfTableVerticalAlignment.Center : vertical == W.VerticalAlignmentValues.Bottom
                    ? PdfCore.PdfTableVerticalAlignment.Bottom : PdfCore.PdfTableVerticalAlignment.Top;
            return new PdfCore.PdfTablePosition(horizontalAnchor, verticalAnchor, alignment, verticalAlignment,
                horizontal.HasValue ? 0 : (position.TablePositionX?.Value ?? 0) / 20D,
                vertical.HasValue ? 0 : (position.TablePositionY?.Value ?? 0) / 20D,
                Math.Max(0, (position.LeftFromText?.Value ?? 0) / 20D),
                Math.Max(0, (position.RightFromText?.Value ?? 0) / 20D),
                Math.Max(0, (position.TopFromText?.Value ?? 0) / 20D),
                Math.Max(0, (position.BottomFromText?.Value ?? 0) / 20D));
        }
    }
}
