using PdfCore = OfficeIMO.Pdf;

namespace OfficeIMO.Word.Pdf {

    public static partial class WordPdfConverterExtensions {
        private static PdfCore.PdfTablePosition CreateNativeTablePosition(OfficeXmlNode position, bool allowOverlap) {
            string? verticalAnchorToken = position.WAttr("vertAnchor");
            PdfCore.PdfTableAnchor horizontalAnchor = position.WAttr("horzAnchor") == "page"
                ? PdfCore.PdfTableAnchor.Page : PdfCore.PdfTableAnchor.Margin;
            PdfCore.PdfTableAnchor verticalAnchor = verticalAnchorToken == "page"
                ? PdfCore.PdfTableAnchor.Page : verticalAnchorToken == "margin"
                    ? PdfCore.PdfTableAnchor.Margin : PdfCore.PdfTableAnchor.Flow;
            string? horizontal = position.WAttr("tblpXSpec");
            global::OfficeIMO.OfficeAlignment alignment = horizontal == "center" ? global::OfficeIMO.OfficeAlignment.Center
                : horizontal == "right" || horizontal == "outside"
                    ? global::OfficeIMO.OfficeAlignment.Right : global::OfficeIMO.OfficeAlignment.Left;
            string? vertical = position.WAttr("tblpYSpec");
            bool relativeVertical = vertical != null && verticalAnchor != PdfCore.PdfTableAnchor.Flow;
            if (!relativeVertical) vertical = null;
            PdfCore.PdfTableVerticalAlignment verticalAlignment = vertical == "center"
                ? PdfCore.PdfTableVerticalAlignment.Center : vertical == "bottom"
                    ? PdfCore.PdfTableVerticalAlignment.Bottom : vertical == "inside"
                    ? PdfCore.PdfTableVerticalAlignment.Inside : vertical == "outside"
                        ? PdfCore.PdfTableVerticalAlignment.Outside : PdfCore.PdfTableVerticalAlignment.Top;
            return new PdfCore.PdfTablePosition(horizontalAnchor, verticalAnchor, alignment, verticalAlignment,
                global::OfficeIMO.OfficeLength.Points(horizontal != null ? 0 : (position.WInt32("tblpX") ?? 0) / 20D),
                global::OfficeIMO.OfficeLength.Points(vertical != null ? 0 : (position.WInt32("tblpY") ?? 0) / 20D),
                global::OfficeIMO.OfficeLength.Points(Math.Max(0, (position.WInt32("leftFromText") ?? 0) / 20D)),
                global::OfficeIMO.OfficeLength.Points(Math.Max(0, (position.WInt32("rightFromText") ?? 0) / 20D)),
                global::OfficeIMO.OfficeLength.Points(Math.Max(0, (position.WInt32("topFromText") ?? 0) / 20D)),
                global::OfficeIMO.OfficeLength.Points(Math.Max(0, (position.WInt32("bottomFromText") ?? 0) / 20D)),
                mirrorHorizontalOnEvenPages: horizontal == "inside" || horizontal == "outside",
                allowOverlap: allowOverlap) { UsesWordOutsideEdges = true };
        }
    }
}
