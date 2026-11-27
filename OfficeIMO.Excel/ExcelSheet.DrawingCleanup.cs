using OfficeIMO.Excel.OpenXml;
using OfficeIMO.OpenXml.Internal;
using static OfficeIMO.Excel.OpenXml.ExcelChartXml;

namespace OfficeIMO.Excel {
    public partial class ExcelSheet {
        internal void CleanupWorksheetDrawingArtifacts() {
            OfficeXmlNode drawing = ExcelPartRef.WorksheetDrawingElement(this);
            if (drawing.IsNone || drawing.RelationshipAttribute("id") is not string drawingRelId || string.IsNullOrWhiteSpace(drawingRelId)) {
                return;
            }

            ExcelPartRef worksheetPart = ExcelPartRef.WorksheetForMetadata(this);
            ExcelPartRef drawingPart = worksheetPart.GetPartById(drawingRelId);
            if (drawingPart.IsNone || drawingPart.Kind != ExcelPartKind.Drawings) {
                drawing.Remove();
                return;
            }

            OfficeXmlNode worksheetDrawing = drawingPart.Root;
            if (worksheetDrawing.IsNone) {
                worksheetPart.DeletePart(drawingPart);
                drawing.Remove();
                return;
            }

            foreach (OfficeXmlNode anchor in worksheetDrawing.Children().ToList()) {
                if (!IsSupportedDrawingAnchor(anchor)) {
                    continue;
                }

                if (!HasValidDrawingTarget(anchor, drawingPart)) {
                    anchor.Remove();
                }
            }

            var referencedRelationshipIds = new HashSet<string>(StringComparer.Ordinal);
            foreach (OfficeXmlNode anchor in worksheetDrawing.Children()) {
                CollectDrawingRelationshipIds(anchor, referencedRelationshipIds);
            }

            foreach (ExcelPartRef chartPart in drawingPart.PartsOfKind(ExcelPartKind.Chart).ToList()) {
                string? relId = drawingPart.GetIdOfPart(chartPart);
                if (relId == null || !referencedRelationshipIds.Contains(relId)) {
                    drawingPart.DeletePart(chartPart);
                }
            }

            foreach (ExcelPartRef imagePart in drawingPart.PartsOfKind(ExcelPartKind.Image).ToList()) {
                string? relId = drawingPart.GetIdOfPart(imagePart);
                if (relId == null || !referencedRelationshipIds.Contains(relId)) {
                    drawingPart.DeletePart(imagePart);
                }
            }

            if (worksheetDrawing.FirstChild.IsNone) {
                worksheetPart.DeletePart(drawingPart);
                drawing.Remove();
                return;
            }

            drawingPart.SaveRoot();
        }

        private static bool IsSupportedDrawingAnchor(OfficeXmlNode anchor)
            => anchor.Is(Xdr, "oneCellAnchor") || anchor.Is(Xdr, "twoCellAnchor") || anchor.Is(Xdr, "absoluteAnchor");

        private static bool HasValidDrawingTarget(OfficeXmlNode anchor, ExcelPartRef drawingPart) {
            OfficeXmlNode picture = anchor.FirstDescendantOutsideExtensions(Xdr, "pic");
            if (picture.Exists) {
                string? embed = picture.Child(Xdr, "blipFill").Child(A, "blip").RelationshipAttribute("embed");
                if (string.IsNullOrWhiteSpace(embed)) {
                    return false;
                }

                return drawingPart.GetPartById(embed).Kind == ExcelPartKind.Image;
            }

            OfficeXmlNode chartReference = anchor.FirstDescendantOutsideExtensions(C, "chart");
            if (chartReference.Exists) {
                string? relId = chartReference.RelationshipAttribute("id");
                if (string.IsNullOrWhiteSpace(relId)) {
                    return false;
                }

                return drawingPart.GetPartById(relId).Kind == ExcelPartKind.Chart;
            }

            return true;
        }

        private static void CollectDrawingRelationshipIds(OfficeXmlNode anchor, ISet<string> relationshipIds) {
            // Besides r:embed on the blip and chart, extensions such as svgBlip keep their own references.
            ExcelDrawingXml.CollectRelationshipAttributeValues(anchor, relationshipIds);
        }
    }
}
