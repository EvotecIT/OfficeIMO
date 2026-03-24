using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using A = DocumentFormat.OpenXml.Drawing;
using C = DocumentFormat.OpenXml.Drawing.Charts;
using Xdr = DocumentFormat.OpenXml.Drawing.Spreadsheet;

namespace OfficeIMO.Excel {
    public partial class ExcelSheet {
        internal void CleanupWorksheetDrawingArtifacts() {
            var ws = WorksheetRoot;
            var drawing = ws.GetFirstChild<Drawing>();
            if (drawing?.Id?.Value is not string drawingRelId || string.IsNullOrWhiteSpace(drawingRelId)) {
                return;
            }

            DrawingsPart drawingPart;
            try {
                drawingPart = (DrawingsPart)_worksheetPart.GetPartById(drawingRelId);
            } catch {
                ws.RemoveChild(drawing);
                return;
            }

            var worksheetDrawing = drawingPart.WorksheetDrawing;
            if (worksheetDrawing == null) {
                _worksheetPart.DeletePart(drawingPart);
                ws.RemoveChild(drawing);
                return;
            }

            foreach (var anchor in worksheetDrawing.ChildElements.ToList()) {
                if (!IsSupportedDrawingAnchor(anchor)) {
                    continue;
                }

                if (!HasValidDrawingTarget(anchor, drawingPart)) {
                    anchor.Remove();
                }
            }

            var referencedRelationshipIds = new HashSet<string>(StringComparer.Ordinal);
            foreach (var anchor in worksheetDrawing.ChildElements) {
                CollectDrawingRelationshipIds(anchor, referencedRelationshipIds);
            }

            foreach (var chartPart in drawingPart.ChartParts.ToList()) {
                string relId = drawingPart.GetIdOfPart(chartPart);
                if (!referencedRelationshipIds.Contains(relId)) {
                    drawingPart.DeletePart(chartPart);
                }
            }

            foreach (var imagePart in drawingPart.ImageParts.ToList()) {
                string relId = drawingPart.GetIdOfPart(imagePart);
                if (!referencedRelationshipIds.Contains(relId)) {
                    drawingPart.DeletePart(imagePart);
                }
            }

            if (!worksheetDrawing.ChildElements.Any()) {
                _worksheetPart.DeletePart(drawingPart);
                ws.RemoveChild(drawing);
                return;
            }

            worksheetDrawing.Save();
        }

        private static bool IsSupportedDrawingAnchor(OpenXmlElement anchor)
            => anchor is Xdr.OneCellAnchor || anchor is Xdr.TwoCellAnchor || anchor is Xdr.AbsoluteAnchor;

        private static bool HasValidDrawingTarget(OpenXmlElement anchor, DrawingsPart drawingPart) {
            var picture = anchor.Descendants<Xdr.Picture>().FirstOrDefault();
            if (picture != null) {
                string? embed = picture.BlipFill?.Blip?.Embed?.Value;
                if (string.IsNullOrWhiteSpace(embed)) {
                    return false;
                }

                try {
                    return drawingPart.GetPartById(embed!) is ImagePart;
                } catch {
                    return false;
                }
            }

            var chartReference = anchor.Descendants<C.ChartReference>().FirstOrDefault();
            if (chartReference != null) {
                string? relId = chartReference.Id?.Value;
                if (string.IsNullOrWhiteSpace(relId)) {
                    return false;
                }

                try {
                    return drawingPart.GetPartById(relId!) is ChartPart;
                } catch {
                    return false;
                }
            }

            return true;
        }

        private static void CollectDrawingRelationshipIds(OpenXmlElement anchor, ISet<string> relationshipIds) {
            foreach (var blip in anchor.Descendants<A.Blip>()) {
                if (!string.IsNullOrWhiteSpace(blip.Embed?.Value)) {
                    relationshipIds.Add(blip.Embed!.Value!);
                }
            }

            foreach (var chartReference in anchor.Descendants<C.ChartReference>()) {
                if (!string.IsNullOrWhiteSpace(chartReference.Id?.Value)) {
                    relationshipIds.Add(chartReference.Id!.Value!);
                }
            }
        }
    }
}
