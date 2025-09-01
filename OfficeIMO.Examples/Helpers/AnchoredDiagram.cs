using System;
using System.Collections.Generic;
using OfficeIMO.Word;
using SixLabors.ImageSharp;

namespace OfficeIMO.Examples.Utils {
    internal static class AnchoredDiagram {
        private static int PtToEmus(double pt) => (int)Math.Round(pt * 12700.0);
        private static double PtToCm(double pt) => pt * 2.54 / 72.0;

        internal static void BuildGrid(
            WordDocument document,
            IList<(ShapeType type, double wPt, double hPt, Color fill, Color stroke, string label)> shapes,
            int cols = 5,
            double startXpt = 30,
            double startYpt = 80,
            double colStepPt = 110,
            double rowStepPt = 100,
            bool addLabels = true,
            bool addHorizontalConnectors = true,
            bool addVerticalConnectors = true,
            (int fromIndex, int toIndex)? elbowConnector = null,
            string? legend = null
        ) {
            var placed = new List<(double left, double top, double w, double h, string name)>();

            for (int i = 0; i < shapes.Count; i++) {
                int row = i / cols;
                int col = i % cols;
                double left = startXpt + col * colStepPt;
                double top = startYpt + row * rowStepPt;
                var (type, w, h, fill, stroke, label) = shapes[i];

                var p = document.AddParagraph("");
                var shp = p.AddShapeDrawing(type, w, h, left, top);
                shp.FillColor = fill;
                shp.StrokeColor = stroke;
                shp.StrokeWeight = 1.5;
                placed.Add((left, top, w, h, string.IsNullOrWhiteSpace(label) ? type.ToString() : label));
            }

            if (addLabels) {
                foreach (var item in placed) {
                    double labelWpt = Math.Max(70, item.w);
                    double labelHpt = 14;
                    double labelLeftPt = item.left + (item.w - labelWpt) / 2.0;
                    double labelTopPt = item.top - (labelHpt + 8);

                    var lp = document.AddParagraph("");
                    var tb = lp.AddTextBox("", WrapTextImage.InFrontOfText);
                    tb.WrapText = WrapTextImage.InFrontOfText;
                    tb.HorizontalPositionRelativeFrom = DocumentFormat.OpenXml.Drawing.Wordprocessing.HorizontalRelativePositionValues.Page;
                    tb.VerticalPositionRelativeFrom = DocumentFormat.OpenXml.Drawing.Wordprocessing.VerticalRelativePositionValues.Page;
                    tb.HorizontalPositionOffset = PtToEmus(labelLeftPt);
                    tb.VerticalPositionOffset = PtToEmus(labelTopPt);
                    tb.WidthCentimeters = PtToCm(labelWpt);
                    tb.HeightCentimeters = PtToCm(labelHpt);
                    var para = tb.Paragraphs.Count > 0 ? tb.Paragraphs[0] : document.AddParagraph("");
                    para.SetText(item.name)
                        .SetAlignment(DocumentFormat.OpenXml.Wordprocessing.JustificationValues.Center);
                }
            }

            if (addHorizontalConnectors) {
                for (int i = 0; i < placed.Count; i++) {
                    int col = i % cols;
                    if (col == cols - 1) continue;
                    var from = placed[i];
                    var to = placed[i + 1];
                    double gapLeft = from.left + from.w;
                    double available = to.left - gapLeft;
                    double arrowH = 12;
                    double arrowW = Math.Max(available - 8, 8);
                    double arrowTop = from.top + (from.h - arrowH) / 2.0;
                    var cp = document.AddParagraph("");
                    var conn = cp.AddShapeDrawing(ShapeType.RightArrow, arrowW, arrowH, gapLeft + 4, arrowTop);
                    conn.FillColor = Color.Gray;
                    conn.StrokeColor = Color.DimGray;
                    conn.StrokeWeight = 1;
                }
            }

            if (addVerticalConnectors) {
                for (int i = 0; i < placed.Count; i++) {
                    int nextRowIndex = i + cols;
                    if (nextRowIndex >= placed.Count) continue;
                    var from = placed[i];
                    var to = placed[nextRowIndex];
                    double gapTop = from.top + from.h;
                    double available = to.top - gapTop;
                    double arrowW = 12;
                    double arrowH = Math.Max(available - 8, 8);
                    double arrowLeft = from.left + (from.w - arrowW) / 2.0;
                    var vp = document.AddParagraph("");
                    var vconn = vp.AddShapeDrawing(ShapeType.DownArrow, arrowW, arrowH, arrowLeft, gapTop + 4);
                    vconn.FillColor = Color.Gray;
                    vconn.StrokeColor = Color.DimGray;
                    vconn.StrokeWeight = 1;
                }
            }

            if (legend != null) {
                double legendLeftPt = startXpt;
                double legendTopPt = Math.Max(30, startYpt - 40);
                double legendWpt = 340;
                double legendHpt = 36;
                var lp = document.AddParagraph("");
                var tb = lp.AddTextBox("", WrapTextImage.InFrontOfText);
                tb.WrapText = WrapTextImage.InFrontOfText;
                tb.HorizontalPositionRelativeFrom = DocumentFormat.OpenXml.Drawing.Wordprocessing.HorizontalRelativePositionValues.Page;
                tb.VerticalPositionRelativeFrom = DocumentFormat.OpenXml.Drawing.Wordprocessing.VerticalRelativePositionValues.Page;
                tb.HorizontalPositionOffset = PtToEmus(legendLeftPt);
                tb.VerticalPositionOffset = PtToEmus(legendTopPt);
                tb.WidthCentimeters = PtToCm(legendWpt);
                tb.HeightCentimeters = PtToCm(legendHpt);
                var para = tb.Paragraphs.Count > 0 ? tb.Paragraphs[0] : document.AddParagraph("");
                para.SetText(legend)
                    .SetAlignment(DocumentFormat.OpenXml.Wordprocessing.JustificationValues.Center);
            }

            if (elbowConnector != null) {
                var (fromIndex, toIndex) = elbowConnector.Value;
                if (fromIndex >= 0 && fromIndex < placed.Count && toIndex >= 0 && toIndex < placed.Count) {
                    var from = placed[fromIndex];
                    var target = placed[toIndex];
                    double elbowX = target.left + target.w / 2.0;
                    double elbowY = from.top + from.h / 2.0;
                    // Horizontal segment to elbowX
                    double hW = Math.Max((elbowX - (from.left + from.w)), 8);
                    double hH = 12;
                    var hp = document.AddParagraph("");
                    var hseg = hp.AddShapeDrawing(ShapeType.RightArrow, hW, hH, from.left + from.w + 4, elbowY - hH / 2.0);
                    hseg.FillColor = Color.Gray;
                    hseg.StrokeColor = Color.DimGray;
                    hseg.StrokeWeight = 1;
                    // Vertical segment down to target
                    double vH = Math.Max(((target.top + target.h / 2.0) - elbowY), 8);
                    double vW = 12;
                    var vp = document.AddParagraph("");
                    var vseg = vp.AddShapeDrawing(ShapeType.DownArrow, vW, vH, elbowX - vW / 2.0, elbowY + 2);
                    vseg.FillColor = Color.Gray;
                    vseg.StrokeColor = Color.DimGray;
                    vseg.StrokeWeight = 1;
                }
            }
        }
    }
}
