using System;
using System.Collections.Generic;
using System.Linq;
using System.Xml.Linq;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using Threaded = DocumentFormat.OpenXml.Office2019.Excel.ThreadedComments;

namespace OfficeIMO.Excel {
    public partial class ExcelSheet {
        /// <summary>Removes overwritten comments that are not themselves part of the moved source.</summary>
        internal void RemoveRangeMoveDestinationComments(
            ExcelReference source,
            int firstRow,
            int firstColumn,
            int lastRow,
            int lastColumn) {
            bool legacyChanged = false;
            Comments? comments = _worksheetPart.WorksheetCommentsPart?.Comments;
            foreach (Comment comment in comments?.CommentList?.Elements<Comment>().ToList()
                ?? new List<Comment>()) {
                if (!TryGetDestinationOnlyCommentCell(
                        comment.Reference?.Value,
                        source,
                        firstRow,
                        firstColumn,
                        lastRow,
                        lastColumn,
                        out int row,
                        out int column)) continue;
                comment.Remove();
                RemoveCommentVmlShape(row, column);
                legacyChanged = true;
            }
            if (legacyChanged) comments!.Save();

            foreach (WorksheetThreadedCommentsPart part in _worksheetPart.WorksheetThreadedCommentsParts.ToList()) {
                Threaded.ThreadedComments? threaded = part.ThreadedComments;
                if (threaded == null) continue;
                bool changed = false;
                foreach (Threaded.ThreadedComment comment in threaded.Elements<Threaded.ThreadedComment>().ToList()) {
                    if (!TryGetDestinationOnlyCommentCell(
                            comment.Ref?.Value,
                            source,
                            firstRow,
                            firstColumn,
                            lastRow,
                            lastColumn,
                            out _,
                            out _)) continue;
                    comment.Remove();
                    changed = true;
                }
                if (!changed) continue;
                if (threaded.Elements<Threaded.ThreadedComment>().Any()) threaded.Save();
                else _worksheetPart.DeletePart(part);
            }

            CleanupCommentArtifacts();
        }

        /// <summary>Remaps legacy note shapes alongside comment references during range mutations.</summary>
        internal void RemapMutationCommentVml(Func<ExcelReference, ExcelReference?> transform) {
            CommentList? comments = _worksheetPart.WorksheetCommentsPart?.Comments?.CommentList;
            if (comments == null) return;

            var removed = new List<(int Row, int Col)>();
            var moved = new List<((int Row, int Col) OldCell, (int Row, int Col) NewCell)>();
            foreach (Comment comment in comments.Elements<Comment>()) {
                if (comment.Reference?.Value is not string value
                    || !ExcelReference.TryParse(value, out ExcelReference? reference)) continue;
                ExcelReference? mapped = transform(reference!);
                var oldCell = (reference!.Start.Row, reference.Start.Column);
                if (mapped == null) {
                    removed.Add(oldCell);
                } else if (mapped.Start.Row != oldCell.Row || mapped.Start.Column != oldCell.Column) {
                    moved.Add((oldCell, (mapped.Start.Row, mapped.Start.Column)));
                }
            }

            if (removed.Count > 0 || moved.Count > 0) {
                RemapCommentVmlShapes(removed, moved);
            }
        }

        private void ValidateRangeMoveCommentVmlAnchors(
            ExcelReference source,
            ExcelReference destination,
            MutationPlanScanBudget? budget) {
            VmlDrawingPart? vmlPart = TryGetCommentVmlPart();
            if (vmlPart == null) return;
            XDocument document = LoadOrCreateVmlDocument(vmlPart);
            XElement? root = document.Root;
            if (root == null) return;
            XNamespace v = XNamespace.Get("urn:schemas-microsoft-com:vml");
            XNamespace x = XNamespace.Get("urn:schemas-microsoft-com:office:excel");
            int rowDelta = destination.Start.Row - source.Start.Row;
            int columnDelta = destination.Start.Column - source.Start.Column;
            foreach (XElement shape in root.Elements(v + "shape")) {
                budget?.Consume();
                XElement? clientData = shape.Element(x + "ClientData");
                if (clientData == null
                    || GetVmlAnchorPlacement(clientData, x) == VmlAnchorPlacement.Absolute
                    || !TryParseVmlCoordinate(clientData.Element(x + "Row")?.Value, out int zeroBasedRow)
                    || !TryParseVmlCoordinate(clientData.Element(x + "Column")?.Value, out int zeroBasedColumn)
                    || !source.Contains(zeroBasedRow + 1, zeroBasedColumn + 1)
                    || !TryParseVmlAnchor(clientData.Element(x + "Anchor"), out int[] values)) {
                    continue;
                }
                if ((long)values[0] + columnDelta < 0
                    || (long)values[4] + columnDelta < 0
                    || (long)values[2] + rowDelta < 0
                    || (long)values[6] + rowDelta < 0
                    || (long)values[0] + columnDelta >= A1.MaxColumns
                    || (long)values[4] + columnDelta >= A1.MaxColumns
                    || (long)values[2] + rowDelta >= A1.MaxRows
                    || (long)values[6] + rowDelta >= A1.MaxRows) {
                    throw new InvalidOperationException(
                        "Range move would place a legacy comment VML anchor outside worksheet limits.");
                }
            }
        }

        private static bool TryGetDestinationOnlyCommentCell(
            string? referenceText,
            ExcelReference source,
            int firstRow,
            int firstColumn,
            int lastRow,
            int lastColumn,
            out int row,
            out int column) {
            row = column = 0;
            if (!ExcelReference.TryParse(referenceText, out ExcelReference? reference)) return false;
            row = reference!.Start.Row;
            column = reference.Start.Column;
            return row >= firstRow
                && row <= lastRow
                && column >= firstColumn
                && column <= lastColumn
                && !source.Contains(row, column);
        }
    }
}
