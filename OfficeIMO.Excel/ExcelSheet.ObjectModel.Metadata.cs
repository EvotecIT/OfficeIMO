using System.Globalization;
using System.Text;
using System.Text.RegularExpressions;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Spreadsheet;
using Xdr = DocumentFormat.OpenXml.Drawing.Spreadsheet;

namespace OfficeIMO.Excel {
    public partial class ExcelSheet {
        private bool ClearCommentsInRange(int firstRow, int firstColumn, int lastRow, int lastColumn) {
            bool changed = false;
            var commentsPart = WorksheetCommentsPartRoot;
            if (commentsPart?.Comments?.CommentList != null) {
                bool removedComment = false;
                var commentList = commentsPart.Comments.CommentList;
                if (CommentListOverlapsRange(commentList, firstRow, firstColumn, lastRow, lastColumn)) {
                    foreach (var comment in commentList.Elements<Comment>().ToList()) {
                        if (comment.Reference?.Value is not string reference) {
                            continue;
                        }

                        var (row, col) = A1.ParseCellRef(reference);
                        if (row >= firstRow && row <= lastRow && col >= firstColumn && col <= lastColumn) {
                            comment.Remove();
                            removedComment = true;
                        }
                    }
                }

                if (removedComment) {
                    commentsPart.Comments.Save();
                    changed = true;
                }
            }

            changed |= RemoveCommentVmlShapesInRange(firstRow, firstColumn, lastRow, lastColumn);
            changed |= CleanupCommentArtifacts();
            return changed;
        }

        private static bool CommentListOverlapsRange(CommentList commentList, int firstRow, int firstColumn, int lastRow, int lastColumn) {
            foreach (var comment in commentList.Elements<Comment>()) {
                if (comment.Reference?.Value is not string reference) {
                    continue;
                }

                var (row, col) = A1.ParseCellRef(reference);
                if (row >= firstRow && row <= lastRow && col >= firstColumn && col <= lastColumn) {
                    return true;
                }
            }

            return false;
        }


        private void RewriteWorksheetFormulaReferences(int firstAffectedRow, int rowDelta) {
            foreach (var cell in WorksheetRoot.Descendants<Cell>()) {
                if (cell.CellFormula?.Text is string formulaText && formulaText.Length > 0) {
                    cell.CellFormula.Text = RewriteShiftedFormulaReferences(formulaText, firstAffectedRow, rowDelta, Name);
                }
            }
        }

        private void RewriteDeletedWorksheetFormulaReferences(int firstDeletedRow, int lastDeletedRow, int rowDelta) {
            foreach (var cell in WorksheetRoot.Descendants<Cell>()) {
                if (cell.CellFormula?.Text is string formulaText && formulaText.Length > 0) {
                    cell.CellFormula.Text = RewriteDeletedFormulaReferences(formulaText, firstDeletedRow, lastDeletedRow, rowDelta, Name);
                }
            }
        }

        private void RemapShiftedRowMetadata(int firstAffectedRow, int rowDelta) {
            RemapShiftedDefinedNames(firstAffectedRow, rowDelta, lastDeletedRow: null);
            RemapShiftedTables(firstAffectedRow, rowDelta, lastDeletedRow: null);
            RemapShiftedComments(firstAffectedRow, rowDelta, lastDeletedRow: null);
            RemapShiftedHyperlinks(firstAffectedRow, rowDelta, lastDeletedRow: null);
            RemapShiftedDataValidations(firstAffectedRow, rowDelta, lastDeletedRow: null);
            RemapShiftedConditionalFormatting(firstAffectedRow, rowDelta, lastDeletedRow: null);
            RemapShiftedSparklines(firstAffectedRow, rowDelta, lastDeletedRow: null);
            RemapShiftedDrawingAnchors(firstAffectedRow, rowDelta, lastDeletedRow: null);
            RemapShiftedChartReferences(firstAffectedRow, rowDelta, lastDeletedRow: null);
        }

        private void RemapDeletedRowMetadata(int firstDeletedRow, int lastDeletedRow, int rowDelta) {
            RemapShiftedDefinedNames(firstDeletedRow, rowDelta, lastDeletedRow);
            RemapShiftedTables(firstDeletedRow, rowDelta, lastDeletedRow);
            RemapShiftedComments(firstDeletedRow, rowDelta, lastDeletedRow);
            RemapShiftedHyperlinks(firstDeletedRow, rowDelta, lastDeletedRow);
            RemapShiftedDataValidations(firstDeletedRow, rowDelta, lastDeletedRow);
            RemapShiftedConditionalFormatting(firstDeletedRow, rowDelta, lastDeletedRow);
            RemapShiftedSparklines(firstDeletedRow, rowDelta, lastDeletedRow);
            RemapShiftedDrawingAnchors(firstDeletedRow, rowDelta, lastDeletedRow);
            RemapShiftedChartReferences(firstDeletedRow, rowDelta, lastDeletedRow);
        }

        private void RemapShiftedDefinedNames(int firstAffectedRow, int rowDelta, int? lastDeletedRow) {
            var definedNames = WorkbookRoot.DefinedNames;
            if (definedNames == null) {
                return;
            }

            bool changed = false;
            foreach (var definedName in definedNames.Elements<DefinedName>()) {
                string? text = definedName.Text;
                if (string.IsNullOrWhiteSpace(text)) {
                    continue;
                }

                string rewritten = lastDeletedRow.HasValue
                    ? RewriteDeletedFormulaReferences(text, firstAffectedRow, lastDeletedRow.Value, rowDelta, Name)
                    : RewriteShiftedFormulaReferences(text, firstAffectedRow, rowDelta, Name);

                if (!string.Equals(text, rewritten, StringComparison.Ordinal)) {
                    definedName.Text = rewritten;
                    changed = true;
                }
            }

            if (changed) {
                WorkbookRoot.Save();
            }
        }

        private void RemapShiftedTables(int firstAffectedRow, int rowDelta, int? lastDeletedRow) {
            foreach (var tableDefinitionPart in _worksheetPart.TableDefinitionParts) {
                var table = tableDefinitionPart.Table;
                if (table == null) {
                    continue;
                }

                bool changed = false;
                if (table.Reference?.Value is string reference
                    && TryRemapShiftedReferenceListRows(reference, firstAffectedRow, rowDelta, lastDeletedRow, out var remapped)
                    && remapped.Count > 0) {
                    string updatedReference = remapped[0];
                    if (!string.Equals(reference, updatedReference, StringComparison.OrdinalIgnoreCase)) {
                        table.Reference = updatedReference;
                        changed = true;
                    }
                }

                var autoFilter = table.GetFirstChild<AutoFilter>();
                if (autoFilter?.Reference?.Value is string filterReference
                    && TryRemapShiftedReferenceListRows(filterReference, firstAffectedRow, rowDelta, lastDeletedRow, out var remappedFilter)
                    && remappedFilter.Count > 0) {
                    string updatedFilterReference = remappedFilter[0];
                    if (!string.Equals(filterReference, updatedFilterReference, StringComparison.OrdinalIgnoreCase)) {
                        autoFilter.Reference = updatedFilterReference;
                        changed = true;
                    }
                }

                if (changed) {
                    table.Save();
                }
            }
        }

        private void RemapShiftedComments(int firstAffectedRow, int rowDelta, int? lastDeletedRow) {
            var commentsPart = WorksheetCommentsPartRoot;
            if (commentsPart?.Comments?.CommentList == null) {
                return;
            }

            var removed = new List<(int Row, int Col)>();
            var moved = new List<((int Row, int Col) OldCell, (int Row, int Col) NewCell)>();
            bool changed = false;
            foreach (var comment in commentsPart.Comments.CommentList.Elements<Comment>().ToList()) {
                if (comment.Reference?.Value is not string reference) {
                    continue;
                }

                var cell = A1.ParseCellRef(reference);
                if (!TryRemapShiftedReferenceRows((cell.Row, cell.Col, cell.Row, cell.Col), firstAffectedRow, rowDelta, lastDeletedRow, out var remapped)) {
                    continue;
                }

                if (remapped == null) {
                    comment.Remove();
                    removed.Add(cell);
                    changed = true;
                    continue;
                }

                string newReference = A1.CellReference(remapped.Value.r1, remapped.Value.c1);
                if (!string.Equals(reference, newReference, StringComparison.OrdinalIgnoreCase)) {
                    comment.Reference = newReference;
                    moved.Add((cell, (remapped.Value.r1, remapped.Value.c1)));
                    changed = true;
                }
            }

            if (!changed) {
                return;
            }

            commentsPart.Comments.Save();
            var shapesToRemove = new HashSet<(int Row, int Col)>();
            foreach (var cell in removed) {
                shapesToRemove.Add(cell);
            }

            foreach (var pair in moved) {
                shapesToRemove.Add(pair.OldCell);
            }

            foreach (var cell in shapesToRemove) {
                RemoveCommentVmlShape(cell.Row, cell.Col);
            }

            foreach (var pair in moved) {
                EnsureCommentVmlShape(pair.NewCell.Row, pair.NewCell.Col);
            }

            CleanupCommentArtifacts();
        }

        private void RemapShiftedHyperlinks(int firstAffectedRow, int rowDelta, int? lastDeletedRow) {
            var hyperlinks = WorksheetRoot.GetFirstChild<Hyperlinks>();
            if (hyperlinks == null) {
                return;
            }

            foreach (var link in hyperlinks.Elements<Hyperlink>().ToList()) {
                if (link.Reference?.Value is not string reference
                    || !TryRemapShiftedReferenceListRows(reference, firstAffectedRow, rowDelta, lastDeletedRow, out var remapped)) {
                    continue;
                }

                if (remapped.Count == 0) {
                    link.Remove();
                    continue;
                }

                link.Reference = remapped[0];
                var insertAfter = link;
                for (int index = 1; index < remapped.Count; index++) {
                    var clone = (Hyperlink)link.CloneNode(true);
                    clone.Reference = remapped[index];
                    hyperlinks.InsertAfter(clone, insertAfter);
                    insertAfter = clone;
                }
            }
        }

        private void RemapShiftedDataValidations(int firstAffectedRow, int rowDelta, int? lastDeletedRow) {
            var validations = WorksheetRoot.GetFirstChild<DataValidations>();
            if (validations == null) {
                return;
            }

            uint count = 0;
            foreach (var validation in validations.Elements<DataValidation>().ToList()) {
                if (validation.SequenceOfReferences?.InnerText is not string references
                    || !TryRemapShiftedReferenceListRows(references, firstAffectedRow, rowDelta, lastDeletedRow, out var remapped)) {
                    count++;
                    continue;
                }

                if (remapped.Count == 0) {
                    validation.Remove();
                    continue;
                }

                validation.SequenceOfReferences = new ListValue<StringValue> { InnerText = string.Join(" ", remapped) };
                count++;
            }

            validations.Count = count;
        }

        private void RemapShiftedConditionalFormatting(int firstAffectedRow, int rowDelta, int? lastDeletedRow) {
            foreach (var conditional in WorksheetRoot.Elements<ConditionalFormatting>().ToList()) {
                if (conditional.SequenceOfReferences?.InnerText is not string references
                    || !TryRemapShiftedReferenceListRows(references, firstAffectedRow, rowDelta, lastDeletedRow, out var remapped)) {
                    continue;
                }

                if (remapped.Count == 0) {
                    conditional.Remove();
                    continue;
                }

                conditional.SequenceOfReferences = new ListValue<StringValue> { InnerText = string.Join(" ", remapped) };
            }
        }

        private void RemapShiftedSparklines(int firstAffectedRow, int rowDelta, int? lastDeletedRow) {
            foreach (var sparkline in WorksheetRoot.Descendants<DocumentFormat.OpenXml.Office2010.Excel.Sparkline>().ToList()) {
                if (sparkline.ReferenceSequence?.Text is string location
                    && TryRemapShiftedReferenceListRows(location, firstAffectedRow, rowDelta, lastDeletedRow, out var remappedLocations)) {
                    if (remappedLocations.Count == 0) {
                        sparkline.Remove();
                        continue;
                    }

                    sparkline.ReferenceSequence.Text = string.Join(" ", remappedLocations);
                }

                if (sparkline.Formula?.Text is string formula
                    && TryRemapShiftedReferenceListRows(formula, firstAffectedRow, rowDelta, lastDeletedRow, out var remappedFormulas)) {
                    if (remappedFormulas.Count == 0) {
                        sparkline.Remove();
                        continue;
                    }

                    sparkline.Formula.Text = string.Join(" ", remappedFormulas);
                }
            }
        }

        private void RemapShiftedDrawingAnchors(int firstAffectedRow, int rowDelta, int? lastDeletedRow) {
            var drawing = _worksheetPart.DrawingsPart?.WorksheetDrawing;
            if (drawing == null) {
                return;
            }

            bool changed = false;
            foreach (var anchor in drawing.ChildElements.ToList()) {
                if (anchor is Xdr.OneCellAnchor oneCellAnchor) {
                    if (!TryRemapDrawingMarkerRow(oneCellAnchor.FromMarker, firstAffectedRow, rowDelta, lastDeletedRow, out bool markerChanged)) {
                        oneCellAnchor.Remove();
                        changed = true;
                        continue;
                    }

                    changed |= markerChanged;
                } else if (anchor is Xdr.TwoCellAnchor twoCellAnchor) {
                    Xdr.EditAsValues placement = twoCellAnchor.EditAs?.Value ?? Xdr.EditAsValues.TwoCell;
                    if (placement == Xdr.EditAsValues.Absolute) {
                        continue;
                    }

                    bool fromKept = TryRemapDrawingMarkerRow(twoCellAnchor.FromMarker, firstAffectedRow, rowDelta, lastDeletedRow, out bool fromChanged);
                    if (!fromKept) {
                        twoCellAnchor.Remove();
                        changed = true;
                        continue;
                    }

                    if (placement == Xdr.EditAsValues.OneCell) {
                        changed |= fromChanged;
                        if (fromChanged) {
                            if (!TryShiftDrawingMarkerRow(twoCellAnchor.ToMarker, rowDelta, out bool toShifted)) {
                                twoCellAnchor.Remove();
                                changed = true;
                                continue;
                            }

                            changed |= toShifted;
                        }

                        continue;
                    }

                    bool toKept = TryRemapDrawingMarkerRow(twoCellAnchor.ToMarker, firstAffectedRow, rowDelta, lastDeletedRow, out bool toChanged);
                    if (!toKept) {
                        twoCellAnchor.Remove();
                        changed = true;
                        continue;
                    }

                    changed |= fromChanged || toChanged;
                }
            }

            if (changed) {
                drawing.Save();
            }
        }

        private static bool TryRemapDrawingMarkerRow(Xdr.MarkerType? marker, int firstAffectedRow, int rowDelta, int? lastDeletedRow, out bool changed) {
            changed = false;
            if (marker?.RowId?.Text is not string rowText
                || !int.TryParse(rowText, NumberStyles.Integer, CultureInfo.InvariantCulture, out int zeroBasedRow)) {
                return true;
            }

            int oneBasedRow = zeroBasedRow + 1;
            if (!TryRemapShiftedReferenceRows((oneBasedRow, 1, oneBasedRow, 1), firstAffectedRow, rowDelta, lastDeletedRow, out var remapped)) {
                return true;
            }

            if (remapped == null) {
                return false;
            }

            int remappedZeroBasedRow = remapped.Value.r1 - 1;
            if (remappedZeroBasedRow != zeroBasedRow) {
                marker.RowId.Text = remappedZeroBasedRow.ToString(CultureInfo.InvariantCulture);
                changed = true;
            }

            return true;
        }

        private static bool TryShiftDrawingMarkerRow(Xdr.MarkerType? marker, int rowDelta, out bool changed) {
            changed = false;
            if (rowDelta == 0
                || marker?.RowId?.Text is not string rowText
                || !int.TryParse(rowText, NumberStyles.Integer, CultureInfo.InvariantCulture, out int zeroBasedRow)) {
                return true;
            }

            int shiftedRow = zeroBasedRow + rowDelta;
            if (shiftedRow < 0 || shiftedRow >= A1.MaxRows) {
                return false;
            }

            if (shiftedRow != zeroBasedRow) {
                marker.RowId.Text = shiftedRow.ToString(CultureInfo.InvariantCulture);
                changed = true;
            }

            return true;
        }

        private void RemapShiftedChartReferences(int firstAffectedRow, int rowDelta, int? lastDeletedRow) {
            var drawingPart = _worksheetPart.DrawingsPart;
            if (drawingPart == null) {
                return;
            }

            foreach (var chartPart in drawingPart.ChartParts) {
                var chartSpace = chartPart.ChartSpace;
                if (chartSpace == null) {
                    continue;
                }

                bool changed = false;
                foreach (var formula in chartSpace.Descendants<DocumentFormat.OpenXml.Drawing.Charts.Formula>()) {
                    string? text = formula.Text;
                    if (string.IsNullOrEmpty(text)) {
                        continue;
                    }

                    string rewritten = lastDeletedRow.HasValue
                        ? RewriteDeletedFormulaReferences(text, firstAffectedRow, lastDeletedRow.Value, rowDelta, Name)
                        : RewriteShiftedFormulaReferences(text, firstAffectedRow, rowDelta, Name);
                    if (!string.Equals(text, rewritten, StringComparison.Ordinal)) {
                        formula.Text = rewritten;
                        changed = true;
                    }
                }

                if (changed) {
                    chartSpace.Save();
                }
            }
        }


        private bool ClearHyperlinksInRange(Worksheet ws, (int r1, int c1, int r2, int c2) bounds) {
            var hyperlinks = ws.GetFirstChild<Hyperlinks>();
            if (hyperlinks == null) return false;
            if (!HyperlinksOverlapRange(hyperlinks, bounds)) return false;

            bool changed = false;
            foreach (var link in hyperlinks.Elements<Hyperlink>().ToList()) {
                if (link.Reference?.Value is string reference) {
                    if (!TryRemoveReferenceOverlap(reference, bounds, out var remaining)) {
                        continue;
                    }

                    if (remaining.Count == 0) {
                        link.Remove();
                        changed = true;
                        continue;
                    }

                    link.Reference = remaining[0];
                    var insertAfter = link;
                    for (int index = 1; index < remaining.Count; index++) {
                        var clone = (Hyperlink)link.CloneNode(true);
                        clone.Reference = remaining[index];
                        hyperlinks.InsertAfter(clone, insertAfter);
                        insertAfter = clone;
                    }

                    changed = true;
                }
            }

            return changed;
        }

        private static bool HyperlinksOverlapRange(Hyperlinks hyperlinks, (int r1, int c1, int r2, int c2) bounds) {
            foreach (var link in hyperlinks.Elements<Hyperlink>()) {
                if (link.Reference?.Value is string reference && ReferenceListOverlaps(reference, bounds)) {
                    return true;
                }
            }

            return false;
        }

        private bool ClearSparklinesInRange((int r1, int c1, int r2, int c2) bounds) {
            if (!SparklinesOverlap(bounds)) return false;

            bool changed = false;
            foreach (var sparkline in WorksheetRoot.Descendants<DocumentFormat.OpenXml.Office2010.Excel.Sparkline>().ToList()) {
                var reference = sparkline.ReferenceSequence?.Text;
                if (!string.IsNullOrWhiteSpace(reference) && TryParseReference(reference!, out var sparklineBounds)) {
                    if (RangesOverlapInclusive(bounds, sparklineBounds)) {
                        sparkline.Remove();
                        changed = true;
                    }
                }
            }

            return changed;
        }

        private bool SparklinesOverlap((int r1, int c1, int r2, int c2) bounds) {
            foreach (var sparkline in WorksheetRoot.Descendants<DocumentFormat.OpenXml.Office2010.Excel.Sparkline>()) {
                var reference = sparkline.ReferenceSequence?.Text;
                if (!string.IsNullOrWhiteSpace(reference)
                    && TryParseReference(reference!, out var sparklineBounds)
                    && RangesOverlapInclusive(bounds, sparklineBounds)) {
                    return true;
                }
            }

            return false;
        }
    }
}
