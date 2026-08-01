using System.Globalization;
using System.Threading;
using System.Text;
using System.Text.RegularExpressions;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using Xdr = DocumentFormat.OpenXml.Drawing.Spreadsheet;
using X14 = DocumentFormat.OpenXml.Office2010.Excel;
using Threaded = DocumentFormat.OpenXml.Office2019.Excel.ThreadedComments;

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
        private void RemapShiftedRowMetadata(int firstAffectedRow, int rowDelta, CancellationToken cancellationToken = default) {
            cancellationToken.ThrowIfCancellationRequested();
            RemapShiftedDataConsolidationReferences(firstAffectedRow, rowDelta, lastDeletedRow: null);
            cancellationToken.ThrowIfCancellationRequested();
            RemapShiftedWebPublishItems(firstAffectedRow, rowDelta, lastDeletedRow: null);
            cancellationToken.ThrowIfCancellationRequested();
            RemapShiftedConnectionParameters(firstAffectedRow, rowDelta, lastDeletedRow: null, cancellationToken);
            cancellationToken.ThrowIfCancellationRequested();
            RemapShiftedPivotSources(firstAffectedRow, rowDelta, lastDeletedRow: null);
            cancellationToken.ThrowIfCancellationRequested();
            RemapShiftedDefinedNames(firstAffectedRow, rowDelta, lastDeletedRow: null);
            cancellationToken.ThrowIfCancellationRequested();
            RemapShiftedTables(firstAffectedRow, rowDelta, lastDeletedRow: null);
            cancellationToken.ThrowIfCancellationRequested();
            RemapShiftedWorksheetRangeMetadata(firstAffectedRow, rowDelta, lastDeletedRow: null);
            cancellationToken.ThrowIfCancellationRequested();
            RemapShiftedPivotLocations(firstAffectedRow, rowDelta, lastDeletedRow: null);
            cancellationToken.ThrowIfCancellationRequested();
            RemapShiftedComments(firstAffectedRow, rowDelta, lastDeletedRow: null);
            cancellationToken.ThrowIfCancellationRequested();
            RemapShiftedThreadedComments(firstAffectedRow, rowDelta, lastDeletedRow: null);
            cancellationToken.ThrowIfCancellationRequested();
            RemapShiftedHyperlinks(firstAffectedRow, rowDelta, lastDeletedRow: null);
            cancellationToken.ThrowIfCancellationRequested();
            RemapShiftedDataValidations(firstAffectedRow, rowDelta, lastDeletedRow: null);
            cancellationToken.ThrowIfCancellationRequested();
            RemapShiftedConditionalFormatting(firstAffectedRow, rowDelta, lastDeletedRow: null);
            cancellationToken.ThrowIfCancellationRequested();
            RemapShiftedSparklines(firstAffectedRow, rowDelta, lastDeletedRow: null);
            cancellationToken.ThrowIfCancellationRequested();
            RemapShiftedDrawingAnchors(firstAffectedRow, rowDelta, lastDeletedRow: null);
            cancellationToken.ThrowIfCancellationRequested();
            RemapShiftedChartReferences(firstAffectedRow, rowDelta, lastDeletedRow: null);
            cancellationToken.ThrowIfCancellationRequested();
            InvalidateWorkbookChartCaches();
        }

        private void RemapDeletedRowMetadata(int firstDeletedRow, int lastDeletedRow, int rowDelta, CancellationToken cancellationToken = default) {
            cancellationToken.ThrowIfCancellationRequested();
            RemapShiftedDataConsolidationReferences(firstDeletedRow, rowDelta, lastDeletedRow);
            cancellationToken.ThrowIfCancellationRequested();
            RemapShiftedWebPublishItems(firstDeletedRow, rowDelta, lastDeletedRow);
            cancellationToken.ThrowIfCancellationRequested();
            RemapShiftedConnectionParameters(firstDeletedRow, rowDelta, lastDeletedRow, cancellationToken);
            cancellationToken.ThrowIfCancellationRequested();
            RemapShiftedPivotSources(firstDeletedRow, rowDelta, lastDeletedRow);
            cancellationToken.ThrowIfCancellationRequested();
            RemapShiftedDefinedNames(firstDeletedRow, rowDelta, lastDeletedRow);
            cancellationToken.ThrowIfCancellationRequested();
            RemapShiftedTables(firstDeletedRow, rowDelta, lastDeletedRow);
            cancellationToken.ThrowIfCancellationRequested();
            RemapShiftedWorksheetRangeMetadata(firstDeletedRow, rowDelta, lastDeletedRow);
            cancellationToken.ThrowIfCancellationRequested();
            RemapShiftedPivotLocations(firstDeletedRow, rowDelta, lastDeletedRow);
            cancellationToken.ThrowIfCancellationRequested();
            RemapShiftedComments(firstDeletedRow, rowDelta, lastDeletedRow);
            cancellationToken.ThrowIfCancellationRequested();
            RemapShiftedThreadedComments(firstDeletedRow, rowDelta, lastDeletedRow);
            cancellationToken.ThrowIfCancellationRequested();
            RemapShiftedHyperlinks(firstDeletedRow, rowDelta, lastDeletedRow);
            cancellationToken.ThrowIfCancellationRequested();
            RemapShiftedDataValidations(firstDeletedRow, rowDelta, lastDeletedRow);
            cancellationToken.ThrowIfCancellationRequested();
            RemapShiftedConditionalFormatting(firstDeletedRow, rowDelta, lastDeletedRow);
            cancellationToken.ThrowIfCancellationRequested();
            RemapShiftedSparklines(firstDeletedRow, rowDelta, lastDeletedRow);
            cancellationToken.ThrowIfCancellationRequested();
            RemapShiftedDrawingAnchors(firstDeletedRow, rowDelta, lastDeletedRow);
            cancellationToken.ThrowIfCancellationRequested();
            RemapShiftedChartReferences(firstDeletedRow, rowDelta, lastDeletedRow);
            cancellationToken.ThrowIfCancellationRequested();
            InvalidateWorkbookChartCaches();
        }

        private bool RemapShiftedDefinedNames(int firstAffectedRow, int rowDelta, int? lastDeletedRow) {
            var definedNames = WorkbookRoot.DefinedNames;
            if (definedNames == null) {
                return false;
            }

            List<Sheet> workbookSheets = WorkbookRoot.Sheets?.Elements<Sheet>().ToList() ?? new List<Sheet>();
            int mutatedSheetIndex = workbookSheets.FindIndex(sheet =>
                string.Equals(sheet.Name?.Value, Name, StringComparison.OrdinalIgnoreCase));
            bool changed = false;
            foreach (var definedName in definedNames.Elements<DefinedName>()) {
                string? text = definedName.Text;
                if (string.IsNullOrWhiteSpace(text)) {
                    continue;
                }

                bool rewriteUnqualifiedReferences = mutatedSheetIndex >= 0
                    && definedName.LocalSheetId?.Value == (uint)mutatedSheetIndex;
                string rewritten = lastDeletedRow.HasValue
                    ? RewriteDeletedFormulaReferences(
                        text,
                        firstAffectedRow,
                        lastDeletedRow.Value,
                        rowDelta,
                        Name,
                        rewriteUnqualifiedReferences)
                    : RewriteShiftedFormulaReferences(
                        text,
                        firstAffectedRow,
                        rowDelta,
                        Name,
                        rewriteUnqualifiedReferences);

                if (!string.Equals(text, rewritten, StringComparison.Ordinal)) {
                    definedName.Text = rewritten;
                    changed = true;
                }
            }

            if (changed) {
                WorkbookRoot.Save();
            }
            return changed;
        }

        private void RemapShiftedConnectionParameters(
            int firstAffectedRow,
            int rowDelta,
            int? lastDeletedRow,
            CancellationToken cancellationToken = default) {
            ConnectionsPart? connectionsPart = WorkbookPartRoot.ConnectionsPart;
            Connections? connections = connectionsPart?.Connections;
            if (connections == null) {
                return;
            }

            HashSet<uint> connectionIds = GetWorksheetQueryConnectionIds(_worksheetPart);
            bool changed = false;
            foreach (Connection connection in connections.Elements<Connection>()) {
                cancellationToken.ThrowIfCancellationRequested();
                foreach (Parameter parameter in connection.Descendants<Parameter>()) {
                    cancellationToken.ThrowIfCancellationRequested();
                    if (parameter.Cell?.Value is not string referenceText
                        || !ExcelReference.TryParse(referenceText, out ExcelReference? reference)
                        || !ConnectionParameterTargetsCurrentSheet(connection, reference!, connectionIds)
                        || !TryRemapConnectionParameterRows(
                            reference!, firstAffectedRow, rowDelta, lastDeletedRow, out ExcelReference? remappedReference)) {
                        continue;
                    }

                    if (remappedReference == null) {
                        parameter.Remove();
                    } else {
                        parameter.Cell = remappedReference.ToString();
                    }
                    changed = true;
                }

                foreach (Parameters parameters in connection.Elements<Parameters>().ToList()) {
                    uint count = (uint)parameters.Elements<Parameter>().Count();
                    if (count == 0U) {
                        parameters.Remove();
                    } else {
                        parameters.Count = count;
                    }
                }
            }

            if (changed) {
                connections.Save();
            }
        }

        private static HashSet<uint> GetWorksheetQueryConnectionIds(
            WorksheetPart worksheetPart,
            MutationPlanScanBudget? budget = null) {
            return new HashSet<uint>(InspectMutationPlanElements(worksheetPart.QueryTableParts, budget)
                .Select(part => part.QueryTable?.ConnectionId?.Value)
                .Where(id => id.HasValue)
                .Select(id => id!.Value));
        }

        private bool ConnectionParameterTargetsCurrentSheet(
            Connection connection,
            ExcelReference reference,
            HashSet<uint> worksheetConnectionIds) {
            if (reference.IsQualified) {
                return IsCurrentSheetQualifier(reference.Qualifier!, Name);
            }
            return connection.Id?.Value is uint id && worksheetConnectionIds.Contains(id);
        }

        private static bool TryRemapConnectionParameterRows(
            ExcelReference reference,
            int firstAffectedRow,
            int rowDelta,
            int? lastDeletedRow,
            out ExcelReference? remapped) {
            reference.GetBounds(out int r1, out int c1, out int r2, out int c2);
            if (!TryRemapShiftedReferenceRows((r1, c1, r2, c2), firstAffectedRow, rowDelta, lastDeletedRow, out var bounds)) {
                remapped = reference;
                return false;
            }
            remapped = bounds == null
                ? null
                : reference.WithCoordinates(reference.Kind, bounds.Value.r1, bounds.Value.c1, bounds.Value.r2, bounds.Value.c2);
            return true;
        }

        private bool RemapShiftedTables(int firstAffectedRow, int rowDelta, int? lastDeletedRow) {
            bool anyChanged = false;
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

                changed |= RemapShiftedSortStateReferences(
                    table,
                    firstAffectedRow,
                    rowDelta,
                    lastDeletedRow);

                if (changed) {
                    table.Save();
                    anyChanged = true;
                }
            }
            return anyChanged;
        }

        private void RemapShiftedWorksheetRangeMetadata(int firstAffectedRow, int rowDelta, int? lastDeletedRow) {
            foreach (AutoFilter autoFilter in WorksheetRoot.Descendants<AutoFilter>().ToList()) {
                if (autoFilter.Reference?.Value is string filterReference
                    && TryRemapShiftedReferenceListRows(
                        filterReference,
                        firstAffectedRow,
                        rowDelta,
                        lastDeletedRow,
                        out List<string> remappedFilterReferences)) {
                    if (remappedFilterReferences.Count == 0) {
                        autoFilter.Remove();
                    } else {
                        autoFilter.Reference = remappedFilterReferences[0];
                    }
                }
            }

            SheetDimension? dimension = WorksheetRoot.GetFirstChild<SheetDimension>();
            if (dimension?.Reference?.Value is string dimensionReference
                && TryRemapShiftedReferenceListRows(
                    dimensionReference,
                    firstAffectedRow,
                    rowDelta,
                    lastDeletedRow,
                    out List<string> remappedDimensionReferences)) {
                if (remappedDimensionReferences.Count == 0) {
                    dimension.Remove();
                } else {
                    dimension.Reference = remappedDimensionReferences[0];
                }
            }

            RemapShiftedProtectedRanges(firstAffectedRow, rowDelta, lastDeletedRow);
            RemapShiftedIgnoredErrors(firstAffectedRow, rowDelta, lastDeletedRow);
            RemapShiftedScenarios(firstAffectedRow, rowDelta, lastDeletedRow);
            RemapShiftedCellWatches(firstAffectedRow, rowDelta, lastDeletedRow);
            RemapShiftedCellSmartTags(firstAffectedRow, rowDelta, lastDeletedRow);
            RemapShiftedSortStateReferences(WorksheetRoot, firstAffectedRow, rowDelta, lastDeletedRow);
            RemapShiftedQueryTableSortStates(firstAffectedRow, rowDelta, lastDeletedRow);
            RemapShiftedSelections(firstAffectedRow, rowDelta, lastDeletedRow);
            RemapShiftedNamedSheetViewFilters(firstAffectedRow, rowDelta, lastDeletedRow);

            foreach (RowBreaks rowBreaks in WorksheetRoot.Descendants<RowBreaks>().ToList()) {
                bool changed = false;
                foreach (Break pageBreak in rowBreaks.Elements<Break>().ToList()) {
                    if (pageBreak.Id?.Value is not uint rowId || rowId == 0U) {
                        continue;
                    }

                    int row = checked((int)rowId);
                    if (lastDeletedRow.HasValue && row >= firstAffectedRow && row <= lastDeletedRow.Value) {
                        pageBreak.Remove();
                        changed = true;
                        continue;
                    }

                    if (row < firstAffectedRow) {
                        continue;
                    }

                    int shiftedRow = row + rowDelta;
                    if (shiftedRow <= 0 || shiftedRow > A1.MaxRows) {
                        pageBreak.Remove();
                    } else {
                        pageBreak.Id = (uint)shiftedRow;
                    }
                    changed = true;
                }

                if (!changed) {
                    continue;
                }

                uint breakCount = (uint)rowBreaks.Elements<Break>().Count();
                if (breakCount == 0U) {
                    rowBreaks.Remove();
                } else {
                    rowBreaks.Count = breakCount;
                    rowBreaks.ManualBreakCount = (uint)rowBreaks.Elements<Break>()
                        .Count(pageBreak => pageBreak.ManualPageBreak?.Value == true);
                }
            }
        }

        private void RemapShiftedPivotLocations(int firstAffectedRow, int rowDelta, int? lastDeletedRow) {
            foreach (PivotTablePart pivotPart in _worksheetPart.PivotTableParts) {
                Location? location = pivotPart.PivotTableDefinition?.Location;
                if (location?.Reference?.Value is not string locationReference
                    || !TryRemapShiftedReferenceListRows(
                        locationReference,
                        firstAffectedRow,
                        rowDelta,
                        lastDeletedRow,
                        out List<string> remappedLocations)) {
                    continue;
                }

                location.Reference = remappedLocations.Count == 0 ? "#REF!" : remappedLocations[0];
                pivotPart.PivotTableDefinition?.Save();
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

            if (changed) {
                commentsPart.Comments.Save();
            }

            RemapCommentVmlShapes(
                removed,
                moved,
                firstAffectedRow,
                structuralRowDelta: rowDelta,
                lastDeletedRow);
            if (changed) {
                CleanupCommentArtifacts();
            }
        }

        private void RemapShiftedThreadedComments(int firstAffectedRow, int rowDelta, int? lastDeletedRow) {
            foreach (WorksheetThreadedCommentsPart part in _worksheetPart.WorksheetThreadedCommentsParts.ToList()) {
                Threaded.ThreadedComments? comments = part.ThreadedComments;
                if (comments == null) {
                    continue;
                }

                List<Threaded.ThreadedComment> allComments = comments.Elements<Threaded.ThreadedComment>().ToList();
                var removedIds = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
                foreach (Threaded.ThreadedComment comment in allComments) {
                    if (comment.Ref?.Value is not string reference) {
                        continue;
                    }

                    (int Row, int Col) cell = A1.ParseCellRef(reference);
                    if (TryRemapShiftedReferenceRows(
                            (cell.Row, cell.Col, cell.Row, cell.Col),
                            firstAffectedRow,
                            rowDelta,
                            lastDeletedRow,
                            out var remapped)
                        && remapped == null
                        && comment.Id?.Value is string id) {
                        removedIds.Add(id);
                    }
                }

                bool added;
                do {
                    added = false;
                    foreach (Threaded.ThreadedComment comment in allComments) {
                        if (comment.Id?.Value is string id
                            && comment.ParentId?.Value is string parentId
                            && removedIds.Contains(parentId)
                            && removedIds.Add(id)) {
                            added = true;
                        }
                    }
                } while (added);

                bool changed = false;
                foreach (Threaded.ThreadedComment comment in allComments) {
                    if (comment.Id?.Value is string id && removedIds.Contains(id)) {
                        comment.Remove();
                        changed = true;
                        continue;
                    }

                    if (comment.Ref?.Value is not string reference) {
                        continue;
                    }

                    (int Row, int Col) cell = A1.ParseCellRef(reference);
                    if (!TryRemapShiftedReferenceRows(
                        (cell.Row, cell.Col, cell.Row, cell.Col),
                        firstAffectedRow,
                        rowDelta,
                        lastDeletedRow,
                        out var remapped)) {
                        continue;
                    }

                    if (remapped == null) {
                        comment.Remove();
                        changed = true;
                        continue;
                    }

                    string newReference = A1.CellReference(remapped.Value.r1, remapped.Value.c1);
                    if (!string.Equals(reference, newReference, StringComparison.OrdinalIgnoreCase)) {
                        comment.Ref = newReference;
                        changed = true;
                    }
                }

                if (!changed) {
                    continue;
                }

                if (comments.Elements<Threaded.ThreadedComment>().Any()) {
                    comments.Save();
                } else {
                    _worksheetPart.DeletePart(part);
                }
            }

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

            if (!hyperlinks.Elements<Hyperlink>().Any()) {
                hyperlinks.Remove();
            }
        }

        private void RemapShiftedConditionalFormatting(int firstAffectedRow, int rowDelta, int? lastDeletedRow) {
            foreach (var conditional in WorksheetRoot.Elements<ConditionalFormatting>().ToList()) {
                if (conditional.SequenceOfReferences?.InnerText is not string references
                    || !TryGetReferenceListAnchorRow(references, out int oldAnchorRow)) {
                    continue;
                }

                string updatedReferences = references;
                if (TryRemapShiftedReferenceListRows(
                    references,
                    firstAffectedRow,
                    rowDelta,
                    lastDeletedRow,
                    out var remapped)) {
                    if (remapped.Count == 0) {
                        conditional.Remove();
                        continue;
                    }

                    updatedReferences = string.Join(" ", remapped);
                    conditional.SequenceOfReferences = new ListValue<StringValue> { InnerText = updatedReferences };
                }

                if (!TryGetReferenceListAnchorRow(updatedReferences, out int newAnchorRow)) {
                    continue;
                }

                int anchorRowDelta = newAnchorRow - oldAnchorRow;
                int relativeFormulaSourceRowDelta = GetRelativeFormulaSourceRowDelta(
                    oldAnchorRow,
                    newAnchorRow,
                    firstAffectedRow,
                    rowDelta,
                    lastDeletedRow);
                foreach (Formula formula in conditional.Descendants<Formula>()) {
                    RewriteAnchoredFormulaText(
                        formula,
                        firstAffectedRow,
                        rowDelta,
                        lastDeletedRow,
                        anchorRowDelta,
                        relativeFormulaSourceRowDelta: relativeFormulaSourceRowDelta);
                }
                foreach (ConditionalFormatValueObject threshold in conditional
                    .Descendants<ConditionalFormatValueObject>()
                    .Where(item => item.Type?.Value == ConditionalFormatValueObjectValues.Formula)) {
                    if (threshold.Val?.Value is not string formulaText || formulaText.Length == 0) {
                        continue;
                    }

                    string rewritten = RewriteAnchoredFormulaReferences(
                        formulaText,
                        firstAffectedRow,
                        rowDelta,
                        lastDeletedRow,
                        Name,
                        anchorRowDelta,
                        relativeReferencesFollowAnchor: false,
                        relativeFormulaSourceRowDelta);
                    if (!string.Equals(formulaText, rewritten, StringComparison.Ordinal)) {
                        threshold.Val = rewritten;
                    }
                }
            }

            RemapShiftedOffice2010ConditionalFormatting(firstAffectedRow, rowDelta, lastDeletedRow);
        }

        private void RemapShiftedSparklines(int firstAffectedRow, int rowDelta, int? lastDeletedRow) {
            foreach (X14.Sparkline sparkline in WorksheetRoot.Descendants<X14.Sparkline>().ToList()) {
                if (sparkline.ReferenceSequence?.Text is string location
                    && TryRemapShiftedReferenceListRows(location, firstAffectedRow, rowDelta, lastDeletedRow, out var remappedLocations)) {
                    if (remappedLocations.Count == 0) {
                        sparkline.Remove();
                        continue;
                    }

                    sparkline.ReferenceSequence.Text = string.Join(" ", remappedLocations);
                }

                if (sparkline.Formula?.Text is string formula && formula.Length > 0) {
                    string rewritten = lastDeletedRow.HasValue
                        ? RewriteDeletedFormulaReferences(
                            formula,
                            firstAffectedRow,
                            lastDeletedRow.Value,
                            rowDelta,
                            Name)
                        : RewriteShiftedFormulaReferences(
                            formula,
                            firstAffectedRow,
                            rowDelta,
                            Name);
                    if (!string.Equals(formula, rewritten, StringComparison.Ordinal)) {
                        sparkline.Formula.Text = rewritten;
                    }
                }
            }

            CleanupEmptySparklineStructures(WorksheetRoot);
        }

        internal static void CleanupEmptySparklineStructures(Worksheet worksheet) {
            foreach (X14.SparklineGroup group in worksheet.Descendants<X14.SparklineGroup>().ToList()) {
                X14.Sparklines? sparklines = group.GetFirstChild<X14.Sparklines>();
                if (sparklines == null || !sparklines.Elements<X14.Sparkline>().Any()) {
                    group.Remove();
                }
            }
            foreach (X14.SparklineGroups groups in worksheet.Descendants<X14.SparklineGroups>().ToList()) {
                if (!groups.Elements<X14.SparklineGroup>().Any()) {
                    groups.Remove();
                }
            }
            foreach (Extension extension in worksheet.Descendants<Extension>().ToList()) {
                if (!extension.ChildElements.Any()) {
                    extension.Remove();
                }
            }
            foreach (ExtensionList extensions in worksheet.Elements<ExtensionList>().ToList()) {
                if (!extensions.Elements<Extension>().Any()) {
                    extensions.Remove();
                }
            }
        }

        private void RemapShiftedDrawingAnchors(int firstAffectedRow, int rowDelta, int? lastDeletedRow) {
            var drawing = _worksheetPart.DrawingsPart?.WorksheetDrawing;
            if (drawing == null) {
                return;
            }

            bool changed = false;
            foreach (OpenXmlElement anchor in drawing.Descendants()
                .Where(element => element is Xdr.OneCellAnchor || element is Xdr.TwoCellAnchor)
                .ToList()) {
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

                    if (placement == Xdr.EditAsValues.OneCell) {
                        int? oldFromRow = TryGetDrawingMarkerRow(twoCellAnchor.FromMarker);
                        bool fromKept = TryRemapDrawingMarkerRow(twoCellAnchor.FromMarker, firstAffectedRow, rowDelta, lastDeletedRow, out bool fromChanged);
                        if (!fromKept) {
                            twoCellAnchor.Remove();
                            changed = true;
                            continue;
                        }

                        changed |= fromChanged;
                        if (fromChanged) {
                            int actualRowDelta = oldFromRow.HasValue
                                && TryGetDrawingMarkerRow(twoCellAnchor.FromMarker) is int newFromRow
                                ? newFromRow - oldFromRow.Value
                                : rowDelta;
                            if (!TryShiftDrawingMarkerRow(twoCellAnchor.ToMarker, actualRowDelta, out bool toShifted)) {
                                twoCellAnchor.Remove();
                                changed = true;
                                continue;
                            }

                            changed |= toShifted;
                        }

                        continue;
                    }

                    bool rangeKept = TryRemapTwoCellAnchorRows(twoCellAnchor, firstAffectedRow, rowDelta, lastDeletedRow, out bool rangeChanged);
                    if (!rangeKept) {
                        twoCellAnchor.Remove();
                        changed = true;
                        continue;
                    }

                    changed |= rangeChanged;
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
                if (!lastDeletedRow.HasValue) {
                    return false;
                }

                int clampedZeroBasedRow = firstAffectedRow - 1;
                if (clampedZeroBasedRow != zeroBasedRow) {
                    marker.RowId.Text = clampedZeroBasedRow.ToString(CultureInfo.InvariantCulture);
                    changed = true;
                }
                return true;
            }

            int remappedZeroBasedRow = remapped.Value.r1 - 1;
            if (remappedZeroBasedRow != zeroBasedRow) {
                marker.RowId.Text = remappedZeroBasedRow.ToString(CultureInfo.InvariantCulture);
                changed = true;
            }

            return true;
        }

        private static bool TryRemapTwoCellAnchorRows(Xdr.TwoCellAnchor anchor, int firstAffectedRow, int rowDelta, int? lastDeletedRow, out bool changed) {
            changed = false;
            if (rowDelta == 0
                || anchor.FromMarker?.RowId?.Text is not string fromRowText
                || anchor.ToMarker?.RowId?.Text is not string toRowText
                || !int.TryParse(fromRowText, NumberStyles.Integer, CultureInfo.InvariantCulture, out int fromZeroBasedRow)
                || !int.TryParse(toRowText, NumberStyles.Integer, CultureInfo.InvariantCulture, out int toZeroBasedRow)) {
                return true;
            }

            int firstSpannedRow = fromZeroBasedRow + 1;
            bool toMarkerInsideRow = long.TryParse(
                    anchor.ToMarker?.RowOffset?.Text,
                    NumberStyles.Integer,
                    CultureInfo.InvariantCulture,
                    out long toRowOffset)
                && toRowOffset != 0L;
            int lastSpannedRow = toZeroBasedRow + (toMarkerInsideRow ? 1 : 0);
            if (lastSpannedRow < firstSpannedRow) {
                return true;
            }

            if (!TryRemapShiftedReferenceRows((firstSpannedRow, 1, lastSpannedRow, 1), firstAffectedRow, rowDelta, lastDeletedRow, out var remapped)) {
                return true;
            }

            if (remapped == null) {
                if (!lastDeletedRow.HasValue) {
                    return false;
                }

                int clampedZeroBasedRow = firstAffectedRow - 1;
                if (clampedZeroBasedRow != fromZeroBasedRow) {
                    anchor.FromMarker.RowId!.Text = clampedZeroBasedRow.ToString(CultureInfo.InvariantCulture);
                    changed = true;
                }
                if (clampedZeroBasedRow != toZeroBasedRow) {
                    anchor.ToMarker!.RowId!.Text = clampedZeroBasedRow.ToString(CultureInfo.InvariantCulture);
                    changed = true;
                }
                return true;
            }

            int remappedFromZeroBasedRow = remapped.Value.r1 - 1;
            int remappedToZeroBasedRow = remapped.Value.r2 - (toMarkerInsideRow ? 1 : 0);
            if (remappedFromZeroBasedRow < 0 || remappedToZeroBasedRow < remappedFromZeroBasedRow || remappedToZeroBasedRow > A1.MaxRows) {
                return false;
            }

            if (remappedFromZeroBasedRow != fromZeroBasedRow) {
                anchor.FromMarker.RowId!.Text = remappedFromZeroBasedRow.ToString(CultureInfo.InvariantCulture);
                changed = true;
            }

            if (remappedToZeroBasedRow != toZeroBasedRow) {
                anchor.ToMarker!.RowId!.Text = remappedToZeroBasedRow.ToString(CultureInfo.InvariantCulture);
                changed = true;
            }

            return true;
        }

        private static int? TryGetDrawingMarkerRow(Xdr.MarkerType? marker) {
            return int.TryParse(
                marker?.RowId?.Text,
                NumberStyles.Integer,
                CultureInfo.InvariantCulture,
                out int row)
                ? row
                : (int?)null;
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
                foreach (OpenXmlLeafTextElement formula in chartSpace.Descendants<OpenXmlLeafTextElement>()
                    .Where(element => string.Equals(element.LocalName, "f", StringComparison.Ordinal))) {
                    bool formulaChanged = RewriteStructuralFormulaText(
                        formula,
                        firstAffectedRow,
                        rowDelta,
                        lastDeletedRow,
                        rewriteUnqualifiedReferences: true);
                    changed |= formulaChanged;
                    if (formulaChanged) {
                        InvalidateChartFormulaCache(formula);
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
