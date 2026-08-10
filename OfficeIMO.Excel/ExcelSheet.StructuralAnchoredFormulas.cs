using DocumentFormat.OpenXml;

namespace OfficeIMO.Excel {
    public partial class ExcelSheet {
        private string RewriteAnchoredFormulaReferences(
            string formula,
            int firstAffectedRow,
            int rowDelta,
            int? lastDeletedRow,
            string sheetName,
            int anchorRowDelta,
            bool relativeReferencesFollowAnchor,
            int relativeFormulaSourceRowDelta,
            int? relativeFormulaAnchorRow = null) {
            if (string.IsNullOrEmpty(formula)) {
                return formula;
            }

            return ExcelFormulaReferenceRewriter.RewriteReferences(formula, match => {
                if (IsFormulaFunctionReferenceToken(formula, match)
                    || !FormulaQualifierTargetsSheet(match.Reference.Qualifier ?? string.Empty, sheetName)) return match.Text;
                if (match.Reference.Kind == ExcelReferenceKind.Cell) {
                    int row = match.Reference.Start.Row;
                    if (!TryMapAnchoredFormulaRow(
                            row,
                            !match.Reference.Start.RowAbsolute,
                            firstAffectedRow,
                            rowDelta,
                            lastDeletedRow,
                            anchorRowDelta,
                            relativeReferencesFollowAnchor,
                            relativeFormulaSourceRowDelta,
                            relativeFormulaAnchorRow,
                            out int targetRow)) return "#REF!";
                    return targetRow == row
                        ? match.Text
                        : FormatFormulaReference(match, targetRow, targetRow);
                }
                if (match.Reference.Kind is ExcelReferenceKind.Range or ExcelReferenceKind.WholeRow) {
                    return RewriteAnchoredFormulaRangeReference(
                        match,
                        firstAffectedRow,
                        rowDelta,
                        lastDeletedRow,
                        anchorRowDelta,
                        relativeReferencesFollowAnchor,
                        relativeFormulaSourceRowDelta,
                        relativeFormulaAnchorRow);
                }
                return match.Text;
            });
        }

        private bool RewriteAnchoredFormulaText(
            OpenXmlLeafTextElement? formula,
            int firstAffectedRow,
            int rowDelta,
            int? lastDeletedRow,
            int anchorRowDelta,
            bool relativeReferencesFollowAnchor = false,
            int relativeFormulaSourceRowDelta = 0,
            int? relativeFormulaAnchorRow = null) {
            if (formula?.Text is not string text || text.Length == 0) {
                return false;
            }

            string rewritten = RewriteAnchoredFormulaReferences(
                text,
                firstAffectedRow,
                rowDelta,
                lastDeletedRow,
                Name,
                anchorRowDelta,
                relativeReferencesFollowAnchor,
                relativeFormulaSourceRowDelta,
                relativeFormulaAnchorRow);
            if (string.Equals(text, rewritten, StringComparison.Ordinal)) {
                return false;
            }

            formula.Text = rewritten;
            return true;
        }

        private static string RewriteAnchoredFormulaRangeReference(
            ExcelFormulaReferenceCandidate match,
            int firstAffectedRow,
            int rowDelta,
            int? lastDeletedRow,
            int anchorRowDelta,
            bool relativeReferencesFollowAnchor,
            int relativeFormulaSourceRowDelta,
            int? relativeFormulaAnchorRow) {
            int startRow = match.Reference.Start.Row;
            int endRow = match.Reference.End.Row;
            bool startRelative = !match.Reference.Start.RowAbsolute;
            bool endRelative = !match.Reference.End.RowAbsolute;

            if (!startRelative && !endRelative) {
                return lastDeletedRow.HasValue
                    ? RewriteDeletedFormulaRangeReference(
                        match,
                        firstAffectedRow,
                        lastDeletedRow.Value,
                        rowDelta)
                    : RewriteShiftedFormulaRangeReference(
                        match,
                        firstAffectedRow,
                        rowDelta);
            }

            bool relativeRangeKeepsFirstDataRowOffsets = relativeReferencesFollowAnchor
                && relativeFormulaAnchorRow.HasValue
                && firstAffectedRow == relativeFormulaAnchorRow.Value
                && (startRelative || endRelative);
            if (lastDeletedRow.HasValue
                && !relativeRangeKeepsFirstDataRowOffsets
                && TryMapAnchoredFormulaRangeRows(
                    startRow,
                    startRelative,
                    endRow,
                    endRelative,
                    firstAffectedRow,
                    rowDelta,
                    lastDeletedRow.Value,
                    relativeFormulaSourceRowDelta,
                    out int survivingStart,
                    out int survivingEnd)) {
                return FormatFormulaReference(match, survivingStart, survivingEnd);
            }

            int targetStart = startRow;
            int targetEnd = endRow;
            bool startMapped = TryMapAnchoredFormulaRow(
                    startRow,
                    startRelative,
                    firstAffectedRow,
                    rowDelta,
                    lastDeletedRow,
                    anchorRowDelta,
                    relativeReferencesFollowAnchor,
                    relativeFormulaSourceRowDelta,
                    relativeFormulaAnchorRow,
                    out targetStart);
            bool endMapped = TryMapAnchoredFormulaRow(
                    endRow,
                    endRelative,
                    firstAffectedRow,
                    rowDelta,
                    lastDeletedRow,
                    anchorRowDelta,
                    relativeReferencesFollowAnchor,
                    relativeFormulaSourceRowDelta,
                    relativeFormulaAnchorRow,
                    out targetEnd);
            if (!startMapped || !endMapped) {
                return "#REF!";
            }

            return FormatFormulaReference(match, targetStart, targetEnd);
        }

        private static bool TryMapAnchoredFormulaRow(
            int row,
            bool relativeRow,
            int firstAffectedRow,
            int rowDelta,
            int? lastDeletedRow,
            int anchorRowDelta,
            bool relativeReferencesFollowAnchor,
            int relativeFormulaSourceRowDelta,
            int? relativeFormulaAnchorRow,
            out int targetRow) {
            targetRow = row;
            if (relativeRow
                && relativeReferencesFollowAnchor
                && relativeFormulaAnchorRow.HasValue
                && firstAffectedRow == relativeFormulaAnchorRow.Value) {
                targetRow += anchorRowDelta;
            } else if (relativeRow
                && relativeReferencesFollowAnchor
                && lastDeletedRow.HasValue
                && row >= firstAffectedRow
                && row <= lastDeletedRow.Value) {
                targetRow += anchorRowDelta;
            } else {
                int sourceRow = relativeRow
                    ? row + relativeFormulaSourceRowDelta
                    : row;
                targetRow = sourceRow;
                if (lastDeletedRow.HasValue) {
                    if (sourceRow >= firstAffectedRow && sourceRow <= lastDeletedRow.Value) {
                        return false;
                    }
                    if (sourceRow > lastDeletedRow.Value) {
                        targetRow += rowDelta;
                    }
                } else if (sourceRow >= firstAffectedRow) {
                    targetRow += rowDelta;
                }
            }

            return targetRow > 0 && targetRow <= A1.MaxRows;
        }

        private static bool TryMapAnchoredFormulaRangeRows(
            int startRow,
            bool startRelative,
            int endRow,
            bool endRelative,
            int firstDeletedRow,
            int rowDelta,
            int lastDeletedRow,
            int relativeFormulaSourceRowDelta,
            out int targetStart,
            out int targetEnd) {
            int sourceStart = startRelative
                ? startRow + relativeFormulaSourceRowDelta
                : startRow;
            int sourceEnd = endRelative
                ? endRow + relativeFormulaSourceRowDelta
                : endRow;
            targetStart = sourceStart;
            targetEnd = sourceEnd;
            if (sourceStart <= 0
                || sourceStart > A1.MaxRows
                || sourceEnd <= 0
                || sourceEnd > A1.MaxRows) {
                return false;
            }

            bool reversed = sourceStart > sourceEnd;
            int firstSourceRow = Math.Min(sourceStart, sourceEnd);
            int lastSourceRow = Math.Max(sourceStart, sourceEnd);
            if (!TryRemapShiftedReferenceRows(
                    (firstSourceRow, 1, lastSourceRow, 1),
                    firstDeletedRow,
                    rowDelta,
                    lastDeletedRow,
                    out var remapped)) {
                return relativeFormulaSourceRowDelta != 0;
            }

            if (remapped == null) {
                targetStart = 0;
                targetEnd = 0;
                return false;
            }

            targetStart = reversed ? remapped.Value.r2 : remapped.Value.r1;
            targetEnd = reversed ? remapped.Value.r1 : remapped.Value.r2;
            return true;
        }

        private static int GetRelativeFormulaSourceRowDelta(
            int oldAnchorRow,
            int newAnchorRow,
            int firstAffectedRow,
            int rowDelta,
            int? lastDeletedRow) {
            int sourceAnchorRow = newAnchorRow;
            if (lastDeletedRow.HasValue) {
                if (newAnchorRow >= firstAffectedRow) {
                    sourceAnchorRow -= rowDelta;
                }
            } else if (newAnchorRow >= firstAffectedRow + rowDelta) {
                sourceAnchorRow -= rowDelta;
            }

            return sourceAnchorRow - oldAnchorRow;
        }

        private static bool FormulaQualifierTargetsSheet(string qualifier, string sheetName) {
            return qualifier.Length == 0 || IsCurrentSheetQualifier(qualifier, sheetName);
        }

        private static bool TryGetReferenceListAnchorRow(string references, out int row) {
            foreach (ReferenceListPart part in SplitReferenceList(references)) {
                if (TryParseReference(part, out var bounds)) {
                    row = bounds.r1;
                    return true;
                }
            }

            row = 0;
            return false;
        }
    }
}
