using System.Globalization;
using System.Text.RegularExpressions;
using DocumentFormat.OpenXml;

namespace OfficeIMO.Excel {
    public partial class ExcelSheet {
        private static string RewriteAnchoredFormulaReferences(
            string formula,
            int firstAffectedRow,
            int rowDelta,
            int? lastDeletedRow,
            string sheetName,
            int anchorRowDelta) {
            if (string.IsNullOrEmpty(formula)) {
                return formula;
            }

            return RewriteFormulaReferencesOutsideStrings(formula, segment => {
                var protectedRanges = new List<string>();
                string rewrittenRanges = ReplaceFormulaRanges(segment, match => {
                    string replacement = RewriteAnchoredFormulaRangeReference(
                        match,
                        firstAffectedRow,
                        rowDelta,
                        lastDeletedRow,
                        sheetName,
                        anchorRowDelta);
                    return ProtectRewrittenFormulaRange(match, replacement, protectedRanges);
                });
                rewrittenRanges = ReplaceFormulaRowRanges(rewrittenRanges, match => {
                    string replacement = RewriteAnchoredFormulaRowRangeReference(
                        match,
                        firstAffectedRow,
                        rowDelta,
                        lastDeletedRow,
                        sheetName,
                        anchorRowDelta);
                    return ProtectRewrittenFormulaRange(match, replacement, protectedRanges);
                });

                string rewritten = ReplaceFormulaReferences(rewrittenRanges, match => {
                    if (!IsValidFormulaColumn(match.Groups["col"].Value)
                        || !int.TryParse(
                            match.Groups["row"].Value,
                            NumberStyles.None,
                            CultureInfo.InvariantCulture,
                            out int row)
                        || row <= 0
                        || row > A1.MaxRows) {
                        return match.Value;
                    }

                    bool relativeRow = match.Groups["rowAbs"].Value.Length == 0;
                    if (!relativeRow && !FormulaQualifierTargetsSheet(match.Groups["sheet"].Value, sheetName)) {
                        return match.Value;
                    }

                    if (!TryMapAnchoredFormulaRow(
                            row,
                            relativeRow,
                            firstAffectedRow,
                            rowDelta,
                            lastDeletedRow,
                            anchorRowDelta,
                            out int targetRow)) {
                        return "#REF!";
                    }

                    return targetRow == row ? match.Value : BuildFormulaReference(match, targetRow);
                });

                for (int index = 0; index < protectedRanges.Count; index++) {
                    rewritten = rewritten.Replace(
                        "\u0001A" + index.ToString(CultureInfo.InvariantCulture) + "\u0002",
                        protectedRanges[index]);
                }

                return rewritten;
            });
        }

        private bool RewriteAnchoredFormulaText(
            OpenXmlLeafTextElement? formula,
            int firstAffectedRow,
            int rowDelta,
            int? lastDeletedRow,
            int anchorRowDelta) {
            if (formula?.Text is not string text || text.Length == 0) {
                return false;
            }

            string rewritten = RewriteAnchoredFormulaReferences(
                text,
                firstAffectedRow,
                rowDelta,
                lastDeletedRow,
                Name,
                anchorRowDelta);
            if (string.Equals(text, rewritten, StringComparison.Ordinal)) {
                return false;
            }

            formula.Text = rewritten;
            return true;
        }

        private static string ProtectRewrittenFormulaRange(
            Match match,
            string replacement,
            List<string> protectedRanges) {
            if (string.Equals(replacement, match.Value, StringComparison.Ordinal)) {
                return match.Value;
            }

            string placeholder = "\u0001A"
                + protectedRanges.Count.ToString(CultureInfo.InvariantCulture)
                + "\u0002";
            protectedRanges.Add(replacement);
            return placeholder;
        }

        private static string RewriteAnchoredFormulaRangeReference(
            Match match,
            int firstAffectedRow,
            int rowDelta,
            int? lastDeletedRow,
            string sheetName,
            int anchorRowDelta) {
            if (!IsValidFormulaColumn(match.Groups["startCol"].Value)
                || !IsValidFormulaColumn(match.Groups["endCol"].Value)
                || !int.TryParse(match.Groups["startRow"].Value, NumberStyles.None, CultureInfo.InvariantCulture, out int startRow)
                || !int.TryParse(match.Groups["endRow"].Value, NumberStyles.None, CultureInfo.InvariantCulture, out int endRow)
                || startRow <= 0
                || startRow > A1.MaxRows
                || endRow <= 0
                || endRow > A1.MaxRows) {
                return match.Value;
            }

            bool startRelative = match.Groups["startRowAbs"].Value.Length == 0;
            bool endRelative = match.Groups["endRowAbs"].Value.Length == 0;
            bool startTargetsSheet = FormulaQualifierTargetsSheet(match.Groups["sheet"].Value, sheetName);
            string endQualifier = match.Groups["endSheet"].Value;
            bool endTargetsSheet = endQualifier.Length == 0
                ? startTargetsSheet
                : FormulaQualifierTargetsSheet(endQualifier, sheetName);

            if (!startRelative && !endRelative && startTargetsSheet && endTargetsSheet) {
                return lastDeletedRow.HasValue
                    ? RewriteDeletedFormulaRangeReference(
                        match,
                        firstAffectedRow,
                        lastDeletedRow.Value,
                        rowDelta,
                        sheetName,
                        rewriteUnqualifiedReferences: true)
                    : RewriteShiftedFormulaRangeReference(
                        match,
                        firstAffectedRow,
                        rowDelta,
                        sheetName,
                        rewriteUnqualifiedReferences: true);
            }

            if ((!startRelative && !startTargetsSheet)
                || (!endRelative && !endTargetsSheet)) {
                if (!startRelative && !endRelative) {
                    return match.Value;
                }
            }

            int targetStart = startRow;
            int targetEnd = endRow;
            bool startMapped = !startRelative && !startTargetsSheet
                || TryMapAnchoredFormulaRow(
                    startRow,
                    startRelative,
                    firstAffectedRow,
                    rowDelta,
                    lastDeletedRow,
                    anchorRowDelta,
                    out targetStart);
            bool endMapped = !endRelative && !endTargetsSheet
                || TryMapAnchoredFormulaRow(
                    endRow,
                    endRelative,
                    firstAffectedRow,
                    rowDelta,
                    lastDeletedRow,
                    anchorRowDelta,
                    out targetEnd);
            if (!startMapped || !endMapped) {
                return "#REF!";
            }

            return BuildAnchoredFormulaRange(match, targetStart, targetEnd);
        }

        private static string RewriteAnchoredFormulaRowRangeReference(
            Match match,
            int firstAffectedRow,
            int rowDelta,
            int? lastDeletedRow,
            string sheetName,
            int anchorRowDelta) {
            if (!int.TryParse(match.Groups["startRow"].Value, NumberStyles.None, CultureInfo.InvariantCulture, out int startRow)
                || !int.TryParse(match.Groups["endRow"].Value, NumberStyles.None, CultureInfo.InvariantCulture, out int endRow)
                || startRow <= 0
                || startRow > A1.MaxRows
                || endRow <= 0
                || endRow > A1.MaxRows) {
                return match.Value;
            }

            bool startRelative = match.Groups["startRowAbs"].Value.Length == 0;
            bool endRelative = match.Groups["endRowAbs"].Value.Length == 0;
            bool startTargetsSheet = FormulaQualifierTargetsSheet(match.Groups["sheet"].Value, sheetName);
            string endQualifier = match.Groups["endSheet"].Value;
            bool endTargetsSheet = endQualifier.Length == 0
                ? startTargetsSheet
                : FormulaQualifierTargetsSheet(endQualifier, sheetName);

            if (!startRelative && !endRelative && startTargetsSheet && endTargetsSheet) {
                return lastDeletedRow.HasValue
                    ? RewriteDeletedFormulaRowRangeReference(
                        match,
                        firstAffectedRow,
                        lastDeletedRow.Value,
                        rowDelta,
                        sheetName,
                        rewriteUnqualifiedReferences: true)
                    : RewriteShiftedFormulaRowRangeReference(
                        match,
                        firstAffectedRow,
                        rowDelta,
                        sheetName,
                        rewriteUnqualifiedReferences: true);
            }

            int targetStart = startRow;
            int targetEnd = endRow;
            bool startMapped = !startRelative && !startTargetsSheet
                || TryMapAnchoredFormulaRow(
                    startRow,
                    startRelative,
                    firstAffectedRow,
                    rowDelta,
                    lastDeletedRow,
                    anchorRowDelta,
                    out targetStart);
            bool endMapped = !endRelative && !endTargetsSheet
                || TryMapAnchoredFormulaRow(
                    endRow,
                    endRelative,
                    firstAffectedRow,
                    rowDelta,
                    lastDeletedRow,
                    anchorRowDelta,
                    out targetEnd);
            if (!startMapped || !endMapped) {
                return "#REF!";
            }

            return BuildFormulaRowRange(match, targetStart, targetEnd);
        }

        private static bool TryMapAnchoredFormulaRow(
            int row,
            bool relativeRow,
            int firstAffectedRow,
            int rowDelta,
            int? lastDeletedRow,
            int anchorRowDelta,
            out int targetRow) {
            targetRow = row;
            if (relativeRow) {
                targetRow += anchorRowDelta;
            } else if (lastDeletedRow.HasValue) {
                if (row >= firstAffectedRow && row <= lastDeletedRow.Value) {
                    return false;
                }
                if (row > lastDeletedRow.Value) {
                    targetRow += rowDelta;
                }
            } else if (row >= firstAffectedRow) {
                targetRow += rowDelta;
            }

            return targetRow > 0 && targetRow <= A1.MaxRows;
        }

        private static bool FormulaQualifierTargetsSheet(string qualifier, string sheetName) {
            return qualifier.Length == 0 || IsCurrentSheetQualifier(qualifier, sheetName);
        }

        private static string BuildAnchoredFormulaRange(Match match, int startRow, int endRow) {
            string startQualifier = match.Groups["sheet"].Value;
            string endQualifier = match.Groups["endSheet"].Value;
            return startQualifier
                + (startQualifier.Length > 0 ? "!" : string.Empty)
                + match.Groups["startColAbs"].Value
                + match.Groups["startCol"].Value
                + match.Groups["startRowAbs"].Value
                + startRow.ToString(CultureInfo.InvariantCulture)
                + ":"
                + endQualifier
                + (endQualifier.Length > 0 ? "!" : string.Empty)
                + match.Groups["endColAbs"].Value
                + match.Groups["endCol"].Value
                + match.Groups["endRowAbs"].Value
                + endRow.ToString(CultureInfo.InvariantCulture);
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
