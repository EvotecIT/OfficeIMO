using OfficeIMO.Excel.Xlsb.Biff12;
using OfficeIMO.Excel.Xlsb.Model;
using System.Globalization;

namespace OfficeIMO.Excel.Xlsb.NameRecords {
    /// <summary>Decodes the common BIFF12 defined-name formula subset.</summary>
    internal static class XlsbDefinedNameFormulaCodec {
        internal static bool TryDecode(
            XlsbDefinedName definedName,
            IReadOnlyList<XlsbExternalSheetReference> externalSheetReferences,
            IReadOnlyList<bool> selfSupportingLinks,
            IReadOnlyList<XlsbWorksheet> worksheets,
            out string? formulaText,
            out string? reason) {
            if (definedName == null) throw new ArgumentNullException(nameof(definedName));
            formulaText = null;
            reason = null;
            if (!definedName.IsSimpleName) {
                reason = "macro, future-function, published, or workbook-parameter defined names";
                return false;
            }

            if (TryDecodeReferenceFormula(
                definedName.FormulaTokens,
                externalSheetReferences,
                selfSupportingLinks,
                worksheets,
                out formulaText,
                out reason)) {
                return true;
            }
            if (reason != null) return false;

            if (XlsbFormulaTextReader.TryRead(definedName.FormulaTokens, out formulaText)) {
                return true;
            }

            reason = "defined-name formulas outside the supported scalar and internal-range subset";
            return false;
        }

        private static bool TryDecodeReferenceFormula(
            byte[] tokens,
            IReadOnlyList<XlsbExternalSheetReference> externalSheetReferences,
            IReadOnlyList<bool> selfSupportingLinks,
            IReadOnlyList<XlsbWorksheet> worksheets,
            out string? formulaText,
            out string? reason) {
            formulaText = null;
            reason = null;
            var stack = new Stack<string>();
            int offset = 0;
            bool sawReference = false;
            while (offset < tokens.Length) {
                byte token = tokens[offset++];
                int tokenKind = token & 0x1F;
                if (tokenKind == 0x1A) {
                    if (!TryReadReference(tokens, ref offset, out ushort externalSheetIndex, out uint row, out ushort columnBits)) {
                        reason = "a truncated three-dimensional cell reference";
                        return false;
                    }
                    if (!TryFormatReference(
                        externalSheetIndex,
                        row,
                        row,
                        columnBits,
                        columnBits,
                        false,
                        externalSheetReferences,
                        selfSupportingLinks,
                        worksheets,
                        out string? reference,
                        out reason)) {
                        return false;
                    }
                    stack.Push(reference!);
                    sawReference = true;
                } else if (tokenKind == 0x1B) {
                    if (!TryReadArea(tokens, ref offset, out ushort externalSheetIndex, out uint firstRow, out uint lastRow, out ushort firstColumnBits, out ushort lastColumnBits)) {
                        reason = "a truncated three-dimensional area reference";
                        return false;
                    }
                    if (!TryFormatReference(
                        externalSheetIndex,
                        firstRow,
                        lastRow,
                        firstColumnBits,
                        lastColumnBits,
                        true,
                        externalSheetReferences,
                        selfSupportingLinks,
                        worksheets,
                        out string? reference,
                        out reason)) {
                        return false;
                    }
                    stack.Push(reference!);
                    sawReference = true;
                } else if (token == 0x10 && stack.Count >= 2) {
                    string right = stack.Pop();
                    string left = stack.Pop();
                    stack.Push(left + "," + right);
                } else {
                    reason = sawReference ? "mixed reference and expression tokens" : null;
                    return false;
                }
            }

            if (!sawReference) return false;
            if (stack.Count != 1) {
                reason = "an invalid reference-union token sequence";
                return false;
            }
            formulaText = stack.Pop();
            return true;
        }

        private static bool TryFormatReference(
            ushort externalSheetIndex,
            uint firstRow,
            uint lastRow,
            ushort firstColumnBits,
            ushort lastColumnBits,
            bool isArea,
            IReadOnlyList<XlsbExternalSheetReference> externalSheetReferences,
            IReadOnlyList<bool> selfSupportingLinks,
            IReadOnlyList<XlsbWorksheet> worksheets,
            out string? reference,
            out string? reason) {
            reference = null;
            reason = null;
            if (externalSheetIndex >= externalSheetReferences.Count) {
                reason = "a missing external-sheet index";
                return false;
            }

            XlsbExternalSheetReference externalSheet = externalSheetReferences[externalSheetIndex];
            if (externalSheet.SupportingLinkIndex >= selfSupportingLinks.Count
                || !selfSupportingLinks[checked((int)externalSheet.SupportingLinkIndex)]) {
                reason = "an external workbook or non-self supporting link";
                return false;
            }
            if (externalSheet.FirstSheetIndex != externalSheet.LastSheetIndex
                || externalSheet.FirstSheetIndex < 0
                || externalSheet.FirstSheetIndex >= worksheets.Count) {
                reason = "a missing or multi-sheet three-dimensional reference";
                return false;
            }
            if (firstRow >= 1_048_576 || lastRow >= 1_048_576 || firstRow > lastRow) {
                reason = "rows outside XLSB worksheet limits";
                return false;
            }
            if ((firstColumnBits & 0xC000) != 0 || (lastColumnBits & 0xC000) != 0) {
                reason = "relative defined-name references";
                return false;
            }

            int firstColumn = firstColumnBits & 0x3FFF;
            int lastColumn = lastColumnBits & 0x3FFF;
            if (firstColumn >= 16_384 || lastColumn >= 16_384 || firstColumn > lastColumn) {
                reason = "columns outside XLSB worksheet limits";
                return false;
            }

            string localReference;
            if (isArea && firstColumn == 0 && lastColumn == 16_383) {
                localReference = "$" + (firstRow + 1).ToString(CultureInfo.InvariantCulture)
                    + ":$" + (lastRow + 1).ToString(CultureInfo.InvariantCulture);
            } else if (isArea && firstRow == 0 && lastRow == 1_048_575) {
                localReference = "$" + A1.ColumnIndexToLetters(firstColumn + 1)
                    + ":$" + A1.ColumnIndexToLetters(lastColumn + 1);
            } else {
                string first = A1.AbsoluteCellReference(checked((int)firstRow + 1), firstColumn + 1);
                string last = A1.AbsoluteCellReference(checked((int)lastRow + 1), lastColumn + 1);
                localReference = !isArea || string.Equals(first, last, StringComparison.Ordinal)
                    ? first
                    : first + ":" + last;
            }

            string sheetName = worksheets[externalSheet.FirstSheetIndex].Name.Replace("'", "''");
            reference = "'" + sheetName + "'!" + localReference;
            return true;
        }

        private static bool TryReadReference(
            byte[] data,
            ref int offset,
            out ushort externalSheetIndex,
            out uint row,
            out ushort columnBits) {
            externalSheetIndex = 0;
            row = 0;
            columnBits = 0;
            return TryReadUInt16(data, ref offset, out externalSheetIndex)
                && TryReadUInt32(data, ref offset, out row)
                && TryReadUInt16(data, ref offset, out columnBits);
        }

        private static bool TryReadArea(
            byte[] data,
            ref int offset,
            out ushort externalSheetIndex,
            out uint firstRow,
            out uint lastRow,
            out ushort firstColumnBits,
            out ushort lastColumnBits) {
            externalSheetIndex = 0;
            firstRow = 0;
            lastRow = 0;
            firstColumnBits = 0;
            lastColumnBits = 0;
            return TryReadUInt16(data, ref offset, out externalSheetIndex)
                && TryReadUInt32(data, ref offset, out firstRow)
                && TryReadUInt32(data, ref offset, out lastRow)
                && TryReadUInt16(data, ref offset, out firstColumnBits)
                && TryReadUInt16(data, ref offset, out lastColumnBits);
        }

        private static bool TryReadUInt16(byte[] data, ref int offset, out ushort value) {
            if (offset > data.Length - 2) {
                value = 0;
                return false;
            }
            value = (ushort)(data[offset] | (data[offset + 1] << 8));
            offset += 2;
            return true;
        }

        private static bool TryReadUInt32(byte[] data, ref int offset, out uint value) {
            if (offset > data.Length - 4) {
                value = 0;
                return false;
            }
            value = (uint)(data[offset]
                | (data[offset + 1] << 8)
                | (data[offset + 2] << 16)
                | (data[offset + 3] << 24));
            offset += 4;
            return true;
        }
    }
}
