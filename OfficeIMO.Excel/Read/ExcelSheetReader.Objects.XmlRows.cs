using System.Globalization;
using System.Diagnostics.CodeAnalysis;
using System.Reflection;
using System.Threading;
using System.Text;
using System.ComponentModel;
using System.Linq.Expressions;
using System.Runtime.Serialization;
using System.Collections.Concurrent;
using System.Collections.Generic;
using System.IO;
using System.Xml;
using DocumentFormat.OpenXml.Spreadsheet;

namespace OfficeIMO.Excel {
    /// <summary>
    /// Object-mapping readers for <see cref="ExcelSheetReader"/>.
    /// </summary>
    public sealed partial class ExcelSheetReader {
        private object?[] ReadXmlRowValues(XmlReader rowReader, int rowIndex, int c1, int c2, int cols, CancellationToken ct) {
            var values = new object?[cols];
            ReadXmlRowValuesInto(rowReader, rowIndex, c1, c2, values, ct);
            return values;
        }

        private void ReadXmlRowValuesInto(XmlReader rowReader, int rowIndex, int c1, int c2, object?[] values, CancellationToken ct) {
            if (rowReader.IsEmptyElement) {
                return;
            }

            int depth = rowReader.Depth;
            bool canCancel = ct.CanBeCanceled;
            int nextColumnIndex = 1;
            int cols = values.Length;
            bool canTrackColumns = cols <= 64;
            ulong allColumnsSeen = canTrackColumns ? CreateAllColumnsSeenMask(cols) : 0UL;
            ulong seenColumns = 0;
            int visitedNodes = 0;
            while (rowReader.Read()) {
                if (canCancel && (++visitedNodes & 1023) == 0) {
                    ct.ThrowIfCancellationRequested();
                }

                if (rowReader.NodeType == XmlNodeType.EndElement && rowReader.Depth == depth && rowReader.LocalName == "row") {
                    return;
                }

                if (rowReader.NodeType != XmlNodeType.Element || rowReader.LocalName != "c") {
                    continue;
                }

                int columnIndex = GetXmlCellColumnIndex(rowReader, ref nextColumnIndex);
                if (columnIndex < c1 || columnIndex > c2) {
                    SkipXmlElement(rowReader, "c");
                    continue;
                }

                object? value;
                if (_opt.CellValueConverter == null) {
                    value = ReadXmlCellValue(rowReader);
                } else {
                    CellRaw raw = ReadXmlCellRaw(rowReader, rowIndex, columnIndex);
                    value = ConvertRaw(raw).TypedValue;
                }

                values[columnIndex - c1] = value;
                if (canTrackColumns && MarkRequestedColumnSeen(columnIndex - c1, allColumnsSeen, ref seenColumns)) {
                    SkipXmlElementContent(rowReader, depth);
                    return;
                }
            }
        }

        private void ReadXmlRowIntoTypedObject<T>(
            XmlReader rowReader,
            int rowIndex,
            int c1,
            int c2,
            TypedPropertyBinding<T>?[] bindings,
            bool canTrackMappedColumns,
            ulong mappedColumns,
            T target,
            CancellationToken ct) {
            if (rowReader.IsEmptyElement) {
                return;
            }

            int depth = rowReader.Depth;
            bool canCancel = ct.CanBeCanceled;
            int nextColumnIndex = 1;
            int convertedCells = 0;
            ulong seenMappedColumns = 0;
            bool canUseOrderedAllMappedExit = canTrackMappedColumns && mappedColumns == CreateAllColumnsSeenMask(bindings.Length);
            int nextExpectedMappedColumn = c1;
            while (rowReader.Read()) {
                if (canCancel && (++convertedCells & 1023) == 0) {
                    ct.ThrowIfCancellationRequested();
                }

                if (rowReader.NodeType == XmlNodeType.EndElement && rowReader.Depth == depth && rowReader.LocalName == "row") {
                    return;
                }

                if (rowReader.NodeType != XmlNodeType.Element || rowReader.LocalName != "c") {
                    continue;
                }

                int columnIndex = GetXmlCellColumnIndex(rowReader, ref nextColumnIndex);
                if (columnIndex < c1 || columnIndex > c2) {
                    if (canUseOrderedAllMappedExit && columnIndex > c2 && nextExpectedMappedColumn <= c2) {
                        canUseOrderedAllMappedExit = false;
                        int orderedSeen = nextExpectedMappedColumn - c1;
                        seenMappedColumns = orderedSeen <= 0 ? 0UL : CreateAllColumnsSeenMask(orderedSeen);
                    }

                    SkipXmlElement(rowReader, "c");
                    continue;
                }

                if (canUseOrderedAllMappedExit && columnIndex != nextExpectedMappedColumn) {
                    canUseOrderedAllMappedExit = false;
                    int orderedSeen = nextExpectedMappedColumn - c1;
                    seenMappedColumns = orderedSeen <= 0 ? 0UL : CreateAllColumnsSeenMask(orderedSeen);
                }

                var binding = bindings[columnIndex - c1];
                if (binding == null) {
                    SkipXmlElement(rowReader, "c");
                    continue;
                }

                ReadXmlCellIntoTypedObject(rowReader, rowIndex, columnIndex, binding, target);
                if (canUseOrderedAllMappedExit) {
                    nextExpectedMappedColumn++;
                }

                if (canUseOrderedAllMappedExit && columnIndex >= c2) {
                    SkipXmlElementContent(rowReader, depth);
                    return;
                }

                if (canTrackMappedColumns && !canUseOrderedAllMappedExit) {
                    seenMappedColumns |= 1UL << (columnIndex - c1);
                    if (seenMappedColumns == mappedColumns) {
                        SkipXmlElementContent(rowReader, depth);
                        return;
                    }
                }
            }
        }

        private static bool TryGetMappedColumnMask<T>(TypedPropertyBinding<T>?[] bindings, out ulong mask) {
            mask = 0;
            for (int i = 0; i < bindings.Length; i++) {
                if (bindings[i] == null) {
                    continue;
                }

                if ((uint)i >= 64u) {
                    mask = 0;
                    return false;
                }

                mask |= 1UL << i;
            }

            return mask != 0;
        }

        private static int GetXmlCellColumnIndex(XmlReader cellReader, ref int nextColumnIndex) {
            string? reference = cellReader.GetAttribute("r");
            int columnIndex = TryGetExpectedSingleLetterColumnIndex(reference, nextColumnIndex, out int expectedColumnIndex)
                ? expectedColumnIndex
                : A1.ParseColumnIndexFromCellReferenceWithKnownRowFast(reference);
            if (columnIndex <= 0) {
                columnIndex = string.IsNullOrEmpty(reference) ? nextColumnIndex : 0;
            }

            if (columnIndex > 0) {
                nextColumnIndex = columnIndex + 1;
            }

            return columnIndex;
        }

        private static bool TryGetExpectedSingleLetterColumnIndex(string? reference, int expectedColumnIndex, out int columnIndex) {
            columnIndex = 0;
            if ((uint)(expectedColumnIndex - 1) >= 26U
                || string.IsNullOrEmpty(reference)) {
                return false;
            }

            string text = reference!;
            if (text.Length < 2) {
                return false;
            }

            char expectedUpper = (char)('A' + expectedColumnIndex - 1);
            char first = text[0];
            if (first != expectedUpper && first != (char)(expectedUpper + 32)) {
                return false;
            }

            char firstRowDigit = text[1];
            int length = text.Length;
            if (firstRowDigit >= '1'
                && firstRowDigit <= '9'
                && (length == 2
                    || (length == 3 && IsAsciiDigit(text[2]))
                    || (length == 4 && IsAsciiDigit(text[2]) && IsAsciiDigit(text[3]))
                    || (length == 5 && IsAsciiDigit(text[2]) && IsAsciiDigit(text[3]) && IsAsciiDigit(text[4]))
                    || (length == 6 && IsAsciiDigit(text[2]) && IsAsciiDigit(text[3]) && IsAsciiDigit(text[4]) && IsAsciiDigit(text[5]))
                    || (length == 7 && IsAsciiDigit(text[2]) && IsAsciiDigit(text[3]) && IsAsciiDigit(text[4]) && IsAsciiDigit(text[5]) && IsAsciiDigit(text[6]))
                    || (length == 8 && IsAsciiDigit(text[2]) && IsAsciiDigit(text[3]) && IsAsciiDigit(text[4]) && IsAsciiDigit(text[5]) && IsAsciiDigit(text[6]) && IsAsciiDigit(text[7])))) {
                columnIndex = expectedColumnIndex;
                return true;
            }

            bool hasNonZeroRowDigit = false;
            for (int i = 1; i < text.Length; i++) {
                char ch = text[i];
                if (ch < '0' || ch > '9') {
                    return false;
                }

                hasNonZeroRowDigit |= ch != '0';
            }

            if (!hasNonZeroRowDigit) {
                return false;
            }

            columnIndex = expectedColumnIndex;
            return true;
        }

        private static bool IsAsciiDigit(char value) {
            return (uint)(value - '0') <= 9U;
        }

        private CellRaw ReadXmlCellRaw(XmlReader cellReader, int rowIndex, int columnIndex) {
            XmlCellKind cellKind = ParseXmlCellKind(cellReader.GetAttribute("t"));
            bool readStyleIndex = _opt.TreatDatesUsingNumberFormat
                && CellKindCanUseDateStyle(cellKind);
            return ReadXmlCellRaw(cellReader, rowIndex, columnIndex, cellKind, readStyleIndex);
        }

        private CellRaw ReadXmlCellRaw<TTarget>(XmlReader cellReader, int rowIndex, int columnIndex, TypedPropertyBinding<TTarget> binding) {
            XmlCellKind cellKind = ParseXmlCellKind(cellReader.GetAttribute("t"));
            bool readStyleIndex = _opt.CellValueConverter != null
                || (_opt.TreatDatesUsingNumberFormat
                && binding.NeedsDateStyleConversion
                && CellKindCanUseDateStyle(cellKind));
            return ReadXmlCellRaw(cellReader, rowIndex, columnIndex, cellKind, readStyleIndex);
        }

        private void ReadXmlCellIntoTypedObject<TTarget>(XmlReader cellReader, int rowIndex, int columnIndex, TypedPropertyBinding<TTarget> binding, TTarget target) {
            if (_opt.CellValueConverter != null || _opt.TypeConverter != null) {
                CellRaw converterRaw = ReadXmlCellRaw(cellReader, rowIndex, columnIndex, binding);
                TrySetRawCellForBinding(converterRaw, binding, target);
                return;
            }

            XmlCellKind cellKind = ParseXmlCellKind(cellReader.GetAttribute("t"));
            uint? styleIndex = null;
            if (_opt.TreatDatesUsingNumberFormat
                && binding.NeedsDateStyleConversion
                && CellKindCanUseDateStyle(cellKind)
                && TryParseUInt(cellReader.GetAttribute("s"), out uint parsedStyle)) {
                if (Styles.HasDateStyles) {
                    styleIndex = parsedStyle;
                }
            }

            if (cellReader.IsEmptyElement) {
                TrySetRawCellForBinding(CreateRawCell(rowIndex, columnIndex, cellKind, styleIndex, hasFormula: false, formulaText: null, rawText: null, inlineText: null), binding, target);
                return;
            }

            int depth = cellReader.Depth;
            string? rawText = null;
            string? inlineText = null;
            string? formulaText = null;
            bool hasFormula = false;
            bool sawUnsupportedNode = false;
            bool hasNode = cellReader.Read();
            while (hasNode) {
                if (cellReader.NodeType == XmlNodeType.EndElement && cellReader.Depth == depth && cellReader.LocalName == "c") {
                    break;
                }

                if (cellReader.NodeType == XmlNodeType.Element) {
                    if (cellReader.LocalName == "v") {
                        if (cellKind == XmlCellKind.SharedString
                            && !hasFormula
                            && inlineText == null
                            && binding.SetString != null
                            && binding.BindingKind == TypedBindingKind.String
                            && _opt.UseCachedFormulaResult) {
                            bool parsedSharedStringIndex = TryReadXmlSharedStringIndexValueAndSkipCell(cellReader, depth, out int sstIndex, out rawText);
                            binding.SetString(target, parsedSharedStringIndex ? GetSharedString(sstIndex) : rawText);
                            return;
                        }

                        rawText = cellReader.ReadElementContentAsString();
                        if (cellReader.NodeType == XmlNodeType.EndElement
                            && cellReader.Depth == depth
                            && cellReader.LocalName == "c"
                            && TrySetSimpleRawTextCell(cellKind, styleIndex, rawText, binding, target)) {
                            return;
                        }

                        hasNode = true;
                        continue;
                    }

                    if (cellReader.LocalName == "f") {
                        hasFormula = true;
                        formulaText = cellReader.ReadElementContentAsString();
                        if (!_opt.UseCachedFormulaResult) {
                            SkipXmlElementContent(cellReader, depth, "c");
                            TrySetRawCellForBinding(CreateRawCell(rowIndex, columnIndex, cellKind, styleIndex, hasFormula: true, formulaText, rawText: null, inlineText: null), binding, target);
                            return;
                        }

                        hasNode = true;
                        continue;
                    }

                    if (cellReader.LocalName == "is") {
                        inlineText = ReadXmlInlineString(cellReader);
                        hasNode = true;
                        continue;
                    }

                    sawUnsupportedNode = true;
                }

                hasNode = cellReader.Read();
            }

            bool preferFormulaText = hasFormula && !_opt.UseCachedFormulaResult && formulaText != null;
            if (!sawUnsupportedNode
                && !hasFormula
                && inlineText == null
                && rawText != null
                && TrySetSimpleRawTextCell(cellKind, styleIndex, rawText, binding, target)) {
                return;
            }

            TrySetRawCellForBinding(CreateRawCell(
                rowIndex,
                columnIndex,
                cellKind,
                styleIndex,
                hasFormula,
                formulaText,
                preferFormulaText ? null : rawText,
                preferFormulaText ? null : inlineText),
                binding,
                target);
        }

        private bool TrySetSimpleRawTextCell<TTarget>(
            XmlCellKind cellKind,
            uint? styleIndex,
            string rawText,
            TypedPropertyBinding<TTarget> binding,
            TTarget target) {
            switch (cellKind) {
                case XmlCellKind.SharedString: {
                    string? text = TryParseSharedStringIndex(rawText, out int sstIndex) ? GetSharedString(sstIndex) : rawText;
                    return TrySetStringTextBinding(text, binding, target);
                }

                case XmlCellKind.Boolean:
                    if (binding.SetBoolean != null && binding.BindingKind == TypedBindingKind.Boolean) {
                        binding.SetBoolean(target, rawText == "1");
                        return true;
                    }

                    if (binding.SetString != null && binding.BindingKind == TypedBindingKind.String) {
                        binding.SetString(target, (rawText == "1").ToString());
                        return true;
                    }

                    return false;

                case XmlCellKind.String:
                case XmlCellKind.InlineString:
                    return TrySetStringTextBinding(rawText, binding, target);

                case XmlCellKind.Date:
                    if (binding.SetDateTime != null
                        && DateTime.TryParse(rawText, _opt.Culture, DateTimeStyles.AssumeLocal, out var dateValue)) {
                        binding.SetDateTime(target, dateValue);
                        return true;
                    }

                    return TrySetStringTextBinding(rawText, binding, target);
            }

            if (_opt.TreatDatesUsingNumberFormat
                && binding.NeedsDateStyleConversion
                && styleIndex is not null
                && Styles.IsDateLike(styleIndex.Value)) {
                if (TryParseInvariantDoubleFast(rawText, out var oa)
                    || double.TryParse(rawText, NumberStyles.Float | NumberStyles.AllowThousands, CultureInfo.InvariantCulture, out oa)) {
                    DateTime dateValue = DateTime.FromOADate(oa);
                    if (binding.SetDateTime != null && binding.BindingKind == TypedBindingKind.DateTime) {
                        binding.SetDateTime(target, dateValue);
                        return true;
                    }

                    if (binding.SetString != null && binding.BindingKind == TypedBindingKind.String) {
                        binding.SetString(target, dateValue.ToString(_opt.Culture));
                        return true;
                    }
                }

                return TrySetStringTextBinding(rawText, binding, target);
            }

            return TrySetNumericTextBinding(rawText, binding, target);
        }

        private static CellRaw CreateRawCell(
            int rowIndex,
            int columnIndex,
            XmlCellKind cellKind,
            uint? styleIndex,
            bool hasFormula,
            string? formulaText,
            string? rawText,
            string? inlineText) {
            return new CellRaw {
                Row = rowIndex,
                Col = columnIndex,
                TypeHint = ToCellValueType(cellKind),
                StyleIndex = styleIndex,
                HasFormula = hasFormula,
                FormulaText = formulaText,
                RawText = rawText,
                InlineText = inlineText
            };
        }

        private CellRaw ReadXmlCellRaw(XmlReader cellReader, int rowIndex, int columnIndex, XmlCellKind cellKind, bool readStyleIndex) {
            var raw = new CellRaw {
                Row = rowIndex,
                Col = columnIndex,
                TypeHint = ToCellValueType(cellKind)
            };

            if (readStyleIndex && TryParseUInt(cellReader.GetAttribute("s"), out uint parsedStyle)) {
                raw.StyleIndex = parsedStyle;
            }

            if (cellReader.IsEmptyElement) {
                return raw;
            }

            int depth = cellReader.Depth;
            string? rawText = null;
            string? inlineText = null;
            string? formulaText = null;
            bool hasFormula = false;
            bool hasNode = cellReader.Read();
            while (hasNode) {
                if (cellReader.NodeType == XmlNodeType.EndElement && cellReader.Depth == depth && cellReader.LocalName == "c") {
                    break;
                }

                if (cellReader.NodeType == XmlNodeType.Element) {
                    if (cellReader.LocalName == "v") {
                        rawText = cellReader.ReadElementContentAsString();
                        hasNode = true;
                        continue;
                    }

                    if (cellReader.LocalName == "f") {
                        hasFormula = true;
                        formulaText = cellReader.ReadElementContentAsString();
                        if (!_opt.UseCachedFormulaResult) {
                            SkipXmlElementContent(cellReader, depth, "c");
                            raw.HasFormula = true;
                            raw.FormulaText = formulaText;
                            return raw;
                        }

                        hasNode = true;
                        continue;
                    }

                    if (cellReader.LocalName == "is") {
                        inlineText = ReadXmlInlineString(cellReader);
                        hasNode = true;
                        continue;
                    }
                }

                hasNode = cellReader.Read();
            }

            bool preferFormulaText = hasFormula && !_opt.UseCachedFormulaResult && formulaText != null;
            raw.HasFormula = hasFormula;
            raw.FormulaText = formulaText;
            raw.RawText = preferFormulaText ? null : rawText;
            raw.InlineText = preferFormulaText ? null : inlineText;
            return raw;
        }

        private static void SkipXmlElement(XmlReader reader, string localName) {
            if (reader.IsEmptyElement) {
                return;
            }

            int depth = reader.Depth;
            while (reader.Read()) {
                if (reader.NodeType == XmlNodeType.EndElement && reader.Depth == depth && reader.LocalName == localName) {
                    return;
                }
            }
        }

        private static void SkipXmlElementContent(XmlReader reader, int depth, string localName) {
            if (reader.NodeType == XmlNodeType.EndElement && reader.Depth == depth && reader.LocalName == localName) {
                return;
            }

            while (reader.Read()) {
                if (reader.NodeType == XmlNodeType.EndElement && reader.Depth == depth && reader.LocalName == localName) {
                    return;
                }
            }
        }

        private static void SkipXmlElementContent(XmlReader reader, int depth) {
            if (reader.NodeType == XmlNodeType.EndElement && reader.Depth == depth) {
                return;
            }

            while (reader.Read()) {
                if (reader.NodeType == XmlNodeType.EndElement && reader.Depth == depth) {
                    return;
                }
            }
        }
    }
}
