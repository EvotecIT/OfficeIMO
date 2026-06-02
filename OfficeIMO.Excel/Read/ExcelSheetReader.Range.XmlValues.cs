using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using System.Data;
using System.Globalization;
using System.IO;
using System.Threading;
using System.Threading.Tasks;
using System.Xml;

namespace OfficeIMO.Excel {
    /// <summary>
    /// Range-based read operations for <see cref="ExcelSheetReader"/>.
    /// </summary>
    public sealed partial class ExcelSheetReader {
        private object? ReadXmlCellValue(XmlReader cellReader) {
            return ReadXmlCellValue(cellReader, cellReader.GetAttribute("t"));
        }

        private object? ReadXmlCellValue(XmlReader cellReader, string? cellType) {
            if (cellReader.IsEmptyElement) {
                return null;
            }

            if (_opt.CellValueConverter == null && cellType == "s") {
                var sharedStringItems = _sharedStringItems ??= _sst.GetItems();
                return ReadXmlSharedStringCellValue(cellReader, _opt.UseCachedFormulaResult, sharedStringItems);
            }

            if (_opt.CellValueConverter == null
                && (string.IsNullOrEmpty(cellType) || cellType == "n")) {
                return ReadXmlNumericCellValue(cellReader);
            }

            XmlCellKind cellKind = ParseXmlCellKind(cellType);
            if (_opt.CellValueConverter != null) {
                CellRaw raw = ReadXmlCellRaw(cellReader, 0, 0, cellKind, readStyleIndex: true);
                return ConvertRaw(raw).TypedValue;
            }

            bool useCachedFormulaResult = _opt.UseCachedFormulaResult;
            if (cellKind == XmlCellKind.SharedString) {
                var sharedStringItems = _sharedStringItems ??= _sst.GetItems();
                return ReadXmlSharedStringCellValue(cellReader, useCachedFormulaResult, sharedStringItems);
            }

            bool numericAsDecimal = _opt.NumericAsDecimal;
            CultureInfo culture = _opt.Culture;
            bool useDateStyle = false;
            if (_opt.TreatDatesUsingNumberFormat && CellKindCanUseDateStyle(cellKind)) {
                string? styleAttribute = cellReader.GetAttribute("s");
                useDateStyle = IsDateStyleAttribute(styleAttribute);
            }

            int depth = cellReader.Depth;
            string? rawText = null;
            string? inlineText = null;
            string? formulaText = null;
            bool hasNode = cellReader.Read();
            while (hasNode) {
                if (cellReader.NodeType == XmlNodeType.EndElement && cellReader.Depth == depth && cellReader.LocalName == "c") {
                    break;
                }

                if (cellReader.NodeType == XmlNodeType.Element) {
                    if (cellReader.LocalName == "v") {
                        if (useCachedFormulaResult) {
                            rawText = ReadXmlValueTextAndSkipCell(cellReader, depth);
                        } else {
                            rawText = ReadXmlValueText(cellReader);
                        }

                        if (useCachedFormulaResult) {
                            if (!numericAsDecimal
                                && !useDateStyle
                                && (cellKind == XmlCellKind.Default || cellKind == XmlCellKind.Number)
                                && (TryParseInvariantDoubleFast(rawText, out double numericValue)
                                    || double.TryParse(rawText, NumberStyles.Float | NumberStyles.AllowThousands, CultureInfo.InvariantCulture, out numericValue))) {
                                return numericValue;
                            }

                            if (TryConvertXmlRawText(cellKind, rawText, useDateStyle, numericAsDecimal, culture, out object? fastValue)) {
                                return fastValue;
                            }
                        }

                        hasNode = true;
                        continue;
                    }

                    if (cellReader.LocalName == "f") {
                        formulaText = cellReader.ReadElementContentAsString();
                        if (!useCachedFormulaResult) {
                            SkipXmlElementContent(cellReader, depth);
                            return formulaText;
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

            if (formulaText != null && !useCachedFormulaResult) {
                return formulaText;
            }

            if (formulaText != null && rawText == null) {
                return formulaText;
            }

            if (cellKind == XmlCellKind.InlineString) {
                return inlineText;
            }

            if (cellKind == XmlCellKind.SharedString) {
                return TryParseSharedStringIndex(rawText, out int sstIndex) ? GetSharedString(sstIndex) : rawText;
            }

            if (cellKind == XmlCellKind.Boolean && rawText != null) {
                return rawText == "1";
            }

            if (cellKind == XmlCellKind.Date && rawText != null) {
                return DateTime.TryParse(rawText, culture, DateTimeStyles.AssumeLocal, out var date)
                    ? date
                    : rawText;
            }

            if (cellKind == XmlCellKind.String) {
                return rawText ?? inlineText;
            }

            if (rawText == null) {
                return inlineText;
            }

            if (useDateStyle
                && (TryParseInvariantDoubleFast(rawText, out double oa)
                    || double.TryParse(rawText, NumberStyles.Float | NumberStyles.AllowThousands, CultureInfo.InvariantCulture, out oa))) {
                return DateTime.FromOADate(oa);
            }

            if (numericAsDecimal
                && TryParseRawDecimal(rawText, culture, out decimal decimalNumber)) {
                return decimalNumber;
            }

            return (TryParseInvariantDoubleFast(rawText, out double number)
                    || double.TryParse(rawText, NumberStyles.Float | NumberStyles.AllowThousands, CultureInfo.InvariantCulture, out number))
                ? number
                : rawText;
        }

        private object? ReadXmlNumericCellValue(XmlReader cellReader) {
            bool useCachedFormulaResult = _opt.UseCachedFormulaResult;
            bool numericAsDecimal = _opt.NumericAsDecimal;
            CultureInfo culture = _opt.Culture;
            bool useDateStyle = _opt.TreatDatesUsingNumberFormat && IsDateStyleAttribute(cellReader.GetAttribute("s"));

            int depth = cellReader.Depth;
            string? rawText = null;
            string? inlineText = null;
            string? formulaText = null;
            bool hasNode = cellReader.Read();
            while (hasNode) {
                if (cellReader.NodeType == XmlNodeType.EndElement && cellReader.Depth == depth && cellReader.LocalName == "c") {
                    break;
                }

                if (cellReader.NodeType == XmlNodeType.Element) {
                    if (cellReader.LocalName == "v") {
                        rawText = useCachedFormulaResult
                            ? ReadXmlValueTextAndSkipCell(cellReader, depth)
                            : ReadXmlValueText(cellReader);

                        if (useCachedFormulaResult) {
                            if (!numericAsDecimal
                                && !useDateStyle
                                && (TryParseInvariantDoubleFast(rawText, out double numericValue)
                                    || double.TryParse(rawText, NumberStyles.Float | NumberStyles.AllowThousands, CultureInfo.InvariantCulture, out numericValue))) {
                                return numericValue;
                            }

                            if (useDateStyle
                                && (TryParseInvariantDoubleFast(rawText, out double oa)
                                    || double.TryParse(rawText, NumberStyles.Float | NumberStyles.AllowThousands, CultureInfo.InvariantCulture, out oa))) {
                                return DateTime.FromOADate(oa);
                            }

                            if (rawText == null) {
                                return null;
                            }

                            if (numericAsDecimal
                                && TryParseRawDecimal(rawText, culture, out decimal decimalNumber)) {
                                return decimalNumber;
                            }

                            return (TryParseInvariantDoubleFast(rawText, out double number)
                                    || double.TryParse(rawText, NumberStyles.Float | NumberStyles.AllowThousands, CultureInfo.InvariantCulture, out number))
                                ? number
                                : rawText;
                        }

                        hasNode = true;
                        continue;
                    }

                    if (cellReader.LocalName == "f") {
                        formulaText = cellReader.ReadElementContentAsString();
                        if (!useCachedFormulaResult) {
                            SkipXmlElementContent(cellReader, depth);
                            return formulaText;
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

            if (formulaText != null && !useCachedFormulaResult) {
                return formulaText;
            }

            if (formulaText != null && rawText == null) {
                return formulaText;
            }

            if (rawText == null) {
                return inlineText;
            }

            if (useDateStyle
                && (TryParseInvariantDoubleFast(rawText, out double oaValue)
                    || double.TryParse(rawText, NumberStyles.Float | NumberStyles.AllowThousands, CultureInfo.InvariantCulture, out oaValue))) {
                return DateTime.FromOADate(oaValue);
            }

            if (numericAsDecimal
                && TryParseRawDecimal(rawText, culture, out decimal rawDecimalNumber)) {
                return rawDecimalNumber;
            }

            return (TryParseInvariantDoubleFast(rawText, out double rawNumber)
                    || double.TryParse(rawText, NumberStyles.Float | NumberStyles.AllowThousands, CultureInfo.InvariantCulture, out rawNumber))
                ? rawNumber
                : rawText;
        }

        private bool TryConvertXmlRawText(
            XmlCellKind cellKind,
            string? rawText,
            bool useDateStyle,
            bool numericAsDecimal,
            CultureInfo culture,
            out object? value) {
            value = null;
            if (rawText == null) {
                return false;
            }

            switch (cellKind) {
                case XmlCellKind.SharedString:
                    value = TryParseSharedStringIndex(rawText, out int sstIndex) ? GetSharedString(sstIndex) : rawText;
                    return true;
                case XmlCellKind.Boolean:
                    value = rawText == "1";
                    return true;
                case XmlCellKind.Date:
                    value = DateTime.TryParse(rawText, culture, DateTimeStyles.AssumeLocal, out var date)
                        ? date
                        : rawText;
                    return true;
                case XmlCellKind.String:
                    value = rawText;
                    return true;
                case XmlCellKind.InlineString:
                    return false;
            }

            if (useDateStyle
                && (TryParseInvariantDoubleFast(rawText, out double oa)
                    || double.TryParse(rawText, NumberStyles.Float | NumberStyles.AllowThousands, CultureInfo.InvariantCulture, out oa))) {
                value = DateTime.FromOADate(oa);
                return true;
            }

            if (numericAsDecimal
                && TryParseRawDecimal(rawText, culture, out decimal decimalNumber)) {
                value = decimalNumber;
                return true;
            }

            value = (TryParseInvariantDoubleFast(rawText, out double number)
                    || double.TryParse(rawText, NumberStyles.Float | NumberStyles.AllowThousands, CultureInfo.InvariantCulture, out number))
                ? number
                : rawText;
            return true;
        }

        private object? ReadXmlSharedStringCellValue(XmlReader cellReader, bool useCachedFormulaResult, List<string> sharedStringItems) {
            int depth = cellReader.Depth;
            string? rawText = null;
            string? formulaText = null;
            bool hasNode = cellReader.Read();
            while (hasNode) {
                if (cellReader.NodeType == XmlNodeType.EndElement && cellReader.Depth == depth && cellReader.LocalName == "c") {
                    break;
                }

                if (cellReader.NodeType == XmlNodeType.Element) {
                    if (cellReader.LocalName == "v") {
                        if (useCachedFormulaResult) {
                            return ReadXmlSharedStringTextAndSkipCell(cellReader, depth, sharedStringItems);
                        }

                        bool parsedSharedStringIndex = TryReadXmlSharedStringIndexValue(cellReader, out int sstIndex, out rawText);
                        hasNode = true;
                        continue;
                    }

                    if (cellReader.LocalName == "f") {
                        formulaText = cellReader.ReadElementContentAsString();
                        if (!useCachedFormulaResult) {
                            SkipXmlElementContent(cellReader, depth);
                            return formulaText;
                        }

                        hasNode = true;
                        continue;
                    }

                    if (cellReader.LocalName == "is") {
                        _ = ReadXmlInlineString(cellReader);
                        hasNode = true;
                        continue;
                    }
                }

                hasNode = cellReader.Read();
            }

            if (formulaText != null && !useCachedFormulaResult) {
                return formulaText;
            }

            if (formulaText != null && rawText == null) {
                return formulaText;
            }

            return TryParseSharedStringIndex(rawText, out int index) ? GetSharedString(index, sharedStringItems) : rawText;
        }

        private static string? ReadXmlSharedStringTextAndSkipCell(XmlReader valueReader, int cellDepth, List<string> sharedStringItems) {
            if (valueReader.IsEmptyElement) {
                SkipXmlElementContent(valueReader, cellDepth);
                return string.Empty;
            }

            int valueDepth = valueReader.Depth;
            if (!valueReader.Read()) {
                return null;
            }

            if (valueReader.NodeType != XmlNodeType.Text
                && valueReader.NodeType != XmlNodeType.SignificantWhitespace
                && valueReader.NodeType != XmlNodeType.Whitespace) {
                SkipXmlElementContent(valueReader, cellDepth);
                return null;
            }

            string text = valueReader.Value;
            int parsed = 0;
            bool hasDigit = false;
            bool parsedFast = true;
            for (int i = 0; i < text.Length; i++) {
                int digit = text[i] - '0';
                if ((uint)digit > 9U || parsed > (int.MaxValue - digit) / 10) {
                    parsedFast = false;
                    break;
                }

                parsed = (parsed * 10) + digit;
                hasDigit = true;
            }

            bool completedCell = valueReader.Read()
                && valueReader.NodeType == XmlNodeType.EndElement
                && valueReader.Depth == valueDepth
                && valueReader.Read()
                && valueReader.NodeType == XmlNodeType.EndElement
                && valueReader.Depth == cellDepth;
            if (!completedCell) {
                SkipXmlElementContent(valueReader, cellDepth);
            }

            if (parsedFast && hasDigit) {
                return GetSharedString(parsed, sharedStringItems);
            }

            return int.TryParse(text, NumberStyles.Integer, CultureInfo.InvariantCulture, out int index)
                ? GetSharedString(index, sharedStringItems)
                : text;
        }

        private static string? ReadXmlValueText(XmlReader valueReader) {
            if (valueReader.IsEmptyElement) {
                return string.Empty;
            }

            int depth = valueReader.Depth;
            if (!valueReader.Read()) {
                return null;
            }

            if (valueReader.NodeType != XmlNodeType.Text
                && valueReader.NodeType != XmlNodeType.SignificantWhitespace
                && valueReader.NodeType != XmlNodeType.Whitespace) {
                SkipXmlElementContent(valueReader, depth);
                return null;
            }

            string text = valueReader.Value;
            SkipXmlElementContent(valueReader, depth);
            return text;
        }

        private static string? ReadXmlValueTextAndSkipCell(XmlReader valueReader, int cellDepth) {
            if (valueReader.IsEmptyElement) {
                SkipXmlElementContent(valueReader, cellDepth);
                return string.Empty;
            }

            int valueDepth = valueReader.Depth;
            if (!valueReader.Read()) {
                return null;
            }

            if (valueReader.NodeType != XmlNodeType.Text
                && valueReader.NodeType != XmlNodeType.SignificantWhitespace
                && valueReader.NodeType != XmlNodeType.Whitespace) {
                SkipXmlElementContent(valueReader, cellDepth);
                return null;
            }

            string text = valueReader.Value;
            if (valueReader.Read()
                && valueReader.NodeType == XmlNodeType.EndElement
                && valueReader.Depth == valueDepth
                && valueReader.Read()
                && valueReader.NodeType == XmlNodeType.EndElement
                && valueReader.Depth == cellDepth) {
                return text;
            }

            SkipXmlElementContent(valueReader, cellDepth);
            return text;
        }

        private static bool TryReadXmlSharedStringIndexValue(XmlReader valueReader, out int index, out string? rawText) {
            index = 0;
            rawText = null;

            if (valueReader.IsEmptyElement) {
                rawText = string.Empty;
                return false;
            }

            int depth = valueReader.Depth;
            if (!valueReader.Read()) {
                return false;
            }

            if (valueReader.NodeType != XmlNodeType.Text
                && valueReader.NodeType != XmlNodeType.SignificantWhitespace
                && valueReader.NodeType != XmlNodeType.Whitespace) {
                SkipXmlElementContent(valueReader, depth);
                return false;
            }

            string text = valueReader.Value;
            rawText = text;
            int parsed = 0;
            bool hasDigit = false;
            for (int i = 0; i < text.Length; i++) {
                int digit = text[i] - '0';
                if ((uint)digit > 9U) {
                    SkipXmlElementContent(valueReader, depth);
                    return int.TryParse(text, NumberStyles.Integer, CultureInfo.InvariantCulture, out index);
                }

                if (parsed > (int.MaxValue - digit) / 10) {
                    SkipXmlElementContent(valueReader, depth);
                    return int.TryParse(text, NumberStyles.Integer, CultureInfo.InvariantCulture, out index);
                }

                parsed = (parsed * 10) + digit;
                hasDigit = true;
            }

            SkipXmlElementContent(valueReader, depth);
            if (!hasDigit) {
                return false;
            }

            index = parsed;
            return true;
        }

        private static bool TryReadXmlSharedStringIndexValueAndSkipCell(XmlReader valueReader, int cellDepth, out int index, out string? rawText) {
            index = 0;
            rawText = null;

            if (valueReader.IsEmptyElement) {
                SkipXmlElementContent(valueReader, cellDepth);
                rawText = string.Empty;
                return false;
            }

            int valueDepth = valueReader.Depth;
            if (!valueReader.Read()) {
                return false;
            }

            if (valueReader.NodeType != XmlNodeType.Text
                && valueReader.NodeType != XmlNodeType.SignificantWhitespace
                && valueReader.NodeType != XmlNodeType.Whitespace) {
                SkipXmlElementContent(valueReader, cellDepth);
                return false;
            }

            string text = valueReader.Value;
            rawText = text;
            int parsed = 0;
            bool hasDigit = false;
            for (int i = 0; i < text.Length; i++) {
                int digit = text[i] - '0';
                if ((uint)digit > 9U) {
                    SkipXmlElementContent(valueReader, cellDepth);
                    return int.TryParse(text, NumberStyles.Integer, CultureInfo.InvariantCulture, out index);
                }

                if (parsed > (int.MaxValue - digit) / 10) {
                    SkipXmlElementContent(valueReader, cellDepth);
                    return int.TryParse(text, NumberStyles.Integer, CultureInfo.InvariantCulture, out index);
                }

                parsed = (parsed * 10) + digit;
                hasDigit = true;
            }

            if (valueReader.Read()
                && valueReader.NodeType == XmlNodeType.EndElement
                && valueReader.Depth == valueDepth
                && valueReader.Read()
                && valueReader.NodeType == XmlNodeType.EndElement
                && valueReader.Depth == cellDepth) {
                if (!hasDigit) {
                    return false;
                }

                index = parsed;
                return true;
            }

            SkipXmlElementContent(valueReader, cellDepth);
            if (!hasDigit) {
                return false;
            }

            index = parsed;
            return true;
        }

        private static string ReadXmlInlineString(XmlReader inlineReader) {
            if (inlineReader.IsEmptyElement) {
                return string.Empty;
            }

            int depth = inlineReader.Depth;
            string? first = null;
            System.Text.StringBuilder? builder = null;
            while (inlineReader.Read()) {
                if (inlineReader.NodeType == XmlNodeType.EndElement && inlineReader.Depth == depth && inlineReader.LocalName == "is") {
                    break;
                }

                if (inlineReader.NodeType != XmlNodeType.Element || inlineReader.LocalName != "t") {
                    continue;
                }

                string text = inlineReader.ReadElementContentAsString();
                if (builder != null) {
                    builder.Append(text);
                } else if (first == null) {
                    first = text;
                } else {
                    builder = new System.Text.StringBuilder(first.Length + text.Length);
                    builder.Append(first);
                    builder.Append(text);
                }
            }

            return builder?.ToString() ?? first ?? string.Empty;
        }

        private static int ParsePositiveIntAttribute(string? value) {
            if (string.IsNullOrEmpty(value)) {
                return 0;
            }

            string text = value!;
            int result = 0;
            for (int i = 0; i < text.Length; i++) {
                int digit = text[i] - '0';
                if ((uint)digit > 9U) {
                    return 0;
                }

                if (result > (int.MaxValue - digit) / 10) {
                    return 0;
                }

                result = (result * 10) + digit;
            }

            return result;
        }

        private static bool TryParseUInt(string? value, out uint result) {
            result = 0;
            if (string.IsNullOrEmpty(value)) {
                return false;
            }

            string text = value!;
            uint parsed = 0;
            for (int i = 0; i < text.Length; i++) {
                uint digit = (uint)(text[i] - '0');
                if (digit > 9U) {
                    return uint.TryParse(text, NumberStyles.Integer, CultureInfo.InvariantCulture, out result);
                }

                if (parsed > (uint.MaxValue - digit) / 10U) {
                    return uint.TryParse(text, NumberStyles.Integer, CultureInfo.InvariantCulture, out result);
                }

                parsed = (parsed * 10U) + digit;
            }

            result = parsed;
            return true;
        }
    }
}
