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
        private enum XmlDataReaderTargetKind : byte {
            None,
            Int32,
            Double,
            DateTime,
            Boolean,
            String
        }

        private enum XmlDataReaderPrimitiveKind : byte {
            None,
            Double,
            DateTime,
            Boolean
        }

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
                return BoxBoolean(rawText == "1");
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
                return FromExcelSerialDate(oa);
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
                        if (useCachedFormulaResult && !numericAsDecimal) {
                            if (TryReadXmlSimpleDoubleAndSkipCell(cellReader, depth, out double simpleNumber, out rawText)) {
                                return useDateStyle ? FromExcelSerialDate(simpleNumber) : simpleNumber;
                            }
                        } else {
                            rawText = useCachedFormulaResult
                                ? ReadXmlValueTextAndSkipCell(cellReader, depth)
                                : ReadXmlValueText(cellReader);
                        }

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
                                return FromExcelSerialDate(oa);
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
                return FromExcelSerialDate(oaValue);
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

        private bool TryReadXmlCellPrimitiveForDataReader(
            XmlReader cellReader,
            string? cellType,
            XmlDataReaderTargetKind targetKind,
            out XmlDataReaderPrimitiveKind primitiveKind,
            out double doubleValue,
            out DateTime dateTimeValue,
            out bool booleanValue,
            out object? objectValue) {
            primitiveKind = XmlDataReaderPrimitiveKind.None;
            doubleValue = 0;
            dateTimeValue = default;
            booleanValue = false;
            objectValue = null;

            if (_opt.CellValueConverter != null || cellReader.IsEmptyElement) {
                return false;
            }

            XmlCellKind cellKind = ParseXmlCellKind(cellType);
            if ((cellKind == XmlCellKind.Default || cellKind == XmlCellKind.Number)
                && !_opt.NumericAsDecimal) {
                bool useDateStyle = _opt.TreatDatesUsingNumberFormat && IsDateStyleAttribute(cellReader.GetAttribute("s"));
                if ((targetKind == XmlDataReaderTargetKind.Int32 || targetKind == XmlDataReaderTargetKind.Double) && !useDateStyle) {
                    return TryReadXmlNumericPrimitiveForDataReader(
                        cellReader,
                        asDate: false,
                        out primitiveKind,
                        out doubleValue,
                        out dateTimeValue,
                        out objectValue);
                }

                if (targetKind == XmlDataReaderTargetKind.DateTime && useDateStyle) {
                    return TryReadXmlNumericPrimitiveForDataReader(
                        cellReader,
                        asDate: true,
                        out primitiveKind,
                        out doubleValue,
                        out dateTimeValue,
                        out objectValue);
                }
            }

            if (cellKind == XmlCellKind.Boolean && targetKind == XmlDataReaderTargetKind.Boolean) {
                return TryReadXmlBooleanPrimitiveForDataReader(cellReader, out primitiveKind, out booleanValue, out objectValue);
            }

            return false;
        }

        private bool TryReadXmlNumericPrimitiveForDataReader(
            XmlReader cellReader,
            bool asDate,
            out XmlDataReaderPrimitiveKind primitiveKind,
            out double doubleValue,
            out DateTime dateTimeValue,
            out object? objectValue) {
            primitiveKind = XmlDataReaderPrimitiveKind.None;
            doubleValue = 0;
            dateTimeValue = default;
            objectValue = null;

            bool useCachedFormulaResult = _opt.UseCachedFormulaResult;
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
                            if (TryReadXmlSimpleDoubleAndSkipCell(cellReader, depth, out double simpleNumber, out rawText)) {
                                if (asDate) {
                                    primitiveKind = XmlDataReaderPrimitiveKind.DateTime;
                                    dateTimeValue = FromExcelSerialDate(simpleNumber);
                                } else {
                                    primitiveKind = XmlDataReaderPrimitiveKind.Double;
                                    doubleValue = simpleNumber;
                                }

                                return true;
                            }
                        } else {
                            rawText = ReadXmlValueText(cellReader);
                        }

                        hasNode = true;
                        continue;
                    }

                    if (cellReader.LocalName == "f") {
                        formulaText = cellReader.ReadElementContentAsString();
                        if (!useCachedFormulaResult) {
                            SkipXmlElementContent(cellReader, depth);
                            objectValue = formulaText;
                            return true;
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
                objectValue = formulaText;
                return true;
            }

            if (formulaText != null && rawText == null) {
                objectValue = formulaText;
                return true;
            }

            if (rawText == null) {
                objectValue = inlineText;
                return true;
            }

            if (TryParseInvariantDoubleFast(rawText, out double number)
                || double.TryParse(rawText, NumberStyles.Float | NumberStyles.AllowThousands, CultureInfo.InvariantCulture, out number)) {
                if (asDate) {
                    primitiveKind = XmlDataReaderPrimitiveKind.DateTime;
                    dateTimeValue = FromExcelSerialDate(number);
                } else {
                    primitiveKind = XmlDataReaderPrimitiveKind.Double;
                    doubleValue = number;
                }

                return true;
            }

            objectValue = rawText;
            return true;
        }

        private bool TryReadXmlBooleanPrimitiveForDataReader(
            XmlReader cellReader,
            out XmlDataReaderPrimitiveKind primitiveKind,
            out bool booleanValue,
            out object? objectValue) {
            primitiveKind = XmlDataReaderPrimitiveKind.None;
            booleanValue = false;
            objectValue = null;

            bool useCachedFormulaResult = _opt.UseCachedFormulaResult;
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
                            primitiveKind = XmlDataReaderPrimitiveKind.Boolean;
                            booleanValue = ReadXmlBooleanValueAndSkipCell(cellReader, depth, out rawText);
                            return true;
                        }

                        rawText = ReadXmlValueText(cellReader);
                        hasNode = true;
                        continue;
                    }

                    if (cellReader.LocalName == "f") {
                        formulaText = cellReader.ReadElementContentAsString();
                        if (!useCachedFormulaResult) {
                            SkipXmlElementContent(cellReader, depth);
                            objectValue = formulaText;
                            return true;
                        }

                        hasNode = true;
                        continue;
                    }
                }

                hasNode = cellReader.Read();
            }

            if (formulaText != null && rawText == null) {
                objectValue = formulaText;
                return true;
            }

            if (rawText == null) {
                return true;
            }

            primitiveKind = XmlDataReaderPrimitiveKind.Boolean;
            booleanValue = rawText == "1";
            return true;
        }

        private bool ReadXmlBooleanValueAndSkipCell(XmlReader valueReader, int cellDepth, out string? rawText) {
            if (!TryReadXmlBufferedValueTextAndSkipCell(valueReader, cellDepth, out char[] buffer, out int length, out rawText)) {
                return rawText == "1";
            }

            if (rawText == null) {
                return length == 1 && buffer[0] == '1';
            }

            return rawText == "1";
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
                    value = BoxBoolean(rawText == "1");
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
                value = FromExcelSerialDate(oa);
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

        private string? ReadXmlSharedStringTextAndSkipCell(XmlReader valueReader, int cellDepth, List<string> sharedStringItems) {
            if (!TryReadXmlBufferedValueTextAndSkipCell(valueReader, cellDepth, out char[] buffer, out int length, out string? rawText)) {
                return rawText;
            }

            if (rawText == null) {
                if (TryParseSharedStringIndex(buffer.AsSpan(0, length), out int parsed)) {
                    return GetSharedString(parsed, sharedStringItems);
                }

                rawText = new string(buffer, 0, length);
                return TryParseSharedStringIndex(rawText, out parsed)
                    ? GetSharedString(parsed, sharedStringItems)
                    : rawText;
            }

            return TryParseSharedStringIndex(rawText, out int index)
                ? GetSharedString(index, sharedStringItems)
                : rawText;
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

        private bool TryReadXmlSimpleDoubleAndSkipCell(XmlReader valueReader, int cellDepth, out double value, out string? rawText) {
            value = 0;
            if (!TryReadXmlBufferedValueTextAndSkipCell(valueReader, cellDepth, out char[] buffer, out int length, out rawText)) {
                return false;
            }

            if (rawText == null) {
                if (TryParseInvariantDoubleFast(buffer.AsSpan(0, length), out value)) {
                    return true;
                }

                rawText = new string(buffer, 0, length);
                return false;
            }

            return TryParseInvariantDoubleFast(rawText, out value);
        }

        private bool TryReadXmlBufferedValueTextAndSkipCell(XmlReader valueReader, int cellDepth, out char[] buffer, out int length, out string? rawText) {
            buffer = _xmlValueTextBuffer ??= new char[64];
            length = 0;
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

            if (!IsXmlTextNode(valueReader.NodeType)) {
                SkipXmlElementContent(valueReader, cellDepth);
                return false;
            }

            System.Text.StringBuilder? builder = null;

            while (true) {
                int read = valueReader.ReadValueChunk(buffer, length, buffer.Length - length);
                if (read == 0) {
                    break;
                }

                length += read;
                if (length != buffer.Length) {
                    continue;
                }

                builder = new System.Text.StringBuilder(buffer.Length * 2);
                builder.Append(buffer, 0, length);
                length = 0;

                while (true) {
                    read = valueReader.ReadValueChunk(buffer, 0, buffer.Length);
                    if (read == 0) {
                        break;
                    }

                    builder.Append(buffer, 0, read);
                }

                break;
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

            if (builder == null) {
                return true;
            }

            rawText = builder.ToString();
            return true;
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

                string text = ReadXmlTextElement(inlineReader);
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

        private static string ReadXmlTextElement(XmlReader textReader) {
            if (textReader.IsEmptyElement) {
                return string.Empty;
            }

            int depth = textReader.Depth;
            string? first = null;
            System.Text.StringBuilder? builder = null;
            while (textReader.Read()) {
                if (textReader.NodeType == XmlNodeType.EndElement && textReader.Depth == depth && textReader.LocalName == "t") {
                    break;
                }

                if (!IsXmlTextNode(textReader.NodeType)) {
                    continue;
                }

                string text = textReader.Value;
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

        private static bool IsXmlTextNode(XmlNodeType nodeType) {
            return nodeType == XmlNodeType.Text
                || nodeType == XmlNodeType.CDATA
                || nodeType == XmlNodeType.SignificantWhitespace
                || nodeType == XmlNodeType.Whitespace;
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
