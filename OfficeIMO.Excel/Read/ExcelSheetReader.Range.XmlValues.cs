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
    internal sealed partial class ExcelSheetReader {
        private enum XmlDataReaderTargetKind : byte {
            None,
            Numeric,
            DateTime,
            Boolean,
            String
        }

        private enum XmlDataReaderPrimitiveKind : byte {
            None,
            Double,
            Decimal,
            CalendarDateSerial,
            ElapsedDateSerial,
            DateTime,
            Boolean
        }

        private object? ReadXmlCellValue(XmlReader cellReader) {
            return ReadXmlCellValue(cellReader, ReadXmlCellTypeAttribute(cellReader));
        }

        private object? ReadXmlCellValue(XmlReader cellReader, string? cellType, bool preserveDateSerial = false) {
            if (cellType == "e" && cellReader.GetAttribute("vm") != null) {
                var raw = ReadXmlCellRaw(cellReader, 0, 0, ParseXmlCellKind(cellType), readStyleIndex: true);
                return ConvertRaw(raw).TypedValue;
            }
            if (cellReader.IsEmptyElement) {
                return null;
            }
            if (preserveDateSerial && _opt.TreatDatesUsingNumberFormat
                && CellKindCanUseDateStyle(ParseXmlCellKind(cellType))
                && IsDateStyleAttribute(ReadXmlStyleAttribute(cellReader))) {
                return ConvertRawForDataReader(ReadXmlCellRaw(cellReader, 0, 0, ParseXmlCellKind(cellType), readStyleIndex: true));
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
            bool calendarStyle = false;
            if (_opt.TreatDatesUsingNumberFormat && CellKindCanUseDateStyle(cellKind)) {
                XmlStyleAttribute styleAttribute = ReadXmlStyleAttribute(cellReader);
                useDateStyle = IsDateStyleAttribute(styleAttribute);
                calendarStyle = useDateStyle && IsCalendarStyleAttribute(styleAttribute);
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
                                && (TryParseInvariantDouble(rawText, out double numericValue)
                                    || double.TryParse(rawText, NumberStyles.Float | NumberStyles.AllowThousands, CultureInfo.InvariantCulture, out numericValue))) {
                                return numericValue;
                            }

                            if (TryConvertXmlRawText(cellKind, rawText, useDateStyle, calendarStyle, numericAsDecimal, culture, out object? fastValue)) {
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
                && (TryParseInvariantDouble(rawText, out double oa)
                    || double.TryParse(rawText, NumberStyles.Float | NumberStyles.AllowThousands, CultureInfo.InvariantCulture, out oa))) {
                return FromExcelSerialDate(oa, calendarStyle);
            }

            if (numericAsDecimal
                && TryParseExcelNumberAsDecimal(rawText, culture, out decimal decimalNumber)) {
                return decimalNumber;
            }

            return (TryParseInvariantDouble(rawText, out double number)
                    || double.TryParse(rawText, NumberStyles.Float | NumberStyles.AllowThousands, CultureInfo.InvariantCulture, out number))
                ? number
                : rawText;
        }

        private object? ReadXmlNumericCellValue(XmlReader cellReader) {
            bool useCachedFormulaResult = _opt.UseCachedFormulaResult;
            bool numericAsDecimal = _opt.NumericAsDecimal;
            CultureInfo culture = _opt.Culture;
            XmlStyleAttribute style = _opt.TreatDatesUsingNumberFormat ? ReadXmlStyleAttribute(cellReader) : default;
            bool useDateStyle = IsDateStyleAttribute(style);
            bool calendarStyle = useDateStyle && IsCalendarStyleAttribute(style);

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
                                if (useDateStyle) return FromExcelSerialDate(simpleNumber, calendarStyle);
                                if (numericAsDecimal && TryConvertExcelNumberToDecimal(simpleNumber, out decimal simpleDecimal)) {
                                    return simpleDecimal;
                                }
                                return simpleNumber;
                            }
                        } else {
                            rawText = ReadXmlValueText(cellReader);
                        }

                        if (useCachedFormulaResult) {
                            if (!numericAsDecimal
                                && !useDateStyle
                                && (TryParseInvariantDouble(rawText, out double numericValue)
                                    || double.TryParse(rawText, NumberStyles.Float | NumberStyles.AllowThousands, CultureInfo.InvariantCulture, out numericValue))) {
                                return numericValue;
                            }

                            if (useDateStyle
                                && (TryParseInvariantDouble(rawText, out double oa)
                                    || double.TryParse(rawText, NumberStyles.Float | NumberStyles.AllowThousands, CultureInfo.InvariantCulture, out oa))) {
                                return FromExcelSerialDate(oa, calendarStyle);
                            }

                            if (rawText == null) {
                                return null;
                            }

                            if (numericAsDecimal
                                && TryParseExcelNumberAsDecimal(rawText, culture, out decimal decimalNumber)) {
                                return decimalNumber;
                            }

                            return (TryParseInvariantDouble(rawText, out double number)
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
                && (TryParseInvariantDouble(rawText, out double oaValue)
                    || double.TryParse(rawText, NumberStyles.Float | NumberStyles.AllowThousands, CultureInfo.InvariantCulture, out oaValue))) {
                return FromExcelSerialDate(oaValue, calendarStyle);
            }

            if (numericAsDecimal
                && TryParseExcelNumberAsDecimal(rawText, culture, out decimal rawDecimalNumber)) {
                return rawDecimalNumber;
            }

            return (TryParseInvariantDouble(rawText, out double rawNumber)
                    || double.TryParse(rawText, NumberStyles.Float | NumberStyles.AllowThousands, CultureInfo.InvariantCulture, out rawNumber))
                ? rawNumber
                : rawText;
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
            bool calendarStyle,
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
                && (TryParseInvariantDouble(rawText, out double oa)
                    || double.TryParse(rawText, NumberStyles.Float | NumberStyles.AllowThousands, CultureInfo.InvariantCulture, out oa))) {
                value = FromExcelSerialDate(oa, calendarStyle);
                return true;
            }

            if (numericAsDecimal
                && TryParseExcelNumberAsDecimal(rawText, culture, out decimal decimalNumber)) {
                value = decimalNumber;
                return true;
            }

            value = (TryParseInvariantDouble(rawText, out double number)
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

        private static int ParsePositiveIntAttribute(ReadOnlySpan<char> value) {
            if (value.IsEmpty) {
                return 0;
            }

            ReadOnlySpan<char> text = value;
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

    }
}
