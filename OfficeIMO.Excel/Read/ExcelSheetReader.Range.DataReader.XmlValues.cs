#nullable enable

using System.Globalization;
using System.Xml;

namespace OfficeIMO.Excel {
    internal sealed partial class ExcelSheetReader {
        private bool TryReadXmlCellPrimitiveForDataReader(
            XmlReader cellReader,
            string? cellType,
            XmlDataReaderTargetKind targetKind,
            out XmlDataReaderPrimitiveKind primitiveKind,
            out double doubleValue,
            out decimal decimalValue,
            out bool booleanValue,
            out object? objectValue) {
            primitiveKind = XmlDataReaderPrimitiveKind.None;
            doubleValue = 0;
            decimalValue = default;
            booleanValue = false;
            objectValue = null;

            if (_opt.CellValueConverter != null || cellReader.IsEmptyElement) {
                return false;
            }

            XmlCellKind cellKind = ParseXmlCellKind(cellType);
            if (cellKind == XmlCellKind.Default || cellKind == XmlCellKind.Number) {
                string? style = cellReader.GetAttribute("s");
                bool useDateStyle = _opt.TreatDatesUsingNumberFormat && IsDateStyleAttribute(style);
                if (targetKind == XmlDataReaderTargetKind.Numeric
                    || (targetKind == XmlDataReaderTargetKind.DateTime && useDateStyle)) {
                    return TryReadXmlNumericPrimitiveForDataReader(
                        cellReader,
                        useDateStyle,
                        useDateStyle && IsCalendarStyleAttribute(style),
                        out primitiveKind,
                        out doubleValue,
                        out decimalValue,
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
            bool dateStyle,
            bool calendarStyle,
            out XmlDataReaderPrimitiveKind primitiveKind,
            out double doubleValue,
            out decimal decimalValue,
            out object? objectValue) {
            primitiveKind = XmlDataReaderPrimitiveKind.None;
            doubleValue = 0;
            decimalValue = default;
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
                                doubleValue = simpleNumber;
                                primitiveKind = ClassifyDataReaderNumber(simpleNumber, dateStyle, calendarStyle, out decimalValue);

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

            if (TryParseInvariantDouble(rawText, out double number)
                || double.TryParse(rawText, NumberStyles.Float | NumberStyles.AllowThousands, CultureInfo.InvariantCulture, out number)) {
                doubleValue = number;
                primitiveKind = ClassifyDataReaderNumber(number, dateStyle, calendarStyle, out decimalValue);

                return true;
            }

            if (!dateStyle && _opt.NumericAsDecimal
                && TryParseExcelNumberAsDecimal(rawText, _opt.Culture, out decimalValue)) {
                primitiveKind = XmlDataReaderPrimitiveKind.Decimal;
                return true;
            }

            objectValue = rawText;
            return true;
        }

        private XmlDataReaderPrimitiveKind ClassifyDataReaderNumber(
            double number, bool dateStyle, bool calendarStyle, out decimal decimalValue) {
            decimalValue = default;
            if (dateStyle) {
                // Keep the serial until a date is requested. Numeric access must also
                // work for serials outside DateTime's range and after date access.
                return calendarStyle ? XmlDataReaderPrimitiveKind.CalendarDateSerial : XmlDataReaderPrimitiveKind.ElapsedDateSerial;
            }
            return _opt.NumericAsDecimal && TryConvertExcelNumberToDecimal(number, out decimalValue)
                ? XmlDataReaderPrimitiveKind.Decimal : XmlDataReaderPrimitiveKind.Double;
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

    }
}
