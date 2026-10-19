using DocumentFormat.OpenXml.Spreadsheet;
using System.Globalization;
using OfficeIMO.Excel.LegacyXls.Biff;
using OfficeIMO.Excel.LegacyXls.Model;

namespace OfficeIMO.Excel.LegacyXls.Write;

internal static partial class LegacyXlsWriter {
    /// <summary>Encodes only the cached result while retaining the original formula tokens and native group definitions.</summary>
    internal static IReadOnlyList<BiffRecord> EncodePreservedFormulaCache(ExcelSheet sheet, Cell cell, IReadOnlyList<BiffRecord> original,
        ExcelDateSystem dateSystem, IReadOnlyDictionary<uint, uint?>? names) {
        byte[] payload = original[0].Payload;
        if (original[0].Type != 0x0006 || payload.Length < 22) throw new InvalidDataException("A retained formula cache has no complete Formula record.");
        int count = BiffRecordReader.ReadUInt16(payload, 20);
        if (count > payload.Length - 22) throw new InvalidDataException("The retained Formula token range is truncated.");
        byte[] tokens = payload.Skip(22).Take(count).ToArray(), extra = payload.Skip(22 + count).ToArray();
        if (names?.Any(item => item.Key != item.Value) == true) BiffFormulaTokenScanner.RewriteNames(tokens, 0, tokens.Length, names);
        ushort row = BiffRecordReader.ReadUInt16(payload, 0), column = BiffRecordReader.ReadUInt16(payload, 2), style = BiffRecordReader.ReadUInt16(payload, 4);
        var converted = ConvertCachedFormulaCell(sheet, cell, row, column, style, dateSystem, tokens, extra, null, null,
            cell.CellReference?.Value ?? string.Empty, false);
        using var output = new MemoryStream();
        WriteCellRecords(output, 0, new[] { converted }, new LegacyXlsSharedStringTable(Array.Empty<LegacyXlsSharedStringEntry>(), 0));
        var result = BiffRecordReader.ReadRecords(output.ToArray(), new()).ToList();
        result[0].Payload[14] = payload[14]; result[0].Payload[15] = payload[15];
        bool stringContinuation = false;
        foreach (var record in original.Skip(1)) {
            if (record.Type == 0x0207) { stringContinuation = true; continue; }
            if (stringContinuation && record.Type == 0x003c) continue;
            stringContinuation = false; result.Add(record);
        }
        return result;
    }

        private static LegacyXlsCell ConvertCachedFormulaCell(
            ExcelSheet sheet,
            Cell cell,
            ushort row,
            ushort column,
            ushort styleIndex,
            ExcelDateSystem dateSystem,
            byte[] tokens,
            byte[] formulaExtraData,
            byte[]? arrayFormulaPayload,
            byte[]? sharedFormulaPayload,
            string address,
            bool isArrayFormula) {
            if (cell.CellValue == null && cell.InlineString == null)
                return LegacyXlsCell.FormulaEmpty(row, column, styleIndex, tokens, formulaExtraData, arrayFormulaPayload, sharedFormulaPayload);
            DocumentFormat.OpenXml.Spreadsheet.CellValues? dataType = cell.DataType?.Value;
            string rawValue = cell.CellValue?.InnerText ?? string.Empty;
            if (dataType == DocumentFormat.OpenXml.Spreadsheet.CellValues.Boolean) {
                if (string.IsNullOrWhiteSpace(rawValue)) {
                    throw new NotSupportedException($"Native XLS saving requires a cached result for formula cell {address}. Enable formula evaluation before saving or save as .xlsx.");
                }

                bool value = rawValue == "1" || rawValue.Equals("true", StringComparison.OrdinalIgnoreCase);
                return LegacyXlsCell.FormulaBoolean(row, column, styleIndex, value, tokens, formulaExtraData, arrayFormulaPayload, sharedFormulaPayload);
            }

            if (dataType == DocumentFormat.OpenXml.Spreadsheet.CellValues.String
                || dataType == DocumentFormat.OpenXml.Spreadsheet.CellValues.SharedString
                || dataType == DocumentFormat.OpenXml.Spreadsheet.CellValues.InlineString) {
                string text = sheet.GetCellText(cell);
                EnsureSupportedFormulaTextLength(text, address);
                return LegacyXlsCell.FormulaText(
                    row,
                    column,
                    styleIndex,
                    text,
                    tokens,
                    formulaExtraData,
                    arrayFormulaPayload: arrayFormulaPayload,
                    sharedFormulaPayload: sharedFormulaPayload);
            }

            if (dataType == DocumentFormat.OpenXml.Spreadsheet.CellValues.Error) {
                if (string.IsNullOrWhiteSpace(rawValue)) {
                    throw new NotSupportedException($"Native XLS saving requires a cached result for formula cell {address}. Enable formula evaluation before saving or save as .xlsx.");
                }

                if (ExcelErrorCode.TryGetCode(rawValue, out byte errorCode)) {
                    return LegacyXlsCell.FormulaError(row, column, styleIndex, errorCode, tokens, formulaExtraData, arrayFormulaPayload, sharedFormulaPayload);
                }

                throw new NotSupportedException($"Native XLS saving does not support cached error result '{rawValue}' for formula cell {address}.");
            }

            if (dataType == DocumentFormat.OpenXml.Spreadsheet.CellValues.Date) {
                if (string.IsNullOrWhiteSpace(rawValue)) {
                    throw new NotSupportedException($"Native XLS saving requires a cached result for formula cell {address}. Enable formula evaluation before saving or save as .xlsx.");
                }

                if (TryParseOpenXmlDate(rawValue, out DateTime date)) {
                    return LegacyXlsCell.FormulaNumber(row, column, styleIndex, ExcelDateSystemConverter.ToSerial(date, dateSystem), tokens, formulaExtraData, arrayFormulaPayload, sharedFormulaPayload);
                }

                throw new NotSupportedException($"Native XLS saving does not support cached date result '{rawValue}' for formula cell {address}.");
            }

            if ((dataType == DocumentFormat.OpenXml.Spreadsheet.CellValues.Number || dataType == null)
                && !string.IsNullOrWhiteSpace(rawValue)
                && double.TryParse(rawValue, NumberStyles.Float, CultureInfo.InvariantCulture, out double number)) {
                return LegacyXlsCell.FormulaNumber(row, column, styleIndex, number, tokens, formulaExtraData, arrayFormulaPayload, sharedFormulaPayload);
            }

            throw new NotSupportedException($"Native XLS saving requires a numeric, Boolean, text, error, or date cached result for formula cell {address}. Enable formula evaluation before saving or save as .xlsx.");
        }

}
