using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading;

namespace OfficeIMO.Excel.Legacy;

internal enum WkRecordLayout {
    Dos,
    QuattroWq2
}

internal abstract class WkRecordSpreadsheetAdapterBase : LegacySpreadsheetAdapterBase {
    protected LegacySpreadsheetModel ParseWkRecords(byte[] data, OfficeLegacyImportLimits limits, string familyName,
        byte expectedProduct0, byte expectedProduct1, CancellationToken cancellationToken,
        WkRecordLayout layout = WkRecordLayout.Dos, bool translateFormulas = true) {
        ValidateBof(data, familyName, expectedProduct0, expectedProduct1);
        var model = new LegacySpreadsheetModel { Quality = OfficeLegacyImportQuality.Structured };
        var sheets = new Dictionary<byte, LegacySpreadsheetSheet>();
        int offset = 0;
        int records = 0;
        int recoveredTextCharacters = 0;
        bool foundEnd = false;
        bool reportedUnsupportedFormula = false;
        bool reportedUnsupportedFormat = false;
        bool reportedStringFormula = false;
        var unsupportedRecordTypes = new HashSet<ushort>();
        while (offset + 4 <= data.Length) {
            cancellationToken.ThrowIfCancellationRequested();
            if (++records > limits.MaxRecords) throw new InvalidDataException("Legacy spreadsheet exceeds the configured record limit.");
            ushort type = OfficeLegacyImportBuffer.ReadUInt16(data, offset);
            ushort length = OfficeLegacyImportBuffer.ReadUInt16(data, offset + 2);
            int payload = offset + 4;
            if (payload > data.Length - length) throw new InvalidDataException($"Truncated {familyName} record 0x{type:X4}.");

            switch (type) {
                case 0x0000:
                    if (offset != 0) throw new InvalidDataException($"The {familyName} record stream contains a duplicate BOF record.");
                    model.Metadata["BofPayload"] = ToHex(data, payload, Math.Min(length, (ushort)8));
                    break;
                case 0x0001:
                    foundEnd = true;
                    break;
                case 0x000B:
                    CaptureName(model, sheets, data, payload, length, limits, ref recoveredTextCharacters);
                    break;
                case 0x000C:
                    ValidateCellHeader(data, payload, length, layout);
                    AddCell(model, sheets, limits, data, payload, null, null, null, ref reportedUnsupportedFormat, layout: layout);
                    break;
                case 0x000D:
                    if (length < DataOffset(layout) + 2) throw new InvalidDataException("Truncated WK integer cell record.");
                    AddCell(model, sheets, limits, data, payload, OfficeLegacyImportBuffer.ReadInt16(data, payload + DataOffset(layout)), null, null, ref reportedUnsupportedFormat, layout: layout);
                    break;
                case 0x000E:
                    if (length < DataOffset(layout) + 8) throw new InvalidDataException("Truncated WK floating-point cell record.");
                    AddCell(model, sheets, limits, data, payload, ReadDouble(data, payload + DataOffset(layout)), null, null, ref reportedUnsupportedFormat, layout: layout);
                    break;
                case 0x000F:
                    int labelOffset = DataOffset(layout);
                    if (length < labelOffset + 1) throw new InvalidDataException("Truncated WK label cell record.");
                    byte prefix = data[payload + labelOffset];
                    string value = layout == WkRecordLayout.QuattroWq2
                        ? ReadPascalAscii(data, payload + labelOffset + 1, length - labelOffset - 1)
                        : ReadNullTerminatedAscii(data, payload + labelOffset + 1, length - labelOffset - 1);
                    AddTextCharacters(ref recoveredTextCharacters, value.Length, limits);
                    OfficeIMO.Excel.ExcelHorizontalAlignment? alignment = prefix == (byte)'^' ? OfficeIMO.Excel.ExcelHorizontalAlignment.Center
                        : prefix == (byte)'"' ? OfficeIMO.Excel.ExcelHorizontalAlignment.Right
                        : prefix == (byte)'\'' ? OfficeIMO.Excel.ExcelHorizontalAlignment.Left
                        : prefix == (byte)'\\' ? OfficeIMO.Excel.ExcelHorizontalAlignment.Fill : null;
                    AddCell(model, sheets, limits, data, payload, value, null, alignment, ref reportedUnsupportedFormat, isText: true, layout: layout);
                    break;
                case 0x0010:
                    ParseFormulaCell(model, sheets, limits, data, payload, length, ref recoveredTextCharacters,
                        ref reportedUnsupportedFormula, ref reportedUnsupportedFormat, layout, translateFormulas);
                    break;
                case 0x0033:
                    if (!reportedStringFormula) {
                        model.Findings.Add(Loss("WK_FORMULA_STRING_CACHE_PARTIAL", "Formula", "A separate formula-string cached-result record was present; numeric and translated formula cells were retained, but this string cache was not associated without a validated producer profile."));
                        reportedStringFormula = true;
                    }
                    break;
                default:
                    if (type == 0x002D || type == 0x002E) {
                        LegacySpreadsheetSheet chartSheet = GetSheet(model, sheets, 0);
                        if (model.Charts.Count >= limits.MaxItems) throw new InvalidDataException("Legacy spreadsheet exceeds the configured chart metadata limit.");
                        model.Charts.Add(new LegacySpreadsheetChartMetadata(chartSheet.Name, "0x" + type.ToString("X4", CultureInfo.InvariantCulture), length));
                    } else unsupportedRecordTypes.Add(type);
                    break;
            }
            offset = payload + length;
            if (foundEnd) break;
        }
        if (!foundEnd) model.Findings.Add(Loss("LEGACY_EOF_MISSING", "Container", "The record stream ended without a recognized EOF record; recovered records were retained."));
        if (unsupportedRecordTypes.Count > 0) {
            string recordIds = string.Join(",", unsupportedRecordTypes.OrderBy(static value => value).Select(static value => "0x" + value.ToString("X4", CultureInfo.InvariantCulture)));
            model.Metadata["UnsupportedRecordTypes"] = recordIds;
            model.Findings.Add(Loss("WK_RECORDS_UNSUPPORTED", "Structure", $"{unsupportedRecordTypes.Count} distinct source record kinds were not projected ({recordIds}); supported cells and metadata were retained."));
        }
        if (model.Charts.Count > 0) model.Findings.Add(Loss("LEGACY_CHART_METADATA_ONLY", "Chart", "Chart records were inventoried as metadata but were not converted into potentially misleading live charts."));
        if (model.RecoveredCellCount == 0) throw new InvalidDataException($"The {familyName} record stream contained no supported cells.");
        return model;
    }

    private static void ParseFormulaCell(LegacySpreadsheetModel model, Dictionary<byte, LegacySpreadsheetSheet> sheets,
        OfficeLegacyImportLimits limits, byte[] data, int payload, int length, ref int recoveredTextCharacters,
        ref bool reportedUnsupportedFormula, ref bool reportedUnsupportedFormat, WkRecordLayout layout, bool translateFormulas) {
        int dataOffset = DataOffset(layout);
        if (length < dataOffset + 10) throw new InvalidDataException("Truncated WK formula cell record envelope.");
        double cached = ReadDouble(data, payload + dataOffset);
        int tokenLength = OfficeLegacyImportBuffer.ReadUInt16(data, payload + dataOffset + 8);
        if (tokenLength != length - dataOffset - 10) throw new InvalidDataException("WK formula token length does not match its containing record.");
        string? formula = null;
        string error = "The selected source profile does not translate this formula dialect.";
        if (translateFormulas) {
            int rowZeroBased = ReadRow(data, payload, layout);
            int columnZeroBased = ReadColumn(data, payload, layout);
            int remainingTextCharacters = limits.MaxTextCharacters - recoveredTextCharacters;
            if (WkFormulaDecoder.TryDecode(data, payload + dataOffset + 10, tokenLength, rowZeroBased, columnZeroBased,
                    limits, remainingTextCharacters, out formula, out error)) {
                AddTextCharacters(ref recoveredTextCharacters, formula!.Length, limits);
                AddCell(model, sheets, limits, data, payload, cached, formula, null, ref reportedUnsupportedFormat, layout: layout);
                return;
            }
        }
        if (formula == null) {
            if (!reportedUnsupportedFormula) {
                model.Findings.Add(Loss("WK_FORMULA_CACHED_FALLBACK", "Formula", "At least one WK formula used an unsupported or invalid token sequence; its finite cached value was retained instead."));
                reportedUnsupportedFormula = true;
            }
            model.Metadata[$"FormulaFallback.{model.RecoveredCellCount + 1}"] = error;
        }
        AddCell(model, sheets, limits, data, payload, cached, formula, null, ref reportedUnsupportedFormat, layout: layout);
    }

    private static void AddCell(LegacySpreadsheetModel model, Dictionary<byte, LegacySpreadsheetSheet> sheets, OfficeLegacyImportLimits limits, byte[] data, int payload, object? value, string? formula, OfficeIMO.Excel.ExcelHorizontalAlignment? alignment, ref bool reportedUnsupportedFormat, bool isText = false, WkRecordLayout layout = WkRecordLayout.Dos) {
        if (model.RecoveredCellCount >= limits.MaxItems) throw new InvalidDataException("Legacy spreadsheet exceeds the configured cell limit.");
        byte format = layout == WkRecordLayout.Dos ? data[payload] : (byte)0;
        int column = ReadColumn(data, payload, layout) + 1;
        byte sheetId = ReadSheet(data, payload, layout);
        int row = ReadRow(data, payload, layout) + 1;
        if (row < 1 || row > 1048576 || column < 1 || column > 16384) throw new InvalidDataException("Legacy cell address is outside the supported workbook model.");
        string? numberFormat = WkCellFormatDecoder.Decode(format, isText);
        if (format != 0 && numberFormat == null && !reportedUnsupportedFormat) {
            model.Findings.Add(Loss("WK_CELL_FORMAT_PARTIAL", "Formatting", "At least one WK cell-format byte used a format family that is not safely mapped; supported numeric formats and label alignment were still projected."));
            reportedUnsupportedFormat = true;
        }
        LegacySpreadsheetSheet sheet = GetSheet(model, sheets, sheetId);
        sheet.Cells.Add(new LegacySpreadsheetCell(row, column, value, formula, format, numberFormat, alignment: alignment));
        model.RecoveredCellCount++;
    }

    private static LegacySpreadsheetSheet GetSheet(LegacySpreadsheetModel model, Dictionary<byte, LegacySpreadsheetSheet> sheets, byte id) {
        if (sheets.TryGetValue(id, out LegacySpreadsheetSheet? sheet)) return sheet;
        sheet = new LegacySpreadsheetSheet(id == 0 ? "Sheet1" : "Sheet" + (id + 1).ToString(CultureInfo.InvariantCulture));
        sheets.Add(id, sheet);
        model.Sheets.Add(sheet);
        return sheet;
    }

    private static void CaptureName(LegacySpreadsheetModel model, Dictionary<byte, LegacySpreadsheetSheet> sheets,
        byte[] data, int payload, int length, OfficeLegacyImportLimits limits, ref int recoveredTextCharacters) {
        if (length < 24) {
            string metadataName = ReadNullTerminatedAscii(data, payload, length).Trim();
            if (metadataName.Length == 0) return;
            AddTextCharacters(ref recoveredTextCharacters, metadataName.Length, limits);
            model.Metadata["UnresolvedName:" + model.Metadata.Count.ToString(CultureInfo.InvariantCulture)] = metadataName;
            model.Findings.Add(Loss("WK_NAME_REFERENCE_UNSUPPORTED", "Name", "A short WK name record was retained as metadata because it did not contain the validated 16-byte-name plus range profile."));
            return;
        }
        string name = data[payload] <= 15 && data[payload] <= length - 1
            ? Encoding.ASCII.GetString(data, payload + 1, data[payload]).Trim()
            : ReadNullTerminatedAscii(data, payload, 16).Trim();
        if (name.Length == 0) return;
        AddTextCharacters(ref recoveredTextCharacters, name.Length, limits);
        int firstColumn = OfficeLegacyImportBuffer.ReadUInt16(data, payload + 16) + 1;
        int firstRow = OfficeLegacyImportBuffer.ReadUInt16(data, payload + 18) + 1;
        int lastColumn = OfficeLegacyImportBuffer.ReadUInt16(data, payload + 20) + 1;
        int lastRow = OfficeLegacyImportBuffer.ReadUInt16(data, payload + 22) + 1;
        if (firstRow > 1048576 || lastRow > 1048576 || firstColumn > 16384 || lastColumn > 16384) throw new InvalidDataException("WK named range is outside the workbook model.");
        if (model.Names.Count >= limits.MaxItems) throw new InvalidDataException("Legacy spreadsheet exceeds the configured name limit.");
        LegacySpreadsheetSheet sheet = GetSheet(model, sheets, 0);
        model.Names.Add(new LegacySpreadsheetName(name, sheet.Name, firstRow, firstColumn, lastRow, lastColumn));
    }

    private static void ValidateBof(byte[] data, string familyName, byte expectedProduct0, byte expectedProduct1) {
        if (data.Length < 6 || OfficeLegacyImportBuffer.ReadUInt16(data, 0) != 0x0000 ||
            OfficeLegacyImportBuffer.ReadUInt16(data, 2) != 2 || data[4] != expectedProduct0 || data[5] != expectedProduct1) {
            throw new InvalidDataException($"The {familyName} record stream does not begin with its validated family BOF record.");
        }
    }

    private static void ValidateCellHeader(byte[] data, int payload, int length, WkRecordLayout layout) {
        if (length < DataOffset(layout) || payload > data.Length - length) throw new InvalidDataException("Truncated WK blank cell record.");
    }

    private static int DataOffset(WkRecordLayout layout) => layout == WkRecordLayout.QuattroWq2 ? 6 : 5;
    private static int ReadColumn(byte[] data, int payload, WkRecordLayout layout) => data[payload + (layout == WkRecordLayout.QuattroWq2 ? 0 : 1)];
    private static byte ReadSheet(byte[] data, int payload, WkRecordLayout layout) => data[payload + (layout == WkRecordLayout.QuattroWq2 ? 1 : 2)];
    private static int ReadRow(byte[] data, int payload, WkRecordLayout layout) => OfficeLegacyImportBuffer.ReadUInt16(data, payload + (layout == WkRecordLayout.QuattroWq2 ? 2 : 3));

    private static double ReadDouble(byte[] data, int offset) {
        if (offset < 0 || offset + 8 > data.Length) throw new InvalidDataException("Truncated legacy floating-point value.");
        double value;
        if (BitConverter.IsLittleEndian) value = BitConverter.ToDouble(data, offset);
        else { var copy = new byte[8]; Buffer.BlockCopy(data, offset, copy, 0, 8); Array.Reverse(copy); value = BitConverter.ToDouble(copy, 0); }
        if (double.IsNaN(value) || double.IsInfinity(value)) throw new InvalidDataException("Legacy cached numeric value is not finite.");
        return value;
    }

    private static void AddTextCharacters(ref int recoveredTextCharacters, int count, OfficeLegacyImportLimits limits) {
        if (count > limits.MaxTextCharacters - recoveredTextCharacters) throw new InvalidDataException("Legacy spreadsheet text exceeds the configured character limit.");
        recoveredTextCharacters += count;
    }

    private static string ReadNullTerminatedAscii(byte[] data, int offset, int length) {
        int available = Math.Min(length, data.Length - offset);
        int count = 0;
        while (count < available && data[offset + count] != 0) count++;
        return Encoding.ASCII.GetString(data, offset, count);
    }

    private static string ReadPascalAscii(byte[] data, int offset, int length) {
        if (length < 1) throw new InvalidDataException("Truncated Pascal string.");
        int count = data[offset];
        if (count > length - 1) throw new InvalidDataException("Pascal string length exceeds its record.");
        return Encoding.ASCII.GetString(data, offset + 1, count);
    }

    private static string ToHex(byte[] data, int offset, int length) {
        var builder = new StringBuilder(length * 2);
        for (int index = 0; index < length; index++) builder.Append(data[offset + index].ToString("X2", CultureInfo.InvariantCulture));
        return builder.ToString();
    }

}
