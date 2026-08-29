using System;
using System.Globalization;
using System.IO;
using System.Text;
using System.Threading;

namespace OfficeIMO.Excel.Legacy;

internal abstract class WkRecordSpreadsheetAdapterBase : LegacySpreadsheetAdapterBase {
    protected LegacySpreadsheetModel ParseWkRecords(byte[] data, OfficeLegacyImportLimits limits, string familyName, CancellationToken cancellationToken) {
        var model = new LegacySpreadsheetModel { Quality = OfficeLegacyImportQuality.Structured };
        var sheet = new LegacySpreadsheetSheet("Sheet1");
        model.Sheets.Add(sheet);
        int offset = 0;
        int records = 0;
        bool reportedFormulaTokens = false;
        bool reportedCellFormatting = false;
        bool foundEnd = false;
        int recoveredTextCharacters = 0;
        while (offset + 4 <= data.Length) {
            cancellationToken.ThrowIfCancellationRequested();
            if (++records > limits.MaxRecords) throw new InvalidDataException("Legacy spreadsheet exceeds the configured record limit.");
            ushort type = OfficeLegacyImportBuffer.ReadUInt16(data, offset);
            ushort length = OfficeLegacyImportBuffer.ReadUInt16(data, offset + 2);
            int payload = offset + 4;
            if (payload > data.Length - length) throw new InvalidDataException($"Truncated {familyName} record 0x{type:X4}.");

            switch (type) {
                case 0x0000:
                    model.Metadata["BofPayload"] = ToHex(data, payload, Math.Min(length, (ushort)8));
                    break;
                case 0x0001:
                    foundEnd = true;
                    break;
                case 0x000B:
                    CaptureName(model, data, payload, length, limits, ref recoveredTextCharacters);
                    break;
                case 0x000D:
                    if (length >= 7) {
                        ReportCellFormattingLoss(model, data[payload], ref reportedCellFormatting);
                        AddCell(model, sheet, limits, data, payload, OfficeLegacyImportBuffer.ReadInt16(data, payload + 5));
                    }
                    break;
                case 0x000E:
                    if (length >= 13) {
                        ReportCellFormattingLoss(model, data[payload], ref reportedCellFormatting);
                        AddCell(model, sheet, limits, data, payload, ReadDouble(data, payload + 5));
                    }
                    break;
                case 0x000F:
                    if (length >= 6) {
                        ReportCellFormattingLoss(model, data[payload], ref reportedCellFormatting);
                        byte prefix = data[payload + 5];
                        int textStart = payload + 6;
                        int textLength = Math.Max(0, length - 6);
                        string value = ReadNullTerminatedAscii(data, textStart, textLength);
                        AddTextCharacters(ref recoveredTextCharacters, value.Length, limits);
                        OfficeIMO.Excel.ExcelHorizontalAlignment? alignment = prefix == (byte)'^'
                            ? OfficeIMO.Excel.ExcelHorizontalAlignment.Center
                            : prefix == (byte)'"' ? OfficeIMO.Excel.ExcelHorizontalAlignment.Right
                            : prefix == (byte)'\'' ? OfficeIMO.Excel.ExcelHorizontalAlignment.Left
                            : prefix == (byte)'\\' ? OfficeIMO.Excel.ExcelHorizontalAlignment.Fill
                            : null;
                        AddCell(model, sheet, limits, data, payload, value, alignment);
                    }
                    break;
                case 0x0010:
                    if (length >= 13) {
                        ReportCellFormattingLoss(model, data[payload], ref reportedCellFormatting);
                        AddCell(model, sheet, limits, data, payload, ReadDouble(data, payload + 5));
                        if (!reportedFormulaTokens) {
                            model.Findings.Add(Loss("LEGACY_FORMULA_CACHED_VALUE", "Formula", "Formula cells use their stored cached value; source token expressions are not projected unless a profile-specific translation is proven safe."));
                            reportedFormulaTokens = true;
                        }
                    }
                    break;
                default:
                    if (type >= 0x002D && type <= 0x0036) {
                        if (model.Charts.Count >= limits.MaxItems) throw new InvalidDataException("Legacy spreadsheet exceeds the configured chart metadata limit.");
                        model.Charts.Add(new LegacySpreadsheetChartMetadata(sheet.Name, "0x" + type.ToString("X4", CultureInfo.InvariantCulture), length));
                    }
                    break;
            }
            offset = payload + length;
            if (foundEnd) break;
        }
        if (!foundEnd) model.Findings.Add(Loss("LEGACY_EOF_MISSING", "Container", "The record stream ended without a recognized EOF record; recovered records were retained."));
        if (model.Charts.Count > 0) model.Findings.Add(Loss("LEGACY_CHART_METADATA_ONLY", "Chart", "Chart records were inventoried as metadata but were not converted into potentially misleading live charts."));
        if (model.RecoveredCellCount == 0) throw new InvalidDataException($"The {familyName} record stream contained no supported cells.");
        DetectInertContent(model, data, cancellationToken);
        return model;
    }

    private static void AddCell(LegacySpreadsheetModel model, LegacySpreadsheetSheet sheet, OfficeLegacyImportLimits limits, byte[] data, int payload, object? value, OfficeIMO.Excel.ExcelHorizontalAlignment? alignment = null) {
        if (model.RecoveredCellCount >= limits.MaxItems) throw new InvalidDataException("Legacy spreadsheet exceeds the configured cell limit.");
        int column = OfficeLegacyImportBuffer.ReadUInt16(data, payload + 1) + 1;
        int row = OfficeLegacyImportBuffer.ReadUInt16(data, payload + 3) + 1;
        if (row < 1 || row > 1048576 || column < 1 || column > 16384) throw new InvalidDataException("Legacy cell address is outside the supported workbook model.");
        sheet.Cells.Add(new LegacySpreadsheetCell(row, column, value, format: data[payload], alignment: alignment));
        model.RecoveredCellCount++;
    }

    private static void CaptureName(LegacySpreadsheetModel model, byte[] data, int payload, int length, OfficeLegacyImportLimits limits, ref int recoveredTextCharacters) {
        string name = ReadNullTerminatedAscii(data, payload, length).Trim();
        if (name.Length == 0) return;
        if (model.Metadata.Count >= limits.MaxItems) throw new InvalidDataException("Legacy spreadsheet exceeds the configured metadata item limit.");
        AddTextCharacters(ref recoveredTextCharacters, name.Length, limits);
        model.Metadata["Name:" + model.Metadata.Count.ToString(CultureInfo.InvariantCulture)] = name;
        model.Findings.Add(Loss("LEGACY_NAME_METADATA_ONLY", "Name", "A source name was retained as metadata because its formula/reference tokens were not safely translatable."));
    }

    private static double ReadDouble(byte[] data, int offset) {
        if (offset < 0 || offset + 8 > data.Length) throw new InvalidDataException("Truncated legacy floating-point value.");
        double value;
        if (BitConverter.IsLittleEndian) {
            value = BitConverter.ToDouble(data, offset);
        } else {
            var copy = new byte[8];
            Buffer.BlockCopy(data, offset, copy, 0, 8);
            Array.Reverse(copy);
            value = BitConverter.ToDouble(copy, 0);
        }
        if (double.IsNaN(value) || double.IsInfinity(value)) throw new InvalidDataException("Legacy cached numeric value is not finite.");
        return value;
    }

    private static void AddTextCharacters(ref int recoveredTextCharacters, int count, OfficeLegacyImportLimits limits) {
        if (count > limits.MaxTextCharacters - recoveredTextCharacters) {
            throw new InvalidDataException("Legacy spreadsheet text exceeds the configured character limit.");
        }
        recoveredTextCharacters += count;
    }

    private static void ReportCellFormattingLoss(LegacySpreadsheetModel model, byte sourceFormat, ref bool reported) {
        if (sourceFormat == 0 || reported) return;
        model.Findings.Add(Loss("LEGACY_CELL_FORMAT_PARTIAL", "Formatting", "A non-default source cell-format byte was detected; supported label alignment was projected, while remaining number and style bits are reported as loss."));
        reported = true;
    }

    private static string ReadNullTerminatedAscii(byte[] data, int offset, int length) {
        int available = Math.Min(length, data.Length - offset);
        int count = 0;
        while (count < available && data[offset + count] != 0) count++;
        return Encoding.ASCII.GetString(data, offset, count).Trim();
    }

    private static string ToHex(byte[] data, int offset, int length) {
        var builder = new StringBuilder(length * 2);
        for (int index = 0; index < length; index++) builder.Append(data[offset + index].ToString("X2", CultureInfo.InvariantCulture));
        return builder.ToString();
    }

    private static void DetectInertContent(LegacySpreadsheetModel model, byte[] data, CancellationToken cancellationToken) {
        string printable = OfficeLegacyImportBuffer.ExtractPrintableText(data, 0, data.Length, Math.Min(data.Length, 256 * 1024), false, cancellationToken: cancellationToken);
        if (printable.IndexOf("http://", StringComparison.OrdinalIgnoreCase) >= 0 || printable.IndexOf("https://", StringComparison.OrdinalIgnoreCase) >= 0) {
            model.InertContent |= OfficeLegacyInertContentKind.ExternalLinks;
            model.Findings.Add(Inert("LEGACY_EXTERNAL_LINK_INERT", "Security", "External link text was discovered but never resolved or refreshed."));
        }
        if (printable.IndexOf("macro", StringComparison.OrdinalIgnoreCase) >= 0) {
            model.InertContent |= OfficeLegacyInertContentKind.Macros | OfficeLegacyInertContentKind.EmbeddedCode;
            model.Findings.Add(Inert("LEGACY_MACRO_INERT", "Security", "Macro markers were discovered but never executed or projected as executable code."));
        }
    }
}
