using OfficeIMO.Excel.LegacyXls.Diagnostics;
using OfficeIMO.Excel.LegacyXls.Model;

namespace OfficeIMO.Excel.LegacyXls.Biff {
    internal static class BiffCalculationSettingsReader {
        internal static bool TryRead(
            BiffRecord record,
            string? sheetName,
            LegacyXlsCalculationSettings settings,
            List<LegacyXlsImportDiagnostic> diagnostics) {
            BiffRecordType type = (BiffRecordType)record.Type;
            switch (type) {
                case BiffRecordType.CalcCount:
                    return TryReadInt16(record, sheetName, LegacyXlsCalculationSettingKind.IterationCount, settings, diagnostics);
                case BiffRecordType.CalcMode:
                    return TryReadCalcMode(record, sheetName, settings, diagnostics);
                case BiffRecordType.CalcPrecision:
                    return TryReadBoolean(record, sheetName, LegacyXlsCalculationSettingKind.FullPrecision, settings, diagnostics);
                case BiffRecordType.CalcRefMode:
                    return TryReadBoolean(record, sheetName, LegacyXlsCalculationSettingKind.A1ReferenceMode, settings, diagnostics);
                case BiffRecordType.CalcDelta:
                    return TryReadDouble(record, sheetName, LegacyXlsCalculationSettingKind.Delta, settings, diagnostics);
                case BiffRecordType.CalcIter:
                    return TryReadBoolean(record, sheetName, LegacyXlsCalculationSettingKind.IterationEnabled, settings, diagnostics);
                case BiffRecordType.CalcSaveRecalc:
                    return TryReadBoolean(record, sheetName, LegacyXlsCalculationSettingKind.RecalculateBeforeSave, settings, diagnostics);
                default:
                    return false;
            }
        }

        private static bool TryReadCalcMode(
            BiffRecord record,
            string? sheetName,
            LegacyXlsCalculationSettings settings,
            List<LegacyXlsImportDiagnostic> diagnostics) {
            if (!TryReadSignedValue(record, sheetName, diagnostics, out short value)) {
                return true;
            }

            LegacyXlsCalculationMode? mode = value switch {
                0 => LegacyXlsCalculationMode.Manual,
                1 => LegacyXlsCalculationMode.Automatic,
                2 => LegacyXlsCalculationMode.AutomaticExceptTables,
                _ => null
            };

            settings.AddRecord(new LegacyXlsCalculationSettingRecord(
                LegacyXlsCalculationSettingKind.Mode,
                sheetName,
                record.Offset,
                record.Type,
                signedValue: value,
                mode: mode));
            if (mode == null) {
                diagnostics.Add(new LegacyXlsImportDiagnostic(
                    LegacyXlsDiagnosticSeverity.Warning,
                    "XLS-BIFF-CALC-MODE-UNKNOWN",
                    $"A CalcMode record contains unknown mode value {value}.",
                    sheetName: sheetName,
                    recordOffset: record.Offset,
                    recordType: record.Type,
                    detailCode: value.ToString(System.Globalization.CultureInfo.InvariantCulture)));
            }

            return true;
        }

        private static bool TryReadInt16(
            BiffRecord record,
            string? sheetName,
            LegacyXlsCalculationSettingKind kind,
            LegacyXlsCalculationSettings settings,
            List<LegacyXlsImportDiagnostic> diagnostics) {
            if (!TryReadSignedValue(record, sheetName, diagnostics, out short value)) {
                return true;
            }

            settings.AddRecord(new LegacyXlsCalculationSettingRecord(
                kind,
                sheetName,
                record.Offset,
                record.Type,
                signedValue: value));
            return true;
        }

        private static bool TryReadBoolean(
            BiffRecord record,
            string? sheetName,
            LegacyXlsCalculationSettingKind kind,
            LegacyXlsCalculationSettings settings,
            List<LegacyXlsImportDiagnostic> diagnostics) {
            if (!TryReadSignedValue(record, sheetName, diagnostics, out short value)) {
                return true;
            }

            settings.AddRecord(new LegacyXlsCalculationSettingRecord(
                kind,
                sheetName,
                record.Offset,
                record.Type,
                signedValue: value,
                booleanValue: value != 0));
            return true;
        }

        private static bool TryReadDouble(
            BiffRecord record,
            string? sheetName,
            LegacyXlsCalculationSettingKind kind,
            LegacyXlsCalculationSettings settings,
            List<LegacyXlsImportDiagnostic> diagnostics) {
            if (record.Payload.Length < 8) {
                AddShortRecordDiagnostic(record, sheetName, diagnostics, expectedBytes: 8);
                return true;
            }

            settings.AddRecord(new LegacyXlsCalculationSettingRecord(
                kind,
                sheetName,
                record.Offset,
                record.Type,
                doubleValue: BiffRecordReader.ReadDouble(record.Payload, 0)));
            return true;
        }

        private static bool TryReadSignedValue(
            BiffRecord record,
            string? sheetName,
            List<LegacyXlsImportDiagnostic> diagnostics,
            out short value) {
            value = 0;
            if (record.Payload.Length < 2) {
                AddShortRecordDiagnostic(record, sheetName, diagnostics, expectedBytes: 2);
                return false;
            }

            value = BiffRecordReader.ReadInt16(record.Payload, 0);
            return true;
        }

        private static void AddShortRecordDiagnostic(
            BiffRecord record,
            string? sheetName,
            List<LegacyXlsImportDiagnostic> diagnostics,
            int expectedBytes) {
            diagnostics.Add(new LegacyXlsImportDiagnostic(
                LegacyXlsDiagnosticSeverity.Warning,
                "XLS-BIFF-CALC-SETTING-SHORT",
                $"A calculation setting record was too short to parse. Expected at least {expectedBytes} payload bytes.",
                sheetName: sheetName,
                recordOffset: record.Offset,
                recordType: record.Type,
                detailCode: $"0x{record.Type:X4}"));
        }
    }
}
