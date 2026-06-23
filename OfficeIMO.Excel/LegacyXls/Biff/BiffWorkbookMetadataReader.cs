using OfficeIMO.Excel.LegacyXls.Diagnostics;
using OfficeIMO.Excel.LegacyXls.Model;

namespace OfficeIMO.Excel.LegacyXls.Biff {
    internal static class BiffWorkbookMetadataReader {
        internal static bool TryRead(BiffRecord record, LegacyXlsWorkbook workbook, List<LegacyXlsImportDiagnostic> diagnostics) {
            switch ((BiffRecordType)record.Type) {
                case BiffRecordType.Backup:
                    if (TryReadBoolean(record, diagnostics, out bool saveBackup)) {
                        workbook.SetSaveBackup(saveBackup);
                    }

                    workbook.AddMetadataRecord(LegacyXlsWorkbookMetadataKind.Backup, record.Offset, record.Type);
                    return true;

                case BiffRecordType.BookBool:
                    if (TryReadUInt16(record, diagnostics, out ushort bookFlags)) {
                        workbook.SetBookOptions(bookFlags);
                    }

                    workbook.AddMetadataRecord(LegacyXlsWorkbookMetadataKind.BookOptions, record.Offset, record.Type);
                    return true;

                case BiffRecordType.CodePage:
                    if (TryReadUInt16(record, diagnostics, out ushort codePage)) {
                        workbook.SetCodePage(codePage);
                    }

                    workbook.AddMetadataRecord(LegacyXlsWorkbookMetadataKind.CodePage, record.Offset, record.Type);
                    return true;

                case BiffRecordType.Country:
                    if (TryReadCountry(record, diagnostics, out LegacyXlsCountryInfo? country)) {
                        workbook.SetCountry(country!.DefaultCountryCode, country.SystemCountryCode);
                    }

                    workbook.AddMetadataRecord(LegacyXlsWorkbookMetadataKind.Country, record.Offset, record.Type);
                    return true;

                case BiffRecordType.HideObj:
                    if (TryReadUInt16(record, diagnostics, out ushort hiddenObjectsMode)) {
                        workbook.SetHiddenObjectsMode(hiddenObjectsMode);
                    }

                    workbook.AddMetadataRecord(LegacyXlsWorkbookMetadataKind.HiddenObjects, record.Offset, record.Type);
                    return true;

                case BiffRecordType.InterfaceHdr:
                    if (TryReadUInt16(record, diagnostics, out ushort interfaceCodePage)) {
                        workbook.SetUserInterfaceCodePage(interfaceCodePage);
                    }

                    workbook.AddMetadataRecord(LegacyXlsWorkbookMetadataKind.InterfaceCodePage, record.Offset, record.Type);
                    return true;

                case BiffRecordType.InterfaceEnd:
                    workbook.AddMetadataRecord(LegacyXlsWorkbookMetadataKind.InterfaceEnd, record.Offset, record.Type);
                    return true;

                case BiffRecordType.PrintSize:
                    if (TryReadUInt16(record, diagnostics, out ushort printSize)) {
                        workbook.SetPrintSize(printSize);
                    }

                    workbook.AddMetadataRecord(LegacyXlsWorkbookMetadataKind.PrintSize, record.Offset, record.Type);
                    return true;

                case BiffRecordType.Prot4Rev:
                    if (TryReadBoolean(record, diagnostics, out bool revisionTrackingLocked)) {
                        workbook.SetRevisionTrackingLocked(revisionTrackingLocked);
                    }

                    workbook.AddMetadataRecord(LegacyXlsWorkbookMetadataKind.RevisionProtection, record.Offset, record.Type);
                    return true;

                case BiffRecordType.UsesElfs:
                    if (TryReadBoolean(record, diagnostics, out bool usesNaturalLanguageFormulas)) {
                        workbook.SetUsesNaturalLanguageFormulas(usesNaturalLanguageFormulas);
                    }

                    workbook.AddMetadataRecord(LegacyXlsWorkbookMetadataKind.NaturalLanguageFormulas, record.Offset, record.Type);
                    return true;

                case BiffRecordType.Window1:
                    if (TryReadWindow(record, diagnostics, out LegacyXlsWorkbookWindow? window)) {
                        workbook.AddWindow(window!);
                    }

                    workbook.AddMetadataRecord(LegacyXlsWorkbookMetadataKind.Window, record.Offset, record.Type);
                    return true;

                case BiffRecordType.WinProtect:
                    if (TryReadBoolean(record, diagnostics, out bool windowsLocked)) {
                        workbook.SetWindowsLocked(windowsLocked);
                    }

                    workbook.AddMetadataRecord(LegacyXlsWorkbookMetadataKind.WindowProtection, record.Offset, record.Type);
                    return true;

                case BiffRecordType.WriteAccess:
                    if (TryReadWriteAccess(record, diagnostics, out string? userName)) {
                        workbook.SetLastWriteUserName(userName);
                    }

                    workbook.AddMetadataRecord(LegacyXlsWorkbookMetadataKind.WriteAccess, record.Offset, record.Type);
                    return true;

                default:
                    return false;
            }
        }

        private static bool TryReadBoolean(BiffRecord record, List<LegacyXlsImportDiagnostic> diagnostics, out bool value) {
            if (!TryReadUInt16(record, diagnostics, out ushort rawValue)) {
                value = false;
                return false;
            }

            if (rawValue != 0 && rawValue != 1) {
                diagnostics.Add(new LegacyXlsImportDiagnostic(
                    LegacyXlsDiagnosticSeverity.Warning,
                    "XLS-BIFF-WORKBOOK-METADATA-VALUE-UNEXPECTED",
                    $"Workbook metadata record 0x{record.Type:X4} contains unexpected Boolean value {rawValue}.",
                    recordOffset: record.Offset,
                    recordType: record.Type));
            }

            value = rawValue != 0;
            return true;
        }

        private static bool TryReadUInt16(BiffRecord record, List<LegacyXlsImportDiagnostic> diagnostics, out ushort value) {
            if (record.Payload.Length < 2) {
                diagnostics.Add(new LegacyXlsImportDiagnostic(
                    LegacyXlsDiagnosticSeverity.Warning,
                    "XLS-BIFF-WORKBOOK-METADATA-SHORT",
                    $"Workbook metadata record 0x{record.Type:X4} is shorter than expected.",
                    recordOffset: record.Offset,
                    recordType: record.Type));
                value = 0;
                return false;
            }

            value = BiffRecordReader.ReadUInt16(record.Payload, 0);
            return true;
        }

        private static bool TryReadCountry(BiffRecord record, List<LegacyXlsImportDiagnostic> diagnostics, out LegacyXlsCountryInfo? country) {
            if (record.Payload.Length < 4) {
                diagnostics.Add(new LegacyXlsImportDiagnostic(
                    LegacyXlsDiagnosticSeverity.Warning,
                    "XLS-BIFF-WORKBOOK-METADATA-SHORT",
                    "The Country workbook metadata record is shorter than expected.",
                    recordOffset: record.Offset,
                    recordType: record.Type));
                country = null;
                return false;
            }

            country = new LegacyXlsCountryInfo(
                BiffRecordReader.ReadUInt16(record.Payload, 0),
                BiffRecordReader.ReadUInt16(record.Payload, 2));
            return true;
        }

        private static bool TryReadWindow(BiffRecord record, List<LegacyXlsImportDiagnostic> diagnostics, out LegacyXlsWorkbookWindow? window) {
            if (record.Payload.Length < 18) {
                diagnostics.Add(new LegacyXlsImportDiagnostic(
                    LegacyXlsDiagnosticSeverity.Warning,
                    "XLS-BIFF-WORKBOOK-METADATA-SHORT",
                    "The Window1 workbook metadata record is shorter than expected.",
                    recordOffset: record.Offset,
                    recordType: record.Type));
                window = null;
                return false;
            }

            ushort flags = BiffRecordReader.ReadUInt16(record.Payload, 8);
            window = new LegacyXlsWorkbookWindow(
                BiffRecordReader.ReadInt16(record.Payload, 0),
                BiffRecordReader.ReadInt16(record.Payload, 2),
                BiffRecordReader.ReadInt16(record.Payload, 4),
                BiffRecordReader.ReadInt16(record.Payload, 6),
                hidden: (flags & 0x0001) != 0,
                minimized: (flags & 0x0002) != 0,
                veryHidden: (flags & 0x0004) != 0,
                horizontalScrollBarVisible: (flags & 0x0008) != 0,
                verticalScrollBarVisible: (flags & 0x0010) != 0,
                sheetTabsVisible: (flags & 0x0020) != 0,
                autoFilterDatesGroupedChronologically: (flags & 0x0040) == 0,
                activeSheetIndex: BiffRecordReader.ReadUInt16(record.Payload, 10),
                firstVisibleSheetTabIndex: BiffRecordReader.ReadUInt16(record.Payload, 12),
                selectedSheetTabCount: BiffRecordReader.ReadUInt16(record.Payload, 14),
                sheetTabRatio: BiffRecordReader.ReadUInt16(record.Payload, 16));
            return true;
        }

        private static bool TryReadWriteAccess(BiffRecord record, List<LegacyXlsImportDiagnostic> diagnostics, out string? userName) {
            try {
                int offset = 0;
                string value = BiffStringReader.ReadUnicodeString(record.Payload, ref offset).TrimEnd('\0', ' ');
                userName = string.IsNullOrWhiteSpace(value) ? null : value;
                return true;
            } catch (InvalidDataException ex) {
                diagnostics.Add(new LegacyXlsImportDiagnostic(
                    LegacyXlsDiagnosticSeverity.Warning,
                    "XLS-BIFF-WRITEACCESS-INVALID",
                    $"The WriteAccess workbook metadata record could not be read. {ex.Message}",
                    recordOffset: record.Offset,
                    recordType: record.Type));
                userName = null;
                return false;
            }
        }
    }
}
