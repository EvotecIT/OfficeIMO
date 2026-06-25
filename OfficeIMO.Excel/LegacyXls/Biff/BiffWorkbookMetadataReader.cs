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

                case BiffRecordType.BookExt:
                    workbook.AddMetadataRecord(LegacyXlsWorkbookMetadataKind.BookExtension, record.Offset, record.Type);
                    return true;

                case BiffRecordType.BuiltInFnGroupCount:
                    if (TryReadBuiltInFunctionGroupCount(record, diagnostics, out ushort builtInFunctionGroupCount)) {
                        workbook.SetBuiltInFunctionGroupCount(builtInFunctionGroupCount);
                    }

                    workbook.AddMetadataRecord(LegacyXlsWorkbookMetadataKind.BuiltInFunctionGroupCount, record.Offset, record.Type);
                    return true;

                case BiffRecordType.CodePage:
                    if (TryReadUInt16(record, diagnostics, out ushort codePage)) {
                        workbook.SetCodePage(codePage);
                    }

                    workbook.AddMetadataRecord(LegacyXlsWorkbookMetadataKind.CodePage, record.Offset, record.Type);
                    return true;

                case BiffRecordType.CodeName:
                    if (BiffCodeNameReader.TryRead(record, sheetName: null, diagnostics, out string? workbookCodeName)) {
                        workbook.SetCodeName(workbookCodeName);
                    }

                    workbook.AddMetadataRecord(LegacyXlsWorkbookMetadataKind.CodeName, record.Offset, record.Type);
                    return true;

                case BiffRecordType.Country:
                    if (TryReadCountry(record, diagnostics, out LegacyXlsCountryInfo? country)) {
                        workbook.SetCountry(country!.DefaultCountryCode, country.SystemCountryCode);
                    }

                    workbook.AddMetadataRecord(LegacyXlsWorkbookMetadataKind.Country, record.Offset, record.Type);
                    return true;

                case BiffRecordType.Dsf:
                    ValidateReservedZero(record, diagnostics);
                    workbook.AddMetadataRecord(LegacyXlsWorkbookMetadataKind.ReservedDsf, record.Offset, record.Type);
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

                case BiffRecordType.Pls:
                    BiffPrinterSettingsReader.Validate(record, sheetName: null, diagnostics);
                    workbook.AddMetadataRecord(LegacyXlsWorkbookMetadataKind.PrinterSettings, record.Offset, record.Type);
                    return true;

                case BiffRecordType.PrintSize:
                    if (TryReadUInt16(record, diagnostics, out ushort printSize)) {
                        workbook.SetPrintSize(printSize);
                    }

                    workbook.AddMetadataRecord(LegacyXlsWorkbookMetadataKind.PrintSize, record.Offset, record.Type);
                    return true;

                case BiffRecordType.RefreshAll:
                    workbook.SetHasRefreshAllMarker();
                    workbook.AddMetadataRecord(LegacyXlsWorkbookMetadataKind.RefreshAll, record.Offset, record.Type);
                    return true;

                case BiffRecordType.Prot4Rev:
                    if (TryReadBoolean(record, diagnostics, out bool revisionTrackingLocked)) {
                        workbook.SetRevisionTrackingLocked(revisionTrackingLocked);
                    }

                    workbook.AddMetadataRecord(LegacyXlsWorkbookMetadataKind.RevisionProtection, record.Offset, record.Type);
                    return true;

                case BiffRecordType.Prot4RevPass:
                    if (TryReadUInt16(record, diagnostics, out ushort revisionTrackingPasswordHash)) {
                        workbook.SetRevisionTrackingPasswordHash(revisionTrackingPasswordHash);
                    }

                    workbook.AddMetadataRecord(LegacyXlsWorkbookMetadataKind.RevisionProtectionPassword, record.Offset, record.Type);
                    return true;

                case BiffRecordType.TabId:
                    if (TryReadSheetTabIds(record, diagnostics, out LegacyXlsSheetTabIdCollection? sheetTabIds)) {
                        workbook.SetSheetTabIds(sheetTabIds!);
                    }

                    workbook.AddMetadataRecord(LegacyXlsWorkbookMetadataKind.SheetTabIds, record.Offset, record.Type);
                    return true;

                case BiffRecordType.UsesElfs:
                    if (TryReadBoolean(record, diagnostics, out bool usesNaturalLanguageFormulas)) {
                        workbook.SetUsesNaturalLanguageFormulas(usesNaturalLanguageFormulas);
                    }

                    workbook.AddMetadataRecord(LegacyXlsWorkbookMetadataKind.NaturalLanguageFormulas, record.Offset, record.Type);
                    return true;

                case BiffRecordType.ObProj:
                    workbook.SetHasVbaProjectMarker();
                    workbook.AddMetadataRecord(LegacyXlsWorkbookMetadataKind.VbaProjectMarker, record.Offset, record.Type);
                    return true;

                case BiffRecordType.ObNoMacros:
                    workbook.SetHasVbaProjectWithoutMacros();
                    workbook.AddMetadataRecord(LegacyXlsWorkbookMetadataKind.VbaProjectNoMacrosMarker, record.Offset, record.Type);
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

                case BiffRecordType.RecalcId:
                    AddFutureMetadataRecord(workbook, record, LegacyXlsWorkbookMetadataKind.RecalculationIdentifier);
                    return true;

                case BiffRecordType.EntExU2:
                    AddFutureMetadataRecord(workbook, record, LegacyXlsWorkbookMetadataKind.ExtendedEncryption);
                    return true;

                case BiffRecordType.ContinueFrt:
                    AddFutureMetadataRecord(workbook, record, LegacyXlsWorkbookMetadataKind.FutureRecordContinuation);
                    return true;

                case BiffRecordType.Compat12:
                    AddFutureMetadataRecord(workbook, record, LegacyXlsWorkbookMetadataKind.Compatibility12);
                    return true;

                case BiffRecordType.NamePublish:
                    AddFutureMetadataRecord(workbook, record, LegacyXlsWorkbookMetadataKind.NamePublish);
                    return true;

                case BiffRecordType.NameCmt:
                    AddFutureMetadataRecord(workbook, record, LegacyXlsWorkbookMetadataKind.NameComment);
                    return true;

                case BiffRecordType.SortData:
                    AddFutureMetadataRecord(workbook, record, LegacyXlsWorkbookMetadataKind.SortData);
                    return true;

                case BiffRecordType.GuidTypeLib:
                    AddFutureMetadataRecord(workbook, record, LegacyXlsWorkbookMetadataKind.TypeLibraryGuid);
                    return true;

                case BiffRecordType.FnGrp12:
                    AddFutureMetadataRecord(workbook, record, LegacyXlsWorkbookMetadataKind.FunctionGroup12);
                    return true;

                case BiffRecordType.NameFnGrp12:
                    AddFutureMetadataRecord(workbook, record, LegacyXlsWorkbookMetadataKind.NameFunctionGroup12);
                    return true;

                case BiffRecordType.MtrSettings:
                    AddFutureMetadataRecord(workbook, record, LegacyXlsWorkbookMetadataKind.MultiThreadedRecalculationSettings);
                    return true;

                case BiffRecordType.CompressPictures:
                    AddFutureMetadataRecord(workbook, record, LegacyXlsWorkbookMetadataKind.CompressPictures);
                    return true;

                case BiffRecordType.HeaderFooter:
                    AddFutureMetadataRecord(workbook, record, LegacyXlsWorkbookMetadataKind.HeaderFooter);
                    return true;

                default:
                    return false;
            }
        }

        private static void AddFutureMetadataRecord(
            LegacyXlsWorkbook workbook,
            BiffRecord record,
            LegacyXlsWorkbookMetadataKind kind) {
            ushort? headerRecordType = record.Payload.Length >= 2
                ? BiffRecordReader.ReadUInt16(record.Payload, 0)
                : null;
            ushort? headerFlags = record.Payload.Length >= 4
                ? BiffRecordReader.ReadUInt16(record.Payload, 2)
                : null;

            workbook.AddFutureMetadataRecord(new LegacyXlsWorkbookFutureMetadataRecord(
                kind,
                record.Offset,
                record.Type,
                record.Payload.Length,
                headerRecordType,
                headerFlags));
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

        private static void ValidateReservedZero(BiffRecord record, List<LegacyXlsImportDiagnostic> diagnostics) {
            if (!TryReadUInt16(record, diagnostics, out ushort rawValue)) {
                return;
            }

            if (rawValue != 0) {
                diagnostics.Add(new LegacyXlsImportDiagnostic(
                    LegacyXlsDiagnosticSeverity.Warning,
                    "XLS-BIFF-WORKBOOK-METADATA-RESERVED-VALUE-UNEXPECTED",
                    $"Reserved workbook metadata record 0x{record.Type:X4} contains unexpected value {rawValue}.",
                    recordOffset: record.Offset,
                    recordType: record.Type));
            }
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

        private static bool TryReadSheetTabIds(
            BiffRecord record,
            List<LegacyXlsImportDiagnostic> diagnostics,
            out LegacyXlsSheetTabIdCollection? sheetTabIds) {
            sheetTabIds = null;
            if ((record.Payload.Length % 2) != 0) {
                diagnostics.Add(new LegacyXlsImportDiagnostic(
                    LegacyXlsDiagnosticSeverity.Warning,
                    "XLS-BIFF-TABID-LENGTH-INVALID",
                    "The TabId record contains a partial sheet tab identifier.",
                    recordOffset: record.Offset,
                    recordType: record.Type));
                return false;
            }

            var ids = new List<ushort>(record.Payload.Length / 2);
            for (int offset = 0; offset < record.Payload.Length; offset += 2) {
                ids.Add(BiffRecordReader.ReadUInt16(record.Payload, offset));
            }

            sheetTabIds = new LegacyXlsSheetTabIdCollection(ids);
            return true;
        }

        private static bool TryReadBuiltInFunctionGroupCount(BiffRecord record, List<LegacyXlsImportDiagnostic> diagnostics, out ushort value) {
            if (!TryReadUInt16(record, diagnostics, out value)) {
                return false;
            }

            if (value != 0x000e && value != 0x0010 && value != 0x0011) {
                diagnostics.Add(new LegacyXlsImportDiagnostic(
                    LegacyXlsDiagnosticSeverity.Warning,
                    "XLS-BIFF-BUILTIN-FNGROUPCOUNT-UNEXPECTED",
                    $"The BuiltInFnGroupCount workbook metadata record contains unexpected value {value}.",
                    recordOffset: record.Offset,
                    recordType: record.Type));
            }

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
