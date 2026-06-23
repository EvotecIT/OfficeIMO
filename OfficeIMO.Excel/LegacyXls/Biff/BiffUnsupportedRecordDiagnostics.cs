using OfficeIMO.Excel.LegacyXls.Diagnostics;
using OfficeIMO.Excel.LegacyXls.Model;

namespace OfficeIMO.Excel.LegacyXls.Biff {
    internal static class BiffUnsupportedRecordDiagnostics {
        internal static LegacyXlsUnsupportedFeature CreateFilePassFeature(BiffRecord record) {
            return new LegacyXlsUnsupportedFeature(
                LegacyXlsUnsupportedFeatureKind.EncryptedWorkbook,
                "XLS-BIFF-FILEPASS-UNSUPPORTED",
                "The workbook contains a FilePass record, which means password-to-open encryption is present. Encrypted legacy XLS import is not supported.",
                recordOffset: record.Offset,
                recordType: record.Type);
        }

        internal static void AddFilePassDiagnostic(BiffRecord record, List<LegacyXlsImportDiagnostic> diagnostics) {
            LegacyXlsUnsupportedFeature feature = CreateFilePassFeature(record);
            diagnostics.Add(new LegacyXlsImportDiagnostic(
                LegacyXlsDiagnosticSeverity.Error,
                feature.Code,
                feature.Description,
                recordOffset: feature.RecordOffset,
                recordType: feature.RecordType));
        }

        internal static LegacyXlsUnsupportedFeature CreateUnsupportedRecordFeature(
            ushort type,
            int offset,
            string? sheetName) {
            GetUnsupportedRecordInfo(type, out LegacyXlsUnsupportedFeatureKind kind, out string code, out string message);
            return new LegacyXlsUnsupportedFeature(
                kind,
                code,
                message,
                sheetName: sheetName,
                recordOffset: offset,
                recordType: type);
        }

        internal static void AddUnsupportedRecordDiagnostic(
            List<LegacyXlsImportDiagnostic> diagnostics,
            ushort type,
            int offset,
            string? sheetName) {
            LegacyXlsUnsupportedFeature feature = CreateUnsupportedRecordFeature(type, offset, sheetName);

            diagnostics.Add(new LegacyXlsImportDiagnostic(
                LegacyXlsDiagnosticSeverity.Info,
                feature.Code,
                feature.Description,
                sheetName: feature.SheetName,
                recordOffset: feature.RecordOffset,
                recordType: feature.RecordType));
        }

        private static void GetUnsupportedRecordInfo(
            ushort type,
            out LegacyXlsUnsupportedFeatureKind kind,
            out string code,
            out string message) {
            kind = LegacyXlsUnsupportedFeatureKind.UnsupportedRecord;
            code = "XLS-BIFF-RECORD-UNSUPPORTED";
            message = $"BIFF record 0x{type:X4} is not imported in this phase.";

            if (type == (ushort)BiffRecordType.HLink) {
                kind = LegacyXlsUnsupportedFeatureKind.Hyperlink;
                code = "XLS-BIFF-FEATURE-HYPERLINK-UNSUPPORTED";
                message = "A hyperlink record is present but its target shape is not supported by the current legacy XLS import phase.";
            } else if (type == (ushort)BiffRecordType.Note) {
                kind = LegacyXlsUnsupportedFeatureKind.Comment;
                code = "XLS-BIFF-FEATURE-COMMENT-UNSUPPORTED";
                message = "Comment records are present but comment import is not implemented in this phase.";
            } else if (type == (ushort)BiffRecordType.Obj
                || type == (ushort)BiffRecordType.DrawingGroup
                || type == (ushort)BiffRecordType.Drawing) {
                kind = LegacyXlsUnsupportedFeatureKind.DrawingObject;
                code = "XLS-BIFF-FEATURE-DRAWING-UNSUPPORTED";
                message = "Drawing or object records are present but drawing import is not implemented in this phase.";
            } else if (IsPivotTableRecord(type)) {
                kind = LegacyXlsUnsupportedFeatureKind.PivotTable;
                code = "XLS-BIFF-FEATURE-PIVOT-TABLE-UNSUPPORTED";
                message = "PivotTable records are present but pivot table import is not implemented in this phase.";
            } else if (IsExternalReferenceRecord(type)) {
                kind = LegacyXlsUnsupportedFeatureKind.ExternalReference;
                code = "XLS-BIFF-FEATURE-EXTERNAL-REFERENCE-UNSUPPORTED";
                message = "External reference records are present but external link import is not implemented in this phase.";
            } else if (IsAutoFilterCriteriaRecord(type)) {
                kind = LegacyXlsUnsupportedFeatureKind.AutoFilterCriteria;
                code = "XLS-BIFF-FEATURE-AUTOFILTER-CRITERIA-UNSUPPORTED";
                message = "AutoFilter criteria records are present but this criteria shape is not supported by the current legacy XLS import phase.";
            } else if (IsDataValidationRecord(type)) {
                kind = LegacyXlsUnsupportedFeatureKind.DataValidation;
                code = "XLS-BIFF-FEATURE-DATA-VALIDATION-UNSUPPORTED";
                message = "Data validation records are present but data validation import is not implemented in this phase.";
            } else if (IsConditionalFormattingRecord(type)) {
                kind = LegacyXlsUnsupportedFeatureKind.ConditionalFormatting;
                code = "XLS-BIFF-FEATURE-CONDITIONAL-FORMATTING-UNSUPPORTED";
                message = "Conditional formatting records are present but conditional formatting import is not implemented in this phase.";
            } else if (type >= 0x1000 && type <= 0x1066) {
                kind = LegacyXlsUnsupportedFeatureKind.Chart;
                code = "XLS-BIFF-FEATURE-CHART-UNSUPPORTED";
                message = "Chart records are present but chart import is not implemented in this phase.";
            }
        }

        internal static LegacyXlsUnsupportedFeature CreateExternalReferenceFeature(BiffRecord record, LegacyXlsExternalReference reference) {
            string description = GetExternalReferenceDescription(reference);
            string target = string.IsNullOrWhiteSpace(reference.Target) ? string.Empty : $" Target: {reference.Target}.";
            return new LegacyXlsUnsupportedFeature(
                LegacyXlsUnsupportedFeatureKind.ExternalReference,
                "XLS-BIFF-FEATURE-EXTERNAL-REFERENCE-UNSUPPORTED",
                $"The workbook contains a {description}. External link import is not implemented in this phase.{target}",
                recordOffset: record.Offset,
                recordType: record.Type);
        }

        internal static void AddExternalReferenceDiagnostic(
            List<LegacyXlsImportDiagnostic> diagnostics,
            BiffRecord record,
            LegacyXlsExternalReference reference) {
            LegacyXlsUnsupportedFeature feature = CreateExternalReferenceFeature(record, reference);
            diagnostics.Add(new LegacyXlsImportDiagnostic(
                LegacyXlsDiagnosticSeverity.Info,
                feature.Code,
                feature.Description,
                recordOffset: feature.RecordOffset,
                recordType: feature.RecordType));
        }

        private static string GetExternalReferenceDescription(LegacyXlsExternalReference reference) {
            return reference.Kind == LegacyXlsExternalReferenceKind.AddIn
                ? "add-in reference"
                : reference.Kind == LegacyXlsExternalReferenceKind.DdeOrOle
                    ? "DDE/OLE reference"
                    : "external workbook reference";
        }

        private static bool IsExternalReferenceRecord(ushort type) {
            return type == (ushort)BiffRecordType.ExternName
                || type == 0x0059 // XCT
                || type == 0x005A // CRN
                || type == 0x01B7 // RefreshAll
                || type == 0x0800 // WebPub
                || type == 0x0802 // DBQueryExt
                || type == 0x0804 // TxtQry
                || type == 0x0875; // DConn
        }

        private static bool IsAutoFilterCriteriaRecord(ushort type) {
            return type == (ushort)BiffRecordType.FilterMode
                || type == (ushort)BiffRecordType.AutoFilterInfo
                || type == (ushort)BiffRecordType.AutoFilter;
        }

        private static bool IsDataValidationRecord(ushort type) {
            return type == (ushort)BiffRecordType.DVal
                || type == (ushort)BiffRecordType.Dv;
        }

        private static bool IsConditionalFormattingRecord(ushort type) {
            return type == (ushort)BiffRecordType.CondFmt
                || type == (ushort)BiffRecordType.Cf
                || type == (ushort)BiffRecordType.Cf12
                || type == (ushort)BiffRecordType.CfEx
                || type == (ushort)BiffRecordType.Dxf;
        }

        private static bool IsPivotTableRecord(ushort type) {
            switch (type) {
                case 0x00B0: // SxView
                case 0x00B1: // Sxvd
                case 0x00B2: // SXVI
                case 0x00B4: // SxIvd
                case 0x00B5: // SXLI
                case 0x00B6: // SXPI
                case 0x00C1: // SXDI
                case 0x00C5: // SXDB
                case 0x00C6: // SXFDB
                case 0x00C7: // SXDBB
                case 0x00C8: // SXNum
                case 0x00C9: // SxBool
                case 0x00CA: // SxErr
                case 0x00CB: // SXInt
                case 0x00CC: // SXString
                case 0x00CD: // SXDtr
                case 0x00CE: // SxNil
                case 0x00CF: // SXTbl
                case 0x00D0: // SXTBRGIITM
                case 0x00D1: // SxTbpg
                case 0x00D5: // SXStreamID
                case 0x00D7: // SXRng
                case 0x00D8: // SxIsxoper
                case 0x00EF: // SxRule
                case 0x00F0: // SXEx
                case 0x00F1: // SxFilt
                case 0x00F2: // SxDXF
                case 0x00F4: // SxItm
                case 0x00F5: // SxName
                case 0x00F6: // SxSelect
                case 0x00F7: // SXPair
                case 0x00F8: // SxFmla
                case 0x00F9: // SxFormat
                case 0x00FF: // SXVDEx
                case 0x0100: // SXFormula
                case 0x0122: // SXDBEx
                case 0x0801: // QsiSXTag
                case 0x0857: // SXViewLink
                case 0x0858: // PivotChartBits
                case 0x0863: // SXAddl
                    return true;
                default:
                    return false;
            }
        }

        internal static void AddUnsupportedSheetTypeDiagnostic(
            List<LegacyXlsImportDiagnostic> diagnostics,
            BiffRecord record,
            LegacyXlsUnsupportedSheet unsupportedSheet) {
            LegacyXlsUnsupportedFeature feature = CreateUnsupportedSheetTypeFeature(record, unsupportedSheet);

            diagnostics.Add(new LegacyXlsImportDiagnostic(
                LegacyXlsDiagnosticSeverity.Info,
                feature.Code,
                feature.Description,
                sheetName: feature.SheetName,
                recordOffset: feature.RecordOffset,
                recordType: feature.RecordType));
        }

        internal static LegacyXlsUnsupportedFeature CreateUnsupportedSheetTypeFeature(BiffRecord record, LegacyXlsUnsupportedSheet unsupportedSheet) {
            GetUnsupportedSheetDiagnostic(unsupportedSheet, out LegacyXlsUnsupportedFeatureKind kind, out string code, out string description);
            return new LegacyXlsUnsupportedFeature(
                kind,
                code,
                $"The workbook contains a {description} entry. This legacy XLS import phase imports worksheet sheets only.",
                sheetName: unsupportedSheet.Name,
                recordOffset: record.Offset,
                recordType: record.Type);
        }

        internal static LegacyXlsUnsupportedFeature CreateUnsupportedDialogSheetFeature(
            int offset,
            LegacyXlsUnsupportedSheet unsupportedSheet) {
            return new LegacyXlsUnsupportedFeature(
                LegacyXlsUnsupportedFeatureKind.DialogSheet,
                "XLS-BIFF-FEATURE-DIALOG-SHEET-UNSUPPORTED",
                "The workbook contains a dialog sheet entry. This legacy XLS import phase imports worksheet sheets only.",
                sheetName: unsupportedSheet.Name,
                recordOffset: offset,
                recordType: (ushort)BiffRecordType.WsBool);
        }

        internal static void AddUnsupportedDialogSheetDiagnostic(
            List<LegacyXlsImportDiagnostic> diagnostics,
            int offset,
            LegacyXlsUnsupportedSheet unsupportedSheet) {
            LegacyXlsUnsupportedFeature feature = CreateUnsupportedDialogSheetFeature(offset, unsupportedSheet);
            diagnostics.Add(new LegacyXlsImportDiagnostic(
                LegacyXlsDiagnosticSeverity.Info,
                feature.Code,
                feature.Description,
                sheetName: feature.SheetName,
                recordOffset: feature.RecordOffset,
                recordType: feature.RecordType));
        }

        private static void GetUnsupportedSheetDiagnostic(
            LegacyXlsUnsupportedSheet unsupportedSheet,
            out LegacyXlsUnsupportedFeatureKind kind,
            out string code,
            out string description) {
            switch (unsupportedSheet.Kind) {
                case LegacyXlsUnsupportedSheetKind.MacroSheet:
                    kind = LegacyXlsUnsupportedFeatureKind.MacroSheet;
                    code = "XLS-BIFF-FEATURE-MACRO-SHEET-UNSUPPORTED";
                    description = "macro sheet";
                    break;
                case LegacyXlsUnsupportedSheetKind.ChartSheet:
                    kind = LegacyXlsUnsupportedFeatureKind.ChartSheet;
                    code = "XLS-BIFF-FEATURE-CHART-SHEET-UNSUPPORTED";
                    description = "chart sheet";
                    break;
                case LegacyXlsUnsupportedSheetKind.VbaModuleSheet:
                    kind = LegacyXlsUnsupportedFeatureKind.VbaModuleSheet;
                    code = "XLS-BIFF-FEATURE-VBA-MODULE-SHEET-UNSUPPORTED";
                    description = "VBA module sheet";
                    break;
                default:
                    kind = LegacyXlsUnsupportedFeatureKind.UnsupportedSheet;
                    code = "XLS-BIFF-FEATURE-SHEET-TYPE-UNSUPPORTED";
                    description = $"sheet type 0x{unsupportedSheet.SheetType:X2}";
                    break;
            }
        }
    }
}
