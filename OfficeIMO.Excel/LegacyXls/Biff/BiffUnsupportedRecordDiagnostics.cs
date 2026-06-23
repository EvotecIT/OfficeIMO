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
            GetUnsupportedRecordInfo(type, out LegacyXlsUnsupportedFeatureKind kind, out string code, out string message, out string detailCode);
            return new LegacyXlsUnsupportedFeature(
                kind,
                code,
                message,
                sheetName: sheetName,
                recordOffset: offset,
                recordType: type,
                detailCode: detailCode);
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
                recordType: feature.RecordType,
                detailCode: feature.DetailCode));
        }

        internal static bool IsPreserveOnlyFeatureRecord(ushort type) {
            GetUnsupportedRecordInfo(type, out LegacyXlsUnsupportedFeatureKind kind, out _, out _, out _);
            return kind != LegacyXlsUnsupportedFeatureKind.UnsupportedRecord;
        }

        private static void GetUnsupportedRecordInfo(
            ushort type,
            out LegacyXlsUnsupportedFeatureKind kind,
            out string code,
            out string message,
            out string detailCode) {
            kind = LegacyXlsUnsupportedFeatureKind.UnsupportedRecord;
            code = "XLS-BIFF-RECORD-UNSUPPORTED";
            message = $"BIFF record 0x{type:X4} is not imported in this phase.";
            detailCode = "BiffRecord:" + GetBiffRecordName(type);

            if (type == (ushort)BiffRecordType.HLink) {
                kind = LegacyXlsUnsupportedFeatureKind.Hyperlink;
                code = "XLS-BIFF-FEATURE-HYPERLINK-UNSUPPORTED";
                message = "A hyperlink record is present but its target shape is not supported by the current legacy XLS import phase.";
                detailCode = "Hyperlink:" + GetBiffRecordName(type);
            } else if (type == (ushort)BiffRecordType.Note) {
                kind = LegacyXlsUnsupportedFeatureKind.Comment;
                code = "XLS-BIFF-FEATURE-COMMENT-UNSUPPORTED";
                message = "Comment records are present but comment import is not implemented in this phase.";
                detailCode = "Comment:" + GetBiffRecordName(type);
            } else if (type == (ushort)BiffRecordType.Obj
                || type == (ushort)BiffRecordType.DrawingGroup
                || type == (ushort)BiffRecordType.Drawing) {
                kind = LegacyXlsUnsupportedFeatureKind.DrawingObject;
                code = "XLS-BIFF-FEATURE-DRAWING-UNSUPPORTED";
                message = "Drawing or object records are present but drawing import is not implemented in this phase.";
                detailCode = "Drawing:" + GetBiffRecordName(type);
            } else if (IsPivotTableRecord(type)) {
                kind = LegacyXlsUnsupportedFeatureKind.PivotTable;
                code = "XLS-BIFF-FEATURE-PIVOT-TABLE-UNSUPPORTED";
                message = "PivotTable records are present but pivot table import is not implemented in this phase.";
                detailCode = "PivotTable:" + GetBiffRecordName(type);
            } else if (IsExternalReferenceRecord(type)) {
                kind = LegacyXlsUnsupportedFeatureKind.ExternalReference;
                code = "XLS-BIFF-FEATURE-EXTERNAL-REFERENCE-UNSUPPORTED";
                message = "External reference records are present but external link import is not implemented in this phase.";
                detailCode = "ExternalReference:" + GetBiffRecordName(type);
            } else if (IsAutoFilterCriteriaRecord(type)) {
                kind = LegacyXlsUnsupportedFeatureKind.AutoFilterCriteria;
                code = "XLS-BIFF-FEATURE-AUTOFILTER-CRITERIA-UNSUPPORTED";
                message = "AutoFilter criteria records are present but this criteria shape is not supported by the current legacy XLS import phase.";
                detailCode = "AutoFilter:" + GetBiffRecordName(type);
            } else if (IsDataValidationRecord(type)) {
                kind = LegacyXlsUnsupportedFeatureKind.DataValidation;
                code = "XLS-BIFF-FEATURE-DATA-VALIDATION-UNSUPPORTED";
                message = "Data validation records are present but data validation import is not implemented in this phase.";
                detailCode = "DataValidation:" + GetBiffRecordName(type);
            } else if (IsConditionalFormattingRecord(type)) {
                kind = LegacyXlsUnsupportedFeatureKind.ConditionalFormatting;
                code = "XLS-BIFF-FEATURE-CONDITIONAL-FORMATTING-UNSUPPORTED";
                message = "Conditional formatting records are present but conditional formatting import is not implemented in this phase.";
                detailCode = "ConditionalFormatting:" + GetBiffRecordName(type);
            } else if (type >= 0x1000 && type <= 0x1066) {
                kind = LegacyXlsUnsupportedFeatureKind.Chart;
                code = "XLS-BIFF-FEATURE-CHART-UNSUPPORTED";
                message = "Chart records are present but chart import is not implemented in this phase.";
                detailCode = "Chart:" + GetBiffRecordName(type);
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
                recordType: record.Type,
                detailCode: "ExternalReference:" + reference.Kind);
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
                recordType: feature.RecordType,
                detailCode: feature.DetailCode));
        }

        private static string GetExternalReferenceDescription(LegacyXlsExternalReference reference) {
            return reference.Kind == LegacyXlsExternalReferenceKind.AddIn
                ? "add-in reference"
                : reference.Kind == LegacyXlsExternalReferenceKind.DdeOrOle
                    ? "DDE/OLE reference"
                    : "external workbook reference";
        }

        private static string GetBiffRecordName(ushort type) {
            switch (type) {
                case 0x000C: return "CalcCount";
                case 0x000D: return "CalcMode";
                case 0x000E: return "Precision";
                case 0x000F: return "RefMode";
                case 0x0010: return "Delta";
                case 0x0011: return "Iteration";
                case 0x0019: return "WinProtect";
                case 0x001D: return "Selection";
                case 0x0033: return "PrintSize";
                case 0x003D: return "Window1";
                case 0x0040: return "Backup";
                case 0x0042: return "CodePage";
                case 0x004D: return "Pls";
                case 0x0059: return "Xct";
                case 0x005A: return "Crn";
                case 0x005C: return "WriteAccess";
                case 0x005D: return "Obj";
                case 0x005F: return "SaveRecalc";
                case 0x0080: return "Guts";
                case 0x0081: return "WsBool";
                case 0x0082: return "GridSet";
                case 0x008C: return "Country";
                case 0x008D: return "HideObj";
                case 0x0090: return "Sort";
                case 0x009B: return "FilterMode";
                case 0x009C: return "FnGroupName";
                case 0x009D: return "AutoFilterInfo";
                case 0x009E: return "AutoFilter";
                case 0x00B0: return "SxView";
                case 0x00B1: return "Sxvd";
                case 0x00B2: return "Sxvi";
                case 0x00B4: return "SxIvd";
                case 0x00B5: return "Sxli";
                case 0x00B6: return "Sxpi";
                case 0x00C1: return "Sxdi";
                case 0x00C5: return "Sxdb";
                case 0x00C6: return "Sxfdb";
                case 0x00C7: return "Sxdbb";
                case 0x00C8: return "Sxnum";
                case 0x00C9: return "Sxbool";
                case 0x00CA: return "Sxerr";
                case 0x00CB: return "Sxint";
                case 0x00CC: return "Sxstring";
                case 0x00CD: return "Sxdtr";
                case 0x00CE: return "Sxnil";
                case 0x00CF: return "Sxtbl";
                case 0x00D0: return "Sxtbrgiitm";
                case 0x00D1: return "Sxtbpg";
                case 0x00D3: return "ObProj";
                case 0x00D5: return "SxStreamId";
                case 0x00D7: return "SxRng";
                case 0x00D8: return "SxIsxoper";
                case 0x00DA: return "BookBool";
                case 0x00E1: return "InterfaceHdr";
                case 0x00E2: return "InterfaceEnd";
                case 0x00EB: return "MsoDrawingGroup";
                case 0x00EC: return "MsoDrawing";
                case 0x00EF: return "SxRule";
                case 0x00F0: return "SxEx";
                case 0x00F1: return "SxFilt";
                case 0x00F2: return "SxDxf";
                case 0x00F4: return "SxItm";
                case 0x00F5: return "SxName";
                case 0x00F6: return "SxSelect";
                case 0x00F7: return "SxPair";
                case 0x00F8: return "SxFmla";
                case 0x00F9: return "SxFormat";
                case 0x00FF: return "SxVdEx";
                case 0x0100: return "SxFormula";
                case 0x0122: return "SxdbEx";
                case 0x013D: return "TabIdConf";
                case 0x0160: return "UsesElfs";
                case 0x0161: return "DsF";
                case 0x01AE: return "SupBook";
                case 0x01AF: return "Prot4Rev";
                case 0x01B0: return "CondFmt";
                case 0x01B1: return "Cf";
                case 0x01B2: return "DVal";
                case 0x01B6: return "TxO";
                case 0x01B7: return "RefreshAll";
                case 0x01B8: return "HLink";
                case 0x01BA: return "CodeName";
                case 0x01BC: return "SxViewEx";
                case 0x01BD: return "Sxth";
                case 0x01BE: return "Dv";
                case 0x020B: return "Index";
                case 0x0293: return "StyleExt";
                case 0x0800: return "WebPub";
                case 0x0801: return "QsiSxTag";
                case 0x0802: return "DbQueryExt";
                case 0x0804: return "TxtQry";
                case 0x0857: return "SxViewLink";
                case 0x0858: return "PivotChartBits";
                case 0x0863: return "SxAddl";
                case 0x0875: return "DConn";
                case 0x087A: return "Cf12";
                case 0x087B: return "CfEx";
                case 0x088D: return "Dxf";
                case 0x1001: return "Units";
                case 0x1002: return "Chart";
                case 0x1003: return "Series";
                case 0x1006: return "DataFormat";
                case 0x1007: return "LineFormat";
                case 0x1009: return "MarkerFormat";
                case 0x100A: return "AreaFormat";
                case 0x100B: return "PieFormat";
                case 0x100D: return "AttachedLabel";
                case 0x1014: return "ChartFormat";
                case 0x1015: return "Legend";
                case 0x1016: return "SeriesList";
                case 0x1017: return "Bar";
                case 0x1018: return "Line";
                case 0x1019: return "Pie";
                case 0x101A: return "Area";
                case 0x101B: return "Scatter";
                case 0x101C: return "ChartLine";
                case 0x101D: return "Axis";
                case 0x101E: return "Tick";
                case 0x101F: return "ValueRange";
                case 0x1020: return "CatSerRange";
                case 0x1021: return "AxisLineFormat";
                case 0x1022: return "ChartFormatLink";
                case 0x1024: return "DefaultText";
                case 0x1025: return "Text";
                case 0x1026: return "FontX";
                case 0x1027: return "ObjectLink";
                case 0x1032: return "Frame";
                case 0x1033: return "Begin";
                case 0x1034: return "End";
                case 0x1035: return "PlotArea";
                case 0x103A: return "Chart3d";
                case 0x1041: return "ShtProps";
                case 0x1044: return "SerToCrt";
                case 0x1045: return "AxesUsed";
                case 0x1046: return "SBaseRef";
                case 0x104F: return "Ifmt";
                case 0x1051: return "Pos";
                case 0x105C: return "SheetExt";
                case 0x105D: return "BookExt";
                case 0x105F: return "Dat";
                case 0x1060: return "PlotGrowth";
                case 0x1064: return "CrErr";
                case 0x1065: return "SeriesFormat";
                default:
                    return $"Record0x{type:X4}";
            }
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
                recordType: feature.RecordType,
                detailCode: feature.DetailCode));
        }

        internal static LegacyXlsUnsupportedFeature CreateUnsupportedSheetTypeFeature(BiffRecord record, LegacyXlsUnsupportedSheet unsupportedSheet) {
            GetUnsupportedSheetDiagnostic(unsupportedSheet, out LegacyXlsUnsupportedFeatureKind kind, out string code, out string description);
            return new LegacyXlsUnsupportedFeature(
                kind,
                code,
                $"The workbook contains a {description} entry. This legacy XLS import phase imports worksheet sheets only.",
                sheetName: unsupportedSheet.Name,
                recordOffset: record.Offset,
                recordType: record.Type,
                detailCode: "Sheet:" + unsupportedSheet.Kind);
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
                recordType: (ushort)BiffRecordType.WsBool,
                detailCode: "Sheet:" + unsupportedSheet.Kind);
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
                recordType: feature.RecordType,
                detailCode: feature.DetailCode));
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
