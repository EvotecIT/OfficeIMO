using OfficeIMO.Excel.LegacyXls.Model;

namespace OfficeIMO.Excel.LegacyXls.Biff {
    internal static class BiffChartMetadataReader {
        internal static bool TryRead(
            BiffRecord record,
            string? sheetName,
            List<LegacyXlsChartRecord> records) {
            if (!BiffUnsupportedRecordDiagnostics.IsChartRecord(record.Type)) {
                return false;
            }

            TryReadChartRectangle(record, out int? chartX, out int? chartY, out int? chartWidth, out int? chartHeight);
            records.Add(new LegacyXlsChartRecord(
                GetKind(record.Type),
                BiffUnsupportedRecordDiagnostics.GetBiffRecordName(record.Type),
                sheetName,
                record.Offset,
                record.Type,
                record.Payload.Length,
                GetChartTypeName(record.Type),
                chartX,
                chartY,
                chartWidth,
                chartHeight));
            return true;
        }

        private static bool TryReadChartRectangle(BiffRecord record, out int? x, out int? y, out int? width, out int? height) {
            x = null;
            y = null;
            width = null;
            height = null;
            if (record.Type != 0x1002 || record.Payload.Length < 16) {
                return false;
            }

            x = BiffRecordReader.ReadInt32(record.Payload, 0);
            y = BiffRecordReader.ReadInt32(record.Payload, 4);
            width = BiffRecordReader.ReadInt32(record.Payload, 8);
            height = BiffRecordReader.ReadInt32(record.Payload, 12);
            return true;
        }

        private static string? GetChartTypeName(ushort type) {
            switch (type) {
                case 0x1017:
                    return "Bar";
                case 0x1018:
                    return "Line";
                case 0x1019:
                    return "Pie";
                case 0x101A:
                    return "Area";
                case 0x101B:
                    return "Scatter";
                case 0x103A:
                    return "ThreeDimensional";
                default:
                    return null;
            }
        }

        private static LegacyXlsChartRecordKind GetKind(ushort type) {
            switch (type) {
                case 0x1001: // Units
                case 0x1002: // Chart
                case 0x1033: // Begin
                case 0x1034: // End
                case 0x1041: // ShtProps
                    return LegacyXlsChartRecordKind.Container;
                case 0x1003: // Series
                case 0x1016: // SeriesList
                case 0x1022: // ChartFormatLink
                case 0x1044: // SerToCrt
                case 0x1046: // SBaseRef
                case 0x1064: // CrErr
                case 0x1065: // SeriesFormat
                    return LegacyXlsChartRecordKind.Series;
                case 0x101D: // Axis
                case 0x101E: // Tick
                case 0x101F: // ValueRange
                case 0x1020: // CatSerRange
                case 0x1021: // AxisLineFormat
                case 0x1045: // AxesUsed
                    return LegacyXlsChartRecordKind.Axis;
                case 0x100D: // AttachedLabel
                case 0x1024: // DefaultText
                case 0x1025: // Text
                case 0x1026: // FontX
                case 0x1027: // ObjectLink
                    return LegacyXlsChartRecordKind.Text;
                case 0x1006: // DataFormat
                case 0x1007: // LineFormat
                case 0x1009: // MarkerFormat
                case 0x100A: // AreaFormat
                case 0x100B: // PieFormat
                case 0x1014: // ChartFormat
                case 0x101C: // ChartLine
                case 0x104F: // Ifmt
                case 0x105F: // Dat
                    return LegacyXlsChartRecordKind.Formatting;
                case 0x1032: // Frame
                case 0x1035: // PlotArea
                case 0x1051: // Pos
                case 0x1060: // PlotGrowth
                    return LegacyXlsChartRecordKind.Layout;
                case 0x1015: // Legend
                case 0x1017: // Bar
                case 0x1018: // Line
                case 0x1019: // Pie
                case 0x101A: // Area
                case 0x101B: // Scatter
                case 0x103A: // Chart3d
                    return LegacyXlsChartRecordKind.ChartType;
                case 0x105C: // SheetExt
                case 0x105D: // BookExt
                    return LegacyXlsChartRecordKind.Extension;
                default:
                    return LegacyXlsChartRecordKind.PreserveOnly;
            }
        }
    }
}
