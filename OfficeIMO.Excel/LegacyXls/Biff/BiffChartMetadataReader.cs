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
            TryReadAxisType(record, out ushort? axisType, out string? axisTypeName);
            TryReadAxesUsedCount(record, out ushort? axesUsedCount);
            TryReadSeries(record, out ushort? seriesCategoryDataType, out string? seriesCategoryDataTypeName, out ushort? seriesValueDataType, out ushort? seriesCategoryCount, out ushort? seriesValueCount, out ushort? seriesBubbleSizeDataType, out ushort? seriesBubbleSizeCount);
            TryReadDataFormat(record, out ushort? dataFormatPointIndex, out ushort? dataFormatSeriesIndex, out ushort? dataFormatOrder, out string? dataFormatTarget);
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
                chartHeight,
                axisType,
                axisTypeName,
                axesUsedCount,
                seriesCategoryDataType,
                seriesCategoryDataTypeName,
                seriesValueDataType,
                seriesCategoryCount,
                seriesValueCount,
                seriesBubbleSizeDataType,
                seriesBubbleSizeCount,
                dataFormatPointIndex,
                dataFormatSeriesIndex,
                dataFormatOrder,
                dataFormatTarget));
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

        private static bool TryReadAxisType(BiffRecord record, out ushort? axisType, out string? axisTypeName) {
            axisType = null;
            axisTypeName = null;
            if (record.Type != 0x101d || record.Payload.Length < 2) {
                return false;
            }

            ushort value = BiffRecordReader.ReadUInt16(record.Payload, 0);
            axisType = value;
            axisTypeName = GetAxisTypeName(value);
            return true;
        }

        private static bool TryReadAxesUsedCount(BiffRecord record, out ushort? axesUsedCount) {
            axesUsedCount = null;
            if (record.Type != 0x1045 || record.Payload.Length < 2) {
                return false;
            }

            axesUsedCount = BiffRecordReader.ReadUInt16(record.Payload, 0);
            return true;
        }

        private static bool TryReadSeries(
            BiffRecord record,
            out ushort? categoryDataType,
            out string? categoryDataTypeName,
            out ushort? valueDataType,
            out ushort? categoryCount,
            out ushort? valueCount,
            out ushort? bubbleSizeDataType,
            out ushort? bubbleSizeCount) {
            categoryDataType = null;
            categoryDataTypeName = null;
            valueDataType = null;
            categoryCount = null;
            valueCount = null;
            bubbleSizeDataType = null;
            bubbleSizeCount = null;
            if (record.Type != 0x1003 || record.Payload.Length < 12) {
                return false;
            }

            categoryDataType = BiffRecordReader.ReadUInt16(record.Payload, 0);
            categoryDataTypeName = GetSeriesCategoryDataTypeName(categoryDataType.Value);
            valueDataType = BiffRecordReader.ReadUInt16(record.Payload, 2);
            categoryCount = BiffRecordReader.ReadUInt16(record.Payload, 4);
            valueCount = BiffRecordReader.ReadUInt16(record.Payload, 6);
            bubbleSizeDataType = BiffRecordReader.ReadUInt16(record.Payload, 8);
            bubbleSizeCount = BiffRecordReader.ReadUInt16(record.Payload, 10);
            return true;
        }

        private static bool TryReadDataFormat(
            BiffRecord record,
            out ushort? pointIndex,
            out ushort? seriesIndex,
            out ushort? order,
            out string? target) {
            pointIndex = null;
            seriesIndex = null;
            order = null;
            target = null;
            if (record.Type != 0x1006 || record.Payload.Length < 8) {
                return false;
            }

            ushort xi = BiffRecordReader.ReadUInt16(record.Payload, 0);
            pointIndex = xi;
            seriesIndex = BiffRecordReader.ReadUInt16(record.Payload, 2);
            order = BiffRecordReader.ReadUInt16(record.Payload, 4);
            target = xi == 0xffff ? "Series" : "Point";
            return true;
        }

        private static string GetAxisTypeName(ushort axisType) {
            switch (axisType) {
                case 0x0000:
                    return "CategoryOrHorizontalValue";
                case 0x0001:
                    return "ValueOrVerticalValue";
                case 0x0002:
                    return "Series";
                default:
                    return $"Unknown:0x{axisType:X4}";
            }
        }

        private static string GetSeriesCategoryDataTypeName(ushort dataType) {
            switch (dataType) {
                case 0x0001:
                    return "Numeric";
                case 0x0003:
                    return "Text";
                default:
                    return $"Unknown:0x{dataType:X4}";
            }
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
