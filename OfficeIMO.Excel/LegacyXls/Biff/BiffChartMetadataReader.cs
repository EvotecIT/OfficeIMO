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
            TryReadCategorySeriesRange(record, out LegacyXlsChartCategorySeriesRange? categorySeriesRange);
            TryReadAxisLineFormat(record, out LegacyXlsChartAxisLineFormat? axisLineFormat);
            TryReadSeries(record, out ushort? seriesCategoryDataType, out string? seriesCategoryDataTypeName, out ushort? seriesValueDataType, out string? seriesValueDataTypeName, out ushort? seriesCategoryCount, out ushort? seriesValueCount, out ushort? seriesBubbleSizeDataType, out string? seriesBubbleSizeDataTypeName, out ushort? seriesBubbleSizeCount);
            TryReadDataFormat(record, out ushort? dataFormatPointIndex, out ushort? dataFormatSeriesIndex, out ushort? dataFormatOrder, out string? dataFormatTarget);
            TryReadNumberFormat(record, out ushort? numberFormatId);
            TryReadFontIndex(record, out ushort? fontIndex);
            TryReadLineFormat(record, out LegacyXlsChartLineFormat? lineFormat);
            TryReadAreaFormat(record, out LegacyXlsChartAreaFormat? areaFormat);
            TryReadMarkerFormat(record, out LegacyXlsChartMarkerFormat? markerFormat);
            TryReadPieFormat(record, out LegacyXlsChartPieFormat? pieFormat);
            TryReadAttachedLabel(record, out LegacyXlsChartAttachedLabel? attachedLabel);
            BiffChartTextMetadataReader.TryReadDefaultText(record, out ushort? defaultTextId, out string? defaultTextTargetName);
            BiffChartTextMetadataReader.TryReadText(record, out LegacyXlsChartText? text);
            BiffChartTextMetadataReader.TryReadObjectLink(record, out LegacyXlsChartObjectLink? objectLink);
            BiffChartTextMetadataReader.TryReadLegend(record, out LegacyXlsChartLegend? legend);
            BiffChartTextMetadataReader.TryReadTick(record, out LegacyXlsChartTick? tick);
            TryReadValueRange(record, out LegacyXlsChartValueRange? valueRange);
            TryReadPosition(record, out LegacyXlsChartPosition? position);
            TryReadFrame(record, out LegacyXlsChartFrame? frame);
            TryReadPlotGrowth(record, out LegacyXlsChartPlotGrowth? plotGrowth);
            TryReadDataTableOptions(record, out LegacyXlsChartDataTableOptions? dataTableOptions);
            TryReadSheetProperties(record, out LegacyXlsChartSheetProperties? sheetProperties);
            TryReadBarOptions(record, out LegacyXlsChartBarOptions? barOptions);
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
                categorySeriesRange,
                axisLineFormat,
                seriesCategoryDataType,
                seriesCategoryDataTypeName,
                seriesValueDataType,
                seriesValueDataTypeName,
                seriesCategoryCount,
                seriesValueCount,
                seriesBubbleSizeDataType,
                seriesBubbleSizeDataTypeName,
                seriesBubbleSizeCount,
                dataFormatPointIndex,
                dataFormatSeriesIndex,
                dataFormatOrder,
                dataFormatTarget,
                numberFormatId,
                fontIndex,
                lineFormat,
                areaFormat,
                markerFormat,
                pieFormat,
                attachedLabel,
                defaultTextId,
                defaultTextTargetName,
                text,
                objectLink,
                legend,
                tick,
                position,
                frame,
                plotGrowth,
                dataTableOptions,
                sheetProperties,
                valueRange,
                barOptions));
            return true;
        }

        private static bool TryReadAttachedLabel(BiffRecord record, out LegacyXlsChartAttachedLabel? attachedLabel) {
            attachedLabel = null;
            if (record.Type != 0x100D || record.Payload.Length < 2) {
                return false;
            }

            attachedLabel = new LegacyXlsChartAttachedLabel(BiffRecordReader.ReadUInt16(record.Payload, 0));
            return true;
        }

        private static bool TryReadPieFormat(BiffRecord record, out LegacyXlsChartPieFormat? pieFormat) {
            pieFormat = null;
            if (record.Type != 0x100B || record.Payload.Length < 2) {
                return false;
            }

            pieFormat = new LegacyXlsChartPieFormat(BiffRecordReader.ReadInt16(record.Payload, 0));
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

        private static bool TryReadCategorySeriesRange(BiffRecord record, out LegacyXlsChartCategorySeriesRange? categorySeriesRange) {
            categorySeriesRange = null;
            if (record.Type != 0x1020 || record.Payload.Length < 8) {
                return false;
            }

            categorySeriesRange = new LegacyXlsChartCategorySeriesRange(
                BiffRecordReader.ReadInt16(record.Payload, 0),
                BiffRecordReader.ReadInt16(record.Payload, 2),
                BiffRecordReader.ReadInt16(record.Payload, 4),
                BiffRecordReader.ReadUInt16(record.Payload, 6));
            return true;
        }

        private static bool TryReadAxisLineFormat(BiffRecord record, out LegacyXlsChartAxisLineFormat? axisLineFormat) {
            axisLineFormat = null;
            if (record.Type != 0x1021 || record.Payload.Length < 2) {
                return false;
            }

            ushort targetId = BiffRecordReader.ReadUInt16(record.Payload, 0);
            axisLineFormat = new LegacyXlsChartAxisLineFormat(targetId, GetAxisLineFormatTargetName(targetId));
            return true;
        }

        private static bool TryReadSeries(
            BiffRecord record,
            out ushort? categoryDataType,
            out string? categoryDataTypeName,
            out ushort? valueDataType,
            out string? valueDataTypeName,
            out ushort? categoryCount,
            out ushort? valueCount,
            out ushort? bubbleSizeDataType,
            out string? bubbleSizeDataTypeName,
            out ushort? bubbleSizeCount) {
            categoryDataType = null;
            categoryDataTypeName = null;
            valueDataType = null;
            valueDataTypeName = null;
            categoryCount = null;
            valueCount = null;
            bubbleSizeDataType = null;
            bubbleSizeDataTypeName = null;
            bubbleSizeCount = null;
            if (record.Type != 0x1003 || record.Payload.Length < 12) {
                return false;
            }

            categoryDataType = BiffRecordReader.ReadUInt16(record.Payload, 0);
            categoryDataTypeName = GetSeriesDataTypeName(categoryDataType.Value);
            valueDataType = BiffRecordReader.ReadUInt16(record.Payload, 2);
            valueDataTypeName = GetSeriesDataTypeName(valueDataType.Value);
            categoryCount = BiffRecordReader.ReadUInt16(record.Payload, 4);
            valueCount = BiffRecordReader.ReadUInt16(record.Payload, 6);
            bubbleSizeDataType = BiffRecordReader.ReadUInt16(record.Payload, 8);
            bubbleSizeDataTypeName = GetSeriesDataTypeName(bubbleSizeDataType.Value);
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

        private static bool TryReadNumberFormat(BiffRecord record, out ushort? numberFormatId) {
            numberFormatId = null;
            if (record.Type != 0x104f || record.Payload.Length < 2) {
                return false;
            }

            numberFormatId = BiffRecordReader.ReadUInt16(record.Payload, 0);
            return true;
        }

        private static bool TryReadFontIndex(BiffRecord record, out ushort? fontIndex) {
            fontIndex = null;
            if (record.Type != 0x1026 || record.Payload.Length < 2) {
                return false;
            }

            fontIndex = BiffRecordReader.ReadUInt16(record.Payload, 0);
            return true;
        }

        private static bool TryReadLineFormat(BiffRecord record, out LegacyXlsChartLineFormat? lineFormat) {
            lineFormat = null;
            if (record.Type != 0x1007 || record.Payload.Length < 12) {
                return false;
            }

            ushort style = BiffRecordReader.ReadUInt16(record.Payload, 4);
            short weight = BiffRecordReader.ReadInt16(record.Payload, 6);
            ushort flags = BiffRecordReader.ReadUInt16(record.Payload, 8);
            lineFormat = new LegacyXlsChartLineFormat(
                ReadLongRgbHex(record.Payload, 0),
                style,
                GetLineStyleName(style),
                weight,
                GetLineWeightName(weight),
                (flags & 0x0001) != 0,
                (flags & 0x0004) != 0,
                (flags & 0x0008) != 0,
                BiffRecordReader.ReadUInt16(record.Payload, 10));
            return true;
        }

        private static bool TryReadAreaFormat(BiffRecord record, out LegacyXlsChartAreaFormat? areaFormat) {
            areaFormat = null;
            if (record.Type != 0x100a || record.Payload.Length < 16) {
                return false;
            }

            ushort pattern = BiffRecordReader.ReadUInt16(record.Payload, 8);
            ushort flags = BiffRecordReader.ReadUInt16(record.Payload, 10);
            areaFormat = new LegacyXlsChartAreaFormat(
                ReadLongRgbHex(record.Payload, 0),
                ReadLongRgbHex(record.Payload, 4),
                pattern,
                GetAreaPatternName(pattern),
                (flags & 0x0001) != 0,
                (flags & 0x0002) != 0,
                BiffRecordReader.ReadUInt16(record.Payload, 12),
                BiffRecordReader.ReadUInt16(record.Payload, 14));
            return true;
        }

        private static bool TryReadMarkerFormat(BiffRecord record, out LegacyXlsChartMarkerFormat? markerFormat) {
            markerFormat = null;
            if (record.Type != 0x1009 || record.Payload.Length < 20) {
                return false;
            }

            ushort markerType = BiffRecordReader.ReadUInt16(record.Payload, 8);
            ushort flags = BiffRecordReader.ReadUInt16(record.Payload, 10);
            markerFormat = new LegacyXlsChartMarkerFormat(
                ReadLongRgbHex(record.Payload, 0),
                ReadLongRgbHex(record.Payload, 4),
                markerType,
                GetMarkerTypeName(markerType),
                (flags & 0x0001) != 0,
                (flags & 0x0010) != 0,
                (flags & 0x0020) != 0,
                BiffRecordReader.ReadUInt16(record.Payload, 12),
                BiffRecordReader.ReadUInt16(record.Payload, 14),
                BiffRecordReader.ReadUInt32(record.Payload, 16));
            return true;
        }

        private static bool TryReadPosition(BiffRecord record, out LegacyXlsChartPosition? position) {
            position = null;
            if (record.Type != 0x1051 || record.Payload.Length < 20) {
                return false;
            }

            ushort topLeftMode = BiffRecordReader.ReadUInt16(record.Payload, 0);
            ushort bottomRightMode = BiffRecordReader.ReadUInt16(record.Payload, 2);
            position = new LegacyXlsChartPosition(
                topLeftMode,
                GetPositionModeName(topLeftMode),
                bottomRightMode,
                GetPositionModeName(bottomRightMode),
                BiffRecordReader.ReadInt16(record.Payload, 4),
                BiffRecordReader.ReadInt16(record.Payload, 8),
                BiffRecordReader.ReadInt16(record.Payload, 12),
                BiffRecordReader.ReadInt16(record.Payload, 16));
            return true;
        }

        private static bool TryReadValueRange(BiffRecord record, out LegacyXlsChartValueRange? valueRange) {
            valueRange = null;
            if (record.Type != 0x101f || record.Payload.Length < 42) {
                return false;
            }

            valueRange = new LegacyXlsChartValueRange(
                BiffRecordReader.ReadDouble(record.Payload, 0),
                BiffRecordReader.ReadDouble(record.Payload, 8),
                BiffRecordReader.ReadDouble(record.Payload, 16),
                BiffRecordReader.ReadDouble(record.Payload, 24),
                BiffRecordReader.ReadDouble(record.Payload, 32),
                BiffRecordReader.ReadUInt16(record.Payload, 40));
            return true;
        }

        private static bool TryReadFrame(BiffRecord record, out LegacyXlsChartFrame? frame) {
            frame = null;
            if (record.Type != 0x1032 || record.Payload.Length < 4) {
                return false;
            }

            ushort frameType = BiffRecordReader.ReadUInt16(record.Payload, 0);
            ushort flags = BiffRecordReader.ReadUInt16(record.Payload, 2);
            frame = new LegacyXlsChartFrame(
                frameType,
                GetFrameTypeName(frameType),
                flags,
                (flags & 0x0001) != 0,
                (flags & 0x0002) != 0);
            return true;
        }

        private static bool TryReadPlotGrowth(BiffRecord record, out LegacyXlsChartPlotGrowth? plotGrowth) {
            plotGrowth = null;
            if (record.Type != 0x1064 || record.Payload.Length < 8) {
                return false;
            }

            plotGrowth = new LegacyXlsChartPlotGrowth(
                BiffRecordReader.ReadInt16(record.Payload, 0),
                BiffRecordReader.ReadUInt16(record.Payload, 2),
                BiffRecordReader.ReadInt16(record.Payload, 4),
                BiffRecordReader.ReadUInt16(record.Payload, 6));
            return true;
        }

        private static bool TryReadDataTableOptions(BiffRecord record, out LegacyXlsChartDataTableOptions? dataTableOptions) {
            dataTableOptions = null;
            if (record.Type != 0x1063 || record.Payload.Length < 2) {
                return false;
            }

            dataTableOptions = new LegacyXlsChartDataTableOptions(BiffRecordReader.ReadUInt16(record.Payload, 0));
            return true;
        }

        private static bool TryReadBarOptions(BiffRecord record, out LegacyXlsChartBarOptions? barOptions) {
            barOptions = null;
            if (record.Type != 0x1017 || record.Payload.Length < 6) {
                return false;
            }

            barOptions = new LegacyXlsChartBarOptions(
                BiffRecordReader.ReadInt16(record.Payload, 0),
                BiffRecordReader.ReadUInt16(record.Payload, 2),
                BiffRecordReader.ReadUInt16(record.Payload, 4));
            return true;
        }

        private static bool TryReadSheetProperties(BiffRecord record, out LegacyXlsChartSheetProperties? sheetProperties) {
            sheetProperties = null;
            if (record.Type != 0x1041 || record.Payload.Length < 4) {
                return false;
            }

            sheetProperties = new LegacyXlsChartSheetProperties(
                BiffRecordReader.ReadUInt16(record.Payload, 0),
                record.Payload[2]);
            return true;
        }

        private static string ReadLongRgbHex(byte[] bytes, int offset) {
            if (offset < 0 || offset + 3 > bytes.Length) throw new InvalidDataException("Unexpected end of BIFF chart color.");
            return "#" + bytes[offset].ToString("X2") + bytes[offset + 1].ToString("X2") + bytes[offset + 2].ToString("X2");
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

        private static string GetSeriesDataTypeName(ushort dataType) {
            switch (dataType) {
                case 0x0001:
                    return "Numeric";
                case 0x0003:
                    return "Text";
                default:
                    return $"Unknown:0x{dataType:X4}";
            }
        }

        private static string GetLineStyleName(ushort style) {
            switch (style) {
                case 0x0000:
                    return "Solid";
                case 0x0001:
                    return "Dash";
                case 0x0002:
                    return "Dot";
                case 0x0003:
                    return "DashDot";
                case 0x0004:
                    return "DashDotDot";
                case 0x0005:
                    return "None";
                case 0x0006:
                    return "DarkGrayPattern";
                case 0x0007:
                    return "MediumGrayPattern";
                case 0x0008:
                    return "LightGrayPattern";
                default:
                    return $"Unknown:0x{style:X4}";
            }
        }

        private static string GetLineWeightName(short weight) {
            switch (weight) {
                case -1:
                    return "Hairline";
                case 0:
                    return "Narrow";
                case 1:
                    return "Medium";
                case 2:
                    return "Wide";
                default:
                    return $"Unknown:0x{unchecked((ushort)weight):X4}";
            }
        }

        private static string GetAreaPatternName(ushort pattern) {
            switch (pattern) {
                case 0x0000:
                    return "None";
                case 0x0001:
                    return "Solid";
                case 0x0002:
                    return "MediumGray";
                case 0x0003:
                    return "DarkGray";
                case 0x0004:
                    return "LightGray";
                case 0x0005:
                    return "HorizontalStripes";
                case 0x0006:
                    return "VerticalStripes";
                case 0x0007:
                    return "DownwardDiagonalStripes";
                case 0x0008:
                    return "UpwardDiagonalStripes";
                case 0x0009:
                    return "Grid";
                case 0x000a:
                    return "Trellis";
                case 0x000b:
                    return "LightHorizontalStripes";
                case 0x000c:
                    return "LightVerticalStripes";
                case 0x000d:
                    return "LightDown";
                case 0x000e:
                    return "LightUp";
                case 0x000f:
                    return "LightGrid";
                case 0x0010:
                    return "LightTrellis";
                case 0x0011:
                    return "Gray125";
                case 0x0012:
                    return "Gray0625";
                default:
                    return $"Unknown:0x{pattern:X4}";
            }
        }

        private static string GetMarkerTypeName(ushort markerType) {
            switch (markerType) {
                case 0x0000:
                    return "None";
                case 0x0001:
                    return "Square";
                case 0x0002:
                    return "Diamond";
                case 0x0003:
                    return "Triangle";
                case 0x0004:
                    return "SquareWithX";
                case 0x0005:
                    return "SquareWithAsterisk";
                case 0x0006:
                    return "ShortBar";
                case 0x0007:
                    return "LongBar";
                case 0x0008:
                    return "Circle";
                case 0x0009:
                    return "SquareWithPlus";
                default:
                    return $"Unknown:0x{markerType:X4}";
            }
        }

        private static string GetPositionModeName(ushort mode) {
            switch (mode) {
                case 0x0000:
                    return "MDFX";
                case 0x0001:
                    return "MDABS";
                case 0x0002:
                    return "MDPARENT";
                case 0x0003:
                    return "MDKTH";
                case 0x0005:
                    return "MDCHART";
                default:
                    return $"Unknown:0x{mode:X4}";
            }
        }

        private static string GetFrameTypeName(ushort frameType) {
            switch (frameType) {
                case 0x0000:
                    return "Frame";
                case 0x0004:
                    return "ShadowFrame";
                default:
                    return $"Unknown:0x{frameType:X4}";
            }
        }

        private static string GetAxisLineFormatTargetName(ushort targetId) {
            switch (targetId) {
                case 0x0000:
                    return "AxisLine";
                case 0x0001:
                    return "MajorGridlines";
                case 0x0002:
                    return "MinorGridlines";
                case 0x0003:
                    return "WallsOrFloor3D";
                default:
                    return $"Unknown:0x{targetId:X4}";
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
                case 0x105F:
                    return "ThreeDimensionalBarShape";
                case 0x1061:
                    return "BarOfPieOrPieOfPie";
                case 0x1067:
                    return "CustomBarOfPieOrPieOfPie";
                default:
                    return null;
            }
        }

        private static LegacyXlsChartRecordKind GetKind(ushort type) {
            switch (type) {
                case 0x0850: // ChartFrtInfo
                case 0x0852: // StartBlock
                case 0x0853: // EndBlock
                    return LegacyXlsChartRecordKind.FutureMetadata;
                case 0x0856: // CatLab
                    return LegacyXlsChartRecordKind.Axis;
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
                case 0x1065: // SIIndex
                    return LegacyXlsChartRecordKind.Series;
                case 0x101D: // Axis
                case 0x101E: // Tick
                case 0x101F: // ValueRange
                case 0x1020: // CatSerRange
                case 0x1021: // AxisLineFormat
                case 0x1045: // AxesUsed
                case 0x1062: // AxcExt
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
                case 0x105B: // SerAuxErrBar
                case 0x105C: // ClrtClient
                case 0x105D: // SerFmt
                case 0x1063: // Dat
                case 0x1066: // GelFrame
                    return LegacyXlsChartRecordKind.Formatting;
                case 0x1032: // Frame
                case 0x1035: // PlotArea
                case 0x1051: // Pos
                case 0x1064: // PlotGrowth
                    return LegacyXlsChartRecordKind.Layout;
                case 0x1015: // Legend
                case 0x1017: // Bar
                case 0x1018: // Line
                case 0x1019: // Pie
                case 0x101A: // Area
                case 0x101B: // Scatter
                case 0x103A: // Chart3d
                case 0x105F: // Chart3DBarShape
                case 0x1061: // BopPop
                case 0x1067: // BopPopCustom
                    return LegacyXlsChartRecordKind.ChartType;
                case 0x1060: // Fbi
                case 0x1068: // Fbi2
                    return LegacyXlsChartRecordKind.Text;
                default:
                    return LegacyXlsChartRecordKind.PreserveOnly;
            }
        }
    }
}
