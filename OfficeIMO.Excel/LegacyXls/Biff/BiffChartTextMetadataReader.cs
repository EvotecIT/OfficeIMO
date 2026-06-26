using OfficeIMO.Excel.LegacyXls.Model;

namespace OfficeIMO.Excel.LegacyXls.Biff {
    internal static class BiffChartTextMetadataReader {
        internal static bool TryReadDefaultText(BiffRecord record, out ushort? defaultTextId, out string? defaultTextTargetName) {
            defaultTextId = null;
            defaultTextTargetName = null;
            if (record.Type != 0x1024 || record.Payload.Length < 2) {
                return false;
            }

            ushort value = BiffRecordReader.ReadUInt16(record.Payload, 0);
            defaultTextId = value;
            defaultTextTargetName = GetDefaultTextTargetName(value);
            return true;
        }

        internal static bool TryReadText(BiffRecord record, out LegacyXlsChartText? text) {
            text = null;
            if (record.Type != 0x1025 || record.Payload.Length < 32) {
                return false;
            }

            ushort flags = BiffRecordReader.ReadUInt16(record.Payload, 24);
            ushort labelFlags = BiffRecordReader.ReadUInt16(record.Payload, 28);
            byte dataLabelPosition = checked((byte)(labelFlags & 0x000f));
            byte readingOrder = checked((byte)((labelFlags >> 14) & 0x0003));
            ushort backgroundMode = BiffRecordReader.ReadUInt16(record.Payload, 2);
            text = new LegacyXlsChartText(
                record.Payload[0],
                GetTextHorizontalAlignmentName(record.Payload[0]),
                record.Payload[1],
                GetTextVerticalAlignmentName(record.Payload[1]),
                backgroundMode,
                GetBackgroundModeName(backgroundMode),
                ReadLongRgbHex(record.Payload, 4),
                BiffRecordReader.ReadInt32(record.Payload, 8),
                BiffRecordReader.ReadInt32(record.Payload, 12),
                BiffRecordReader.ReadInt32(record.Payload, 16),
                BiffRecordReader.ReadInt32(record.Payload, 20),
                flags,
                GetTextFlagNames(flags),
                BiffRecordReader.ReadUInt16(record.Payload, 26),
                dataLabelPosition,
                GetDataLabelPositionName(dataLabelPosition),
                readingOrder,
                GetReadingOrderName(readingOrder),
                BiffRecordReader.ReadUInt16(record.Payload, 30));
            return true;
        }

        internal static bool TryReadObjectLink(BiffRecord record, out LegacyXlsChartObjectLink? objectLink) {
            objectLink = null;
            if (record.Type != 0x1027 || record.Payload.Length < 6) {
                return false;
            }

            ushort linkedObject = BiffRecordReader.ReadUInt16(record.Payload, 0);
            objectLink = new LegacyXlsChartObjectLink(
                linkedObject,
                GetObjectLinkTargetName(linkedObject),
                BiffRecordReader.ReadUInt16(record.Payload, 2),
                BiffRecordReader.ReadUInt16(record.Payload, 4));
            return true;
        }

        internal static bool TryReadLegend(BiffRecord record, out LegacyXlsChartLegend? legend) {
            legend = null;
            if (record.Type != 0x1015 || record.Payload.Length < 20) {
                return false;
            }

            legend = new LegacyXlsChartLegend(
                BiffRecordReader.ReadUInt32(record.Payload, 0),
                BiffRecordReader.ReadUInt32(record.Payload, 4),
                BiffRecordReader.ReadUInt32(record.Payload, 8),
                BiffRecordReader.ReadUInt32(record.Payload, 12),
                record.Payload[17],
                BiffRecordReader.ReadUInt16(record.Payload, 18));
            return true;
        }

        internal static bool TryReadTick(BiffRecord record, out LegacyXlsChartTick? tick) {
            tick = null;
            if (record.Type != 0x101e || record.Payload.Length < 30) {
                return false;
            }

            ushort flags = BiffRecordReader.ReadUInt16(record.Payload, 24);
            byte rotationMode = checked((byte)((flags >> 2) & 0x0007));
            byte readingOrder = checked((byte)((flags >> 14) & 0x0003));
            tick = new LegacyXlsChartTick(
                record.Payload[0],
                GetTickLocationName(record.Payload[0]),
                record.Payload[1],
                GetTickLocationName(record.Payload[1]),
                record.Payload[2],
                GetTickLabelLocationName(record.Payload[2]),
                record.Payload[3],
                GetBackgroundModeName(record.Payload[3]),
                ReadLongRgbHex(record.Payload, 4),
                flags,
                rotationMode,
                GetTickRotationModeName(rotationMode),
                (flags & 0x0001) != 0,
                (flags & 0x0002) != 0,
                (flags & 0x0020) != 0,
                readingOrder,
                GetReadingOrderName(readingOrder),
                BiffRecordReader.ReadUInt16(record.Payload, 26),
                BiffRecordReader.ReadUInt16(record.Payload, 28));
            return true;
        }

        private static string ReadLongRgbHex(byte[] bytes, int offset) {
            if (offset < 0 || offset + 3 > bytes.Length) throw new InvalidDataException("Unexpected end of BIFF chart color.");
            return "#" + bytes[offset].ToString("X2") + bytes[offset + 1].ToString("X2") + bytes[offset + 2].ToString("X2");
        }

        private static string GetDefaultTextTargetName(ushort value) {
            switch (value) {
                case 0x0000:
                    return "ChartGroupTextWithoutValueOrPercent";
                case 0x0001:
                    return "ChartGroupTextWithValueOrPercent";
                case 0x0002:
                    return "ChartUnscaledText";
                case 0x0003:
                    return "ChartScaledText";
                default:
                    return $"DefaultText:0x{value:X4}";
            }
        }

        private static string GetTextHorizontalAlignmentName(byte value) {
            switch (value) {
                case 0x01:
                    return "Left";
                case 0x02:
                    return "Center";
                case 0x03:
                    return "Right";
                case 0x04:
                    return "Justify";
                case 0x07:
                    return "Distributed";
                default:
                    return $"HorizontalAlignment:0x{value:X2}";
            }
        }

        private static string GetTextVerticalAlignmentName(byte value) {
            switch (value) {
                case 0x01:
                    return "Top";
                case 0x02:
                    return "Center";
                case 0x03:
                    return "Bottom";
                case 0x04:
                    return "Justify";
                case 0x07:
                    return "Distributed";
                default:
                    return $"VerticalAlignment:0x{value:X2}";
            }
        }

        private static string GetBackgroundModeName(ushort value) {
            switch (value) {
                case 0x0001:
                    return "Transparent";
                case 0x0002:
                    return "Opaque";
                default:
                    return $"BackgroundMode:0x{value:X4}";
            }
        }

        private static IReadOnlyList<string> GetTextFlagNames(ushort flags) {
            var names = new List<string>();
            AddFlagName(names, flags, 0x0001, "AutoColor");
            AddFlagName(names, flags, 0x0002, "ShowKey");
            AddFlagName(names, flags, 0x0004, "ShowValue");
            AddFlagName(names, flags, 0x0010, "AutoText");
            AddFlagName(names, flags, 0x0020, "Generated");
            AddFlagName(names, flags, 0x0040, "Deleted");
            AddFlagName(names, flags, 0x0080, "AutoMode");
            AddFlagName(names, flags, 0x0800, "ShowLabelAndPercent");
            AddFlagName(names, flags, 0x1000, "ShowPercent");
            AddFlagName(names, flags, 0x2000, "ShowBubbleSizes");
            AddFlagName(names, flags, 0x4000, "ShowLabel");
            return names;
        }

        private static void AddFlagName(List<string> names, ushort flags, ushort mask, string name) {
            if ((flags & mask) != 0) {
                names.Add(name);
            }
        }

        private static string GetDataLabelPositionName(byte value) {
            switch (value) {
                case 0x00:
                    return "Auto";
                case 0x01:
                    return "OutsideEnd";
                case 0x02:
                    return "InsideEnd";
                case 0x03:
                    return "Center";
                case 0x04:
                    return "InsideBase";
                case 0x05:
                    return "Above";
                case 0x06:
                    return "Below";
                case 0x07:
                    return "Left";
                case 0x08:
                    return "Right";
                case 0x09:
                    return "AutoPie";
                case 0x0a:
                    return "MovedByUser";
                default:
                    return $"DataLabelPosition:0x{value:X2}";
            }
        }

        private static string GetReadingOrderName(byte value) {
            switch (value) {
                case 0x00:
                    return "Context";
                case 0x01:
                    return "LeftToRight";
                case 0x02:
                    return "RightToLeft";
                default:
                    return $"ReadingOrder:0x{value:X2}";
            }
        }

        private static string GetObjectLinkTargetName(ushort value) {
            switch (value) {
                case 0x0001:
                    return "EntireChart";
                case 0x0002:
                    return "ValueAxis";
                case 0x0003:
                    return "CategoryAxis";
                case 0x0004:
                    return "SeriesOrDataPoint";
                case 0x0007:
                    return "SeriesAxis";
                case 0x000c:
                    return "DisplayUnitsLabels";
                default:
                    return $"ObjectLink:0x{value:X4}";
            }
        }

        private static string GetTickLocationName(byte value) {
            switch (value) {
                case 0x00:
                    return "None";
                case 0x01:
                    return "Inside";
                case 0x02:
                    return "Outside";
                case 0x03:
                    return "Crossing";
                default:
                    return $"TickLocation:0x{value:X2}";
            }
        }

        private static string GetTickLabelLocationName(byte value) {
            switch (value) {
                case 0x00:
                    return "None";
                case 0x01:
                    return "Low";
                case 0x02:
                    return "High";
                case 0x03:
                    return "NextToAxis";
                default:
                    return $"TickLabelLocation:0x{value:X2}";
            }
        }

        private static string GetTickRotationModeName(byte value) {
            switch (value) {
                case 0x00:
                    return "UseRotation";
                case 0x01:
                    return "Stacked";
                case 0x02:
                    return "RotatedCounterClockwise";
                case 0x03:
                    return "RotatedClockwise";
                default:
                    return $"TickRotation:0x{value:X2}";
            }
        }
    }
}
