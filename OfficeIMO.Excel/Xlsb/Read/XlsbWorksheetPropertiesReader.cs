using OfficeIMO.Excel.Xlsb.Biff12;
using OfficeIMO.Excel.Xlsb.Model;

namespace OfficeIMO.Excel.Xlsb.Read {
    /// <summary>Decodes standard worksheet properties stored in BrtWsProp.</summary>
    internal static class XlsbWorksheetPropertiesReader {
        internal static XlsbWorksheetProperties Read(XlsbRecord record, XlsbImportOptions options) {
            if (record.Data.Length < 23) {
                throw new InvalidDataException($"The BrtWsProp record at offset {record.Offset} is truncated.");
            }
            var cursor = new XlsbBinaryCursor(record.Data);
            uint flags = (uint)(cursor.ReadByte() | (cursor.ReadByte() << 8) | (cursor.ReadByte() << 16));
            if ((flags & 0xFC0A06U) != 0 || (flags & 0x000010U) != 0) {
                throw new InvalidDataException($"The BrtWsProp record at offset {record.Offset} contains reserved flags or identifies a dialog sheet.");
            }

            byte colorFlags = cursor.ReadByte();
            byte colorType = (byte)(colorFlags >> 1);
            byte colorIndex = cursor.ReadByte();
            short tint = cursor.ReadInt16();
            byte red = cursor.ReadByte();
            byte green = cursor.ReadByte();
            byte blue = cursor.ReadByte();
            byte alpha = cursor.ReadByte();
            if (colorType > 3
                || (colorType == 2 && (colorFlags & 0x01) == 0)
                || (colorType == 3 && colorIndex > 11)) {
                throw new InvalidDataException($"The BrtWsProp record at offset {record.Offset} contains an invalid tab color.");
            }

            uint synchronizedRow = cursor.ReadUInt32();
            uint synchronizedColumn = cursor.ReadUInt32();
            bool synchronizeHorizontal = (flags & 0x001000U) != 0;
            bool synchronizeVertical = (flags & 0x002000U) != 0;
            if ((!synchronizeHorizontal && !synchronizeVertical
                    && (synchronizedRow != uint.MaxValue || synchronizedColumn != uint.MaxValue))
                || (synchronizedRow != uint.MaxValue && synchronizedRow >= A1.MaxRows)
                || (synchronizedColumn != uint.MaxValue && synchronizedColumn >= A1.MaxColumns)) {
                throw new InvalidDataException($"The BrtWsProp record at offset {record.Offset} contains invalid synchronized-scroll coordinates.");
            }

            string codeName = cursor.ReadWideString(Math.Min(options.MaxStringCharacters, 31));
            if (cursor.Remaining != 0) {
                throw new InvalidDataException($"The BrtWsProp record at offset {record.Offset} has {cursor.Remaining} unexpected trailing payload bytes.");
            }

            return new XlsbWorksheetProperties {
                ShowAutomaticPageBreaks = (flags & 0x000001U) != 0,
                Published = (flags & 0x000008U) != 0,
                ApplyOutlineStyles = (flags & 0x000020U) != 0,
                SummaryRowsBelow = (flags & 0x000040U) != 0,
                SummaryColumnsRight = (flags & 0x000080U) != 0,
                FitToPage = (flags & 0x000100U) != 0,
                ShowOutlineSymbols = (flags & 0x000400U) != 0,
                SynchronizeHorizontal = synchronizeHorizontal,
                SynchronizeVertical = synchronizeVertical,
                TransitionEvaluation = (flags & 0x004000U) != 0,
                TransitionEntry = (flags & 0x008000U) != 0,
                FilterMode = (flags & 0x010000U) != 0,
                CalculateConditionalFormatting = (flags & 0x020000U) != 0,
                TabColor = new XlsbColor(colorType, colorIndex, tint, red, green, blue, alpha),
                SynchronizedRow = synchronizedRow,
                SynchronizedColumn = synchronizedColumn,
                CodeName = codeName
            };
        }
    }
}
