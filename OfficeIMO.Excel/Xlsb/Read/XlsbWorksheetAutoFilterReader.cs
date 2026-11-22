using DocumentFormat.OpenXml.Spreadsheet;
using OfficeIMO.Excel.Xlsb.Biff12;
using OfficeIMO.Excel.Xlsb.Model;
using static OfficeIMO.Excel.Xlsb.Biff12.XlsbAutoFilterCriteria;

namespace OfficeIMO.Excel.Xlsb.Read {
    /// <summary>Decodes worksheet AutoFilter columns: equality lists, custom, top 10, dynamic and icon criteria.</summary>
    internal static class XlsbWorksheetAutoFilterReader {
        internal static XlsbAutoFilterColumn ReadColumn(XlsbRecord record, XlsbAutoFilter autoFilter) {
            if (record.Data.Length != 6) {
                throw new InvalidDataException($"The BrtBeginFilterColumn record at offset {record.Offset} has invalid payload length {record.Data.Length}.");
            }
            var cursor = new XlsbBinaryCursor(record.Data);
            uint columnId = cursor.ReadUInt32();
            ushort flags = cursor.ReadUInt16();
            uint width = checked((uint)(autoFilter.Range.LastColumn - autoFilter.Range.FirstColumn + 1));
            if (columnId >= width || autoFilter.Columns.Any(column => column.ColumnId == columnId) || (flags & 0xFFFC) != 0) {
                throw new InvalidDataException($"The BrtBeginFilterColumn record at offset {record.Offset} contains an invalid column or reserved flags.");
            }
            return new XlsbAutoFilterColumn(columnId) {
                HasUnsupportedContent = (flags & 0x0003) != 0
            };
        }

        internal static void ReadBeginFilters(XlsbRecord record, XlsbAutoFilterColumn column) {
            if (record.Data.Length != 8) {
                throw new InvalidDataException($"The BrtBeginFilters record at offset {record.Offset} has invalid payload length {record.Data.Length}.");
            }
            var cursor = new XlsbBinaryCursor(record.Data);
            column.IncludeBlank = XlsbBinaryValueReader.ReadUInt32Boolean(cursor, record, "fBlank");
            cursor.ReadUInt32(); // unused
            if (column.HasValueList || column.Criterion != null) column.HasUnsupportedContent = true;
            column.HasValueList = true;
        }

        internal static string ReadValue(XlsbRecord record) {
            var cursor = new XlsbBinaryCursor(record.Data);
            string value = cursor.ReadWideString(255);
            if (value.Length == 0 || cursor.Remaining != 0) {
                throw new InvalidDataException($"The BrtFilter record at offset {record.Offset} contains an invalid value.");
            }
            return value;
        }

        /// <summary>
        /// Reads a custom (BrtBeginCustomFilters, BrtCustomFilter, BrtEndCustomFilters), top 10, dynamic or icon filter
        /// record into the column's criterion. Criteria that SpreadsheetML cannot state the same way mark the column
        /// unsupported, so its native records are preserved instead of projected.
        /// </summary>
        internal static void ReadCriterion(XlsbRecord record, XlsbAutoFilterColumn column) {
            switch (record.Type) {
                case BrtBeginCustomFilters: {
                    if (record.Data.Length != 4 || column.InCustomFilters) Invalid(record);
                    uint join = BitConverter.ToUInt32(record.Data, 0);
                    if (join > 1) Invalid(record);
                    StartCriterion(column, new CustomFilters { And = join == 1 ? true : null });
                    column.InCustomFilters = true;
                    break;
                }
                case BrtCustomFilter:
                    if (!column.InCustomFilters) Invalid(record);
                    ReadCustomFilter(record, column);
                    break;
                case BrtEndCustomFilters:
                    if (record.Data.Length != 0 || !column.InCustomFilters) Invalid(record);
                    column.InCustomFilters = false;
                    if (column.Criterion is CustomFilters custom && !custom.HasChildren) column.HasUnsupportedContent = true;
                    break;
                case BrtTop10Filter:
                    if (record.Data.Length != 17 || column.InCustomFilters) Invalid(record);
                    ReadTop10(record, column);
                    break;
                case BrtDynamicFilter:
                    if (record.Data.Length != 21 || column.InCustomFilters) Invalid(record);
                    ReadDynamic(record, column);
                    break;
                case BrtIconFilter:
                    if (record.Data.Length != 8 || column.InCustomFilters) Invalid(record);
                    ReadIcon(record, column);
                    break;
                default:
                    throw new ArgumentOutOfRangeException(nameof(record), record.Type, "The record is not an AutoFilter criterion.");
            }
        }

        private static void ReadCustomFilter(XlsbRecord record, XlsbAutoFilterColumn column) {
            if (record.Data.Length < 10) Invalid(record);
            byte valueType = record.Data[0];
            byte comparison = record.Data[1];
            if (!TryGetOperator(comparison, out FilterOperatorValues filterOperator)) Invalid(record);
            string? value = null;
            if (valueType == NumberValue) {
                if (record.Data.Length != 10) Invalid(record);
                value = InvariantNumberText.Get(BitConverter.ToDouble(record.Data, 2));
            } else if (valueType == StringValue) {
                var cursor = new XlsbBinaryCursor(record.Data);
                cursor.ReadBytes(10);
                value = cursor.ReadWideString(32_767);
                if (cursor.Remaining != 0) Invalid(record);
                // A string comparison whose text reads as a number would be re-encoded as a numeric one; keep it native.
                if (OfficeIMO.Core.Internal.OfficeRoundTripText.TryParse(value, System.Globalization.NumberStyles.Float, System.Globalization.CultureInfo.InvariantCulture, out double _)) value = null;
            } else if (valueType is not (0x08 or 0x0C or 0x0E) || record.Data.Length != 10) {
                Invalid(record);
            }
            // Boolean, blank and non-blank comparisons have no single SpreadsheetML customFilter value.
            if (value == null || column.Criterion is not CustomFilters custom || custom.ChildElements.Count >= 2) {
                column.HasUnsupportedContent = true;
                return;
            }
            custom.Append(new CustomFilter { Operator = filterOperator, Val = value });
        }

        private static void ReadTop10(XlsbRecord record, XlsbAutoFilterColumn column) {
            byte flags = record.Data[0];
            double value = BitConverter.ToDouble(record.Data, 1);
            double filterValue = BitConverter.ToDouble(record.Data, 9);
            bool applied = (flags & 0x04) != 0;
            StartCriterion(column, new Top10 {
                Top = (flags & 0x01) != 0,
                Percent = (flags & 0x02) != 0,
                Val = value,
                FilterValue = applied ? filterValue : null
            });
        }

        private static void ReadDynamic(XlsbRecord record, XlsbAutoFilterColumn column) {
            uint code = BitConverter.ToUInt32(record.Data, 0);
            bool applied = (record.Data[4] & 0x01) != 0;
            double value = BitConverter.ToDouble(record.Data, 5);
            double maximum = BitConverter.ToDouble(record.Data, 13);
            if (!TryGetDynamicType(code, out DynamicFilterValues type)) {
                column.HasUnsupportedContent = true;
                return;
            }
            bool carriesValue = applied && (IsAverage(code) || IsDateRange(code));
            StartCriterion(column, new DynamicFilter {
                Type = type,
                Val = carriesValue ? value : null,
                MaxVal = applied && IsDateRange(code) ? maximum : null
            });
        }

        private static void ReadIcon(XlsbRecord record, XlsbAutoFilterColumn column) {
            uint iconSet = BitConverter.ToUInt32(record.Data, 0);
            int icon = BitConverter.ToInt32(record.Data, 4);
            if (!TryGetIconSet(iconSet, out IconSetValues set, out int iconCount) || icon < 0 || icon >= iconCount) {
                column.HasUnsupportedContent = true;
                return;
            }
            StartCriterion(column, new IconFilter { IconSet = set, IconId = (uint)icon });
        }

        private static void StartCriterion(XlsbAutoFilterColumn column, DocumentFormat.OpenXml.OpenXmlElement criterion) {
            if (column.HasValueList || column.Criterion != null) {
                column.HasUnsupportedContent = true;
                return;
            }
            column.Criterion = criterion;
        }

        private static void Invalid(XlsbRecord record) =>
            throw new InvalidDataException($"The AutoFilter criterion record {record.Type} at offset {record.Offset} is invalid or out of place.");
    }
}
