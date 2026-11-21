using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Spreadsheet;
using OfficeIMO.Excel.Xlsb.Biff12;
using OfficeIMO.Excel.Xlsb.Model;
using System.Globalization;

namespace OfficeIMO.Excel.Xlsb.Write {
    /// <summary>Validates and writes worksheet AutoFilter ranges and their equality-list, custom, top 10, dynamic and icon criteria.</summary>
    internal static class XlsbWorksheetAutoFilterWriter {
        private const int BrtBeginAFilter = 161;
        private const int BrtEndAFilter = 162;
        private const int BrtBeginFilterColumn = 163;
        private const int BrtEndFilterColumn = 164;
        private const int BrtBeginFilters = 165;
        private const int BrtEndFilters = 166;
        private const int BrtFilter = 167;

        internal static void Write(Stream output, AutoFilter? autoFilter, string sheetName) {
            ArgumentNullException.ThrowIfNull(output);
            if (autoFilter == null) return;
            IReadOnlyList<XlsbGeneratedRecord> records = CreateRecords(autoFilter, sheetName);
            foreach (XlsbGeneratedRecord record in records) {
                XlsbRecordWriter.Write(output, record.Type, record.Payload);
            }
        }

        internal static void Validate(AutoFilter? autoFilter, string sheetName) {
            if (autoFilter != null) CreateRecords(autoFilter, sheetName);
        }

        internal static bool TryGetRange(AutoFilter? autoFilter, out XlsbCellRange? range) {
            range = null;
            if (autoFilter?.Reference?.Value is not string reference) return false;
            return TryParseRange(reference, out range);
        }

        internal static IReadOnlyList<XlsbGeneratedRecord> CreateRecords(AutoFilter autoFilter, string sheetName) {
            if (autoFilter.HasChildren && autoFilter.ChildElements.Any(element => element is not FilterColumn)) {
                throw new NotSupportedException($"Native XLSB generation supports filterColumn children only in the worksheet AutoFilter on worksheet '{sheetName}'.");
            }
            EnsureOnlyAttributes(autoFilter, sheetName, "ref");
            if (!TryGetRange(autoFilter, out XlsbCellRange? range) || range == null) {
                throw new NotSupportedException($"Native XLSB generation cannot encode AutoFilter range '{autoFilter.Reference?.Value}' on worksheet '{sheetName}'.");
            }

            var records = new List<XlsbGeneratedRecord> {
                new XlsbGeneratedRecord(BrtBeginAFilter, CreateRangePayload(range))
            };
            var seenColumns = new HashSet<uint>();
            uint width = checked((uint)(range.LastColumn - range.FirstColumn + 1));
            foreach (FilterColumn column in autoFilter.Elements<FilterColumn>()) {
                EnsureOnlyAttributes(column, sheetName, "colId", "hiddenButton", "showButton");
                if (column.ColumnId?.Value is not uint columnId || columnId >= width || !seenColumns.Add(columnId)) {
                    throw new NotSupportedException($"Native XLSB generation requires unique AutoFilter column ids within the filtered range on worksheet '{sheetName}'.");
                }
                if (column.HiddenButton?.Value == true || column.ShowButton?.Value == false) {
                    throw new NotSupportedException($"Native XLSB generation does not yet support hidden or relocated AutoFilter buttons on worksheet '{sheetName}'.");
                }
                if (column.ChildElements.Count != 1) {
                    throw new NotSupportedException($"Native XLSB generation requires exactly one criterion per AutoFilter column on worksheet '{sheetName}'.");
                }

                records.Add(new XlsbGeneratedRecord(BrtBeginFilterColumn, CreateFilterColumnPayload(columnId)));
                switch (column.FirstChild) {
                    case Filters filters:
                        AppendFilters(records, filters, sheetName);
                        break;
                    case CustomFilters customFilters:
                        AppendCustomFilters(records, customFilters, sheetName);
                        break;
                    case Top10 top10:
                        records.Add(new XlsbGeneratedRecord(XlsbAutoFilterCriteria.BrtTop10Filter, CreateTop10Payload(top10, sheetName)));
                        break;
                    case DynamicFilter dynamicFilter:
                        records.Add(new XlsbGeneratedRecord(XlsbAutoFilterCriteria.BrtDynamicFilter, CreateDynamicPayload(dynamicFilter, sheetName)));
                        break;
                    case IconFilter iconFilter:
                        records.Add(new XlsbGeneratedRecord(XlsbAutoFilterCriteria.BrtIconFilter, CreateIconPayload(iconFilter, sheetName)));
                        break;
                    case ColorFilter:
                        throw new NotSupportedException($"Native XLSB generation cannot write the AutoFilter color filter on worksheet '{sheetName}' because it refers to differential formats, which native XLSB styles do not write yet.");
                    default:
                        throw new NotSupportedException($"Native XLSB generation does not support AutoFilter criterion '{column.FirstChild?.LocalName}' on worksheet '{sheetName}'.");
                }
                records.Add(new XlsbGeneratedRecord(BrtEndFilterColumn, Array.Empty<byte>()));
            }
            records.Add(new XlsbGeneratedRecord(BrtEndAFilter, Array.Empty<byte>()));
            return records.AsReadOnly();
        }

        private static void AppendFilters(List<XlsbGeneratedRecord> records, Filters filters, string sheetName) {
            EnsureOnlyAttributes(filters, sheetName, "blank");
            if (filters.ChildElements.Any(element => element is not Filter)) {
                throw new NotSupportedException($"Native XLSB generation supports text equality values only in worksheet AutoFilters on worksheet '{sheetName}'.");
            }

            using var beginPayload = new MemoryStream(8);
            WriteUInt32(beginPayload, filters.Blank?.Value == true ? 1U : 0U);
            WriteUInt32(beginPayload, 0U);
            records.Add(new XlsbGeneratedRecord(BrtBeginFilters, beginPayload.ToArray()));
            foreach (Filter filter in filters.Elements<Filter>()) {
                EnsureOnlyAttributes(filter, sheetName, "val");
                if (filter.HasChildren) {
                    throw new NotSupportedException($"Native XLSB generation does not support child content in worksheet filter values on worksheet '{sheetName}'.");
                }
                string value = filter.Val?.Value ?? string.Empty;
                if (value.Length == 0 || value.Length > 255) {
                    throw new NotSupportedException($"Native XLSB generation requires worksheet filter values between 1 and 255 characters on worksheet '{sheetName}'.");
                }
                records.Add(new XlsbGeneratedRecord(BrtFilter, CreateWideStringPayload(value)));
            }
            records.Add(new XlsbGeneratedRecord(BrtEndFilters, Array.Empty<byte>()));
        }

        private static void AppendCustomFilters(List<XlsbGeneratedRecord> records, CustomFilters customFilters, string sheetName) {
            EnsureOnlyAttributes(customFilters, sheetName, "and");
            CustomFilter[] conditions = customFilters.Elements<CustomFilter>().ToArray();
            if (conditions.Length is < 1 or > 2 || customFilters.ChildElements.Count != conditions.Length) {
                throw new NotSupportedException($"Native XLSB generation requires one or two custom AutoFilter conditions on worksheet '{sheetName}'.");
            }
            // fAnd (2.4.40) is 1 when the conditions are joined with AND. Desktop Excel reads it that way; the value
            // table in the specification lists 0 for AND, which Excel shows as an OR filter.
            records.Add(new XlsbGeneratedRecord(XlsbAutoFilterCriteria.BrtBeginCustomFilters,
                BitConverter.GetBytes(customFilters.And?.Value == true ? 1U : 0U)));
            foreach (CustomFilter condition in conditions) {
                EnsureOnlyAttributes(condition, sheetName, "operator", "val");
                if (condition.HasChildren) {
                    throw new NotSupportedException($"Native XLSB generation does not support child content in custom AutoFilter conditions on worksheet '{sheetName}'.");
                }
                string value = condition.Val?.Value
                    ?? throw new NotSupportedException($"Native XLSB generation requires a value for each custom AutoFilter condition on worksheet '{sheetName}'.");
                using var payload = new MemoryStream(10 + 4 + value.Length * 2);
                bool number = OfficeIMO.Core.Internal.OfficeRoundTripText.TryParseFinite(value, NumberStyles.Float, CultureInfo.InvariantCulture, out double numeric);
                payload.WriteByte(number ? XlsbAutoFilterCriteria.NumberValue : XlsbAutoFilterCriteria.StringValue);
                payload.WriteByte(XlsbAutoFilterCriteria.GetOperatorCode(condition.Operator?.Value ?? FilterOperatorValues.Equal));
                payload.Write(BitConverter.GetBytes(number ? numeric : 0D), 0, 8);
                if (!number) {
                    if (value.Length > 255) {
                        throw new NotSupportedException($"Native XLSB generation requires custom AutoFilter text of up to 255 characters on worksheet '{sheetName}'.");
                    }
                    byte[] text = CreateWideStringPayload(value);
                    payload.Write(text, 0, text.Length);
                }
                records.Add(new XlsbGeneratedRecord(XlsbAutoFilterCriteria.BrtCustomFilter, payload.ToArray()));
            }
            records.Add(new XlsbGeneratedRecord(XlsbAutoFilterCriteria.BrtEndCustomFilters, Array.Empty<byte>()));
        }

        private static byte[] CreateTop10Payload(Top10 top10, string sheetName) {
            EnsureOnlyAttributes(top10, sheetName, "top", "percent", "val", "filterVal");
            double value = top10.Val?.Value ?? 0D;
            if (top10.HasChildren || value < 1D || value > 500D) {
                throw new NotSupportedException($"Native XLSB generation requires a top 10 AutoFilter value from 1 to 500 on worksheet '{sheetName}'.");
            }
            double? filterValue = top10.FilterValue?.Value;
            var payload = new byte[17];
            payload[0] = (byte)((top10.Top?.Value ?? true ? 0x01 : 0)
                | (top10.Percent?.Value == true ? 0x02 : 0)
                | (filterValue.HasValue ? 0x04 : 0)); // fApplied: xNumFilter holds the computed threshold
            Buffer.BlockCopy(BitConverter.GetBytes(value), 0, payload, 1, 8);
            Buffer.BlockCopy(BitConverter.GetBytes(filterValue ?? 0D), 0, payload, 9, 8);
            return payload;
        }

        private static byte[] CreateDynamicPayload(DynamicFilter dynamicFilter, string sheetName) {
            EnsureOnlyAttributes(dynamicFilter, sheetName, "type", "val", "maxVal");
            if (dynamicFilter.HasChildren || dynamicFilter.Type?.Value is not DynamicFilterValues type
                || !XlsbAutoFilterCriteria.TryGetDynamicCode(type, out uint code)) {
                throw new NotSupportedException($"Native XLSB generation requires a known dynamic AutoFilter type on worksheet '{sheetName}'.");
            }
            double? value = dynamicFilter.Val?.Value;
            double? maximum = dynamicFilter.MaxVal?.Value;
            // fApplied marks xNumValue (and for date ranges xNumValueMax) as computed; other types carry neither.
            bool valid = XlsbAutoFilterCriteria.IsDateRange(code)
                ? value.HasValue == maximum.HasValue
                : !maximum.HasValue && (!value.HasValue || XlsbAutoFilterCriteria.IsAverage(code));
            if (!valid) {
                throw new NotSupportedException($"Native XLSB generation cannot store the computed values of dynamic AutoFilter type '{dynamicFilter.Type.InnerText}' on worksheet '{sheetName}'.");
            }
            var payload = new byte[21];
            Buffer.BlockCopy(BitConverter.GetBytes(code), 0, payload, 0, 4);
            payload[4] = value.HasValue ? (byte)0x01 : (byte)0;
            Buffer.BlockCopy(BitConverter.GetBytes(value ?? 0D), 0, payload, 5, 8);
            Buffer.BlockCopy(BitConverter.GetBytes(maximum ?? 0D), 0, payload, 13, 8);
            return payload;
        }

        private static byte[] CreateIconPayload(IconFilter iconFilter, string sheetName) {
            EnsureOnlyAttributes(iconFilter, sheetName, "iconSet", "iconId");
            IconSetValues iconSet = iconFilter.IconSet?.Value ?? IconSetValues.ThreeArrows;
            if (iconFilter.HasChildren || !XlsbAutoFilterCriteria.TryGetIconSetCode(iconSet, out uint code, out int iconCount)
                || iconFilter.IconId?.Value is not uint icon || icon >= iconCount) {
                throw new NotSupportedException($"Native XLSB generation requires an icon AutoFilter with an icon of its icon set on worksheet '{sheetName}'.");
            }
            var payload = new byte[8];
            Buffer.BlockCopy(BitConverter.GetBytes(code), 0, payload, 0, 4);
            Buffer.BlockCopy(BitConverter.GetBytes(icon), 0, payload, 4, 4);
            return payload;
        }

        private static bool TryParseRange(string reference, out XlsbCellRange? range) {
            range = null;
            if (ExcelA1.TryParseRange(reference, out int firstRow, out int firstColumn, out int lastRow, out int lastColumn)) {
                range = new XlsbCellRange(firstRow, lastRow, firstColumn, lastColumn);
                return true;
            }
            if (ExcelA1.TryParseCellReferenceFast(reference, out int row, out int column)) {
                range = new XlsbCellRange(row, row, column, column);
                return true;
            }
            return false;
        }

        private static byte[] CreateRangePayload(XlsbCellRange range) {
            using var output = new MemoryStream(16);
            WriteUInt32(output, checked((uint)(range.FirstRow - 1)));
            WriteUInt32(output, checked((uint)(range.LastRow - 1)));
            WriteUInt32(output, checked((uint)(range.FirstColumn - 1)));
            WriteUInt32(output, checked((uint)(range.LastColumn - 1)));
            return output.ToArray();
        }

        private static byte[] CreateFilterColumnPayload(uint columnId) {
            using var output = new MemoryStream(6);
            WriteUInt32(output, columnId);
            output.WriteByte(0);
            output.WriteByte(0);
            return output.ToArray();
        }

        private static byte[] CreateWideStringPayload(string value) {
            using var output = new MemoryStream(4 + value.Length * 2);
            WriteUInt32(output, checked((uint)value.Length));
            byte[] bytes = Encoding.Unicode.GetBytes(value);
            output.Write(bytes, 0, bytes.Length);
            return output.ToArray();
        }

        private static void EnsureOnlyAttributes(OpenXmlElement element, string sheetName, params string[] allowedNames) {
            OpenXmlAttribute? unsupported =
                XlsbOpenXmlAttributeValidator.FindUnsupported(element, allowedNames);
            if (unsupported.HasValue) {
                throw new NotSupportedException($"Native XLSB generation does not yet support AutoFilter attribute '{unsupported.Value.LocalName}' on worksheet '{sheetName}'.");
            }
        }

        private static void WriteUInt32(Stream output, uint value) {
            output.WriteByte((byte)value);
            output.WriteByte((byte)(value >> 8));
            output.WriteByte((byte)(value >> 16));
            output.WriteByte((byte)(value >> 24));
        }
    }
}
