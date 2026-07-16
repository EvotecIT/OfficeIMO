using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Spreadsheet;
using OfficeIMO.Excel.Xlsb.Biff12;
using OfficeIMO.Excel.Xlsb.Model;

namespace OfficeIMO.Excel.Xlsb.Write {
    /// <summary>Validates and writes worksheet AutoFilter ranges and equality-list criteria.</summary>
    internal static class XlsbWorksheetAutoFilterWriter {
        private const int BrtBeginAFilter = 161;
        private const int BrtEndAFilter = 162;
        private const int BrtBeginFilterColumn = 163;
        private const int BrtEndFilterColumn = 164;
        private const int BrtBeginFilters = 165;
        private const int BrtEndFilters = 166;
        private const int BrtFilter = 167;

        internal static void Write(Stream output, AutoFilter? autoFilter, string sheetName) {
            if (output == null) throw new ArgumentNullException(nameof(output));
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

        private static IReadOnlyList<XlsbGeneratedRecord> CreateRecords(AutoFilter autoFilter, string sheetName) {
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
                Filters[] filters = column.Elements<Filters>().ToArray();
                if (column.ChildElements.Any(element => element is not Filters) || filters.Length != 1) {
                    throw new NotSupportedException($"Native XLSB generation currently supports one equality-list filters element per AutoFilter column on worksheet '{sheetName}'.");
                }

                records.Add(new XlsbGeneratedRecord(BrtBeginFilterColumn, CreateFilterColumnPayload(columnId)));
                AppendFilters(records, filters[0], sheetName);
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

        private static bool TryParseRange(string reference, out XlsbCellRange? range) {
            range = null;
            if (A1.TryParseRange(reference, out int firstRow, out int firstColumn, out int lastRow, out int lastColumn)) {
                range = new XlsbCellRange(firstRow, lastRow, firstColumn, lastColumn);
                return true;
            }
            if (A1.TryParseCellReferenceFast(reference, out int row, out int column)) {
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
            var allowed = new HashSet<string>(allowedNames, StringComparer.Ordinal);
            OpenXmlAttribute? unsupported = element.GetAttributes()
                .Cast<OpenXmlAttribute?>()
                .FirstOrDefault(attribute => attribute.HasValue
                    && !string.Equals(attribute.Value.NamespaceUri, "http://www.w3.org/2000/xmlns/", StringComparison.Ordinal)
                    && !allowed.Contains(attribute.Value.LocalName));
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
