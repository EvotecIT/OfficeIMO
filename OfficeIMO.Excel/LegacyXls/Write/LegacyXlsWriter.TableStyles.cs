using DocumentFormat.OpenXml.Spreadsheet;
using System.Text;

namespace OfficeIMO.Excel.LegacyXls.Write {
    internal static partial class LegacyXlsWriter {
        private const uint LegacyXlsBuiltInTableStyleCount = 145;

        internal static bool SupportsWorkbookTableStyles(Stylesheet? stylesheet, out string? reason) {
            reason = null;
            TableStyles? tableStyles = stylesheet?.TableStyles;
            if (tableStyles == null) {
                return true;
            }

            List<DocumentFormat.OpenXml.Spreadsheet.TableStyle> customStyles = tableStyles
                .Elements<DocumentFormat.OpenXml.Spreadsheet.TableStyle>()
                .ToList();
            if (tableStyles.ChildElements.Count != customStyles.Count) {
                reason = "table styles";
                return false;
            }

            if (!IsSupportedTableStyleNameLength(tableStyles.DefaultTableStyle?.Value)
                || !IsSupportedTableStyleNameLength(tableStyles.DefaultPivotStyle?.Value)) {
                reason = "table style name lengths outside BIFF8 limits";
                return false;
            }

            if (tableStyles.Count?.Value is uint declaredCount && declaredCount != customStyles.Count) {
                reason = "custom table style counts outside native subset";
                return false;
            }

            foreach (DocumentFormat.OpenXml.Spreadsheet.TableStyle customStyle in customStyles) {
                if (!IsSupportedCustomTableStyle(customStyle, out reason)) {
                    return false;
                }
            }

            return true;
        }

        private static IReadOnlyList<TableStyleRecord> CreateTableStyleRecords(Stylesheet? stylesheet) {
            TableStyles? tableStyles = stylesheet?.TableStyles;
            if (tableStyles == null) {
                return Array.Empty<TableStyleRecord>();
            }

            string? defaultTableStyle = NormalizeTableStyleName(tableStyles.DefaultTableStyle?.Value);
            string? defaultPivotStyle = NormalizeTableStyleName(tableStyles.DefaultPivotStyle?.Value);
            List<DocumentFormat.OpenXml.Spreadsheet.TableStyle> customStyles = tableStyles
                .Elements<DocumentFormat.OpenXml.Spreadsheet.TableStyle>()
                .ToList();
            if (defaultTableStyle == null && defaultPivotStyle == null && customStyles.Count == 0) {
                return Array.Empty<TableStyleRecord>();
            }

            var records = new List<TableStyleRecord>(1 + customStyles.Sum(style => 1 + style.Elements<TableStyleElement>().Count()));
            records.Add(new TableStyleRecord(0x088e, CreateTableStylesPayload(defaultTableStyle, defaultPivotStyle, customStyles.Count)));

            foreach (DocumentFormat.OpenXml.Spreadsheet.TableStyle customStyle in customStyles) {
                IReadOnlyList<TableStyleElement> elements = customStyle.Elements<TableStyleElement>().ToList();
                records.Add(new TableStyleRecord(0x088f, CreateTableStylePayload(customStyle, elements.Count)));
                foreach (TableStyleElement element in elements) {
                    records.Add(new TableStyleRecord(0x0890, CreateTableStyleElementPayload(element)));
                }
            }

            return records;
        }

        private static byte[] CreateTableStylesPayload(string? defaultTableStyle, string? defaultPivotStyle, int customStyleCount) {
            using var stream = new MemoryStream();
            WriteUInt16(stream, 0x088e);
            WriteUInt16(stream, 0);
            WriteUInt32(stream, 0);
            WriteUInt32(stream, 0);
            WriteUInt32(stream, checked(LegacyXlsBuiltInTableStyleCount + (uint)customStyleCount));
            WriteUInt16(stream, checked((ushort)(defaultTableStyle?.Length ?? 0)));
            WriteUInt16(stream, checked((ushort)(defaultPivotStyle?.Length ?? 0)));
            WriteUnicodeCharacters(stream, defaultTableStyle);
            WriteUnicodeCharacters(stream, defaultPivotStyle);
            return stream.ToArray();
        }

        private static byte[] CreateTableStylePayload(DocumentFormat.OpenXml.Spreadsheet.TableStyle tableStyle, int elementCount) {
            using var stream = new MemoryStream();
            WriteFutureRecordHeader(stream, 0x088f);
            ushort flags = 0;
            if (tableStyle.Pivot?.Value == true) {
                flags |= 0x0002;
            }

            if (tableStyle.Table?.Value == true) {
                flags |= 0x0004;
            }

            WriteUInt16(stream, flags);
            WriteUInt32(stream, checked((uint)elementCount));
            string name = tableStyle.Name!.Value!;
            WriteUInt16(stream, checked((ushort)name.Length));
            WriteUnicodeCharacters(stream, name);
            return stream.ToArray();
        }

        private static byte[] CreateTableStyleElementPayload(TableStyleElement element) {
            using var stream = new MemoryStream();
            WriteFutureRecordHeader(stream, 0x0890);
            if (!TryGetBiffTableStyleElementType(GetTableStyleElementTypeValue(element), out uint elementType)) {
                throw new NotSupportedException("The table style element type is outside the native XLS subset.");
            }

            WriteUInt32(stream, elementType);
            WriteUInt32(stream, element.Size?.Value ?? 0U);
            WriteUInt32(stream, element.FormatId?.Value ?? 0U);
            return stream.ToArray();
        }

        private static void WriteFutureRecordHeader(Stream stream, ushort recordType) {
            WriteUInt16(stream, recordType);
            WriteUInt16(stream, 0);
            WriteUInt32(stream, 0);
            WriteUInt32(stream, 0);
        }

        private static bool IsSupportedCustomTableStyle(DocumentFormat.OpenXml.Spreadsheet.TableStyle tableStyle, out string? reason) {
            reason = null;
            if (string.IsNullOrWhiteSpace(tableStyle.Name?.Value)) {
                reason = "custom table style names";
                return false;
            }

            if (!IsSupportedTableStyleNameLength(tableStyle.Name!.Value)) {
                reason = "custom table style name lengths outside BIFF8 limits";
                return false;
            }

            List<TableStyleElement> elements = tableStyle.Elements<TableStyleElement>().ToList();
            if (tableStyle.ChildElements.Count != elements.Count) {
                reason = "custom table style metadata";
                return false;
            }

            if (tableStyle.Count?.Value is uint declaredElementCount && declaredElementCount != elements.Count) {
                reason = "custom table style element counts outside native subset";
                return false;
            }

            foreach (TableStyleElement element in elements) {
                if (element.Type?.Value == null || !TryGetBiffTableStyleElementType(GetTableStyleElementTypeValue(element), out _)) {
                    reason = "custom table style element types outside BIFF8 limits";
                    return false;
                }
            }

            return true;
        }

        private static string GetTableStyleElementTypeValue(TableStyleElement element) {
            return element.Type?.InnerText ?? element.Type?.Value.ToString() ?? string.Empty;
        }

        private static bool TryGetBiffTableStyleElementType(string value, out uint elementType) {
            switch (value) {
                case "wholeTable":
                case "WholeTable":
                    elementType = 0x00000000;
                    return true;
                case "headerRow":
                case "HeaderRow":
                    elementType = 0x00000001;
                    return true;
                case "totalRow":
                case "TotalRow":
                    elementType = 0x00000002;
                    return true;
                case "firstColumn":
                case "FirstColumn":
                    elementType = 0x00000003;
                    return true;
                case "lastColumn":
                case "LastColumn":
                    elementType = 0x00000004;
                    return true;
                case "firstRowStripe":
                case "FirstRowStripe":
                    elementType = 0x00000005;
                    return true;
                case "secondRowStripe":
                case "SecondRowStripe":
                    elementType = 0x00000006;
                    return true;
                case "firstColumnStripe":
                case "FirstColumnStripe":
                    elementType = 0x00000007;
                    return true;
                case "secondColumnStripe":
                case "SecondColumnStripe":
                    elementType = 0x00000008;
                    return true;
                case "firstHeaderCell":
                case "FirstHeaderCell":
                    elementType = 0x00000009;
                    return true;
                case "lastHeaderCell":
                case "LastHeaderCell":
                    elementType = 0x0000000A;
                    return true;
                case "firstTotalCell":
                case "FirstTotalCell":
                    elementType = 0x0000000B;
                    return true;
                case "lastTotalCell":
                case "LastTotalCell":
                    elementType = 0x0000000C;
                    return true;
                case "firstSubtotalColumn":
                case "FirstSubtotalColumn":
                    elementType = 0x0000000D;
                    return true;
                case "secondSubtotalColumn":
                case "SecondSubtotalColumn":
                    elementType = 0x0000000E;
                    return true;
                case "thirdSubtotalColumn":
                case "ThirdSubtotalColumn":
                    elementType = 0x0000000F;
                    return true;
                case "firstSubtotalRow":
                case "FirstSubtotalRow":
                    elementType = 0x00000010;
                    return true;
                case "secondSubtotalRow":
                case "SecondSubtotalRow":
                    elementType = 0x00000011;
                    return true;
                case "thirdSubtotalRow":
                case "ThirdSubtotalRow":
                    elementType = 0x00000012;
                    return true;
                case "blankRow":
                case "BlankRow":
                    elementType = 0x00000013;
                    return true;
                case "firstColumnSubheading":
                case "FirstColumnSubheading":
                    elementType = 0x00000014;
                    return true;
                case "secondColumnSubheading":
                case "SecondColumnSubheading":
                    elementType = 0x00000015;
                    return true;
                case "thirdColumnSubheading":
                case "ThirdColumnSubheading":
                    elementType = 0x00000016;
                    return true;
                case "firstRowSubheading":
                case "FirstRowSubheading":
                    elementType = 0x00000017;
                    return true;
                case "secondRowSubheading":
                case "SecondRowSubheading":
                    elementType = 0x00000018;
                    return true;
                case "thirdRowSubheading":
                case "ThirdRowSubheading":
                    elementType = 0x00000019;
                    return true;
                case "pageFieldLabels":
                case "PageFieldLabels":
                    elementType = 0x0000001A;
                    return true;
                case "pageFieldValues":
                case "PageFieldValues":
                    elementType = 0x0000001B;
                    return true;
                default:
                    elementType = 0;
                    return false;
            }
        }

        private static string? NormalizeTableStyleName(string? value) {
            return string.IsNullOrEmpty(value) ? null : value;
        }

        private static bool IsSupportedTableStyleNameLength(string? value) {
            if (string.IsNullOrEmpty(value)) {
                return true;
            }

            long payloadLength = 20L + (long)value!.Length * 2L;
            return value.Length <= ushort.MaxValue && payloadLength <= ushort.MaxValue;
        }

        private static void WriteUnicodeCharacters(Stream stream, string? value) {
            if (string.IsNullOrEmpty(value)) {
                return;
            }

            byte[] bytes = Encoding.Unicode.GetBytes(value!);
            stream.Write(bytes, 0, bytes.Length);
        }

        private readonly struct TableStyleRecord {
            internal TableStyleRecord(ushort recordType, byte[] payload) {
                RecordType = recordType;
                Payload = payload;
            }

            internal ushort RecordType { get; }

            internal byte[] Payload { get; }
        }
    }
}
