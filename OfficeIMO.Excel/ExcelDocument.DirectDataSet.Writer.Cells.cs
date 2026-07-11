using System.Globalization;
using System.Data;
using System.IO.Compression;
using System.Text;
using System.Threading;

namespace OfficeIMO.Excel {
    public partial class ExcelDocument {
        private static partial class DirectDataSetWorkbookWriter {
            private static void WriteDirectValueCell(
                TextWriter writer,
                string rowReference,
                string cellReferencePrefix,
                object? value,
                string? styleAttribute,
                bool valueStyleColumn,
                DirectCellValueKind cellValueKind,
                bool useCellValueNumberFormats,
                Func<DateTimeOffset, DateTime> dateTimeOffsetWriteStrategy,
                ExcelDateSystem dateSystem,
                DirectSharedStringTable? sharedStrings) {
                if (valueStyleColumn) {
                    WriteCell(writer, rowReference, cellReferencePrefix, value, styleAttribute, valueStyleColumn, cellValueKind, useCellValueNumberFormats, dateTimeOffsetWriteStrategy, dateSystem, sharedStrings);
                } else {
                    WriteCell(writer, rowReference, cellReferencePrefix, value, styleAttribute, cellValueKind, dateTimeOffsetWriteStrategy, dateSystem, sharedStrings);
                }
            }

            internal static string? CreateStyleAttributeForValue(object? value, bool useCellValueNumberFormats) {
                switch (value) {
                    case DateTime:
                    case DateTimeOffset:
                        return useCellValueNumberFormats ? CellValueDateStyleAttribute : DateStyleAttribute;
                    case TimeSpan:
                        return useCellValueNumberFormats ? CellValueTimeStyleAttribute : TimeStyleAttribute;
#if NET6_0_OR_GREATER
                    case DateOnly:
                        return useCellValueNumberFormats ? CellValueDateStyleAttribute : DateStyleAttribute;
                    case TimeOnly:
                        return useCellValueNumberFormats ? CellValueTimeStyleAttribute : TimeStyleAttribute;
#endif
                    default:
                        return null;
                }
            }

            internal static string GetDateStyleAttribute(bool useCellValueNumberFormats)
                => useCellValueNumberFormats ? CellValueDateStyleAttribute : DateStyleAttribute;

            internal static string GetTimeStyleAttribute(bool useCellValueNumberFormats)
                => useCellValueNumberFormats ? CellValueTimeStyleAttribute : TimeStyleAttribute;

            private static void WriteCell(TextWriter writer, string rowReference, string cellReferencePrefix, object? value, string? styleAttribute, Func<DateTimeOffset, DateTime> dateTimeOffsetWriteStrategy, ExcelDateSystem dateSystem, DirectSharedStringTable? sharedStrings) {
                writer.Write(cellReferencePrefix);
                writer.Write(rowReference);
                writer.Write('"');
                if (styleAttribute != null) {
                    writer.Write(styleAttribute);
                }

                WriteCellValue(writer, value, dateTimeOffsetWriteStrategy, dateSystem, sharedStrings);
            }

            private static void WriteCell(TextWriter writer, string rowReference, string cellReferencePrefix, object? value, string? styleAttribute, DirectCellValueKind cellValueKind, Func<DateTimeOffset, DateTime> dateTimeOffsetWriteStrategy, ExcelDateSystem dateSystem, DirectSharedStringTable? sharedStrings) {
                writer.Write(cellReferencePrefix);
                writer.Write(rowReference);
                writer.Write('"');
                if (styleAttribute != null) {
                    writer.Write(styleAttribute);
                }

                WriteCellValue(writer, value, cellValueKind, dateTimeOffsetWriteStrategy, dateSystem, sharedStrings);
            }

            private static void WriteCell(TextWriter writer, string rowReference, string cellReferencePrefix, object? value, string? styleAttribute, bool useValueStyle, DirectCellValueKind cellValueKind, bool useCellValueNumberFormats, Func<DateTimeOffset, DateTime> dateTimeOffsetWriteStrategy, ExcelDateSystem dateSystem, DirectSharedStringTable? sharedStrings) {
                writer.Write(cellReferencePrefix);
                writer.Write(rowReference);
                writer.Write('"');
                string? effectiveStyleAttribute = styleAttribute ?? (useValueStyle ? CreateStyleAttributeForValue(value, useCellValueNumberFormats) : null);
                if (effectiveStyleAttribute != null) {
                    writer.Write(effectiveStyleAttribute);
                }

                if (useValueStyle) {
                    WriteCellValue(writer, value, dateTimeOffsetWriteStrategy, dateSystem, sharedStrings);
                } else {
                    WriteCellValue(writer, value, cellValueKind, dateTimeOffsetWriteStrategy, dateSystem, sharedStrings);
                }
            }

            private static void WriteCellValue(TextWriter writer, object? value, DirectCellValueKind cellValueKind, Func<DateTimeOffset, DateTime> dateTimeOffsetWriteStrategy, ExcelDateSystem dateSystem, DirectSharedStringTable? sharedStrings) {
                if (value == null || value == DBNull.Value) {
                    writer.Write(" t=\"str\"><v/></c>");
                    return;
                }

                switch (cellValueKind) {
                    case DirectCellValueKind.Formula:
                        if (value is DirectFormulaCellValue formulaValue) {
                            WriteFormulaCellValue(writer, formulaValue);
                            return;
                        }

                        break;
                    case DirectCellValueKind.Object:
                        if (value is DirectTypedCellValue typedValue) {
                            WriteTypedCellValue(writer, typedValue);
                            return;
                        }

                        break;
                    case DirectCellValueKind.String:
                        if (value is string stringValue) {
                            WriteStringCellValue(writer, stringValue, sharedStrings);
                            return;
                        }

                        break;
                    case DirectCellValueKind.Boolean:
                        if (value is bool boolValue) {
                            writer.Write(boolValue ? " t=\"b\"><v>1</v></c>" : " t=\"b\"><v>0</v></c>");
                            return;
                        }

                        break;
                    case DirectCellValueKind.DateTime:
                        if (value is DateTime dateTime) {
                            WriteRawValueCell(writer, ExcelDateSystemConverter.ToSerial(dateTime, dateSystem));
                            return;
                        }

                        break;
                    case DirectCellValueKind.DateTimeOffset:
                        if (value is DateTimeOffset dateTimeOffset) {
                            WriteDateTimeOffsetCellValue(writer, dateTimeOffset, dateTimeOffsetWriteStrategy, dateSystem);
                            return;
                        }

                        break;
                    case DirectCellValueKind.TimeSpan:
                        if (value is TimeSpan timeSpan) {
                            WriteRawValueCell(writer, timeSpan.TotalDays);
                            return;
                        }

                        break;
                    case DirectCellValueKind.Double:
                        if (value is double doubleValue) {
                            WriteRawValueCell(writer, doubleValue);
                            return;
                        }

                        break;
                    case DirectCellValueKind.Float:
                        if (value is float floatValue) {
                            WriteRawValueCell(writer, floatValue);
                            return;
                        }

                        break;
                    case DirectCellValueKind.Decimal:
                        if (value is decimal decimalValue) {
                            WriteRawValueCell(writer, decimalValue);
                            return;
                        }

                        break;
                    case DirectCellValueKind.SByte:
                        if (value is sbyte sbyteValue) {
                            WriteRawValueCell(writer, sbyteValue);
                            return;
                        }

                        break;
                    case DirectCellValueKind.Byte:
                        if (value is byte byteValue) {
                            WriteRawValueCell(writer, byteValue);
                            return;
                        }

                        break;
                    case DirectCellValueKind.Int16:
                        if (value is short shortValue) {
                            WriteRawValueCell(writer, shortValue);
                            return;
                        }

                        break;
                    case DirectCellValueKind.UInt16:
                        if (value is ushort ushortValue) {
                            WriteRawValueCell(writer, ushortValue);
                            return;
                        }

                        break;
                    case DirectCellValueKind.Int32:
                        if (value is int intValue) {
                            WriteRawValueCell(writer, intValue);
                            return;
                        }

                        break;
                    case DirectCellValueKind.UInt32:
                        if (value is uint uintValue) {
                            WriteRawValueCell(writer, uintValue);
                            return;
                        }

                        break;
                    case DirectCellValueKind.Int64:
                        if (value is long longValue) {
                            WriteRawValueCell(writer, longValue);
                            return;
                        }

                        break;
                    case DirectCellValueKind.UInt64:
                        if (value is ulong ulongValue) {
                            WriteRawValueCell(writer, ulongValue);
                            return;
                        }

                        break;
#if NET6_0_OR_GREATER
                    case DirectCellValueKind.DateOnly:
                        if (value is DateOnly dateOnly) {
                            WriteRawValueCell(writer, ExcelDateSystemConverter.ToSerial(dateOnly.ToDateTime(TimeOnly.MinValue), dateSystem));
                            return;
                        }

                        break;
                    case DirectCellValueKind.TimeOnly:
                        if (value is TimeOnly timeOnly) {
                            WriteRawValueCell(writer, timeOnly.ToTimeSpan().TotalDays);
                            return;
                        }

                        break;
#endif
                }

                WriteCellValue(writer, value, dateTimeOffsetWriteStrategy, dateSystem, sharedStrings);
            }

            internal static void WriteCellValue(TextWriter writer, object? value, Func<DateTimeOffset, DateTime> dateTimeOffsetWriteStrategy, ExcelDateSystem dateSystem, DirectSharedStringTable? sharedStrings) {
                switch (value) {
                    case null:
                    case DBNull:
                        writer.Write(" t=\"str\"><v/></c>");
                        return;
                    case DirectFormulaCellValue formulaValue:
                        WriteFormulaCellValue(writer, formulaValue);
                        return;
                    case DirectTypedCellValue typedValue:
                        WriteTypedCellValue(writer, typedValue);
                        return;
                    case string stringValue:
                        WriteStringCellValue(writer, stringValue, sharedStrings);
                        return;
                    case bool boolValue:
                        writer.Write(boolValue ? " t=\"b\"><v>1</v></c>" : " t=\"b\"><v>0</v></c>");
                        return;
                    case DateTime dateTime:
                        WriteRawValueCell(writer, ExcelDateSystemConverter.ToSerial(dateTime, dateSystem));
                        return;
                    case DateTimeOffset dateTimeOffset:
                        WriteDateTimeOffsetCellValue(writer, dateTimeOffset, dateTimeOffsetWriteStrategy, dateSystem);
                        return;
                    case TimeSpan timeSpan:
                        WriteRawValueCell(writer, timeSpan.TotalDays);
                        return;
                    case double doubleValue:
                        WriteRawValueCell(writer, doubleValue);
                        return;
                    case float floatValue:
                        WriteRawValueCell(writer, floatValue);
                        return;
                    case decimal decimalValue:
                        WriteRawValueCell(writer, decimalValue);
                        return;
                    case sbyte sbyteValue:
                        WriteRawValueCell(writer, sbyteValue);
                        return;
                    case byte byteValue:
                        WriteRawValueCell(writer, byteValue);
                        return;
                    case short shortValue:
                        WriteRawValueCell(writer, shortValue);
                        return;
                    case ushort ushortValue:
                        WriteRawValueCell(writer, ushortValue);
                        return;
                    case int intValue:
                        WriteRawValueCell(writer, intValue);
                        return;
                    case uint uintValue:
                        WriteRawValueCell(writer, uintValue);
                        return;
                    case long longValue:
                        WriteRawValueCell(writer, longValue);
                        return;
                    case ulong ulongValue:
                        WriteRawValueCell(writer, ulongValue);
                        return;
#if NET6_0_OR_GREATER
                    case DateOnly dateOnly:
                        WriteRawValueCell(writer, ExcelDateSystemConverter.ToSerial(dateOnly.ToDateTime(TimeOnly.MinValue), dateSystem));
                        return;
                    case TimeOnly timeOnly:
                        WriteRawValueCell(writer, timeOnly.ToTimeSpan().TotalDays);
                        return;
#endif
                    default:
                        WriteStringCell(writer, value.ToString() ?? string.Empty, validateLength: true);
                        return;
                }
            }

            private static void WriteFormulaCellValue(TextWriter writer, string formula) {
                writer.Write("><f>");
                WriteEscaped(writer, formula);
                writer.Write("</f></c>");
            }

            private static void WriteFormulaCellValue(TextWriter writer, DirectFormulaCellValue formula) {
                if (!string.IsNullOrEmpty(formula.FormulaXml)) {
                    writer.Write('>');
                    writer.Write(formula.FormulaXml);
                    if (formula.CachedValue != null) {
                        writer.Write("<v>");
                        WriteEscaped(writer, formula.CachedValue);
                        writer.Write("</v>");
                    }

                    writer.Write("</c>");
                    return;
                }

                WriteFormulaCellValue(writer, formula.Formula);
            }

            private static void WriteTypedCellValue(TextWriter writer, DirectTypedCellValue typed) {
                writer.Write(" t=\"");
                WriteEscaped(writer, typed.DataType);
                writer.Write("\">");
                if (!string.IsNullOrEmpty(typed.InlineStringXml)) {
                    writer.Write(typed.InlineStringXml);
                } else if (typed.Value != null) {
                    writer.Write("<v>");
                    WriteEscaped(writer, typed.Value);
                    writer.Write("</v>");
                } else {
                    writer.Write("<v/>");
                }

                writer.Write("</c>");
            }

            internal static void WriteStringCellValue(TextWriter writer, string value, DirectSharedStringTable? sharedStrings) {
                if (sharedStrings != null && sharedStrings.TryGetIndex(value, out int sharedStringIndex)) {
                    WriteSharedStringCell(writer, sharedStringIndex);
                } else {
                    WriteStringCell(writer, value, validateLength: true);
                }
            }

            internal static void WriteDateTimeOffsetCellValue(TextWriter writer, DateTimeOffset value, Func<DateTimeOffset, DateTime> dateTimeOffsetWriteStrategy, ExcelDateSystem dateSystem) {
                if (!TryGetDateTimeOffsetSerial(value, dateTimeOffsetWriteStrategy, dateSystem, out double dateTimeOffsetSerial)) {
                    WriteStringCell(writer, value.ToString("o", CultureInfo.InvariantCulture), validateLength: true);
                    return;
                }

                WriteRawValueCell(writer, dateTimeOffsetSerial);
            }

            private static bool TryGetDateTimeOffsetSerial(DateTimeOffset value, Func<DateTimeOffset, DateTime> dateTimeOffsetWriteStrategy, ExcelDateSystem dateSystem, out double serial) {
                try {
                    if (value.UtcDateTime < ExcelMinimumSupportedDateTimeOffset) {
                        serial = 0D;
                        return false;
                    }

                    serial = ExcelDateSystemConverter.ToSerial(dateTimeOffsetWriteStrategy(value), dateSystem);
                    return true;
                } catch (ArgumentException) {
                    serial = 0D;
                    return false;
                } catch (OverflowException) {
                    serial = 0D;
                    return false;
                }
            }

            private static void WriteStringCell(TextWriter writer, string text, bool validateLength) {
                if (validateLength) {
                    CoerceValueHelper.ValidateSharedStringLength(text, "value");
                }

                if (text.Length == 0) {
                    writer.Write(" t=\"inlineStr\"><is><t/></is></c>");
                    return;
                }

                writer.Write(" t=\"inlineStr\"><is><t");
                if (NeedsPreserveSpace(text)) {
                    writer.Write(" xml:space=\"preserve\"");
                }

                writer.Write('>');
                WriteSanitizedEscaped(writer, text);
                writer.Write("</t></is></c>");
            }

            private static void WriteSharedStringCell(TextWriter writer, int sharedStringIndex) {
                if ((uint)sharedStringIndex < (uint)SharedStringCellCache.Length) {
                    string? cached = SharedStringCellCache[sharedStringIndex];
                    if (cached == null) {
                        cached = " t=\"s\"><v>" + InvariantNumberText.Get(sharedStringIndex) + "</v></c>";
                        SharedStringCellCache[sharedStringIndex] = cached;
                    }

                    writer.Write(cached);
                    return;
                }

                writer.Write(" t=\"s\"><v>");
                WriteInvariant(writer, sharedStringIndex);
                writer.Write("</v></c>");
            }

            internal static void WriteRawValueCell(TextWriter writer, double value) {
                writer.Write("><v>");
                WriteInvariant(writer, value);
                writer.Write("</v></c>");
            }

            internal static void WriteRawValueCell(TextWriter writer, float value) {
                writer.Write("><v>");
                WriteInvariant(writer, value);
                writer.Write("</v></c>");
            }

            internal static void WriteRawValueCell(TextWriter writer, decimal value) {
                writer.Write("><v>");
                WriteInvariant(writer, value);
                writer.Write("</v></c>");
            }

            internal static void WriteRawValueCell(TextWriter writer, int value) {
                if (TryWriteCachedRawNonNegativeIntegerCell(writer, value)) {
                    return;
                }

                writer.Write("><v>");
                WriteInvariant(writer, value);
                writer.Write("</v></c>");
            }

            internal static void WriteRawValueCell(TextWriter writer, long value) {
                if (TryWriteCachedRawNonNegativeIntegerCell(writer, value)) {
                    return;
                }

                writer.Write("><v>");
                WriteInvariant(writer, value);
                writer.Write("</v></c>");
            }

            internal static void WriteRawValueCell(TextWriter writer, ulong value) {
                if (TryWriteCachedRawNonNegativeIntegerCell(writer, value)) {
                    return;
                }

                writer.Write("><v>");
                WriteInvariant(writer, value);
                writer.Write("</v></c>");
            }

            private static bool TryWriteCachedRawNonNegativeIntegerCell(TextWriter writer, int value) {
                var cache = RawNonNegativeIntegerCellCache;
                if ((uint)value >= (uint)cache.Length) {
                    return false;
                }

                writer.Write(cache[value]);
                return true;
            }

            private static bool TryWriteCachedRawNonNegativeIntegerCell(TextWriter writer, long value) {
                var cache = RawNonNegativeIntegerCellCache;
                if ((ulong)value >= (ulong)cache.Length) {
                    return false;
                }

                writer.Write(cache[(int)value]);
                return true;
            }

            private static bool TryWriteCachedRawNonNegativeIntegerCell(TextWriter writer, ulong value) {
                var cache = RawNonNegativeIntegerCellCache;
                if (value >= (ulong)cache.Length) {
                    return false;
                }

                writer.Write(cache[(int)value]);
                return true;
            }

            private static string[] CreateRawNonNegativeIntegerCellCache() {
                var cache = new string[CachedRawIntegerCellLimit];
                for (int i = 0; i < cache.Length; i++) {
                    cache[i] = "><v>" + InvariantNumberText.Get(i) + "</v></c>";
                }

                return cache;
            }

            private static void WriteInvariant(TextWriter writer, double value) {
#if NET6_0_OR_GREATER
                Span<char> buffer = stackalloc char[32];
                if (value.TryFormat(buffer, out int written, provider: CultureInfo.InvariantCulture)) {
                    writer.Write(buffer.Slice(0, written));
                    return;
                }
#endif
                writer.Write(value.ToString(CultureInfo.InvariantCulture));
            }

            private static void WriteColumnWidth(TextWriter writer, double value) {
#if NET6_0_OR_GREATER
                Span<char> buffer = stackalloc char[32];
                if (value.TryFormat(buffer, out int written, "0.###", CultureInfo.InvariantCulture)) {
                    writer.Write(buffer.Slice(0, written));
                    return;
                }
#endif
                writer.Write(value.ToString("0.###", CultureInfo.InvariantCulture));
            }

            private static void WriteInvariant(TextWriter writer, float value) {
#if NET6_0_OR_GREATER
                Span<char> buffer = stackalloc char[32];
                if (value.TryFormat(buffer, out int written, provider: CultureInfo.InvariantCulture)) {
                    writer.Write(buffer.Slice(0, written));
                    return;
                }
#endif
                writer.Write(value.ToString(CultureInfo.InvariantCulture));
            }

            private static void WriteInvariant(TextWriter writer, decimal value) {
#if NET6_0_OR_GREATER
                Span<char> buffer = stackalloc char[32];
                if (value.TryFormat(buffer, out int written, provider: CultureInfo.InvariantCulture)) {
                    writer.Write(buffer.Slice(0, written));
                    return;
                }
#endif
                writer.Write(value.ToString(CultureInfo.InvariantCulture));
            }

            private static void WriteInvariant(TextWriter writer, int value) {
                if (InvariantNumberText.TryGet(value, out string text)) {
                    writer.Write(text);
                    return;
                }

#if NET6_0_OR_GREATER
                Span<char> buffer = stackalloc char[16];
                if (value.TryFormat(buffer, out int written, provider: CultureInfo.InvariantCulture)) {
                    writer.Write(buffer.Slice(0, written));
                    return;
                }
#endif
                writer.Write(value.ToString(CultureInfo.InvariantCulture));
            }

            private static void WriteInvariant(TextWriter writer, long value) {
                if (InvariantNumberText.TryGet(value, out string text)) {
                    writer.Write(text);
                    return;
                }

#if NET6_0_OR_GREATER
                Span<char> buffer = stackalloc char[32];
                if (value.TryFormat(buffer, out int written, provider: CultureInfo.InvariantCulture)) {
                    writer.Write(buffer.Slice(0, written));
                    return;
                }
#endif
                writer.Write(value.ToString(CultureInfo.InvariantCulture));
            }

            private static void WriteInvariant(TextWriter writer, ulong value) {
                if (InvariantNumberText.TryGet(value, out string text)) {
                    writer.Write(text);
                    return;
                }

#if NET6_0_OR_GREATER
                Span<char> buffer = stackalloc char[32];
                if (value.TryFormat(buffer, out int written, provider: CultureInfo.InvariantCulture)) {
                    writer.Write(buffer.Slice(0, written));
                    return;
                }
#endif
                writer.Write(value.ToString(CultureInfo.InvariantCulture));
            }

            private static void WriteTextEntry(ZipArchive archive, string path, string text) {
                var entry = archive.CreateEntry(path, CompressionLevel.Fastest);
                using var stream = entry.Open();
#if NET6_0_OR_GREATER
                int byteCount = Utf8NoBom.GetByteCount(text);
                if (byteCount <= StackallocTextEntryByteLimit) {
                    Span<byte> stackBytes = stackalloc byte[byteCount];
                    int written = Utf8NoBom.GetBytes(text.AsSpan(), stackBytes);
                    stream.Write(stackBytes.Slice(0, written));
                    return;
                }
#endif
                byte[] bytes = Utf8NoBom.GetBytes(text);
                stream.Write(bytes, 0, bytes.Length);
            }

            private static void AppendEscaped(StringBuilder builder, string value) {
                int escapeIndex = IndexOfXmlEscape(value);
                if (escapeIndex < 0) {
                    builder.Append(value);
                    return;
                }

                int start = 0;
                while (escapeIndex >= 0) {
                    if (escapeIndex > start) {
                        builder.Append(value, start, escapeIndex - start);
                    }

                    AppendEscapedCharacter(builder, value[escapeIndex]);
                    start = escapeIndex + 1;
                    escapeIndex = IndexOfXmlEscape(value, start);
                }

                if (start < value.Length) {
                    builder.Append(value, start, value.Length - start);
                }
            }

            private static void WriteEscaped(TextWriter writer, string value) {
                int escapeIndex = IndexOfXmlEscape(value);
                if (escapeIndex < 0) {
                    writer.Write(value);
                    return;
                }

                int start = 0;
                while (escapeIndex >= 0) {
                    if (escapeIndex > start) {
                        WriteSlice(writer, value, start, escapeIndex - start);
                    }

                    WriteEscapedCharacter(writer, value[escapeIndex]);
                    start = escapeIndex + 1;
                    escapeIndex = IndexOfXmlEscape(value, start);
                }

                if (start < value.Length) {
                    WriteSlice(writer, value, start, value.Length - start);
                }
            }

            private static void WriteSanitizedEscaped(TextWriter writer, string value) {
                int start = 0;
                for (int i = 0; i < value.Length; i++) {
                    char current = value[i];
                    if (!IsInvalidXmlControl(current) && !IsXmlTextEscape(current)) {
                        continue;
                    }

                    if (i > start) {
                        WriteSlice(writer, value, start, i - start);
                    }

                    if (!IsInvalidXmlControl(current)) {
                        WriteEscapedTextCharacter(writer, current);
                    }

                    start = i + 1;
                }

                if (start < value.Length) {
                    WriteSlice(writer, value, start, value.Length - start);
                }
            }

            private static void WriteSlice(TextWriter writer, string value, int startIndex, int length) {
#if NET6_0_OR_GREATER
                writer.Write(value.AsSpan(startIndex, length));
#else
                writer.Write(value.Substring(startIndex, length));
#endif
            }

            private static int IndexOfXmlEscape(string value, int startIndex = 0) {
                for (int i = startIndex; i < value.Length; i++) {
                    if (IsXmlEscape(value[i])) {
                        return i;
                    }
                }

                return -1;
            }

            private static bool IsInvalidXmlControl(char value)
                => value < 0x20 && value != '\t' && value != '\n' && value != '\r';

            private static bool IsXmlEscape(char value)
                => value is '&' or '<' or '>' or '"' or '\'';

            private static bool IsXmlTextEscape(char value)
                => value is '&' or '<' or '>';

            private static bool NeedsPreserveSpace(string value) {
                return value.Length > 0 && (char.IsWhiteSpace(value[0]) || char.IsWhiteSpace(value[value.Length - 1]));
            }

            private static void AppendEscapedCharacter(StringBuilder builder, char value) {
                switch (value) {
                    case '&':
                        builder.Append("&amp;");
                        break;
                    case '<':
                        builder.Append("&lt;");
                        break;
                    case '>':
                        builder.Append("&gt;");
                        break;
                    case '"':
                        builder.Append("&quot;");
                        break;
                    case '\'':
                        builder.Append("&apos;");
                        break;
                }
            }

            private static void WriteEscapedCharacter(TextWriter writer, char value) {
                switch (value) {
                    case '&':
                        writer.Write("&amp;");
                        break;
                    case '<':
                        writer.Write("&lt;");
                        break;
                    case '>':
                        writer.Write("&gt;");
                        break;
                    case '"':
                        writer.Write("&quot;");
                        break;
                    case '\'':
                        writer.Write("&apos;");
                        break;
                }
            }

            private static void WriteEscapedTextCharacter(TextWriter writer, char value) {
                switch (value) {
                    case '&':
                        writer.Write("&amp;");
                        break;
                    case '<':
                        writer.Write("&lt;");
                        break;
                    case '>':
                        writer.Write("&gt;");
                        break;
                }
            }
        }

    }
}
