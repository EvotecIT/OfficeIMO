using System.Data;
using System.Globalization;
using System.ComponentModel;
using System.Text;
using System.Threading;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;

namespace OfficeIMO.Excel {
    public partial class ExcelDocument {
        private sealed partial class DirectDataSetTableModel {
            internal double[] CalculateColumnWidths(
                bool includeHeaders,
                Func<DateTimeOffset, DateTime> dateTimeOffsetWriteStrategy,
                CancellationToken ct,
                IReadOnlyList<int>? columnIndexes = null) {
                int columnCount = ColumnCount;
                var widths = new double[columnCount];
                if (columnCount == 0) {
                    return widths;
                }

                int[]? selectedColumnIndexes = CreateAutoFitSelectedColumnIndexes(columnCount, columnIndexes);
                if (selectedColumnIndexes != null && selectedColumnIndexes.Length == 0) {
                    return widths;
                }

                if (includeHeaders) {
                    if (selectedColumnIndexes == null) {
                        for (int columnIndex = 0; columnIndex < columnCount; columnIndex++) {
                            widths[columnIndex] = Math.Max(widths[columnIndex], EstimateAutoFitWidth(GetColumnName(columnIndex)));
                        }
                    } else {
                        for (int i = 0; i < selectedColumnIndexes.Length; i++) {
                            int columnIndex = selectedColumnIndexes[i];
                            widths[columnIndex] = Math.Max(widths[columnIndex], EstimateAutoFitWidth(GetColumnName(columnIndex)));
                        }
                    }
                }

                AutoFitWidthKind[] widthKinds = CreateAutoFitWidthKinds();
                Dictionary<string, double>?[]? stringWidthCaches = null;
                int rowCount = RowCount;
                DataRowCollection? sourceRows = _sourceTable?.Rows;
                bool canCancel = ct.CanBeCanceled;
                if (sourceRows != null) {
                    for (int rowIndex = 0; rowIndex < rowCount; rowIndex++) {
                        if (canCancel) {
                            ct.ThrowIfCancellationRequested();
                        }

                        DataRow sourceRow = sourceRows[rowIndex];
                        if (selectedColumnIndexes == null) {
                            for (int columnIndex = 0; columnIndex < columnCount; columnIndex++) {
                                object? value = sourceRow[columnIndex];
                                widths[columnIndex] = Math.Max(
                                    widths[columnIndex],
                                    EstimateAutoFitWidth(value, widthKinds[columnIndex], dateTimeOffsetWriteStrategy, ref stringWidthCaches, columnIndex, columnCount));
                            }
                        } else {
                            for (int i = 0; i < selectedColumnIndexes.Length; i++) {
                                int columnIndex = selectedColumnIndexes[i];
                                object? value = sourceRow[columnIndex];
                                widths[columnIndex] = Math.Max(
                                    widths[columnIndex],
                                    EstimateAutoFitWidth(value, widthKinds[columnIndex], dateTimeOffsetWriteStrategy, ref stringWidthCaches, columnIndex, columnCount));
                            }
                        }
                    }
                } else if (TryGetBufferedRows(out DirectBufferedRows bufferedRows)) {
                    for (int rowIndex = 0; rowIndex < rowCount; rowIndex++) {
                        if (canCancel) {
                            ct.ThrowIfCancellationRequested();
                        }

                        object?[] bufferedRow = bufferedRows[rowIndex];
                        if (selectedColumnIndexes == null) {
                            for (int columnIndex = 0; columnIndex < columnCount; columnIndex++) {
                                object? value = bufferedRow[columnIndex];
                                widths[columnIndex] = Math.Max(
                                    widths[columnIndex],
                                    EstimateAutoFitWidth(value, widthKinds[columnIndex], dateTimeOffsetWriteStrategy, ref stringWidthCaches, columnIndex, columnCount));
                            }
                        } else {
                            for (int i = 0; i < selectedColumnIndexes.Length; i++) {
                                int columnIndex = selectedColumnIndexes[i];
                                object? value = bufferedRow[columnIndex];
                                widths[columnIndex] = Math.Max(
                                    widths[columnIndex],
                                    EstimateAutoFitWidth(value, widthKinds[columnIndex], dateTimeOffsetWriteStrategy, ref stringWidthCaches, columnIndex, columnCount));
                            }
                        }
                    }
                } else if (_exactDictionaryRows != null) {
                    string[] columnNames = CreateColumnNameArray();
                    for (int rowIndex = 0; rowIndex < rowCount; rowIndex++) {
                        if (canCancel) {
                            ct.ThrowIfCancellationRequested();
                        }

                        Dictionary<string, object?> row = _exactDictionaryRows[rowIndex];
                        if (selectedColumnIndexes == null) {
                            for (int columnIndex = 0; columnIndex < columnCount; columnIndex++) {
                                object? value = row.TryGetValue(columnNames[columnIndex], out object? dictionaryValue)
                                    ? dictionaryValue
                                    : null;
                                widths[columnIndex] = Math.Max(
                                    widths[columnIndex],
                                    EstimateAutoFitWidth(value == DBNull.Value ? null : value, widthKinds[columnIndex], dateTimeOffsetWriteStrategy, ref stringWidthCaches, columnIndex, columnCount));
                            }
                        } else {
                            for (int i = 0; i < selectedColumnIndexes.Length; i++) {
                                int columnIndex = selectedColumnIndexes[i];
                                object? value = row.TryGetValue(columnNames[columnIndex], out object? dictionaryValue)
                                    ? dictionaryValue
                                    : null;
                                widths[columnIndex] = Math.Max(
                                    widths[columnIndex],
                                    EstimateAutoFitWidth(value == DBNull.Value ? null : value, widthKinds[columnIndex], dateTimeOffsetWriteStrategy, ref stringWidthCaches, columnIndex, columnCount));
                            }
                        }
                    }
                } else if (_dictionaryRows != null) {
                    string[] columnNames = CreateColumnNameArray();
                    for (int rowIndex = 0; rowIndex < rowCount; rowIndex++) {
                        if (canCancel) {
                            ct.ThrowIfCancellationRequested();
                        }

                        IReadOnlyDictionary<string, object?> row = _dictionaryRows[rowIndex];
                        if (selectedColumnIndexes == null) {
                            for (int columnIndex = 0; columnIndex < columnCount; columnIndex++) {
                                object? value = row.TryGetValue(columnNames[columnIndex], out object? dictionaryValue)
                                    ? dictionaryValue
                                    : null;
                                widths[columnIndex] = Math.Max(
                                    widths[columnIndex],
                                    EstimateAutoFitWidth(value == DBNull.Value ? null : value, widthKinds[columnIndex], dateTimeOffsetWriteStrategy, ref stringWidthCaches, columnIndex, columnCount));
                            }
                        } else {
                            for (int i = 0; i < selectedColumnIndexes.Length; i++) {
                                int columnIndex = selectedColumnIndexes[i];
                                object? value = row.TryGetValue(columnNames[columnIndex], out object? dictionaryValue)
                                    ? dictionaryValue
                                    : null;
                                widths[columnIndex] = Math.Max(
                                    widths[columnIndex],
                                    EstimateAutoFitWidth(value == DBNull.Value ? null : value, widthKinds[columnIndex], dateTimeOffsetWriteStrategy, ref stringWidthCaches, columnIndex, columnCount));
                            }
                        }
                    }
                } else if (_legacyDictionaryRows != null) {
                    string[] columnNames = CreateColumnNameArray();
                    for (int rowIndex = 0; rowIndex < rowCount; rowIndex++) {
                        if (canCancel) {
                            ct.ThrowIfCancellationRequested();
                        }

                        System.Collections.IDictionary row = _legacyDictionaryRows[rowIndex];
                        if (selectedColumnIndexes == null) {
                            for (int columnIndex = 0; columnIndex < columnCount; columnIndex++) {
                                object? value = GetValue(rowIndex, columnIndex);
                                widths[columnIndex] = Math.Max(
                                    widths[columnIndex],
                                    EstimateAutoFitWidth(value == DBNull.Value ? null : value, widthKinds[columnIndex], dateTimeOffsetWriteStrategy, ref stringWidthCaches, columnIndex, columnCount));
                            }
                        } else {
                            for (int i = 0; i < selectedColumnIndexes.Length; i++) {
                                int columnIndex = selectedColumnIndexes[i];
                                object? value = GetValue(rowIndex, columnIndex);
                                widths[columnIndex] = Math.Max(
                                    widths[columnIndex],
                                    EstimateAutoFitWidth(value == DBNull.Value ? null : value, widthKinds[columnIndex], dateTimeOffsetWriteStrategy, ref stringWidthCaches, columnIndex, columnCount));
                            }
                        }
                    }
                } else {
                    for (int rowIndex = 0; rowIndex < rowCount; rowIndex++) {
                        if (canCancel) {
                            ct.ThrowIfCancellationRequested();
                        }

                        if (selectedColumnIndexes == null) {
                            for (int columnIndex = 0; columnIndex < columnCount; columnIndex++) {
                                object? value = GetValue(rowIndex, columnIndex);
                                widths[columnIndex] = Math.Max(
                                    widths[columnIndex],
                                    EstimateAutoFitWidth(value, widthKinds[columnIndex], dateTimeOffsetWriteStrategy, ref stringWidthCaches, columnIndex, columnCount));
                            }
                        } else {
                            for (int i = 0; i < selectedColumnIndexes.Length; i++) {
                                int columnIndex = selectedColumnIndexes[i];
                                object? value = GetValue(rowIndex, columnIndex);
                                widths[columnIndex] = Math.Max(
                                    widths[columnIndex],
                                    EstimateAutoFitWidth(value, widthKinds[columnIndex], dateTimeOffsetWriteStrategy, ref stringWidthCaches, columnIndex, columnCount));
                            }
                        }
                    }
                }

                return widths;
            }

            private static int[]? CreateAutoFitSelectedColumnIndexes(int columnCount, IReadOnlyList<int>? columnIndexes) {
                if (columnIndexes == null) {
                    return null;
                }

                var selected = new int[columnIndexes.Count];
                int selectedCount = 0;
                bool allColumnsInOrder = columnIndexes.Count == columnCount;
                for (int i = 0; i < columnIndexes.Count; i++) {
                    int columnIndex = columnIndexes[i];
                    if (columnIndex <= 0 || columnIndex > columnCount) {
                        continue;
                    }

                    allColumnsInOrder &= columnIndex == i + 1;
                    selected[selectedCount++] = columnIndex - 1;
                }

                if (selectedCount == 0) {
                    return Array.Empty<int>();
                }

                if (allColumnsInOrder) {
                    return null;
                }

                if (selectedCount != selected.Length) {
                    Array.Resize(ref selected, selectedCount);
                }

                return selected;
            }

            private IReadOnlyList<object?[]> GetBufferedRowsForReuse() {
                if (_arrayRows != null) {
                    return _arrayRows;
                }

                return _listRows!;
            }

            private AutoFitWidthKind[] CreateAutoFitWidthKinds() {
                var kinds = new AutoFitWidthKind[_columns!.Length];
                for (int i = 0; i < kinds.Length; i++) {
                    kinds[i] = GetAutoFitWidthKind(_columns[i].DataType);
                }

                return kinds;
            }

            private static AutoFitWidthKind GetAutoFitWidthKind(Type dataType) {
                if (dataType == typeof(string)) return AutoFitWidthKind.String;
                if (dataType == typeof(bool)) return AutoFitWidthKind.Boolean;
                if (dataType == typeof(DateTime)) return AutoFitWidthKind.DateTime;
                if (dataType == typeof(DateTimeOffset)) return AutoFitWidthKind.DateTimeOffset;
                if (dataType == typeof(TimeSpan)) return AutoFitWidthKind.TimeSpan;
                if (dataType == typeof(double)) return AutoFitWidthKind.Double;
                if (dataType == typeof(float)) return AutoFitWidthKind.Float;
                if (dataType == typeof(decimal)) return AutoFitWidthKind.Decimal;
                if (dataType == typeof(sbyte)) return AutoFitWidthKind.SByte;
                if (dataType == typeof(byte)) return AutoFitWidthKind.Byte;
                if (dataType == typeof(short)) return AutoFitWidthKind.Int16;
                if (dataType == typeof(ushort)) return AutoFitWidthKind.UInt16;
                if (dataType == typeof(int)) return AutoFitWidthKind.Int32;
                if (dataType == typeof(uint)) return AutoFitWidthKind.UInt32;
                if (dataType == typeof(long)) return AutoFitWidthKind.Int64;
                if (dataType == typeof(ulong)) return AutoFitWidthKind.UInt64;
#if NET6_0_OR_GREATER
                if (dataType == typeof(DateOnly)) return AutoFitWidthKind.DateOnly;
                if (dataType == typeof(TimeOnly)) return AutoFitWidthKind.TimeOnly;
#endif
                return AutoFitWidthKind.Object;
            }

            private static double EstimateAutoFitWidth(string text) {
                if (string.IsNullOrEmpty(text)) {
                    return 0D;
                }

                if (text.IndexOf('\r') < 0 && text.IndexOf('\n') < 0) {
                    return EstimateAutoFitWidthFromLength(text.Length);
                }

                int maxLineLength = 0;
                int currentLineLength = 0;
                for (int i = 0; i < text.Length; i++) {
                    char current = text[i];
                    if (current == '\r' || current == '\n') {
                        if (currentLineLength > maxLineLength) {
                            maxLineLength = currentLineLength;
                        }

                        currentLineLength = 0;
                        if (current == '\r' && i + 1 < text.Length && text[i + 1] == '\n') {
                            i++;
                        }
                    } else {
                        currentLineLength++;
                    }
                }

                if (currentLineLength > maxLineLength) {
                    maxLineLength = currentLineLength;
                }

                if (maxLineLength == 0) {
                    return 0D;
                }

                return Math.Min(255D, Math.Max(1D, maxLineLength + 2D));
            }

            private static double EstimateAutoFitWidth(object? value, Func<DateTimeOffset, DateTime> dateTimeOffsetWriteStrategy) {
                switch (value) {
                    case null:
                    case DBNull:
                        return 0D;
                    case string stringValue:
                        return EstimateAutoFitWidth(stringValue);
                    case bool boolValue:
                        return EstimateAutoFitWidthFromLength(boolValue ? 4 : 5);
                    case DateTime dateTime:
                        _ = dateTime;
                        return EstimateAutoFitWidthFromLength(16);
                    case DateTimeOffset dateTimeOffset:
                        try {
                            _ = dateTimeOffsetWriteStrategy(dateTimeOffset);
                            return EstimateAutoFitWidthFromLength(16);
                        } catch (ArgumentException) {
                            return EstimateAutoFitWidth(dateTimeOffset.ToString("o", CultureInfo.InvariantCulture));
                        } catch (OverflowException) {
                            return EstimateAutoFitWidth(dateTimeOffset.ToString("o", CultureInfo.InvariantCulture));
                        }
                    case TimeSpan timeSpan:
                        return EstimateAutoFitWidthFromLength(CountFormattedCharacters(timeSpan));
                    case double doubleValue:
                        return EstimateAutoFitWidthFromLength(CountFormattedCharacters(doubleValue));
                    case float floatValue:
                        return EstimateAutoFitWidthFromLength(CountFormattedCharacters(floatValue));
                    case decimal decimalValue:
                        return EstimateAutoFitWidthFromLength(CountFormattedCharacters(decimalValue));
                    case sbyte sbyteValue:
                        return EstimateAutoFitWidthFromLength(CountSignedIntegerCharacters(sbyteValue));
                    case byte byteValue:
                        return EstimateAutoFitWidthFromLength(CountUnsignedIntegerCharacters(byteValue));
                    case short shortValue:
                        return EstimateAutoFitWidthFromLength(CountSignedIntegerCharacters(shortValue));
                    case ushort ushortValue:
                        return EstimateAutoFitWidthFromLength(CountUnsignedIntegerCharacters(ushortValue));
                    case int intValue:
                        return EstimateAutoFitWidthFromLength(CountSignedIntegerCharacters(intValue));
                    case uint uintValue:
                        return EstimateAutoFitWidthFromLength(CountUnsignedIntegerCharacters(uintValue));
                    case long longValue:
                        return EstimateAutoFitWidthFromLength(CountSignedIntegerCharacters(longValue));
                    case ulong ulongValue:
                        return EstimateAutoFitWidthFromLength(CountUnsignedIntegerCharacters(ulongValue));
#if NET6_0_OR_GREATER
                    case DateOnly dateOnly:
                        _ = dateOnly;
                        return EstimateAutoFitWidthFromLength(10);
                    case TimeOnly timeOnly:
                        _ = timeOnly;
                        return EstimateAutoFitWidthFromLength(8);
#endif
                    case IFormattable formattable:
                        return EstimateAutoFitWidth(formattable.ToString(null, CultureInfo.InvariantCulture) ?? string.Empty);
                    default:
                        return EstimateAutoFitWidth(value.ToString() ?? string.Empty);
                }
            }

            private static double EstimateAutoFitWidth(object? value, AutoFitWidthKind widthKind, Func<DateTimeOffset, DateTime> dateTimeOffsetWriteStrategy) {
                if (value == null || value == DBNull.Value) {
                    return 0D;
                }

                switch (widthKind) {
                    case AutoFitWidthKind.String:
                        return value is string stringValue ? EstimateAutoFitWidth(stringValue) : EstimateAutoFitWidth(value, dateTimeOffsetWriteStrategy);
                    case AutoFitWidthKind.Boolean:
                        return value is bool boolValue ? EstimateAutoFitWidthFromLength(boolValue ? 4 : 5) : EstimateAutoFitWidth(value, dateTimeOffsetWriteStrategy);
                    case AutoFitWidthKind.DateTime:
                        return value is DateTime ? EstimateAutoFitWidthFromLength(16) : EstimateAutoFitWidth(value, dateTimeOffsetWriteStrategy);
                    case AutoFitWidthKind.DateTimeOffset:
                        return value is DateTimeOffset dateTimeOffsetValue ? EstimateDateTimeOffsetAutoFitWidth(dateTimeOffsetValue, dateTimeOffsetWriteStrategy) : EstimateAutoFitWidth(value, dateTimeOffsetWriteStrategy);
                    case AutoFitWidthKind.TimeSpan:
                        return value is TimeSpan timeSpanValue ? EstimateAutoFitWidthFromLength(CountFormattedCharacters(timeSpanValue)) : EstimateAutoFitWidth(value, dateTimeOffsetWriteStrategy);
                    case AutoFitWidthKind.Double:
                        return value is double doubleValue ? EstimateAutoFitWidthFromLength(CountFormattedCharacters(doubleValue)) : EstimateAutoFitWidth(value, dateTimeOffsetWriteStrategy);
                    case AutoFitWidthKind.Float:
                        return value is float floatValue ? EstimateAutoFitWidthFromLength(CountFormattedCharacters(floatValue)) : EstimateAutoFitWidth(value, dateTimeOffsetWriteStrategy);
                    case AutoFitWidthKind.Decimal:
                        return value is decimal decimalValue ? EstimateAutoFitWidthFromLength(CountFormattedCharacters(decimalValue)) : EstimateAutoFitWidth(value, dateTimeOffsetWriteStrategy);
                    case AutoFitWidthKind.SByte:
                        return value is sbyte sbyteValue ? EstimateAutoFitWidthFromLength(CountSignedIntegerCharacters(sbyteValue)) : EstimateAutoFitWidth(value, dateTimeOffsetWriteStrategy);
                    case AutoFitWidthKind.Byte:
                        return value is byte byteValue ? EstimateAutoFitWidthFromLength(CountUnsignedIntegerCharacters(byteValue)) : EstimateAutoFitWidth(value, dateTimeOffsetWriteStrategy);
                    case AutoFitWidthKind.Int16:
                        return value is short shortValue ? EstimateAutoFitWidthFromLength(CountSignedIntegerCharacters(shortValue)) : EstimateAutoFitWidth(value, dateTimeOffsetWriteStrategy);
                    case AutoFitWidthKind.UInt16:
                        return value is ushort ushortValue ? EstimateAutoFitWidthFromLength(CountUnsignedIntegerCharacters(ushortValue)) : EstimateAutoFitWidth(value, dateTimeOffsetWriteStrategy);
                    case AutoFitWidthKind.Int32:
                        return value is int intValue ? EstimateAutoFitWidthFromLength(CountSignedIntegerCharacters(intValue)) : EstimateAutoFitWidth(value, dateTimeOffsetWriteStrategy);
                    case AutoFitWidthKind.UInt32:
                        return value is uint uintValue ? EstimateAutoFitWidthFromLength(CountUnsignedIntegerCharacters(uintValue)) : EstimateAutoFitWidth(value, dateTimeOffsetWriteStrategy);
                    case AutoFitWidthKind.Int64:
                        return value is long longValue ? EstimateAutoFitWidthFromLength(CountSignedIntegerCharacters(longValue)) : EstimateAutoFitWidth(value, dateTimeOffsetWriteStrategy);
                    case AutoFitWidthKind.UInt64:
                        return value is ulong ulongValue ? EstimateAutoFitWidthFromLength(CountUnsignedIntegerCharacters(ulongValue)) : EstimateAutoFitWidth(value, dateTimeOffsetWriteStrategy);
#if NET6_0_OR_GREATER
                    case AutoFitWidthKind.DateOnly:
                        return value is DateOnly ? EstimateAutoFitWidthFromLength(10) : EstimateAutoFitWidth(value, dateTimeOffsetWriteStrategy);
                    case AutoFitWidthKind.TimeOnly:
                        return value is TimeOnly ? EstimateAutoFitWidthFromLength(8) : EstimateAutoFitWidth(value, dateTimeOffsetWriteStrategy);
#endif
                    default:
                        return EstimateAutoFitWidth(value, dateTimeOffsetWriteStrategy);
                }
            }

            private static double EstimateAutoFitWidth(
                object? value,
                AutoFitWidthKind widthKind,
                Func<DateTimeOffset, DateTime> dateTimeOffsetWriteStrategy,
                ref Dictionary<string, double>?[]? stringWidthCaches,
                int columnIndex,
                int columnCount) {
                if (value is string stringValue && stringValue.Length > 0) {
                    stringWidthCaches ??= new Dictionary<string, double>?[columnCount];
                    Dictionary<string, double>? cache = stringWidthCaches[columnIndex];
                    if (cache == null) {
                        cache = new Dictionary<string, double>(StringComparer.Ordinal);
                        stringWidthCaches[columnIndex] = cache;
                    }

                    if (cache.TryGetValue(stringValue, out double cachedWidth)) {
                        return cachedWidth;
                    }

                    double width = EstimateAutoFitWidth(value, widthKind, dateTimeOffsetWriteStrategy);
                    if (cache.Count < MaxAutoFitStringWidthCacheEntriesPerColumn) {
                        cache[stringValue] = width;
                    }

                    return width;
                }

                return EstimateAutoFitWidth(value, widthKind, dateTimeOffsetWriteStrategy);
            }

            private static double EstimateDateTimeOffsetAutoFitWidth(DateTimeOffset value, Func<DateTimeOffset, DateTime> dateTimeOffsetWriteStrategy) {
                try {
                    _ = dateTimeOffsetWriteStrategy(value);
                    return EstimateAutoFitWidthFromLength(16);
                } catch (ArgumentException) {
                    return EstimateAutoFitWidth(value.ToString("o", CultureInfo.InvariantCulture));
                } catch (OverflowException) {
                    return EstimateAutoFitWidth(value.ToString("o", CultureInfo.InvariantCulture));
                }
            }

            private static double EstimateAutoFitWidthFromLength(int length) {
                if (length <= 0) {
                    return 0D;
                }

                return Math.Min(255D, Math.Max(1D, length + 2D));
            }

            private static int CountFormattedCharacters(double value) {
#if NET6_0_OR_GREATER
                Span<char> buffer = stackalloc char[32];
                if (value.TryFormat(buffer, out int written, provider: CultureInfo.InvariantCulture)) {
                    return written;
                }
#endif
                return value.ToString(CultureInfo.InvariantCulture).Length;
            }

            private static int CountFormattedCharacters(float value) {
#if NET6_0_OR_GREATER
                Span<char> buffer = stackalloc char[32];
                if (value.TryFormat(buffer, out int written, provider: CultureInfo.InvariantCulture)) {
                    return written;
                }
#endif
                return value.ToString(CultureInfo.InvariantCulture).Length;
            }

            private static int CountFormattedCharacters(decimal value) {
#if NET6_0_OR_GREATER
                Span<char> buffer = stackalloc char[64];
                if (value.TryFormat(buffer, out int written, provider: CultureInfo.InvariantCulture)) {
                    return written;
                }
#endif
                return value.ToString(CultureInfo.InvariantCulture).Length;
            }

            private static int CountFormattedCharacters(TimeSpan value) {
#if NET6_0_OR_GREATER
                Span<char> buffer = stackalloc char[32];
                if (value.TryFormat(buffer, out int written, "c", CultureInfo.InvariantCulture)) {
                    return written;
                }
#endif
                return value.ToString("c", CultureInfo.InvariantCulture).Length;
            }

            private static int CountSignedIntegerCharacters(long value) {
                if (value < 0) {
                    ulong magnitude = (ulong)(-(value + 1)) + 1UL;
                    return 1 + CountUnsignedIntegerCharacters(magnitude);
                }

                return CountUnsignedIntegerCharacters((ulong)value);
            }

            private static int CountUnsignedIntegerCharacters(ulong value) {
                int count = 1;
                while (value >= 10UL) {
                    value /= 10UL;
                    count++;
                }

                return count;
            }
        }
    }
}
