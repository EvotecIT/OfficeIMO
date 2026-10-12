using System.Globalization;
using System.Threading;
using OfficeIMO.SharedSource.IO;

namespace OfficeIMO.Excel {
    public partial class ExcelDocument {
        /// <summary>
        /// Writes strongly typed values into the current package-native worksheet row.
        /// </summary>
        public sealed partial class ExcelTabularRowWriter {
            private readonly TextWriter _writer;
            private readonly PooledUtf8TextWriter? _bufferedWriter;
            private readonly bool _includeCellReferences;
            private readonly string[] _cellReferencePrefixes;
            private readonly string?[]? _styleAttributes;
            private readonly bool[]? _valueStyleColumns;
            private readonly bool _useCellValueNumberFormats;
            private readonly Func<DateTimeOffset, DateTime> _dateTimeOffsetWriteStrategy;
            private readonly ExcelDateSystem _dateSystem;
            private readonly DirectDataSetWorkbookWriter.DirectSharedStringTable? _sharedStrings;
            private int _rowIndex;
            private int _columnIndex;
            private string _rowReference = string.Empty;

            private ExcelTabularRowWriter(
                TextWriter writer,
                int startRowIndex,
                bool includeCellReferences,
                string[] cellReferencePrefixes,
                string?[]? styleAttributes,
                bool[]? valueStyleColumns,
                bool useCellValueNumberFormats,
                Func<DateTimeOffset, DateTime> dateTimeOffsetWriteStrategy,
                ExcelDateSystem dateSystem,
                DirectDataSetWorkbookWriter.DirectSharedStringTable? sharedStrings,
                ExcelTabularStylePlan? declaredStyles) {
                _writer = writer;
                _bufferedWriter = writer as PooledUtf8TextWriter;
                _rowIndex = startRowIndex;
                _includeCellReferences = includeCellReferences;
                _cellReferencePrefixes = cellReferencePrefixes;
                _styleAttributes = styleAttributes;
                _valueStyleColumns = valueStyleColumns;
                _useCellValueNumberFormats = useCellValueNumberFormats;
                _dateTimeOffsetWriteStrategy = dateTimeOffsetWriteStrategy;
                _dateSystem = dateSystem;
                _sharedStrings = sharedStrings;
                _declaredStyles = declaredStyles;
            }

            internal static ExcelTabularRowWriter Create(
                TextWriter writer,
                int startRowIndex,
                bool includeCellReferences,
                string[] cellReferencePrefixes,
                string?[]? styleAttributes,
                bool[]? valueStyleColumns,
                bool useCellValueNumberFormats,
                Func<DateTimeOffset, DateTime> dateTimeOffsetWriteStrategy,
                ExcelDateSystem dateSystem,
                object? sharedStrings,
                ExcelTabularStylePlan? declaredStyles = null) {
                return new ExcelTabularRowWriter(
                    writer,
                    startRowIndex,
                    includeCellReferences,
                    cellReferencePrefixes,
                    styleAttributes,
                    valueStyleColumns,
                    useCellValueNumberFormats,
                    dateTimeOffsetWriteStrategy,
                    dateSystem,
                    (DirectDataSetWorkbookWriter.DirectSharedStringTable?)sharedStrings,
                    declaredStyles);
            }

            internal void BeginRow() {
                _columnIndex = 0;
                _rowActive = true;
                _rowOpened = false;
                _rowStyle = _declaredStyles?.DefaultRowStyle;
                _nextCellStyle = null;
                if (_includeCellReferences) _rowReference = InvariantNumberText.Get(_rowIndex);
            }

            internal void EndRow() {
                if (_columnIndex != _cellReferencePrefixes.Length) {
                    throw new InvalidOperationException(
                        "The row writer produced " + _columnIndex.ToString(CultureInfo.InvariantCulture)
                        + " cells for " + _cellReferencePrefixes.Length.ToString(CultureInfo.InvariantCulture) + " headers.");
                }

                _writer.Write("</row>");
                _rowActive = false;
                _rowIndex++;
            }

            /// <summary>Writes a text cell.</summary>
            public ExcelTabularRowWriter Write(string? value) {
                BeginCell();
                if (value == null) _writer.Write(" t=\"str\"><v/></c>");
                else DirectDataSetWorkbookWriter.WriteStringCellValue(_writer, value, _sharedStrings);
                return this;
            }

            /// <summary>Writes a Boolean cell.</summary>
            public ExcelTabularRowWriter Write(bool value) {
                BeginCell();
                _writer.Write(value ? " t=\"b\"><v>1</v></c>" : " t=\"b\"><v>0</v></c>");
                return this;
            }

            /// <summary>Writes a date and time cell.</summary>
            public ExcelTabularRowWriter Write(DateTime value) {
                BeginCell(DirectDataSetWorkbookWriter.GetDateStyleAttribute(_useCellValueNumberFormats));
                DirectDataSetWorkbookWriter.WriteDateTimeSerialCell(_writer, value, _dateSystem);
                return this;
            }

            /// <summary>Writes a date, time, and offset cell.</summary>
            public ExcelTabularRowWriter Write(DateTimeOffset value) {
                BeginCell(DirectDataSetWorkbookWriter.GetDateStyleAttribute(_useCellValueNumberFormats));
                DirectDataSetWorkbookWriter.WriteDateTimeOffsetCellValue(_writer, value, _dateTimeOffsetWriteStrategy, _dateSystem);
                return this;
            }

            /// <summary>Writes a duration cell.</summary>
            public ExcelTabularRowWriter Write(TimeSpan value) { BeginCell(DirectDataSetWorkbookWriter.GetTimeStyleAttribute(_useCellValueNumberFormats)); DirectDataSetWorkbookWriter.WriteRawValueCell(_writer, value.TotalDays); return this; }
            /// <summary>Writes a double-precision numeric cell.</summary>
            public ExcelTabularRowWriter Write(double value) { BeginCell(); DirectDataSetWorkbookWriter.WriteRawValueCell(_writer, value); return this; }
            /// <summary>Writes a single-precision numeric cell.</summary>
            public ExcelTabularRowWriter Write(float value) { BeginCell(); DirectDataSetWorkbookWriter.WriteRawValueCell(_writer, value); return this; }
            /// <summary>Writes a decimal cell.</summary>
            public ExcelTabularRowWriter Write(decimal value) { BeginCell(); DirectDataSetWorkbookWriter.WriteRawValueCell(_writer, value); return this; }
            /// <summary>Writes a signed 8-bit integer cell.</summary>
            public ExcelTabularRowWriter Write(sbyte value) { BeginCell(); DirectDataSetWorkbookWriter.WriteRawValueCell(_writer, (int)value); return this; }
            /// <summary>Writes an unsigned 8-bit integer cell.</summary>
            public ExcelTabularRowWriter Write(byte value) { BeginCell(); DirectDataSetWorkbookWriter.WriteRawValueCell(_writer, (int)value); return this; }
            /// <summary>Writes a signed 16-bit integer cell.</summary>
            public ExcelTabularRowWriter Write(short value) { BeginCell(); DirectDataSetWorkbookWriter.WriteRawValueCell(_writer, (int)value); return this; }
            /// <summary>Writes an unsigned 16-bit integer cell.</summary>
            public ExcelTabularRowWriter Write(ushort value) { BeginCell(); DirectDataSetWorkbookWriter.WriteRawValueCell(_writer, (int)value); return this; }
            /// <summary>Writes a signed 32-bit integer cell.</summary>
            public ExcelTabularRowWriter Write(int value) { BeginCell(); DirectDataSetWorkbookWriter.WriteRawValueCell(_writer, value); return this; }
            /// <summary>Writes an unsigned 32-bit integer cell.</summary>
            public ExcelTabularRowWriter Write(uint value) { BeginCell(); DirectDataSetWorkbookWriter.WriteRawValueCell(_writer, (long)value); return this; }
            /// <summary>Writes a signed 64-bit integer cell.</summary>
            public ExcelTabularRowWriter Write(long value) { BeginCell(); DirectDataSetWorkbookWriter.WriteRawValueCell(_writer, value); return this; }
            /// <summary>Writes an unsigned 64-bit integer cell.</summary>
            public ExcelTabularRowWriter Write(ulong value) { BeginCell(); DirectDataSetWorkbookWriter.WriteRawValueCell(_writer, value); return this; }

            /// <summary>Writes a cell using the runtime value type.</summary>
            public ExcelTabularRowWriter Write(object? value) {
                BeginCell(value, useRuntimeStyle: true);
                DirectDataSetWorkbookWriter.WriteCellValue(_writer, value, _dateTimeOffsetWriteStrategy, _dateSystem, _sharedStrings);
                return this;
            }

            private void BeginCell(object? value = null, bool useRuntimeStyle = false) {
                BeginCellCore(useRuntimeStyle ? DirectDataSetWorkbookWriter.CreateStyleAttributeForValue(value, _useCellValueNumberFormats) : null);
            }

            private void BeginCell(string styleAttribute) {
                BeginCellCore(styleAttribute);
            }

            private void BeginCellCore(string? runtimeStyleAttribute) {
                EnsureActiveRow();
                if (_columnIndex >= _cellReferencePrefixes.Length) {
                    throw new InvalidOperationException(
                        "The row writer produced more than "
                        + _cellReferencePrefixes.Length.ToString(CultureInfo.InvariantCulture) + " cells.");
                }

                int columnIndex = _columnIndex++;
                OpenRow();
                if (_includeCellReferences) {
                    WriteReferenceStart(_cellReferencePrefixes[columnIndex], "\"");
                } else {
                    _writer.Write("<c");
                }

                string? styleAttribute;
                if (_declaredStyles != null) {
                    styleAttribute = GetDeclaredCellAttribute(columnIndex, runtimeStyleAttribute ?? _styleAttributes?[columnIndex]);
                } else {
                    styleAttribute = _styleAttributes?[columnIndex];
                    if (styleAttribute == null && ((_valueStyleColumns?[columnIndex] ?? false) || runtimeStyleAttribute != null)) styleAttribute = runtimeStyleAttribute;
                }

                if (styleAttribute != null) _writer.Write(styleAttribute);
            }

            private void WriteReferenceStart(string prefix, string suffix) {
                if (_bufferedWriter != null) {
                    _bufferedWriter.WriteFragments(prefix, _rowReference, suffix);
                } else {
                    _writer.Write(prefix);
                    _writer.Write(_rowReference);
                    _writer.Write(suffix);
                }
            }
        }

        private static partial class DirectDataSetWorkbookWriter {
            private static void WriteObjectRows(
                TextWriter writer,
                IDirectObjectRows rows,
                int startRowIndex,
                bool includeCellReferences,
                string[] cellReferencePrefixes,
                string?[]? styleAttributes,
                bool[]? valueStyleColumns,
                bool useCellValueNumberFormats,
                Func<DateTimeOffset, DateTime> dateTimeOffsetWriteStrategy,
                ExcelDateSystem dateSystem,
                DirectSharedStringTable? sharedStrings,
                CancellationToken ct,
                ExcelTabularStylePlan? declaredStyles = null) {
                var rowWriter = ExcelTabularRowWriter.Create(
                    writer,
                    startRowIndex,
                    includeCellReferences,
                    cellReferencePrefixes,
                    styleAttributes,
                    valueStyleColumns,
                    useCellValueNumberFormats,
                    dateTimeOffsetWriteStrategy,
                    dateSystem,
                    sharedStrings,
                    declaredStyles);
                rows.WriteRows(rowWriter, ct);
            }
        }
    }
}
