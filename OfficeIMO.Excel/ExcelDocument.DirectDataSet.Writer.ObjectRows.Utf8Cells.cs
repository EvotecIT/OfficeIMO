#if NET8_0_OR_GREATER
using System.Globalization;
using System.Text;
using OfficeIMO.SharedSource.IO;

namespace OfficeIMO.Excel {
    public partial class ExcelDocument {
        public sealed partial class ExcelTabularRowWriter {
            // Text up to this length is escaped straight into the UTF-8 buffer; its worst case stays well inside it.
            private const int Utf8TextLimit = 4096;
            // Room for a number or date cell after its start tag, the longest being a 17-digit double with exponent.
            private const int Utf8NumberBytes = 48;
            // On .NET 8 and later, cells go straight into the worksheet's UTF-8 buffer instead of through characters.
            private readonly PooledUtf8TextWriter? _utf8;
            private readonly byte[] _rowDigits = new byte[16];
            private int _rowDigitCount;

            /// <summary>Writes the start of the next cell and reserves <paramref name="valueBytes"/> more after it.</summary>
            /// <returns>The buffer; <paramref name="start"/> is where the cell begins and <paramref name="position"/> where its value goes.</returns>
            private byte[] BeginUtf8Cell(string? runtimeStyleAttribute, int valueBytes, out int start, out int position) {
                EnsureActiveRow();
                if (_columnIndex >= _cellReferencePrefixes.Length) ThrowTooManyCells();
                int columnIndex = _columnIndex++;
                OpenRow();
                string? style = ResolveCellStyle(columnIndex, runtimeStyleAttribute);
                string prefix = _cellReferencePrefixes[columnIndex];
                byte[] buffer = _utf8!.ReserveUtf8(prefix.Length + _rowDigitCount + (style?.Length ?? 0) + 4 + valueBytes, out start);
                if (_includeCellReferences) {
                    position = WriteUtf8Reference(buffer, start, prefix);
                    buffer[position++] = (byte)'"';
                } else {
                    buffer[start] = (byte)'<';
                    buffer[start + 1] = (byte)'c';
                    position = start + 2;
                }
                if (style != null) position += Utf8Ascii(style, buffer, position);
                return buffer;
            }

            private void WriteUtf8Cell(string? runtimeStyleAttribute, ReadOnlySpan<byte> rest) {
                byte[] buffer = BeginUtf8Cell(runtimeStyleAttribute, rest.Length, out int start, out int position);
                rest.CopyTo(buffer.AsSpan(position));
                _utf8!.CommitUtf8(position + rest.Length - start);
            }

            private void WriteUtf8Number(int value) {
                byte[] buffer = BeginUtf8Cell(null, Utf8NumberBytes, out int start, out int position);
                position += DirectDataSetWorkbookWriter.WriteUtf8NumberCell(buffer.AsSpan(position), value);
                _utf8!.CommitUtf8(position - start);
            }

            private void WriteUtf8Number(long value) {
                byte[] buffer = BeginUtf8Cell(null, Utf8NumberBytes, out int start, out int position);
                position += DirectDataSetWorkbookWriter.WriteUtf8NumberCell(buffer.AsSpan(position), value);
                _utf8!.CommitUtf8(position - start);
            }

            private void WriteUtf8Number(double value) {
                byte[] buffer = BeginUtf8Cell(null, Utf8NumberBytes, out int start, out int position);
                position += DirectDataSetWorkbookWriter.WriteUtf8NumberCell(buffer.AsSpan(position), value);
                _utf8!.CommitUtf8(position - start);
            }

            private void WriteUtf8DateTime(DateTime value) {
                byte[] buffer = BeginUtf8Cell(DirectDataSetWorkbookWriter.GetDateStyleAttribute(_useCellValueNumberFormats),
                    Utf8NumberBytes, out int start, out int position);
                position += DirectDataSetWorkbookWriter.WriteUtf8DateTimeCell(buffer.AsSpan(position), value, _dateSystem);
                _utf8!.CommitUtf8(position - start);
            }

            /// <summary>
            /// Writes printable ASCII text as an inline string. Other text (controls, non-ASCII, underscores that may
            /// need x-string escaping), shared strings and long values return false for the character path.
            /// </summary>
            private bool TryWriteUtf8Text(string? value) {
                if (value == null) {
                    WriteUtf8Cell(null, " t=\"str\"><v/></c>"u8);
                    return true;
                }
                if (_sharedStrings != null || value.Length > Utf8TextLimit) return false;
                ReadOnlySpan<char> text = value.AsSpan();
                if (text.IndexOfAnyExceptInRange(' ', '~') >= 0 || text.IndexOf('_') >= 0) return false;
                if (text.IsEmpty) {
                    WriteUtf8Cell(null, " t=\"inlineStr\"><is><t/></is></c>"u8);
                    return true;
                }
                // Each character takes at most five bytes, as &amp;, plus 56 bytes of markup.
                byte[] buffer = BeginUtf8Cell(null, text.Length * 5 + 64, out int start, out int position);
                Span<byte> output = buffer.AsSpan();
                " t=\"inlineStr\"><is><t"u8.CopyTo(output.Slice(position));
                position += 21;
                if (text[0] == ' ' || text[text.Length - 1] == ' ') {
                    " xml:space=\"preserve\""u8.CopyTo(output.Slice(position));
                    position += 21;
                }
                output[position++] = (byte)'>';
                while (!text.IsEmpty) {
                    int special = text.IndexOfAny('&', '<', '>');
                    ReadOnlySpan<char> plain = special < 0 ? text : text.Slice(0, special);
                    position += Encoding.UTF8.GetBytes(plain, output.Slice(position));
                    if (special < 0) break;
                    ReadOnlySpan<byte> entity = text[special] switch {
                        '&' => "&amp;"u8,
                        '<' => "&lt;"u8,
                        _ => "&gt;"u8
                    };
                    entity.CopyTo(output.Slice(position));
                    position += entity.Length;
                    text = text.Slice(special + 1);
                }
                "</t></is></c>"u8.CopyTo(output.Slice(position));
                position += 13;
                _utf8!.CommitUtf8(position - start);
                return true;
            }

            /// <summary>Writes a reference prefix such as <c>&lt;c r="B</c> followed by the row number; returns the next position.</summary>
            private int WriteUtf8Reference(byte[] buffer, int position, string prefix) {
                position += Utf8Ascii(prefix, buffer, position);
                _rowDigits.AsSpan(0, _rowDigitCount).CopyTo(buffer.AsSpan(position));
                return position + _rowDigitCount;
            }

            /// <summary>Copies markup the writer composed itself, which is always ASCII.</summary>
            private static int Utf8Ascii(string value, byte[] buffer, int position) =>
                Encoding.UTF8.GetBytes(value, buffer.AsSpan(position));
        }

        private static partial class DirectDataSetWorkbookWriter {
            /// <summary>Writes <c>&gt;&lt;v&gt;</c>, the invariant digits and the closing tags of a numeric cell.</summary>
            internal static int WriteUtf8NumberCell(Span<byte> output, long value) {
                "><v>"u8.CopyTo(output);
                value.TryFormat(output.Slice(4), out int written, default, CultureInfo.InvariantCulture);
                return CloseUtf8Value(output, 4 + written);
            }

            /// <summary>Writes a double as <see cref="WriteRawValueCell(TextWriter, double)"/> does.</summary>
            internal static int WriteUtf8NumberCell(Span<byte> output, double value) {
                if (TryGetTwoDecimalScaledValue(value, out long scaledInteger)) {
                    "><v>"u8.CopyTo(output);
                    int position = 4;
                    ulong magnitude;
                    if (scaledInteger < 0) {
                        output[position++] = (byte)'-';
                        magnitude = (ulong)(-scaledInteger);
                    } else {
                        magnitude = (ulong)scaledInteger;
                    }
                    (magnitude / 100UL).TryFormat(output.Slice(position), out int written, default, CultureInfo.InvariantCulture);
                    position += written;
                    uint fraction = (uint)(magnitude % 100UL);
                    if (fraction != 0) {
                        output[position++] = (byte)'.';
                        output[position++] = (byte)('0' + (fraction / 10U));
                        uint ones = fraction % 10U;
                        if (ones != 0) output[position++] = (byte)('0' + ones);
                    }
                    return CloseUtf8Value(output, position);
                }
                if (double.IsNaN(value) || double.IsInfinity(value)) {
                    return Encoding.UTF8.GetBytes(NonFiniteNumberErrorCell, output);
                }
                "><v>"u8.CopyTo(output);
                value.TryFormat(output.Slice(4), out int digits, "R", CultureInfo.InvariantCulture);
                return CloseUtf8Value(output, 4 + digits);
            }

            /// <summary>Writes a date serial as <see cref="WriteDateTimeSerialCell"/> does.</summary>
            internal static int WriteUtf8DateTimeCell(Span<byte> output, DateTime value, ExcelDateSystem dateSystem) {
                DateTime fastPathMinimum = dateSystem == ExcelDateSystem.NineteenFour
                    ? NineteenFourFastDateSerialMinimum
                    : NineteenHundredFastDateSerialMinimum;
                if (value < fastPathMinimum) return WriteUtf8NumberCell(output, ExcelDateSystemConverter.ToSerial(value, dateSystem));
                const long oaEpochDay = 693_593L;
                const ulong millisecondsPerDay = 86_400_000UL;
                long ticks = value.Ticks;
                int daySerial = checked((int)((ticks / TimeSpan.TicksPerDay) - oaEpochDay));
                if (dateSystem == ExcelDateSystem.NineteenFour) daySerial -= (int)ExcelDateSystemConverter.Date1904OffsetDays;
                if (daySerial == 0) return WriteUtf8NumberCell(output, ExcelDateSystemConverter.ToSerial(value, dateSystem));

                long timeTicks = ticks % TimeSpan.TicksPerDay;
                int dayDigits = CountDigits(daySerial);
                ulong fraction = 0;
                int fractionDigitCount = 0;
                ulong millisecondsAfterMidnight = (ulong)(timeTicks / TimeSpan.TicksPerMillisecond);
                if (millisecondsAfterMidnight != 0) {
                    fractionDigitCount = 17 - dayDigits;
                    ulong fractionScale = DecimalPowers[fractionDigitCount];
                    UInt128 scaledMilliseconds = (UInt128)millisecondsAfterMidnight * fractionScale;
                    fraction = (ulong)(scaledMilliseconds / millisecondsPerDay);
                    if ((scaledMilliseconds % millisecondsPerDay) * 2 >= millisecondsPerDay) fraction++;
                    if (fraction == fractionScale) {
                        daySerial++;
                        fraction = 0;
                    }
                }

                "><v>"u8.CopyTo(output);
                daySerial.TryFormat(output.Slice(4), out int written, default, CultureInfo.InvariantCulture);
                int position = 4 + written;
                if (fraction != 0) {
                    Span<byte> digits = output.Slice(position + 1, fractionDigitCount);
                    for (int index = fractionDigitCount - 1; index >= 0; index--) {
                        digits[index] = (byte)('0' + (fraction % 10));
                        fraction /= 10;
                    }
                    int digitCount = fractionDigitCount;
                    while (digitCount > 0 && digits[digitCount - 1] == (byte)'0') digitCount--;
                    output[position] = (byte)'.';
                    position += 1 + digitCount;
                }
                return CloseUtf8Value(output, position);
            }

            private static int CountDigits(int value) {
                int digits = 1;
                while (value >= 10) {
                    value /= 10;
                    digits++;
                }
                return digits;
            }

            private static int CloseUtf8Value(Span<byte> output, int position) {
                "</v></c>"u8.CopyTo(output.Slice(position));
                return position + 8;
            }
        }
    }
}
#endif
