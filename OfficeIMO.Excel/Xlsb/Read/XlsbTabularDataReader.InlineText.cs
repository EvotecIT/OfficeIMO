using System.Buffers.Binary;
using System.Text;

namespace OfficeIMO.Excel.Xlsb.Read {
    internal sealed partial class XlsbTabularDataReader {
#if NET8_0_OR_GREATER
        private InlineTextSlice[] _inlineTextSlices = null!;
#endif

        /// <summary>
        /// Keeps a validated inline string in the owned worksheet until canonical text is needed.
        /// Projected readers retain their ordinary eager decoding behavior.
        /// </summary>
        private void StoreInlineText(int ordinal, byte[] bytes, int position) {
#if NET8_0_OR_GREATER
            if (_inlineTextSlices.Length != 0 && !_options.InferSchema
                && _options.CellValueConverter == null) {
                int characterCount = checked((int)BinaryPrimitives.ReadUInt32LittleEndian(
                    bytes.AsSpan(position, sizeof(uint))));
                _inlineTextSlices[ordinal] = new InlineTextSlice(
                    position + sizeof(uint), checked(characterCount * sizeof(char)));
                return;
            }
#endif
            _strings[ordinal] = ReadValidatedWideString(bytes, position);
        }

        /// <summary>
        /// Materializes canonical text once; returned strings remain independent of worksheet rentals.
        /// </summary>
        private string GetCanonicalText(int ordinal) {
            string? value = _strings[ordinal];
#if NET8_0_OR_GREATER
            if (value == null && _inlineTextSlices.Length != 0) {
                InlineTextSlice slice = _inlineTextSlices[ordinal];
                if (slice.IsPresent) {
                    value = Encoding.Unicode.GetString(
                        _worksheetPart.Buffer, slice.Offset, slice.ByteCount);
                    _strings[ordinal] = value;
                }
            }
#endif
            return value!;
        }

#if NET8_0_OR_GREATER
        private readonly struct InlineTextSlice {
            internal InlineTextSlice(int offset, int byteCount) {
                Offset = offset;
                ByteCount = byteCount;
            }

            // Inline payloads follow a record and cell header, so zero denotes no slice.
            internal bool IsPresent => Offset != 0;
            internal int Offset { get; }
            internal int ByteCount { get; }
        }
#endif
    }
}
