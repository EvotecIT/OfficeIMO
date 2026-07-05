using System.Text;
using OfficeIMO.Excel.LegacyXls.Model;

namespace OfficeIMO.Excel.LegacyXls.Biff {
    /// <summary>
    /// Reads BIFF TxO text and formatting-run payloads from following Continue records.
    /// </summary>
    internal sealed class BiffTextObjectContinueReader {
        private readonly StringBuilder _builder = new();
        private readonly List<byte> _formattingBytes = new();
        private int _remainingCharacters;
        private int _remainingFormattingBytes;
        private IReadOnlyList<LegacyXlsCommentFormattingRun>? _formattingRuns;

        internal BiffTextObjectContinueReader(int textCharacters, int formattingRunBytes) {
            if (textCharacters < 0) {
                throw new ArgumentOutOfRangeException(nameof(textCharacters));
            }

            if (formattingRunBytes < 0) {
                throw new ArgumentOutOfRangeException(nameof(formattingRunBytes));
            }

            _remainingCharacters = textCharacters;
            _remainingFormattingBytes = formattingRunBytes;
        }

        internal string Text => _builder.ToString();

        internal IReadOnlyList<LegacyXlsCommentFormattingRun> FormattingRuns => _formattingRuns ??= ParseFormattingRuns();

        internal bool Complete => _remainingCharacters == 0 && _remainingFormattingBytes == 0;

        internal bool TryRead(byte[] payload) {
            int offset = 0;
            if (_remainingCharacters > 0) {
                if (payload.Length == 0) {
                    return false;
                }

                byte options = payload[offset++];
                bool isUtf16 = (options & 0x01) != 0;
                int bytesPerCharacter = isUtf16 ? 2 : 1;
                int availableCharacters = (payload.Length - offset) / bytesPerCharacter;
                int charactersToRead = Math.Min(_remainingCharacters, availableCharacters);
                int bytesToRead = checked(charactersToRead * bytesPerCharacter);
                if (isUtf16) {
                    _builder.Append(Encoding.Unicode.GetString(payload, offset, bytesToRead));
                } else {
                    for (int i = 0; i < bytesToRead; i++) {
                        _builder.Append((char)payload[offset + i]);
                    }
                }

                offset += bytesToRead;
                _remainingCharacters -= charactersToRead;
            }

            if (_remainingCharacters == 0 && _remainingFormattingBytes > 0) {
                int formattingBytesToRead = Math.Min(_remainingFormattingBytes, payload.Length - offset);
                for (int i = 0; i < formattingBytesToRead; i++) {
                    _formattingBytes.Add(payload[offset + i]);
                }

                _remainingFormattingBytes -= formattingBytesToRead;
            }

            return true;
        }

        private IReadOnlyList<LegacyXlsCommentFormattingRun> ParseFormattingRuns() {
            if (_formattingBytes.Count < 16 || _formattingBytes.Count % 8 != 0) {
                return Array.Empty<LegacyXlsCommentFormattingRun>();
            }

            int runCount = (_formattingBytes.Count / 8) - 1;
            if (runCount <= 0) {
                return Array.Empty<LegacyXlsCommentFormattingRun>();
            }

            var runs = new List<LegacyXlsCommentFormattingRun>(runCount);
            byte[] bytes = _formattingBytes.ToArray();
            for (int i = 0; i < runCount; i++) {
                int offset = i * 8;
                ushort startCharacter = BiffRecordReader.ReadUInt16(bytes, offset);
                ushort fontIndex = BiffRecordReader.ReadUInt16(bytes, offset + 2);
                if (startCharacter > Text.Length) {
                    continue;
                }

                if (runs.Count > 0 && startCharacter <= runs[runs.Count - 1].StartCharacter) {
                    continue;
                }

                runs.Add(new LegacyXlsCommentFormattingRun(startCharacter, fontIndex));
            }

            return runs;
        }
    }
}
