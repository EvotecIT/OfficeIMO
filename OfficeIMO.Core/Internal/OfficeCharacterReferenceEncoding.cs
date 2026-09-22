using System;
using System.Globalization;
using System.IO;
using System.Text;

namespace OfficeIMO.Core.Internal;

/// <summary>Escapes Unicode scalar values that a strict legacy text encoding cannot represent.</summary>
internal static class OfficeCharacterReferenceEncoding {
    internal static Encoding WithCharacterReferenceFallback(Encoding encoding) {
        if (encoding == null) throw new ArgumentNullException(nameof(encoding));
        Encoding copy = (Encoding)encoding.Clone();
        copy.EncoderFallback = new CharacterReferenceFallback();
        return copy;
    }

    /// <summary>
    /// Replaces unrepresentable scalar values with hexadecimal character references while preserving valid text.
    /// The supplied encoding must use <see cref="EncoderFallback.ExceptionFallback"/>.
    /// </summary>
    internal static string EscapeUnrepresentableCharacters(string value, Encoding encoding) {
        if (value == null) throw new ArgumentNullException(nameof(value));
        if (encoding == null) throw new ArgumentNullException(nameof(encoding));

        StringBuilder? escaped = null;
        var scalarBuffer = new char[2];
        for (int index = 0; index < value.Length;) {
            int characterCount = 1;
            int codePoint;
            char current = value[index];
            if (char.IsHighSurrogate(current)) {
                if (index + 1 >= value.Length || !char.IsLowSurrogate(value[index + 1])) {
                    throw new InvalidDataException("The text contains an invalid Unicode surrogate.");
                }
                characterCount = 2;
                codePoint = char.ConvertToUtf32(current, value[index + 1]);
            } else if (char.IsLowSurrogate(current)) {
                throw new InvalidDataException("The text contains an invalid Unicode surrogate.");
            } else {
                codePoint = current;
            }
            scalarBuffer[0] = current;
            if (characterCount == 2) scalarBuffer[1] = value[index + 1];

            bool representable;
            try {
                encoding.GetByteCount(scalarBuffer, 0, characterCount);
                representable = true;
            } catch (EncoderFallbackException) {
                representable = false;
            }

            if (!representable) {
                if (escaped == null) {
                    escaped = new StringBuilder(value.Length + 16);
                    escaped.Append(value, 0, index);
                }
                escaped.Append("&#x");
                escaped.Append(codePoint.ToString("X", CultureInfo.InvariantCulture));
                escaped.Append(';');
            } else if (escaped != null) {
                escaped.Append(value, index, characterCount);
            }
            index += characterCount;
        }
        return escaped?.ToString() ?? value;
    }

    private sealed class CharacterReferenceFallback : EncoderFallback {
        public override int MaxCharCount => 10;

        public override EncoderFallbackBuffer CreateFallbackBuffer() => new CharacterReferenceFallbackBuffer();
    }

    private sealed class CharacterReferenceFallbackBuffer : EncoderFallbackBuffer {
        private string _replacement = string.Empty;
        private int _position;

        public override bool Fallback(char charUnknown, int index) {
            if (Remaining != 0) return false;
            if (char.IsSurrogate(charUnknown)) throw new InvalidDataException("The text contains an invalid Unicode surrogate.");
            SetReplacement(charUnknown);
            return true;
        }

        public override bool Fallback(char charUnknownHigh, char charUnknownLow, int index) {
            if (Remaining != 0) return false;
            if (!char.IsSurrogatePair(charUnknownHigh, charUnknownLow)) {
                throw new InvalidDataException("The text contains an invalid Unicode surrogate.");
            }
            SetReplacement(char.ConvertToUtf32(charUnknownHigh, charUnknownLow));
            return true;
        }

        public override char GetNextChar() => _position < _replacement.Length ? _replacement[_position++] : '\0';

        public override bool MovePrevious() {
            if (_position == 0) return false;
            _position--;
            return true;
        }

        public override int Remaining => _replacement.Length - _position;

        public override void Reset() {
            _replacement = string.Empty;
            _position = 0;
        }

        private void SetReplacement(int codePoint) {
            _replacement = "&#x" + codePoint.ToString("X", CultureInfo.InvariantCulture) + ";";
            _position = 0;
        }
    }
}
