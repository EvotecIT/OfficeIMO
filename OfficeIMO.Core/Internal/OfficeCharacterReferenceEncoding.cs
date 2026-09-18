using System;
using System.Globalization;
using System.IO;
using System.Text;

namespace OfficeIMO.Core.Internal;

/// <summary>Escapes Unicode scalar values that a strict legacy text encoding cannot represent.</summary>
internal static class OfficeCharacterReferenceEncoding {
    /// <summary>
    /// Replaces unrepresentable scalar values with hexadecimal character references while preserving valid text.
    /// The supplied encoding must use <see cref="EncoderFallback.ExceptionFallback"/>.
    /// </summary>
    internal static string EscapeUnrepresentableCharacters(string value, Encoding encoding) {
        if (value == null) throw new ArgumentNullException(nameof(value));
        if (encoding == null) throw new ArgumentNullException(nameof(encoding));

        StringBuilder? escaped = null;
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

            bool representable;
            try {
                encoding.GetByteCount(value.Substring(index, characterCount));
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
}
