#nullable enable

using System.Globalization;
using System.Text;
using System.Xml;

namespace OfficeIMO.Excel {
    internal sealed partial class ExcelSheetReader {
        // Keep this separate from coordinates: validation retains a cell reference
        // while reading its type and style, including when constructing errors.
        [ThreadStatic]
        private static char[]? _xmlAttributeTextBuffer;

        private readonly struct XmlStyleAttribute {
            internal readonly bool Present;
            internal readonly bool Valid;
            internal readonly uint Index;

            internal XmlStyleAttribute(ReadOnlySpan<char> text, bool present) {
                Present = present;
                Valid = TryParseUInt(text, out uint index);
                Index = index;
            }
        }

        private static XmlStyleAttribute ReadXmlStyleAttribute(XmlReader reader) {
            ReadOnlySpan<char> text = ReadXmlAttributeText(reader, "s", out bool present);
            return new XmlStyleAttribute(text, present);
        }

        private static bool TryReadXmlStyleIndex(XmlReader reader, out uint index) {
            XmlStyleAttribute style = ReadXmlStyleAttribute(reader);
            index = style.Index;
            return style.Valid;
        }

        private bool IsDateStyleAttribute(XmlStyleAttribute style) =>
            style.Valid && Styles.IsDateLike(style.Index);

        private bool IsCalendarStyleAttribute(XmlStyleAttribute style) =>
            style.Valid && Styles.IsDateSystemShiftStyle(style.Index);

        private static string? ReadXmlCellTypeAttribute(XmlReader reader) {
            ReadOnlySpan<char> text = ReadXmlAttributeText(reader, "t", out bool present);
            if (!present) return null;
            if (text.Length == 0) return string.Empty;
            if (text.Length == 1) {
                switch (text[0]) {
                    case 'b': return "b";
                    case 'd': return "d";
                    case 'e': return "e";
                    case 'n': return "n";
                    case 's': return "s";
                }
            } else if (text.SequenceEqual("str".AsSpan())) {
                return "str";
            } else if (text.SequenceEqual("inlineStr".AsSpan())) {
                return "inlineStr";
            }
            return text.ToString();
        }

        // The span is consumed before another metadata attribute is read on this
        // thread. Long and unusual values preserve the ordinary XML semantics.
        private static ReadOnlySpan<char> ReadXmlAttributeText(XmlReader reader, string name, out bool present) {
            if (!reader.CanReadValueChunk) {
                string? value = reader.GetAttribute(name);
                present = value != null;
                return value.AsSpan();
            }
            present = reader.MoveToAttribute(name);
            if (!present) return default;
            try {
                const int chunkLength = 32;
                char[] buffer = _xmlAttributeTextBuffer ??= new char[chunkLength * 2];
                int length = reader.ReadValueChunk(buffer, 0, chunkLength);
                int following = reader.ReadValueChunk(buffer, chunkLength, chunkLength);
                if (following == 0) return buffer.AsSpan(0, length);
                var builder = new StringBuilder(length + following);
                builder.Append(buffer, 0, length);
                builder.Append(buffer, chunkLength, following);
                while ((length = reader.ReadValueChunk(buffer, 0, buffer.Length)) != 0)
                    builder.Append(buffer, 0, length);
                return builder.ToString().AsSpan();
            } finally {
                reader.MoveToElement();
            }
        }

        private static bool TryParseUInt(ReadOnlySpan<char> text, out uint result) {
            result = 0;
            if (text.IsEmpty) {
                return false;
            }

            uint parsed = 0;
            for (int i = 0; i < text.Length; i++) {
                uint digit = (uint)(text[i] - '0');
                if (digit > 9U) {
                    return uint.TryParse(text.ToString(), NumberStyles.Integer, CultureInfo.InvariantCulture, out result);
                }

                if (parsed > (uint.MaxValue - digit) / 10U) {
                    return uint.TryParse(text.ToString(), NumberStyles.Integer, CultureInfo.InvariantCulture, out result);
                }

                parsed = (parsed * 10U) + digit;
            }

            result = parsed;
            return true;
        }
    }
}
