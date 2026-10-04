#nullable enable

using System.Text;
using System.Xml;

namespace OfficeIMO.Excel {
    internal sealed partial class ExcelSheetReader {
        // Parsing and validation consume each reference synchronously, before the next
        // reference read on this thread. Separate buffers keep parallel readers independent.
        [ThreadStatic]
        private static char[]? _xmlCoordinateTextBuffer;

        private readonly ref struct XmlCoordinateReference {
            internal readonly ReadOnlySpan<char> Text;
            private readonly bool _hasValue;

            internal XmlCoordinateReference(string? value) {
                Text = value.AsSpan();
                _hasValue = value != null;
            }

            internal XmlCoordinateReference(ReadOnlySpan<char> value) {
                Text = value;
                _hasValue = true;
            }

            public override string ToString() => _hasValue ? Text.ToString() : "(unknown cell)";
        }

        /// <summary>
        /// Reads the unqualified coordinate attribute without materializing ordinary
        /// references. The returned span is valid until the next call on this thread.
        /// </summary>
        private static XmlCoordinateReference ReadXmlReferenceAttribute(XmlReader reader) {
            if (!reader.CanReadValueChunk) return new XmlCoordinateReference(reader.GetAttribute("r"));
            if (!reader.MoveToAttribute("r")) return default;

            try {
                const int chunkLength = 32;
                char[] buffer = _xmlCoordinateTextBuffer ??= new char[chunkLength * 2];
                int length = reader.ReadValueChunk(buffer, 0, chunkLength);
                // Framework attribute chunks can be one character short when a
                // surrogate pair crosses the boundary; shorter values are complete.
                if (length < chunkLength - 1) return new XmlCoordinateReference(buffer.AsSpan(0, length));
                // A short chunk can end before a surrogate pair, rather than at EOF.
                // Keep the first chunk intact while checking for additional content.
                int followingLength = reader.ReadValueChunk(buffer, chunkLength, chunkLength);
                if (followingLength == 0) {
                    return new XmlCoordinateReference(buffer.AsSpan(0, length));
                }

                var builder = new StringBuilder(length + followingLength);
                builder.Append(buffer, 0, length);
                builder.Append(buffer, chunkLength, followingLength);
                while ((length = reader.ReadValueChunk(buffer, 0, buffer.Length)) != 0) {
                    builder.Append(buffer, 0, length);
                }
                return new XmlCoordinateReference(builder.ToString());
            } finally {
                reader.MoveToElement();
            }
        }
    }
}
