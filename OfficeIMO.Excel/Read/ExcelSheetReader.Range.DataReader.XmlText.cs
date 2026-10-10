using System.Xml;

namespace OfficeIMO.Excel {
    internal sealed partial class ExcelSheetReader {
        /// <summary>
        /// Charges decoded text before constructing XML-streaming row values. A
        /// single allowance also covers every row retained for out-of-order input.
        /// </summary>
        private sealed class XmlDataReaderTextBudget {
            private readonly long _maximumCharacters;
            private readonly Action _checkCancellation;
            private readonly char[] _characters = new char[4096];
            private long _remainingCharacters;

            internal XmlDataReaderTextBudget(long maximumCharacters, Action checkCancellation) {
                _maximumCharacters = maximumCharacters;
                _remainingCharacters = maximumCharacters;
                _checkCancellation = checkCancellation;
            }

            internal bool WasExceeded { get; private set; }

            internal void Reset() => _remainingCharacters = _maximumCharacters;

            internal void Charge(string? text) {
                if (text != null) Charge(text.Length);
            }

            private void Charge(int count) {
                _checkCancellation();
                if (count > _remainingCharacters) {
                    WasExceeded = true;
                    throw ExcelReadLimitFailure.Create(
                        $"XML data-reader buffering exceeds {nameof(ExcelReadOptions.MaxXmlDataReaderBufferedCharacters)} "
                        + $"({_maximumCharacters} characters).");
                }

                _remainingCharacters -= count;
            }

            internal string? ReadAttribute(XmlReader reader, string name) {
                if (!reader.MoveToAttribute(name)) return null;
                try {
                    string? first = null;
                    StringBuilder? builder = null;
                    AppendCurrentText(reader, ref first, ref builder);
                    return builder?.ToString() ?? first ?? string.Empty;
                } finally {
                    reader.MoveToElement();
                }
            }

            internal string ReadElementText(XmlReader reader, bool advancePastEnd) {
                _checkCancellation();
                if (reader.IsEmptyElement) {
                    if (advancePastEnd) reader.Read();
                    return string.Empty;
                }

                int depth = reader.Depth;
                string? first = null;
                StringBuilder? builder = null;
                while (reader.Read()) {
                    _checkCancellation();
                    if (reader.NodeType == XmlNodeType.EndElement && reader.Depth == depth) {
                        if (advancePastEnd) reader.Read();
                        return builder?.ToString() ?? first ?? string.Empty;
                    }

                    if (IsXmlTextNode(reader.NodeType)) {
                        AppendCurrentText(reader, ref first, ref builder);
                    } else if (reader.NodeType == XmlNodeType.Element) {
                        throw new XmlException("Cell text elements must contain text only.");
                    }
                }

                throw new XmlException("Unexpected end of XML while reading cell text.");
            }

            internal string ReadInlineString(XmlReader reader) {
                _checkCancellation();
                if (reader.IsEmptyElement) {
                    reader.Read();
                    return string.Empty;
                }

                int depth = reader.Depth;
                int richRunDepth = -1;
                string? first = null;
                StringBuilder? builder = null;
                while (reader.Read()) {
                    _checkCancellation();
                    if (reader.NodeType == XmlNodeType.EndElement && reader.Depth == depth && reader.LocalName == "is") {
                        return builder?.ToString() ?? first ?? string.Empty;
                    }
                    if (!IsXmlInlineStringTextElement(reader, depth, ref richRunDepth)) continue;

                    string text = ReadElementText(reader, advancePastEnd: false);
                    if (builder != null) {
                        builder.Append(text);
                    } else if (first == null) {
                        first = text;
                    } else {
                        builder = new StringBuilder(first.Length + text.Length);
                        builder.Append(first);
                        builder.Append(text);
                    }
                }

                throw new XmlException("Unexpected end of XML while reading an inline string.");
            }

            private void AppendCurrentText(XmlReader reader, ref string? first, ref StringBuilder? builder) {
                int read;
                while ((read = reader.ReadValueChunk(_characters, 0, _characters.Length)) != 0) {
                    // XmlReader can already buffer a CDATA/attribute node, but
                    // OfficeIMO never copies its complete value before charging.
                    // Never request Value/ReadContentAsString before this check: those
                    // APIs can first allocate the complete decompressed text node.
                    Charge(read);
                    if (builder != null) {
                        builder.Append(_characters, 0, read);
                    } else if (first == null) {
                        first = new string(_characters, 0, read);
                    } else {
                        builder = new StringBuilder(first.Length + read);
                        builder.Append(first);
                        builder.Append(_characters, 0, read);
                    }
                }
            }
        }
    }
}
