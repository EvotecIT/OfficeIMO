#nullable enable

using System.Xml;

namespace OfficeIMO.Excel {
    internal sealed partial class ExcelSheetReader {
        private static string? ReadXmlValueText(XmlReader valueReader) {
            if (valueReader.IsEmptyElement) {
                // Callers resume on this node, so <v/> must advance as well.
                valueReader.Read();
                return string.Empty;
            }

            int depth = valueReader.Depth;
            if (!valueReader.Read()) return null;
            if (!IsXmlTextNode(valueReader.NodeType)) {
                SkipXmlElementContent(valueReader, depth);
                return null;
            }

            // A simple XML value may span ordinary text, entities, and CDATA.
            // The XML reader keeps this on the value end element for the caller.
            return valueReader.ReadContentAsString();
        }

        private static string? ReadXmlValueTextAndSkipCell(XmlReader valueReader, int cellDepth) {
            string? text = ReadXmlValueText(valueReader);
            SkipXmlElementContent(valueReader, cellDepth);
            return text;
        }

        private bool TryReadXmlSimpleDoubleAndSkipCell(XmlReader valueReader, int cellDepth, out double value, out string? rawText) {
            value = 0;
            if (!TryReadXmlBufferedValueTextAndSkipCell(valueReader, cellDepth, out char[] buffer, out int length, out rawText)) {
                return false;
            }

            if (rawText == null) {
                if (TryParseInvariantDouble(buffer.AsSpan(0, length), out value)) {
                    return true;
                }

                rawText = new string(buffer, 0, length);
                return false;
            }

            return TryParseInvariantDouble(rawText, out value);
        }

        private bool TryReadXmlBufferedValueTextAndSkipCell(XmlReader valueReader, int cellDepth, out char[] buffer, out int length, out string? rawText) {
            buffer = _xmlValueTextBuffer ??= new char[64];
            length = 0;
            rawText = null;

            if (valueReader.IsEmptyElement) {
                SkipXmlElementContent(valueReader, cellDepth);
                rawText = string.Empty;
                return false;
            }

            int valueDepth = valueReader.Depth;
            if (!valueReader.Read()) {
                return false;
            }

            if (!IsXmlTextNode(valueReader.NodeType)) {
                SkipXmlElementContent(valueReader, cellDepth);
                return false;
            }

            System.Text.StringBuilder? builder = null;

            while (true) {
                int read = valueReader.ReadValueChunk(buffer, length, buffer.Length - length);
                if (read == 0) {
                    break;
                }

                length += read;
                if (length != buffer.Length) {
                    continue;
                }

                builder = new System.Text.StringBuilder(buffer.Length * 2);
                builder.Append(buffer, 0, length);
                length = 0;

                while (true) {
                    read = valueReader.ReadValueChunk(buffer, 0, buffer.Length);
                    if (read == 0) {
                        break;
                    }

                    builder.Append(buffer, 0, read);
                }

                break;
            }

            bool hasFollowingNode = valueReader.Read();
            if (hasFollowingNode && !(valueReader.NodeType == XmlNodeType.EndElement && valueReader.Depth == valueDepth)) {
                // Preserve split text/CDATA without allocating for the common single-text-node value.
                string tail = valueReader.ReadContentAsString();
                if (builder == null) {
                    builder = new System.Text.StringBuilder(length + tail.Length);
                    builder.Append(buffer, 0, length);
                }
                builder.Append(tail);
                length = 0;
            }

            bool completedCell = valueReader.NodeType == XmlNodeType.EndElement
                && valueReader.Depth == valueDepth
                && valueReader.Read()
                && valueReader.NodeType == XmlNodeType.EndElement
                && valueReader.Depth == cellDepth;
            if (!completedCell) {
                SkipXmlElementContent(valueReader, cellDepth);
            }

            if (builder == null) {
                return true;
            }

            rawText = builder.ToString();
            return true;
        }

        private static bool TryReadXmlSharedStringIndexValue(XmlReader valueReader, out int index, out string? rawText) {
            rawText = ReadXmlValueText(valueReader);
            return TryParseSharedStringIndex(rawText, out index);
        }

        private static bool TryReadXmlSharedStringIndexValueAndSkipCell(XmlReader valueReader, int cellDepth, out int index, out string? rawText) {
            rawText = ReadXmlValueTextAndSkipCell(valueReader, cellDepth);
            return TryParseSharedStringIndex(rawText, out index);
        }

        private static string ReadXmlInlineString(XmlReader inlineReader, XmlDataReaderTextBudget? textBudget = null) {
            if (textBudget != null) return textBudget.ReadInlineString(inlineReader);
            if (inlineReader.IsEmptyElement) {
                // Cell readers resume on the current node after this helper returns.
                // Advance past <is/> so they cannot repeatedly consume the same element.
                inlineReader.Read();
                return string.Empty;
            }

            int depth = inlineReader.Depth;
            int richRunDepth = -1;
            string? first = null;
            System.Text.StringBuilder? builder = null;
            while (inlineReader.Read()) {
                if (inlineReader.NodeType == XmlNodeType.EndElement && inlineReader.Depth == depth && inlineReader.LocalName == "is") {
                    break;
                }

                if (!IsXmlInlineStringTextElement(inlineReader, depth, ref richRunDepth)) {
                    continue;
                }

                string text = ReadXmlTextElement(inlineReader);
                if (builder != null) {
                    builder.Append(text);
                } else if (first == null) {
                    first = text;
                } else {
                    builder = new System.Text.StringBuilder(first.Length + text.Length);
                    builder.Append(first);
                    builder.Append(text);
                }
            }

            return builder?.ToString() ?? first ?? string.Empty;
        }

        /// <summary>Includes visible inline text in direct text elements and rich runs, excluding extension and phonetic payloads.</summary>
        private static bool IsXmlInlineStringTextElement(XmlReader reader, int inlineStringDepth, ref int richRunDepth) =>
            SpreadsheetXmlContent.IsRichTextElement(reader, inlineStringDepth, ref richRunDepth);

        private static string ReadXmlTextElement(XmlReader textReader) {
            if (textReader.IsEmptyElement) {
                return string.Empty;
            }

            int depth = textReader.Depth;
            string? first = null;
            System.Text.StringBuilder? builder = null;
            while (textReader.Read()) {
                if (textReader.NodeType == XmlNodeType.EndElement && textReader.Depth == depth && textReader.LocalName == "t") {
                    break;
                }

                if (textReader.NodeType == XmlNodeType.Element) {
                    throw new XmlException("Inline string text elements must contain text only.");
                }

                if (!IsXmlTextNode(textReader.NodeType)) {
                    continue;
                }

                string text = textReader.Value;
                if (builder != null) {
                    builder.Append(text);
                } else if (first == null) {
                    first = text;
                } else {
                    builder = new System.Text.StringBuilder(first.Length + text.Length);
                    builder.Append(first);
                    builder.Append(text);
                }
            }

            return builder?.ToString() ?? first ?? string.Empty;
        }

        private static bool IsXmlTextNode(XmlNodeType nodeType) {
            return nodeType == XmlNodeType.Text
                || nodeType == XmlNodeType.CDATA
                || nodeType == XmlNodeType.SignificantWhitespace
                || nodeType == XmlNodeType.Whitespace;
        }
    }
}
