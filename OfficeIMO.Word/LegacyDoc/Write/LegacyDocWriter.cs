using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Wordprocessing;
using OfficeIMO.Shared;
using System.Text;

namespace OfficeIMO.Word.LegacyDoc.Write {
    internal static class LegacyDocWriter {
        private const int FibLength = 0x1AA;
        private const int TextOffset = 0x200;
        private const ushort WordDocumentMagic = 0xA5EC;
        private const ushort Word97FibVersion = 0x00D9;
        private const ushort OneTableStreamFlag = 0x0200;

        internal static byte[] WriteDocument(WordDocument document) {
            if (document == null) throw new ArgumentNullException(nameof(document));

            string bodyText = BuildBodyText(document);
            byte[] wordDocumentStream = CreateWordDocumentStream(bodyText);
            byte[] tableStream = CreateTableStream(bodyText.Length);
            IReadOnlyList<OfficeCompoundStream> propertyStreams = LegacyDocPropertySetWriter.CreateDocumentPropertyStreams(document);
            var streams = new List<OfficeCompoundStream>(propertyStreams.Count + 2) {
                new OfficeCompoundStream("WordDocument", wordDocumentStream),
                new OfficeCompoundStream("1Table", tableStream)
            };
            streams.AddRange(propertyStreams);

            return OfficeCompoundFileWriter.Write(streams);
        }

        private static string BuildBodyText(WordDocument document) {
            var wordDocument = document._wordprocessingDocument;
            if (wordDocument == null) {
                throw new InvalidOperationException("Document couldn't be saved as legacy DOC because WordDocument wasn't provided.");
            }

            wordDocument.Save();
            var mainPart = wordDocument.MainDocumentPart;
            var body = mainPart?.Document?.Body;
            if (body == null) {
                throw new InvalidOperationException("Document couldn't be saved as legacy DOC because the document body is missing.");
            }

            ThrowIfUnsupportedDocumentParts(document, mainPart);

            var paragraphs = new List<string>();
            foreach (OpenXmlElement child in body.ChildElements) {
                switch (child) {
                    case Paragraph paragraph:
                        paragraphs.Add(ExtractPlainParagraphText(paragraph));
                        break;
                    case SectionProperties:
                        break;
                    default:
                        throw new NotSupportedException($"Native DOC saving currently supports only plain body paragraphs. Unsupported body element: {child.LocalName}.");
                }
            }

            return paragraphs.Count == 0
                ? "\r"
                : string.Join("\r", paragraphs) + "\r";
        }

        private static void ThrowIfUnsupportedDocumentParts(WordDocument document, DocumentFormat.OpenXml.Packaging.MainDocumentPart? mainPart) {
            if (document.TablesIncludingNestedTables.Count > 0) {
                throw new NotSupportedException("Native DOC saving currently supports only plain paragraphs. Tables are not supported yet.");
            }

            if (document.Sections.Count > 1) {
                throw new NotSupportedException("Native DOC saving currently supports a single section only.");
            }

            if (mainPart == null) {
                return;
            }

            if (mainPart.HeaderParts.Any() || mainPart.FooterParts.Any()) {
                throw new NotSupportedException("Native DOC saving currently supports body paragraphs only. Headers and footers are not supported yet.");
            }

            if (HasUserFootnotes(mainPart) || HasUserEndnotes(mainPart)) {
                throw new NotSupportedException("Native DOC saving currently supports body paragraphs only. Footnotes and endnotes are not supported yet.");
            }

            if (mainPart.ImageParts.Any()) {
                throw new NotSupportedException("Native DOC saving currently supports text only. Images are not supported yet.");
            }

            if (mainPart.ChartParts.Any()) {
                throw new NotSupportedException("Native DOC saving currently supports text only. Charts are not supported yet.");
            }
        }

        private static bool HasUserFootnotes(DocumentFormat.OpenXml.Packaging.MainDocumentPart mainPart) {
            var footnotes = mainPart.FootnotesPart?.Footnotes;
            return footnotes != null && footnotes.Elements<Footnote>().Any(footnote => footnote.Type == null);
        }

        private static bool HasUserEndnotes(DocumentFormat.OpenXml.Packaging.MainDocumentPart mainPart) {
            var endnotes = mainPart.EndnotesPart?.Endnotes;
            return endnotes != null && endnotes.Elements<Endnote>().Any(endnote => endnote.Type == null);
        }

        private static string ExtractPlainParagraphText(Paragraph paragraph) {
            if (paragraph.ParagraphProperties != null && paragraph.ParagraphProperties.HasChildren) {
                throw new NotSupportedException("Native DOC saving currently supports unformatted paragraphs only. Paragraph properties are not supported yet.");
            }

            var text = new StringBuilder();
            foreach (OpenXmlElement child in paragraph.ChildElements) {
                switch (child) {
                    case ParagraphProperties:
                        break;
                    case Run run:
                        AppendPlainRunText(text, run);
                        break;
                    default:
                        throw new NotSupportedException($"Native DOC saving currently supports only plain text runs. Unsupported paragraph element: {child.LocalName}.");
                }
            }

            return text.ToString();
        }

        private static void AppendPlainRunText(StringBuilder text, Run run) {
            if (run.RunProperties != null && run.RunProperties.HasChildren) {
                throw new NotSupportedException("Native DOC saving currently supports unformatted runs only. Run properties are not supported yet.");
            }

            foreach (OpenXmlElement child in run.ChildElements) {
                switch (child) {
                    case RunProperties:
                        break;
                    case Text textNode:
                        text.Append(textNode.Text);
                        break;
                    case TabChar:
                        text.Append('\t');
                        break;
                    default:
                        throw new NotSupportedException($"Native DOC saving currently supports plain text and tabs only. Unsupported run element: {child.LocalName}.");
                }
            }
        }

        private static byte[] CreateWordDocumentStream(string text) {
            byte[] textBytes = Encoding.Unicode.GetBytes(text);
            var stream = new byte[Math.Max(FibLength, TextOffset + textBytes.Length)];
            WriteUInt16(stream, 0x00, WordDocumentMagic);
            WriteUInt16(stream, 0x02, Word97FibVersion);
            WriteUInt16(stream, 0x0A, OneTableStreamFlag);
            WriteInt32(stream, 0x4C, text.Length);
            WriteInt32(stream, 0x1A2, 0);
            WriteInt32(stream, 0x1A6, 21);
            Buffer.BlockCopy(textBytes, 0, stream, TextOffset, textBytes.Length);
            return stream;
        }

        private static byte[] CreateTableStream(int characterCount) {
            var table = new byte[21];
            table[0] = 0x02;
            WriteInt32(table, 1, 16);
            WriteInt32(table, 5, 0);
            WriteInt32(table, 9, characterCount);
            WriteUInt16(table, 13, 0);
            WriteUInt32(table, 15, TextOffset);
            WriteUInt16(table, 19, 0);
            return table;
        }

        private static void WriteUInt16(byte[] bytes, int offset, ushort value) {
            bytes[offset] = (byte)value;
            bytes[offset + 1] = (byte)(value >> 8);
        }

        private static void WriteInt32(byte[] bytes, int offset, int value) {
            bytes[offset] = (byte)value;
            bytes[offset + 1] = (byte)(value >> 8);
            bytes[offset + 2] = (byte)(value >> 16);
            bytes[offset + 3] = (byte)(value >> 24);
        }

        private static void WriteUInt32(byte[] bytes, int offset, uint value) {
            bytes[offset] = (byte)value;
            bytes[offset + 1] = (byte)(value >> 8);
            bytes[offset + 2] = (byte)(value >> 16);
            bytes[offset + 3] = (byte)(value >> 24);
        }
    }
}
