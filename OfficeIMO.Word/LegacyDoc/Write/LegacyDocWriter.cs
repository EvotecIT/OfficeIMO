using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Wordprocessing;
using OfficeIMO.Shared;
using System.Text;

namespace OfficeIMO.Word.LegacyDoc.Write {
    internal static class LegacyDocWriter {
        private const int FibLength = 0x1AA;
        private const int TextOffset = 0x200;
        private const int OleSectorSize = 512;
        private const int ClxLength = 21;
        private const int ChpxPlcLength = 12;
        private const ushort SprmCFBold = 0x0835;
        private const ushort SprmCFItalic = 0x0836;
        private const ushort WordDocumentMagic = 0xA5EC;
        private const ushort Word97FibVersion = 0x00D9;
        private const ushort OneTableStreamFlag = 0x0200;

        internal static byte[] WriteDocument(WordDocument document) {
            if (document == null) throw new ArgumentNullException(nameof(document));

            LegacyDocWritableBody body = BuildBody(document);
            byte[] wordDocumentStream = CreateWordDocumentStream(body);
            byte[] tableStream = CreateTableStream(body);
            IReadOnlyList<OfficeCompoundStream> propertyStreams = LegacyDocPropertySetWriter.CreateDocumentPropertyStreams(document);
            var streams = new List<OfficeCompoundStream>(propertyStreams.Count + 2) {
                new OfficeCompoundStream("WordDocument", wordDocumentStream),
                new OfficeCompoundStream("1Table", tableStream)
            };
            streams.AddRange(propertyStreams);

            return OfficeCompoundFileWriter.Write(streams);
        }

        private static LegacyDocWritableBody BuildBody(WordDocument document) {
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

            var text = new StringBuilder();
            var runs = new List<LegacyDocWritableRun>();
            int paragraphCount = 0;
            foreach (OpenXmlElement child in body.ChildElements) {
                switch (child) {
                    case Paragraph paragraph:
                        AppendParagraph(text, runs, paragraph);
                        paragraphCount++;
                        break;
                    case SectionProperties:
                        break;
                    default:
                        throw new NotSupportedException($"Native DOC saving currently supports body paragraphs with bold and italic text runs. Unsupported body element: {child.LocalName}.");
                }
            }

            if (paragraphCount == 0) {
                text.Append('\r');
            }

            return new LegacyDocWritableBody(text.ToString(), runs);
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

        private static void AppendParagraph(StringBuilder text, List<LegacyDocWritableRun> runs, Paragraph paragraph) {
            if (paragraph.ParagraphProperties != null && paragraph.ParagraphProperties.HasChildren) {
                throw new NotSupportedException("Native DOC saving currently supports unformatted paragraphs only. Paragraph properties are not supported yet.");
            }

            foreach (OpenXmlElement child in paragraph.ChildElements) {
                switch (child) {
                    case ParagraphProperties:
                        break;
                    case Run run:
                        AppendSupportedRunText(text, runs, run);
                        break;
                    default:
                        throw new NotSupportedException($"Native DOC saving currently supports only text runs with bold and italic formatting. Unsupported paragraph element: {child.LocalName}.");
                }
            }

            text.Append('\r');
        }

        private static void AppendSupportedRunText(StringBuilder text, List<LegacyDocWritableRun> runs, Run run) {
            LegacyDocWritableFormatting formatting = ReadSupportedRunFormatting(run.RunProperties);

            foreach (OpenXmlElement child in run.ChildElements) {
                switch (child) {
                    case RunProperties:
                        break;
                    case Text textNode:
                        AppendFormattedText(text, runs, textNode.Text, formatting);
                        break;
                    case TabChar:
                        AppendFormattedText(text, runs, "\t", formatting);
                        break;
                    default:
                        throw new NotSupportedException($"Native DOC saving currently supports text and tabs only. Unsupported run element: {child.LocalName}.");
                }
            }
        }

        private static LegacyDocWritableFormatting ReadSupportedRunFormatting(RunProperties? runProperties) {
            if (runProperties == null || !runProperties.HasChildren) {
                return LegacyDocWritableFormatting.Plain;
            }

            bool bold = false;
            bool italic = false;
            foreach (OpenXmlElement property in runProperties.ChildElements) {
                switch (property) {
                    case Bold boldProperty:
                        bold = IsEnabled(boldProperty);
                        break;
                    case BoldComplexScript boldComplexScript:
                        bold = IsEnabled(boldComplexScript);
                        break;
                    case Italic italicProperty:
                        italic = IsEnabled(italicProperty);
                        break;
                    case ItalicComplexScript italicComplexScript:
                        italic = IsEnabled(italicComplexScript);
                        break;
                    default:
                        throw new NotSupportedException($"Native DOC saving currently supports only bold and italic run formatting. Unsupported run property: {property.LocalName}.");
                }
            }

            return new LegacyDocWritableFormatting(bold, italic);
        }

        private static bool IsEnabled(OnOffType property) {
            return property.Val == null || property.Val.Value;
        }

        private static void AppendFormattedText(
            StringBuilder text,
            List<LegacyDocWritableRun> runs,
            string? value,
            LegacyDocWritableFormatting formatting) {
            if (string.IsNullOrEmpty(value)) {
                return;
            }

            string textValue = value!;
            int start = text.Length;
            text.Append(textValue);
            if (!formatting.HasFormatting) {
                return;
            }

            int length = textValue.Length;
            if (runs.Count > 0) {
                LegacyDocWritableRun previous = runs[runs.Count - 1];
                if (previous.EndCharacter == start && previous.Formatting.Equals(formatting)) {
                    runs[runs.Count - 1] = previous.Extend(length);
                    return;
                }
            }

            runs.Add(new LegacyDocWritableRun(start, length, formatting));
        }

        private static byte[] CreateWordDocumentStream(LegacyDocWritableBody body) {
            byte[] textBytes = Encoding.Unicode.GetBytes(body.Text);
            int chpxFkpOffset = body.HasCharacterFormatting
                ? AlignToSector(TextOffset + textBytes.Length)
                : 0;
            int streamLength = body.HasCharacterFormatting
                ? chpxFkpOffset + OleSectorSize
                : TextOffset + textBytes.Length;
            var stream = new byte[Math.Max(FibLength, streamLength)];
            WriteUInt16(stream, 0x00, WordDocumentMagic);
            WriteUInt16(stream, 0x02, Word97FibVersion);
            WriteUInt16(stream, 0x0A, OneTableStreamFlag);
            WriteInt32(stream, 0x4C, body.Text.Length);
            WriteInt32(stream, 0xFA, body.HasCharacterFormatting ? ClxLength : 0);
            WriteInt32(stream, 0xFE, body.HasCharacterFormatting ? ChpxPlcLength : 0);
            WriteInt32(stream, 0x1A2, 0);
            WriteInt32(stream, 0x1A6, ClxLength);
            Buffer.BlockCopy(textBytes, 0, stream, TextOffset, textBytes.Length);
            if (body.HasCharacterFormatting) {
                WriteChpxFkp(stream, chpxFkpOffset, body.CreateFormattingSegments());
            }

            return stream;
        }

        private static byte[] CreateTableStream(LegacyDocWritableBody body) {
            var table = new byte[body.HasCharacterFormatting ? ClxLength + ChpxPlcLength : ClxLength];
            table[0] = 0x02;
            WriteInt32(table, 1, 16);
            WriteInt32(table, 5, 0);
            WriteInt32(table, 9, body.Text.Length);
            WriteUInt16(table, 13, 0);
            WriteUInt32(table, 15, TextOffset);
            WriteUInt16(table, 19, 0);

            if (body.HasCharacterFormatting) {
                int chpxFkpOffset = AlignToSector(TextOffset + Encoding.Unicode.GetByteCount(body.Text));
                WriteInt32(table, ClxLength, TextOffset);
                WriteInt32(table, ClxLength + 4, TextOffset + (body.Text.Length * 2));
                WriteInt32(table, ClxLength + 8, chpxFkpOffset / OleSectorSize);
            }

            return table;
        }

        private static void WriteChpxFkp(byte[] stream, int pageOffset, IReadOnlyList<LegacyDocWritableSegment> segments) {
            if (segments.Count == 0 || segments.Count > byte.MaxValue) {
                throw new NotSupportedException("Native DOC saving currently supports bold and italic formatting only when it fits in one character-format page.");
            }

            int rgbOffset = pageOffset + ((segments.Count + 1) * 4);
            int chpxOffset = AlignToEven((segments.Count + 1) * 4 + segments.Count);

            for (int index = 0; index < segments.Count; index++) {
                LegacyDocWritableSegment segment = segments[index];
                WriteInt32(stream, pageOffset + (index * 4), TextOffset + (segment.StartCharacter * 2));
                if (segment.Formatting.HasFormatting) {
                    byte[] chpx = CreateChpx(segment.Formatting);
                    chpxOffset = AlignToEven(chpxOffset);
                    if (chpxOffset + chpx.Length >= OleSectorSize - 1 || chpxOffset / 2 > byte.MaxValue) {
                        throw new NotSupportedException("Native DOC saving currently supports bold and italic formatting only when it fits in one character-format page.");
                    }

                    Buffer.BlockCopy(chpx, 0, stream, pageOffset + chpxOffset, chpx.Length);
                    stream[rgbOffset + index] = (byte)(chpxOffset / 2);
                    chpxOffset += chpx.Length;
                }
            }

            LegacyDocWritableSegment lastSegment = segments[segments.Count - 1];
            WriteInt32(stream, pageOffset + (segments.Count * 4), TextOffset + (lastSegment.EndCharacter * 2));
            stream[pageOffset + OleSectorSize - 1] = (byte)segments.Count;
        }

        private static byte[] CreateChpx(LegacyDocWritableFormatting formatting) {
            var grpprl = new List<byte>(6);
            if (formatting.Bold) {
                grpprl.Add((byte)(SprmCFBold & 0xFF));
                grpprl.Add((byte)(SprmCFBold >> 8));
                grpprl.Add(1);
            }

            if (formatting.Italic) {
                grpprl.Add((byte)(SprmCFItalic & 0xFF));
                grpprl.Add((byte)(SprmCFItalic >> 8));
                grpprl.Add(1);
            }

            var chpx = new byte[grpprl.Count + 1];
            chpx[0] = (byte)grpprl.Count;
            grpprl.CopyTo(chpx, 1);
            return chpx;
        }

        private static int AlignToSector(int value) {
            return ((value + OleSectorSize - 1) / OleSectorSize) * OleSectorSize;
        }

        private static int AlignToEven(int value) {
            return value % 2 == 0 ? value : value + 1;
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

        private readonly struct LegacyDocWritableBody {
            internal LegacyDocWritableBody(string text, IReadOnlyList<LegacyDocWritableRun> formattedRuns) {
                Text = text;
                FormattedRuns = formattedRuns;
            }

            internal string Text { get; }

            internal IReadOnlyList<LegacyDocWritableRun> FormattedRuns { get; }

            internal bool HasCharacterFormatting => FormattedRuns.Count > 0;

            internal IReadOnlyList<LegacyDocWritableSegment> CreateFormattingSegments() {
                var segments = new List<LegacyDocWritableSegment>();
                int character = 0;
                foreach (LegacyDocWritableRun run in FormattedRuns.OrderBy(item => item.StartCharacter)) {
                    if (run.StartCharacter > character) {
                        AddSegment(segments, character, run.StartCharacter - character, LegacyDocWritableFormatting.Plain);
                    }

                    AddSegment(segments, run.StartCharacter, run.Length, run.Formatting);
                    character = run.EndCharacter;
                }

                if (character < Text.Length) {
                    AddSegment(segments, character, Text.Length - character, LegacyDocWritableFormatting.Plain);
                }

                return segments;
            }

            private static void AddSegment(
                List<LegacyDocWritableSegment> segments,
                int startCharacter,
                int length,
                LegacyDocWritableFormatting formatting) {
                if (length <= 0) {
                    return;
                }

                if (segments.Count > 0) {
                    LegacyDocWritableSegment previous = segments[segments.Count - 1];
                    if (previous.EndCharacter == startCharacter && previous.Formatting.Equals(formatting)) {
                        segments[segments.Count - 1] = previous.Extend(length);
                        return;
                    }
                }

                segments.Add(new LegacyDocWritableSegment(startCharacter, length, formatting));
            }
        }

        private readonly struct LegacyDocWritableFormatting : IEquatable<LegacyDocWritableFormatting> {
            internal static readonly LegacyDocWritableFormatting Plain = new LegacyDocWritableFormatting(false, false);

            internal LegacyDocWritableFormatting(bool bold, bool italic) {
                Bold = bold;
                Italic = italic;
            }

            internal bool Bold { get; }

            internal bool Italic { get; }

            internal bool HasFormatting => Bold || Italic;

            public bool Equals(LegacyDocWritableFormatting other) {
                return Bold == other.Bold && Italic == other.Italic;
            }

            public override bool Equals(object? obj) {
                return obj is LegacyDocWritableFormatting other && Equals(other);
            }

            public override int GetHashCode() {
                int hash = 17;
                hash = (hash * 31) + Bold.GetHashCode();
                hash = (hash * 31) + Italic.GetHashCode();
                return hash;
            }
        }

        private readonly struct LegacyDocWritableRun {
            internal LegacyDocWritableRun(int startCharacter, int length, LegacyDocWritableFormatting formatting) {
                StartCharacter = startCharacter;
                Length = length;
                Formatting = formatting;
            }

            internal int StartCharacter { get; }

            internal int Length { get; }

            internal int EndCharacter => StartCharacter + Length;

            internal LegacyDocWritableFormatting Formatting { get; }

            internal LegacyDocWritableRun Extend(int additionalLength) {
                return new LegacyDocWritableRun(StartCharacter, Length + additionalLength, Formatting);
            }
        }

        private readonly struct LegacyDocWritableSegment {
            internal LegacyDocWritableSegment(int startCharacter, int length, LegacyDocWritableFormatting formatting) {
                StartCharacter = startCharacter;
                Length = length;
                Formatting = formatting;
            }

            internal int StartCharacter { get; }

            internal int Length { get; }

            internal int EndCharacter => StartCharacter + Length;

            internal LegacyDocWritableFormatting Formatting { get; }

            internal LegacyDocWritableSegment Extend(int additionalLength) {
                return new LegacyDocWritableSegment(StartCharacter, Length + additionalLength, Formatting);
            }
        }
    }
}
