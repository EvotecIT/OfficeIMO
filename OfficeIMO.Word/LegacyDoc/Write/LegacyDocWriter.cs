using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Wordprocessing;
using OfficeIMO.Shared;
using OfficeIMO.Word.LegacyDoc.Model;
using System.Text;

namespace OfficeIMO.Word.LegacyDoc.Write {
    internal static partial class LegacyDocWriter {
        private const int FibLength = 0x1AA;
        private const int TextOffset = 0x800;
        private const int OleSectorSize = 512;
        private const int OleMiniStreamCutoffSize = 4096;
        private const int ClxLength = 21;
        private const int ChpxPlcLength = 12;
        private const int PapxPlcLength = 12;
        private const int SedLength = 12;
        private const uint CompressedTextFlag = 0x40000000;
        private const ushort SprmCFBold = 0x0835;
        private const ushort SprmCFItalic = 0x0836;
        private const ushort SprmCFStrike = 0x0837;
        private const ushort SprmCFOutline = 0x0838;
        private const ushort SprmCFShadow = 0x0839;
        private const ushort SprmCFSmallCaps = 0x083A;
        private const ushort SprmCFCaps = 0x083B;
        private const ushort SprmCFEmboss = 0x0858;
        private const ushort SprmCHighlight = 0x2A0C;
        private const ushort SprmCKul = 0x2A3E;
        private const ushort SprmCIss = 0x2A48;
        private const ushort SprmCHps = 0x4A43;
        private const ushort SprmCRgFtc0 = 0x4A4F;
        private const ushort SprmCFDStrike = 0x2A53;
        private const ushort SprmCCv = 0x6870;
        private const ushort SprmPJc = 0x2461;
        private const ushort DefaultPcdFlags = 0x0310;
        private const ushort WordDocumentMagic = 0xA5EC;
        private const ushort Word97FibVersion = 0x00C1;
        private const ushort Word97FibBackVersion = 0x00BF;
        private const ushort DefaultLanguageId = 0x0409;
        private const ushort FibRgW97WordCount = 0x000E;
        private const ushort FibRgLw97DwordCount = 0x0016;
        private const ushort FibRgFcLcb97Size = 0x00B7;
        private const ushort OneTableStreamFlag = 0x0200;
        private const ushort ExtendedCharacterFlag = 0x1000;
        private const int FcPlcfSedOffset = 0xCA;
        private const int LcbPlcfSedOffset = 0xCE;
        private const int FcPlcfBtePapxOffset = 0x102;
        private const int LcbPlcfBtePapxOffset = 0x106;
        private const int FcSttbfFfnOffset = 0x112;
        private const int LcbSttbfFfnOffset = 0x116;

        internal static byte[] WriteDocument(WordDocument document) {
            if (document == null) throw new ArgumentNullException(nameof(document));

            ThrowIfUnsupportedLegacyDocImportState(document);

            LegacyDocWritableBody body = BuildBody(document);
            byte[] wordDocumentStream = PadToRegularOleStream(CreateWordDocumentStream(body));
            byte[] tableStream = PadToRegularOleStream(CreateTableStream(body));
            IReadOnlyList<OfficeCompoundStream> propertyStreams = LegacyDocPropertySetWriter.CreateDocumentPropertyStreams(document);
            var streams = new List<OfficeCompoundStream>(propertyStreams.Count + 2) {
                new OfficeCompoundStream("WordDocument", wordDocumentStream),
                new OfficeCompoundStream("1Table", tableStream)
            };
            foreach (OfficeCompoundStream propertyStream in propertyStreams) {
                streams.Add(new OfficeCompoundStream(propertyStream.Name, PadToRegularOleStream(propertyStream.Bytes)));
            }

            return OfficeCompoundFileWriter.Write(streams);
        }

        private static void ThrowIfUnsupportedLegacyDocImportState(WordDocument document) {
            if (!document.WasLoadedFromLegacyDoc || document.LegacyDocUnsupportedFeatures.Count == 0) {
                return;
            }

            string codes = string.Join(
                ", ",
                document.LegacyDocUnsupportedFeatures
                    .Select(feature => feature.Code)
                    .Where(code => !string.IsNullOrWhiteSpace(code))
                    .Distinct(StringComparer.Ordinal)
                    .Take(5));
            string detail = string.IsNullOrWhiteSpace(codes)
                ? "unsupported or preserve-only features"
                : $"unsupported or preserve-only features ({codes})";

            throw new NotSupportedException($"Native DOC saving is blocked because this document was imported from a legacy DOC with {detail}. Save as DOCX after reviewing LegacyDocUnsupportedFeatures, or remove and recreate the unsupported content before saving as DOC.");
        }

        private static byte[] PadToRegularOleStream(byte[] bytes) {
            if (bytes.Length >= OleMiniStreamCutoffSize) {
                return bytes;
            }

            byte[] padded = new byte[OleMiniStreamCutoffSize];
            Buffer.BlockCopy(bytes, 0, padded, 0, bytes.Length);
            return padded;
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
            var paragraphFormats = new List<LegacyDocWritableParagraph>();
            LegacyDocSectionFormat finalSectionFormat = LegacyDocSectionFormat.Default;
            var sections = new List<LegacyDocWritableSection>();
            SectionMarkValues? pendingSectionBreakType = null;
            int bodyContentCount = 0;
            foreach (OpenXmlElement child in body.ChildElements) {
                switch (child) {
                    case Paragraph paragraph:
                        if (!IsPureSectionBreakParagraph(paragraph)) {
                            AppendParagraph(text, runs, paragraphFormats, paragraph);
                            bodyContentCount++;
                        }

                        SectionProperties? paragraphSectionProperties = paragraph.ParagraphProperties?.SectionProperties;
                        if (paragraphSectionProperties != null) {
                            LegacyDocSectionFormat paragraphSectionFormat = ReadSupportedSectionProperties(paragraphSectionProperties);
                            AddSection(sections, text.Length, paragraphSectionFormat.WithSectionBreakType(pendingSectionBreakType));
                            pendingSectionBreakType = paragraphSectionFormat.SectionBreakType;
                        }

                        break;
                    case Table table:
                        AppendTable(text, runs, paragraphFormats, table);
                        bodyContentCount++;
                        break;
                    case SectionProperties sectionProperties:
                        finalSectionFormat = ReadSupportedSectionProperties(sectionProperties);
                        break;
                    default:
                        throw new NotSupportedException($"Native DOC saving currently supports body paragraphs and simple tables with bold, italic, strikethrough, double-strikethrough, outline, shadow, emboss, caps/small-caps, superscript/subscript, underline, highlight, font size, color, and font family text runs. Unsupported body element: {child.LocalName}.");
                }
            }

            if (bodyContentCount == 0) {
                text.Append('\r');
            }

            AddSection(sections, text.Length, finalSectionFormat.WithSectionBreakType(pendingSectionBreakType));
            return new LegacyDocWritableBody(text.ToString(), runs, paragraphFormats, sections);
        }

        private static void ThrowIfUnsupportedDocumentParts(WordDocument document, DocumentFormat.OpenXml.Packaging.MainDocumentPart? mainPart) {
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

        private static bool IsPureSectionBreakParagraph(Paragraph paragraph) {
            ParagraphProperties? paragraphProperties = paragraph.ParagraphProperties;
            if (paragraphProperties?.SectionProperties == null) {
                return false;
            }

            return paragraph.ChildElements.All(element => element is ParagraphProperties)
                && paragraphProperties.ChildElements.All(element => element is SectionProperties);
        }

        private static void AddSection(List<LegacyDocWritableSection> sections, int endCharacter, LegacyDocSectionFormat format) {
            if (sections.Count > 0 && endCharacter < sections[sections.Count - 1].EndCharacter) {
                throw new NotSupportedException("Native DOC saving cannot write sections with non-monotonic text ranges.");
            }

            if (sections.Count > 0 && endCharacter == sections[sections.Count - 1].EndCharacter) {
                sections[sections.Count - 1] = new LegacyDocWritableSection(endCharacter, format);
                return;
            }

            sections.Add(new LegacyDocWritableSection(endCharacter, format));
        }

        private static void AppendParagraph(StringBuilder text, List<LegacyDocWritableRun> runs, List<LegacyDocWritableParagraph> paragraphFormats, Paragraph paragraph) {
            LegacyDocWritableParagraphFormatting paragraphFormatting = ReadSupportedParagraphFormatting(paragraph.ParagraphProperties);
            int paragraphStart = text.Length;

            foreach (OpenXmlElement child in paragraph.ChildElements) {
                switch (child) {
                    case ParagraphProperties:
                        break;
                    case Run run:
                        AppendSupportedRunText(text, runs, run);
                        break;
                    default:
                        throw new NotSupportedException($"Native DOC saving currently supports only text runs with bold, italic, strikethrough, double-strikethrough, outline, shadow, emboss, caps/small-caps, superscript/subscript, underline, highlight, font size, color, and font family formatting. Unsupported paragraph element: {child.LocalName}.");
                }
            }

            text.Append('\r');
            if (paragraphFormatting.HasFormatting) {
                paragraphFormats.Add(new LegacyDocWritableParagraph(paragraphStart, text.Length - paragraphStart, paragraphFormatting));
            }
        }

        private static byte[] CreateWordDocumentStream(LegacyDocWritableBody body) {
            bool compressedText = CanWriteCompressedText(body.Text);
            int bytesPerCharacter = compressedText ? 1 : 2;
            byte[] textBytes = compressedText ? EncodeCompressedText(body.Text) : Encoding.Unicode.GetBytes(body.Text);
            byte[] fontTable = CreateFontTable(body.FontFamilies);
            int chpxFkpOffset = body.HasCharacterFormatting
                ? AlignToSector(TextOffset + textBytes.Length)
                : 0;
            int papxFkpOffset = body.HasParagraphFormatting
                ? AlignToSector(body.HasCharacterFormatting ? chpxFkpOffset + OleSectorSize : TextOffset + textBytes.Length)
                : 0;
            int sectionDataOffset = GetSectionDataOffset(body, textBytes.Length, chpxFkpOffset, papxFkpOffset);
            IReadOnlyList<LegacyDocWritableSectionRecord> sectionRecords = CreateSectionRecords(body, sectionDataOffset);
            int streamLength = body.HasParagraphFormatting
                ? papxFkpOffset + OleSectorSize
                : body.HasCharacterFormatting
                    ? chpxFkpOffset + OleSectorSize
                    : TextOffset + textBytes.Length;
            if (body.HasSectionDescriptors) {
                streamLength = Math.Max(streamLength, sectionRecords.Count == 0 ? sectionDataOffset : sectionRecords.Max(record => record.EndOffset));
            }
            var stream = new byte[Math.Max(FibLength, streamLength)];
            WriteUInt16(stream, 0x00, WordDocumentMagic);
            WriteUInt16(stream, 0x02, Word97FibVersion);
            WriteUInt16(stream, 0x06, DefaultLanguageId);
            WriteUInt16(stream, 0x0A, OneTableStreamFlag | ExtendedCharacterFlag);
            WriteUInt16(stream, 0x0C, Word97FibBackVersion);
            WriteInt32(stream, 0x18, TextOffset);
            WriteInt32(stream, 0x1C, TextOffset + textBytes.Length);
            WriteUInt16(stream, 0x20, FibRgW97WordCount);
            WriteUInt16(stream, 0x3C, DefaultLanguageId);
            WriteUInt16(stream, 0x3E, FibRgLw97DwordCount);
            WriteInt32(stream, 0x4C, body.Text.Length);
            WriteUInt16(stream, 0x98, FibRgFcLcb97Size);
            WriteInt32(stream, 0xFA, body.HasCharacterFormatting ? ClxLength : 0);
            WriteInt32(stream, 0xFE, body.HasCharacterFormatting ? ChpxPlcLength : 0);
            WriteInt32(stream, FcPlcfBtePapxOffset, body.HasParagraphFormatting ? body.PapxPlcOffsetInTableStream : 0);
            WriteInt32(stream, LcbPlcfBtePapxOffset, body.HasParagraphFormatting ? PapxPlcLength : 0);
            WriteInt32(stream, FcPlcfSedOffset, body.HasSectionDescriptors ? body.SedPlcOffsetInTableStream : 0);
            WriteInt32(stream, LcbPlcfSedOffset, body.HasSectionDescriptors ? body.SedPlcLength : 0);
            WriteInt32(stream, FcSttbfFfnOffset, body.HasFontTable ? body.FontTableOffsetInTableStream : 0);
            WriteInt32(stream, LcbSttbfFfnOffset, fontTable.Length);
            WriteInt32(stream, 0x1A2, 0);
            WriteInt32(stream, 0x1A6, ClxLength);
            Buffer.BlockCopy(textBytes, 0, stream, TextOffset, textBytes.Length);
            if (body.HasCharacterFormatting) {
                WriteChpxFkp(stream, chpxFkpOffset, body.CreateFormattingSegments(), body.FontFamilyIndexes, bytesPerCharacter);
            }

            if (body.HasParagraphFormatting) {
                LegacyDocParagraphFormattingWriter.WritePapxFkp(stream, papxFkpOffset, TextOffset, OleSectorSize, body.CreateParagraphSegments(), bytesPerCharacter);
            }

            if (body.HasSectionDescriptors) {
                foreach (LegacyDocWritableSectionRecord record in sectionRecords) {
                    if (record.Sepx.Length == 0) {
                        continue;
                    }

                    Buffer.BlockCopy(record.Sepx, 0, stream, record.SepxOffset, record.Sepx.Length);
                }
            }

            return stream;
        }

        private static byte[] CreateTableStream(LegacyDocWritableBody body) {
            byte[] fontTable = CreateFontTable(body.FontFamilies);
            bool compressedText = CanWriteCompressedText(body.Text);
            int bytesPerCharacter = compressedText ? 1 : 2;
            int textByteLength = checked(body.Text.Length * bytesPerCharacter);
            var table = new byte[ClxLength + (body.HasCharacterFormatting ? ChpxPlcLength : 0) + (body.HasParagraphFormatting ? PapxPlcLength : 0) + (body.HasSectionDescriptors ? body.SedPlcLength : 0) + fontTable.Length];
            table[0] = 0x02;
            WriteInt32(table, 1, 16);
            WriteInt32(table, 5, 0);
            WriteInt32(table, 9, body.Text.Length);
            WriteUInt16(table, 13, DefaultPcdFlags);
            WriteUInt32(table, 15, compressedText ? CompressedTextFlag | (uint)(TextOffset * 2) : TextOffset);
            WriteUInt16(table, 19, 0);

            if (body.HasCharacterFormatting) {
                int chpxFkpOffset = AlignToSector(TextOffset + textByteLength);
                WriteInt32(table, ClxLength, TextOffset);
                WriteInt32(table, ClxLength + 4, TextOffset + textByteLength);
                WriteInt32(table, ClxLength + 8, chpxFkpOffset / OleSectorSize);
            }

            if (body.HasParagraphFormatting) {
                int chpxFkpOffset = body.HasCharacterFormatting
                    ? AlignToSector(TextOffset + textByteLength)
                    : 0;
                int papxFkpOffset = AlignToSector(body.HasCharacterFormatting ? chpxFkpOffset + OleSectorSize : TextOffset + textByteLength);
                WriteInt32(table, body.PapxPlcOffsetInTableStream, TextOffset);
                WriteInt32(table, body.PapxPlcOffsetInTableStream + 4, TextOffset + textByteLength);
                WriteInt32(table, body.PapxPlcOffsetInTableStream + 8, papxFkpOffset / OleSectorSize);
            }

            if (body.HasSectionDescriptors) {
                int chpxFkpOffset = body.HasCharacterFormatting
                    ? AlignToSector(TextOffset + textByteLength)
                    : 0;
                int papxFkpOffset = body.HasParagraphFormatting
                    ? AlignToSector(body.HasCharacterFormatting ? chpxFkpOffset + OleSectorSize : TextOffset + textByteLength)
                    : 0;
                int sepxOffset = AlignToEven(body.HasParagraphFormatting
                    ? papxFkpOffset + OleSectorSize
                    : body.HasCharacterFormatting
                        ? chpxFkpOffset + OleSectorSize
                        : TextOffset + textByteLength);
                WritePlcfSed(table, body.SedPlcOffsetInTableStream, CreateSectionRecords(body, sepxOffset));
            }

            if (fontTable.Length > 0) {
                Buffer.BlockCopy(fontTable, 0, table, body.FontTableOffsetInTableStream, fontTable.Length);
            }

            return table;
        }

        private static bool CanWriteCompressedText(string text) {
            foreach (char character in text) {
                if (character > 0x7F) {
                    return false;
                }
            }

            return true;
        }

        private static byte[] EncodeCompressedText(string text) {
            byte[] bytes = new byte[text.Length];
            for (int i = 0; i < text.Length; i++) {
                bytes[i] = (byte)text[i];
            }

            return bytes;
        }

        private static int AlignToSector(int value) {
            return ((value + OleSectorSize - 1) / OleSectorSize) * OleSectorSize;
        }

        private static int AlignToEven(int value) {
            return value % 2 == 0 ? value : value + 1;
        }

        private static int GetSectionDataOffset(LegacyDocWritableBody body, int textByteLength, int chpxFkpOffset, int papxFkpOffset) {
            return AlignToEven(body.HasParagraphFormatting
                ? papxFkpOffset + OleSectorSize
                : body.HasCharacterFormatting
                    ? chpxFkpOffset + OleSectorSize
                    : TextOffset + textByteLength);
        }

        private static IReadOnlyList<LegacyDocWritableSectionRecord> CreateSectionRecords(LegacyDocWritableBody body, int firstSepxOffset) {
            var records = new List<LegacyDocWritableSectionRecord>(body.Sections.Count);
            int sepxOffset = firstSepxOffset;
            foreach (LegacyDocWritableSection section in body.Sections) {
                byte[] sepx = section.Format.HasFormatting ? CreateSepx(section.Format) : Array.Empty<byte>();
                int recordSepxOffset = 0;
                if (sepx.Length > 0) {
                    recordSepxOffset = sepxOffset;
                    sepxOffset = AlignToEven(sepxOffset + sepx.Length);
                }

                records.Add(new LegacyDocWritableSectionRecord(section.EndCharacter, recordSepxOffset, sepx));
            }

            return records;
        }

        private static void WriteUInt16(byte[] bytes, int offset, ushort value) {
            bytes[offset] = (byte)value;
            bytes[offset + 1] = (byte)(value >> 8);
        }

        private static void WriteUInt16(Stream stream, ushort value) {
            stream.WriteByte((byte)(value & 0xFF));
            stream.WriteByte((byte)(value >> 8));
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
            internal LegacyDocWritableBody(
                string text,
                IReadOnlyList<LegacyDocWritableRun> formattedRuns,
                IReadOnlyList<LegacyDocWritableParagraph> formattedParagraphs,
                IReadOnlyList<LegacyDocWritableSection> sections) {
                Text = text;
                FormattedRuns = formattedRuns;
                FormattedParagraphs = formattedParagraphs;
                Sections = sections;
                FontFamilies = formattedRuns
                    .Select(run => run.Formatting.FontFamily)
                    .Where(fontFamily => !string.IsNullOrWhiteSpace(fontFamily))
                    .Select(fontFamily => fontFamily!)
                    .Distinct(StringComparer.OrdinalIgnoreCase)
                    .ToArray();
                FontFamilyIndexes = FontFamilies
                    .Select((fontFamily, index) => new { fontFamily, index })
                    .ToDictionary(item => item.fontFamily, item => item.index, StringComparer.OrdinalIgnoreCase);
            }

            internal string Text { get; }

            internal IReadOnlyList<LegacyDocWritableRun> FormattedRuns { get; }

            internal IReadOnlyList<LegacyDocWritableParagraph> FormattedParagraphs { get; }

            internal IReadOnlyList<LegacyDocWritableSection> Sections { get; }

            internal IReadOnlyList<string> FontFamilies { get; }

            internal IReadOnlyDictionary<string, int> FontFamilyIndexes { get; }

            internal bool HasCharacterFormatting => FormattedRuns.Count > 0;

            internal bool HasParagraphFormatting => FormattedParagraphs.Count > 0;

            internal bool HasFontTable => FontFamilies.Count > 0;

            internal bool HasSectionDescriptors => Sections.Count > 1 || Sections.Any(section => section.Format.HasFormatting);

            internal int PapxPlcOffsetInTableStream => ClxLength + (HasCharacterFormatting ? ChpxPlcLength : 0);

            internal int SedPlcOffsetInTableStream => ClxLength + (HasCharacterFormatting ? ChpxPlcLength : 0) + (HasParagraphFormatting ? PapxPlcLength : 0);

            internal int SedPlcLength => 4 + (Sections.Count * (4 + SedLength));

            internal int FontTableOffsetInTableStream => ClxLength + (HasCharacterFormatting ? ChpxPlcLength : 0) + (HasParagraphFormatting ? PapxPlcLength : 0) + (HasSectionDescriptors ? SedPlcLength : 0);

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

            internal IReadOnlyList<LegacyDocWritableParagraphSegment> CreateParagraphSegments() {
                var segments = new List<LegacyDocWritableParagraphSegment>();
                int character = 0;
                foreach (LegacyDocWritableParagraph paragraph in FormattedParagraphs.OrderBy(item => item.StartCharacter)) {
                    if (paragraph.StartCharacter > character) {
                        AddParagraphSegment(segments, character, paragraph.StartCharacter - character, LegacyDocWritableParagraphFormatting.Plain);
                    }

                    AddParagraphSegment(segments, paragraph.StartCharacter, paragraph.Length, paragraph.Formatting);
                    character = paragraph.EndCharacter;
                }

                if (character < Text.Length) {
                    AddParagraphSegment(segments, character, Text.Length - character, LegacyDocWritableParagraphFormatting.Plain);
                }

                return segments;
            }

            private static void AddParagraphSegment(
                List<LegacyDocWritableParagraphSegment> segments,
                int startCharacter,
                int length,
                LegacyDocWritableParagraphFormatting formatting) {
                if (length <= 0) {
                    return;
                }

                if (segments.Count > 0) {
                    LegacyDocWritableParagraphSegment previous = segments[segments.Count - 1];
                    if (previous.EndCharacter == startCharacter && previous.Formatting.Equals(formatting)) {
                        segments[segments.Count - 1] = previous.Extend(length);
                        return;
                    }
                }

                segments.Add(new LegacyDocWritableParagraphSegment(startCharacter, length, formatting));
            }
        }

        private readonly struct LegacyDocWritableSection {
            internal LegacyDocWritableSection(int endCharacter, LegacyDocSectionFormat format) {
                EndCharacter = endCharacter;
                Format = format;
            }

            internal int EndCharacter { get; }

            internal LegacyDocSectionFormat Format { get; }
        }

        private readonly struct LegacyDocWritableSectionRecord {
            internal LegacyDocWritableSectionRecord(int endCharacter, int sepxOffset, byte[] sepx) {
                EndCharacter = endCharacter;
                SepxOffset = sepxOffset;
                Sepx = sepx;
            }

            internal int EndCharacter { get; }

            internal int SepxOffset { get; }

            internal byte[] Sepx { get; }

            internal int EndOffset => Sepx.Length == 0 ? SepxOffset : SepxOffset + Sepx.Length;
        }

    }
}
