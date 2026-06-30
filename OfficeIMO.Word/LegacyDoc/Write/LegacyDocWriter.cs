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
        private const ushort SprmCFImprint = 0x0854;
        private const ushort SprmCFSmallCaps = 0x083A;
        private const ushort SprmCFCaps = 0x083B;
        private const ushort SprmCFVanish = 0x083C;
        private const ushort SprmCFEmboss = 0x0858;
        private const ushort SprmCFSpec = 0x0855;
        private const ushort SprmCHighlight = 0x2A0C;
        private const ushort SprmCKul = 0x2A3E;
        private const ushort SprmCIss = 0x2A48;
        private const ushort SprmCHps = 0x4A43;
        private const ushort SprmCRgFtc0 = 0x4A4F;
        private const ushort SprmCFDStrike = 0x2A53;
        private const ushort SprmCCv = 0x6870;
        private const ushort SprmPJc = 0x2461;
        private const ushort DefaultPcdFlags = 0x0310;
        private const ushort FootnotePcdFlags = 0x0330;
        private const ushort WordDocumentMagic = 0xA5EC;
        private const ushort Word97FibVersion = 0x00C1;
        private const ushort Word97FibBackVersion = 0x00BF;
        private const ushort DefaultLanguageId = 0x0409;
        private const ushort FibRgW97WordCount = 0x000E;
        private const ushort FibRgLw97DwordCount = 0x0016;
        private const ushort FibRgFcLcb97Size = 0x00B7;
        private const ushort OneTableStreamFlag = 0x0200;
        private const ushort ExtendedCharacterFlag = 0x1000;
        private const ushort DefaultFibFlags = 0x1200;
        private const int FcPlcfSedOffset = 0xCA;
        private const int LcbPlcfSedOffset = 0xCE;
        private const int FcPlcffndRefOffset = 0xAA;
        private const int LcbPlcffndRefOffset = 0xAE;
        private const int FcPlcffndTxtOffset = 0xB2;
        private const int LcbPlcffndTxtOffset = 0xB6;
        private const int FcPlcfendRefOffset = 0x20A;
        private const int LcbPlcfendRefOffset = 0x20E;
        private const int FcPlcfendTxtOffset = 0x212;
        private const int LcbPlcfendTxtOffset = 0x216;
        private const int FcStshfOffset = 0xA2;
        private const int LcbStshfOffset = 0xA6;
        private const int FcDopOffset = 0x192;
        private const int LcbDopOffset = 0x196;
        private const int FcPlcfBtePapxOffset = 0x102;
        private const int LcbPlcfBtePapxOffset = 0x106;
        private const int FcSttbfFfnOffset = 0x112;
        private const int LcbSttbfFfnOffset = 0x116;
        private const int FcPlcfHddOffset = 0xF2;
        private const int LcbPlcfHddOffset = 0xF6;
        private const int DopBaseLength = 8;
        private const int DopBaseEndnotePlacementLength = 56;
        private const int DopBaseEndnotePlacementOffset = 52;
        private const int DopBaseEndnotePlacementShift = 16;
        private const ushort FacingPagesDopFlag = 0x0001;
        private static readonly byte[] PlainParagraphPapx = { 0x00, 0x01, 0x00, 0x00 };
        private static readonly byte[] FootnoteTextParagraphPapx = { 0x00, 0x01, 0x23, 0x00 };

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
            LegacyDocWritableFootnotes footnotes = ReadSupportedFootnotes(mainPart!);
            LegacyDocWritableEndnotes endnotes = ReadSupportedEndnotes(mainPart!);
            LegacyDocWritableStyleSheet styleSheet = CreateWritableStyleSheet(mainPart!, body);
            IReadOnlyDictionary<string, Style> tableStyleDefinitions = ReadTableStyleDefinitions(mainPart!);
            LegacyDocSectionFormat finalSectionFormat = LegacyDocSectionFormat.Default;
            var sections = new List<LegacyDocWritableSection>();
            SectionMarkValues? pendingSectionBreakType = null;
            int bodyContentCount = 0;
            foreach (OpenXmlElement child in body.ChildElements) {
                switch (child) {
                    case Paragraph paragraph:
                        if (!IsPureSectionBreakParagraph(paragraph)) {
                            AppendParagraph(text, runs, paragraphFormats, paragraph, styleSheet.StyleIndexes, footnotes, endnotes);
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
                        AppendTable(text, runs, paragraphFormats, table, styleSheet.StyleIndexes, tableStyleDefinitions, footnotes, endnotes);
                        bodyContentCount++;
                        break;
                    case SectionProperties sectionProperties:
                        finalSectionFormat = ReadSupportedSectionProperties(sectionProperties);
                        break;
                    default:
                        throw new NotSupportedException($"Native DOC saving currently supports body paragraphs and simple tables with bold, italic, strikethrough, double-strikethrough, outline, shadow, emboss, imprint, hidden text, caps/small-caps, superscript/subscript, underline, highlight, font size, color, and font family text runs. Unsupported body element: {child.LocalName}.");
                }
            }

            if (bodyContentCount == 0) {
                text.Append('\r');
            }

            AddSection(sections, text.Length, finalSectionFormat.WithSectionBreakType(pendingSectionBreakType));
            footnotes.ThrowIfUnreferencedFootnotesRemain();
            endnotes.ThrowIfUnreferencedEndnotesRemain();
            bool hasNoteReferences = footnotes.HasReferences || endnotes.HasReferences;
            LegacyDocWritableHeaderFooterStories headerFooterStories = BuildHeaderFooterStories(document, mainPart!, hasNoteReferences, styleSheet.StyleIndexes);
            int terminalCharacterPadding = hasNoteReferences ? 1 : 0;
            LegacyDocWritableFootnoteStories footnoteStories = footnotes.CreateStories(text.Length, headerFooterStories.Text.Length, terminalCharacterPadding);
            LegacyDocWritableEndnoteStories endnoteStories = endnotes.CreateStories(text.Length, footnoteStories.Text.Length, headerFooterStories.Text.Length, terminalCharacterPadding);
            return new LegacyDocWritableBody(text.ToString(), runs, paragraphFormats, sections, styleSheet, footnoteStories, endnoteStories, headerFooterStories, HasEvenAndOddHeaders(mainPart!), ReadDocumentEndnotePosition(sections));
        }

        private static bool HasEvenAndOddHeaders(DocumentFormat.OpenXml.Packaging.MainDocumentPart mainPart) {
            Settings? settings = mainPart.DocumentSettingsPart?.Settings;
            return settings?.Elements<EvenAndOddHeaders>().Any(IsOnOffEnabled) == true;
        }

        private static EndnotePositionValues? ReadDocumentEndnotePosition(IReadOnlyList<LegacyDocWritableSection> sections) {
            EndnotePositionValues? position = null;
            foreach (LegacyDocWritableSection section in sections) {
                EndnotePositionValues? sectionPosition = section.Format.EndnotePosition;
                if (sectionPosition == null) {
                    continue;
                }

                if (position != null && position.Value != sectionPosition.Value) {
                    throw new NotSupportedException("Native DOC saving supports only one endnote placement for the whole document.");
                }

                position = sectionPosition;
            }

            return position;
        }

        private static void ThrowIfUnsupportedDocumentParts(WordDocument document, DocumentFormat.OpenXml.Packaging.MainDocumentPart? mainPart) {
            if (mainPart == null) {
                return;
            }

            ThrowIfUnsupportedReviewMarkup(mainPart);

            if (mainPart.ImageParts.Any()) {
                throw new NotSupportedException("Native DOC saving currently supports text only. Images are not supported yet.");
            }

            if (mainPart.ChartParts.Any()) {
                throw new NotSupportedException("Native DOC saving currently supports text only. Charts are not supported yet.");
            }
        }

        private static void ThrowIfUnsupportedReviewMarkup(DocumentFormat.OpenXml.Packaging.MainDocumentPart mainPart) {
            Settings? settings = mainPart.DocumentSettingsPart?.Settings;
            if (settings?.GetFirstChild<TrackRevisions>() != null) {
                throw new NotSupportedException("Native DOC saving currently does not support revision tracking settings. Disable tracking, accept or reject revisions, or save as DOCX before saving as DOC.");
            }

            Body? body = mainPart.Document?.Body;
            if (body != null && HasTrackedRevisionMarkup(body)) {
                throw new NotSupportedException("Native DOC saving currently does not support tracked revision markup. Accept or reject revisions, or save as DOCX before saving as DOC.");
            }

            if (HasComments(mainPart, body)) {
                throw new NotSupportedException("Native DOC saving currently does not support comments. Remove comments, or save as DOCX before saving as DOC.");
            }
        }

        private static bool HasTrackedRevisionMarkup(Body body) {
            return body.Descendants<InsertedRun>().Any()
                || body.Descendants<DeletedRun>().Any()
                || body.Descendants<MoveFromRun>().Any()
                || body.Descendants<MoveToRun>().Any();
        }

        private static bool HasComments(DocumentFormat.OpenXml.Packaging.MainDocumentPart mainPart, Body? body) {
            if (mainPart.WordprocessingCommentsPart?.Comments?.Elements<Comment>().Any() == true) {
                return true;
            }

            return body != null
                && (body.Descendants<CommentRangeStart>().Any()
                    || body.Descendants<CommentRangeEnd>().Any()
                    || body.Descendants<CommentReference>().Any());
        }

        private static bool IsUserEndnote(Endnote endnote) {
            return endnote.Type == null || endnote.Type.Value == FootnoteEndnoteValues.Normal;
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

        private static void AppendParagraph(StringBuilder text, List<LegacyDocWritableRun> runs, List<LegacyDocWritableParagraph> paragraphFormats, Paragraph paragraph, IReadOnlyDictionary<string, ushort> styleIndexes, LegacyDocWritableFootnotes footnotes, LegacyDocWritableEndnotes endnotes) {
            LegacyDocWritableParagraphFormatting paragraphFormatting = ReadSupportedParagraphFormatting(paragraph.ParagraphProperties, styleIndexes);
            int paragraphStart = text.Length;

            foreach (OpenXmlElement child in paragraph.ChildElements) {
                switch (child) {
                    case ParagraphProperties:
                        break;
                    case Run run:
                        AppendSupportedRunText(text, runs, run, footnotes, endnotes);
                        break;
                    default:
                        throw new NotSupportedException($"Native DOC saving currently supports only text runs with bold, italic, strikethrough, double-strikethrough, outline, shadow, emboss, imprint, hidden text, caps/small-caps, superscript/subscript, underline, highlight, font size, color, and font family formatting. Unsupported paragraph element: {child.LocalName}.");
                }
            }

            text.Append('\r');
            if (paragraphFormatting.HasFormatting) {
                paragraphFormats.Add(new LegacyDocWritableParagraph(paragraphStart, text.Length - paragraphStart, paragraphFormatting));
            }
        }

        private static byte[] CreateWordDocumentStream(LegacyDocWritableBody body) {
            bool compressedText = CanWriteCompressedText(body.StoredText);
            int bytesPerCharacter = compressedText ? 1 : 2;
            byte[] textBytes = compressedText ? EncodeCompressedText(body.StoredText) : Encoding.Unicode.GetBytes(body.StoredText);
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
            WriteUInt16(stream, 0x0A, DefaultFibFlags);
            WriteUInt16(stream, 0x0C, Word97FibBackVersion);
            WriteInt32(stream, 0x18, TextOffset);
            WriteInt32(stream, 0x1C, TextOffset + textBytes.Length);
            WriteUInt16(stream, 0x20, FibRgW97WordCount);
            WriteUInt16(stream, 0x3C, DefaultLanguageId);
            WriteUInt16(stream, 0x3E, FibRgLw97DwordCount);
            WriteInt32(stream, 0x4C, body.Text.Length);
            WriteInt32(stream, 0x50, body.FootnoteText.Length);
            WriteInt32(stream, 0x54, body.HeaderFooterText.Length);
            WriteInt32(stream, 0x60, body.EndnoteText.Length);
            WriteUInt16(stream, 0x98, FibRgFcLcb97Size);
            WriteInt32(stream, FcStshfOffset, body.HasStyleSheet ? body.StyleSheetOffsetInTableStream : 0);
            WriteInt32(stream, LcbStshfOffset, body.StyleSheet.Bytes.Length);
            WriteInt32(stream, 0xFA, body.HasCharacterFormatting ? ClxLength : 0);
            WriteInt32(stream, 0xFE, body.HasCharacterFormatting ? ChpxPlcLength : 0);
            WriteInt32(stream, FcPlcfBtePapxOffset, body.HasParagraphFormatting ? body.PapxPlcOffsetInTableStream : 0);
            WriteInt32(stream, LcbPlcfBtePapxOffset, body.HasParagraphFormatting ? PapxPlcLength : 0);
            WriteInt32(stream, FcPlcfSedOffset, body.HasSectionDescriptors ? body.SedPlcOffsetInTableStream : 0);
            WriteInt32(stream, LcbPlcfSedOffset, body.HasSectionDescriptors ? body.SedPlcLength : 0);
            WriteInt32(stream, FcPlcffndRefOffset, body.HasFootnotes ? body.PlcffndRefOffsetInTableStream : 0);
            WriteInt32(stream, LcbPlcffndRefOffset, body.HasFootnotes ? body.PlcffndRef.Length : 0);
            WriteInt32(stream, FcPlcffndTxtOffset, body.HasFootnotes ? body.PlcffndTxtOffsetInTableStream : 0);
            WriteInt32(stream, LcbPlcffndTxtOffset, body.HasFootnotes ? body.PlcffndTxt.Length : 0);
            WriteInt32(stream, FcPlcfendRefOffset, body.HasEndnotes ? body.PlcfendRefOffsetInTableStream : 0);
            WriteInt32(stream, LcbPlcfendRefOffset, body.HasEndnotes ? body.PlcfendRef.Length : 0);
            WriteInt32(stream, FcPlcfendTxtOffset, body.HasEndnotes ? body.PlcfendTxtOffsetInTableStream : 0);
            WriteInt32(stream, LcbPlcfendTxtOffset, body.HasEndnotes ? body.PlcfendTxt.Length : 0);
            WriteInt32(stream, FcPlcfHddOffset, body.HasHeaderFooterStories ? body.PlcfHddOffsetInTableStream : 0);
            WriteInt32(stream, LcbPlcfHddOffset, body.HasHeaderFooterStories ? body.PlcfHdd.Length : 0);
            WriteInt32(stream, FcSttbfFfnOffset, body.HasFontTable ? body.FontTableOffsetInTableStream : 0);
            WriteInt32(stream, LcbSttbfFfnOffset, fontTable.Length);
            WriteInt32(stream, FcDopOffset, body.HasDocumentOptions ? body.DopOffsetInTableStream : 0);
            WriteInt32(stream, LcbDopOffset, body.HasDocumentOptions ? body.DopLength : 0);
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
            bool compressedText = CanWriteCompressedText(body.StoredText);
            int bytesPerCharacter = compressedText ? 1 : 2;
            int textByteLength = checked(body.StoredText.Length * bytesPerCharacter);
            int pieceTableByteLength = checked(body.PieceTableCharacterCount * bytesPerCharacter);
            var table = new byte[body.FontTableOffsetInTableStream + fontTable.Length];
            table[0] = 0x02;
            WriteInt32(table, 1, 16);
            WriteInt32(table, 5, 0);
            WriteInt32(table, 9, body.PieceTableCharacterCount);
            WriteUInt16(table, 13, body.HasFootnotes ? FootnotePcdFlags : DefaultPcdFlags);
            WriteUInt32(table, 15, compressedText ? CompressedTextFlag | (uint)(TextOffset * 2) : TextOffset);
            WriteUInt16(table, 19, 0);

            if (body.HasCharacterFormatting) {
                int chpxFkpOffset = AlignToSector(TextOffset + textByteLength);
                WriteInt32(table, ClxLength, TextOffset);
                WriteInt32(table, ClxLength + 4, TextOffset + pieceTableByteLength);
                WriteInt32(table, ClxLength + 8, chpxFkpOffset / OleSectorSize);
            }

            if (body.HasParagraphFormatting) {
                int chpxFkpOffset = body.HasCharacterFormatting
                    ? AlignToSector(TextOffset + textByteLength)
                    : 0;
                int papxFkpOffset = AlignToSector(body.HasCharacterFormatting ? chpxFkpOffset + OleSectorSize : TextOffset + textByteLength);
                WriteInt32(table, body.PapxPlcOffsetInTableStream, TextOffset);
                WriteInt32(table, body.PapxPlcOffsetInTableStream + 4, TextOffset + pieceTableByteLength);
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

            if (body.HasFootnotes) {
                Buffer.BlockCopy(body.PlcffndRef, 0, table, body.PlcffndRefOffsetInTableStream, body.PlcffndRef.Length);
                Buffer.BlockCopy(body.PlcffndTxt, 0, table, body.PlcffndTxtOffsetInTableStream, body.PlcffndTxt.Length);
            }

            if (body.HasHeaderFooterStories) {
                Buffer.BlockCopy(body.PlcfHdd, 0, table, body.PlcfHddOffsetInTableStream, body.PlcfHdd.Length);
            }

            if (body.HasEndnotes) {
                Buffer.BlockCopy(body.PlcfendRef, 0, table, body.PlcfendRefOffsetInTableStream, body.PlcfendRef.Length);
                Buffer.BlockCopy(body.PlcfendTxt, 0, table, body.PlcfendTxtOffsetInTableStream, body.PlcfendTxt.Length);
            }

            if (body.HasDocumentOptions) {
                byte[] dop = CreateDopBase(body);
                Buffer.BlockCopy(dop, 0, table, body.DopOffsetInTableStream, dop.Length);
            }

            if (body.HasStyleSheet) {
                Buffer.BlockCopy(body.StyleSheet.Bytes, 0, table, body.StyleSheetOffsetInTableStream, body.StyleSheet.Bytes.Length);
            }

            if (fontTable.Length > 0) {
                Buffer.BlockCopy(fontTable, 0, table, body.FontTableOffsetInTableStream, fontTable.Length);
            }

            return table;
        }

        private static byte[] CreateDopBase(LegacyDocWritableBody body) {
            var dop = new byte[body.DopLength];
            if (body.FacingPages) {
                WriteUInt16(dop, 0, FacingPagesDopFlag);
            }

            if (body.EndnotePosition != null) {
                uint placement = (uint)GetEndnotePositionOperand(body.EndnotePosition.Value)!.Value;
                WriteUInt32(dop, DopBaseEndnotePlacementOffset, placement << DopBaseEndnotePlacementShift);
            }

            return dop;
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
                IReadOnlyList<LegacyDocWritableSection> sections,
                LegacyDocWritableStyleSheet styleSheet,
                LegacyDocWritableFootnoteStories footnoteStories,
                LegacyDocWritableEndnoteStories endnoteStories,
                LegacyDocWritableHeaderFooterStories headerFooterStories,
                bool facingPages,
                EndnotePositionValues? endnotePosition) {
                Text = text;
                FormattedRuns = formattedRuns;
                FormattedParagraphs = formattedParagraphs;
                Sections = sections;
                StyleSheet = styleSheet;
                FootnoteText = footnoteStories.Text;
                PlcffndRef = footnoteStories.PlcffndRef;
                PlcffndTxt = footnoteStories.PlcffndTxt;
                FootnoteMarkerPositions = footnoteStories.MarkerPositions;
                FootnoteFormattedRuns = footnoteStories.FormattedRuns;
                EndnoteText = endnoteStories.Text;
                PlcfendRef = endnoteStories.PlcfendRef;
                PlcfendTxt = endnoteStories.PlcfendTxt;
                EndnoteMarkerPositions = endnoteStories.MarkerPositions;
                EndnoteFormattedRuns = endnoteStories.FormattedRuns;
                HeaderFooterText = headerFooterStories.Text;
                PlcfHdd = headerFooterStories.PlcfHdd;
                HeaderFooterMarkerPositions = headerFooterStories.MarkerPositions;
                HeaderFooterFormattedRuns = headerFooterStories.FormattedRuns;
                HeaderFooterFormattedParagraphs = headerFooterStories.FormattedParagraphs;
                FacingPages = facingPages;
                EndnotePosition = endnotePosition;
                FontFamilies = styleSheet.FontFamilies
                    .Concat(formattedRuns.Select(run => run.Formatting.FontFamily))
                    .Concat(FootnoteFormattedRuns.Select(run => run.Formatting.FontFamily))
                    .Concat(HeaderFooterFormattedRuns.Select(run => run.Formatting.FontFamily))
                    .Concat(EndnoteFormattedRuns.Select(run => run.Formatting.FontFamily))
                    .Where(fontFamily => !string.IsNullOrWhiteSpace(fontFamily))
                    .Select(fontFamily => fontFamily!)
                    .Distinct(StringComparer.OrdinalIgnoreCase)
                    .ToArray();
                FontFamilyIndexes = FontFamilies
                    .Select((fontFamily, index) => new { fontFamily, index })
                    .ToDictionary(item => item.fontFamily, item => item.index, StringComparer.OrdinalIgnoreCase);
            }

            internal string Text { get; }

            internal string HeaderFooterText { get; }

            internal string FootnoteText { get; }

            internal string EndnoteText { get; }

            internal string FullText => Text + FootnoteText + HeaderFooterText + EndnoteText;

            internal bool HasNoteStories => HasFootnotes || HasEndnotes;

            internal string StoredText => HasNoteStories ? FullText + "\r" : FullText;

            internal int PieceTableCharacterCount => HasNoteStories ? FullText.Length + 1 : FullText.Length;

            internal byte[] PlcffndRef { get; }

            internal byte[] PlcffndTxt { get; }

            internal IReadOnlyList<int> FootnoteMarkerPositions { get; }

            internal IReadOnlyList<LegacyDocWritableRun> FootnoteFormattedRuns { get; }

            internal byte[] PlcfendRef { get; }

            internal byte[] PlcfendTxt { get; }

            internal IReadOnlyList<int> EndnoteMarkerPositions { get; }

            internal IReadOnlyList<LegacyDocWritableRun> EndnoteFormattedRuns { get; }

            internal IReadOnlyList<int> HeaderFooterMarkerPositions { get; }

            internal IReadOnlyList<LegacyDocWritableRun> HeaderFooterFormattedRuns { get; }

            internal IReadOnlyList<LegacyDocWritableParagraph> HeaderFooterFormattedParagraphs { get; }

            internal byte[] PlcfHdd { get; }

            internal bool FacingPages { get; }

            internal EndnotePositionValues? EndnotePosition { get; }

            internal IReadOnlyList<LegacyDocWritableRun> FormattedRuns { get; }

            internal IReadOnlyList<LegacyDocWritableParagraph> FormattedParagraphs { get; }

            internal IReadOnlyList<LegacyDocWritableSection> Sections { get; }

            internal LegacyDocWritableStyleSheet StyleSheet { get; }

            internal IReadOnlyList<string> FontFamilies { get; }

            internal IReadOnlyDictionary<string, int> FontFamilyIndexes { get; }

            internal bool HasCharacterFormatting => FormattedRuns.Count > 0 || FootnoteFormattedRuns.Count > 0 || HeaderFooterFormattedRuns.Count > 0 || EndnoteFormattedRuns.Count > 0;

            internal bool HasParagraphFormatting => FormattedParagraphs.Count > 0 || HeaderFooterFormattedParagraphs.Count > 0 || HasNoteStories;

            internal bool HasFontTable => FontFamilies.Count > 0;

            internal bool HasStyleSheet => StyleSheet.Bytes.Length > 0;

            internal bool HasSectionDescriptors => Sections.Count > 1 || Sections.Any(section => section.Format.HasFormatting);

            internal bool HasHeaderFooterStories => HeaderFooterText.Length > 0 && PlcfHdd.Length > 0;

            internal bool HasFootnotes => FootnoteText.Length > 0 && PlcffndRef.Length > 0 && PlcffndTxt.Length > 0;

            internal bool HasEndnotes => EndnoteText.Length > 0 && PlcfendRef.Length > 0 && PlcfendTxt.Length > 0;

            internal bool HasDocumentOptions => FacingPages || EndnotePosition != null;

            internal int DopLength => EndnotePosition != null ? DopBaseEndnotePlacementLength : DopBaseLength;

            internal int PapxPlcOffsetInTableStream => ClxLength + (HasCharacterFormatting ? ChpxPlcLength : 0);

            internal int SedPlcOffsetInTableStream => ClxLength + (HasCharacterFormatting ? ChpxPlcLength : 0) + (HasParagraphFormatting ? PapxPlcLength : 0);

            internal int SedPlcLength => 4 + (Sections.Count * (4 + SedLength));

            private int AfterSectionDataOffsetInTableStream => ClxLength + (HasCharacterFormatting ? ChpxPlcLength : 0) + (HasParagraphFormatting ? PapxPlcLength : 0) + (HasSectionDescriptors ? SedPlcLength : 0);

            internal int PlcffndRefOffsetInTableStream => AfterSectionDataOffsetInTableStream;

            internal int PlcffndTxtOffsetInTableStream => AfterSectionDataOffsetInTableStream + (HasFootnotes ? PlcffndRef.Length : 0);

            private int AfterFootnoteDataOffsetInTableStream => AfterSectionDataOffsetInTableStream + (HasFootnotes ? PlcffndRef.Length + PlcffndTxt.Length : 0);

            internal int PlcfHddOffsetInTableStream => AfterFootnoteDataOffsetInTableStream;

            private int AfterHeaderFooterDataOffsetInTableStream => AfterFootnoteDataOffsetInTableStream + (HasHeaderFooterStories ? PlcfHdd.Length : 0);

            internal int PlcfendRefOffsetInTableStream => AfterHeaderFooterDataOffsetInTableStream;

            internal int PlcfendTxtOffsetInTableStream => AfterHeaderFooterDataOffsetInTableStream + (HasEndnotes ? PlcfendRef.Length : 0);

            private int AfterEndnoteDataOffsetInTableStream => AfterHeaderFooterDataOffsetInTableStream + (HasEndnotes ? PlcfendRef.Length + PlcfendTxt.Length : 0);

            internal int DopOffsetInTableStream => HasDocumentOptions ? AlignToEven(AfterEndnoteDataOffsetInTableStream) : AfterEndnoteDataOffsetInTableStream;

            private int AfterDocumentOptionsOffsetInTableStream => HasDocumentOptions ? DopOffsetInTableStream + DopLength : AfterEndnoteDataOffsetInTableStream;

            internal int StyleSheetOffsetInTableStream => HasStyleSheet ? AlignToEven(AfterDocumentOptionsOffsetInTableStream) : AfterDocumentOptionsOffsetInTableStream;

            internal int FontTableOffsetInTableStream => HasStyleSheet
                ? StyleSheetOffsetInTableStream + StyleSheet.Bytes.Length
                : AfterDocumentOptionsOffsetInTableStream;

            internal IReadOnlyList<LegacyDocWritableSegment> CreateFormattingSegments() {
                var segments = new List<LegacyDocWritableSegment>();
                int character = 0;
                foreach (LegacyDocWritableRun run in CreateFormattedRuns().OrderBy(item => item.StartCharacter)) {
                    if (run.StartCharacter > character) {
                        AddSegment(segments, character, run.StartCharacter - character, LegacyDocWritableFormatting.Plain);
                    }

                    AddSegment(segments, run.StartCharacter, run.Length, run.Formatting);
                    character = run.EndCharacter;
                }

                if (character < PieceTableCharacterCount) {
                    AddSegment(segments, character, PieceTableCharacterCount - character, LegacyDocWritableFormatting.Plain);
                }

                return segments;
            }

            private IReadOnlyList<LegacyDocWritableRun> CreateFormattedRuns() {
                if (FootnoteMarkerPositions.Count == 0
                    && FootnoteFormattedRuns.Count == 0
                    && HeaderFooterMarkerPositions.Count == 0
                    && HeaderFooterFormattedRuns.Count == 0
                    && EndnoteMarkerPositions.Count == 0
                    && EndnoteFormattedRuns.Count == 0) {
                    return FormattedRuns;
                }

                var runs = new List<LegacyDocWritableRun>(
                    FormattedRuns.Count
                    + FootnoteMarkerPositions.Count
                    + FootnoteFormattedRuns.Count
                    + HeaderFooterMarkerPositions.Count
                    + HeaderFooterFormattedRuns.Count
                    + EndnoteMarkerPositions.Count
                    + EndnoteFormattedRuns.Count);
                runs.AddRange(FormattedRuns);
                int footnoteStartCharacter = Text.Length;
                foreach (LegacyDocWritableRun run in FootnoteFormattedRuns) {
                    runs.Add(new LegacyDocWritableRun(footnoteStartCharacter + run.StartCharacter, run.Length, run.Formatting));
                }

                foreach (int markerPosition in FootnoteMarkerPositions) {
                    runs.Add(new LegacyDocWritableRun(footnoteStartCharacter + markerPosition, 1, LegacyDocWritableFormatting.SpecialCharacter));
                }

                int headerFooterStartCharacter = Text.Length + FootnoteText.Length;
                foreach (LegacyDocWritableRun run in HeaderFooterFormattedRuns) {
                    runs.Add(new LegacyDocWritableRun(headerFooterStartCharacter + run.StartCharacter, run.Length, run.Formatting));
                }

                foreach (int markerPosition in HeaderFooterMarkerPositions) {
                    runs.Add(new LegacyDocWritableRun(headerFooterStartCharacter + markerPosition, 1, LegacyDocWritableFormatting.SpecialCharacter));
                }

                int endnoteStartCharacter = Text.Length + FootnoteText.Length + HeaderFooterText.Length;
                foreach (LegacyDocWritableRun run in EndnoteFormattedRuns) {
                    runs.Add(new LegacyDocWritableRun(endnoteStartCharacter + run.StartCharacter, run.Length, run.Formatting));
                }

                foreach (int markerPosition in EndnoteMarkerPositions) {
                    runs.Add(new LegacyDocWritableRun(endnoteStartCharacter + markerPosition, 1, LegacyDocWritableFormatting.SpecialCharacter));
                }

                return runs;
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
                if (HasNoteStories || HeaderFooterFormattedParagraphs.Count > 0) {
                    return CreateFootnoteAwareParagraphSegments();
                }

                var segments = new List<LegacyDocWritableParagraphSegment>();
                int character = 0;
                foreach (LegacyDocWritableParagraph paragraph in FormattedParagraphs.OrderBy(item => item.StartCharacter)) {
                    if (paragraph.StartCharacter > character) {
                        AddParagraphSegment(segments, character, paragraph.StartCharacter - character, LegacyDocWritableParagraphFormatting.Plain);
                    }

                    AddParagraphSegment(segments, paragraph.StartCharacter, paragraph.Length, paragraph.Formatting);
                    character = paragraph.EndCharacter;
                }

                if (character < PieceTableCharacterCount) {
                    AddParagraphSegment(segments, character, PieceTableCharacterCount - character, LegacyDocWritableParagraphFormatting.Plain);
                }

                return segments;
            }

            private IReadOnlyList<LegacyDocWritableParagraphSegment> CreateFootnoteAwareParagraphSegments() {
                var segments = new List<LegacyDocWritableParagraphSegment>();
                AddBodyParagraphSegments(segments);
                AddStoryParagraphSegments(
                    segments,
                    FootnoteText,
                    Text.Length,
                    paragraph => paragraph.Length > 0 && paragraph[0] == LegacyDocFootnoteReader.FootnoteReferenceCharacter
                        ? FootnoteTextParagraphPapx
                        : PlainParagraphPapx);
                AddStoryParagraphSegments(segments, HeaderFooterText, Text.Length + FootnoteText.Length, CreateHeaderFooterParagraphFormatter());
                AddStoryParagraphSegments(
                    segments,
                    EndnoteText,
                    Text.Length + FootnoteText.Length + HeaderFooterText.Length,
                    paragraph => paragraph.Length > 0 && paragraph[0] == LegacyDocFootnoteReader.FootnoteReferenceCharacter
                        ? FootnoteTextParagraphPapx
                        : PlainParagraphPapx);
                AddRawParagraphSegment(segments, FullText.Length, PieceTableCharacterCount - FullText.Length, PlainParagraphPapx);
                return segments;
            }

            private Func<LegacyDocWritableParagraphRange, object> CreateHeaderFooterParagraphFormatter() {
                if (HeaderFooterFormattedParagraphs.Count == 0) {
                    return _ => PlainParagraphPapx;
                }

                LegacyDocWritableParagraph[] formattedParagraphs = HeaderFooterFormattedParagraphs
                    .OrderBy(item => item.StartCharacter)
                    .ToArray();
                int headerFooterStartCharacter = Text.Length + FootnoteText.Length;
                int formattedIndex = 0;
                return paragraph => {
                    while (formattedIndex < formattedParagraphs.Length
                        && headerFooterStartCharacter + formattedParagraphs[formattedIndex].EndCharacter <= paragraph.Start) {
                        formattedIndex++;
                    }

                    if (formattedIndex < formattedParagraphs.Length
                        && headerFooterStartCharacter + formattedParagraphs[formattedIndex].StartCharacter == paragraph.Start
                        && formattedParagraphs[formattedIndex].Length == paragraph.Length) {
                        return formattedParagraphs[formattedIndex].Formatting;
                    }

                    return PlainParagraphPapx;
                };
            }

            private void AddBodyParagraphSegments(List<LegacyDocWritableParagraphSegment> segments) {
                var formattedParagraphs = FormattedParagraphs
                    .OrderBy(item => item.StartCharacter)
                    .ToArray();
                int formattedIndex = 0;
                AddStoryParagraphSegments(
                    segments,
                    Text,
                    0,
                    paragraph => {
                        while (formattedIndex < formattedParagraphs.Length
                            && formattedParagraphs[formattedIndex].EndCharacter <= paragraph.Start) {
                            formattedIndex++;
                        }

                        if (formattedIndex < formattedParagraphs.Length
                            && formattedParagraphs[formattedIndex].StartCharacter == paragraph.Start
                            && formattedParagraphs[formattedIndex].Length == paragraph.Length) {
                            return formattedParagraphs[formattedIndex].Formatting;
                        }

                        return PlainParagraphPapx;
                    });
            }

            private static void AddStoryParagraphSegments(
                List<LegacyDocWritableParagraphSegment> segments,
                string story,
                int storyStart,
                Func<LegacyDocWritableParagraphRange, object> selectParagraphFormat) {
                int paragraphStart = 0;
                for (int index = 0; index < story.Length; index++) {
                    if (story[index] != '\r') {
                        continue;
                    }

                    AddStoryParagraphSegment(segments, story, storyStart, paragraphStart, index + 1, selectParagraphFormat);
                    paragraphStart = index + 1;
                }

                if (paragraphStart < story.Length) {
                    AddStoryParagraphSegment(segments, story, storyStart, paragraphStart, story.Length, selectParagraphFormat);
                }
            }

            private static void AddStoryParagraphSegment(
                List<LegacyDocWritableParagraphSegment> segments,
                string story,
                int storyStart,
                int paragraphStart,
                int paragraphEnd,
                Func<LegacyDocWritableParagraphRange, object> selectParagraphFormat) {
                int length = paragraphEnd - paragraphStart;
                if (length <= 0) {
                    return;
                }

                var paragraph = new LegacyDocWritableParagraphRange(storyStart + paragraphStart, length, story.Substring(paragraphStart, length));
                object paragraphFormat = selectParagraphFormat(paragraph);
                if (paragraphFormat is LegacyDocWritableParagraphFormatting formatting) {
                    AddParagraphSegment(segments, paragraph.Start, paragraph.Length, formatting);
                } else if (paragraphFormat is byte[] papxOverride) {
                    AddRawParagraphSegment(segments, paragraph.Start, paragraph.Length, papxOverride);
                } else {
                    throw new InvalidOperationException("The generated DOC paragraph segment formatter returned an unsupported value.");
                }
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
                    if (previous.EndCharacter == startCharacter && previous.CanMergeWith(formatting)) {
                        segments[segments.Count - 1] = previous.Extend(length);
                        return;
                    }
                }

                segments.Add(new LegacyDocWritableParagraphSegment(startCharacter, length, formatting));
            }

            private static void AddRawParagraphSegment(
                List<LegacyDocWritableParagraphSegment> segments,
                int startCharacter,
                int length,
                byte[] papxOverride) {
                if (length <= 0) {
                    return;
                }

                segments.Add(new LegacyDocWritableParagraphSegment(startCharacter, length, papxOverride));
            }
        }

        private readonly struct LegacyDocWritableParagraphRange {
            internal LegacyDocWritableParagraphRange(int start, int length, string text) {
                Start = start;
                Length = length;
                Text = text;
            }

            internal int Start { get; }

            internal int Length { get; }

            internal string Text { get; }

            internal char this[int index] => Text[index];
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
