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
        private const int PapxPlcLength = 12;
        private const ushort SprmCFBold = 0x0835;
        private const ushort SprmCFItalic = 0x0836;
        private const ushort SprmCKul = 0x2A3E;
        private const ushort SprmCHps = 0x4A43;
        private const ushort SprmCRgFtc0 = 0x4A4F;
        private const ushort SprmCCv = 0x6870;
        private const ushort SprmPJc = 0x2461;
        private const ushort WordDocumentMagic = 0xA5EC;
        private const ushort Word97FibVersion = 0x00D9;
        private const ushort OneTableStreamFlag = 0x0200;
        private const int FcPlcfBtePapxOffset = 0x102;
        private const int LcbPlcfBtePapxOffset = 0x106;
        private const int FcSttbfFfnOffset = 0x112;
        private const int LcbSttbfFfnOffset = 0x116;

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
            var paragraphFormats = new List<LegacyDocWritableParagraph>();
            int paragraphCount = 0;
            foreach (OpenXmlElement child in body.ChildElements) {
                switch (child) {
                    case Paragraph paragraph:
                        AppendParagraph(text, runs, paragraphFormats, paragraph);
                        paragraphCount++;
                        break;
                    case SectionProperties:
                        break;
                    default:
                        throw new NotSupportedException($"Native DOC saving currently supports body paragraphs with bold, italic, underline, font size, color, and font family text runs. Unsupported body element: {child.LocalName}.");
                }
            }

            if (paragraphCount == 0) {
                text.Append('\r');
            }

            return new LegacyDocWritableBody(text.ToString(), runs, paragraphFormats);
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
                        throw new NotSupportedException($"Native DOC saving currently supports only text runs with bold, italic, underline, font size, color, and font family formatting. Unsupported paragraph element: {child.LocalName}.");
                }
            }

            text.Append('\r');
            if (paragraphFormatting.HasFormatting) {
                paragraphFormats.Add(new LegacyDocWritableParagraph(paragraphStart, text.Length - paragraphStart, paragraphFormatting));
            }
        }

        private static LegacyDocWritableParagraphFormatting ReadSupportedParagraphFormatting(ParagraphProperties? paragraphProperties) {
            if (paragraphProperties == null || !paragraphProperties.HasChildren) {
                return LegacyDocWritableParagraphFormatting.Plain;
            }

            byte? alignment = null;
            foreach (OpenXmlElement property in paragraphProperties.ChildElements) {
                switch (property) {
                    case Justification justification:
                        alignment = ReadSupportedParagraphAlignment(justification);
                        break;
                    default:
                        throw new NotSupportedException($"Native DOC saving currently supports only paragraph alignment. Unsupported paragraph property: {property.LocalName}.");
                }
            }

            return new LegacyDocWritableParagraphFormatting(alignment);
        }

        private static byte? ReadSupportedParagraphAlignment(Justification justification) {
            JustificationValues value = justification.Val?.Value ?? JustificationValues.Left;
            if (value == JustificationValues.Left) {
                return 0;
            } else if (value == JustificationValues.Center) {
                return 1;
            } else if (value == JustificationValues.Right) {
                return 2;
            } else if (value == JustificationValues.Both) {
                return 3;
            }

            throw new NotSupportedException($"Native DOC saving does not support paragraph alignment '{value}'.");
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
            byte? underline = null;
            int? fontSizeHalfPoints = null;
            string? colorHex = null;
            string? fontFamily = null;
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
                    case Underline underlineProperty:
                        underline = ReadSupportedUnderline(underlineProperty);
                        break;
                    case FontSize fontSize:
                        fontSizeHalfPoints = ReadFontSizeHalfPoints(fontSize);
                        break;
                    case Color color:
                        colorHex = ReadSupportedColorHex(color);
                        break;
                    case RunFonts runFonts:
                        fontFamily = ReadSupportedRunFontFamily(runFonts);
                        break;
                    default:
                        throw new NotSupportedException($"Native DOC saving currently supports only bold, italic, underline, font size, color, and font family run formatting. Unsupported run property: {property.LocalName}.");
                }
            }

            return new LegacyDocWritableFormatting(bold, italic, underline, fontSizeHalfPoints, colorHex, fontFamily);
        }

        private static bool IsEnabled(OnOffType property) {
            return property.Val == null || property.Val.Value;
        }

        private static byte? ReadSupportedUnderline(Underline underline) {
            UnderlineValues value = underline.Val?.Value ?? UnderlineValues.Single;
            if (value == UnderlineValues.None) {
                return null;
            } else if (value == UnderlineValues.Single) {
                return 1;
            } else if (value == UnderlineValues.Words) {
                return 2;
            } else if (value == UnderlineValues.Double) {
                return 3;
            } else if (value == UnderlineValues.Dotted) {
                return 4;
            } else if (value == UnderlineValues.Thick) {
                return 6;
            } else if (value == UnderlineValues.Dash) {
                return 7;
            } else if (value == UnderlineValues.DotDash) {
                return 8;
            } else if (value == UnderlineValues.DotDotDash) {
                return 9;
            } else if (value == UnderlineValues.Wave) {
                return 10;
            } else if (value == UnderlineValues.DottedHeavy) {
                return 11;
            } else if (value == UnderlineValues.DashedHeavy) {
                return 12;
            } else if (value == UnderlineValues.DashDotHeavy) {
                return 13;
            } else if (value == UnderlineValues.DashDotDotHeavy) {
                return 14;
            } else if (value == UnderlineValues.WavyHeavy) {
                return 15;
            } else if (value == UnderlineValues.DashLong) {
                return 16;
            } else if (value == UnderlineValues.WavyDouble) {
                return 17;
            } else if (value == UnderlineValues.DashLongHeavy) {
                return 18;
            }

            throw new NotSupportedException($"Native DOC saving does not support underline style '{value}'.");
        }

        private static int ReadFontSizeHalfPoints(FontSize fontSize) {
            string? value = fontSize.Val?.Value;
            if (string.IsNullOrWhiteSpace(value) || !int.TryParse(value, System.Globalization.NumberStyles.Integer, System.Globalization.CultureInfo.InvariantCulture, out int halfPoints)) {
                throw new NotSupportedException("Native DOC saving supports font size only when it is stored as a numeric half-point value.");
            }

            return halfPoints;
        }

        private static string? ReadSupportedColorHex(Color color) {
            string? value = color.Val?.Value;
            if (string.IsNullOrWhiteSpace(value) || string.Equals(value, "auto", StringComparison.OrdinalIgnoreCase)) {
                return null;
            }

            string colorValue = value!;
            string hex = colorValue.Trim().TrimStart('#').ToLowerInvariant();
            if (hex.Length != 6 || hex.Any(character => !Uri.IsHexDigit(character))) {
                throw new NotSupportedException("Native DOC saving supports text color only when it is stored as a 6-digit RGB hex value.");
            }

            return hex;
        }

        private static string? ReadSupportedRunFontFamily(RunFonts runFonts) {
            string? ascii = NormalizeFontFamily(runFonts.Ascii?.Value);
            string? highAnsi = NormalizeFontFamily(runFonts.HighAnsi?.Value);
            string? eastAsia = NormalizeFontFamily(runFonts.EastAsia?.Value);
            string? complexScript = NormalizeFontFamily(runFonts.ComplexScript?.Value);

            string? fontFamily = ascii ?? highAnsi;
            if (fontFamily == null) {
                if (eastAsia != null || complexScript != null) {
                    throw new NotSupportedException("Native DOC saving currently supports font family only for ASCII/HighAnsi text runs.");
                }

                return null;
            }

            if ((highAnsi != null && !string.Equals(fontFamily, highAnsi, StringComparison.OrdinalIgnoreCase))
                || (eastAsia != null && !string.Equals(fontFamily, eastAsia, StringComparison.OrdinalIgnoreCase))
                || (complexScript != null && !string.Equals(fontFamily, complexScript, StringComparison.OrdinalIgnoreCase))) {
                throw new NotSupportedException("Native DOC saving currently supports a single font family per text run. Multiple script-specific font families are not supported yet.");
            }

            return fontFamily;
        }

        private static string? NormalizeFontFamily(string? value) {
            if (string.IsNullOrWhiteSpace(value)) {
                return null;
            }

            return value!.Trim();
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
            byte[] fontTable = CreateFontTable(body.FontFamilies);
            int chpxFkpOffset = body.HasCharacterFormatting
                ? AlignToSector(TextOffset + textBytes.Length)
                : 0;
            int papxFkpOffset = body.HasParagraphFormatting
                ? AlignToSector(body.HasCharacterFormatting ? chpxFkpOffset + OleSectorSize : TextOffset + textBytes.Length)
                : 0;
            int streamLength = body.HasParagraphFormatting
                ? papxFkpOffset + OleSectorSize
                : body.HasCharacterFormatting
                    ? chpxFkpOffset + OleSectorSize
                    : TextOffset + textBytes.Length;
            var stream = new byte[Math.Max(FibLength, streamLength)];
            WriteUInt16(stream, 0x00, WordDocumentMagic);
            WriteUInt16(stream, 0x02, Word97FibVersion);
            WriteUInt16(stream, 0x0A, OneTableStreamFlag);
            WriteInt32(stream, 0x4C, body.Text.Length);
            WriteInt32(stream, 0xFA, body.HasCharacterFormatting ? ClxLength : 0);
            WriteInt32(stream, 0xFE, body.HasCharacterFormatting ? ChpxPlcLength : 0);
            WriteInt32(stream, FcPlcfBtePapxOffset, body.HasParagraphFormatting ? body.PapxPlcOffsetInTableStream : 0);
            WriteInt32(stream, LcbPlcfBtePapxOffset, body.HasParagraphFormatting ? PapxPlcLength : 0);
            WriteInt32(stream, FcSttbfFfnOffset, body.HasFontTable ? body.FontTableOffsetInTableStream : 0);
            WriteInt32(stream, LcbSttbfFfnOffset, fontTable.Length);
            WriteInt32(stream, 0x1A2, 0);
            WriteInt32(stream, 0x1A6, ClxLength);
            Buffer.BlockCopy(textBytes, 0, stream, TextOffset, textBytes.Length);
            if (body.HasCharacterFormatting) {
                WriteChpxFkp(stream, chpxFkpOffset, body.CreateFormattingSegments(), body.FontFamilyIndexes);
            }

            if (body.HasParagraphFormatting) {
                LegacyDocParagraphFormattingWriter.WritePapxFkp(stream, papxFkpOffset, TextOffset, OleSectorSize, body.CreateParagraphSegments());
            }

            return stream;
        }

        private static byte[] CreateTableStream(LegacyDocWritableBody body) {
            byte[] fontTable = CreateFontTable(body.FontFamilies);
            var table = new byte[ClxLength + (body.HasCharacterFormatting ? ChpxPlcLength : 0) + (body.HasParagraphFormatting ? PapxPlcLength : 0) + fontTable.Length];
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

            if (body.HasParagraphFormatting) {
                int chpxFkpOffset = body.HasCharacterFormatting
                    ? AlignToSector(TextOffset + Encoding.Unicode.GetByteCount(body.Text))
                    : 0;
                int papxFkpOffset = AlignToSector(body.HasCharacterFormatting ? chpxFkpOffset + OleSectorSize : TextOffset + Encoding.Unicode.GetByteCount(body.Text));
                WriteInt32(table, body.PapxPlcOffsetInTableStream, TextOffset);
                WriteInt32(table, body.PapxPlcOffsetInTableStream + 4, TextOffset + (body.Text.Length * 2));
                WriteInt32(table, body.PapxPlcOffsetInTableStream + 8, papxFkpOffset / OleSectorSize);
            }

            if (fontTable.Length > 0) {
                Buffer.BlockCopy(fontTable, 0, table, body.FontTableOffsetInTableStream, fontTable.Length);
            }

            return table;
        }

        private static void WriteChpxFkp(byte[] stream, int pageOffset, IReadOnlyList<LegacyDocWritableSegment> segments, IReadOnlyDictionary<string, int> fontFamilyIndexes) {
            if (segments.Count == 0 || segments.Count > byte.MaxValue) {
                throw new NotSupportedException("Native DOC saving currently supports run formatting only when it fits in one character-format page.");
            }

            int rgbOffset = pageOffset + ((segments.Count + 1) * 4);
            int chpxOffset = AlignToEven((segments.Count + 1) * 4 + segments.Count);

            for (int index = 0; index < segments.Count; index++) {
                LegacyDocWritableSegment segment = segments[index];
                WriteInt32(stream, pageOffset + (index * 4), TextOffset + (segment.StartCharacter * 2));
                if (segment.Formatting.HasFormatting) {
                    byte[] chpx = CreateChpx(segment.Formatting, fontFamilyIndexes);
                    chpxOffset = AlignToEven(chpxOffset);
                    if (chpxOffset + chpx.Length >= OleSectorSize - 1 || chpxOffset / 2 > byte.MaxValue) {
                        throw new NotSupportedException("Native DOC saving currently supports run formatting only when it fits in one character-format page.");
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

        private static byte[] CreateChpx(LegacyDocWritableFormatting formatting, IReadOnlyDictionary<string, int> fontFamilyIndexes) {
            var grpprl = new List<byte>(18);
            if (formatting.Bold) {
                AddSingleByteSprm(grpprl, SprmCFBold, 1);
            }

            if (formatting.Italic) {
                AddSingleByteSprm(grpprl, SprmCFItalic, 1);
            }

            if (formatting.Underline != null) {
                AddSingleByteSprm(grpprl, SprmCKul, formatting.Underline.Value);
            }

            if (formatting.FontSizeHalfPoints != null) {
                AddUInt16Sprm(grpprl, SprmCHps, checked((ushort)formatting.FontSizeHalfPoints.Value));
            }

            if (formatting.ColorHex != null) {
                AddColorRefSprm(grpprl, formatting.ColorHex);
            }

            if (formatting.FontFamily != null) {
                if (!fontFamilyIndexes.TryGetValue(formatting.FontFamily, out int fontIndex)) {
                    throw new InvalidOperationException("The DOC font table does not contain a formatted run font family.");
                }

                AddUInt16Sprm(grpprl, SprmCRgFtc0, checked((ushort)fontIndex));
            }

            var chpx = new byte[grpprl.Count + 1];
            chpx[0] = (byte)grpprl.Count;
            grpprl.CopyTo(chpx, 1);
            return chpx;
        }

        private static byte[] CreateFontTable(IReadOnlyList<string> fontFamilies) {
            if (fontFamilies.Count == 0) {
                return Array.Empty<byte>();
            }

            if (fontFamilies.Count > ushort.MaxValue) {
                throw new NotSupportedException("Native DOC saving supports only documents whose font table fits in a Word 97-2003 STTBF.");
            }

            using var stream = new MemoryStream();
            WriteUInt16(stream, checked((ushort)fontFamilies.Count));
            WriteUInt16(stream, 0);

            foreach (string fontFamily in fontFamilies) {
                byte[] ffn = CreateFfn(fontFamily);
                if (ffn.Length > byte.MaxValue) {
                    throw new NotSupportedException($"Native DOC saving cannot write font family '{fontFamily}' because its DOC font-table record is too long.");
                }

                stream.WriteByte(checked((byte)ffn.Length));
                stream.Write(ffn, 0, ffn.Length);
            }

            return stream.ToArray();
        }

        private static byte[] CreateFfn(string fontFamily) {
            if (string.IsNullOrWhiteSpace(fontFamily)) {
                throw new NotSupportedException("Native DOC saving cannot write an empty font family name.");
            }

            byte[] nameBytes = Encoding.Unicode.GetBytes(fontFamily + '\0');
            var ffn = new byte[39 + nameBytes.Length];
            ffn[1] = 0x90;
            ffn[2] = 0x01;
            Buffer.BlockCopy(nameBytes, 0, ffn, 39, nameBytes.Length);
            return ffn;
        }

        private static void AddSingleByteSprm(List<byte> grpprl, ushort sprm, byte operand) {
            grpprl.Add((byte)(sprm & 0xFF));
            grpprl.Add((byte)(sprm >> 8));
            grpprl.Add(operand);
        }

        private static void AddUInt16Sprm(List<byte> grpprl, ushort sprm, ushort operand) {
            grpprl.Add((byte)(sprm & 0xFF));
            grpprl.Add((byte)(sprm >> 8));
            grpprl.Add((byte)(operand & 0xFF));
            grpprl.Add((byte)(operand >> 8));
        }

        private static void AddColorRefSprm(List<byte> grpprl, string colorHex) {
            grpprl.Add((byte)(SprmCCv & 0xFF));
            grpprl.Add((byte)(SprmCCv >> 8));
            grpprl.Add(Convert.ToByte(colorHex.Substring(0, 2), 16));
            grpprl.Add(Convert.ToByte(colorHex.Substring(2, 2), 16));
            grpprl.Add(Convert.ToByte(colorHex.Substring(4, 2), 16));
            grpprl.Add(0);
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
            internal LegacyDocWritableBody(string text, IReadOnlyList<LegacyDocWritableRun> formattedRuns, IReadOnlyList<LegacyDocWritableParagraph> formattedParagraphs) {
                Text = text;
                FormattedRuns = formattedRuns;
                FormattedParagraphs = formattedParagraphs;
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

            internal IReadOnlyList<string> FontFamilies { get; }

            internal IReadOnlyDictionary<string, int> FontFamilyIndexes { get; }

            internal bool HasCharacterFormatting => FormattedRuns.Count > 0;

            internal bool HasParagraphFormatting => FormattedParagraphs.Count > 0;

            internal bool HasFontTable => FontFamilies.Count > 0;

            internal int PapxPlcOffsetInTableStream => ClxLength + (HasCharacterFormatting ? ChpxPlcLength : 0);

            internal int FontTableOffsetInTableStream => ClxLength + (HasCharacterFormatting ? ChpxPlcLength : 0) + (HasParagraphFormatting ? PapxPlcLength : 0);

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

        private readonly struct LegacyDocWritableFormatting : IEquatable<LegacyDocWritableFormatting> {
            internal static readonly LegacyDocWritableFormatting Plain = new LegacyDocWritableFormatting(false, false, null, null, null, null);

            internal LegacyDocWritableFormatting(bool bold, bool italic, byte? underline, int? fontSizeHalfPoints, string? colorHex, string? fontFamily) {
                Bold = bold;
                Italic = italic;
                Underline = underline;
                FontSizeHalfPoints = fontSizeHalfPoints;
                ColorHex = colorHex;
                FontFamily = fontFamily;
            }

            internal bool Bold { get; }

            internal bool Italic { get; }

            internal byte? Underline { get; }

            internal int? FontSizeHalfPoints { get; }

            internal string? ColorHex { get; }

            internal string? FontFamily { get; }

            internal bool HasFormatting => Bold || Italic || Underline != null || FontSizeHalfPoints != null || ColorHex != null || FontFamily != null;

            public bool Equals(LegacyDocWritableFormatting other) {
                return Bold == other.Bold
                    && Italic == other.Italic
                    && Underline == other.Underline
                    && FontSizeHalfPoints == other.FontSizeHalfPoints
                    && string.Equals(ColorHex, other.ColorHex, StringComparison.OrdinalIgnoreCase)
                    && string.Equals(FontFamily, other.FontFamily, StringComparison.OrdinalIgnoreCase);
            }

            public override bool Equals(object? obj) {
                return obj is LegacyDocWritableFormatting other && Equals(other);
            }

            public override int GetHashCode() {
                int hash = 17;
                hash = (hash * 31) + Bold.GetHashCode();
                hash = (hash * 31) + Italic.GetHashCode();
                hash = (hash * 31) + Underline.GetHashCode();
                hash = (hash * 31) + FontSizeHalfPoints.GetHashCode();
                hash = (hash * 31) + StringComparer.OrdinalIgnoreCase.GetHashCode(ColorHex ?? string.Empty);
                hash = (hash * 31) + StringComparer.OrdinalIgnoreCase.GetHashCode(FontFamily ?? string.Empty);
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
