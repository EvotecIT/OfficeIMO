namespace OfficeIMO.Word.LegacyDoc.Model {
    /// <summary>Decodes the bounded, non-complex Windows Word 6/95 FIB and its in-stream records.</summary>
    internal static class LegacyWord69Reader {
        private const ushort WesternSignature = 0xA5DC;
        private const int FibSize = 0x192;
        private const ushort HasPictures = 0x0008;

        internal static bool IsProfile(byte[] bytes) {
            if (bytes.Length < 2) return false;
            return U16(new LegacyWord69Bytes(bytes), 0) is WesternSignature or 0xA697 or 0xA698 or 0xA699;
        }

        internal static bool TryRead(byte[] source, LegacyDocImportOptions options, out Result? result,
            out string? errorCode, out string? error) {
            result = null;
            errorCode = null;
            error = null;
            if (source.Length < FibSize) {
                errorCode = "doc-word69-fib-invalid";
                error = "The Word 6/95 File Information Block is truncated.";
                return false;
            }

            var bytes = new LegacyWord69Bytes(source);
            ushort flags = U16(bytes, 0x0A);
            if ((flags & 0x0100) != 0) {
                errorCode = "doc-encrypted";
                error = "Password-protected Word 6/95 files cannot be imported.";
                return false;
            }

            if ((flags & 0x0004) != 0) {
                errorCode = "doc-word69-complex-unsupported";
                error = "Fast-saved Word 6/95 files require the legacy piece-table profile.";
                return false;
            }

            if (U16(bytes, 0) != WesternSignature || U16(bytes, 2) is not (101 or 104)
                || bytes[0x12] != 0 || U16(bytes, 0x14) != 0 || U16(bytes, 0x16) != 0
                || (flags & (0x1000 | 0x0002)) != 0 || !IsWesternLanguage(U16(bytes, 6))) {
                errorCode = "doc-word69-profile-unsupported";
                error = "The Word 6/95 source is outside the supported Windows ANSI document profile.";
                return false;
            }

            try {
                result = Read(bytes, flags, options);
                return true;
            } catch (InvalidDataException exception) {
                errorCode = "doc-word69-structure-invalid";
                error = exception.Message;
                return false;
            }
        }

        private static Result Read(LegacyWord69Bytes bytes, ushort flags, LegacyDocImportOptions options) {
            int textStart = I32(bytes, 0x18);
            int textEnd = I32(bytes, 0x1C);
            int bodyLength = I32(bytes, 0x34);
            int usedEnd = I32(bytes, 0x20);
            long storyLength = 0;
            for (int offset = 0x34; offset < 0x54; offset += 4) {
                int count = I32(bytes, offset);
                if (count < 0) throw new InvalidDataException("The Word 6/95 FIB has a negative story length.");
                storyLength += count;
            }

            if (textStart < FibSize || textEnd < textStart || usedEnd < textEnd || usedEnd > bytes.Length
                || storyLength > textEnd - textStart || storyLength > options.MaxDecodedCharacters) {
                throw new InvalidDataException("The Word 6/95 FIB text extent exceeds the stream or decoded text limit.");
            }

            if (I32(bytes, 0x164) != 0) {
                throw new InvalidDataException("The Word 6/95 FIB declares a piece table in a non-complex document.");
            }

            var allCharacters = new LegacyDocTextCharacter[(int)storyLength];
            for (int cp = 0; cp < allCharacters.Length; cp++) {
                int fc = textStart + cp;
                allCharacters[cp] = new LegacyDocTextCharacter(LegacyDocPieceTable.DecodeWindows1252(bytes[fc]), fc, cp);
            }
            var body = new ArraySegment<LegacyDocTextCharacter>(allCharacters, 0, bodyLength);

            var findings = new List<LegacyDocUnsupportedFeature>();
            void Report(LegacyDocUnsupportedFeatureKind kind, string code, string message) =>
                findings.Add(new LegacyDocUnsupportedFeature(kind, code, message));

            var properties = new LegacyWord69Properties();
            IReadOnlyList<string> fonts = LegacyWord69Properties.ReadFonts(ReadRange(bytes, usedEnd, 0xD0), out var charsets);
            var styleSheet = LegacyDocStyleSheet.ReadWord69(ReadRange(bytes, usedEnd, 0x60), fonts, properties);
            var characters = LegacyWord69Formatting.ReadCharacterRanges(bytes, usedEnd, ReadRange(bytes, usedEnd, 0xB8),
                fonts, properties, options.MaxParagraphPropertyWorkBytes, out var contextRanges);
            var paragraphs = LegacyWord69Formatting.ReadParagraphRanges(bytes, usedEnd, ReadRange(bytes, usedEnd, 0xC0),
                properties, options.MaxParagraphPropertyWorkBytes);

            if (LegacyDocField.HasUnprojectedFields(allCharacters)) {
                Report(LegacyDocUnsupportedFeatureKind.Formatting, "doc-word69-fields-unsupported",
                    "The Word 6/95 document contains fields whose instructions are reduced to cached text.");
            }
            if (LegacyDocField.HasUnresolvedStyleReferences(allCharacters, styleSheet)) {
                Report(LegacyDocUnsupportedFeatureKind.Formatting, "doc-word69-field-reference-unresolved",
                    "A Word 6/95 STYLEREF field has no matching qualified editable style.");
            }
            int equationPosition = LegacyDocField.FindEquationFieldCharacterPosition(allCharacters);
            if (equationPosition >= 0) {
                findings.Add(new LegacyDocUnsupportedFeature(LegacyDocUnsupportedFeatureKind.Formatting,
                    "doc-word69-equation-appearance-unqualified",
                    $"The Word 6/95 EQ field at character position {equationPosition} retains its source instruction and cached result, but native Word appearance is not qualified across regional settings.",
                    "WordDocument", "Field:EQ:NativeAppearance"));
            }

            if (LegacyWord69Encoding.HasUnsupportedText(bytes.Slice(textStart, allCharacters.Length), textStart,
                charsets, contextRanges, paragraphs, styleSheet, properties.StyleFonts, properties.StyleLanguages,
                properties.DefaultFontIndex, U16(bytes, 6))) {
                Report(LegacyDocUnsupportedFeatureKind.Formatting, "doc-word69-text-codepage-unsupported",
                    "The Word 6/95 document uses a text font outside the Windows-1252 import profile.");
            }

            if (properties.HasUnsupportedFormatting) {
                Report(LegacyDocUnsupportedFeatureKind.Formatting, "doc-word69-formatting-unsupported",
                    "The Word 6/95 document contains formatting properties outside the supported profile.");
            }
            if (I32(bytes, 0x38) != 0) Report(LegacyDocUnsupportedFeatureKind.Footnote, "doc-word69-footnotes", "Word 6/95 footnote stories are not projected.");
            if (I32(bytes, 0x40) != 0 || I32(bytes, 0x11C) > 2 || I32(bytes, 0x124) != 0 || I32(bytes, 0x12C) != 0) {
                Report(LegacyDocUnsupportedFeatureKind.VbaProject, "doc-word69-wordbasic", "Legacy WordBasic command and macro content is not projected.");
            }
            if (I32(bytes, 0x44) != 0) Report(LegacyDocUnsupportedFeatureKind.Comment, "doc-word69-comments", "Word 6/95 annotation stories are not projected.");
            if (I32(bytes, 0x48) != 0) Report(LegacyDocUnsupportedFeatureKind.Endnote, "doc-word69-endnotes", "Word 6/95 endnote stories are not projected.");
            if (I32(bytes, 0x4C) != 0 || I32(bytes, 0x50) != 0) {
                Report(LegacyDocUnsupportedFeatureKind.TextBox, "doc-word69-textboxes", "Word 6/95 text box stories are not projected.");
            }

            if ((flags & HasPictures) != 0 || allCharacters.Any(character => character.Character == '\u0001')) {
                Report(LegacyDocUnsupportedFeatureKind.Picture, "doc-word69-pictures", "Word 6/95 inline graphics and embedded payloads are not projected.");
            }
            if (I32(bytes, 0x104) != 0 || I32(bytes, 0x10C) != 0 || I32(bytes, 0x114) != 0) {
                Report(LegacyDocUnsupportedFeatureKind.Bookmark, "doc-word69-bookmarks", "Word 6/95 bookmark records are not projected.");
            }

            LegacyWord69Bytes documentProperties = ReadRange(bytes, usedEnd, 0x150);
            if (!documentProperties.IsEmpty && documentProperties.Length < 12) {
                throw new InvalidDataException("The Word 6/95 document properties are truncated.");
            }
            if (!documentProperties.IsEmpty && (documentProperties[7] & 0x02) != 0) {
                throw new InvalidDataException("The Word 6/95 source has edit protection outside the supported import profile.");
            }

            ushort nativeTabStop = documentProperties.IsEmpty ? (ushort)0 : U16(documentProperties, 10);
            int? defaultTabStop = nativeTabStop == 0 ? null : nativeTabStop;
            var sections = LegacyWord69Sections.Read(bytes, usedEnd, ReadRange(bytes, usedEnd, 0x88),
                bodyLength, textEnd - textStart, out var sectionHeaderMasks, out bool sectionUnsupported);
            if (sectionUnsupported) {
                Report(LegacyDocUnsupportedFeatureKind.Section, "doc-word69-sections",
                    "The Word 6/95 document contains section properties outside the supported profile.");
            }
            byte separatorMask = documentProperties.IsEmpty ? (byte)0 : documentProperties[1];
            var headerFooterStories = LegacyWord69HeaderFooter.Read(ReadRange(bytes, usedEnd, 0xB0),
                allCharacters, bodyLength + I32(bytes, 0x38), I32(bytes, 0x3C), sectionHeaderMasks,
                separatorMask, characters, paragraphs, out bool runningStoriesUnsupported);
            if (runningStoriesUnsupported) {
                Report(LegacyDocUnsupportedFeatureKind.HeaderFooter, "doc-word69-running-stories",
                    "The Word 6/95 document contains running stories outside the supported header/footer directory.");
            }
            bool differentOddAndEvenPages = !documentProperties.IsEmpty && (documentProperties[0] & 1) != 0;
            return new Result(body, characters, paragraphs, styleSheet, findings, defaultTabStop,
                sections, headerFooterStories, differentOddAndEvenPages);
        }

        internal static bool IsWesternLanguage(ushort language) =>
            (language & 0x03FF) is 0 or 3 or 6 or 7 or 9 or 10 or 11 or 12 or 15 or 16 or 19 or 20 or 22 or 29 or 45 or 54 or 56;

        private static LegacyWord69Bytes ReadRange(LegacyWord69Bytes bytes, int usedEnd, int directoryOffset) {
            int start = I32(bytes, directoryOffset);
            int length = I32(bytes, directoryOffset + 4);
            if (length == 0) return default;
            if (start < FibSize || length < 0 || start > usedEnd - length) {
                throw new InvalidDataException($"The Word 6/95 FIB record at 0x{directoryOffset:X} points outside the used stream.");
            }

            return bytes.Slice(start, length);
        }

        internal static ushort U16(LegacyWord69Bytes data, int offset) =>
            (ushort)(data[offset] | (data[offset + 1] << 8));

        internal static int I32(LegacyWord69Bytes data, int offset) =>
            data[offset] | (data[offset + 1] << 8) | (data[offset + 2] << 16) | (data[offset + 3] << 24);

        internal sealed record Result(
            IReadOnlyList<LegacyDocTextCharacter> Body,
            IReadOnlyList<LegacyDocCharacterFormatRange> Characters,
            IReadOnlyList<LegacyDocParagraphFormatRange> Paragraphs,
            LegacyDocStyleSheet StyleSheet,
            IReadOnlyList<LegacyDocUnsupportedFeature> Findings,
            int? DefaultTabStop,
            IReadOnlyList<LegacyDocSection> Sections,
            IReadOnlyList<LegacyDocHeaderFooterStory> HeaderFooterStories,
            bool DifferentOddAndEvenPages);
    }
}
