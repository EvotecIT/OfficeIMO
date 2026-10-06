using System.Text;

namespace OfficeIMO.Word.LegacyDoc.Write {
    internal static partial class LegacyDocWriter {
        private static byte[] CreateWordDocumentStream(LegacyDocWritableBody body, bool isTemplate) {
            bool compressedText = CanWriteCompressedText(body.StoredText);
            int bytesPerCharacter = compressedText ? 1 : 2;
            byte[] textBytes = compressedText ? EncodeCompressedText(body.StoredText) : Encoding.Unicode.GetBytes(body.StoredText);
            byte[] fontTable = CreateFontTable(body.FontFamilies);
            IReadOnlyList<IReadOnlyList<LegacyDocWritableSegment>> chpxPages = body.ChpxPages;
            int chpxFkpOffset = body.HasCharacterFormatting
                ? AlignToSector(TextOffset + textBytes.Length)
                : 0;
            IReadOnlyList<IReadOnlyList<LegacyDocWritableParagraphSegment>> papxPages = body.PapxPages;
            int papxFkpOffset = body.HasParagraphFormatting
                ? AlignToSector(body.HasCharacterFormatting ? chpxFkpOffset + (chpxPages.Count * OleSectorSize) : TextOffset + textBytes.Length)
                : 0;
            int sectionDataOffset = GetSectionDataOffset(body, textBytes.Length, chpxFkpOffset, chpxPages.Count, papxFkpOffset, papxPages.Count);
            IReadOnlyList<LegacyDocWritableSectionRecord> sectionRecords = CreateSectionRecords(body, sectionDataOffset);
            int streamLength = body.HasParagraphFormatting
                ? papxFkpOffset + (papxPages.Count * OleSectorSize)
                : body.HasCharacterFormatting
                    ? chpxFkpOffset + (chpxPages.Count * OleSectorSize)
                    : TextOffset + textBytes.Length;
            if (body.HasSectionDescriptors) {
                streamLength = Math.Max(streamLength, sectionRecords.Count == 0 ? sectionDataOffset : sectionRecords.Max(record => record.EndOffset));
            }
            var stream = new byte[Math.Max(FibLength, streamLength)];
            WriteUInt16(stream, 0x00, WordDocumentMagic);
            bool hasNestedTables = papxPages.Any(page => page.Any(segment => segment.Formatting.TableDepth > 1));
            // Word 97 interprets nested cell marks as ordinary paragraph marks.
            // Declare the Word 2000 format and its matching FIB extension when needed.
            ushort fibVersion = hasNestedTables ? (ushort)0x00D9 : Word97FibVersion;
            ushort fibPairCount = hasNestedTables ? (ushort)0x006C : (ushort)0x005D;
            WriteUInt16(stream, 0x02, fibVersion);
            WriteUInt16(stream, 0x06, DefaultLanguageId);
            ushort fibFlags = DefaultFibFlags;
            if (body.HasPictures) fibFlags = unchecked((ushort)(fibFlags | HasPicturesFibFlag));
            if (isTemplate) fibFlags = unchecked((ushort)(fibFlags | TemplateFibFlag));
            WriteUInt16(stream, 0x0A, fibFlags);
            WriteUInt16(stream, 0x0C, Word97FibBackVersion);
            WriteInt32(stream, 0x18, TextOffset);
            WriteInt32(stream, 0x1C, TextOffset + textBytes.Length);
            WriteUInt16(stream, 0x20, FibRgW97WordCount);
            WriteUInt16(stream, 0x3C, DefaultLanguageId);
            WriteUInt16(stream, 0x3E, FibRgLw97DwordCount);
            WriteInt32(stream, 0x4C, body.Text.Length);
            WriteInt32(stream, 0x50, body.FootnoteText.Length);
            WriteInt32(stream, 0x54, body.HeaderFooterText.Length);
            WriteInt32(stream, 0x5C, body.CommentText.Length);
            WriteInt32(stream, 0x60, body.EndnoteText.Length);
            WriteUInt16(stream, 0x98, fibPairCount);
            int fibExtensionOffset = 0x9A + fibPairCount * 8;
            WriteUInt16(stream, fibExtensionOffset, hasNestedTables ? (ushort)2 : (ushort)0);
            if (hasNestedTables) {
                WriteUInt16(stream, fibExtensionOffset + 2, fibVersion);
                WriteUInt16(stream, fibExtensionOffset + 4, 0);
            }
            WriteInt32(stream, FcStshfOffset, body.HasStyleSheet ? body.StyleSheetOffsetInTableStream : 0);
            WriteInt32(stream, LcbStshfOffset, body.StyleSheet.Bytes.Length);
            WriteInt32(stream, 0xFA, body.HasCharacterFormatting ? ClxLength : 0);
            WriteInt32(stream, 0xFE, body.HasCharacterFormatting ? body.ChpxPlcLength : 0);
            WriteInt32(stream, FcPlcfBtePapxOffset, body.HasParagraphFormatting ? body.PapxPlcOffsetInTableStream : 0);
            WriteInt32(stream, LcbPlcfBtePapxOffset, body.HasParagraphFormatting ? body.PapxPlcLength : 0);
            WriteInt32(stream, FcPlcfSedOffset, body.HasSectionDescriptors ? body.SedPlcOffsetInTableStream : 0);
            WriteInt32(stream, LcbPlcfSedOffset, body.HasSectionDescriptors ? body.SedPlcLength : 0);
            WriteInt32(stream, FcPlcffndRefOffset, body.HasFootnotes ? body.PlcffndRefOffsetInTableStream : 0);
            WriteInt32(stream, LcbPlcffndRefOffset, body.HasFootnotes ? body.PlcffndRef.Length : 0);
            WriteInt32(stream, FcPlcffndTxtOffset, body.HasFootnotes ? body.PlcffndTxtOffsetInTableStream : 0);
            WriteInt32(stream, LcbPlcffndTxtOffset, body.HasFootnotes ? body.PlcffndTxt.Length : 0);
            WriteInt32(stream, FcPlcfandRefOffset, body.HasComments ? body.PlcfandRefOffsetInTableStream : 0);
            WriteInt32(stream, LcbPlcfandRefOffset, body.HasComments ? body.PlcfandRef.Length : 0);
            WriteInt32(stream, FcPlcfandTxtOffset, body.HasComments ? body.PlcfandTxtOffsetInTableStream : 0);
            WriteInt32(stream, LcbPlcfandTxtOffset, body.HasComments ? body.PlcfandTxt.Length : 0);
            WriteInt32(stream, FcPlcfendRefOffset, body.HasEndnotes ? body.PlcfendRefOffsetInTableStream : 0);
            WriteInt32(stream, LcbPlcfendRefOffset, body.HasEndnotes ? body.PlcfendRef.Length : 0);
            WriteInt32(stream, FcPlcfendTxtOffset, body.HasEndnotes ? body.PlcfendTxtOffsetInTableStream : 0);
            WriteInt32(stream, LcbPlcfendTxtOffset, body.HasEndnotes ? body.PlcfendTxt.Length : 0);
            WriteInt32(stream, FcPlcfHddOffset, body.HasHeaderFooterStories ? body.PlcfHddOffsetInTableStream : 0);
            WriteInt32(stream, LcbPlcfHddOffset, body.HasHeaderFooterStories ? body.PlcfHdd.Length : 0);
            body.FieldTables.WriteFibRecords(stream, body.FieldTablesOffsetInTableStream);
            WriteInt32(stream, FcSttbfBkmkOffset, body.HasBookmarks ? body.SttbfBkmkOffsetInTableStream : 0);
            WriteInt32(stream, LcbSttbfBkmkOffset, body.HasBookmarks ? body.SttbfBkmk.Length : 0);
            WriteInt32(stream, FcSttbfRMarkOffset, body.HasRevisions ? body.SttbfRMarkOffsetInTableStream : 0);
            WriteInt32(stream, LcbSttbfRMarkOffset, body.SttbfRMark.Length);
            WriteInt32(stream, FcPlcfBkfOffset, body.HasBookmarks ? body.PlcfBkfOffsetInTableStream : 0);
            WriteInt32(stream, LcbPlcfBkfOffset, body.HasBookmarks ? body.PlcfBkf.Length : 0);
            WriteInt32(stream, FcPlcfBklOffset, body.HasBookmarks ? body.PlcfBklOffsetInTableStream : 0);
            WriteInt32(stream, LcbPlcfBklOffset, body.HasBookmarks ? body.PlcfBkl.Length : 0);
            WriteInt32(stream, FcSttbfFfnOffset, body.HasFontTable ? body.FontTableOffsetInTableStream : 0);
            WriteInt32(stream, LcbSttbfFfnOffset, fontTable.Length);
            WriteInt32(stream, FcDopOffset, body.HasDocumentOptions ? body.DopOffsetInTableStream : 0);
            WriteInt32(stream, LcbDopOffset, body.HasDocumentOptions ? body.DopLength : 0);
            WriteInt32(stream, 0x1A2, 0);
            WriteInt32(stream, 0x1A6, ClxLength);
            Buffer.BlockCopy(textBytes, 0, stream, TextOffset, textBytes.Length);
            if (body.HasCharacterFormatting) {
                for (int pageIndex = 0; pageIndex < chpxPages.Count; pageIndex++) {
                    WriteChpxFkp(
                        stream,
                        chpxFkpOffset + (pageIndex * OleSectorSize),
                        chpxPages[pageIndex],
                        body.FontFamilyIndexes,
                        body.RevisionAuthorIndexes,
                        bytesPerCharacter);
                }
            }

            if (body.HasParagraphFormatting) {
                for (int pageIndex = 0; pageIndex < papxPages.Count; pageIndex++) {
                    LegacyDocParagraphFormattingWriter.WritePapxFkp(
                        stream,
                        papxFkpOffset + (pageIndex * OleSectorSize),
                        TextOffset,
                        OleSectorSize,
                        papxPages[pageIndex],
                        bytesPerCharacter);
                }
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
                WriteChpxBtePlc(table, body, chpxFkpOffset, bytesPerCharacter);
            }

            if (body.HasParagraphFormatting) {
                int chpxFkpOffset = body.HasCharacterFormatting
                    ? AlignToSector(TextOffset + textByteLength)
                    : 0;
                int papxFkpOffset = AlignToSector(body.HasCharacterFormatting ? chpxFkpOffset + (body.ChpxPageCount * OleSectorSize) : TextOffset + textByteLength);
                WritePapxBtePlc(table, body, papxFkpOffset, bytesPerCharacter);
            }

            if (body.HasSectionDescriptors) {
                int chpxFkpOffset = body.HasCharacterFormatting
                    ? AlignToSector(TextOffset + textByteLength)
                    : 0;
                int papxFkpOffset = body.HasParagraphFormatting
                    ? AlignToSector(body.HasCharacterFormatting ? chpxFkpOffset + (body.ChpxPageCount * OleSectorSize) : TextOffset + textByteLength)
                    : 0;
                int sepxOffset = AlignToEven(body.HasParagraphFormatting
                    ? papxFkpOffset + (body.PapxPageCount * OleSectorSize)
                    : body.HasCharacterFormatting
                        ? chpxFkpOffset + (body.ChpxPageCount * OleSectorSize)
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

            if (body.HasComments) {
                Buffer.BlockCopy(body.PlcfandRef, 0, table, body.PlcfandRefOffsetInTableStream, body.PlcfandRef.Length);
                Buffer.BlockCopy(body.PlcfandTxt, 0, table, body.PlcfandTxtOffsetInTableStream, body.PlcfandTxt.Length);
            }

            if (body.HasEndnotes) {
                Buffer.BlockCopy(body.PlcfendRef, 0, table, body.PlcfendRefOffsetInTableStream, body.PlcfendRef.Length);
                Buffer.BlockCopy(body.PlcfendTxt, 0, table, body.PlcfendTxtOffsetInTableStream, body.PlcfendTxt.Length);
            }

            if (body.HasDocumentOptions) {
                byte[] dop = CreateDopBase(body);
                Buffer.BlockCopy(dop, 0, table, body.DopOffsetInTableStream, dop.Length);
            }

            if (body.HasBookmarks) {
                Buffer.BlockCopy(body.SttbfBkmk, 0, table, body.SttbfBkmkOffsetInTableStream, body.SttbfBkmk.Length);
                Buffer.BlockCopy(body.PlcfBkf, 0, table, body.PlcfBkfOffsetInTableStream, body.PlcfBkf.Length);
                Buffer.BlockCopy(body.PlcfBkl, 0, table, body.PlcfBklOffsetInTableStream, body.PlcfBkl.Length);
            }

            if (body.HasRevisions) {
                Buffer.BlockCopy(body.SttbfRMark, 0, table, body.SttbfRMarkOffsetInTableStream, body.SttbfRMark.Length);
            }

            if (body.HasStyleSheet) {
                Buffer.BlockCopy(body.StyleSheet.Bytes, 0, table, body.StyleSheetOffsetInTableStream, body.StyleSheet.Bytes.Length);
            }

            if (fontTable.Length > 0) {
                Buffer.BlockCopy(fontTable, 0, table, body.FontTableOffsetInTableStream, fontTable.Length);
            }

            body.FieldTables.WriteTableBytes(table, body.FieldTablesOffsetInTableStream);

            return table;
        }
    }
}
