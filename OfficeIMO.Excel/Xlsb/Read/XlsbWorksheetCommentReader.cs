using OfficeIMO.Excel.Xlsb.Biff12;
using OfficeIMO.Excel.Xlsb.Model;

namespace OfficeIMO.Excel.Xlsb.Read {
    /// <summary>Relationship and content types of the legacy comment parts (MS-XLSB 2.1.7.8, ISO/IEC 29500-1 18.1).</summary>
    internal static class XlsbCommentRelationships {
        internal const string Comments = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/comments";
        internal const string VmlDrawing = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/vmlDrawing";
        internal const string CommentsContentType = "application/vnd.ms-excel.comments";
        internal const string VmlDrawingContentType = "application/vnd.openxmlformats-officedocument.vmlDrawing";
    }

    /// <summary>Decodes the COMMENTS record sequence of a Comments part (MS-XLSB 2.1.7.8).</summary>
    internal static class XlsbWorksheetCommentReader {
        internal const int BrtBeginComments = 628;
        internal const int BrtEndComments = 629;
        internal const int BrtBeginCommentAuthors = 630;
        internal const int BrtEndCommentAuthors = 631;
        internal const int BrtCommentAuthor = 632;
        internal const int BrtBeginCommentList = 633;
        internal const int BrtEndCommentList = 634;
        internal const int BrtBeginComment = 635;
        internal const int BrtEndComment = 636;
        internal const int BrtCommentText = 637;
        internal const int BrtFRTBegin = 35;
        internal const int BrtFRTEnd = 36;
        internal const int BrtACBegin = 37;
        internal const int BrtACEnd = 38;
        internal const int MaximumAuthorLength = 54;
        internal const int MaximumTextLength = 0x7FFF;

        /// <summary>
        /// COMMENTS = BrtBeginComments COMMENTAUTHORS COMMENTLIST *FRT BrtEndComments, with
        /// COMMENT = [ACUID] BrtBeginComment [BrtCommentText] BrtEndComment.
        /// </summary>
        internal static XlsbComments Read(IReadOnlyList<XlsbRecord> records, string partName, string relationshipId) {
            int index = 0;
            XlsbRecord Next() {
                if (index >= records.Count) throw Invalid(partName, "ends before BrtEndComments");
                return records[index++];
            }
            void Expect(int type, string name) {
                XlsbRecord record = Next();
                if (record.Type != type) throw Invalid(partName, $"has record {record.Type} where {name} is required");
                if (record.Data.Length != 0) throw Invalid(partName, $"has a {name} record with a payload");
            }

            Expect(BrtBeginComments, "BrtBeginComments");
            Expect(BrtBeginCommentAuthors, "BrtBeginCommentAuthors");
            var authors = new List<string>();
            while (index < records.Count && records[index].Type == BrtCommentAuthor) {
                var cursor = new XlsbBinaryCursor(records[index++].Data);
                string author = cursor.ReadWideString(MaximumAuthorLength);
                if (cursor.Remaining != 0) throw Invalid(partName, "has an unexpected BrtCommentAuthor payload");
                authors.Add(author);
            }
            Expect(BrtEndCommentAuthors, "BrtEndCommentAuthors");
            Expect(BrtBeginCommentList, "BrtBeginCommentList");

            var comments = new List<XlsbComment>();
            var cells = new HashSet<(int Row, int Column)>();
            while (index < records.Count && records[index].Type != BrtEndCommentList) {
                int start = index;
                if (records[index].Type == BrtACBegin) SkipBlock(records, ref index, BrtACBegin, BrtACEnd, partName);
                XlsbRecord begin = Next();
                if (begin.Type != BrtBeginComment) throw Invalid(partName, $"has record {begin.Type} where BrtBeginComment is required");
                (int row, int column, int author) = ReadBeginComment(begin.Data, authors.Count, partName);
                if (!cells.Add((row, column))) throw Invalid(partName, $"has more than one comment on row {row}, column {column}");
                string text = string.Empty;
                XlsbTextRun[]? runs = null;
                if (index < records.Count && records[index].Type == BrtCommentText) {
                    (text, runs) = ReadCommentText(records[index++].Data, partName);
                }
                Expect(BrtEndComment, "BrtEndComment");
                comments.Add(new XlsbComment(row, column, author, text, runs, Encode(records, start, index)));
            }
            Expect(BrtEndCommentList, "BrtEndCommentList");
            int trailingStart = index;
            while (index < records.Count && records[index].Type == BrtFRTBegin) SkipBlock(records, ref index, BrtFRTBegin, BrtFRTEnd, partName);
            byte[] trailing = Encode(records, trailingStart, index);
            Expect(BrtEndComments, "BrtEndComments");
            if (index != records.Count) throw Invalid(partName, "has records after BrtEndComments");
            return new XlsbComments(partName, relationshipId, authors.AsReadOnly(), comments.AsReadOnly(), trailing);
        }

        // BrtBeginComment (2.4.30): iauthor, an UncheckedRfX naming one cell, and a GUID that only shared workbooks use.
        private static (int Row, int Column, int Author) ReadBeginComment(byte[] payload, int authorCount, string partName) {
            if (payload.Length != 36) throw Invalid(partName, "has a BrtBeginComment record of the wrong size");
            int author = BitConverter.ToInt32(payload, 0);
            uint firstRow = BitConverter.ToUInt32(payload, 4);
            uint lastRow = BitConverter.ToUInt32(payload, 8);
            uint firstColumn = BitConverter.ToUInt32(payload, 12);
            uint lastColumn = BitConverter.ToUInt32(payload, 16);
            if (author < 0 || author >= authorCount) throw Invalid(partName, "has a comment whose author index is out of range");
            if (firstRow != lastRow || firstColumn != lastColumn || firstRow >= 1_048_576U || firstColumn >= 16_384U) {
                throw Invalid(partName, "has a comment that does not name one worksheet cell");
            }
            return (checked((int)firstRow + 1), checked((int)firstColumn + 1), author);
        }

        // BrtCommentText (2.4.341): a RichStr with fRichStr = 1 and fExtStr = 0.
        private static (string Text, XlsbTextRun[]? Runs) ReadCommentText(byte[] payload, string partName) {
            var cursor = new XlsbBinaryCursor(payload);
            byte flags = cursor.ReadByte();
            if ((flags & 0x01) == 0 || (flags & 0x02) != 0) throw Invalid(partName, "has comment text that is not a RichStr without phonetic data");
            string text = cursor.ReadWideString(MaximumTextLength);
            uint count = cursor.ReadUInt32();
            if (count > MaximumTextLength || cursor.Remaining != count * 4U) throw Invalid(partName, "has malformed comment text runs");
            if (count == 0) return (text, null);
            var runs = new XlsbTextRun[count];
            for (int run = 0; run < runs.Length; run++) {
                ushort start = cursor.ReadUInt16();
                ushort font = cursor.ReadUInt16();
                if (start >= text.Length || (run > 0 && start <= runs[run - 1].Start)) throw Invalid(partName, "has malformed comment text runs");
                runs[run] = new XlsbTextRun(start, font);
            }
            return (text, runs);
        }

        // Skips one block from begin through its matching end record; FRT blocks may nest.
        private static void SkipBlock(IReadOnlyList<XlsbRecord> records, ref int index, int beginType, int endType, string partName) {
            int depth = 0;
            do {
                if (index >= records.Count) throw Invalid(partName, "has an unterminated future-record block");
                int type = records[index++].Type;
                if (type == beginType) depth++;
                else if (type == endType) depth--;
            } while (depth > 0);
        }

        private static byte[] Encode(IReadOnlyList<XlsbRecord> records, int start, int end) {
            if (start == end) return Array.Empty<byte>();
            using var output = new MemoryStream();
            for (int index = start; index < end; index++) XlsbRecordWriter.Write(output, records[index].Type, records[index].Data);
            return output.ToArray();
        }

        private static InvalidDataException Invalid(string partName, string detail) =>
            new InvalidDataException($"The XLSB comments part '{partName}' {detail}.");
    }
}
