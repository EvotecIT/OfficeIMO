using OfficeIMO.Excel.Xlsb.Biff12;

namespace OfficeIMO.Excel.Xlsb.Write {
    /// <summary>
    /// Replaces one kind of worksheet record collection after the cell table, such as DVALS or RWBRK, when a loaded XLSB
    /// worksheet is rewritten. Every existing collection of that kind is removed; the replacement records are written where
    /// the first one was, or at the collection's place in the MS-XLSB WORKSHEET grammar (section 2.1.7.62) when there was none.
    /// </summary>
    internal sealed class XlsbRecordBlockEdit {
        private const int BrtEndSheetData = 146;
        private const int BrtEndSheet = 130;

        // Record types that open each optional collection of the WORKSHEET rule after CELLTABLE, in grammar order.
        private static readonly int[] GrammarOrder = {
            663, // BrtSheetCalcProp
            678, // BrtSheetProtectionIso
            535, // BrtSheetProtection
            680, // BrtRangeProtectionIso
            536, // BrtRangeProtection
            500, // BrtBeginScenMan
            161, // BrtBeginAFilter
            530, // BrtBeginSortState
            495, // BrtBeginDCon
            422, // BrtBeginUserShViews
            177, // BrtBeginMergeCells
            537, // BrtPhoneticInfo
            461, // BrtBeginConditionalFormatting
            573, // BrtBeginDVals
            494, // BrtHLink
            477, // BrtPrintOptions
            476, // BrtMargins
            478, // BrtPageSetup
            479, // BrtBeginHeaderFooter
            392, // BrtBeginRwBrk
            394, // BrtBeginColBrk
            625, // BrtBigName
            605, // BrtBeginCellWatches
            648, // BrtBeginCellIgnoreECs
            594, // BrtBeginSmartTags
            550, // BrtDrawing
            551, // BrtLegacyDrawing
            552, // BrtLegacyDrawingHF
            562, // BrtBkHim
            638, // BrtBeginOleObjects
            643, // BrtBeginActiveXControls
            554, // BrtBeginWebPubItems
            660, // BrtBeginListParts
            35,  // BrtFRTBegin (FRTWORKSHEET)
        };

        // Closing records of the collections above that span more than one record.
        private static readonly Dictionary<int, int> CollectionEnds = new Dictionary<int, int> {
            [500] = 501, [161] = 162, [530] = 531, [495] = 496, [422] = 425, [177] = 178, [461] = 462, [573] = 574,
            [479] = 480, [392] = 393, [394] = 395, [605] = 606, [648] = 650, [594] = 595, [638] = 640, [643] = 645,
            [554] = 555, [660] = 662, [35] = 36, [BrtACBegin] = BrtACEnd,
        };

        private const int BrtACBegin = 37;
        private const int BrtACEnd = 38;

        /// <param name="beginType">The record that opens the collection, or the single record for a one-record element.</param>
        /// <param name="endType">The record that closes the collection; equal to <paramref name="beginType"/> for one record.</param>
        /// <param name="replacement">The records to write instead; empty removes the collection.</param>
        internal XlsbRecordBlockEdit(int beginType, int endType, IReadOnlyList<XlsbGeneratedRecord> replacement) {
            if (Array.IndexOf(GrammarOrder, beginType) < 0) {
                throw new ArgumentOutOfRangeException(nameof(beginType), beginType, "The record does not open a worksheet collection after the cell table.");
            }
            if (CollectionEnds.TryGetValue(beginType, out int collectionEnd) ? collectionEnd != endType : endType != beginType) {
                throw new ArgumentOutOfRangeException(nameof(endType), endType, "The record does not close the worksheet collection.");
            }
            BeginType = beginType;
            EndType = endType;
            Replacement = replacement ?? throw new ArgumentNullException(nameof(replacement));
        }

        internal int BeginType { get; }

        internal int EndType { get; }

        internal IReadOnlyList<XlsbGeneratedRecord> Replacement { get; }

        /// <summary>Applies <paramref name="edits"/> to a complete worksheet part and returns the new part.</summary>
        internal static byte[] Apply(byte[] part, IReadOnlyList<XlsbRecordBlockEdit> edits, long maximumBytes) {
            if (edits.Count == 0) return part;
            List<XlsbRecord> records;
            using (var input = new MemoryStream(part, writable: false)) {
                records = XlsbRecordReader.ReadAll(input).ToList();
            }
            var output = new List<(int Type, byte[] Payload)>(records.Count);
            int tableEnd = records.FindIndex(record => record.Type == BrtEndSheetData);
            if (tableEnd < 0) throw new InvalidDataException("The XLSB worksheet does not contain a BrtEndSheetData record.");
            for (int index = 0; index <= tableEnd; index++) output.Add((records[index].Type, records[index].Data));

            var pending = new List<XlsbRecordBlockEdit>(edits);
            var written = new HashSet<XlsbRecordBlockEdit>();
            for (int index = tableEnd + 1; index < records.Count; index++) {
                // An ACUID prefix (BrtACBegin .. BrtACEnd) belongs to the record after it, as in [ACUID] BrtHLink, so the
                // prefix and that record form one unit and new collections are inserted before the prefix.
                int keyIndex = index;
                if (records[index].Type == BrtACBegin) keyIndex = FindCollectionEnd(records, index) + 1;
                if (keyIndex >= records.Count) throw new InvalidDataException("The XLSB worksheet ends inside an ACUID prefix.");
                XlsbRecord record = records[keyIndex];
                int end = FindCollectionEnd(records, keyIndex);
                // Write edits whose collection belongs before this top-level record and was not present, in grammar order.
                int rank = Rank(record.Type);
                if (rank >= 0 || record.Type == BrtEndSheet) {
                    foreach (XlsbRecordBlockEdit edit in pending.Where(edit => record.Type == BrtEndSheet || Rank(edit.BeginType) < rank)
                        .OrderBy(edit => Rank(edit.BeginType)).ToArray()) {
                        if (written.Add(edit)) Append(output, edit.Replacement);
                        pending.Remove(edit);
                    }
                }

                XlsbRecordBlockEdit? owner = edits.FirstOrDefault(edit => edit.BeginType == record.Type);
                if (owner == null) {
                    for (int item = index; item <= end; item++) output.Add((records[item].Type, records[item].Data));
                } else {
                    if (written.Add(owner)) Append(output, owner.Replacement);
                    pending.Remove(owner);
                }
                index = end;
            }

            using var stream = new OfficeIMO.Core.Internal.OfficeBoundedMemoryStream(maximumBytes);
            foreach ((int type, byte[] payload) in output) XlsbRecordWriter.Write(stream, type, payload);
            return stream.ToArray();
        }

        private static int Rank(int type) => Array.IndexOf(GrammarOrder, type);

        // The last record of the top-level collection that starts at index, counting nested collections of the same kind.
        private static int FindCollectionEnd(IReadOnlyList<XlsbRecord> records, int index) {
            int begin = records[index].Type;
            if (!CollectionEnds.TryGetValue(begin, out int endType)) return index;
            int depth = 0;
            for (int item = index; item < records.Count; item++) {
                if (records[item].Type == begin) depth++;
                else if (records[item].Type == endType && --depth == 0) return item;
            }
            throw new InvalidDataException($"The XLSB worksheet collection opened by record {begin} has no closing record {endType}.");
        }

        private static void Append(List<(int Type, byte[] Payload)> output, IReadOnlyList<XlsbGeneratedRecord> records) {
            foreach (XlsbGeneratedRecord record in records) output.Add((record.Type, record.Payload));
        }
    }
}
