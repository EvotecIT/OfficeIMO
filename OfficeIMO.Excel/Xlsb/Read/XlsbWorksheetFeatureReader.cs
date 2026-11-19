using OfficeIMO.Excel.Xlsb.Biff12;
using OfficeIMO.Excel.Xlsb.Model;

namespace OfficeIMO.Excel.Xlsb.Read {
    /// <summary>
    /// Decodes worksheet record collections after the cell table that project directly to SpreadsheetML elements, and
    /// adds those elements to <see cref="XlsbSheet.FeatureElements"/>.
    /// </summary>
    internal sealed class XlsbWorksheetFeatureReader {
        private readonly XlsbWorksheetPageBreakReader _pageBreaks;
        private readonly XlsbWorksheetIgnoredErrorReader _ignoredErrors;
        private readonly XlsbWorksheetConditionalFormattingReader _conditionalFormatting;
        private readonly XlsbWorksheetDataValidationReader _dataValidations;

        internal XlsbWorksheetFeatureReader(XlsbSheet sheet, XlsbWorkbook workbook, int sheetIndex, string partName) {
            Sheet = sheet;
            Workbook = workbook;
            SheetIndex = sheetIndex;
            PartName = partName;
            _pageBreaks = new XlsbWorksheetPageBreakReader(this);
            _ignoredErrors = new XlsbWorksheetIgnoredErrorReader(this);
            _conditionalFormatting = new XlsbWorksheetConditionalFormattingReader(this);
            _dataValidations = new XlsbWorksheetDataValidationReader(this);
        }

        internal XlsbSheet Sheet { get; }

        internal XlsbWorkbook Workbook { get; }

        internal int SheetIndex { get; }

        internal string PartName { get; }

        /// <summary>Consumes <paramref name="record"/> when it belongs to a decoded feature collection.</summary>
        internal bool TryRead(XlsbRecord record) =>
            _conditionalFormatting.TryRead(record)
            || _dataValidations.TryRead(record)
            || _pageBreaks.TryRead(record)
            || _ignoredErrors.TryRead(record)
            || TryReadCommentDrawing(record);

        /// <summary>
        /// Consumes BrtLegacyDrawing (MS-XLSB 2.4.703) when the worksheet has a Comments part and the record names its
        /// VML drawing; the comments projection then owns both parts. Other legacy drawings stay preserved records.
        /// </summary>
        private bool TryReadCommentDrawing(XlsbRecord record) {
            const int BrtLegacyDrawing = 551;
            if (record.Type != BrtLegacyDrawing) return false;
            if (!Sheet.Relationships.Values.Any(item => item.Type == XlsbCommentRelationships.Comments && !item.IsExternal)) return false;
            var cursor = new XlsbBinaryCursor(record.Data);
            string id = cursor.ReadWideString(255);
            if (cursor.Remaining != 0) throw new InvalidDataException($"BrtLegacyDrawing in '{PartName}' has an unexpected payload.");
            if (!Sheet.Relationships.TryGetValue(id, out var relationship) || relationship.IsExternal
                || relationship.Type != XlsbCommentRelationships.VmlDrawing) return false;
            if (Sheet.LegacyDrawingRelationshipId != null) throw new InvalidDataException($"The worksheet part '{PartName}' contains more than one BrtLegacyDrawing record.");
            Sheet.LegacyDrawingRelationshipId = id;
            return true;
        }

        /// <summary>Called after the last worksheet record; throws when a collection is still open.</summary>
        internal void Complete() {
            _conditionalFormatting.Complete();
            _pageBreaks.Complete();
            _ignoredErrors.Complete();
            _dataValidations.Complete();
        }

        /// <summary>
        /// Decodes the cce/rgce/cb/rgcb layout of a DVParsedFormula, CFParsedFormula or similar structure at
        /// <paramref name="offset"/> and advances past it. Returns null when the tokens cannot be shown as text.
        /// </summary>
        /// <remarks>Pass the zero-based <paramref name="relativeAnchor"/> cell when the structure stores relative references as
        /// PtgRefN/PtgAreaN offsets from it, as data validation and conditional formatting formulas do.</remarks>
        internal string? ReadFormula(byte[] payload, ref int offset, (int Row, int Column)? relativeAnchor = null) {
            if (offset > payload.Length - 4) throw new InvalidDataException($"A formula in '{PartName}' is truncated.");
            int cce = checked((int)BitConverter.ToUInt32(payload, offset));
            offset += 4;
            if (cce < 0 || offset > payload.Length - cce - 4) throw new InvalidDataException($"A formula in '{PartName}' is truncated.");
            byte[] tokens = new byte[cce];
            Buffer.BlockCopy(payload, offset, tokens, 0, cce);
            offset += cce;
            int cb = checked((int)BitConverter.ToUInt32(payload, offset));
            offset += 4;
            if (cb < 0 || offset > payload.Length - cb) throw new InvalidDataException($"A formula in '{PartName}' is truncated.");
            byte[] extra = new byte[cb];
            Buffer.BlockCopy(payload, offset, extra, 0, cb);
            offset += cb;
            if (cce == 0) return string.Empty;
            if (relativeAnchor.HasValue) {
                try {
                    XlsbRelativeReferenceTokens.ToCoordinates(tokens, relativeAnchor.Value.Row, relativeAnchor.Value.Column);
                } catch (NotSupportedException) {
                    return null;
                }
            }
            return XlsbFormulaTextReader.TryRead(tokens, out string? text, Workbook, SheetIndex, extra) ? text : null;
        }

        internal static XlsbCellRange ReadRfX(byte[] payload, ref int offset, string partName) {
            if (offset > payload.Length - 16) throw new InvalidDataException($"A range in '{partName}' is truncated.");
            uint firstRow = BitConverter.ToUInt32(payload, offset);
            uint lastRow = BitConverter.ToUInt32(payload, offset + 4);
            uint firstColumn = BitConverter.ToUInt32(payload, offset + 8);
            uint lastColumn = BitConverter.ToUInt32(payload, offset + 12);
            offset += 16;
            if (firstRow > lastRow || lastRow >= 1_048_576U || firstColumn > lastColumn || lastColumn >= 16_384U) {
                throw new InvalidDataException($"A range in '{partName}' is outside the worksheet.");
            }
            return new XlsbCellRange(checked((int)firstRow + 1), checked((int)lastRow + 1), checked((int)firstColumn + 1), checked((int)lastColumn + 1));
        }

        /// <summary>Reads an UncheckedSqRfX (MS-XLSB 2.5.156) as the space-separated A1 list SpreadsheetML uses.</summary>
        internal static string ReadSqRfX(byte[] payload, ref int offset, string partName) {
            if (offset > payload.Length - 4) throw new InvalidDataException($"A range list in '{partName}' is truncated.");
            int count = BitConverter.ToInt32(payload, offset);
            offset += 4;
            if (count < 0 || count > (payload.Length - offset) / 16) throw new InvalidDataException($"A range list in '{partName}' is invalid.");
            var parts = new string[count];
            for (int index = 0; index < count; index++) parts[index] = ReadRfX(payload, ref offset, partName).ToA1Reference();
            return string.Join(" ", parts);
        }
    }
}
