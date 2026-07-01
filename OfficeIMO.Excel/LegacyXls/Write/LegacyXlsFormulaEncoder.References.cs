namespace OfficeIMO.Excel.LegacyXls.Write {
    internal static partial class LegacyXlsFormulaEncoder {
        private static bool TryEncodeSheetQualifiedReference(
            string text,
            bool allowArea,
            LegacyXlsFormulaNameIndex nameIndex,
            out byte[] tokens) {
            tokens = Array.Empty<byte>();
            if (!SheetNameLookup.TryParseSheetQualifiedReference(text, out string sheetName, out string localReference, allowExternalWorkbookReferences: true)
                || !TryResolveExternSheetIndex(sheetName, nameIndex, out ushort externSheetIndex)) {
                return false;
            }

            if (allowArea && TryParseAreaReference(localReference, out FormulaReference first, out FormulaReference last)) {
                tokens = BuildArea3dReferenceToken(externSheetIndex, first, last);
                return true;
            }

            if (TryParseCellReference(localReference, out FormulaReference reference)) {
                tokens = BuildCell3dReferenceToken(externSheetIndex, reference);
                return true;
            }

            return false;
        }

        private static bool TryResolveExternSheetIndex(
            string sheetName,
            LegacyXlsFormulaNameIndex nameIndex,
            out ushort externSheetIndex) {
            externSheetIndex = 0;
            if (nameIndex.TryGetExternSheetIndex(sheetName, out externSheetIndex)) {
                return true;
            }

            int rangeSeparator = sheetName.IndexOf(':');
            if (rangeSeparator > 0 && rangeSeparator < sheetName.Length - 1 && sheetName.IndexOf(':', rangeSeparator + 1) < 0) {
                string firstSheetName = sheetName.Substring(0, rangeSeparator).Trim();
                string lastSheetName = sheetName.Substring(rangeSeparator + 1).Trim();
                return nameIndex.TryGetExternSheetRangeIndex(firstSheetName, lastSheetName, out externSheetIndex);
            }

            return false;
        }

        private static byte[] BuildCell3dReferenceToken(ushort externSheetIndex, FormulaReference reference) {
            byte[] token = new byte[7];
            token[0] = 0x5a;
            WriteUInt16(token, 1, externSheetIndex);
            WriteUInt16(token, 3, reference.Row);
            WriteUInt16(token, 5, reference.ColumnBits);
            return token;
        }

        private static byte[] BuildArea3dReferenceToken(ushort externSheetIndex, FormulaReference first, FormulaReference last) {
            byte[] token = new byte[11];
            token[0] = 0x5b;
            WriteUInt16(token, 1, externSheetIndex);
            WriteUInt16(token, 3, first.Row);
            WriteUInt16(token, 5, last.Row);
            WriteUInt16(token, 7, first.ColumnBits);
            WriteUInt16(token, 9, last.ColumnBits);
            return token;
        }
    }
}
