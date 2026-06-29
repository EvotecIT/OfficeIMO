using System.Text;

namespace OfficeIMO.Excel.LegacyXls.Biff {
    internal static class BiffHyperlinkTooltipReader {
        private const int FrtRefHeaderNoGrbitLength = 10;

        internal static bool TryRead(byte[] payload, out BiffHyperlinkTooltip? tooltip) {
            tooltip = null;
            if (payload.Length <= FrtRefHeaderNoGrbitLength) {
                return false;
            }

            ushort recordType = BiffRecordReader.ReadUInt16(payload, 0);
            if (recordType != (ushort)BiffRecordType.HLinkTooltip) {
                return false;
            }

            ushort firstRow = BiffRecordReader.ReadUInt16(payload, 2);
            ushort lastRow = BiffRecordReader.ReadUInt16(payload, 4);
            ushort firstColumn = BiffRecordReader.ReadUInt16(payload, 6);
            ushort lastColumn = BiffRecordReader.ReadUInt16(payload, 8);
            if (lastRow < firstRow || lastColumn < firstColumn) {
                return false;
            }

            int textByteCount = payload.Length - FrtRefHeaderNoGrbitLength;
            if (textByteCount <= 0 || textByteCount % 2 != 0) {
                return false;
            }

            string text = Encoding.Unicode.GetString(payload, FrtRefHeaderNoGrbitLength, textByteCount).TrimEnd('\0');
            if (string.IsNullOrWhiteSpace(text)) {
                return false;
            }

            tooltip = new BiffHyperlinkTooltip(
                firstRow + 1,
                firstColumn + 1,
                lastRow + 1,
                lastColumn + 1,
                text);
            return true;
        }
    }

    internal sealed class BiffHyperlinkTooltip {
        internal BiffHyperlinkTooltip(int startRow, int startColumn, int endRow, int endColumn, string text) {
            StartRow = startRow;
            StartColumn = startColumn;
            EndRow = endRow;
            EndColumn = endColumn;
            Text = text;
        }

        internal int StartRow { get; }

        internal int StartColumn { get; }

        internal int EndRow { get; }

        internal int EndColumn { get; }

        internal string Text { get; }
    }
}
