namespace OfficeIMO.Excel.Xlsb.Write;

internal static partial class XlsbWorksheetPartWriter {
    /// <summary>Streams the cell table; the optional BrtWsDim is omitted for an unknown-length source.</summary>
    internal sealed class StreamingRowSink : ExcelTabularRowSink {
        private readonly XlsbDirectRecordWriter _writer;
        private readonly int _columns;

        internal StreamingRowSink(Stream output, int columns) {
            _columns = columns;
            _writer = new XlsbDirectRecordWriter(output);
            _writer.WriteRecord(129); // BrtBeginSheet
            _writer.WriteRecord(BrtBeginSheetData);
        }

        internal override void BeginRow(int oneBasedRow) =>
            WriteDirectRowHeader(_writer, oneBasedRow - 1, _columns);
        internal override void EndRow() { }
        internal override void WriteText(string value) {
            CoerceValueHelper.ValidateSharedStringLength(value, nameof(value));
            _writer.WriteTextCell(BrtCellSt, Column, value);
        }
        internal override void WriteNumber(double value) => _writer.WriteNumberCell(BrtCellReal, Column, value, Style);
        internal override void WriteBoolean(bool value) => _writer.WriteBooleanCell(BrtCellBool, Column, value);
        internal override void WriteBlank() {
            _writer.WriteHeader(BrtCellBlank, 8);
            _writer.WriteUInt32(checked((uint)Column));
            _writer.WriteUInt32(Style);
        }
        internal override void Complete() {
            _writer.WriteRecord(BrtEndSheetData);
            _writer.WriteRecord(BrtEndSheet);
            _writer.Flush();
        }
        public override void Dispose() => _writer.Dispose();
    }
}
