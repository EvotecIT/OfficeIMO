using OfficeIMO.Excel.Xlsb.Model;

namespace OfficeIMO.Excel.Xlsb.Write {
    internal enum XlsbWriteCellKind {
        Blank,
        Number,
        Text,
        Boolean,
        Error,
        FormulaNumber,
        FormulaText,
        FormulaBoolean,
        FormulaError
    }

    internal sealed class XlsbWriteCell {
        internal XlsbWriteCell(
            int row,
            int column,
            uint styleIndex,
            XlsbWriteCellKind kind,
            object? value,
            byte[]? formulaPayload = null,
            int? sourceRecordType = null,
            byte[]? sourceRecordData = null) {
            Row = row;
            Column = column;
            StyleIndex = styleIndex;
            Kind = kind;
            Value = value;
            FormulaPayload = formulaPayload;
            SourceRecordType = sourceRecordType;
            SourceRecordData = sourceRecordData;
        }

        internal XlsbGeneratedRecord? FormulaGroup { get; set; }

        internal int Row { get; }

        internal int Column { get; }

        internal uint StyleIndex { get; }

        internal XlsbWriteCellKind Kind { get; }

        internal object? Value { get; }

        internal byte[]? FormulaPayload { get; }

        internal int? SourceRecordType { get; }

        internal byte[]? SourceRecordData { get; }

        internal static XlsbWriteCell PreserveSource(XlsbCell source, uint? styleIndex = null) {
            if (source == null) throw new ArgumentNullException(nameof(source));
            byte[] data = source.SourceRecordData ?? throw new InvalidDataException("The XLSB source cell has no preserved record payload.");
            if (styleIndex.HasValue && styleIndex.Value != source.StyleIndex) {
                if (data.Length < 8 || styleIndex.Value > 0xFFFFFF) throw new InvalidDataException("Invalid native XLSB style reference.");
                data = (byte[])data.Clone(); for (int index = 0; index < 3; index++) data[4 + index] = (byte)(styleIndex.Value >> (index * 8));
            }
            return new XlsbWriteCell(
                source.Row,
                source.Column,
                styleIndex ?? source.StyleIndex,
                ToWriteKind(source),
                source.Value,
                source.FormulaPayloadBytes,
                source.SourceRecordType,
                data);
        }

        private static XlsbWriteCellKind ToWriteKind(XlsbCell source) {
            bool formula = source.FormulaBytes != null;
            switch (source.Kind) {
                case XlsbCellValueKind.Blank:
                    return XlsbWriteCellKind.Blank;
                case XlsbCellValueKind.Number:
                    return formula ? XlsbWriteCellKind.FormulaNumber : XlsbWriteCellKind.Number;
                case XlsbCellValueKind.Text:
                    return formula ? XlsbWriteCellKind.FormulaText : XlsbWriteCellKind.Text;
                case XlsbCellValueKind.Boolean:
                    return formula ? XlsbWriteCellKind.FormulaBoolean : XlsbWriteCellKind.Boolean;
                case XlsbCellValueKind.Error:
                    return formula ? XlsbWriteCellKind.FormulaError : XlsbWriteCellKind.Error;
                default:
                    throw new InvalidOperationException($"Unsupported XLSB cell kind {source.Kind}.");
            }
        }
    }
}
