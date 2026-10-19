using OfficeIMO.Excel.LegacyXls.Model;

namespace OfficeIMO.Excel.LegacyXls.Write;

internal static partial class LegacyXlsWriter {
        private readonly struct LegacyXlsCell {
            private LegacyXlsCell(
                ushort row,
                ushort column,
                ushort styleIndex,
                LegacyXlsCellKind kind,
                string? textValue,
                double numberValue,
                bool booleanValue,
                byte errorValue,
                byte[]? formulaTokens,
                byte[]? formulaExtraData,
                byte[]? arrayFormulaPayload,
                byte[]? sharedFormulaPayload,
                IReadOnlyList<LegacyXlsTextFormattingRun>? textFormattingRuns,
                uint sharedStringIndex = uint.MaxValue) {
                Row = row;
                Column = column;
                StyleIndex = styleIndex;
                Kind = kind;
                TextValue = textValue;
                NumberValue = numberValue;
                BooleanValue = booleanValue;
                ErrorValue = errorValue;
                FormulaTokens = formulaTokens ?? Array.Empty<byte>();
                FormulaExtraData = formulaExtraData ?? Array.Empty<byte>();
                ArrayFormulaPayload = arrayFormulaPayload ?? Array.Empty<byte>();
                SharedFormulaPayload = sharedFormulaPayload ?? Array.Empty<byte>();
                TextFormattingRuns = textFormattingRuns ?? Array.Empty<LegacyXlsTextFormattingRun>();
                SharedStringIndex = sharedStringIndex;
            }

            internal ushort Row { get; }

            internal ushort Column { get; }

            internal ushort StyleIndex { get; }

            internal LegacyXlsCellKind Kind { get; }

            internal string? TextValue { get; }

            internal double NumberValue { get; }

            internal bool BooleanValue { get; }

            internal byte ErrorValue { get; }

            internal byte[] FormulaTokens { get; }

            internal byte[] FormulaExtraData { get; }

            internal byte[] ArrayFormulaPayload { get; }

            internal byte[] SharedFormulaPayload { get; }

            internal IReadOnlyList<LegacyXlsTextFormattingRun> TextFormattingRuns { get; }

            internal uint SharedStringIndex { get; }

            internal LegacyXlsCell WithSharedStringIndex(uint sharedStringIndex) =>
                new LegacyXlsCell(
                    Row,
                    Column,
                    StyleIndex,
                    Kind,
                    TextValue,
                    NumberValue,
                    BooleanValue,
                    ErrorValue,
                    FormulaTokens,
                    FormulaExtraData,
                    ArrayFormulaPayload,
                    SharedFormulaPayload,
                    TextFormattingRuns,
                    sharedStringIndex);

            internal static LegacyXlsCell Blank(ushort row, ushort column, ushort styleIndex) =>
                new LegacyXlsCell(row, column, styleIndex, LegacyXlsCellKind.Blank, null, 0, false, 0, null, null, null, null, null);

            internal static LegacyXlsCell Text(ushort row, ushort column, ushort styleIndex, string value, IReadOnlyList<LegacyXlsTextFormattingRun>? textFormattingRuns = null) =>
                new LegacyXlsCell(row, column, styleIndex, LegacyXlsCellKind.Text, value, 0, false, 0, null, null, null, null, textFormattingRuns);

            internal static LegacyXlsCell Number(ushort row, ushort column, ushort styleIndex, double value) =>
                new LegacyXlsCell(row, column, styleIndex, LegacyXlsCellKind.Number, null, value, false, 0, null, null, null, null, null);

            internal static LegacyXlsCell Boolean(ushort row, ushort column, ushort styleIndex, bool value) =>
                new LegacyXlsCell(row, column, styleIndex, LegacyXlsCellKind.Boolean, null, 0, value, 0, null, null, null, null, null);

            internal static LegacyXlsCell Error(ushort row, ushort column, ushort styleIndex, byte errorValue) =>
                new LegacyXlsCell(row, column, styleIndex, LegacyXlsCellKind.Error, null, 0, false, errorValue, null, null, null, null, null);

            internal static LegacyXlsCell FormulaNumber(ushort row, ushort column, ushort styleIndex, double value, byte[] formulaTokens, byte[]? formulaExtraData = null, byte[]? arrayFormulaPayload = null, byte[]? sharedFormulaPayload = null) =>
                new LegacyXlsCell(row, column, styleIndex, LegacyXlsCellKind.FormulaNumber, null, value, false, 0, formulaTokens, formulaExtraData, arrayFormulaPayload, sharedFormulaPayload, null);

            internal static LegacyXlsCell FormulaBoolean(ushort row, ushort column, ushort styleIndex, bool value, byte[] formulaTokens, byte[]? formulaExtraData = null, byte[]? arrayFormulaPayload = null, byte[]? sharedFormulaPayload = null) =>
                new LegacyXlsCell(row, column, styleIndex, LegacyXlsCellKind.FormulaBoolean, null, 0, value, 0, formulaTokens, formulaExtraData, arrayFormulaPayload, sharedFormulaPayload, null);

            internal static LegacyXlsCell FormulaText(ushort row, ushort column, ushort styleIndex, string value, byte[] formulaTokens, byte[]? formulaExtraData = null, byte[]? arrayFormulaPayload = null, byte[]? sharedFormulaPayload = null) =>
                new LegacyXlsCell(row, column, styleIndex, LegacyXlsCellKind.FormulaText, value, 0, false, 0, formulaTokens, formulaExtraData, arrayFormulaPayload, sharedFormulaPayload, null);

            internal static LegacyXlsCell FormulaEmpty(ushort row, ushort column, ushort styleIndex, byte[] tokens, byte[] extra, byte[]? array, byte[]? shared) =>
                new LegacyXlsCell(row, column, styleIndex, LegacyXlsCellKind.FormulaEmpty, null, 0, false, 0, tokens, extra, array, shared, null);

            internal static LegacyXlsCell FormulaError(ushort row, ushort column, ushort styleIndex, byte errorValue, byte[] formulaTokens, byte[]? formulaExtraData = null, byte[]? arrayFormulaPayload = null, byte[]? sharedFormulaPayload = null) =>
                new LegacyXlsCell(row, column, styleIndex, LegacyXlsCellKind.FormulaError, null, 0, false, errorValue, formulaTokens, formulaExtraData, arrayFormulaPayload, sharedFormulaPayload, null);
        }

        private enum LegacyXlsCellKind {
            Blank,
            Text,
            Number,
            Boolean,
            Error,
            FormulaNumber,
            FormulaBoolean,
            FormulaText,
            FormulaError,
            FormulaEmpty
        }

}
