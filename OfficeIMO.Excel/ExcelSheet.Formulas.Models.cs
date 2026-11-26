using System.Globalization;
using System.Text;
using System.Text.RegularExpressions;

namespace OfficeIMO.Excel {
    public partial class ExcelSheet {
        internal readonly struct FormulaArgumentValue {
            private readonly string? _text;
            // A calculated number spells its text only when something reads it.
            private readonly bool _textFromNumber;

            internal FormulaArgumentValue(double? number, string? text, bool isUnresolvedFormula = false, bool isError = false,
                ExcelCellDataKind? sourceCellKind = null, bool isBoolean = false, bool isUnevaluatedFormulaCache = false)
                : this(number, text, false, isUnresolvedFormula, isError, sourceCellKind, isBoolean, isUnevaluatedFormulaCache) {
            }

            private FormulaArgumentValue(double? number, string? text, bool textFromNumber, bool isUnresolvedFormula, bool isError,
                ExcelCellDataKind? sourceCellKind, bool isBoolean, bool isUnevaluatedFormulaCache) {
                Number = number;
                _text = text;
                _textFromNumber = textFromNumber && number.HasValue;
                IsUnresolvedFormula = isUnresolvedFormula;
                IsError = isError;
                SourceCellKind = sourceCellKind;
                IsBoolean = isBoolean || sourceCellKind == ExcelCellDataKind.Boolean;
                IsUnevaluatedFormulaCache = isUnevaluatedFormulaCache;
            }

            /// <summary>
            /// A calculated number whose text is its invariant spelling, the same value as
            /// <c>new FormulaArgumentValue(number, InvariantNumberText.Get(number))</c> without formatting it up front.
            /// </summary>
            internal static FormulaArgumentValue FromNumber(double number, bool isBoolean = false) =>
                new FormulaArgumentValue(number, null, true, false, false, null, isBoolean, false);

            internal double? Number { get; }
            internal string? Text => _textFromNumber ? InvariantNumberText.Get(Number!.Value) : _text;
            /// <summary>Whether <see cref="Text"/> is the number's invariant spelling, not yet formatted.</summary>
            internal bool IsTextNumberSpelling => _textFromNumber;
            internal bool IsUnresolvedFormula { get; }
            internal bool IsError { get; }
            internal bool IsBoolean { get; }
            // Preserve referenced cell types for functions whose coercion differs from ordinary numeric arguments.
            internal ExcelCellDataKind? SourceCellKind { get; }
            internal bool IsNumericAggregateValue => Number.HasValue
                && SourceCellKind != ExcelCellDataKind.Boolean && SourceCellKind != ExcelCellDataKind.Text;
            internal string? ErrorCode => IsError ? Text : null;
            internal bool HasValue => Number.HasValue || _text != null || IsError;

            internal static FormulaArgumentValue Boolean(bool value) => new FormulaArgumentValue(
                value ? 1d : 0d, value ? "1" : "0", isBoolean: true);

            // Only references apply reference coercion: a direct TRUE argument still counts as numeric.
            internal FormulaArgumentValue AsReferencedValue() => new FormulaArgumentValue(
                Number, _text, _textFromNumber, IsUnresolvedFormula, IsError,
                IsError ? ExcelCellDataKind.Error : IsBoolean ? ExcelCellDataKind.Boolean : SourceCellKind == ExcelCellDataKind.Text ? ExcelCellDataKind.Text
                    : Number.HasValue ? ExcelCellDataKind.Number
                    : _text != null ? ExcelCellDataKind.Text : ExcelCellDataKind.Blank, IsBoolean, IsUnevaluatedFormulaCache);

            // Value-returning functions keep their typed payload but do not
            // inherit the source cell's reference-specific aggregate coercion.
            internal FormulaArgumentValue AsScalarValue() => new FormulaArgumentValue(
                Number, _text, _textFromNumber, IsUnresolvedFormula, IsError, null, IsBoolean, IsUnevaluatedFormulaCache);
            internal bool IsUnevaluatedFormulaCache { get; }

            internal FormulaArgumentValue WithUnevaluatedFormulaCache() =>
                new FormulaArgumentValue(Number, _text, _textFromNumber, IsUnresolvedFormula, IsError,
                    SourceCellKind, IsBoolean, true);

            internal static FormulaArgumentValue UnresolvedFormula() {
                return new FormulaArgumentValue(null, null, isUnresolvedFormula: true);
            }

            internal static FormulaArgumentValue Error(string errorCode) {
                return new FormulaArgumentValue(null, errorCode, isError: true);
            }
        }

        /// <summary>Formula results of one calculation pass, keyed by sheet number, row and column.</summary>
        internal sealed class FormulaEvaluationResults : Dictionary<long, FormulaArgumentValue> {
            private static long _passes;
            private readonly Dictionary<string, int> _sheetNumbers = new(StringComparer.OrdinalIgnoreCase);

            /// <summary>A number unique to this pass in the process; never zero.</summary>
            internal long Pass { get; } = System.Threading.Interlocked.Increment(ref _passes);

            internal int GetSheetNumber(string sheetName) {
                if (!_sheetNumbers.TryGetValue(sheetName, out int number)) {
                    number = _sheetNumbers.Count + 1;
                    _sheetNumbers.Add(sheetName, number);
                }
                return number;
            }

            // Rows need 21 bits and columns 15, so the sheet number starts at bit 36.
            internal static long Key(int sheet, int row, int column) => ((long)sheet << 36) | ((long)row << 15) | (uint)column;
        }

        private enum FormulaCriteriaOperator {
            Equal,
            NotEqual,
            GreaterThan,
            LessThan,
            GreaterThanOrEqual,
            LessThanOrEqual
        }

        private readonly struct FormulaCriteria {
            internal FormulaCriteria(string op, string text, double? number, string? errorCode = null) {
                Operator = op;
                Text = text;
                Number = number;
                ErrorCode = errorCode;
                Kind = op switch {
                    ">" => FormulaCriteriaOperator.GreaterThan,
                    "<" => FormulaCriteriaOperator.LessThan,
                    ">=" => FormulaCriteriaOperator.GreaterThanOrEqual,
                    "<=" => FormulaCriteriaOperator.LessThanOrEqual,
                    "<>" => FormulaCriteriaOperator.NotEqual,
                    _ => FormulaCriteriaOperator.Equal
                };
                IsTextPattern = text.IndexOf('*') >= 0 || text.IndexOf('?') >= 0 || text.IndexOf('~') >= 0;
            }

            internal string Operator { get; }
            /// <summary>The operator, resolved once so each compared cell switches on a number.</summary>
            internal FormulaCriteriaOperator Kind { get; }
            /// <summary>Whether the text holds a wildcard or escape, found once for every compared cell.</summary>
            internal bool IsTextPattern { get; }
            internal string Text { get; }
            internal double? Number { get; }
            internal string? ErrorCode { get; }
        }
    }
}
