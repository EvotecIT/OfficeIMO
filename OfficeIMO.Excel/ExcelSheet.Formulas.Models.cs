using System.Globalization;
using System.Text;
using System.Text.RegularExpressions;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;

namespace OfficeIMO.Excel {
    public partial class ExcelSheet {
        private readonly struct FormulaArgumentValue {
            internal FormulaArgumentValue(double? number, string? text, bool isUnresolvedFormula = false, bool isError = false,
                ExcelCellDataKind? sourceCellKind = null, bool isBoolean = false, bool isUnevaluatedFormulaCache = false) {
                Number = number;
                Text = text;
                IsUnresolvedFormula = isUnresolvedFormula;
                IsError = isError;
                SourceCellKind = sourceCellKind;
                IsBoolean = isBoolean || sourceCellKind == ExcelCellDataKind.Boolean;
                IsUnevaluatedFormulaCache = isUnevaluatedFormulaCache;
            }

            internal double? Number { get; }
            internal string? Text { get; }
            internal bool IsUnresolvedFormula { get; }
            internal bool IsError { get; }
            internal bool IsBoolean { get; }
            // Preserve referenced cell types for functions whose coercion differs from ordinary numeric arguments.
            internal ExcelCellDataKind? SourceCellKind { get; }
            internal bool IsNumericAggregateValue => Number.HasValue
                && SourceCellKind != ExcelCellDataKind.Boolean && SourceCellKind != ExcelCellDataKind.Text;
            internal string? ErrorCode => IsError ? Text : null;
            internal bool HasValue => Number.HasValue || Text != null || IsError;

            internal static FormulaArgumentValue Boolean(bool value) => new FormulaArgumentValue(
                value ? 1d : 0d, value ? "1" : "0", isBoolean: true);

            // Only references apply reference coercion: a direct TRUE argument still counts as numeric.
            internal FormulaArgumentValue AsReferencedValue() => new FormulaArgumentValue(
                Number, Text, IsUnresolvedFormula, IsError,
                sourceCellKind: IsError ? ExcelCellDataKind.Error : IsBoolean ? ExcelCellDataKind.Boolean : SourceCellKind == ExcelCellDataKind.Text ? ExcelCellDataKind.Text
                    : Number.HasValue ? ExcelCellDataKind.Number
                    : Text != null ? ExcelCellDataKind.Text : ExcelCellDataKind.Blank, isBoolean: IsBoolean, isUnevaluatedFormulaCache: IsUnevaluatedFormulaCache);
            internal bool IsUnevaluatedFormulaCache { get; }

            internal FormulaArgumentValue WithUnevaluatedFormulaCache() =>
                new FormulaArgumentValue(Number, Text, IsUnresolvedFormula, IsError,
                    sourceCellKind: SourceCellKind, isBoolean: IsBoolean, isUnevaluatedFormulaCache: true);

            internal static FormulaArgumentValue UnresolvedFormula() {
                return new FormulaArgumentValue(null, null, isUnresolvedFormula: true);
            }

            internal static FormulaArgumentValue Error(string errorCode) {
                return new FormulaArgumentValue(null, errorCode, isError: true);
            }
        }

        private readonly struct FormulaCriteria {
            internal FormulaCriteria(string op, string text, double? number) {
                Operator = op;
                Text = text;
                Number = number;
            }

            internal string Operator { get; }
            internal string Text { get; }
            internal double? Number { get; }
        }
    }
}
