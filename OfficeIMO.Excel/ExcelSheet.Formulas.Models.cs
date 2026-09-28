using System.Globalization;
using System.Text;
using System.Text.RegularExpressions;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;

namespace OfficeIMO.Excel {
    public partial class ExcelSheet {
        private readonly struct FormulaArgumentValue {
            internal FormulaArgumentValue(double? number, string? text, bool isUnresolvedFormula = false, bool isError = false, bool isBoolean = false,
                bool isUnevaluatedFormulaCache = false) {
                Number = number;
                Text = text;
                IsUnresolvedFormula = isUnresolvedFormula;
                IsError = isError;
                IsBoolean = isBoolean;
                IsUnevaluatedFormulaCache = isUnevaluatedFormulaCache;
            }

            internal double? Number { get; }
            internal string? Text { get; }
            internal bool IsUnresolvedFormula { get; }
            internal bool IsError { get; }
            internal bool IsBoolean { get; }
            internal bool IsUnevaluatedFormulaCache { get; }
            internal string? ErrorCode => IsError ? Text : null;
            internal bool HasValue => Number.HasValue || Text != null || IsError;

            internal FormulaArgumentValue WithUnevaluatedFormulaCache() =>
                new FormulaArgumentValue(Number, Text, IsUnresolvedFormula, IsError, IsBoolean,
                    isUnevaluatedFormulaCache: true);

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
