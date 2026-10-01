using System.Globalization;
using System.Text;
using System.Text.RegularExpressions;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;

namespace OfficeIMO.Excel {
    public partial class ExcelSheet {
        private readonly struct FormulaArgumentValue {
            internal FormulaArgumentValue(double? number, string? text, bool isUnresolvedFormula = false, bool isError = false,
                ExcelCellDataKind? sourceCellKind = null) {
                Number = number;
                Text = text;
                IsUnresolvedFormula = isUnresolvedFormula;
                IsError = isError;
                SourceCellKind = sourceCellKind;
            }

            internal double? Number { get; }
            internal string? Text { get; }
            internal bool IsUnresolvedFormula { get; }
            internal bool IsError { get; }
            // Preserve referenced cell types for functions whose coercion differs from ordinary numeric arguments.
            internal ExcelCellDataKind? SourceCellKind { get; }
            internal bool IsNumericAggregateValue => Number.HasValue
                && SourceCellKind != ExcelCellDataKind.Boolean && SourceCellKind != ExcelCellDataKind.Text;
            internal string? ErrorCode => IsError ? Text : null;
            internal bool HasValue => Number.HasValue || Text != null || IsError;

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
