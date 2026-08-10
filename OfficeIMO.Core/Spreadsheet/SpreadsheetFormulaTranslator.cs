using System;
using System.Collections.Generic;
using System.Text;

namespace OfficeIMO.Spreadsheet;

internal static class SpreadsheetFormulaTranslator {
    internal static SpreadsheetFormulaTranslationResult Translate(
        SpreadsheetFormulaSyntaxTree tree,
        SpreadsheetFormulaDialect targetDialect) {
        if (tree.Dialect == targetDialect) {
            return new SpreadsheetFormulaTranslationResult(
                tree.Text,
                tree.Dialect,
                targetDialect,
                tree.Diagnostics);
        }

        var diagnostics = new List<SpreadsheetFormulaDiagnostic>(tree.Diagnostics);
        var output = new StringBuilder(tree.Text.Length + 8);
        output.Append(targetDialect == SpreadsheetFormulaDialect.OpenFormula ? "of:=" : "=");
        AppendChildren(tree.Root.Children, tree.Dialect, targetDialect, output, diagnostics);
        return new SpreadsheetFormulaTranslationResult(
            output.ToString(),
            tree.Dialect,
            targetDialect,
            diagnostics);
    }

    private static void AppendChildren(
        IReadOnlyList<SpreadsheetFormulaSyntaxNode> nodes,
        SpreadsheetFormulaDialect sourceDialect,
        SpreadsheetFormulaDialect targetDialect,
        StringBuilder output,
        ICollection<SpreadsheetFormulaDiagnostic> diagnostics) {
        foreach (SpreadsheetFormulaSyntaxNode node in nodes) {
            if (node.Kind != SpreadsheetFormulaSyntaxKind.Token) {
                AppendChildren(node.Children, sourceDialect, targetDialect, output, diagnostics);
                continue;
            }

            switch (node.TokenKind) {
                case SpreadsheetFormulaTokenKind.Prefix:
                    break;
                case SpreadsheetFormulaTokenKind.Reference:
                    AppendReference(node, targetDialect, output, diagnostics);
                    break;
                case SpreadsheetFormulaTokenKind.ArgumentSeparator:
                    output.Append(targetDialect == SpreadsheetFormulaDialect.OpenFormula ? ';' : ',');
                    break;
                case SpreadsheetFormulaTokenKind.ArrayColumnSeparator:
                    output.Append(targetDialect == SpreadsheetFormulaDialect.OpenFormula ? ';' : ',');
                    break;
                case SpreadsheetFormulaTokenKind.ArrayRowSeparator:
                    output.Append(targetDialect == SpreadsheetFormulaDialect.OpenFormula ? '|' : ';');
                    break;
                case SpreadsheetFormulaTokenKind.UnionOperator:
                    output.Append(targetDialect == SpreadsheetFormulaDialect.OpenFormula ? '~' : ',');
                    break;
                case SpreadsheetFormulaTokenKind.IntersectionOperator:
                    output.Append(targetDialect == SpreadsheetFormulaDialect.OpenFormula ? '!' : ' ');
                    break;
                case SpreadsheetFormulaTokenKind.ErrorLiteral:
                    AppendErrorLiteral(node, sourceDialect, targetDialect, output);
                    break;
                case SpreadsheetFormulaTokenKind.Unsupported:
                    output.Append(node.Text);
                    diagnostics.Add(new SpreadsheetFormulaDiagnostic(
                        "FORMULA_TRANSLATION_UNSUPPORTED",
                        SpreadsheetFormulaDiagnosticSeverity.Error,
                        $"The syntax '{node.Text}' has no safe {targetDialect} translation.",
                        node.Position,
                        node.Text.Length));
                    break;
                default:
                    output.Append(node.Text);
                    break;
            }
        }
    }

    private static void AppendReference(
        SpreadsheetFormulaSyntaxNode node,
        SpreadsheetFormulaDialect targetDialect,
        StringBuilder output,
        ICollection<SpreadsheetFormulaDiagnostic> diagnostics) {
        if (node.Reference == null) {
            output.Append(node.Text);
            diagnostics.Add(new SpreadsheetFormulaDiagnostic(
                "FORMULA_TRANSLATION_REFERENCE",
                SpreadsheetFormulaDiagnosticSeverity.Error,
                "A formula reference was not represented by typed address syntax.",
                node.Position,
                node.Text.Length));
            return;
        }

        if (targetDialect == SpreadsheetFormulaDialect.OpenFormula) {
            output.Append('[')
                .Append(node.Reference.Format(SpreadsheetAddressDialect.OpenDocument))
                .Append(']');
        } else {
            if (!FitsExcelBounds(node.Reference)) {
                output.Append(node.Text);
                diagnostics.Add(new SpreadsheetFormulaDiagnostic(
                    "FORMULA_TRANSLATION_REFERENCE_BOUNDS",
                    SpreadsheetFormulaDiagnosticSeverity.Error,
                    "The reference exceeds Excel's row or column bounds.",
                    node.Position,
                    node.Text.Length));
                return;
            }
            if (!node.Reference.TryFormat(SpreadsheetAddressDialect.ExcelA1, out string formatted)) {
                output.Append(node.Text);
                diagnostics.Add(new SpreadsheetFormulaDiagnostic(
                    "FORMULA_TRANSLATION_REFERENCE_SHEETS",
                    SpreadsheetFormulaDiagnosticSeverity.Error,
                    "The reference uses relative or different worksheet endpoints that Excel A1 syntax cannot preserve.",
                    node.Position,
                    node.Text.Length));
                return;
            }
            output.Append(formatted);
        }
    }

    private static bool FitsExcelBounds(SpreadsheetRangeReference reference) =>
        FitsExcelBounds(reference.Start) && (reference.End == null || FitsExcelBounds(reference.End));

    private static bool FitsExcelBounds(SpreadsheetCellReference reference) =>
        (!reference.Column.HasValue || reference.Column.Value <= 16_384)
        && (!reference.Row.HasValue || reference.Row.Value <= 1_048_576);

    private static void AppendErrorLiteral(
        SpreadsheetFormulaSyntaxNode node,
        SpreadsheetFormulaDialect sourceDialect,
        SpreadsheetFormulaDialect targetDialect,
        StringBuilder output) {
        string text = node.Text;
        if (sourceDialect == SpreadsheetFormulaDialect.OpenFormula &&
            targetDialect == SpreadsheetFormulaDialect.ExcelA1 &&
            text.Length >= 2 && text[0] == '[' && text[text.Length - 1] == ']') {
            output.Append(text, 1, text.Length - 2);
            return;
        }
        if (sourceDialect == SpreadsheetFormulaDialect.ExcelA1 &&
            targetDialect == SpreadsheetFormulaDialect.OpenFormula &&
            string.Equals(text, "#REF!", StringComparison.OrdinalIgnoreCase)) {
            output.Append('[').Append(text).Append(']');
            return;
        }
        output.Append(text);
    }
}