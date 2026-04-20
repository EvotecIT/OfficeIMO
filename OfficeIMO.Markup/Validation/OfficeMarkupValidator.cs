namespace OfficeIMO.Markup;

/// <summary>
/// Validates that profile-specific AST nodes are only used in compatible authoring profiles.
/// </summary>
public static class OfficeMarkupValidator {
    public static IReadOnlyList<OfficeMarkupDiagnostic> Validate(OfficeMarkupDocument document) {
        if (document == null) {
            throw new ArgumentNullException(nameof(document));
        }

        var diagnostics = new List<OfficeMarkupDiagnostic>();
        foreach (var block in document.DescendantsAndSelf()) {
            if (!IsAllowed(document.Profile, block.Kind)) {
                diagnostics.Add(new OfficeMarkupDiagnostic(
                    OfficeMarkupDiagnosticSeverity.Error,
                    $"Node '{block.Kind}' is not allowed in the '{document.Profile}' OfficeIMO markup profile.",
                    block));
            }

            ValidateRequiredFields(block, diagnostics);
        }

        return diagnostics;
    }

    private static bool IsAllowed(OfficeMarkupProfile profile, OfficeMarkupNodeKind kind) {
        if (IsCommon(kind)) {
            return true;
        }

        switch (profile) {
            case OfficeMarkupProfile.Common:
                return false;
            case OfficeMarkupProfile.Presentation:
                return kind == OfficeMarkupNodeKind.Slide
                    || kind == OfficeMarkupNodeKind.Chart
                    || kind == OfficeMarkupNodeKind.TextBox
                    || kind == OfficeMarkupNodeKind.Columns
                    || kind == OfficeMarkupNodeKind.Column
                    || kind == OfficeMarkupNodeKind.Card;
            case OfficeMarkupProfile.Document:
                return kind == OfficeMarkupNodeKind.PageBreak
                    || kind == OfficeMarkupNodeKind.Section
                    || kind == OfficeMarkupNodeKind.HeaderFooter
                    || kind == OfficeMarkupNodeKind.TableOfContents
                    || kind == OfficeMarkupNodeKind.Chart;
            case OfficeMarkupProfile.Workbook:
                return kind == OfficeMarkupNodeKind.Sheet
                    || kind == OfficeMarkupNodeKind.Range
                    || kind == OfficeMarkupNodeKind.Formula
                    || kind == OfficeMarkupNodeKind.NamedTable
                    || kind == OfficeMarkupNodeKind.Chart
                    || kind == OfficeMarkupNodeKind.Formatting;
            default:
                return false;
        }
    }

    private static bool IsCommon(OfficeMarkupNodeKind kind) {
        return kind == OfficeMarkupNodeKind.Heading
            || kind == OfficeMarkupNodeKind.Paragraph
            || kind == OfficeMarkupNodeKind.List
            || kind == OfficeMarkupNodeKind.Code
            || kind == OfficeMarkupNodeKind.Image
            || kind == OfficeMarkupNodeKind.Table
            || kind == OfficeMarkupNodeKind.Diagram
            || kind == OfficeMarkupNodeKind.Extension
            || kind == OfficeMarkupNodeKind.RawMarkdown;
    }

    private static void ValidateRequiredFields(OfficeMarkupBlock block, IList<OfficeMarkupDiagnostic> diagnostics) {
        switch (block) {
            case OfficeMarkupSheetBlock sheet when string.IsNullOrWhiteSpace(sheet.Name):
                diagnostics.Add(Required(block, "Workbook sheet nodes require a non-empty name."));
                break;
            case OfficeMarkupRangeBlock range when string.IsNullOrWhiteSpace(range.Address):
                diagnostics.Add(Required(block, "Workbook range nodes require an address."));
                break;
            case OfficeMarkupFormulaBlock formula when string.IsNullOrWhiteSpace(formula.Cell) || string.IsNullOrWhiteSpace(formula.Expression):
                diagnostics.Add(Required(block, "Workbook formula nodes require both cell and expression."));
                break;
            case OfficeMarkupNamedTableBlock table when string.IsNullOrWhiteSpace(table.Name) || string.IsNullOrWhiteSpace(table.Range):
                diagnostics.Add(Required(block, "Workbook named table nodes require both name and range."));
                break;
        }
    }

    private static OfficeMarkupDiagnostic Required(OfficeMarkupBlock block, string message) =>
        new OfficeMarkupDiagnostic(OfficeMarkupDiagnosticSeverity.Error, message, block);
}
