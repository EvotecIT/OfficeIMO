namespace OfficeIMO.Markup;

/// <summary>
/// Emits starter C# code from the semantic OfficeIMO markup AST.
/// </summary>
public sealed partial class OfficeMarkupCSharpEmitter {
    public string Emit(OfficeMarkupDocument document, OfficeMarkupEmitterOptions? options = null) {
        if (document == null) {
            throw new ArgumentNullException(nameof(document));
        }

        options ??= new OfficeMarkupEmitterOptions();
        var sb = new StringBuilder();
        if (options.IncludeHeader) {
            sb.AppendLine("// Generated from OfficeIMO.Markup semantic AST.");
            sb.AppendLine("// Extend the emitted code when the authored markup reaches beyond Markdown.");
        }

        switch (document.Profile) {
            case OfficeMarkupProfile.Presentation:
                EmitPresentation(document, options, sb);
                break;
            case OfficeMarkupProfile.Workbook:
                EmitWorkbook(document, options, sb);
                break;
            case OfficeMarkupProfile.Document:
            case OfficeMarkupProfile.Common:
            default:
                EmitWordDocument(document, options, sb);
                break;
        }

        return sb.ToString().TrimEnd();
    }
}
