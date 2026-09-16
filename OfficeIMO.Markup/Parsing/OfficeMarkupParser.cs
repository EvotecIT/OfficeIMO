using OfficeIMO.Markdown;

namespace OfficeIMO.Markup;

/// <summary>
/// Parses Markdown plus explicit OfficeIMO semantic fenced blocks into a profile-aware AST.
/// </summary>
public static partial class OfficeMarkupParser {
    private static readonly string[] OfficeLanguages = {
        "officeimo",
        "officeimo-presentation",
        "officeimo-document",
        "officeimo-workbook",
        "officeimo-slide",
        "officeimo-page-break",
        "officeimo-pagebreak",
        "officeimo-section",
        "officeimo-header",
        "officeimo-footer",
        "officeimo-toc",
        "officeimo-sheet",
        "officeimo-range",
        "officeimo-formula",
        "officeimo-table",
        "officeimo-chart",
        "officeimo-format"
    };
    private static readonly string[] MermaidLanguages = { "mermaid" };

    /// <summary>Parses front matter, Markdown, and Office directives into a semantic document and diagnostics.</summary>
    /// <remarks>Null input is treated as empty text. Front-matter profile metadata can override the option's fallback profile; validation runs unless disabled in the options.</remarks>
    public static OfficeMarkupParseResult Parse(string markup, OfficeMarkupParserOptions? options = null) {
        options ??= new OfficeMarkupParserOptions();
        var source = (markup ?? string.Empty).Replace("\r\n", "\n").Replace('\r', '\n');
        var diagnostics = new List<OfficeMarkupDiagnostic>();
        var metadata = ExtractFrontMatter(ref source);
        var profile = ResolveProfile(options.Profile, metadata);
        var document = new OfficeMarkupDocument(profile);
        CopyAttributes(metadata, document.Metadata);

        if (!TryMapOfficeSyntax(source, document, profile, diagnostics)) {
            var markdownOptions = CreateMarkdownOptions(options);
            var markdownDocument = MarkdownReader.ParseSemanticProjection(source, markdownOptions);
            MapMarkdownBlocks(markdownDocument.Blocks, document.Blocks, profile, diagnostics);
        }

        if (options.Validate) {
            diagnostics.AddRange(OfficeMarkupValidator.Validate(document));
        }

        return new OfficeMarkupParseResult(document, diagnostics);
    }
}
