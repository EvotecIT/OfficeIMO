using DocumentFormat.OpenXml.Wordprocessing;

namespace OfficeIMO.Word;

public partial class WordDocument {
    /// <summary>Adds a requested library style and its dependencies without replacing authored definitions.</summary>
    internal void EnsureStyleDefinition(string styleId) {
        if (FileOpenAccess == FileAccess.Read || string.IsNullOrEmpty(styleId)) return;
        var visited = new HashSet<string>(StringComparer.Ordinal);
        EnsureStyleDefinition(styleId, visited);
    }

    private void EnsureStyleDefinition(string styleId, HashSet<string> visited) {
        if (!visited.Add(styleId)) return;
        var styles = _wordprocessingDocument.MainDocumentPart?.StyleDefinitionsPart?.Styles;
        Style? definition = styles?.Elements<Style>().FirstOrDefault(style =>
            string.Equals(style.StyleId?.Value, styleId, StringComparison.Ordinal));
        if (definition == null) {
            definition = FindLibraryStyleDefinition(styleId);
            if (definition == null) return; // An external/custom ID may intentionally be resolved by the document consumer.
            EnsureStyleDefinitionsInitialized();
            styles = _wordprocessingDocument.MainDocumentPart?.StyleDefinitionsPart?.Styles;
            if (styles == null) return;
            // Selecting a style for content must preserve both existing defaults and their absence.
            definition.Default = null;
            styles.Append(definition);
        }

        foreach (string? dependency in new[] {
                     definition.BasedOn?.Val?.Value,
                     definition.LinkedStyle?.Val?.Value,
                     definition.NextParagraphStyle?.Val?.Value }) {
            if (!string.IsNullOrEmpty(dependency)) EnsureStyleDefinition(dependency!, visited);
        }
    }

    private Style? FindLibraryStyleDefinition(string styleId) {
        Style? registered = WordParagraphStyle.CustomStyles.FirstOrDefault(style =>
            string.Equals(style.StyleId?.Value, styleId, StringComparison.Ordinal));
        if (registered != null) return registered;

        if (WordTableStyles.TryGetStyle(styleId, out var tableStyle))
            return tableStyle == WordTableStyle.OfficeIMOModern
                ? WordThemeDefaults.CreateTableStyle(EffectiveCreationTheme)
                : WordTableStyles.GetStyleDefinition(tableStyle);

        foreach (WordParagraphStyles paragraphStyle in global::OfficeIMO.Internal.EnumCompat.GetValues<WordParagraphStyles>()) {
            if (paragraphStyle != WordParagraphStyles.Custom &&
                string.Equals(paragraphStyle.ToStringStyle(), styleId, StringComparison.Ordinal))
                return WordParagraphStyle.GetOpenXmlStyleDefinition(paragraphStyle);
        }

        return (Style?)BuiltInStyleDefinitions.Value.Elements<Style>().FirstOrDefault(style =>
            string.Equals(style.StyleId?.Value, styleId, StringComparison.Ordinal))?.CloneNode(true);
    }

    /// <summary>Installs styles referenced by a newly generated content block, without scanning the loaded document.</summary>
    internal void EnsureContentStyleDefinitions(OpenXmlElement content) {
        foreach (OpenXmlElement reference in content.Descendants().Where(element =>
                     element is ParagraphStyleId or RunStyle or TableStyle)) {
            string? styleId = reference switch {
                ParagraphStyleId paragraph => paragraph.Val?.Value,
                RunStyle run => run.Val?.Value,
                TableStyle table => table.Val?.Value,
                _ => null
            };
            if (!string.IsNullOrEmpty(styleId)) EnsureStyleDefinition(styleId!);
        }
    }
}
