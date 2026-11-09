using DocumentFormat.OpenXml.Wordprocessing;

namespace OfficeIMO.Word;

public partial class WordDocument {
    private StyleDefinitionResolutionContext? _styleDefinitionResolutionContext;

    /// <summary>Adds a requested library style and its dependencies without replacing authored definitions.</summary>
    internal void EnsureStyleDefinition(string styleId) {
        if (FileOpenAccess == FileAccess.Read || string.IsNullOrEmpty(styleId)) return;
        (_styleDefinitionResolutionContext ?? new StyleDefinitionResolutionContext(this)).Ensure(styleId);
    }

    /// <summary>Shares style resolution only while an owned projection controls the document's style DOM.</summary>
    private StyleDefinitionResolutionScope BeginStyleDefinitionResolutionScope() => new(this);

    private sealed class StyleDefinitionResolutionScope : IDisposable {
        private readonly WordDocument _document;
        private readonly StyleDefinitionResolutionContext? _previous;
        private bool _disposed;

        internal StyleDefinitionResolutionScope(WordDocument document) {
            _document = document;
            _previous = document._styleDefinitionResolutionContext;
            document._styleDefinitionResolutionContext = new StyleDefinitionResolutionContext(document, indexDefinitions: true);
        }

        public void Dispose() {
            if (_disposed) return;
            _document._styleDefinitionResolutionContext = _previous;
            _disposed = true;
        }
    }

    private sealed class StyleDefinitionResolutionContext {
        private readonly WordDocument _document;
        private readonly Dictionary<string, Style>? _definitions;
        private readonly HashSet<string> _visited = new(StringComparer.Ordinal);

        internal StyleDefinitionResolutionContext(WordDocument document, bool indexDefinitions = false) {
            _document = document;
            if (!indexDefinitions) return;
            _definitions = new Dictionary<string, Style>(StringComparer.Ordinal);
            var styles = document._wordprocessingDocument.MainDocumentPart?.StyleDefinitionsPart?.Styles;
            if (styles == null) return;
            foreach (Style definition in styles.Elements<Style>()) {
                string? id = definition.StyleId?.Value;
                if (!string.IsNullOrEmpty(id) && !_definitions.ContainsKey(id!)) _definitions.Add(id!, definition);
            }
        }

        internal void Ensure(string styleId) {
            if (!_visited.Add(styleId)) return;
            Style? definition;
            if (_definitions != null) {
                _definitions.TryGetValue(styleId, out definition);
            } else {
                var styles = _document._wordprocessingDocument.MainDocumentPart?.StyleDefinitionsPart?.Styles;
                definition = styles?.Elements<Style>().FirstOrDefault(style =>
                    string.Equals(style.StyleId?.Value, styleId, StringComparison.Ordinal));
            }
            if (definition == null) {
                definition = _document.FindLibraryStyleDefinition(styleId);
                if (definition == null) return; // An external/custom ID may intentionally be resolved by the document consumer.
                _document.EnsureStyleDefinitionsInitialized();
                var styles = _document._wordprocessingDocument.MainDocumentPart?.StyleDefinitionsPart?.Styles;
                if (styles == null) return;
                // Selecting a style for content must preserve both existing defaults and their absence.
                definition.Default = null;
                styles.Append(definition);
                _definitions?.Add(styleId, definition);
            }

            EnsureDependency(definition.BasedOn?.Val?.Value);
            EnsureDependency(definition.LinkedStyle?.Val?.Value);
            EnsureDependency(definition.NextParagraphStyle?.Val?.Value);
        }

        private void EnsureDependency(string? styleId) {
            if (!string.IsNullOrEmpty(styleId)) Ensure(styleId!);
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
        if (FileOpenAccess == FileAccess.Read) return;
        var context = _styleDefinitionResolutionContext ?? new StyleDefinitionResolutionContext(this);
        foreach (OpenXmlElement reference in content.Descendants().Where(element =>
                     element is ParagraphStyleId or RunStyle or TableStyle)) {
            string? styleId = reference switch {
                ParagraphStyleId paragraph => paragraph.Val?.Value,
                RunStyle run => run.Val?.Value,
                TableStyle table => table.Val?.Value,
                _ => null
            };
            if (!string.IsNullOrEmpty(styleId)) context.Ensure(styleId!);
        }
    }
}
