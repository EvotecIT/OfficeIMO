namespace OfficeIMO.Markdown;

public sealed partial class MarkdownNativeDocument {
    /// <summary>Enumerates source-backed reference-style link definition fields in document order.</summary>
    public IEnumerable<MarkdownNativeReferenceLinkDefinitionField> EnumerateReferenceLinkDefinitionFields() {
        for (var i = 0; i < ReferenceLinkDefinitions.Count; i++) {
            foreach (var field in EnumerateReferenceLinkDefinitionFields(ReferenceLinkDefinitions[i])) {
                yield return field;
            }
        }
    }

    /// <summary>Enumerates source-backed reference-style link definition fields with the supplied field name in document order.</summary>
    public IEnumerable<MarkdownNativeReferenceLinkDefinitionField> EnumerateReferenceLinkDefinitionFields(string name) {
        if (string.IsNullOrWhiteSpace(name)) {
            yield break;
        }

        foreach (var field in EnumerateReferenceLinkDefinitionFields()) {
            if (string.Equals(field.Name, name, StringComparison.OrdinalIgnoreCase)) {
                yield return field;
            }
        }
    }

    /// <summary>Finds the first reference-style link definition whose source span contains the supplied 1-based line and column.</summary>
    public MarkdownReferenceLinkDefinition? FindReferenceLinkDefinitionAtPosition(int lineNumber, int columnNumber) {
        for (var i = 0; i < ReferenceLinkDefinitions.Count; i++) {
            var definition = ReferenceLinkDefinitions[i];
            if (ContainsPosition(definition.SourceSpan, lineNumber, columnNumber)
                || ReferenceDefinitionFieldsContainPosition(definition, lineNumber, columnNumber)) {
                return definition;
            }
        }

        return null;
    }

    /// <summary>Finds the first reference-style link definition field whose span contains the supplied 1-based line and column.</summary>
    public MarkdownNativeReferenceLinkDefinitionField? FindReferenceLinkDefinitionFieldAtPosition(int lineNumber, int columnNumber) {
        foreach (var field in EnumerateReferenceLinkDefinitionFields()) {
            if (field.SourceSpan.ContainsPosition(lineNumber, columnNumber)) {
                return field;
            }
        }

        return null;
    }

    /// <summary>Creates a non-mutating source edit that replaces a reference-style link definition field.</summary>
    public MarkdownNativeSourceEdit CreateReplaceEdit(MarkdownNativeReferenceLinkDefinitionField field, string replacementMarkdown) {
        if (field == null) {
            throw new ArgumentNullException(nameof(field));
        }

        return CreateReplaceEdit(field.SourceSpan, replacementMarkdown);
    }

    internal static IEnumerable<MarkdownNativeReferenceLinkDefinitionField> EnumerateReferenceLinkDefinitionFields(MarkdownReferenceLinkDefinition definition) {
        if (definition.OpeningMarkerSourceSpan.HasValue) {
            yield return new MarkdownNativeReferenceLinkDefinitionField("openingMarker", "[", definition.OpeningMarkerSourceSpan.Value, definition);
        }

        if (definition.LabelSourceSpan.HasValue) {
            yield return new MarkdownNativeReferenceLinkDefinitionField("label", definition.Label, definition.LabelSourceSpan.Value, definition);
        }

        if (definition.SeparatorMarkerSourceSpan.HasValue) {
            yield return new MarkdownNativeReferenceLinkDefinitionField("separatorMarker", "]:", definition.SeparatorMarkerSourceSpan.Value, definition);
        }

        if (definition.UrlSourceSpan.HasValue) {
            yield return new MarkdownNativeReferenceLinkDefinitionField("url", definition.Url, definition.UrlSourceSpan.Value, definition);
        }

        if (definition.TitleSourceSpan.HasValue) {
            yield return new MarkdownNativeReferenceLinkDefinitionField("title", definition.Title, definition.TitleSourceSpan.Value, definition);
        }
    }

    private static bool ReferenceDefinitionFieldsContainPosition(MarkdownReferenceLinkDefinition definition, int lineNumber, int columnNumber) {
        foreach (var field in EnumerateReferenceLinkDefinitionFields(definition)) {
            if (field.SourceSpan.ContainsPosition(lineNumber, columnNumber)) {
                return true;
            }
        }

        return false;
    }
}
