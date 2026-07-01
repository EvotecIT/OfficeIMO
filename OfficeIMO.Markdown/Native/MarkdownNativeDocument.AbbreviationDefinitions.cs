namespace OfficeIMO.Markdown;

public sealed partial class MarkdownNativeDocument {
    /// <summary>Enumerates source-backed abbreviation definition fields in document order.</summary>
    public IEnumerable<MarkdownNativeAbbreviationDefinitionField> EnumerateAbbreviationDefinitionFields() {
        for (var i = 0; i < AbbreviationDefinitions.Count; i++) {
            foreach (var field in EnumerateAbbreviationDefinitionFields(AbbreviationDefinitions[i])) {
                yield return field;
            }
        }
    }

    /// <summary>Enumerates source-backed abbreviation definition fields with the supplied field name in document order.</summary>
    public IEnumerable<MarkdownNativeAbbreviationDefinitionField> EnumerateAbbreviationDefinitionFields(string name) {
        if (string.IsNullOrWhiteSpace(name)) {
            yield break;
        }

        foreach (var field in EnumerateAbbreviationDefinitionFields()) {
            if (string.Equals(field.Name, name, StringComparison.OrdinalIgnoreCase)) {
                yield return field;
            }
        }
    }

    /// <summary>Finds the first abbreviation definition whose source span contains the supplied 1-based line and column.</summary>
    public MarkdownAbbreviationDefinition? FindAbbreviationDefinitionAtPosition(int lineNumber, int columnNumber) {
        for (var i = 0; i < AbbreviationDefinitions.Count; i++) {
            var definition = AbbreviationDefinitions[i];
            if (ContainsPosition(definition.SourceSpan, lineNumber, columnNumber)
                || AbbreviationDefinitionFieldsContainPosition(definition, lineNumber, columnNumber)) {
                return definition;
            }
        }

        return null;
    }

    /// <summary>Finds the first abbreviation definition field whose span contains the supplied 1-based line and column.</summary>
    public MarkdownNativeAbbreviationDefinitionField? FindAbbreviationDefinitionFieldAtPosition(int lineNumber, int columnNumber) {
        foreach (var field in EnumerateAbbreviationDefinitionFields()) {
            if (field.SourceSpan.ContainsPosition(lineNumber, columnNumber)) {
                return field;
            }
        }

        return null;
    }

    /// <summary>Creates a source slice over the normalized markdown text that backs an abbreviation definition.</summary>
    public bool TryCreateSourceSlice(MarkdownAbbreviationDefinition abbreviationDefinition, out MarkdownSourceSlice slice) {
        if (abbreviationDefinition == null || !abbreviationDefinition.SourceSpan.HasValue) {
            slice = default;
            return false;
        }

        return TryCreateSourceSlice(abbreviationDefinition.SourceSpan.Value, out slice);
    }

    /// <summary>Creates a source slice over the normalized markdown text that backs an abbreviation definition field.</summary>
    public bool TryCreateSourceSlice(MarkdownNativeAbbreviationDefinitionField field, out MarkdownSourceSlice slice) {
        if (field == null) {
            slice = default;
            return false;
        }

        return TryCreateSourceSlice(field.SourceSpan, out slice);
    }

    /// <summary>Creates a source slice over the original reader input that backs an abbreviation definition when trivia was preserved.</summary>
    public bool TryCreateOriginalSourceSlice(MarkdownAbbreviationDefinition abbreviationDefinition, out MarkdownSourceSlice slice) {
        return TryCreateOriginalSourceSlice(abbreviationDefinition, out slice, out _);
    }

    /// <summary>Creates a source slice over the original reader input that backs an abbreviation definition when trivia was preserved.</summary>
    public bool TryCreateOriginalSourceSlice(
        MarkdownAbbreviationDefinition abbreviationDefinition,
        out MarkdownSourceSlice slice,
        out MarkdownOriginalSourceSliceFailureReason failureReason) {
        if (abbreviationDefinition == null || !abbreviationDefinition.SourceSpan.HasValue) {
            slice = default;
            failureReason = MarkdownOriginalSourceSliceFailureReason.SourceSpanUnavailable;
            return false;
        }

        return TryCreateOriginalSourceSlice(abbreviationDefinition.SourceSpan.Value, out slice, out failureReason);
    }

    /// <summary>Creates a source slice over the original reader input that backs an abbreviation definition field when trivia was preserved.</summary>
    public bool TryCreateOriginalSourceSlice(MarkdownNativeAbbreviationDefinitionField field, out MarkdownSourceSlice slice) {
        return TryCreateOriginalSourceSlice(field, out slice, out _);
    }

    /// <summary>Creates a source slice over the original reader input that backs an abbreviation definition field when trivia was preserved.</summary>
    public bool TryCreateOriginalSourceSlice(
        MarkdownNativeAbbreviationDefinitionField field,
        out MarkdownSourceSlice slice,
        out MarkdownOriginalSourceSliceFailureReason failureReason) {
        if (field == null) {
            slice = default;
            failureReason = MarkdownOriginalSourceSliceFailureReason.SourceSpanUnavailable;
            return false;
        }

        return TryCreateOriginalSourceSlice(field.SourceSpan, out slice, out failureReason);
    }

    /// <summary>Creates a non-mutating source edit that replaces an abbreviation definition.</summary>
    public MarkdownNativeSourceEdit CreateReplaceEdit(MarkdownAbbreviationDefinition abbreviationDefinition, string replacementMarkdown) {
        if (abbreviationDefinition == null) {
            throw new ArgumentNullException(nameof(abbreviationDefinition));
        }

        if (!abbreviationDefinition.SourceSpan.HasValue) {
            throw new InvalidOperationException("The abbreviation definition does not have a source span.");
        }

        return CreateReplaceEdit(abbreviationDefinition.SourceSpan.Value, replacementMarkdown);
    }

    /// <summary>Creates a non-mutating source edit that replaces an abbreviation definition field.</summary>
    public MarkdownNativeSourceEdit CreateReplaceEdit(MarkdownNativeAbbreviationDefinitionField field, string replacementMarkdown) {
        if (field == null) {
            throw new ArgumentNullException(nameof(field));
        }

        return CreateReplaceEdit(field.SourceSpan, replacementMarkdown);
    }

    internal static IEnumerable<MarkdownNativeAbbreviationDefinitionField> EnumerateAbbreviationDefinitionFields(MarkdownAbbreviationDefinition definition) {
        if (definition.OpeningMarkerSourceSpan.HasValue) {
            yield return new MarkdownNativeAbbreviationDefinitionField("openingMarker", "*[", definition.OpeningMarkerSourceSpan.Value, definition);
        }

        if (definition.LabelSourceSpan.HasValue) {
            yield return new MarkdownNativeAbbreviationDefinitionField("label", definition.Label, definition.LabelSourceSpan.Value, definition);
        }

        if (definition.SeparatorMarkerSourceSpan.HasValue) {
            yield return new MarkdownNativeAbbreviationDefinitionField("separatorMarker", "]:", definition.SeparatorMarkerSourceSpan.Value, definition);
        }

        if (definition.TitleSourceSpan.HasValue) {
            yield return new MarkdownNativeAbbreviationDefinitionField("title", definition.Title, definition.TitleSourceSpan.Value, definition);
        }
    }

    private static bool AbbreviationDefinitionFieldsContainPosition(MarkdownAbbreviationDefinition definition, int lineNumber, int columnNumber) {
        foreach (var field in EnumerateAbbreviationDefinitionFields(definition)) {
            if (field.SourceSpan.ContainsPosition(lineNumber, columnNumber)) {
                return true;
            }
        }

        return false;
    }
}
