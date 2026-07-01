namespace OfficeIMO.Markdown;

public sealed partial class MarkdownNativeDocument {
    /// <summary>
    /// Creates a source mapping for a source span, including normalized text and original text when it maps safely.
    /// </summary>
    public bool TryCreateSourceMapping(MarkdownSourceSpan sourceSpan, out MarkdownSourceMapping mapping) =>
        ParseResult.TryCreateSourceMapping(sourceSpan, out mapping);

    /// <summary>
    /// Creates a source mapping for a native block, including normalized text and original text when it maps safely.
    /// </summary>
    public bool TryCreateSourceMapping(MarkdownNativeBlock block, out MarkdownSourceMapping mapping) {
        if (block == null || !block.SourceSpan.HasValue) {
            mapping = default;
            return false;
        }

        return TryCreateSourceMapping(block.SourceSpan.Value, block.SyntaxNode, out mapping);
    }

    /// <summary>
    /// Creates a source mapping for a native block source field, including normalized text and original text when it maps safely.
    /// </summary>
    public bool TryCreateSourceMapping(MarkdownNativeBlockSourceField field, out MarkdownSourceMapping mapping) {
        if (field == null) {
            mapping = default;
            return false;
        }

        return TryCreateSourceMapping(field.SourceSpan, field.Block.SyntaxNode, out mapping);
    }

    /// <summary>
    /// Creates a source mapping for document-level source trivia, including normalized text and original text when it maps safely.
    /// </summary>
    public bool TryCreateSourceMapping(MarkdownNativeSourceTrivia trivia, out MarkdownSourceMapping mapping) {
        if (trivia == null) {
            mapping = default;
            return false;
        }

        return TryCreateSourceMapping(trivia.SourceSpan, syntaxNode: null, out mapping);
    }

    /// <summary>
    /// Creates a source mapping for a native inline, including normalized text and original text when it maps safely.
    /// </summary>
    public bool TryCreateSourceMapping(MarkdownNativeInline inline, out MarkdownSourceMapping mapping) {
        if (inline == null || !inline.SourceSpan.HasValue) {
            mapping = default;
            return false;
        }

        return TryCreateSourceMapping(inline.SourceSpan.Value, inline.SyntaxNode, out mapping);
    }

    /// <summary>
    /// Creates a source mapping for source-backed inline metadata, including normalized text and original text when it maps safely.
    /// </summary>
    public bool TryCreateSourceMapping(MarkdownNativeInlineMetadata metadata, out MarkdownSourceMapping mapping) {
        if (metadata == null || !metadata.SourceSpan.HasValue) {
            mapping = default;
            return false;
        }

        return TryCreateSourceMapping(metadata.SourceSpan.Value, metadata.SyntaxNode, out mapping);
    }

    /// <summary>
    /// Creates a source mapping for a native list item content span, including normalized text and original text when it maps safely.
    /// </summary>
    public bool TryCreateSourceMapping(MarkdownNativeListItem listItem, out MarkdownSourceMapping mapping) {
        if (listItem == null || !listItem.ContentSourceSpan.HasValue) {
            mapping = default;
            return false;
        }

        return TryCreateSourceMapping(listItem.ContentSourceSpan.Value, listItem.SyntaxNode, out mapping);
    }

    /// <summary>
    /// Creates a source mapping for a paragraph owned by a native list item, including normalized text and original text when it maps safely.
    /// </summary>
    public bool TryCreateSourceMapping(MarkdownNativeListItemParagraph paragraph, out MarkdownSourceMapping mapping) {
        if (paragraph == null || !paragraph.SourceSpan.HasValue) {
            mapping = default;
            return false;
        }

        return TryCreateSourceMapping(paragraph.SourceSpan.Value, paragraph.SyntaxNode, out mapping);
    }

    /// <summary>
    /// Creates a source mapping for a reference-style link definition, including normalized text and original text when it maps safely.
    /// </summary>
    public bool TryCreateSourceMapping(MarkdownReferenceLinkDefinition referenceDefinition, out MarkdownSourceMapping mapping) {
        if (referenceDefinition == null || !referenceDefinition.SourceSpan.HasValue) {
            mapping = default;
            return false;
        }

        return TryCreateSourceMapping(referenceDefinition.SourceSpan.Value, syntaxNode: null, out mapping);
    }

    /// <summary>
    /// Creates a source mapping for a reference-style link definition field, including normalized text and original text when it maps safely.
    /// </summary>
    public bool TryCreateSourceMapping(MarkdownNativeReferenceLinkDefinitionField field, out MarkdownSourceMapping mapping) {
        if (field == null) {
            mapping = default;
            return false;
        }

        return TryCreateSourceMapping(field.SourceSpan, syntaxNode: null, out mapping);
    }

    /// <summary>
    /// Creates a source mapping for an abbreviation definition, including normalized text and original text when it maps safely.
    /// </summary>
    public bool TryCreateSourceMapping(MarkdownAbbreviationDefinition abbreviationDefinition, out MarkdownSourceMapping mapping) {
        if (abbreviationDefinition == null || !abbreviationDefinition.SourceSpan.HasValue) {
            mapping = default;
            return false;
        }

        return TryCreateSourceMapping(abbreviationDefinition.SourceSpan.Value, syntaxNode: null, out mapping);
    }

    /// <summary>
    /// Creates a source mapping for an abbreviation definition field, including normalized text and original text when it maps safely.
    /// </summary>
    public bool TryCreateSourceMapping(MarkdownNativeAbbreviationDefinitionField field, out MarkdownSourceMapping mapping) {
        if (field == null) {
            mapping = default;
            return false;
        }

        return TryCreateSourceMapping(field.SourceSpan, syntaxNode: null, out mapping);
    }

    /// <summary>
    /// Creates a source mapping for a native table cell, including normalized text and original text when it maps safely.
    /// </summary>
    public bool TryCreateSourceMapping(MarkdownNativeTableCell tableCell, out MarkdownSourceMapping mapping) {
        if (tableCell == null || !tableCell.SourceSpan.HasValue) {
            mapping = default;
            return false;
        }

        return TryCreateSourceMapping(tableCell.SourceSpan.Value, tableCell.SyntaxNode, out mapping);
    }

    /// <summary>
    /// Creates a source mapping for a native table row, including normalized text and original text when it maps safely.
    /// </summary>
    public bool TryCreateSourceMapping(MarkdownNativeTableRow tableRow, out MarkdownSourceMapping mapping) {
        if (tableRow == null || !tableRow.SourceSpan.HasValue) {
            mapping = default;
            return false;
        }

        return TryCreateSourceMapping(tableRow.SourceSpan.Value, tableRow.SyntaxNode, out mapping);
    }

    /// <summary>
    /// Creates a source mapping for a native definition-list group, including normalized text and original text when it maps safely.
    /// </summary>
    public bool TryCreateSourceMapping(MarkdownNativeDefinitionListGroup definitionGroup, out MarkdownSourceMapping mapping) {
        if (definitionGroup == null || !definitionGroup.SourceSpan.HasValue) {
            mapping = default;
            return false;
        }

        return TryCreateSourceMapping(definitionGroup.SourceSpan.Value, definitionGroup.SyntaxNode, out mapping);
    }

    /// <summary>
    /// Creates a source mapping for a native definition-list term, including normalized text and original text when it maps safely.
    /// </summary>
    public bool TryCreateSourceMapping(MarkdownNativeDefinitionListTerm definitionTerm, out MarkdownSourceMapping mapping) {
        if (definitionTerm == null || !definitionTerm.SourceSpan.HasValue) {
            mapping = default;
            return false;
        }

        return TryCreateSourceMapping(definitionTerm.SourceSpan.Value, definitionTerm.SyntaxNode, out mapping);
    }

    /// <summary>
    /// Creates a source mapping for a native definition-list definition body, including normalized text and original text when it maps safely.
    /// </summary>
    public bool TryCreateSourceMapping(MarkdownNativeDefinitionListDefinition definition, out MarkdownSourceMapping mapping) {
        if (definition == null || !definition.SourceSpan.HasValue) {
            mapping = default;
            return false;
        }

        return TryCreateSourceMapping(definition.SourceSpan.Value, definition.SyntaxNode, out mapping);
    }

    private bool TryCreateSourceMapping(
        MarkdownSourceSpan sourceSpan,
        MarkdownSyntaxNode? syntaxNode,
        out MarkdownSourceMapping mapping) {
        if (!ParseResult.TryCreateSourceMapping(sourceSpan, out mapping)) {
            return false;
        }

        if (syntaxNode?.IsGenerated == true) {
            mapping = mapping.WithOriginalFailure(MarkdownOriginalSourceSliceFailureReason.GeneratedSyntaxNode);
        }

        return true;
    }
}
