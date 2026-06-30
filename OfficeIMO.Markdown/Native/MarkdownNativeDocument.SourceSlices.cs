namespace OfficeIMO.Markdown;

public sealed partial class MarkdownNativeDocument {
    /// <summary>
    /// Creates a source slice over the normalized markdown text that backs the supplied source span.
    /// </summary>
    public bool TryCreateSourceSlice(MarkdownSourceSpan sourceSpan, out MarkdownSourceSlice slice) =>
        ParseResult.TryCreateSourceSlice(sourceSpan, out slice);

    private bool TryCreateSourceSlice(MarkdownSourceSpan? sourceSpan, out MarkdownSourceSlice slice) {
        if (!sourceSpan.HasValue) {
            slice = default;
            return false;
        }

        return TryCreateSourceSlice(sourceSpan.Value, out slice);
    }

    /// <summary>
    /// Creates a source slice over the normalized markdown text that backs a native block.
    /// </summary>
    public bool TryCreateSourceSlice(MarkdownNativeBlock block, out MarkdownSourceSlice slice) {
        if (block == null) {
            slice = default;
            return false;
        }

        return TryCreateSourceSlice(block.SourceSpan, out slice);
    }

    /// <summary>
    /// Creates a source slice over the normalized markdown text that backs a native block source field.
    /// </summary>
    public bool TryCreateSourceSlice(MarkdownNativeBlockSourceField field, out MarkdownSourceSlice slice) {
        if (field == null) {
            slice = default;
            return false;
        }

        return TryCreateSourceSlice(field.SourceSpan, out slice);
    }

    /// <summary>
    /// Creates a source slice over the normalized markdown text that backs document-level source trivia.
    /// </summary>
    public bool TryCreateSourceSlice(MarkdownNativeSourceTrivia trivia, out MarkdownSourceSlice slice) {
        if (trivia == null) {
            slice = default;
            return false;
        }

        if (trivia.SourceSpan.StartOffset.HasValue && trivia.SourceSpan.EndOffset.HasValue) {
            return MarkdownSourceSlice.TryCreateFromOffsets(
                SourceMarkdown,
                trivia.SourceSpan,
                MarkdownSourceTextKind.Normalized,
                trivia.SourceSpan.StartOffset.Value,
                trivia.SourceSpan.EndOffset.Value,
                out slice);
        }

        return TryCreateSourceSlice(trivia.SourceSpan, out slice);
    }

    /// <summary>
    /// Creates a source slice over the normalized markdown text that backs a native inline.
    /// </summary>
    public bool TryCreateSourceSlice(MarkdownNativeInline inline, out MarkdownSourceSlice slice) {
        if (inline == null || !inline.SourceSpan.HasValue) {
            slice = default;
            return false;
        }

        return TryCreateSourceSlice(inline.SourceSpan.Value, out slice);
    }

    /// <summary>
    /// Creates a source slice over the normalized markdown text that backs source-backed inline metadata.
    /// </summary>
    public bool TryCreateSourceSlice(MarkdownNativeInlineMetadata metadata, out MarkdownSourceSlice slice) {
        if (metadata == null || !metadata.SourceSpan.HasValue) {
            slice = default;
            return false;
        }

        return TryCreateSourceSlice(metadata.SourceSpan.Value, out slice);
    }

    /// <summary>
    /// Creates a source slice over the normalized markdown text that backs the source-backed content span of a native list item.
    /// </summary>
    public bool TryCreateSourceSlice(MarkdownNativeListItem listItem, out MarkdownSourceSlice slice) {
        if (listItem == null) {
            slice = default;
            return false;
        }

        return TryCreateSourceSlice(listItem.ContentSourceSpan, out slice);
    }

    /// <summary>
    /// Creates a source slice over the normalized markdown text that backs a paragraph owned by a native list item.
    /// </summary>
    public bool TryCreateSourceSlice(MarkdownNativeListItemParagraph paragraph, out MarkdownSourceSlice slice) {
        if (paragraph == null) {
            slice = default;
            return false;
        }

        return TryCreateSourceSlice(paragraph.SourceSpan, out slice);
    }

    /// <summary>
    /// Creates a source slice over the normalized markdown text that backs a reference-style link definition.
    /// </summary>
    public bool TryCreateSourceSlice(MarkdownReferenceLinkDefinition referenceDefinition, out MarkdownSourceSlice slice) {
        if (referenceDefinition == null) {
            slice = default;
            return false;
        }

        return TryCreateSourceSlice(referenceDefinition.SourceSpan, out slice);
    }

    /// <summary>
    /// Creates a source slice over the normalized markdown text that backs a reference-style link definition field.
    /// </summary>
    public bool TryCreateSourceSlice(MarkdownNativeReferenceLinkDefinitionField field, out MarkdownSourceSlice slice) {
        if (field == null) {
            slice = default;
            return false;
        }

        return TryCreateSourceSlice(field.SourceSpan, out slice);
    }

    /// <summary>
    /// Creates a source slice over the normalized markdown text that backs a native table cell.
    /// </summary>
    public bool TryCreateSourceSlice(MarkdownNativeTableCell tableCell, out MarkdownSourceSlice slice) {
        if (tableCell == null) {
            slice = default;
            return false;
        }

        return TryCreateSourceSlice(tableCell.SourceSpan, out slice);
    }

    /// <summary>
    /// Creates a source slice over the normalized markdown text that backs a native definition-list group.
    /// </summary>
    public bool TryCreateSourceSlice(MarkdownNativeDefinitionListGroup definitionGroup, out MarkdownSourceSlice slice) {
        if (definitionGroup == null) {
            slice = default;
            return false;
        }

        return TryCreateSourceSlice(definitionGroup.SourceSpan, out slice);
    }

    /// <summary>
    /// Creates a source slice over the normalized markdown text that backs a native definition-list term.
    /// </summary>
    public bool TryCreateSourceSlice(MarkdownNativeDefinitionListTerm definitionTerm, out MarkdownSourceSlice slice) {
        if (definitionTerm == null) {
            slice = default;
            return false;
        }

        return TryCreateSourceSlice(definitionTerm.SourceSpan, out slice);
    }

    /// <summary>
    /// Creates a source slice over the normalized markdown text that backs a native definition-list definition body.
    /// </summary>
    public bool TryCreateSourceSlice(MarkdownNativeDefinitionListDefinition definition, out MarkdownSourceSlice slice) {
        if (definition == null) {
            slice = default;
            return false;
        }

        return TryCreateSourceSlice(definition.SourceSpan, out slice);
    }

    /// <summary>
    /// Creates a source slice over the original reader input that backs the supplied source span when trivia was preserved.
    /// </summary>
    public bool TryCreateOriginalSourceSlice(MarkdownSourceSpan sourceSpan, out MarkdownSourceSlice slice) =>
        ParseResult.TryCreateOriginalSourceSlice(sourceSpan, out slice);

    /// <summary>
    /// Creates a source slice over the original reader input that backs the supplied source span when trivia was preserved.
    /// </summary>
    public bool TryCreateOriginalSourceSlice(
        MarkdownSourceSpan sourceSpan,
        out MarkdownSourceSlice slice,
        out MarkdownOriginalSourceSliceFailureReason failureReason) =>
        ParseResult.TryCreateOriginalSourceSlice(sourceSpan, out slice, out failureReason);

    private bool TryCreateOriginalSourceSlice(
        MarkdownSourceSpan? sourceSpan,
        out MarkdownSourceSlice slice,
        out MarkdownOriginalSourceSliceFailureReason failureReason) {
        if (!sourceSpan.HasValue) {
            slice = default;
            failureReason = MarkdownOriginalSourceSliceFailureReason.SourceSpanUnavailable;
            return false;
        }

        return TryCreateOriginalSourceSlice(sourceSpan.Value, out slice, out failureReason);
    }

    /// <summary>
    /// Creates a source slice over the original reader input that backs a native block when trivia was preserved.
    /// </summary>
    public bool TryCreateOriginalSourceSlice(MarkdownNativeBlock block, out MarkdownSourceSlice slice) {
        return TryCreateOriginalSourceSlice(block, out slice, out _);
    }

    /// <summary>
    /// Creates a source slice over the original reader input that backs a native block when trivia was preserved.
    /// </summary>
    public bool TryCreateOriginalSourceSlice(
        MarkdownNativeBlock block,
        out MarkdownSourceSlice slice,
        out MarkdownOriginalSourceSliceFailureReason failureReason) {
        if (block == null) {
            slice = default;
            failureReason = MarkdownOriginalSourceSliceFailureReason.SourceSpanUnavailable;
            return false;
        }

        if (block.SyntaxNode.IsGenerated) {
            return ParseResult.TryCreateOriginalSourceSlice(block.SyntaxNode, out slice, out failureReason);
        }

        return TryCreateOriginalSourceSlice(block.SourceSpan, out slice, out failureReason);
    }

    /// <summary>
    /// Creates a source slice over the original reader input that backs a native block source field when trivia was preserved.
    /// </summary>
    public bool TryCreateOriginalSourceSlice(MarkdownNativeBlockSourceField field, out MarkdownSourceSlice slice) {
        return TryCreateOriginalSourceSlice(field, out slice, out _);
    }

    /// <summary>
    /// Creates a source slice over the original reader input that backs a native block source field when trivia was preserved.
    /// </summary>
    public bool TryCreateOriginalSourceSlice(
        MarkdownNativeBlockSourceField field,
        out MarkdownSourceSlice slice,
        out MarkdownOriginalSourceSliceFailureReason failureReason) {
        if (field == null) {
            slice = default;
            failureReason = MarkdownOriginalSourceSliceFailureReason.SourceSpanUnavailable;
            return false;
        }

        if (field.Block.SyntaxNode.IsGenerated) {
            return ParseResult.TryCreateOriginalSourceSlice(field.Block.SyntaxNode, out slice, out failureReason);
        }

        return TryCreateOriginalSourceSlice(field.SourceSpan, out slice, out failureReason);
    }

    /// <summary>
    /// Creates a source slice over the original reader input that backs document-level source trivia when trivia was preserved.
    /// </summary>
    public bool TryCreateOriginalSourceSlice(MarkdownNativeSourceTrivia trivia, out MarkdownSourceSlice slice) {
        return TryCreateOriginalSourceSlice(trivia, out slice, out _);
    }

    /// <summary>
    /// Creates a source slice over the original reader input that backs document-level source trivia when trivia was preserved.
    /// </summary>
    public bool TryCreateOriginalSourceSlice(
        MarkdownNativeSourceTrivia trivia,
        out MarkdownSourceSlice slice,
        out MarkdownOriginalSourceSliceFailureReason failureReason) {
        if (trivia == null) {
            slice = default;
            failureReason = MarkdownOriginalSourceSliceFailureReason.SourceSpanUnavailable;
            return false;
        }

        return TryCreateOriginalSourceSlice(trivia.SourceSpan, out slice, out failureReason);
    }

    /// <summary>
    /// Creates a source slice over the original reader input that backs a native inline when trivia was preserved.
    /// </summary>
    public bool TryCreateOriginalSourceSlice(MarkdownNativeInline inline, out MarkdownSourceSlice slice) {
        return TryCreateOriginalSourceSlice(inline, out slice, out _);
    }

    /// <summary>
    /// Creates a source slice over the original reader input that backs a native inline when trivia was preserved.
    /// </summary>
    public bool TryCreateOriginalSourceSlice(
        MarkdownNativeInline inline,
        out MarkdownSourceSlice slice,
        out MarkdownOriginalSourceSliceFailureReason failureReason) {
        if (inline == null || !inline.SourceSpan.HasValue) {
            slice = default;
            failureReason = MarkdownOriginalSourceSliceFailureReason.SourceSpanUnavailable;
            return false;
        }

        return TryCreateOriginalSourceSlice(inline.SourceSpan.Value, out slice, out failureReason);
    }

    /// <summary>
    /// Creates a source slice over the original reader input that backs source-backed inline metadata when trivia was preserved.
    /// </summary>
    public bool TryCreateOriginalSourceSlice(MarkdownNativeInlineMetadata metadata, out MarkdownSourceSlice slice) {
        return TryCreateOriginalSourceSlice(metadata, out slice, out _);
    }

    /// <summary>
    /// Creates a source slice over the original reader input that backs source-backed inline metadata when trivia was preserved.
    /// </summary>
    public bool TryCreateOriginalSourceSlice(
        MarkdownNativeInlineMetadata metadata,
        out MarkdownSourceSlice slice,
        out MarkdownOriginalSourceSliceFailureReason failureReason) {
        if (metadata == null || !metadata.SourceSpan.HasValue) {
            slice = default;
            failureReason = MarkdownOriginalSourceSliceFailureReason.SourceSpanUnavailable;
            return false;
        }

        return TryCreateOriginalSourceSlice(metadata.SourceSpan.Value, out slice, out failureReason);
    }

    /// <summary>
    /// Creates a source slice over the original reader input that backs the source-backed content span of a native list item when trivia was preserved.
    /// </summary>
    public bool TryCreateOriginalSourceSlice(MarkdownNativeListItem listItem, out MarkdownSourceSlice slice) {
        return TryCreateOriginalSourceSlice(listItem, out slice, out _);
    }

    /// <summary>
    /// Creates a source slice over the original reader input that backs the source-backed content span of a native list item when trivia was preserved.
    /// </summary>
    public bool TryCreateOriginalSourceSlice(
        MarkdownNativeListItem listItem,
        out MarkdownSourceSlice slice,
        out MarkdownOriginalSourceSliceFailureReason failureReason) {
        if (listItem == null) {
            slice = default;
            failureReason = MarkdownOriginalSourceSliceFailureReason.SourceSpanUnavailable;
            return false;
        }

        if (listItem.SyntaxNode.IsGenerated) {
            return ParseResult.TryCreateOriginalSourceSlice(listItem.SyntaxNode, out slice, out failureReason);
        }

        return TryCreateOriginalSourceSlice(listItem.ContentSourceSpan, out slice, out failureReason);
    }

    /// <summary>
    /// Creates a source slice over the original reader input that backs a paragraph owned by a native list item when trivia was preserved.
    /// </summary>
    public bool TryCreateOriginalSourceSlice(MarkdownNativeListItemParagraph paragraph, out MarkdownSourceSlice slice) {
        return TryCreateOriginalSourceSlice(paragraph, out slice, out _);
    }

    /// <summary>
    /// Creates a source slice over the original reader input that backs a paragraph owned by a native list item when trivia was preserved.
    /// </summary>
    public bool TryCreateOriginalSourceSlice(
        MarkdownNativeListItemParagraph paragraph,
        out MarkdownSourceSlice slice,
        out MarkdownOriginalSourceSliceFailureReason failureReason) {
        if (paragraph == null) {
            slice = default;
            failureReason = MarkdownOriginalSourceSliceFailureReason.SourceSpanUnavailable;
            return false;
        }

        if (paragraph.SyntaxNode?.IsGenerated == true) {
            return ParseResult.TryCreateOriginalSourceSlice(paragraph.SyntaxNode, out slice, out failureReason);
        }

        return TryCreateOriginalSourceSlice(paragraph.SourceSpan, out slice, out failureReason);
    }

    /// <summary>
    /// Creates a source slice over the original reader input that backs a reference-style link definition when trivia was preserved.
    /// </summary>
    public bool TryCreateOriginalSourceSlice(MarkdownReferenceLinkDefinition referenceDefinition, out MarkdownSourceSlice slice) {
        return TryCreateOriginalSourceSlice(referenceDefinition, out slice, out _);
    }

    /// <summary>
    /// Creates a source slice over the original reader input that backs a reference-style link definition when trivia was preserved.
    /// </summary>
    public bool TryCreateOriginalSourceSlice(
        MarkdownReferenceLinkDefinition referenceDefinition,
        out MarkdownSourceSlice slice,
        out MarkdownOriginalSourceSliceFailureReason failureReason) {
        if (referenceDefinition == null) {
            slice = default;
            failureReason = MarkdownOriginalSourceSliceFailureReason.SourceSpanUnavailable;
            return false;
        }

        return TryCreateOriginalSourceSlice(referenceDefinition.SourceSpan, out slice, out failureReason);
    }

    /// <summary>
    /// Creates a source slice over the original reader input that backs a reference-style link definition field when trivia was preserved.
    /// </summary>
    public bool TryCreateOriginalSourceSlice(MarkdownNativeReferenceLinkDefinitionField field, out MarkdownSourceSlice slice) {
        return TryCreateOriginalSourceSlice(field, out slice, out _);
    }

    /// <summary>
    /// Creates a source slice over the original reader input that backs a reference-style link definition field when trivia was preserved.
    /// </summary>
    public bool TryCreateOriginalSourceSlice(
        MarkdownNativeReferenceLinkDefinitionField field,
        out MarkdownSourceSlice slice,
        out MarkdownOriginalSourceSliceFailureReason failureReason) {
        if (field == null) {
            slice = default;
            failureReason = MarkdownOriginalSourceSliceFailureReason.SourceSpanUnavailable;
            return false;
        }

        return TryCreateOriginalSourceSlice(field.SourceSpan, out slice, out failureReason);
    }

    /// <summary>
    /// Creates a source slice over the original reader input that backs a native table cell when trivia was preserved.
    /// </summary>
    public bool TryCreateOriginalSourceSlice(MarkdownNativeTableCell tableCell, out MarkdownSourceSlice slice) {
        return TryCreateOriginalSourceSlice(tableCell, out slice, out _);
    }

    /// <summary>
    /// Creates a source slice over the original reader input that backs a native table cell when trivia was preserved.
    /// </summary>
    public bool TryCreateOriginalSourceSlice(
        MarkdownNativeTableCell tableCell,
        out MarkdownSourceSlice slice,
        out MarkdownOriginalSourceSliceFailureReason failureReason) {
        if (tableCell == null) {
            slice = default;
            failureReason = MarkdownOriginalSourceSliceFailureReason.SourceSpanUnavailable;
            return false;
        }

        if (tableCell.SyntaxNode?.IsGenerated == true) {
            return ParseResult.TryCreateOriginalSourceSlice(tableCell.SyntaxNode, out slice, out failureReason);
        }

        return TryCreateOriginalSourceSlice(tableCell.SourceSpan, out slice, out failureReason);
    }

    /// <summary>
    /// Creates a source slice over the original reader input that backs a native definition-list group when trivia was preserved.
    /// </summary>
    public bool TryCreateOriginalSourceSlice(MarkdownNativeDefinitionListGroup definitionGroup, out MarkdownSourceSlice slice) {
        return TryCreateOriginalSourceSlice(definitionGroup, out slice, out _);
    }

    /// <summary>
    /// Creates a source slice over the original reader input that backs a native definition-list group when trivia was preserved.
    /// </summary>
    public bool TryCreateOriginalSourceSlice(
        MarkdownNativeDefinitionListGroup definitionGroup,
        out MarkdownSourceSlice slice,
        out MarkdownOriginalSourceSliceFailureReason failureReason) {
        if (definitionGroup == null) {
            slice = default;
            failureReason = MarkdownOriginalSourceSliceFailureReason.SourceSpanUnavailable;
            return false;
        }

        if (definitionGroup.SyntaxNode?.IsGenerated == true) {
            return ParseResult.TryCreateOriginalSourceSlice(definitionGroup.SyntaxNode, out slice, out failureReason);
        }

        return TryCreateOriginalSourceSlice(definitionGroup.SourceSpan, out slice, out failureReason);
    }

    /// <summary>
    /// Creates a source slice over the original reader input that backs a native definition-list term when trivia was preserved.
    /// </summary>
    public bool TryCreateOriginalSourceSlice(MarkdownNativeDefinitionListTerm definitionTerm, out MarkdownSourceSlice slice) {
        return TryCreateOriginalSourceSlice(definitionTerm, out slice, out _);
    }

    /// <summary>
    /// Creates a source slice over the original reader input that backs a native definition-list term when trivia was preserved.
    /// </summary>
    public bool TryCreateOriginalSourceSlice(
        MarkdownNativeDefinitionListTerm definitionTerm,
        out MarkdownSourceSlice slice,
        out MarkdownOriginalSourceSliceFailureReason failureReason) {
        if (definitionTerm == null) {
            slice = default;
            failureReason = MarkdownOriginalSourceSliceFailureReason.SourceSpanUnavailable;
            return false;
        }

        if (definitionTerm.SyntaxNode?.IsGenerated == true) {
            return ParseResult.TryCreateOriginalSourceSlice(definitionTerm.SyntaxNode, out slice, out failureReason);
        }

        return TryCreateOriginalSourceSlice(definitionTerm.SourceSpan, out slice, out failureReason);
    }

    /// <summary>
    /// Creates a source slice over the original reader input that backs a native definition-list definition body when trivia was preserved.
    /// </summary>
    public bool TryCreateOriginalSourceSlice(MarkdownNativeDefinitionListDefinition definition, out MarkdownSourceSlice slice) {
        return TryCreateOriginalSourceSlice(definition, out slice, out _);
    }

    /// <summary>
    /// Creates a source slice over the original reader input that backs a native definition-list definition body when trivia was preserved.
    /// </summary>
    public bool TryCreateOriginalSourceSlice(
        MarkdownNativeDefinitionListDefinition definition,
        out MarkdownSourceSlice slice,
        out MarkdownOriginalSourceSliceFailureReason failureReason) {
        if (definition == null) {
            slice = default;
            failureReason = MarkdownOriginalSourceSliceFailureReason.SourceSpanUnavailable;
            return false;
        }

        if (definition.SyntaxNode?.IsGenerated == true) {
            return ParseResult.TryCreateOriginalSourceSlice(definition.SyntaxNode, out slice, out failureReason);
        }

        return TryCreateOriginalSourceSlice(definition.SourceSpan, out slice, out failureReason);
    }
}
