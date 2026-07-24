using AngleSharp.Dom;
using OfficeIMO.Markdown;

namespace OfficeIMO.Markdown.Html;

internal sealed partial class HtmlToMarkdownConverter {
    private static DetailsBlock ConvertDetailsElement(IElement element, ConversionContext context) {
        SummaryBlock? summary = null;
        var summaryElement = element.Children.FirstOrDefault(child => HasEffectiveTagName(child, context, "SUMMARY"));
        if (summaryElement != null) {
            summary = new SummaryBlock(NormalizeInlineSequenceForBlock(ConvertInlineNodesToInlineSequence(summaryElement.ChildNodes, context)));
        }

        var details = new DetailsBlock(summary, open: element.HasAttribute("open"));
        foreach (var child in element.ChildNodes) {
            if (ReferenceEquals(child, summaryElement)) {
                continue;
            }

            foreach (var block in ConvertNodesToBlocks(new[] { child }, context)) {
                details.ChildBlocks.Add(block);
            }
        }

        return details;
    }

    private static DefinitionListBlock ConvertDefinitionListElement(IElement element, ConversionContext context) {
        var list = new DefinitionListBlock();
        var pendingTerms = new List<InlineSequence>();
        bool hasDefinitionsForCurrentGroup = false;
        List<DefinitionListDefinition>? pendingDefinitions = null;

        void FlushPendingGroup() {
            if (pendingTerms.Count == 0 || pendingDefinitions == null || pendingDefinitions.Count == 0) {
                pendingTerms.Clear();
                pendingDefinitions = null;
                hasDefinitionsForCurrentGroup = false;
                return;
            }

            long groupExpansionCount = (long)pendingTerms.Count * pendingDefinitions.Count;
            long totalExpansionCount = context.DefinitionListEntryExpansionCount + groupExpansionCount;
            if (totalExpansionCount > context.Options.MaxDefinitionListEntryExpansions) {
                throw new InvalidOperationException(
                    $"HTML definition lists exceed the configured entry expansion limit of {context.Options.MaxDefinitionListEntryExpansions}.");
            }

            list.AddGroup(new DefinitionListGroup(pendingTerms, pendingDefinitions));
            context.DefinitionListEntryExpansionCount = (int)totalExpansionCount;
            pendingTerms.Clear();
            pendingDefinitions = null;
            hasDefinitionsForCurrentGroup = false;
        }

        foreach (var child in element.Children) {
            if (HasEffectiveTagName(child, context, "DT")) {
                if (hasDefinitionsForCurrentGroup) {
                    FlushPendingGroup();
                }

                var term = NormalizeInlineSequenceForBlock(ConvertInlineNodesToInlineSequence(child.ChildNodes, context));
                if (HasVisibleInlineContent(term)) {
                    pendingTerms.Add(term);
                }
                continue;
            }

            if (HasEffectiveTagName(child, context, "DD") && pendingTerms.Count > 0) {
                pendingDefinitions ??= new List<DefinitionListDefinition>();
                pendingDefinitions.Add(new DefinitionListDefinition(ConvertDefinitionValueToBlocks(child, context)));
                hasDefinitionsForCurrentGroup = true;
            }
        }

        FlushPendingGroup();
        return list;
    }

    private static IReadOnlyList<IMarkdownBlock> ConvertDefinitionValueToBlocks(IElement element, ConversionContext context) {
        if (HasDirectBlockChildren(element, context)) {
            return ConvertNodesToBlocks(element.ChildNodes, context);
        }

        var inlineSequence = NormalizeInlineSequenceForBlock(ConvertInlineNodesToInlineSequence(element.ChildNodes, context));
        if (!HasVisibleInlineContent(inlineSequence)) {
            return Array.Empty<IMarkdownBlock>();
        }

        return new IMarkdownBlock[] { new ParagraphBlock(inlineSequence) };
    }

    private static InlineSequence NormalizeInlineSequenceForBlock(InlineSequence? source) {
        return source ?? new InlineSequence { AutoSpacing = false };
    }

    private static bool HasVisibleInlineContent(InlineSequence? sequence) {
        if (sequence == null || sequence.Nodes.Count == 0) {
            return false;
        }

        foreach (var node in sequence.Nodes) {
            switch (node) {
                case null:
                    continue;
                case TextRun textRun when string.IsNullOrWhiteSpace(textRun.Text):
                    continue;
                default:
                    return true;
            }
        }

        return false;
    }
}
