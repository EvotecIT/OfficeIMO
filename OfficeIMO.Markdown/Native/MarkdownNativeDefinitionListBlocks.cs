namespace OfficeIMO.Markdown;

/// <summary>
/// Native projection for a definition list block.
/// </summary>
public sealed class MarkdownNativeDefinitionListBlock : MarkdownNativeBlock {
    internal MarkdownNativeDefinitionListBlock(
        DefinitionListBlock definitionList,
        MarkdownSyntaxNode syntaxNode,
        ICollection<MarkdownNativeDiagnostic> diagnostics)
        : base(MarkdownNativeBlockKind.DefinitionList, definitionList, syntaxNode) {
        DefinitionList = definitionList;
        Groups = BuildGroups(definitionList, syntaxNode, diagnostics);
        Children = Groups
            .SelectMany(static group => group.Definitions)
            .SelectMany(static definition => definition.Children)
            .ToArray();
    }

    /// <summary>Source definition list block.</summary>
    public DefinitionListBlock DefinitionList { get; }

    /// <summary>Definition groups in document order.</summary>
    public IReadOnlyList<MarkdownNativeDefinitionListGroup> Groups { get; }

    /// <summary>Flattened native definition body children in document order.</summary>
    public IReadOnlyList<MarkdownNativeBlock> Children { get; }

    private static IReadOnlyList<MarkdownNativeDefinitionListGroup> BuildGroups(
        DefinitionListBlock definitionList,
        MarkdownSyntaxNode syntaxNode,
        ICollection<MarkdownNativeDiagnostic> diagnostics) {
        if (definitionList.Groups.Count == 0) {
            return Array.Empty<MarkdownNativeDefinitionListGroup>();
        }

        var groupNodes = syntaxNode.Children
            .Where(static child => child.Kind == MarkdownSyntaxKind.DefinitionGroup)
            .ToArray();
        var groups = new List<MarkdownNativeDefinitionListGroup>(definitionList.Groups.Count);
        for (int i = 0; i < definitionList.Groups.Count; i++) {
            var group = definitionList.Groups[i];
            var groupNode = i < groupNodes.Length ? groupNodes[i] : null;
            groups.Add(new MarkdownNativeDefinitionListGroup(group, groupNode, diagnostics));
        }

        return groups;
    }
}

/// <summary>
/// Native projection for a grouped definition-list item.
/// </summary>
public sealed class MarkdownNativeDefinitionListGroup {
    internal MarkdownNativeDefinitionListGroup(
        DefinitionListGroup group,
        MarkdownSyntaxNode? syntaxNode,
        ICollection<MarkdownNativeDiagnostic> diagnostics) {
        Group = group ?? throw new ArgumentNullException(nameof(group));
        SyntaxNode = syntaxNode;
        SourceSpan = syntaxNode?.SourceSpan ?? group.SourceSpan;
        Terms = BuildTerms(group, syntaxNode);
        Definitions = BuildDefinitions(group, syntaxNode, diagnostics);
    }

    /// <summary>Source semantic definition-list group.</summary>
    public DefinitionListGroup Group { get; }

    /// <summary>Syntax node that produced this group when available.</summary>
    public MarkdownSyntaxNode? SyntaxNode { get; }

    /// <summary>Source span in the normalized markdown text when available.</summary>
    public MarkdownSourceSpan? SourceSpan { get; }

    /// <summary>Terms in this definition group.</summary>
    public IReadOnlyList<MarkdownNativeDefinitionListTerm> Terms { get; }

    /// <summary>Definitions shared by the terms in this group.</summary>
    public IReadOnlyList<MarkdownNativeDefinitionListDefinition> Definitions { get; }

    private static IReadOnlyList<MarkdownNativeDefinitionListTerm> BuildTerms(
        DefinitionListGroup group,
        MarkdownSyntaxNode? syntaxNode) {
        if (group.TermItems.Count == 0) {
            return Array.Empty<MarkdownNativeDefinitionListTerm>();
        }

        var termNodes = syntaxNode?.Children
            .Where(static child => child.Kind == MarkdownSyntaxKind.DefinitionTerm)
            .ToArray() ?? Array.Empty<MarkdownSyntaxNode>();
        var terms = new List<MarkdownNativeDefinitionListTerm>(group.TermItems.Count);
        for (int i = 0; i < group.TermItems.Count; i++) {
            var termNode = i < termNodes.Length ? termNodes[i] : null;
            terms.Add(new MarkdownNativeDefinitionListTerm(group.TermItems[i], termNode));
        }

        return terms;
    }

    private static IReadOnlyList<MarkdownNativeDefinitionListDefinition> BuildDefinitions(
        DefinitionListGroup group,
        MarkdownSyntaxNode? syntaxNode,
        ICollection<MarkdownNativeDiagnostic> diagnostics) {
        if (group.Definitions.Count == 0) {
            return Array.Empty<MarkdownNativeDefinitionListDefinition>();
        }

        var definitionNodes = syntaxNode?.Children
            .Where(static child => child.Kind == MarkdownSyntaxKind.DefinitionValue)
            .ToArray() ?? Array.Empty<MarkdownSyntaxNode>();
        var definitions = new List<MarkdownNativeDefinitionListDefinition>(group.Definitions.Count);
        for (int i = 0; i < group.Definitions.Count; i++) {
            var definitionNode = i < definitionNodes.Length ? definitionNodes[i] : null;
            definitions.Add(new MarkdownNativeDefinitionListDefinition(group.Definitions[i], definitionNode, diagnostics));
        }

        return definitions;
    }
}

/// <summary>
/// Native projection for a definition-list term.
/// </summary>
public sealed class MarkdownNativeDefinitionListTerm {
    internal MarkdownNativeDefinitionListTerm(DefinitionListTerm term, MarkdownSyntaxNode? syntaxNode) {
        TermObject = term ?? new DefinitionListTerm();
        SyntaxNode = syntaxNode;
        SourceSpan = syntaxNode?.SourceSpan ?? TermObject.SourceSpan;
        Text = TermObject.Text;
        Markdown = TermObject.Markdown;
        InlineRuns = syntaxNode != null
            ? MarkdownNativeInlineProjection.FromInlineContainer(syntaxNode)
            : Array.Empty<MarkdownNativeInline>();
    }

    /// <summary>Source semantic term.</summary>
    public DefinitionListTerm TermObject { get; }

    /// <summary>Source term inline sequence.</summary>
    public InlineSequence Term => TermObject.Inlines;

    /// <summary>Source term inline sequence.</summary>
    public InlineSequence Inlines => TermObject.Inlines;

    /// <summary>Syntax node that produced this term when available.</summary>
    public MarkdownSyntaxNode? SyntaxNode { get; }

    /// <summary>Source span in the normalized markdown text when available.</summary>
    public MarkdownSourceSpan? SourceSpan { get; }

    /// <summary>Plain-text term content.</summary>
    public string Text { get; }

    /// <summary>Markdown term content.</summary>
    public string Markdown { get; }

    /// <summary>AST-backed native inline projection for the term.</summary>
    public IReadOnlyList<MarkdownNativeInline> InlineRuns { get; }
}

/// <summary>
/// Native projection for a definition-list definition body.
/// </summary>
public sealed class MarkdownNativeDefinitionListDefinition {
    internal MarkdownNativeDefinitionListDefinition(
        DefinitionListDefinition definition,
        MarkdownSyntaxNode? syntaxNode,
        ICollection<MarkdownNativeDiagnostic> diagnostics) {
        Definition = definition ?? new DefinitionListDefinition();
        SyntaxNode = syntaxNode;
        SourceSpan = syntaxNode?.SourceSpan ?? Definition.SourceSpan;
        Markdown = Definition.RenderMarkdown();
        BlankLineSourceSpans = Definition.BlankLineSourceSpans;
        ContinuationIndentSourceSpans = Definition.ContinuationIndentSourceSpans;
        Children = syntaxNode != null
            ? MarkdownNativeProjectionFactory.CreateChildren(syntaxNode, diagnostics)
            : Array.Empty<MarkdownNativeBlock>();
    }

    /// <summary>Source semantic definition body.</summary>
    public DefinitionListDefinition Definition { get; }

    /// <summary>Syntax node that produced this definition body when available.</summary>
    public MarkdownSyntaxNode? SyntaxNode { get; }

    /// <summary>Source span in the normalized markdown text when available.</summary>
    public MarkdownSourceSpan? SourceSpan { get; }

    /// <summary>Markdown representation of the definition body.</summary>
    public string Markdown { get; }

    /// <summary>Source spans for blank separator lines inside this definition body.</summary>
    public IReadOnlyList<MarkdownSourceSpan> BlankLineSourceSpans { get; }

    /// <summary>Source spans for indentation stripped from definition continuation lines.</summary>
    public IReadOnlyList<MarkdownSourceSpan> ContinuationIndentSourceSpans { get; }

    /// <summary>Native definition body children.</summary>
    public IReadOnlyList<MarkdownNativeBlock> Children { get; }
}
