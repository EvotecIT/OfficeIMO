using System.Threading;

namespace OfficeIMO.Latex;

internal sealed class LatexSemanticModel {
    internal LatexSemanticModel(
        IReadOnlyList<LatexCommand> commands,
        IReadOnlyList<LatexEnvironment> environments,
        IReadOnlyList<LatexMath> math,
        IReadOnlyList<LatexHeading> headings,
        IReadOnlyList<LatexParagraph> paragraphs,
        IReadOnlyList<LatexList> lists,
        IReadOnlyList<LatexFigure> figures,
        IReadOnlyList<LatexTable> tables,
        IReadOnlyList<LatexCitation> citations,
        IReadOnlyList<LatexReference> references,
        IReadOnlyList<LatexLabel> labels,
        IReadOnlyList<LatexTheorem> theorems,
        IReadOnlyList<LatexFootnote> footnotes,
        IReadOnlyList<LatexMacroDefinition> macroDefinitions) {
        Commands = commands;
        Environments = environments;
        Math = math;
        Headings = headings;
        Paragraphs = paragraphs;
        Lists = lists;
        Figures = figures;
        Tables = tables;
        Citations = citations;
        References = references;
        Labels = labels;
        Theorems = theorems;
        Footnotes = footnotes;
        MacroDefinitions = macroDefinitions;
    }

    internal IReadOnlyList<LatexCommand> Commands { get; }
    internal IReadOnlyList<LatexEnvironment> Environments { get; }
    internal IReadOnlyList<LatexMath> Math { get; }
    internal IReadOnlyList<LatexHeading> Headings { get; }
    internal IReadOnlyList<LatexParagraph> Paragraphs { get; }
    internal IReadOnlyList<LatexList> Lists { get; }
    internal IReadOnlyList<LatexFigure> Figures { get; }
    internal IReadOnlyList<LatexTable> Tables { get; }
    internal IReadOnlyList<LatexCitation> Citations { get; }
    internal IReadOnlyList<LatexReference> References { get; }
    internal IReadOnlyList<LatexLabel> Labels { get; }
    internal IReadOnlyList<LatexTheorem> Theorems { get; }
    internal IReadOnlyList<LatexFootnote> Footnotes { get; }
    internal IReadOnlyList<LatexMacroDefinition> MacroDefinitions { get; }
}

internal static partial class LatexSemanticBuilder {
    internal static LatexSemanticModel Build(
        LatexSourceText source,
        LatexSyntaxTree syntaxTree,
        LatexDocumentProfile profile,
        CancellationToken cancellationToken) {
        cancellationToken.ThrowIfCancellationRequested();
        var commandSyntax = new List<LatexSyntaxNode>();
        var environmentSyntax = new List<LatexSyntaxNode>();
        var mathSyntax = new List<LatexSyntaxNode>();
        foreach (LatexSyntaxNode node in syntaxTree.Root.DescendantsAndSelf()) {
            cancellationToken.ThrowIfCancellationRequested();
            switch (node.Kind) {
                case LatexSyntaxKind.Command: commandSyntax.Add(node); break;
                case LatexSyntaxKind.Environment: environmentSyntax.Add(node); break;
                case LatexSyntaxKind.Math: mathSyntax.Add(node); break;
            }
        }
        var commandMap = new Dictionary<LatexSyntaxNode, LatexCommand>();
        for (int index = 0; index < commandSyntax.Count; index++) {
            cancellationToken.ThrowIfCancellationRequested();
            commandMap[commandSyntax[index]] = new LatexCommand(commandSyntax[index], source);
        }

        var environments = new List<LatexEnvironment>(environmentSyntax.Count);
        for (int index = 0; index < environmentSyntax.Count; index++) {
            cancellationToken.ThrowIfCancellationRequested();
            LatexSyntaxNode syntax = environmentSyntax[index];
            LatexSyntaxNode beginSyntax = syntax.Children.First(static child => child.Kind == LatexSyntaxKind.Command);
            LatexSyntaxNode? endSyntax = syntax.Children.LastOrDefault(child =>
                child.Kind == LatexSyntaxKind.Command && string.Equals(child.Value, "end", StringComparison.Ordinal) &&
                string.Equals(commandMap[child].GetRequiredArgument(0)?.Content, syntax.Value, StringComparison.Ordinal));
            environments.Add(new LatexEnvironment(
                syntax,
                commandMap[beginSyntax],
                endSyntax == null ? null : commandMap[endSyntax],
                source));
        }

        var math = new List<LatexMath>(mathSyntax.Count);
        for (int index = 0; index < mathSyntax.Count; index++) math.Add(new LatexMath(mathSyntax[index], source));
        math.AddRange(environments.Where(static environment => environment.IsMath).Select(static environment => new LatexMath(environment)));

        LatexCommand[] commands = commandMap.Values.OrderBy(static command => command.Syntax.StartOffset).ToArray();
        LatexEnvironment[] orderedEnvironments = environments.OrderBy(static environment => environment.Syntax.StartOffset).ToArray();
        LatexMath[] orderedMath = math.OrderBy(static item => item.Syntax.StartOffset).ToArray();
        if (profile == LatexDocumentProfile.PreserveOnly) {
            return new LatexSemanticModel(
                commands,
                orderedEnvironments,
                orderedMath,
                Array.Empty<LatexHeading>(),
                Array.Empty<LatexParagraph>(),
                Array.Empty<LatexList>(),
                Array.Empty<LatexFigure>(),
                Array.Empty<LatexTable>(),
                Array.Empty<LatexCitation>(),
                Array.Empty<LatexReference>(),
                Array.Empty<LatexLabel>(),
                Array.Empty<LatexTheorem>(),
                Array.Empty<LatexFootnote>(),
                Array.Empty<LatexMacroDefinition>());
        }

        LatexCommand[] activeCommands = commands.Where(static command => IsActiveSyntax(command.Syntax)).ToArray();
        orderedEnvironments = orderedEnvironments.Where(static environment => IsActiveSyntax(environment.Syntax)).ToArray();
        orderedMath = orderedMath.Where(static item => IsActiveSyntax(item.Syntax)).ToArray();
        var headings = new List<LatexHeading>();
        foreach (LatexCommand command in activeCommands) {
            cancellationToken.ThrowIfCancellationRequested();
            if (TryGetHeadingLevel(command.Name, out int level) && command.GetRequiredArgument(0) != null) {
                headings.Add(new LatexHeading(command, level));
            }
        }

        LatexEnvironment? body = orderedEnvironments.FirstOrDefault(static environment => string.Equals(environment.Name, "document", StringComparison.Ordinal));
        cancellationToken.ThrowIfCancellationRequested();
        IReadOnlyList<LatexParagraph> paragraphs = body == null
            ? Array.Empty<LatexParagraph>()
            : BuildParagraphs(source, body, headings, orderedEnvironments, orderedMath, activeCommands, cancellationToken);
        cancellationToken.ThrowIfCancellationRequested();
        var directCommands = IndexDirectEnvironmentCommands(activeCommands, cancellationToken);
        IReadOnlyList<LatexList> lists = BuildLists(source, orderedEnvironments, directCommands, cancellationToken);
        cancellationToken.ThrowIfCancellationRequested();
        IReadOnlyList<LatexFigure> figures = BuildFigures(orderedEnvironments, directCommands, cancellationToken);
        cancellationToken.ThrowIfCancellationRequested();
        IReadOnlyList<LatexTable> tables = BuildTables(source, orderedEnvironments, cancellationToken);
        cancellationToken.ThrowIfCancellationRequested();
        IReadOnlyList<LatexCitation> citations = BuildCitations(activeCommands);
        IReadOnlyList<LatexReference> references = BuildReferences(activeCommands);
        IReadOnlyList<LatexLabel> labels = BuildLabels(activeCommands);
        IReadOnlyList<LatexTheorem> theorems = BuildTheorems(orderedEnvironments, directCommands, cancellationToken);
        IReadOnlyList<LatexFootnote> footnotes = BuildFootnotes(activeCommands, cancellationToken);
        IReadOnlyList<LatexMacroDefinition> macros = BuildMacroDefinitions(activeCommands, cancellationToken);
        cancellationToken.ThrowIfCancellationRequested();
        return new LatexSemanticModel(
            commands,
            orderedEnvironments,
            orderedMath,
            headings,
            paragraphs,
            lists,
            figures,
            tables,
            citations,
            references,
            labels,
            theorems,
            footnotes,
            macros);
    }

    private static IReadOnlyList<LatexList> BuildLists(
        LatexSourceText source,
        IReadOnlyList<LatexEnvironment> environments,
        IReadOnlyDictionary<LatexSyntaxNode, List<LatexCommand>> commands, CancellationToken cancellationToken) {
        var lists = new List<LatexList>();
        for (int environmentIndex = 0; environmentIndex < environments.Count; environmentIndex++) {
            cancellationToken.ThrowIfCancellationRequested();
            LatexEnvironment environment = environments[environmentIndex];
            LatexListKind kind;
            if (string.Equals(environment.Name, "itemize", StringComparison.Ordinal)) kind = LatexListKind.Unordered;
            else if (string.Equals(environment.Name, "enumerate", StringComparison.Ordinal)) kind = LatexListKind.Ordered;
            else if (string.Equals(environment.Name, "description", StringComparison.Ordinal)) kind = LatexListKind.Description;
            else continue;

            IReadOnlyList<LatexCommand> itemCommands = commands.TryGetValue(
                environment.Syntax,
                out List<LatexCommand>? itemCommandsForEnvironment)
                ? itemCommandsForEnvironment.Where(static command => string.Equals(command.Name, "item", StringComparison.Ordinal)).ToArray()
                : Array.Empty<LatexCommand>();
            var listItems = new List<LatexListItem>();
            for (int index = 0; index < itemCommands.Count; index++) {
                cancellationToken.ThrowIfCancellationRequested();
                int start = itemCommands[index].Syntax.EndOffset;
                int end = index + 1 < itemCommands.Count
                    ? itemCommands[index + 1].Syntax.StartOffset
                    : environment.ContentSpan.End.Offset;
                TrimWhitespace(source.Text, ref start, ref end);
                listItems.Add(new LatexListItem(
                    itemCommands[index],
                    source.CreateSpan(start, end),
                    source.Text.Substring(start, end - start)));
            }
            lists.Add(new LatexList(environment, kind, listItems));
        }
        return lists;
    }

    /// <summary>Groups active commands by their nearest environment for bounded semantic binding.</summary>
    private static Dictionary<LatexSyntaxNode, List<LatexCommand>> IndexDirectEnvironmentCommands(
        IReadOnlyList<LatexCommand> commands, CancellationToken cancellationToken) {
        var result = new Dictionary<LatexSyntaxNode, List<LatexCommand>>();
        foreach (LatexCommand command in commands) {
            cancellationToken.ThrowIfCancellationRequested();
            LatexSyntaxNode? parent = command.Syntax.Parent;
            while (parent != null && parent.Kind != LatexSyntaxKind.Environment) parent = parent.Parent;
            if (parent == null) continue;
            if (!result.TryGetValue(parent, out List<LatexCommand>? nested)) {
                nested = new List<LatexCommand>();
                result.Add(parent, nested);
            }
            nested.Add(command);
        }
        return result;
    }

    private static IReadOnlyList<LatexFigure> BuildFigures(
        IReadOnlyList<LatexEnvironment> environments,
        IReadOnlyDictionary<LatexSyntaxNode, List<LatexCommand>> commands,
        CancellationToken cancellationToken) {
        var figures = new List<LatexFigure>();
        foreach (LatexEnvironment environment in environments.Where(static environment => environment.Name == "figure")) {
            cancellationToken.ThrowIfCancellationRequested();
            IReadOnlyList<LatexCommand> nested = commands.TryGetValue(environment.Syntax, out List<LatexCommand>? items)
                ? items : Array.Empty<LatexCommand>();
            var images = new List<LatexImage>();
            LatexCommand? caption = null;
            LatexCommand? label = null;
            foreach (LatexCommand command in nested) {
                cancellationToken.ThrowIfCancellationRequested();
                if (command.Name == "includegraphics" && command.GetRequiredArgument(0) != null) images.Add(new LatexImage(command));
                if (command.Name == "caption" && caption == null) caption = command;
                if (command.Name == "label" && label == null) label = command;
            }
            figures.Add(new LatexFigure(environment, images.ToArray(), caption, label));
        }
        return figures;
    }

    private static IReadOnlyList<LatexTable> BuildTables(LatexSourceText source, IReadOnlyList<LatexEnvironment> environments, CancellationToken cancellationToken) {
        var tables = new List<LatexTable>();
        foreach (LatexEnvironment environment in environments.Where(static environment => string.Equals(environment.Name, "tabular", StringComparison.Ordinal))) {
            cancellationToken.ThrowIfCancellationRequested();
            string columnSpecification = environment.BeginCommand.GetRequiredArgument(1)?.Content ?? string.Empty;
            tables.Add(new LatexTable(environment, columnSpecification, ParseTableRows(source, environment, cancellationToken)));
        }
        return tables;
    }

    private static IReadOnlyList<LatexTableRow> ParseTableRows(LatexSourceText source, LatexEnvironment environment, CancellationToken cancellationToken) {
        var rows = new List<LatexTableRow>();
        var currentCells = new List<LatexTableCell>();
        int cellStart = environment.ContentSpan.Start.Offset;
        int ignoreUntil = cellStart;
        foreach (LatexSyntaxNode node in environment.Syntax.Children) {
            cancellationToken.ThrowIfCancellationRequested();
            int nodeStart = node.StartOffset;
            if (nodeStart < environment.ContentSpan.Start.Offset || nodeStart >= environment.ContentSpan.End.Offset ||
                nodeStart < ignoreUntil) continue;

            if (node.Kind == LatexSyntaxKind.Text && string.Equals(node.OriginalText, "&", StringComparison.Ordinal)) {
                AddTableCell(source, cellStart, nodeStart, rows.Count, currentCells.Count, currentCells, true);
                cellStart = node.EndOffset;
                continue;
            }

            if (node.Kind == LatexSyntaxKind.Command && string.Equals(node.Value, "\\", StringComparison.Ordinal)) {
                AddTableCell(source, cellStart, nodeStart, rows.Count, currentCells.Count, currentCells, true);
                if (currentCells.Count > 0 && !IsRuleOnlyRow(currentCells)) {
                    rows.Add(new LatexTableRow(rows.Count, currentCells.ToArray()));
                }
                currentCells = new List<LatexTableCell>();
                int rowContentStart = node.EndOffset;
                while (rowContentStart < environment.ContentSpan.End.Offset && source.Text[rowContentStart] == '*') rowContentStart++;
                if (rowContentStart < environment.ContentSpan.End.Offset && source.Text[rowContentStart] == '[') {
                    SkipBalanced(source.Text, ref rowContentStart, '[', ']', environment.ContentSpan.End.Offset, cancellationToken);
                }
                cellStart = rowContentStart;
                ignoreUntil = rowContentStart;
            }
        }
        AddTableCell(source, cellStart, environment.ContentSpan.End.Offset, rows.Count, currentCells.Count, currentCells);
        if (currentCells.Count > 0 && !IsRuleOnlyRow(currentCells)) rows.Add(new LatexTableRow(rows.Count, currentCells.ToArray()));
        return rows;
    }

    private static void AddTableCell(
        LatexSourceText source,
        int start,
        int end,
        int row,
        int column,
        List<LatexTableCell> cells,
        bool separator = false) {
        TrimWhitespace(source.Text, ref start, ref end);
        if (end <= start && cells.Count == 0 && !separator) return;
        cells.Add(new LatexTableCell(source.CreateSpan(start, end), source.Text.Substring(start, end - start), row, column));
    }

    private static bool IsRuleOnlyRow(IReadOnlyList<LatexTableCell> cells) {
        if (cells.Count != 1) return false;
        string value = cells[0].Content.Trim();
        return value == "\\hline" || value == "\\toprule" || value == "\\midrule" || value == "\\bottomrule";
    }

    private static IReadOnlyList<LatexCitation> BuildCitations(IReadOnlyList<LatexCommand> commands) =>
        commands.Where(static command => IsCitationCommand(command.Name) && command.GetRequiredArgument(0) != null)
            .Select(command => new LatexCitation(command, SplitComma(command.GetRequiredArgument(0)!.Content))).ToArray();

    private static IReadOnlyList<LatexReference> BuildReferences(IReadOnlyList<LatexCommand> commands) =>
        commands.Where(static command => IsReferenceCommand(command.Name) && command.GetRequiredArgument(0) != null)
            .Select(command => new LatexReference(command, command.GetRequiredArgument(0)!.Content)).ToArray();

    private static IReadOnlyList<LatexLabel> BuildLabels(IReadOnlyList<LatexCommand> commands) =>
        commands.Where(static command => string.Equals(command.Name, "label", StringComparison.Ordinal) && command.GetRequiredArgument(0) != null)
            .Select(command => new LatexLabel(command, command.GetRequiredArgument(0)!.Content)).ToArray();

    private static IReadOnlyList<LatexTheorem> BuildTheorems(
        IReadOnlyList<LatexEnvironment> environments,
        IReadOnlyDictionary<LatexSyntaxNode, List<LatexCommand>> commands,
        CancellationToken cancellationToken) {
        var theorems = new List<LatexTheorem>();
        foreach (LatexEnvironment environment in environments.Where(static environment => IsTheoremEnvironment(environment.Name))) {
            cancellationToken.ThrowIfCancellationRequested();
            LatexCommand? label = null;
            if (commands.TryGetValue(environment.Syntax, out List<LatexCommand>? nested)) {
                foreach (LatexCommand command in nested) {
                    cancellationToken.ThrowIfCancellationRequested();
                    if (command.Name == "label") { label = command; break; }
                }
            }
            theorems.Add(new LatexTheorem(environment, label));
        }
        return theorems;
    }

    internal static bool IsActiveSyntax(LatexSyntaxNode node) {
        for (LatexSyntaxNode? parent = node.Parent; parent != null; parent = parent.Parent) {
            if (parent.Kind == LatexSyntaxKind.Command &&
                (parent.Value == "newcommand" || parent.Value == "renewcommand" || parent.Value == "providecommand" || parent.Value == "newtheorem")) return false;
        }
        return true;
    }

    internal static bool IsInsideCommandArgument(LatexSyntaxNode node) {
        for (LatexSyntaxNode? parent = node.Parent; parent != null; parent = parent.Parent) {
            if (parent.Kind == LatexSyntaxKind.Command) return true;
        }
        return false;
    }

    // Container semantics own direct commands even when the whole container is
    // inside an inline body. An intervening command still owns its argument.
    internal static bool IsInsideCommandArgumentBeforeEnvironment(LatexSyntaxNode node) {
        for (LatexSyntaxNode? parent = node.Parent; parent != null; parent = parent.Parent) {
            if (parent.Kind == LatexSyntaxKind.Environment) return false;
            if (parent.Kind == LatexSyntaxKind.Command) return true;
        }
        return false;
    }

    private static void TrimWhitespace(string source, ref int start, ref int end) {
        while (start < end && char.IsWhiteSpace(source[start])) start++;
        while (end > start && char.IsWhiteSpace(source[end - 1])) end--;
    }

    private static void SkipBalanced(string source, ref int index, char open, char close, int end, CancellationToken cancellationToken) {
        if (index >= end || source[index] != open) return;
        int depth = 1;
        index++;
        while (index < end && depth > 0) {
            if ((index & 1023) == 0) cancellationToken.ThrowIfCancellationRequested();
            if (source[index] == '\\') { index += Math.Min(2, end - index); continue; }
            if (source[index] == open) depth++;
            else if (source[index] == close) depth--;
            index++;
        }
    }

    private static IReadOnlyList<string> SplitComma(string value) =>
        value.Split(new[] { ',' }, StringSplitOptions.RemoveEmptyEntries).Select(static item => item.Trim()).Where(static item => item.Length > 0).ToArray();

    private static bool IsCitationCommand(string name) =>
        name == "cite" || name == "citep" || name == "citet" || name == "nocite";

    private static bool IsReferenceCommand(string name) =>
        name == "ref" || name == "pageref" || name == "autoref" || name == "eqref";

    private static bool IsTheoremEnvironment(string name) =>
        name == "theorem" || name == "lemma" || name == "proposition" || name == "corollary" ||
        name == "definition" || name == "remark" || name == "proof";

    private static IReadOnlyList<LatexParagraph> BuildParagraphs(
        LatexSourceText source,
        LatexEnvironment body,
        IReadOnlyList<LatexHeading> headings,
        IReadOnlyList<LatexEnvironment> environments,
        IReadOnlyList<LatexMath> math,
        IReadOnlyList<LatexCommand> commands, CancellationToken cancellationToken) {
        var blocked = new List<LatexSourceSpan>();
        blocked.AddRange(headings.Where(static heading => !IsInsideFootnote(heading.Command.Syntax))
            .Select(static heading => heading.Command.Syntax.Span));
        LatexCommand[] labels = commands.Where(static command => string.Equals(command.Name, "label", StringComparison.Ordinal)).ToArray();
        int labelIndex = 0;
        for (int index = 0; index < headings.Count; index++) {
            cancellationToken.ThrowIfCancellationRequested();
            if (IsInsideFootnote(headings[index].Command.Syntax)) continue;
            LatexSourceSpan headingSpan = headings[index].Command.Syntax.Span;
            while (labelIndex < labels.Length && labels[labelIndex].Syntax.StartOffset < headingSpan.End.Offset) labelIndex++;
            LatexCommand? label = labelIndex < labels.Length
                && labels[labelIndex].Syntax.EndOffset <= body.ContentSpan.End.Offset
                && IsWhitespaceOnly(source.Text, headingSpan.End.Offset, labels[labelIndex].Syntax.StartOffset)
                    ? labels[labelIndex]
                    : null;
            if (label != null) blocked.Add(label.Syntax.Span);
        }
        blocked.AddRange(body.Syntax.DescendantsAndSelf()
            .Where(static node => node.Kind == LatexSyntaxKind.Command && IsActiveSyntax(node) && !IsInsideFootnote(node) && string.Equals(node.Value, "maketitle", StringComparison.Ordinal))
            .Select(static node => node.Span));
        blocked.AddRange(body.Syntax.DescendantsAndSelf()
            .Where(static node => node.Kind == LatexSyntaxKind.Verbatim && IsActiveSyntax(node) && !IsInsideCommandArgument(node) &&
                !string.Equals(node.Value, "verb", StringComparison.Ordinal))
            .Select(static node => node.Span));
        blocked.AddRange(environments.Where(environment => !ReferenceEquals(environment, body) && !IsInsideFootnote(environment.Syntax) &&
            environment.Syntax.StartOffset >= body.ContentSpan.Start.Offset &&
            environment.Syntax.EndOffset <= body.ContentSpan.End.Offset).Select(static environment => environment.Syntax.Span));
        blocked.AddRange(math.Where(static item => !IsInsideFootnote(item.Syntax) && item.Kind != LatexMathKind.InlineDollar && item.Kind != LatexMathKind.InlineParentheses && item.Kind != LatexMathKind.Environment)
            .Select(static item => item.Syntax.Span));
        blocked = Merge(blocked
            .Where(span => span.End.Offset > body.ContentSpan.Start.Offset && span.Start.Offset < body.ContentSpan.End.Offset)
            .Select(span => source.CreateSpan(Math.Max(span.Start.Offset, body.ContentSpan.Start.Offset),
                Math.Min(span.End.Offset, body.ContentSpan.End.Offset)))
            .OrderBy(static span => span.Start.Offset).ToList());

        var paragraphs = new List<LatexParagraph>();
        LatexSourceSpan[] inlineBodies = commands.Where(static command => command.Name == "footnote")
            .Select(static command => command.Syntax.Span).ToArray();
        int cursor = body.ContentSpan.Start.Offset;
        for (int index = 0; index < blocked.Count; index++) {
            cancellationToken.ThrowIfCancellationRequested();
            LatexSourceSpan span = blocked[index];
            if (span.Start.Offset > cursor) AddParagraphSegments(source, cursor, span.Start.Offset, paragraphs, cancellationToken, inlineBodies);
            cursor = Math.Max(cursor, span.End.Offset);
        }
        if (cursor < body.ContentSpan.End.Offset) AddParagraphSegments(source, cursor, body.ContentSpan.End.Offset, paragraphs, cancellationToken, inlineBodies);
        return paragraphs;
    }

    private static bool IsWhitespaceOnly(string source, int start, int end) {
        for (int index = start; index < end; index++) {
            if (!char.IsWhiteSpace(source[index])) return false;
        }
        return true;
    }

    /// <summary>Builds source-backed paragraph segments for a bounded container body without reparsing or rebasing offsets.</summary>
    internal static IReadOnlyList<LatexParagraph> BuildParagraphsInSpan(
        LatexSourceText source, LatexSourceSpan span, CancellationToken cancellationToken,
        IReadOnlyList<LatexSourceSpan>? inlineBodies = null) {
        var paragraphs = new List<LatexParagraph>();
        AddParagraphSegments(source, span.Start.Offset, span.End.Offset, paragraphs, cancellationToken, inlineBodies);
        return paragraphs;
    }

    private static void AddParagraphSegments(LatexSourceText source, int start, int end, List<LatexParagraph> paragraphs,
        CancellationToken cancellationToken, IReadOnlyList<LatexSourceSpan>? inlineBodies = null) {
        int segmentStart = start;
        int index = start;
        int low = 0, high = inlineBodies?.Count ?? 0;
        while (low < high) {
            int middle = low + (high - low) / 2;
            if (inlineBodies![middle].Start.Offset < start) low = middle + 1;
            else high = middle;
        }
        int inlineIndex = low;
        while (index < end) {
            if ((index & 1023) == 0) cancellationToken.ThrowIfCancellationRequested();
            while (inlineBodies != null && inlineIndex < inlineBodies.Count && inlineBodies[inlineIndex].Start.Offset < index) inlineIndex++;
            if (inlineBodies != null && inlineIndex < inlineBodies.Count &&
                inlineBodies[inlineIndex].Start.Offset == index && inlineBodies[inlineIndex].End.Offset <= end) {
                index = inlineBodies[inlineIndex++].End.Offset;
                cancellationToken.ThrowIfCancellationRequested();
                continue;
            }
            if (source.Text[index] == '\\') { index += Math.Min(2, end - index); continue; }
            if (source.Text[index] == '%') {
                while (index < end && source.Text[index] != '\r' && source.Text[index] != '\n') {
                    if ((index & 1023) == 0) cancellationToken.ThrowIfCancellationRequested();
                    index++;
                }
                if (TryReadLineEnding(source.Text, index, end, out int commentEnding)) {
                    int nextLine = index + commentEnding;
                    while (nextLine < end && (source.Text[nextLine] == ' ' || source.Text[nextLine] == '\t')) nextLine++;
                    if (TryReadLineEnding(source.Text, nextLine, end, out int blankEnding)) {
                        AddTrimmedParagraph(source, segmentStart, index, paragraphs);
                        segmentStart = nextLine + blankEnding;
                        index = segmentStart;
                    } else {
                        index += commentEnding;
                    }
                }
                continue;
            }
            if (!TryReadLineEnding(source.Text, index, end, out int firstLength)) { index++; continue; }
            int lookahead = index + firstLength;
            while (lookahead < end && (source.Text[lookahead] == ' ' || source.Text[lookahead] == '\t')) lookahead++;
            if (!TryReadLineEnding(source.Text, lookahead, end, out int secondLength)) { index += firstLength; continue; }
            AddTrimmedParagraph(source, segmentStart, index, paragraphs);
            segmentStart = lookahead + secondLength;
            index = segmentStart;
        }
        AddTrimmedParagraph(source, segmentStart, end, paragraphs);
    }

    private static void AddTrimmedParagraph(LatexSourceText source, int start, int end, List<LatexParagraph> paragraphs) {
        while (start < end && char.IsWhiteSpace(source.Text[start])) start++;
        while (end > start && char.IsWhiteSpace(source.Text[end - 1])) end--;
        if (end <= start) return;
        LatexSourceSpan span = source.CreateSpan(start, end);
        paragraphs.Add(new LatexParagraph(span, source.Text.Substring(start, end - start)));
    }

    private static bool TryReadLineEnding(string source, int index, int end, out int length) {
        length = 0;
        if (index >= end) return false;
        if (source[index] == '\r') { length = index + 1 < end && source[index + 1] == '\n' ? 2 : 1; return true; }
        if (source[index] == '\n') { length = 1; return true; }
        return false;
    }

    private static List<LatexSourceSpan> Merge(List<LatexSourceSpan> spans) {
        if (spans.Count < 2) return spans;
        var result = new List<LatexSourceSpan>();
        LatexSourceSpan current = spans[0];
        for (int index = 1; index < spans.Count; index++) {
            LatexSourceSpan next = spans[index];
            if (next.Start.Offset <= current.End.Offset) {
                current = new LatexSourceSpan(current.Start,
                    next.End.Offset > current.End.Offset ? next.End : current.End);
            } else {
                result.Add(current);
                current = next;
            }
        }
        result.Add(current);
        return result;
    }

    private static bool TryGetHeadingLevel(string name, out int level) {
        switch (name) {
            case "part": level = 0; return true;
            case "chapter": level = 1; return true;
            case "section": level = 2; return true;
            case "subsection": level = 3; return true;
            case "subsubsection": level = 4; return true;
            case "paragraph": level = 5; return true;
            case "subparagraph": level = 6; return true;
            default: level = 0; return false;
        }
    }
}
