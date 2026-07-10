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
    internal IReadOnlyList<LatexMacroDefinition> MacroDefinitions { get; }
}

internal static class LatexSemanticBuilder {
    internal static LatexSemanticModel Build(
        LatexSourceText source,
        LatexSyntaxTree syntaxTree,
        LatexDocumentProfile profile) {
        LatexSyntaxNode[] commandSyntax = syntaxTree.Root.DescendantsAndSelf()
            .Where(static node => node.Kind == LatexSyntaxKind.Command)
            .ToArray();
        var commandMap = new Dictionary<LatexSyntaxNode, LatexCommand>();
        for (int index = 0; index < commandSyntax.Length; index++) {
            commandMap[commandSyntax[index]] = new LatexCommand(commandSyntax[index], source);
        }

        LatexSyntaxNode[] environmentSyntax = syntaxTree.Root.DescendantsAndSelf()
            .Where(static node => node.Kind == LatexSyntaxKind.Environment)
            .ToArray();
        var environments = new List<LatexEnvironment>(environmentSyntax.Length);
        for (int index = 0; index < environmentSyntax.Length; index++) {
            LatexSyntaxNode syntax = environmentSyntax[index];
            LatexSyntaxNode beginSyntax = syntax.Children.First(static child => child.Kind == LatexSyntaxKind.Command);
            LatexSyntaxNode? endSyntax = syntax.Children.LastOrDefault(child =>
                child.Kind == LatexSyntaxKind.Command && string.Equals(child.Value, "end", StringComparison.Ordinal));
            environments.Add(new LatexEnvironment(
                syntax,
                commandMap[beginSyntax],
                endSyntax == null ? null : commandMap[endSyntax],
                source));
        }

        var math = syntaxTree.Root.DescendantsAndSelf()
            .Where(static node => node.Kind == LatexSyntaxKind.Math)
            .Select(node => new LatexMath(node, source))
            .ToList();
        math.AddRange(environments.Where(static environment => environment.IsMath).Select(static environment => new LatexMath(environment)));

        LatexCommand[] commands = commandMap.Values.OrderBy(static command => command.Syntax.Span.Start.Offset).ToArray();
        LatexEnvironment[] orderedEnvironments = environments.OrderBy(static environment => environment.Syntax.Span.Start.Offset).ToArray();
        LatexMath[] orderedMath = math.OrderBy(static item => item.Syntax.Span.Start.Offset).ToArray();
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
                Array.Empty<LatexMacroDefinition>());
        }

        var headings = new List<LatexHeading>();
        foreach (LatexCommand command in commands) {
            if (TryGetHeadingLevel(command.Name, out int level) && command.GetRequiredArgument(0) != null) {
                headings.Add(new LatexHeading(command, level));
            }
        }

        LatexEnvironment? body = environments.FirstOrDefault(static environment => string.Equals(environment.Name, "document", StringComparison.Ordinal));
        IReadOnlyList<LatexParagraph> paragraphs = body == null
            ? Array.Empty<LatexParagraph>()
            : BuildParagraphs(source, body, headings, orderedEnvironments, orderedMath, commands);
        IReadOnlyList<LatexList> lists = BuildLists(source, orderedEnvironments, commands);
        IReadOnlyList<LatexFigure> figures = BuildFigures(orderedEnvironments, commands);
        IReadOnlyList<LatexTable> tables = BuildTables(source, orderedEnvironments);
        IReadOnlyList<LatexCitation> citations = BuildCitations(commands);
        IReadOnlyList<LatexReference> references = BuildReferences(commands);
        IReadOnlyList<LatexLabel> labels = BuildLabels(commands);
        IReadOnlyList<LatexTheorem> theorems = BuildTheorems(orderedEnvironments, commands);
        IReadOnlyList<LatexMacroDefinition> macros = BuildMacroDefinitions(commands);
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
            macros);
    }

    private static IReadOnlyList<LatexList> BuildLists(
        LatexSourceText source,
        IReadOnlyList<LatexEnvironment> environments,
        IReadOnlyList<LatexCommand> commands) {
        var lists = new List<LatexList>();
        for (int environmentIndex = 0; environmentIndex < environments.Count; environmentIndex++) {
            LatexEnvironment environment = environments[environmentIndex];
            LatexListKind kind;
            if (string.Equals(environment.Name, "itemize", StringComparison.Ordinal)) kind = LatexListKind.Unordered;
            else if (string.Equals(environment.Name, "enumerate", StringComparison.Ordinal)) kind = LatexListKind.Ordered;
            else if (string.Equals(environment.Name, "description", StringComparison.Ordinal)) kind = LatexListKind.Description;
            else continue;

            LatexCommand[] itemCommands = commands.Where(command => string.Equals(command.Name, "item", StringComparison.Ordinal) &&
                    IsDirectlyInside(command.Syntax, environment.Syntax))
                .OrderBy(static command => command.Syntax.Span.Start.Offset)
                .ToArray();
            var items = new List<LatexListItem>();
            for (int index = 0; index < itemCommands.Length; index++) {
                int start = itemCommands[index].Syntax.Span.End.Offset;
                int end = index + 1 < itemCommands.Length
                    ? itemCommands[index + 1].Syntax.Span.Start.Offset
                    : environment.ContentSpan.End.Offset;
                TrimWhitespace(source.Text, ref start, ref end);
                items.Add(new LatexListItem(
                    itemCommands[index],
                    source.CreateSpan(start, end),
                    source.Text.Substring(start, end - start)));
            }
            lists.Add(new LatexList(environment, kind, items));
        }
        return lists;
    }

    private static IReadOnlyList<LatexFigure> BuildFigures(
        IReadOnlyList<LatexEnvironment> environments,
        IReadOnlyList<LatexCommand> commands) {
        var figures = new List<LatexFigure>();
        foreach (LatexEnvironment environment in environments.Where(static environment => string.Equals(environment.Name, "figure", StringComparison.Ordinal))) {
            LatexCommand[] nested = commands.Where(command => IsDirectlyInside(command.Syntax, environment.Syntax)).ToArray();
            LatexImage[] images = nested.Where(static command => string.Equals(command.Name, "includegraphics", StringComparison.Ordinal) && command.GetRequiredArgument(0) != null)
                .Select(static command => new LatexImage(command)).ToArray();
            figures.Add(new LatexFigure(
                environment,
                images,
                nested.FirstOrDefault(static command => string.Equals(command.Name, "caption", StringComparison.Ordinal)),
                nested.FirstOrDefault(static command => string.Equals(command.Name, "label", StringComparison.Ordinal))));
        }
        return figures;
    }

    private static IReadOnlyList<LatexTable> BuildTables(LatexSourceText source, IReadOnlyList<LatexEnvironment> environments) {
        var tables = new List<LatexTable>();
        foreach (LatexEnvironment environment in environments.Where(static environment => string.Equals(environment.Name, "tabular", StringComparison.Ordinal))) {
            string columnSpecification = environment.BeginCommand.GetRequiredArgument(1)?.Content ?? string.Empty;
            tables.Add(new LatexTable(environment, columnSpecification, ParseTableRows(source, environment)));
        }
        return tables;
    }

    private static IReadOnlyList<LatexTableRow> ParseTableRows(LatexSourceText source, LatexEnvironment environment) {
        var rows = new List<LatexTableRow>();
        var currentCells = new List<LatexTableCell>();
        int cellStart = environment.ContentSpan.Start.Offset;
        int braceDepth = 0;
        int index = cellStart;
        while (index < environment.ContentSpan.End.Offset) {
            char current = source.Text[index];
            if (current == '%') {
                while (index < environment.ContentSpan.End.Offset && source.Text[index] != '\r' && source.Text[index] != '\n') index++;
                continue;
            }
            if (current == '\\') {
                if (index + 1 < environment.ContentSpan.End.Offset && source.Text[index + 1] == '\\' && braceDepth == 0) {
                    AddTableCell(source, cellStart, index, rows.Count, currentCells.Count, currentCells);
                    if (currentCells.Count > 0 && !IsRuleOnlyRow(currentCells)) rows.Add(new LatexTableRow(rows.Count, currentCells.ToArray()));
                    currentCells = new List<LatexTableCell>();
                    index += 2;
                    while (index < environment.ContentSpan.End.Offset && source.Text[index] == '*') index++;
                    if (index < environment.ContentSpan.End.Offset && source.Text[index] == '[') SkipBalanced(source.Text, ref index, '[', ']', environment.ContentSpan.End.Offset);
                    cellStart = index;
                    continue;
                }
                index += Math.Min(2, environment.ContentSpan.End.Offset - index);
                continue;
            }
            if (current == '{') braceDepth++;
            else if (current == '}' && braceDepth > 0) braceDepth--;
            else if (current == '&' && braceDepth == 0) {
                AddTableCell(source, cellStart, index, rows.Count, currentCells.Count, currentCells);
                cellStart = index + 1;
            }
            index++;
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
        List<LatexTableCell> cells) {
        TrimWhitespace(source.Text, ref start, ref end);
        if (end <= start && cells.Count == 0) return;
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
        IReadOnlyList<LatexCommand> commands) {
        return environments.Where(static environment => IsTheoremEnvironment(environment.Name))
            .Select(environment => new LatexTheorem(
                environment,
                commands.FirstOrDefault(command => string.Equals(command.Name, "label", StringComparison.Ordinal) && IsDirectlyInside(command.Syntax, environment.Syntax))))
            .ToArray();
    }

    private static IReadOnlyList<LatexMacroDefinition> BuildMacroDefinitions(IReadOnlyList<LatexCommand> commands) {
        var candidates = new List<MacroCandidate>();
        foreach (LatexCommand command in commands.Where(static command =>
                     string.Equals(command.Name, "newcommand", StringComparison.Ordinal) ||
                     string.Equals(command.Name, "renewcommand", StringComparison.Ordinal) ||
                     string.Equals(command.Name, "providecommand", StringComparison.Ordinal))) {
            LatexArgument? nameArgument = command.GetRequiredArgument(0);
            LatexArgument? bodyArgument = command.GetRequiredArgument(1);
            if (nameArgument == null || bodyArgument == null) continue;
            string name = nameArgument.Content.Trim();
            if (name.StartsWith("\\", StringComparison.Ordinal)) name = name.Substring(1);
            if (!IsSimpleControlWord(name)) continue;
            int parameterCount = 0;
            string? defaultValue = null;
            LatexArgument[] optional = command.Arguments.Where(static argument => argument.IsOptional).ToArray();
            bool isWellFormed = optional.Length == 0 || int.TryParse(optional[0].Content.Trim(), out parameterCount);
            if (optional.Length > 1) defaultValue = optional[1].Content;
            if (defaultValue != null && parameterCount < 1) isWellFormed = false;
            candidates.Add(new MacroCandidate(command, name, parameterCount, defaultValue, bodyArgument.Content, isWellFormed));
        }

        var allowedLocalNames = new HashSet<string>(candidates.Select(static candidate => candidate.Name), StringComparer.Ordinal);
        var safeCandidates = new bool[candidates.Count];
        while (true) {
            for (int index = 0; index < candidates.Count; index++) {
                MacroCandidate candidate = candidates[index];
                safeCandidates[index] = candidate.IsWellFormed && candidate.ParameterCount >= 0 && candidate.ParameterCount <= 9 &&
                    IsSafeMacroBody(candidate.Body, candidate.ParameterCount, allowedLocalNames) &&
                    (candidate.DefaultValue == null || IsSafeMacroBody(candidate.DefaultValue, 0, allowedLocalNames));
            }

            var nextAllowedNames = new HashSet<string>(StringComparer.Ordinal);
            foreach (IGrouping<string, int> group in candidates
                         .Select(static (candidate, index) => new { candidate.Name, Index = index })
                         .GroupBy(static item => item.Name, static item => item.Index, StringComparer.Ordinal)) {
                if (group.All(index => safeCandidates[index])) nextAllowedNames.Add(group.Key);
            }
            if (allowedLocalNames.SetEquals(nextAllowedNames)) break;
            allowedLocalNames = nextAllowedNames;
        }

        var definitions = new List<LatexMacroDefinition>(candidates.Count);
        for (int index = 0; index < candidates.Count; index++) {
            MacroCandidate candidate = candidates[index];
            definitions.Add(new LatexMacroDefinition(
                candidate.Command,
                candidate.Name,
                candidate.ParameterCount,
                candidate.DefaultValue,
                candidate.Body,
                safeCandidates[index]));
        }
        return definitions;
    }

    private static bool IsDirectlyInside(LatexSyntaxNode node, LatexSyntaxNode environment) {
        LatexSyntaxNode? current = node.Parent;
        while (current != null) {
            if (current.Kind == LatexSyntaxKind.Environment) return ReferenceEquals(current, environment);
            current = current.Parent;
        }
        return false;
    }

    private static void TrimWhitespace(string source, ref int start, ref int end) {
        while (start < end && char.IsWhiteSpace(source[start])) start++;
        while (end > start && char.IsWhiteSpace(source[end - 1])) end--;
    }

    private static void SkipBalanced(string source, ref int index, char open, char close, int end) {
        if (index >= end || source[index] != open) return;
        int depth = 1;
        index++;
        while (index < end && depth > 0) {
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

    private static bool IsSimpleControlWord(string value) {
        if (value.Length == 0) return false;
        for (int index = 0; index < value.Length; index++) {
            char current = value[index];
            if (!((current >= 'a' && current <= 'z') || (current >= 'A' && current <= 'Z') || current == '@')) return false;
        }
        return true;
    }

    private static bool IsSafeMacroBody(string body, int parameterCount, IReadOnlyCollection<string> localNames) {
        for (int index = 0; index < body.Length; index++) {
            if (body[index] != '\\' || index + 1 >= body.Length) continue;
            if (!IsControlWordCharacter(body[index + 1])) {
                index++;
                continue;
            }
            int nameStart = ++index;
            while (index + 1 < body.Length && IsControlWordCharacter(body[index + 1])) index++;
            string name = body.Substring(nameStart, index - nameStart + 1);
            if (!localNames.Contains(name) && !IsSafeReplacementCommand(name)) return false;
        }
        for (int index = 0; index + 1 < body.Length; index++) {
            if (body[index] != '#') continue;
            if (body[index + 1] == '#') { index++; continue; }
            if (body[index + 1] < '1' || body[index + 1] > '9' || body[index + 1] - '0' > parameterCount) return false;
        }
        return true;
    }

    private static bool IsSafeReplacementCommand(string name) {
        switch (name) {
            case "textbf":
            case "textit":
            case "emph":
            case "texttt":
            case "underline":
            case "textsuperscript":
            case "textsubscript":
            case "mathrm":
            case "mathbf":
            case "mathit":
            case "mathsf":
            case "mathtt":
            case "operatorname":
            case "frac":
            case "sqrt":
            case "ensuremath":
            case "protect":
            case "ref":
            case "pageref":
            case "autoref":
            case "eqref":
            case "cite":
            case "citep":
            case "citet":
            case "url":
            case "href":
            case "footnote":
            case "newline":
            case "linebreak":
                return true;
            default:
                return false;
        }
    }

    private static bool IsControlWordCharacter(char value) =>
        (value >= 'a' && value <= 'z') || (value >= 'A' && value <= 'Z') || value == '@';

    private sealed class MacroCandidate {
        internal MacroCandidate(
            LatexCommand command,
            string name,
            int parameterCount,
            string? defaultValue,
            string body,
            bool isWellFormed) {
            Command = command;
            Name = name;
            ParameterCount = parameterCount;
            DefaultValue = defaultValue;
            Body = body;
            IsWellFormed = isWellFormed;
        }

        internal LatexCommand Command { get; }
        internal string Name { get; }
        internal int ParameterCount { get; }
        internal string? DefaultValue { get; }
        internal string Body { get; }
        internal bool IsWellFormed { get; }
    }

    private static IReadOnlyList<LatexParagraph> BuildParagraphs(
        LatexSourceText source,
        LatexEnvironment body,
        IReadOnlyList<LatexHeading> headings,
        IReadOnlyList<LatexEnvironment> environments,
        IReadOnlyList<LatexMath> math,
        IReadOnlyList<LatexCommand> commands) {
        var blocked = new List<LatexSourceSpan>();
        blocked.AddRange(headings.Select(static heading => heading.Command.Syntax.Span));
        for (int index = 0; index < headings.Count; index++) {
            LatexSourceSpan headingSpan = headings[index].Command.Syntax.Span;
            LatexCommand? label = commands.FirstOrDefault(command =>
                string.Equals(command.Name, "label", StringComparison.Ordinal) &&
                command.Syntax.Span.Start.Offset >= headingSpan.End.Offset &&
                command.Syntax.Span.End.Offset <= body.ContentSpan.End.Offset &&
                IsWhitespaceOnly(source.Text, headingSpan.End.Offset, command.Syntax.Span.Start.Offset));
            if (label != null) blocked.Add(label.Syntax.Span);
        }
        blocked.AddRange(body.Syntax.DescendantsAndSelf()
            .Where(static node => node.Kind == LatexSyntaxKind.Command && string.Equals(node.Value, "maketitle", StringComparison.Ordinal))
            .Select(static node => node.Span));
        blocked.AddRange(environments.Where(environment => !ReferenceEquals(environment, body) &&
            environment.Syntax.Span.Start.Offset >= body.ContentSpan.Start.Offset &&
            environment.Syntax.Span.End.Offset <= body.ContentSpan.End.Offset).Select(static environment => environment.Syntax.Span));
        blocked.AddRange(math.Where(static item => item.Kind != LatexMathKind.InlineDollar && item.Kind != LatexMathKind.InlineParentheses && item.Kind != LatexMathKind.Environment)
            .Select(static item => item.Syntax.Span));
        blocked = Merge(blocked.OrderBy(static span => span.Start.Offset).ToList());

        var paragraphs = new List<LatexParagraph>();
        int cursor = body.ContentSpan.Start.Offset;
        for (int index = 0; index < blocked.Count; index++) {
            LatexSourceSpan span = blocked[index];
            if (span.Start.Offset > cursor) AddParagraphSegments(source, cursor, span.Start.Offset, paragraphs);
            cursor = Math.Max(cursor, span.End.Offset);
        }
        if (cursor < body.ContentSpan.End.Offset) AddParagraphSegments(source, cursor, body.ContentSpan.End.Offset, paragraphs);
        return paragraphs;
    }

    private static bool IsWhitespaceOnly(string source, int start, int end) {
        for (int index = start; index < end; index++) {
            if (!char.IsWhiteSpace(source[index])) return false;
        }
        return true;
    }

    private static void AddParagraphSegments(LatexSourceText source, int start, int end, List<LatexParagraph> paragraphs) {
        int segmentStart = start;
        int index = start;
        while (index < end) {
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
