namespace OfficeIMO.Latex.Markdown;

internal static class LatexInlineToMarkdownConverter {
    internal static InlineSequence Convert(
        LatexDocument document,
        LatexSourceSpan span,
        List<LatexMarkdownConversionDiagnostic> diagnostics) {
        var target = new InlineSequence { AutoSpacing = false };
        AddRange(target, document, span.Start.Offset, span.End.Offset, diagnostics);
        return target;
    }

    internal static InlineSequence ConvertExcluding(
        LatexDocument document,
        LatexSourceSpan span,
        IEnumerable<LatexSourceSpan> excludedSpans,
        List<LatexMarkdownConversionDiagnostic> diagnostics) {
        var target = new InlineSequence { AutoSpacing = false };
        int cursor = span.Start.Offset;
        foreach (LatexSourceSpan excluded in excludedSpans
                     .Where(excluded => excluded.End.Offset > span.Start.Offset && excluded.Start.Offset < span.End.Offset)
                     .OrderBy(static excluded => excluded.Start.Offset)) {
            int excludedStart = Math.Max(cursor, excluded.Start.Offset);
            int excludedEnd = Math.Min(span.End.Offset, excluded.End.Offset);
            if (excludedStart > cursor) AddRange(target, document, cursor, excludedStart, diagnostics);
            cursor = Math.Max(cursor, excludedEnd);
        }
        if (cursor < span.End.Offset) AddRange(target, document, cursor, span.End.Offset, diagnostics);
        return target;
    }

    private static void AddRange(
        InlineSequence target,
        LatexDocument document,
        int start,
        int end,
        List<LatexMarkdownConversionDiagnostic> diagnostics) {
        var candidates = new List<InlineCandidate>();
        candidates.AddRange(document.Commands
            .Where(command => LatexSemanticBuilder.IsActiveSyntax(command.Syntax) && command.Syntax.Span.Start.Offset >= start && command.Syntax.Span.End.Offset <= end)
            .Select(static command => new InlineCandidate(command.Syntax.Span, command, null)));
        candidates.AddRange(document.Math
            .Where(math => math.Syntax.Span.Start.Offset >= start && math.Syntax.Span.End.Offset <= end && math.Kind != LatexMathKind.Environment)
            .Select(static math => new InlineCandidate(math.Syntax.Span, null, math)));
        candidates.AddRange(document.SyntaxTree.Root.DescendantsAndSelf()
            .Where(node => node.Kind == LatexSyntaxKind.Verbatim && LatexSemanticBuilder.IsActiveSyntax(node) &&
                node.Span.Start.Offset >= start && node.Span.End.Offset <= end)
            .Select(static node => new InlineCandidate(node.Span, null, null, node)));
        InlineCandidate[] ordered = candidates.OrderBy(static candidate => candidate.Span.Start.Offset)
            .ThenByDescending(static candidate => candidate.Span.End.Offset)
            .ToArray();

        int cursor = start;
        for (int index = 0; index < ordered.Length; index++) {
            InlineCandidate candidate = ordered[index];
            if (candidate.Span.Start.Offset < cursor) continue;
            AddPlain(target, document.Source.Text.Substring(cursor, candidate.Span.Start.Offset - cursor));
            if (candidate.Command != null) AddCommand(target, document, candidate.Command, diagnostics);
            else if (candidate.Math != null) AddMath(target, document, candidate.Math, diagnostics);
            else if (candidate.Verbatim != null) AddVerbatim(target, candidate.Verbatim, diagnostics);
            cursor = candidate.Span.End.Offset;
        }
        if (cursor < end) AddPlain(target, document.Source.Text.Substring(cursor, end - cursor));
    }

    private static void AddCommand(
        InlineSequence target,
        LatexDocument document,
        LatexCommand command,
        List<LatexMarkdownConversionDiagnostic> diagnostics) {
        LatexArgument? first = command.GetRequiredArgument(0);
        LatexCommandSyntaxSignature? signature = LatexProfileSyntaxCatalog.GetCommand(command.Name);
        if (signature != null && command.Arguments.Count(static argument => !argument.IsOptional) < signature.Arguments.Count(static argument => argument == LatexArgumentGroupKind.Required)) {
            target.AddRaw(new CodeSpanInline(LatexToMarkdownConverter.ExtractVisibleSource(document, command.Syntax, command.Syntax.Span, diagnostics)));
            Report(diagnostics, "LATEXMD112", LatexMarkdownConversionOutcome.SourceFallback, "command-arguments:" + command.Name,
                "The bounded profile requires braced arguments; the incomplete or unbraced command was retained as source.", command.Syntax.Span);
            return;
        }
        switch (command.Name) {
            case "textbf":
                target.AddRaw(new BoldSequenceInline(ConvertArgument(document, first, diagnostics)));
                break;
            case "textit":
            case "emph":
                target.AddRaw(new ItalicSequenceInline(ConvertArgument(document, first, diagnostics)));
                break;
            case "texttt":
                target.AddRaw(new CodeSpanInline(ConvertScalarArgument(document, command, first, diagnostics)));
                break;
            case "underline":
                target.AddRaw(new UnderlineInline(ConvertScalarArgument(document, command, first, diagnostics)));
                break;
            case "textsuperscript":
                target.AddRaw(new SuperscriptSequenceInline(ConvertArgument(document, first, diagnostics)));
                break;
            case "textsubscript":
                target.AddRaw(new SubscriptSequenceInline(ConvertArgument(document, first, diagnostics)));
                break;
            case "sout":
                target.AddRaw(new StrikethroughSequenceInline(ConvertArgument(document, first, diagnostics)));
                break;
            case "href": {
                LatexArgument? label = command.GetRequiredArgument(1);
                target.AddRaw(new LinkInline(ConvertArgument(document, label ?? first, diagnostics), LatexLiteralText.Decode(ReadArgumentSource(document, first, diagnostics)), null));
                break;
            }
            case "url":
                string url = LatexLiteralText.Decode(ReadArgumentSource(document, first, diagnostics));
                target.AddRaw(new LinkInline(url, url, null));
                break;
            case "ref":
            case "pageref":
            case "autoref":
            case "eqref":
                string reference = ReadArgumentSource(document, first, diagnostics);
                target.AddRaw(new LinkInline(reference, "#" + reference, null));
                Report(diagnostics, "LATEXMD113", LatexMarkdownConversionOutcome.Simplified, "reference:" + command.Name,
                    "The reference key was retained as a link; TeX counters, page numbers, prefixes, and equation formatting were not evaluated.", command.Syntax.Span);
                break;
            case "cite":
            case "citep":
            case "citet":
                target.AddRaw(new MarkdownTextRun("[" + ReadArgumentSource(document, first, diagnostics) + "]"));
                Report(diagnostics, "LATEXMD102", LatexMarkdownConversionOutcome.Simplified, "citation",
                    "Citation keys were retained as visible text; bibliography style and numbering require a TeX processor.", command.Syntax.Span);
                break;
            case "includegraphics":
                ReportGraphicsOptions(command, diagnostics);
                string image = LatexLiteralText.Decode(ReadArgumentSource(document, first, diagnostics));
                target.AddRaw(new ImageInline(image, image));
                break;
            case "label":
                target.AddRaw(new HtmlRawInline("<a id=\"" + EscapeHtml(ReadArgumentSource(document, first, diagnostics)) + "\"></a>"));
                break;
            case "textbackslash": target.AddRaw(new MarkdownTextRun("\\")); break;
            case "textasciitilde": target.AddRaw(new MarkdownTextRun("~")); break;
            case "textasciicircum": target.AddRaw(new MarkdownTextRun("^")); break;
            case "%": target.AddRaw(new MarkdownTextRun("%")); break;
            case "&": target.AddRaw(new MarkdownTextRun("&")); break;
            case "_": target.AddRaw(new MarkdownTextRun("_")); break;
            case "#": target.AddRaw(new MarkdownTextRun("#")); break;
            case "$": target.AddRaw(new MarkdownTextRun("$")); break;
            case "{": target.AddRaw(new MarkdownTextRun("{")); break;
            case "}": target.AddRaw(new MarkdownTextRun("}")); break;
            case "\\":
            case "newline":
            case "linebreak":
                target.AddRaw(new HardBreakInline());
                break;
            default:
                target.AddRaw(new CodeSpanInline(LatexToMarkdownConverter.ExtractVisibleSource(document, command.Syntax, command.Syntax.Span, diagnostics)));
                Report(diagnostics, "LATEXMD109", LatexMarkdownConversionOutcome.SourceFallback, "command:" + command.Name,
                    "Unknown or package-specific command retained as inline LaTeX source.", command.Syntax.Span);
                break;
        }
    }

    internal static string ReadArgumentSource(LatexDocument document, LatexArgument? argument,
        List<LatexMarkdownConversionDiagnostic> diagnostics) => argument == null ? string.Empty
        : LatexToMarkdownConverter.ExtractVisibleSource(document, argument.Syntax, argument.ContentSpan, diagnostics);

    private static string ConvertScalarArgument(
        LatexDocument document,
        LatexCommand command,
        LatexArgument? argument,
        List<LatexMarkdownConversionDiagnostic> diagnostics) {
        InlineSequence children = ConvertArgument(document, argument, diagnostics);
        if (children.Nodes.Any(static inline => !(inline is MarkdownTextRun))) {
            Report(diagnostics, "LATEXMD115", LatexMarkdownConversionOutcome.Simplified, "inline-formatting:" + command.Name,
                "Nested inline formatting was reduced to plain text inside the target code or underline node.", command.Syntax.Span);
        }
        return InlinePlainText.Extract(children);
    }

    private static InlineSequence ConvertArgument(
        LatexDocument document,
        LatexArgument? argument,
        List<LatexMarkdownConversionDiagnostic> diagnostics) =>
        argument == null
            ? new InlineSequence { AutoSpacing = false }
            : Convert(document, argument.ContentSpan, diagnostics);

    private static void AddMath(
        InlineSequence target,
        LatexDocument document,
        LatexMath math,
        List<LatexMarkdownConversionDiagnostic> diagnostics) {
        target.AddRaw(new CodeSpanInline(LatexToMarkdownConverter.ExtractVisibleSource(document, math.Syntax, math.ContentSpan, diagnostics)));
        Report(diagnostics, "LATEXMD101", LatexMarkdownConversionOutcome.Simplified, "inline-math",
            "LaTeX math source was transported in a code span; TeX layout was not evaluated.", math.Syntax.Span);
    }

    private static void AddVerbatim(
        InlineSequence target,
        LatexSyntaxNode syntax,
        List<LatexMarkdownConversionDiagnostic> diagnostics) {
        if (string.Equals(syntax.Value, "comment", StringComparison.Ordinal)) {
            Report(diagnostics, "LATEXMD110", LatexMarkdownConversionOutcome.Omitted, "comment-environment",
                "The LaTeX comment environment was omitted and its body was not exposed as Markdown text.", syntax.Span);
            return;
        }
        target.AddRaw(new CodeSpanInline(GetVerbatimContent(syntax)));
        Report(diagnostics, "LATEXMD111", LatexMarkdownConversionOutcome.Simplified, "verbatim",
            "Opaque LaTeX verbatim content was retained as code without TeX environment semantics.", syntax.Span);
    }

    internal static string GetVerbatimContent(LatexSyntaxNode syntax) {
        string source = syntax.OriginalText;
        if (string.Equals(syntax.Value, "verb", StringComparison.Ordinal)) {
            int delimiter = source.StartsWith("\\verb*", StringComparison.Ordinal) ? 6 : 5;
            if (source.Length <= delimiter) return string.Empty;
            int contentStart = delimiter + 1;
            bool terminated = source.Length > contentStart && source[source.Length - 1] == source[delimiter];
            int inlineContentEnd = terminated ? source.Length - 1 : source.Length;
            return inlineContentEnd > contentStart ? source.Substring(contentStart, inlineContentEnd - contentStart) : string.Empty;
        }
        if (!LatexVerbatimSyntax.TryReadEnvironmentOpening(
                source, 0, out string environmentName, out int start)
            || !string.Equals(environmentName, syntax.Value, StringComparison.Ordinal)) return source;
        if (string.Equals(syntax.Value, "minted", StringComparison.Ordinal)) {
            int argumentStart = start;
            SkipInterArgumentWhitespace(source, ref argumentStart);
            TrySkipDelimitedArgument(source, ref argumentStart, '[', ']');
            SkipInterArgumentWhitespace(source, ref argumentStart);
            if (TrySkipDelimitedArgument(source, ref argumentStart, '{', '}')) start = argumentStart;
        } else if (string.Equals(syntax.Value, "lstlisting", StringComparison.Ordinal)
            || string.Equals(syntax.Value, "Verbatim", StringComparison.Ordinal)) {
            int argumentStart = start;
            SkipInterArgumentWhitespace(source, ref argumentStart);
            if (TrySkipDelimitedArgument(source, ref argumentStart, '[', ']')) start = argumentStart;
        }
        int contentEnd = LatexVerbatimSyntax.TryFindEnvironmentClosing(
            source, start, environmentName, out int closingStart, out int closingEnd)
            && closingEnd == source.Length
                ? closingStart
                : source.Length;
        int length = contentEnd - start;
        return length > 0 ? source.Substring(start, length) : string.Empty;
    }

    private static void SkipInterArgumentWhitespace(string source, ref int cursor) {
        while (cursor < source.Length && char.IsWhiteSpace(source[cursor])) cursor++;
    }

    private static bool TrySkipDelimitedArgument(string source, ref int cursor, char open, char close) {
        if (cursor >= source.Length || source[cursor] != open) return false;
        int depth = 0;
        for (int index = cursor; index < source.Length; index++) {
            char current = source[index];
            if (current == '\\' && index + 1 < source.Length) {
                index++;
                continue;
            }
            if (current == open) depth++;
            else if (current == close && --depth == 0) {
                cursor = index + 1;
                return true;
            }
        }
        return false;
    }

    private static void AddPlain(InlineSequence target, string value) {
        var text = new StringBuilder();
        for (int index = 0; index < value.Length; index++) {
            char current = value[index];
            if (current == '%') {
                Flush(target, text);
                while (index + 1 < value.Length && value[index + 1] != '\r' && value[index + 1] != '\n') index++;
                if (index + 1 < value.Length && value[index + 1] == '\r') index++;
                if (index + 1 < value.Length && value[index + 1] == '\n') index++;
                continue;
            }
            if (current == '\r' || current == '\n') {
                Flush(target, text);
                if (current == '\r' && index + 1 < value.Length && value[index + 1] == '\n') index++;
                target.AddRaw(new SoftBreakInline());
                continue;
            }
            if (current == '~') text.Append(' ');
            else if (current != '{' && current != '}') text.Append(current);
        }
        Flush(target, text);
    }

    private static void Flush(InlineSequence target, StringBuilder text) {
        if (text.Length == 0) return;
        target.AddRaw(new MarkdownTextRun(text.ToString()));
        text.Clear();
    }

    internal static void ReportGraphicsOptions(LatexCommand command, List<LatexMarkdownConversionDiagnostic> diagnostics) {
        if (string.IsNullOrWhiteSpace(command.GetOptionalArgument(0)?.Content)) return;
        Report(diagnostics, "LATEXMD114", LatexMarkdownConversionOutcome.Simplified, "graphics-options",
            "The graphics resource was retained; TeX size, rotation, crop, and placement options were not evaluated.", command.Syntax.Span);
    }

    private static string EscapeHtml(string value) =>
        value.Replace("&", "&amp;").Replace("\"", "&quot;").Replace("<", "&lt;").Replace(">", "&gt;");

    private static void Report(
        List<LatexMarkdownConversionDiagnostic> diagnostics,
        string code,
        LatexMarkdownConversionOutcome outcome,
        string feature,
        string message,
        LatexSourceSpan span) =>
        diagnostics.Add(new LatexMarkdownConversionDiagnostic(code, outcome, feature, message, span));

    private sealed class InlineCandidate {
        internal InlineCandidate(LatexSourceSpan span, LatexCommand? command, LatexMath? math,
            LatexSyntaxNode? verbatim = null) {
            Span = span;
            Command = command;
            Math = math;
            Verbatim = verbatim;
        }
        internal LatexSourceSpan Span { get; }
        internal LatexCommand? Command { get; }
        internal LatexMath? Math { get; }
        internal LatexSyntaxNode? Verbatim { get; }
    }
}
