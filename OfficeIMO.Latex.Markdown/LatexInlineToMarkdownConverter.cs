namespace OfficeIMO.Latex.Markdown;

internal static class LatexInlineToMarkdownConverter {
    internal static InlineSequence Convert(
        LatexProjectionContext context,
        LatexSourceSpan span,
        List<LatexMarkdownConversionDiagnostic> diagnostics) {
        context.CheckCancellation();
        var target = new InlineSequence { AutoSpacing = false };
        AddRange(target, context, span.Start.Offset, span.End.Offset, diagnostics);
        return target;
    }

    internal static InlineSequence ConvertExcluding(
        LatexProjectionContext context,
        LatexSourceSpan span,
        IEnumerable<LatexSourceSpan> excludedSpans,
        List<LatexMarkdownConversionDiagnostic> diagnostics) {
        return Convert(context.WithExcludedSpans(excludedSpans), span, diagnostics);
    }

    private static void AddRange(
        InlineSequence target,
        LatexProjectionContext context,
        int start,
        int end,
        List<LatexMarkdownConversionDiagnostic> diagnostics) {
        int cursor = start;
        foreach (LatexInlineCandidate candidate in context.InlineCandidates(start, end)) {
            context.CheckCancellation();
            if (candidate.Span.Start.Offset < cursor) continue;
            AddPlain(target, context, cursor, candidate.Span.Start.Offset);
            if (context.IsExcluded(candidate.Span)) { cursor = candidate.Span.End.Offset; continue; }
            if (candidate.Command != null) AddCommand(target, context, candidate.Command, diagnostics);
            else if (candidate.Math != null) AddMath(target, context, candidate.Math, diagnostics);
            else if (candidate.Verbatim != null) AddVerbatim(target, context, candidate.Verbatim, diagnostics);
            cursor = candidate.Span.End.Offset;
        }
        if (cursor < end) AddPlain(target, context, cursor, end);
    }

    private static void AddCommand(
        InlineSequence target,
        LatexProjectionContext context,
        LatexCommand command,
        List<LatexMarkdownConversionDiagnostic> diagnostics) {
        LatexArgument? first = command.GetRequiredArgument(0);
        LatexCommandSyntaxSignature? signature = LatexProfileSyntaxCatalog.GetCommand(command.Name);
        if (signature != null && command.Arguments.Count(static argument => !argument.IsOptional) < signature.Arguments.Count(static argument => argument == LatexArgumentGroupKind.Required)) {
            target.AddRaw(new CodeSpanInline(LatexToMarkdownConverter.ExtractVisibleSource(context, command.Syntax.Span, diagnostics)));
            Report(diagnostics, "LATEXMD112", LatexMarkdownConversionOutcome.SourceFallback, "command-arguments:" + command.Name,
                "The bounded profile requires braced arguments; the incomplete or unbraced command was retained as source.", command.Syntax.Span);
            return;
        }
        switch (command.Name) {
            case "textbf":
                target.AddRaw(new BoldSequenceInline(ConvertArgument(context, first, diagnostics)));
                break;
            case "textit":
            case "emph":
                target.AddRaw(new ItalicSequenceInline(ConvertArgument(context, first, diagnostics)));
                break;
            case "texttt":
                target.AddRaw(new CodeSpanInline(ConvertScalarArgument(context, command, first, diagnostics)));
                break;
            case "underline":
                target.AddRaw(new UnderlineInline(ConvertScalarArgument(context, command, first, diagnostics)));
                break;
            case "textsuperscript":
                target.AddRaw(new SuperscriptSequenceInline(ConvertArgument(context, first, diagnostics)));
                break;
            case "textsubscript":
                target.AddRaw(new SubscriptSequenceInline(ConvertArgument(context, first, diagnostics)));
                break;
            case "sout":
                target.AddRaw(new StrikethroughSequenceInline(ConvertArgument(context, first, diagnostics)));
                break;
            case "href": {
                LatexArgument? label = command.GetRequiredArgument(1);
                target.AddRaw(new LinkInline(ConvertArgument(context, label ?? first, diagnostics), LatexLiteralText.Decode(ReadArgumentSource(context, first, diagnostics), context.CancellationToken), null));
                break;
            }
            case "url":
                string url = LatexLiteralText.Decode(ReadArgumentSource(context, first, diagnostics), context.CancellationToken);
                target.AddRaw(new LinkInline(url, url, null));
                break;
            case "ref":
            case "pageref":
            case "autoref":
            case "eqref":
                string reference = ReadArgumentSource(context, first, diagnostics);
                target.AddRaw(new LinkInline(reference, "#" + reference, null));
                Report(diagnostics, "LATEXMD113", LatexMarkdownConversionOutcome.Simplified, "reference:" + command.Name,
                    "The reference key was retained as a link; TeX counters, page numbers, prefixes, and equation formatting were not evaluated.", command.Syntax.Span);
                break;
            case "cite":
            case "citep":
            case "citet":
                target.AddRaw(new MarkdownTextRun("[" + ReadArgumentSource(context, first, diagnostics) + "]"));
                Report(diagnostics, "LATEXMD102", LatexMarkdownConversionOutcome.Simplified, "citation",
                    "Citation keys were retained as visible text; bibliography style and numbering require a TeX processor.", command.Syntax.Span);
                break;
            case "includegraphics":
                ReportGraphicsOptions(command, diagnostics);
                string image = LatexLiteralText.Decode(ReadArgumentSource(context, first, diagnostics), context.CancellationToken);
                target.AddRaw(new ImageInline(image, image));
                break;
            case "label":
                target.AddRaw(new HtmlRawInline("<a id=\"" + EscapeHtml(ReadArgumentSource(context, first, diagnostics)) + "\"></a>"));
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
                target.AddRaw(new CodeSpanInline(LatexToMarkdownConverter.ExtractVisibleSource(context, command.Syntax.Span, diagnostics)));
                Report(diagnostics, "LATEXMD109", LatexMarkdownConversionOutcome.SourceFallback, "command:" + command.Name,
                    "Unknown or package-specific command retained as inline LaTeX source.", command.Syntax.Span);
                break;
        }
    }

    internal static string ReadArgumentSource(LatexProjectionContext context, LatexArgument? argument,
        List<LatexMarkdownConversionDiagnostic> diagnostics) => argument == null ? string.Empty
        : LatexToMarkdownConverter.ExtractVisibleSource(context, argument.ContentSpan, diagnostics);

    private static string ConvertScalarArgument(
        LatexProjectionContext context,
        LatexCommand command,
        LatexArgument? argument,
        List<LatexMarkdownConversionDiagnostic> diagnostics) {
        return ReadDisplayArgument(context, argument, diagnostics, "inline-formatting:" + command.Name);
    }

    internal static string ReadDisplayArgument(LatexProjectionContext context, LatexArgument? argument,
        List<LatexMarkdownConversionDiagnostic> diagnostics, string feature) {
        var localDiagnostics = new List<LatexMarkdownConversionDiagnostic>();
        InlineSequence children = ConvertArgument(context, argument, localDiagnostics);
        foreach (LatexMarkdownConversionDiagnostic diagnostic in localDiagnostics) {
            diagnostics.Add(diagnostic.Code == "LATEXMD110"
                ? new LatexMarkdownConversionDiagnostic("LATEXMD210", diagnostic.Outcome, diagnostic.Feature, diagnostic.Message, diagnostic.LatexSpan)
                : diagnostic);
        }
        if (children.Nodes.Any(static inline => !(inline is MarkdownTextRun) && !(inline is SoftBreakInline))) {
            Report(diagnostics, "LATEXMD115", LatexMarkdownConversionOutcome.Simplified, feature,
                "Inline formatting was reduced to its visible plain text in the target scalar value.", argument!.ContentSpan);
        } else if (argument != null) {
            // The literal route also retains the original CR/LF convention in metadata.
            string visible = context.ExtractResidual(argument.ContentSpan, context.Comments(argument.ContentSpan).Select(static comment => comment.Span));
            return LatexLiteralText.Decode(visible, context.CancellationToken, preserveTilde: false);
        }
        return InlinePlainText.Extract(children);
    }

    private static InlineSequence ConvertArgument(
        LatexProjectionContext context,
        LatexArgument? argument,
        List<LatexMarkdownConversionDiagnostic> diagnostics) =>
        argument == null
            ? new InlineSequence { AutoSpacing = false }
            : Convert(context, argument.ContentSpan, diagnostics);

    private static void AddMath(
        InlineSequence target,
        LatexProjectionContext context,
        LatexMath math,
        List<LatexMarkdownConversionDiagnostic> diagnostics) {
        target.AddRaw(new CodeSpanInline(LatexToMarkdownConverter.ExtractVisibleSource(context, math.ContentSpan, diagnostics)));
        Report(diagnostics, "LATEXMD101", LatexMarkdownConversionOutcome.Simplified, "inline-math",
            "LaTeX math source was transported in a code span; TeX layout was not evaluated.", math.Syntax.Span);
    }

    private static void AddVerbatim(
        InlineSequence target,
        LatexProjectionContext context,
        LatexSyntaxNode syntax,
        List<LatexMarkdownConversionDiagnostic> diagnostics) {
        if (string.Equals(syntax.Value, "comment", StringComparison.Ordinal)) {
            Report(diagnostics, "LATEXMD110", LatexMarkdownConversionOutcome.Omitted, "comment-environment",
                "The LaTeX comment environment was omitted and its body was not exposed as Markdown text.", syntax.Span);
            return;
        }
        target.AddRaw(new CodeSpanInline(GetVerbatimContent(context, syntax)));
        Report(diagnostics, "LATEXMD111", LatexMarkdownConversionOutcome.Simplified, "verbatim",
            "Opaque LaTeX verbatim content was retained as code without TeX environment semantics.", syntax.Span);
    }

    internal static string GetVerbatimContent(LatexProjectionContext context, LatexSyntaxNode syntax) {
        context.CheckCancellation();
        string source = syntax.OriginalText;
        context.CheckCancellation();
        if (string.Equals(syntax.Value, "verb", StringComparison.Ordinal)) {
            int delimiter = source.StartsWith("\\verb*", StringComparison.Ordinal) ? 6 : 5;
            if (source.Length <= delimiter) return string.Empty;
            int contentStart = delimiter + 1;
            bool terminated = source.Length > contentStart && source[source.Length - 1] == source[delimiter];
            int inlineContentEnd = terminated ? source.Length - 1 : source.Length;
            return inlineContentEnd > contentStart ? source.Substring(contentStart, inlineContentEnd - contentStart) : string.Empty;
        }
        if (!LatexVerbatimSyntax.TryReadEnvironmentOpening(
                source, 0, out string environmentName, out int start, context.CancellationToken)
            || !string.Equals(environmentName, syntax.Value, StringComparison.Ordinal)) return source;
        if (string.Equals(syntax.Value, "minted", StringComparison.Ordinal)) {
            int argumentStart = start;
            SkipInterArgumentWhitespace(context, source, ref argumentStart);
            TrySkipDelimitedArgument(context, source, ref argumentStart, '[', ']');
            SkipInterArgumentWhitespace(context, source, ref argumentStart);
            if (TrySkipDelimitedArgument(context, source, ref argumentStart, '{', '}')) start = argumentStart;
        } else if (string.Equals(syntax.Value, "lstlisting", StringComparison.Ordinal)
            || string.Equals(syntax.Value, "Verbatim", StringComparison.Ordinal)) {
            int argumentStart = start;
            SkipInterArgumentWhitespace(context, source, ref argumentStart);
            if (TrySkipDelimitedArgument(context, source, ref argumentStart, '[', ']')) start = argumentStart;
        }
        int contentEnd = LatexVerbatimSyntax.TryFindEnvironmentClosing(
            source, start, environmentName, out int closingStart, out int closingEnd, context.CancellationToken)
            && closingEnd == source.Length
                ? closingStart
                : source.Length;
        int length = contentEnd - start;
        string result = length > 0 ? source.Substring(start, length) : string.Empty;
        context.CheckCancellation();
        return result;
    }

    private static void SkipInterArgumentWhitespace(LatexProjectionContext context, string source, ref int cursor) {
        while (cursor < source.Length && char.IsWhiteSpace(source[cursor])) {
            if ((cursor & 1023) == 0) context.CheckCancellation();
            cursor++;
        }
    }

    private static bool TrySkipDelimitedArgument(LatexProjectionContext context, string source, ref int cursor, char open, char close) {
        if (cursor >= source.Length || source[cursor] != open) return false;
        int depth = 0;
        for (int index = cursor; index < source.Length; index++) {
            if ((index & 1023) == 0) context.CheckCancellation();
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

    private static void AddPlain(InlineSequence target, LatexProjectionContext context, int start, int end) {
        string value = context.Document.Source.Text;
        var text = new StringBuilder();
        for (int index = start; index < end; index++) {
            if ((index & 1023) == 0) context.CheckCancellation();
            char current = value[index];
            if (current == '%') {
                Flush(target, text);
                while (index + 1 < end && value[index + 1] != '\r' && value[index + 1] != '\n') { if ((index & 1023) == 0) context.CheckCancellation(); index++; }
                if (index + 1 < end && value[index + 1] == '\r') index++;
                if (index + 1 < end && value[index + 1] == '\n') index++;
                continue;
            }
            if (current == '\r' || current == '\n') {
                Flush(target, text);
                if (current == '\r' && index + 1 < end && value[index + 1] == '\n') index++;
                target.AddRaw(new SoftBreakInline());
                continue;
            }
            if (current == '~') text.Append(' ');
            else if (current != '{' && current != '}') text.Append(current);
        }
        Flush(target, text);
        context.CheckCancellation();
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

}
