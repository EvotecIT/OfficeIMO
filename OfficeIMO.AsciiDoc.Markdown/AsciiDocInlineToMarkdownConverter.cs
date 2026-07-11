namespace OfficeIMO.AsciiDoc.Markdown;

internal static class AsciiDocInlineToMarkdownConverter {
    internal static InlineSequence Convert(
        AsciiDocInlineSequence source,
        AsciiDocDocumentAttributes attributes,
        AsciiDocToMarkdownOptions options,
        List<AsciiDocMarkdownConversionDiagnostic> diagnostics,
        AsciiDocBlock owner) {
        var target = new InlineSequence { AutoSpacing = false };
        for (int index = 0; index < source.Items.Count; index++) {
            Add(target, source.Items[index], attributes, options, diagnostics, owner);
        }
        return target;
    }

    private static void Add(
        InlineSequence target,
        AsciiDocInline source,
        AsciiDocDocumentAttributes attributes,
        AsciiDocToMarkdownOptions options,
        List<AsciiDocMarkdownConversionDiagnostic> diagnostics,
        AsciiDocBlock owner) {
        switch (source) {
            case AsciiDocTextInline text:
                AddText(target, Unescape(text.Text));
                break;
            case AsciiDocFormattedInline formatted:
                AddFormatted(target, formatted, attributes, options, diagnostics, owner);
                break;
            case AsciiDocAttributeReferenceInline reference:
                AddAttributeReference(target, reference, attributes, options, diagnostics, owner);
                break;
            case AsciiDocCrossReferenceInline crossReference:
                target.AddRaw(new LinkInline(
                    crossReference.Text ?? crossReference.Target,
                    NormalizeCrossReference(crossReference.Target),
                    null));
                break;
            case AsciiDocAnchorInline anchor:
                target.AddRaw(new HtmlRawInline("<a id=\"" + EscapeHtmlAttribute(anchor.Id) + "\"></a>"));
                break;
            case AsciiDocStemInline stem:
                target.AddRaw(new CodeSpanInline(stem.Expression));
                Report(diagnostics, owner, "ADOCMD103", "inline-stem", "Inline STEM converted to a code span because Markdown has no portable math inline contract.");
                break;
            case AsciiDocMacroInline macro:
                AddMacro(target, macro, diagnostics, owner);
                break;
            case AsciiDocPassthroughInline passthrough:
                AddText(target, passthrough.Content);
                break;
            default:
                AddText(target, source.OriginalText);
                Report(diagnostics, owner, "ADOCMD109", source.GetType().Name, "Inline source was retained as visible text.");
                break;
        }
    }

    private static void AddFormatted(
        InlineSequence target,
        AsciiDocFormattedInline source,
        AsciiDocDocumentAttributes attributes,
        AsciiDocToMarkdownOptions options,
        List<AsciiDocMarkdownConversionDiagnostic> diagnostics,
        AsciiDocBlock owner) {
        InlineSequence nested = Convert(source.Content, attributes, options, diagnostics, owner);
        switch (source.Style) {
            case AsciiDocInlineStyle.Strong: target.AddRaw(new BoldSequenceInline(nested)); break;
            case AsciiDocInlineStyle.Emphasis: target.AddRaw(new ItalicSequenceInline(nested)); break;
            case AsciiDocInlineStyle.Monospace: target.AddRaw(new CodeSpanInline(PlainText(source.Content))); break;
            case AsciiDocInlineStyle.Highlight: target.AddRaw(new HighlightSequenceInline(nested)); break;
            case AsciiDocInlineStyle.Subscript: target.AddRaw(new SubscriptSequenceInline(nested)); break;
            case AsciiDocInlineStyle.Superscript: target.AddRaw(new SuperscriptSequenceInline(nested)); break;
        }
    }

    private static void AddAttributeReference(
        InlineSequence target,
        AsciiDocAttributeReferenceInline source,
        AsciiDocDocumentAttributes attributes,
        AsciiDocToMarkdownOptions options,
        List<AsciiDocMarkdownConversionDiagnostic> diagnostics,
        AsciiDocBlock owner) {
        if (!options.ExpandDocumentAttributes) {
            AddText(target, source.OriginalText);
            return;
        }
        var substitutionOptions = new AsciiDocAttributeSubstitutionOptions {
            UndefinedAttributeBehavior = options.UndefinedAttributeBehavior
        };
        AsciiDocAttributeSubstitutionResult result = AsciiDocAttributeSubstitutor.Substitute(source.OriginalText, attributes, substitutionOptions);
        AddText(target, result.Value);
        for (int index = 0; index < result.Diagnostics.Count; index++) {
            Report(diagnostics, owner, "ADOCMD101", "attribute-reference", result.Diagnostics[index].Message);
        }
    }

    private static void AddMacro(
        InlineSequence target,
        AsciiDocMacroInline source,
        List<AsciiDocMarkdownConversionDiagnostic> diagnostics,
        AsciiDocBlock owner) {
        string label = FirstAttribute(source.AttributeList) ?? source.Target;
        if (string.Equals(source.Name, "image", StringComparison.Ordinal)) {
            target.AddRaw(new ImageInline(label, source.Target));
            return;
        }
        if (string.Equals(source.Name, "link", StringComparison.Ordinal)) {
            target.AddRaw(new LinkInline(label, source.Target, null));
            return;
        }
        if (string.Equals(source.Name, "xref", StringComparison.Ordinal)) {
            target.AddRaw(new LinkInline(label, NormalizeCrossReference(source.Target), null));
            return;
        }
        AddText(target, source.OriginalText);
        Report(diagnostics, owner, "ADOCMD102", "inline-macro:" + source.Name, "Unknown inline macro retained as visible source text.");
    }

    private static void AddText(InlineSequence target, string value) {
        int start = 0;
        for (int index = 0; index < value.Length; index++) {
            if (value[index] != '\r' && value[index] != '\n') continue;
            if (index > start) target.AddRaw(new TextRun(value.Substring(start, index - start)));
            if (value[index] == '\r' && index + 1 < value.Length && value[index + 1] == '\n') index++;
            target.AddRaw(new SoftBreakInline());
            start = index + 1;
        }
        if (start < value.Length) target.AddRaw(new TextRun(value.Substring(start)));
    }

    private static string PlainText(AsciiDocInlineSequence sequence) {
        var output = new System.Text.StringBuilder();
        for (int index = 0; index < sequence.Items.Count; index++) {
            AsciiDocInline item = sequence.Items[index];
            if (item is AsciiDocTextInline text) output.Append(text.Text);
            else if (item is AsciiDocFormattedInline formatted) output.Append(PlainText(formatted.Content));
            else if (item is AsciiDocPassthroughInline pass) output.Append(pass.Content);
            else output.Append(item.OriginalText);
        }
        return output.ToString();
    }

    private static string Unescape(string value) {
        var output = new System.Text.StringBuilder(value.Length);
        for (int index = 0; index < value.Length; index++) {
            if (value[index] == '\\' && index + 1 < value.Length && IsEscapable(value[index + 1])) index++;
            output.Append(value[index]);
        }
        return output.ToString();
    }

    private static bool IsEscapable(char value) => "*_'`#+~^{}[]<>\\".IndexOf(value) >= 0;

    private static string NormalizeCrossReference(string target) {
        if (target.IndexOf('#') >= 0 || target.IndexOf('/') >= 0 || target.IndexOf('.') >= 0) return target;
        return "#" + target;
    }

    private static string EscapeHtmlAttribute(string value) =>
        value.Replace("&", "&amp;").Replace("\"", "&quot;").Replace("<", "&lt;").Replace(">", "&gt;");

    private static string? FirstAttribute(string value) {
        if (string.IsNullOrWhiteSpace(value)) return null;
        AsciiDocElementAttributes attributes = AsciiDocAttributeListParser.Parse(value);
        return attributes.Entries.FirstOrDefault(static entry => entry.Kind == AsciiDocElementAttributeKind.Positional)?.Value;
    }

    private static void Report(
        List<AsciiDocMarkdownConversionDiagnostic> diagnostics,
        AsciiDocBlock owner,
        string code,
        string feature,
        string message) {
        diagnostics.Add(new AsciiDocMarkdownConversionDiagnostic(
            code,
            AsciiDocMarkdownDiagnosticSeverity.Warning,
            AsciiDocMarkdownConversionOutcome.Simplified,
            feature,
            message,
            owner.Span));
    }
}
