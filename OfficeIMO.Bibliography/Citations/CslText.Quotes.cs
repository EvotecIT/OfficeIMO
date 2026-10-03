using System.Xml.Linq;

namespace OfficeIMO.Bibliography;

internal readonly partial struct CslText {
    /// <summary>Resolves style and input quotation levels while retaining formatting inside them.</summary>
    internal CslText ResolveQuotes(CslLocale locale, int maximumCharacters, CancellationToken token) {
        if (!Html.Contains("data-csl-quote") && Html.IndexOfAny(new[] { '‘', '’', '“', '”' }) < 0 &&
            !Html.Contains("&quot;") && !Html.Contains("&#39;")) return this;
        XElement root = CslStyle.ReadXml("<root>" + Html + "</root>", (int)Math.Min(int.MaxValue, (long)maximumCharacters + 13), 1024, token);
        ResolveQuoteElements(root, locale, 0, maximumCharacters, token);
        string html = string.Concat(root.Nodes().Select(node => node is XText text ? Escape(text.Value) : node.ToString(SaveOptions.DisableFormatting)));
        if (html.Length > maximumCharacters) throw new InvalidDataException("CSL rendering exceeds MaximumIntermediateCharacters.");
        return new CslText(HtmlPlain(html, token), html, Attempted, Rendered);
    }

    private static void ResolveQuoteElements(XElement element, CslLocale locale, int level, int maximumCharacters, CancellationToken token) {
        token.ThrowIfCancellationRequested();
        bool quoted = (string?)element.Attribute("data-csl-quote") == "true";
        if (quoted) {
            // Decorate owns both boundary text nodes; input quotes within the
            // body remain distinct and are resolved at the next quotation level.
            string open = locale.Term("open-quote"), close = locale.Term("close-quote");
            if (open.Length > 0) {
                XText first = (XText)element.FirstNode!;
                first.Value = first.Value.Substring(open.Length);
            }
            if (close.Length > 0) {
                XText last = (XText)element.LastNode!;
                last.Value = last.Value.Substring(0, last.Value.Length - close.Length);
            }
        }
        if ((string?)element.Attribute("data-csl-input") == "true") ResolveInputQuotes(element, locale, level, maximumCharacters, token);
        else foreach (XElement child in element.Elements()) ResolveQuoteElements(child, locale, level + (quoted ? 1 : 0), maximumCharacters, token);
        if (quoted) {
            bool displayQuotes = (string?)element.Attribute("data-csl-display-quotes") == "true";
            string open = QuoteTerm(locale, level, false);
            element.AddFirst(displayQuotes ? new XElement("span", new XAttribute("data-csl-display-edge", "start"), open) : new XText(open));
            XElement close = ClosingQuote(QuoteTerm(locale, level, true));
            if (displayQuotes) close.Add(new XAttribute("data-csl-display-edge", "end"));
            element.Add(close);
            element.Attribute("data-csl-quote")!.Remove();
            element.Attribute("data-csl-display-quotes")?.Remove();
        }
    }

    private static string QuoteTerm(CslLocale locale, int level, bool closing) =>
        locale.Term((closing ? "close-" : "open-") + ((level & 1) == 0 ? "quote" : "inner-quote"));

    // Keep provenance until punctuation is localized: an apostrophe can have
    // exactly the same spelling as a locale's closing quotation mark.
    private static XElement ClosingQuote(string value) => new XElement("span", new XAttribute("data-csl-quote-end", "true"), value);

    private static void ResolveInputQuotes(XElement element, CslLocale locale, int level, int maximumCharacters, CancellationToken token) {
        XText[] nodes = element.DescendantNodes().OfType<XText>().ToArray();
        string source = string.Concat(nodes.Select(node => node.Value));
        var replacements = new Dictionary<int, string>();
        var pending = new Stack<(int Offset, int Kind)>();
        var openings = new HashSet<int>();
        var closings = new HashSet<int>();
        for (int index = 0; index < source.Length; index++) {
            if ((index & 1023) == 0) token.ThrowIfCancellationRequested();
            char current = source[index];
            int kind = current == '\'' || current == '‘' || current == '’' ? 1 : current == '"' || current == '“' || current == '”' ? 2 : 0;
            if (kind == 0) continue;
            char previous = index == 0 ? '\0' : source[index - 1];
            char next = index + 1 == source.Length ? '\0' : source[index + 1];
            bool apostrophe = kind == 1 && (char.IsLetterOrDigit(previous) && (char.IsLetterOrDigit(next) || next == '\'') ||
                current == '’' && char.IsLetterOrDigit(previous) && char.IsLetterOrDigit(next));
            if (apostrophe) { if (current == '\'') replacements[index] = "’"; continue; }
            bool closing = current == '’' || current == '”' ||
                (current == '\'' || current == '"') && !char.IsWhiteSpace(previous) && previous != '\0' &&
                (next == '\0' || char.IsWhiteSpace(next) || char.IsPunctuation(next));
            bool opening = current == '‘' || current == '“' ||
                (current == '\'' || current == '"') && next != '\0' && !char.IsWhiteSpace(next) && next != current &&
                (previous == '\0' || char.IsWhiteSpace(previous) || char.IsPunctuation(previous));
            if (closing && pending.Any(candidate => candidate.Kind == kind)) {
                // An unmatched inner candidate (for example the apostrophe in
                // "ETFA '09") must not prevent a complete enclosing pair.
                while (pending.Peek().Kind != kind) {
                    var unmatched = pending.Pop();
                    if (source[unmatched.Offset] == '\'') replacements[unmatched.Offset] = "’";
                }
                var start = pending.Pop();
                openings.Add(start.Offset);
                closings.Add(index);
            } else if (opening) {
                if (pending.Count >= 64) throw new InvalidDataException("CSL input quotation nesting exceeds 64.");
                pending.Push((index, kind));
            } else if (current == '\'') replacements[index] = "’";
        }
        // A lone single quote is an apostrophe (including abbreviated years).
        // Unmatched double quotes retain the caller's original text.
        foreach (var unmatched in pending) if (source[unmatched.Offset] == '\'') replacements[unmatched.Offset] = "’";
        int matchedLevel = level;
        for (int index = 0; index < source.Length; index++) {
            if ((index & 1023) == 0) token.ThrowIfCancellationRequested();
            if (openings.Contains(index)) replacements[index] = QuoteTerm(locale, matchedLevel++, false);
            else if (closings.Contains(index)) replacements[index] = QuoteTerm(locale, --matchedLevel, true);
        }
        int offset = 0;
        long total = 0;
        foreach (XText node in nodes) {
            token.ThrowIfCancellationRequested();
            string text = node.Value;
            var output = new StringBuilder(text.Length);
            var segments = new List<XNode>();
            for (int index = 0; index < text.Length; index++) {
                if ((index & 1023) == 0) token.ThrowIfCancellationRequested();
                bool replaced = replacements.TryGetValue(offset + index, out string? replacement);
                total += replaced ? replacement!.Length : 1;
                if (total > maximumCharacters) throw new InvalidDataException("CSL rendering exceeds MaximumIntermediateCharacters.");
                if (closings.Contains(offset + index)) {
                    if (output.Length > 0) { segments.Add(new XText(output.ToString())); output.Clear(); }
                    segments.Add(ClosingQuote(replacement!));
                } else if (replaced) output.Append(replacement);
                else output.Append(text[index]);
            }
            offset += text.Length;
            if (segments.Count == 0) node.Value = output.ToString();
            else {
                if (output.Length > 0) segments.Add(new XText(output.ToString()));
                node.ReplaceWith(segments);
            }
        }
    }
}
