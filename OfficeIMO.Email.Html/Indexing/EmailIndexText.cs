using OfficeIMO.Html.Dom;

namespace OfficeIMO.Email;

/// <summary>Creates network-free indexing text with conservative, inspectable mail-region markers.</summary>
public static class EmailIndexText {
    /// <summary>Projects text without opening attachments or modifying the source. This is not an authorship detector.</summary>
    public static EmailIndexTextResult Create(EmailDocument source, EmailIndexTextOptions? options = null) {
        if (source == null) throw new ArgumentNullException(nameof(source));
        var effective = options ?? new EmailIndexTextOptions();
        if (effective.MaxSourceChars <= 0) throw new ArgumentOutOfRangeException(nameof(effective.MaxSourceChars));
        if (effective.MaxProjectionChars <= 0) throw new ArgumentOutOfRangeException(nameof(effective.MaxProjectionChars));
        if (effective.MaxTextChars <= 0) throw new ArgumentOutOfRangeException(nameof(effective.MaxTextChars));
        var builder = new EmailIndexTextBuilder(effective.MaxTextChars);
        var diagnostics = new List<EmailDiagnostic>();
        EmailBodySourceKind kind;
        if (!string.IsNullOrEmpty(source.Body.Text) && (effective.PreferPlainText ||
            string.IsNullOrWhiteSpace(source.Body.Html) && string.IsNullOrWhiteSpace(source.Body.Rtf))) {
            kind = EmailBodySourceKind.PlainText;
            CheckSource(source.Body.Text!);
            bool signature = false;
            using var reader = new StringReader(source.Body.Text!);
            string? line;
            while ((line = reader.ReadLine()) != null && !builder.Truncated) {
                bool quoted = line.TrimStart().StartsWith(">", StringComparison.Ordinal);
                if (!quoted && line == "-- ") signature = true;
                var region = quoted ? EmailIndexRegionKind.Quoted : signature ? EmailIndexRegionKind.Signature : EmailIndexRegionKind.Unclassified;
                string reason = quoted ? "plain-quote-prefix" : signature ? "plain-signature-separator" : "unclassified";
                builder.Append(line, region, reason, preserveWhitespace: true);
                builder.LineBreak(region, reason);
            }
        } else {
            CheckSource(!string.IsNullOrWhiteSpace(source.Body.Html) ? source.Body.Html! : source.Body.Rtf ?? string.Empty);
            var htmlOptions = HtmlConversionDocumentOptions.CreateUntrustedProfile();
            htmlOptions.Limits.MaxInputCharacters = effective.MaxProjectionChars;
            var projection = EmailBodyProjection.Create(source, new EmailBodyProjectionOptions {
                IncludeResources = false, MaxBodySourceCharacters = effective.MaxSourceChars, HtmlOptions = htmlOptions });
            kind = projection.SourceKind;
            diagnostics.AddRange(projection.Diagnostics);
            if (kind != EmailBodySourceKind.None) {
                var document = projection.Document.CreateDocumentForConversion();
                Walk(document.Body ?? (HtmlNode)document, EmailIndexRegionKind.Unclassified, "unclassified", false);
            }
        }
        if (builder.Truncated) diagnostics.Add(new EmailDiagnostic("EMAIL_INDEX_TEXT_TRUNCATED", "Indexing text reached its character limit.", location: "message/body"));
        return builder.Build(kind, effective.ExcludeQuotes, effective.ExcludeSignatures, diagnostics.AsReadOnly());

        void CheckSource(string value) {
            if (value.Length > effective.MaxSourceChars) throw new ArgumentException("The selected email body exceeds MaxSourceChars.", nameof(source));
        }
        void Walk(HtmlNode node, EmailIndexRegionKind region, string reason, bool preformatted) {
            if (builder.Truncated) return;
            if (node.Kind == HtmlNodeKind.Text) { builder.Append(node.TextContent, region, reason, preformatted); return; }
            var element = node as HtmlElement;
            string name = element?.LocalName ?? string.Empty;
            if (name == "head" || name == "script" || name == "style" || name == "template" ||
                element?.HasAttribute("hidden") == true || element?.GetAttribute("aria-hidden") == "true") return;
            bool block = IsBlock(name);
            if (block || name == "br") builder.LineBreak(region, reason);
            if (region == EmailIndexRegionKind.Unclassified && element != null) {
                string[] classes = (element.GetAttribute("class") ?? string.Empty).Split(new[] { ' ', '\t', '\r', '\n', '\f' }, StringSplitOptions.RemoveEmptyEntries);
                string? quote = classes.FirstOrDefault(value => value == "gmail_quote" || value == "yahoo_quoted");
                string? signature = classes.FirstOrDefault(value => value == "gmail_signature" || value == "moz-signature");
                if (name == "blockquote" || quote != null) { region = EmailIndexRegionKind.Quoted; reason = quote ?? "blockquote"; }
                else if (signature != null) { region = EmailIndexRegionKind.Signature; reason = signature; }
            }
            foreach (HtmlNode child in node.ChildNodes) {
                Walk(child, region, reason, preformatted || name == "pre");
                if (builder.Truncated) break;
            }
            if (block) builder.LineBreak(region, reason);
            else if (name == "td" || name == "th") builder.Append(" ", region, reason, false);
        }
    }

    private static bool IsBlock(string name) => name == "p" || name == "div" || name == "blockquote" ||
        name == "pre" || name == "li" || name == "tr" || name == "section" || name == "article" ||
        name == "header" || name == "footer" || name == "h1" || name == "h2" || name == "h3" ||
        name == "h4" || name == "h5" || name == "h6";
}
