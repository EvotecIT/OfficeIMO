using OfficeIMO.Html.Dom;

namespace OfficeIMO.Email;

/// <summary>Controls safe, resource-free HTML quotation alongside the core composition policy.</summary>
public sealed class EmailHtmlCompositionOptions {
    /// <summary>Recipient, threading, timestamp and quotation policy.</summary>
    public EmailCompositionOptions Composition { get; set; } = new EmailCompositionOptions();
    /// <summary>Maximum original body characters accepted before HTML/RTF projection.</summary>
    public int MaxSourceChars { get; set; } = 2 * 1024 * 1024;
    /// <summary>Maximum final HTML characters, including authored text, quotation, encoding and document markup.</summary>
    public int MaxProjectionChars { get; set; } = 16 * 1024 * 1024;
}

/// <summary>A draft containing safe HTML, a plain-text alternative, and composition evidence.</summary>
public sealed class EmailHtmlCompositionResult {
    internal EmailHtmlCompositionResult(EmailDocument document, IReadOnlyList<EmailDiagnostic> diagnostics) {
        Document = document; Diagnostics = diagnostics;
    }
    /// <summary>Independent draft; recipients and threading are chosen by the core composer.</summary>
    public EmailDocument Document { get; }
    /// <summary>Recipient, quotation and omitted resource diagnostics.</summary>
    public IReadOnlyList<EmailDiagnostic> Diagnostics { get; }
}

/// <summary>Quotes the original through the canonical untrusted HTML policy without copying attachments or loading resources.</summary>
public static class EmailHtmlComposer {
    /// <summary>Creates a reply with safely encoded authored text and a bounded HTML quotation.</summary>
    public static EmailHtmlCompositionResult Reply(EmailDocument original, EmailAddress from, string text,
        EmailHtmlCompositionOptions? options = null) => Compose(original, from, text, 0, options);
    /// <summary>Creates a reply to all intended original recipients, with safe HTML quotation.</summary>
    public static EmailHtmlCompositionResult ReplyAll(EmailDocument original, EmailAddress from, string text,
        EmailHtmlCompositionOptions? options = null) => Compose(original, from, text, 1, options);
    /// <summary>Creates a forward draft with safe HTML quotation and no inherited recipients or attachments.</summary>
    public static EmailHtmlCompositionResult Forward(EmailDocument original, EmailAddress from, string text,
        EmailHtmlCompositionOptions? options = null) => Compose(original, from, text, 2, options);

    private static EmailHtmlCompositionResult Compose(EmailDocument original, EmailAddress from, string text,
        int mode, EmailHtmlCompositionOptions? options) {
        var effective = options ?? new EmailHtmlCompositionOptions();
        if (effective.MaxSourceChars <= 0) throw new ArgumentOutOfRangeException(nameof(options), "MaxSourceChars must be positive.");
        if (effective.MaxProjectionChars <= 0) throw new ArgumentOutOfRangeException(nameof(options), "MaxProjectionChars must be positive.");
        var policy = effective.Composition ?? throw new ArgumentException("A composition policy is required.", nameof(options));
        var corePolicy = new EmailCompositionOptions { QuoteOriginal = false, MaxQuoteChars = policy.MaxQuoteChars,
            MaxReferences = policy.MaxReferences, Date = policy.Date };
        foreach (string address in policy.OwnAddresses) corePolicy.OwnAddresses.Add(address);
        var composition = mode == 0 ? EmailComposer.Reply(original, from, text, corePolicy)
            : mode == 1 ? EmailComposer.ReplyAll(original, from, text, corePolicy) : EmailComposer.Forward(original, from, text, corePolicy);
        var diagnostics = new List<EmailDiagnostic>(composition.Diagnostics);
        string html = TextHtml(text, effective.MaxProjectionChars);
        if (policy.QuoteOriginal && policy.MaxQuoteChars > 0) {
            // Prefer the rich source for both alternatives so they quote the same representation.
            var index = EmailIndexText.Create(original, new EmailIndexTextOptions {
                PreferPlainText = false, MaxSourceChars = effective.MaxSourceChars,
                MaxProjectionChars = effective.MaxProjectionChars, MaxTextChars = policy.MaxQuoteChars });
            diagnostics.AddRange(index.Diagnostics);
            if (index.SourceKind == EmailBodySourceKind.None) return Finish();
            if (index.Truncated || index.SourceKind == EmailBodySourceKind.PlainText) {
                AddQuote(TextHtml(index.FullText, effective.MaxProjectionChars));
                if (index.Truncated) diagnostics.Add(new EmailDiagnostic("EMAIL_COMPOSITION_TEXT_QUOTE_FALLBACK", "The rich quotation exceeded its bound; bounded text was quoted instead."));
                return Finish();
            }
            var htmlPolicy = HtmlConversionDocumentOptions.CreateUntrustedProfile();
            htmlPolicy.Limits.MaxInputCharacters = effective.MaxProjectionChars;
            var projection = EmailBodyProjection.Create(original, new EmailBodyProjectionOptions {
                IncludeResources = false, IncludeResourceReferences = false,
                MaxBodySourceCharacters = effective.MaxSourceChars, HtmlOptions = htmlPolicy });
            diagnostics.AddRange(projection.Diagnostics);
            var document = projection.Document.CreateDocumentForConversion();
            var body = document.Body;
            // Outlook conditional comments can carry hidden VML resources outside the URL policy's element tree.
            foreach (var comment in document.Descendants().Where(node => node.Kind == HtmlNodeKind.Comment).ToArray()) comment.Remove();
            foreach (var element in document.Descendants().OfType<HtmlElement>().ToArray()) {
                element.RemoveAttribute("style");
                string name = element.LocalName;
                if (name == "img" || name == "picture" || name == "source" || name == "audio" || name == "video" ||
                    name == "object" || name == "embed" || name == "iframe" || name == "svg" || name == "style" || name == "link" ||
                    name == "input" || name == "button" || name == "select" || name == "textarea" || name == "meta") {
                    element.Remove();
                }
            }
            diagnostics.Add(new EmailDiagnostic("EMAIL_COMPOSITION_RESOURCES_OMITTED",
                "Resource references, resource elements, form controls and inline attachments are excluded by quotation policy."));
            string quote = body?.InnerHtml ?? string.Empty;
            // Never clip markup mid-tag. A large rich quote becomes a scalar-safe text quotation.
            if (index.Truncated || quote.Length > policy.MaxQuoteChars) {
                quote = TextHtml(index.FullText, effective.MaxProjectionChars);
                diagnostics.Add(new EmailDiagnostic("EMAIL_COMPOSITION_TEXT_QUOTE_FALLBACK", "The rich quotation exceeded its bound; bounded text was quoted instead."));
            }
            AddQuote(quote);

            void AddQuote(string quotation) {
                EnsureProjectionLength((long)html.Length + quotation.Length + "<blockquote></blockquote>".Length, effective.MaxProjectionChars);
                html += "<blockquote>" + quotation + "</blockquote>";
                composition.Document.Body.Text += "\r\n\r\n" + string.Join("\r\n", index.FullText.TrimEnd('\n').Split('\n').Select(line => "> " + line));
            }
        }
        return Finish();

        EmailHtmlCompositionResult Finish() {
            const string prefix = "<!DOCTYPE html><html><head><meta charset=\"utf-8\"></head><body>";
            const string suffix = "</body></html>";
            EnsureProjectionLength((long)prefix.Length + html.Length + suffix.Length, effective.MaxProjectionChars);
            composition.Document.Body.Html = prefix + html + suffix;
            composition.Document.Body.HtmlCharset = "utf-8";
            return new EmailHtmlCompositionResult(composition.Document, diagnostics.AsReadOnly());
        }
    }
    private static string TextHtml(string text, int maximum) {
        if (text == null) throw new ArgumentNullException(nameof(text));
        EnsureProjectionLength(text.Length, maximum);
        string normalized = text.Replace("\r\n", "\n").Replace('\r', '\n');
        var output = new System.Text.StringBuilder("<div>");
        for (int offset = 0; offset < normalized.Length;) {
            int count = Math.Min(4096, normalized.Length - offset);
            if (offset + count < normalized.Length && char.IsHighSurrogate(normalized[offset + count - 1])) count--;
            string encoded = OfficeHtmlText.Escape(normalized.Substring(offset, count)).Replace("\n", "<br>");
            EnsureProjectionLength((long)output.Length + encoded.Length + "</div>".Length, maximum);
            output.Append(encoded);
            offset += count;
        }
        EnsureProjectionLength((long)output.Length + "</div>".Length, maximum);
        return output.Append("</div>").ToString();
    }

    private static void EnsureProjectionLength(long length, int maximum) {
        if (length > maximum) throw new InvalidDataException("The composed HTML exceeds MaxProjectionChars.");
    }
}
