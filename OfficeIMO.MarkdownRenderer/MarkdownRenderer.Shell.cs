using OfficeIMO.Markdown;

namespace OfficeIMO.MarkdownRenderer;

public static partial class MarkdownRenderer {
    private static string BuildOverflowBodyHtml(HtmlOptions htmlOptions, string message) {
        string msg = System.Net.WebUtility.HtmlEncode(message ?? "Content too large.");
        string inner = $"<blockquote class=\"callout warning\" data-omd=\"overflow\"><p>{msg}</p></blockquote>";

        var bodyClass = htmlOptions.BodyClass;
        if (bodyClass != null) {
            bodyClass = bodyClass.Trim();
            if (bodyClass.Length > 0) {
                string cls = System.Net.WebUtility.HtmlEncode(bodyClass);
                return $"<article class=\"{cls}\">{inner}</article>";
            }
        }

        return $"<div data-omd=\"overflow\">{inner}</div>";
    }

    /// <summary>
    /// Builds a self-contained HTML document that preloads CSS and scripts once (Prism/Mermaid),
    /// and exposes a global <c>updateContent(newBodyHtml)</c> function for incremental updates.
    /// </summary>
    public static string BuildShellHtml(string? title = null, MarkdownRendererOptions? options = null) {
        options ??= new MarkdownRendererOptions();
        var htmlOptions = options.HtmlOptions ?? new HtmlOptions { Kind = HtmlKind.Fragment };

        // Build head assets (CSS + optional Prism assets) from OfficeIMO.Markdown.
        // This intentionally uses an empty doc; content is pushed later via updateContent(...).
        var empty = MarkdownDoc.Create();
        var parts = empty.ToHtmlParts(htmlOptions);

        var sb = new StringBuilder(16 * 1024);
        sb.Append("<!DOCTYPE html><html lang=\"en\"><head><meta charset=\"utf-8\"><meta name=\"viewport\" content=\"width=device-width,initial-scale=1\">");
        if (!string.IsNullOrWhiteSpace(options.ContentSecurityPolicy)) {
            sb.Append("<meta http-equiv=\"Content-Security-Policy\" content=\"")
              .Append(System.Net.WebUtility.HtmlEncode(options.ContentSecurityPolicy!.Trim()))
              .Append("\">");
        }
        sb.Append("<title>").Append(System.Net.WebUtility.HtmlEncode(title ?? "Markdown")).Append("</title>");
        if (!string.IsNullOrEmpty(parts.Css)) sb.Append("<style>\n").Append(parts.Css).Append("\n</style>");
        if (!string.IsNullOrEmpty(parts.Head)) sb.Append(parts.Head);
        if (!string.IsNullOrWhiteSpace(options.ShellCss)) {
            sb.Append("<style data-omd=\"shell\">")
              .Append("\n")
              .Append(options.ShellCss)
              .Append("\n</style>");
        }

        var assetMode = htmlOptions.AssetMode;
        var externalTextResolver = htmlOptions.ExternalTextResolver;

        if (options.Math?.Enabled == true) {
            sb.Append(BuildMathBootstrap(options.Math, assetMode, externalTextResolver));
        }

        if (options.Mermaid?.Enabled == true) {
            sb.Append(BuildMermaidBootstrap(options.Mermaid, assetMode, externalTextResolver));
        }

        if (options.Chart?.Enabled == true) {
            sb.Append(BuildChartBootstrap(options.Chart, assetMode, externalTextResolver));
        }

        AppendCustomShellHeadHtml(sb, options, assetMode);

        sb.Append("</head><body>");
        sb.Append("<div id=\"omdRoot\"></div>");
        sb.Append("<script>\n").Append(BuildIncrementalUpdateScript(options)).Append("\n</script>");
        sb.Append("</body></html>");
        return sb.ToString();
    }

    /// <summary>
    /// Returns a JavaScript snippet that calls <c>updateContent(...)</c> with a properly escaped string literal.
    /// </summary>
    public static string BuildUpdateScript(string bodyHtml) {
        return "updateContent(" + JavaScriptString.SingleQuoted(bodyHtml ?? string.Empty) + ");";
    }

    /// <summary>
    /// Convenience helper for hosts: renders Markdown to an HTML fragment and returns the JavaScript snippet
    /// that updates the shell (calls <c>updateContent(...)</c>).
    /// </summary>
    public static string RenderUpdateScript(string markdown, MarkdownRendererOptions? options = null) {
        var bodyHtml = RenderBodyHtml(markdown ?? string.Empty, options);
        return BuildUpdateScript(bodyHtml);
    }

    /// <summary>
    /// Wraps an existing HTML fragment in a chat "bubble" container (optional).
    /// This is purely a formatting helper: it does not change Markdown parsing rules.
    /// </summary>
    public static string WrapAsChatBubble(string bodyHtml, ChatMessageRole role = ChatMessageRole.Assistant) {
        string roleClass = role switch {
            ChatMessageRole.User => "omd-role-user",
            ChatMessageRole.System => "omd-role-system",
            _ => "omd-role-assistant"
        };

        // bodyHtml is expected to be the output of RenderBodyHtml (typically an <article class="markdown-body"> wrapper).
        // Keep it as-is and add a single outer container so host UIs don't have to author HTML around each message.
        return $"<div class=\"omd-chat-row {roleClass}\"><div class=\"omd-chat-bubble\">{bodyHtml ?? string.Empty}</div></div>";
    }

    /// <summary>
    /// Convenience helper: renders Markdown then wraps the result in a chat bubble.
    /// </summary>
    public static string RenderChatBubbleBodyHtml(string markdown, ChatMessageRole role = ChatMessageRole.Assistant, MarkdownRendererOptions? options = null) {
        var bodyHtml = RenderBodyHtml(markdown ?? string.Empty, options);
        return WrapAsChatBubble(bodyHtml, role);
    }

    /// <summary>
    /// Convenience helper: renders Markdown as a chat bubble and returns an update script snippet.
    /// </summary>
    public static string RenderChatBubbleUpdateScript(string markdown, ChatMessageRole role = ChatMessageRole.Assistant, MarkdownRendererOptions? options = null) {
        return BuildUpdateScript(RenderChatBubbleBodyHtml(markdown, role, options));
    }
}
