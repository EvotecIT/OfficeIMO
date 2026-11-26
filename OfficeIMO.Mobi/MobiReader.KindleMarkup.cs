using AngleSharp.Html;
using AngleSharp.Html.Parser;
using AngleSharp.Html.Parser.Tokens;
using AngleSharp.Text;

namespace OfficeIMO.Mobi;

public static partial class MobiReader {
    private static string NormalizeKindleEmptyRawElements(string markup, ref int remainingTokens,
        out HashSet<int>? placeholders, CancellationToken token) {
        placeholders = null;
        if (markup.IndexOf("/>", StringComparison.Ordinal) < 0) return markup;

        var source = new TextSource(markup);
        var tokenizer = new HtmlTokenizer(source, HtmlEntityProvider.ResolverExtended);
        var transitions = new Stack<(string EndTag, string? PreviousNamespace)>();
        string? foreign = null;
        char[]? normalized = null;
        while (true) {
            token.ThrowIfCancellationRequested();
            if (--remainingTokens < 0)
                throw new InvalidDataException("Kindle XHTML exceeds its aggregate token limit.");
            HtmlToken current = tokenizer.Get();
            if (current.Type == HtmlTokenType.EndOfFile) break;
            if (current is not HtmlTagToken tag) continue;
            if (current.Type == HtmlTokenType.EndTag) {
                if (transitions.Count != 0 && transitions.Peek().EndTag == tag.Name) {
                    foreign = transitions.Pop().PreviousNamespace;
                    tokenizer.IsAcceptingCharacterData = transitions.Count != 0;
                }
                continue;
            }
            if (tag.Name is "svg" or "math" || IsHtmlIntegrationPoint(foreign, tag)) {
                if (!tag.IsSelfClosing) {
                    if (transitions.Count == 64)
                        throw new InvalidDataException("Kindle XHTML exceeds its namespace nesting limit.");
                    transitions.Push((tag.Name, foreign));
                    foreign = tag.Name is "svg" or "math" ? tag.Name : null;
                    tokenizer.IsAcceptingCharacterData = true;
                }
                continue;
            }
            if (foreign != null) continue;

            HtmlParseMode mode = tag.Name switch {
                "style" or "xmp" or "iframe" or "noembed" or "noframes" => HtmlParseMode.Rawtext,
                "script" => HtmlParseMode.Script,
                "textarea" or "title" => HtmlParseMode.RCData,
                "plaintext" => HtmlParseMode.Plaintext,
                _ => HtmlParseMode.PCData
            };
            if (mode == HtmlParseMode.PCData) continue;
            if (!tag.IsSelfClosing) {
                tokenizer.State = mode;
                continue;
            }

            // A void element retains the ID and source position without starting a raw-text
            // span. Use a head-safe placeholder for elements that can be parsed in the head.
            string placeholder = tag.Name is "style" or "title" or "script" or "noframes" ? "link" : "wbr";
            normalized ??= markup.ToCharArray();
            int start = tag.Position.Index;
            int nameStart = start + 1;
            placeholder.CopyTo(0, normalized, nameStart, placeholder.Length);
            for (int i = nameStart + placeholder.Length; i < nameStart + tag.Name.Length; i++)
                normalized[i] = ' ';
            (placeholders ??= new HashSet<int>()).Add(start);
        }
        return normalized == null ? markup : new string(normalized);
    }

    private static bool IsHtmlIntegrationPoint(string? foreign, HtmlTagToken tag) =>
        foreign == "svg" && (tag.Name is "foreignobject" or "desc" or "title") ||
        foreign == "math" && (tag.Name is "mi" or "mo" or "mn" or "ms" or "mtext" ||
            tag.Name == "annotation-xml" &&
            (tag.GetAttribute("encoding")?.ToLowerInvariant() is "text/html" or "application/xhtml+xml"));
}
