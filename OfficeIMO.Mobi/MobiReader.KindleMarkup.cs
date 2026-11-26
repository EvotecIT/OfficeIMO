using System.Text;
using AngleSharp.Html;
using AngleSharp.Html.Parser;
using AngleSharp.Html.Parser.Tokens;
using AngleSharp.Text;
using OfficeIMO.Html;

namespace OfficeIMO.Mobi;

public static partial class MobiReader {
    private static string ExpandKindleEmptyElements(string markup, ref int remainingTokens,
        out List<(int Offset, int Length)>? insertedClosers, CancellationToken token) {
        insertedClosers = null;
        if (markup.IndexOf("/>", StringComparison.Ordinal) < 0) return markup;

        var source = new TextSource(markup);
        using var tokenizer = new HtmlTokenizer(source, HtmlEntityProvider.ResolverExtended);
        var transitions = new Stack<(string EndTag, string? PreviousNamespace)>();
        string? foreign = null;
        StringBuilder? expanded = null;
        int copied = 0;
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

            if (!tag.IsSelfClosing) {
                tokenizer.State = tag.Name switch {
                    "style" or "xmp" or "iframe" or "noembed" or "noframes" => HtmlParseMode.Rawtext,
                    "script" => HtmlParseMode.Script,
                    "textarea" or "title" => HtmlParseMode.RCData,
                    "plaintext" => HtmlParseMode.Plaintext,
                    _ => HtmlParseMode.PCData
                };
                continue;
            }
            if (HtmlDomElementFacts.IsVoidElement(tag.Name) || tag.Name is "basefont" or "bgsound" or "frame" or "keygen" or "mbp:pagebreak")
                continue;

            expanded ??= new StringBuilder(markup.Length);
            string name = tag.Name;
            if (name == "plaintext") {
                // HTML's plaintext state cannot be closed. Keep the attributes on an empty span.
                int nameStart = tag.Position.Index + 1;
                expanded.Append(markup, copied, nameStart - copied).Append("span").Append(' ', name.Length - "span".Length);
                copied = nameStart + name.Length;
                name = "span";
            }

            // Keep the original opening-tag length and record each inserted close so
            // Kindle positions can be translated from the original byte offsets.
            int end = source.Index;
            expanded.Append(markup, copied, end - 2 - copied).Append(" ></").Append(name).Append('>');
            copied = end;
            (insertedClosers ??= new List<(int, int)>()).Add((end, name.Length + 3));
        }
        return expanded == null ? markup : expanded.Append(markup, copied, markup.Length - copied).ToString();
    }

    private static bool IsHtmlIntegrationPoint(string? foreign, HtmlTagToken tag) =>
        foreign == "svg" && (tag.Name is "foreignobject" or "desc" or "title") ||
        foreign == "math" && (tag.Name is "mi" or "mo" or "mn" or "ms" or "mtext" ||
            tag.Name == "annotation-xml" &&
            (tag.GetAttribute("encoding")?.ToLowerInvariant() is "text/html" or "application/xhtml+xml"));
}
