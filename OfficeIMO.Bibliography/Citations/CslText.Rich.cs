using System.Net;
using System.Xml;
using System.Xml.Linq;

namespace OfficeIMO.Bibliography;

internal readonly partial struct CslText {
    internal static CslText Rich(string source, bool variable = true, CancellationToken cancellationToken = default) {
        // Decode entities as text, preserving the distinction between actual
        // input tags and an escaped literal such as &lt;i&gt;.
        cancellationToken.ThrowIfCancellationRequested();
        if (source.Length > 4 * 1024 * 1024) throw new InvalidDataException("CSL rich text exceeds the maximum field length of four million characters.");
        string encoded = RichXmlCarrier(source, cancellationToken);
        try {
            XElement root = CslStyle.ReadXml("<root>" + encoded + "</root>", checked(source.Length * 6 + 13), 64, cancellationToken);
            string html = string.Concat(root.Nodes().Select(node => RichNode(node, cancellationToken)));
            string plain = HtmlPlain(html, cancellationToken);
            return new CslText(plain, WrapInputQuotes(html, plain), variable ? 1 : 0, variable && plain.Length > 0 ? 1 : 0);
        } catch (XmlException) {
            string plain = WebUtility.HtmlDecode(source);
            return new CslText(plain, WrapInputQuotes(Escape(plain), plain), variable ? 1 : 0, variable && plain.Length > 0 ? 1 : 0);
        }
    }

    /// <summary>Converts HTML entities into XML-safe text without activating escaped markup.</summary>
    private static string RichXmlCarrier(string source, CancellationToken token) {
        if (source.IndexOf('&') < 0) return source;
        var output = new StringBuilder(source.Length);
        for (int index = 0; index < source.Length; index++) {
            if ((index & 1023) == 0) token.ThrowIfCancellationRequested();
            if (source[index] != '&') { output.Append(source[index]); continue; }
            // The supported HTML entity names are short. Bound each scan so
            // malformed ampersands cannot repeatedly traverse the whole field.
            int end = index + 1;
            int limit = Math.Min(source.Length, index + 32);
            while (end < limit && (char.IsLetterOrDigit(source[end]) || source[end] == '#')) end++;
            if (end < limit && source[end] == ';') {
                string entity = source.Substring(index, end - index + 1);
                string decoded = WebUtility.HtmlDecode(entity);
                if (decoded != entity) {
                    output.Append(Escape(decoded));
                    index = end;
                    continue;
                }
            }
            output.Append("&amp;");
        }
        return output.ToString();
    }

    private static string WrapInputQuotes(string html, string plain) => plain.IndexOfAny(new[] { '\'', '"', '‘', '’', '“', '”' }) >= 0 ?
        "<span data-csl-input=\"true\">" + html + "</span>" : html;

    private static string RichNode(XNode node, CancellationToken token) {
        token.ThrowIfCancellationRequested();
        if (node is XText text) return Escape(text.Value);
        if (!(node is XElement element)) return string.Empty;
        if (element.Name.Namespace != XNamespace.None) return Escape(element.ToString(SaveOptions.DisableFormatting));
        string body = string.Concat(element.Nodes().Select(child => RichNode(child, token)));
        switch (element.Name.LocalName) {
            case "i": case "b": return "<" + element.Name.LocalName + " data-csl-flip=\"true\">" + body + "</" + element.Name.LocalName + ">";
            case "sup": case "sub": return "<" + element.Name.LocalName + ">" + body + "</" + element.Name.LocalName + ">";
            case "sc": return "<span style=\"font-variant:small-caps\" data-csl-flip=\"true\">" + body + "</span>";
            case "span":
                return RichSpan(element, body);
            default: return Escape(element.ToString(SaveOptions.DisableFormatting));
        }
    }

    /// <summary>Combines supported input attributes instead of letting a protection class hide formatting.</summary>
    private static string RichSpan(XElement element, string body) {
        string[] classes = ((string?)element.Attribute("class") ?? string.Empty).Split((char[]?)null, StringSplitOptions.RemoveEmptyEntries);
        bool nocase = classes.Contains("nocase"), nodecor = classes.Contains("nodecor");
        bool smallCaps = HasSmallCaps((string?)element.Attribute("style"));
        if (!nocase && !nodecor && !smallCaps) return body;
        string attributes = nocase || nodecor ? " class=\"" + (nocase && nodecor ? "nocase nodecor" : nocase ? "nocase" : "nodecor") + "\"" : string.Empty;
        if (smallCaps) attributes += " style=\"font-variant:small-caps\" data-csl-flip=\"true\"";
        return "<span" + attributes + ">" + body + "</span>";
    }

    /// <summary>Reads the supported CSS property without forwarding input declarations into output.</summary>
    private static bool HasSmallCaps(string? css) {
        if (css == null) return false;
        string? variant = null;
        foreach (string declaration in css.Split(';')) {
            int colon = declaration.IndexOf(':');
            if (colon > 0 && declaration.Substring(0, colon).Trim().Equals("font-variant", StringComparison.OrdinalIgnoreCase))
                variant = declaration.Substring(colon + 1).Trim();
        }
        return string.Equals(variant, "small-caps", StringComparison.OrdinalIgnoreCase);
    }
}
