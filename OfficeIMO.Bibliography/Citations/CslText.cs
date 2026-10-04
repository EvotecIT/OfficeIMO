using System.Net;
using System.Xml.Linq;

namespace OfficeIMO.Bibliography;

internal readonly partial struct CslText {
    internal CslText(string plain, string html, int attempted = 0, int rendered = 0) {
        Plain = plain; Html = html; Attempted = attempted; Rendered = rendered;
    }
    internal string Plain { get; }
    internal string Html { get; }
    internal int Attempted { get; }
    internal int Rendered { get; }
    internal bool IsEmpty => string.IsNullOrEmpty(Plain);
    internal static CslText Empty => Literal(string.Empty);
    internal static CslText Literal(string text, bool variable = false) => new CslText(text, Escape(text), variable ? 1 : 0, variable && text.Length > 0 ? 1 : 0);
    internal static string Escape(string text) => WebUtility.HtmlEncode(text);

    internal static CslText Join(IEnumerable<CslText> source, string delimiter, bool suppressEmptyVariables = false, int maximumCharacters = int.MaxValue) {
        int attempts = 0, rendered = 0;
        var plain = new StringBuilder();
        var html = new StringBuilder();
        foreach (CslText value in source) {
            attempts = checked(attempts + value.Attempted); rendered = checked(rendered + value.Rendered);
            if (value.IsEmpty) continue;
            string separator = delimiter;
            if (plain.Length == 0) separator = string.Empty;
            string encoded = Escape(separator);
            MergePunctuationBoundary(plain, html, ref separator, ref encoded);
            if ((long)plain.Length + separator.Length > maximumCharacters || (long)html.Length + encoded.Length > maximumCharacters)
                throw new InvalidDataException("CSL rendering exceeds MaximumIntermediateCharacters.");
            plain.Append(separator); html.Append(encoded);
            string body = value.Plain, markup = value.Html;
            MergePunctuationBoundary(plain, html, ref body, ref markup);
            if ((long)plain.Length + body.Length > maximumCharacters || (long)html.Length + markup.Length > maximumCharacters)
                throw new InvalidDataException("CSL rendering exceeds MaximumIntermediateCharacters.");
            plain.Append(body); html.Append(markup);
        }
        if (suppressEmptyVariables && attempts > 0 && rendered == 0) return new CslText(string.Empty, string.Empty, attempts, 0);
        return new CslText(plain.ToString(), html.ToString(), attempts, rendered);
    }

    internal CslText Affix(string prefix, string suffix) {
        if (IsEmpty) return this;
        if (prefix.Length == 0 && suffix.Length == 0) return this;
        return Join(new[] { Literal(prefix), this, Literal(suffix) }, string.Empty);
    }

    internal CslText TrimStart(CancellationToken token) {
        bool trimming = true;
        string html = TransformHtml(Html, (value, _) => {
            if (!trimming) return value;
            string trimmed = value.TrimStart();
            if (trimmed.Length > 0) trimming = false;
            return trimmed;
        }, token);
        return new CslText(Plain.TrimStart(), html, Attempted, Rendered);
    }

    internal CslText Decorate(XElement element, CslLocale locale, bool sorting = false, bool includeAffixes = true, CultureInfo? textCulture = null, CancellationToken cancellationToken = default) {
        if (IsEmpty) return this;
        bool layout = element.Name == CslStyle.Namespace + "layout" && !sorting;
        CslText initial = layout ? AffixDisplay((string?)element.Attribute("prefix") ?? string.Empty, (string?)element.Attribute("suffix") ?? string.Empty) : this;
        string plain = initial.Plain, html = initial.Html;
        if ((string?)element.Attribute("strip-periods") == "true") { plain = plain.Replace(".", string.Empty); html = TransformHtml(html, (value, _) => value.Replace(".", string.Empty), cancellationToken); }
        string? textCase = (string?)element.Attribute("text-case");
        if (textCase != null) {
            var unprotected = new StringBuilder();
            TransformHtml(html, (value, protectedCase) => {
                if (!protectedCase) unprotected.Append(value);
                return value;
            }, cancellationToken);
            CultureInfo culture = textCulture ?? locale.Culture;
            string unprotectedText = unprotected.ToString();
            bool allUpper = unprotectedText == unprotectedText.ToUpper(culture);
            string cased = ChangeCase(HtmlPlain(html, cancellationToken), textCase, culture, allUpper, cancellationToken);
            int offset = 0;
            html = TransformHtml(html, (value, protectedCase) => {
                int start = offset; offset += value.Length;
                return protectedCase || cased.Length != plain.Length ? value : cased.Substring(start, value.Length);
            }, cancellationToken);
            plain = HtmlPlain(html, cancellationToken);
        }
        if (!sorting) {
            html = Wrap(html, "font-style", (string?)element.Attribute("font-style"));
            html = Wrap(html, "font-weight", (string?)element.Attribute("font-weight"));
            html = Wrap(html, "font-variant", (string?)element.Attribute("font-variant"));
            html = Wrap(html, "text-decoration", (string?)element.Attribute("text-decoration"));
            html = Wrap(html, "vertical-align", (string?)element.Attribute("vertical-align"));
            if ((string?)element.Attribute("quotes") == "true") {
                string open = locale.Term("open-quote"), close = locale.Term("close-quote");
                plain = open + plain + close;
                html = "<span data-csl-quote=\"true\"" + (html.Contains("<div") ? " data-csl-display-quotes=\"true\" data-csl-display-owner=\"true\"" : string.Empty) + ">" + Escape(open) + html + Escape(close) + "</span>";
            }
        }
        var result = new CslText(plain, html, Attempted, Rendered);
        if (!sorting && !layout && includeAffixes) result = result.AffixDisplay((string?)element.Attribute("prefix") ?? string.Empty, (string?)element.Attribute("suffix") ?? string.Empty);
        return !sorting && (string?)element.Attribute("display") is string display ? result.WithDisplay(display) : result;
    }

    internal CslText WithDisplay(string display) => IsEmpty ? this :
        new CslText(Plain, "<div class=\"csl-" + Escape(display) + "\">" + Html + "</div>", Attempted, Rendered);

    private static string Wrap(string html, string property, string? value) {
        if (value == null) return html;
        if (property == "font-style" && value == "italic") return "<i>" + html + "</i>";
        if (property == "font-weight" && value == "bold") return "<b>" + html + "</b>";
        if (property == "vertical-align" && value == "sup") return "<sup>" + html + "</sup>";
        if (property == "vertical-align" && value == "sub") return "<sub>" + html + "</sub>";
        string css = property == "vertical-align" && value == "baseline" ? "baseline" : value;
        return "<span style=\"" + property + ":" + Escape(css) + "\">" + html + "</span>";
    }
}
