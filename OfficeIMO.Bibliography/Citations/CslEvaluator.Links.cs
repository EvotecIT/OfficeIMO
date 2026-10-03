using System.Xml.Linq;

namespace OfficeIMO.Bibliography;

internal sealed partial class CslEvaluator {
    private CslText LinkBibliographyText(CslText value, XElement element, CslContext context, ref XElement formatting) {
        if (!_options.LinkBibliographyIdentifiers || _options.OutputFormat != CslOutputFormat.Html ||
            context.Scope != XElementScope.Bibliography || context.Sorting || value.IsEmpty || element.Name.LocalName != "text") return value;
        string? variable = Attr(element, "variable");
        string? uriPrefix = variable switch {
            "URL" => string.Empty,
            "DOI" => "https://doi.org/",
            "PMID" => "https://www.ncbi.nlm.nih.gov/pubmed/",
            "PMCID" => "https://www.ncbi.nlm.nih.gov/pmc/articles/",
            _ => null
        };
        if (uriPrefix == null) return value;
        string identifier = System.Net.WebUtility.HtmlDecode(context.Record.Scalar(variable!));
        string target = uriPrefix + identifier;
        for (int index = 0; index < target.Length; index++) {
            if ((index & 1023) == 0) _token.ThrowIfCancellationRequested();
            if (char.IsControl(target[index])) return value;
        }
        if (!Uri.TryCreate(target, UriKind.Absolute, out Uri? uri) ||
            uri.Scheme != Uri.UriSchemeHttp && uri.Scheme != Uri.UriSchemeHttps || string.IsNullOrEmpty(uri.Host)) return value;

        // A URI prefix belongs inside the anchor. Descriptive affixes stay in
        // ordinary text and are subsequently applied by Decorate.
        string prefix = Attr(formatting, "prefix") ?? string.Empty;
        if (uriPrefix.Length > 0 && string.Equals(prefix, uriPrefix, StringComparison.OrdinalIgnoreCase)) {
            value = value.Affix(prefix, string.Empty);
            formatting = new XElement(formatting);
            formatting.Attribute("prefix")?.Remove();
        }
        return new CslText(value.Plain, "<a href=\"" + CslText.Escape(target) + "\">" + value.Html + "</a>", value.Attempted, value.Rendered);
    }
}
