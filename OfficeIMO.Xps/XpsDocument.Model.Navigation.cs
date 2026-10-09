namespace OfficeIMO.Xps;

internal sealed partial class XpsDocumentModelProjection {
    private void AddLinks(XpsPage page, int pageIndex, XElement markup) {
        foreach (var element in XpsStoryFragmentsReader.PageElements(markup)) {
            _budget.Charge();
            if (element.Attribute("FixedPage.NavigateUri") is not XAttribute attribute) continue;
            (int PageIndex, string? Name, string? Uri)? target;
            try { target = _document.ResolveNavigation(page.PartName, attribute.Value); }
            catch (InvalidDataException) { target = null; }
            catch (UriFormatException) { target = null; }
            if (!target.HasValue) { Diagnostic("XpsNavigationUnavailable", "Unresolved or unsupported native link: " + attribute.Value, pageIndex); continue; }
            var value = target.Value;
            var text = new StringBuilder();
            foreach (var child in XpsStoryFragmentsReader.PageElements(element)) {
                _budget.Charge();
                if (child.Name.LocalName != "Glyphs") continue;
                string valueText = XpsPage.Unescape((string?)child.Attribute("UnicodeString") ?? string.Empty);
                _budget.Text(valueText.Length); text.Append(valueText);
            }
            _links.Add(new OfficeDocumentModelLink { Id = "xps-link-" + _links.Count.ToString(CultureInfo.InvariantCulture),
                Kind = value.Uri != null ? "hyperlink" : "internal-link", Uri = value.Uri, DestinationName = value.Name,
                DestinationPageNumber = value.PageIndex < 0 ? null : value.PageIndex + 1, Text = text.ToString(),
                Location = Location(pageIndex, null, "fixed-page-link") });
        }
    }
}
