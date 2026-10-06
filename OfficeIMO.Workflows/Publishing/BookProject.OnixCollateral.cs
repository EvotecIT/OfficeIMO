using System.Globalization;
using System.Text;
using System.Xml;
using System.Xml.Linq;

namespace OfficeIMO.Workflows;

public sealed partial class BookProject {
    private static XElement? BuildOnixCollateral(IReadOnlyList<BookOnixCollateralText> items, IReadOnlyList<BookOnixSupportingResource> resources, IReadOnlyList<XElement> contributors, CancellationToken cancellationToken) {
        ArgumentNullException.ThrowIfNull(items);
        if (items.Count > 64) throw new ArgumentException("At most 64 collateral text items are supported.", nameof(items));
        ArgumentNullException.ThrowIfNull(resources);
        if (resources.Count > 64) throw new ArgumentException("At most 64 supporting resources are supported.", nameof(resources));
        if (items.Count == 0 && resources.Count == 0) return null;
        XNamespace ns = OnixNamespace;
        var result = new XElement(ns + "CollateralDetail");
        int textBudget = 524288;
        foreach (var item in items) {
            cancellationToken.ThrowIfCancellationRequested();
            ArgumentNullException.ThrowIfNull(item);
            if (!Enum.IsDefined(item.Type)) throw new ArgumentOutOfRangeException(nameof(item.Type));
            var content = new XElement(ns + "TextContent", new XElement(ns + "SequenceNumber", result.Elements().Count() + 1),
                new XElement(ns + "TextType", ((int)item.Type).ToString("00", CultureInfo.InvariantCulture)));
            content.Add(BuildOnixContentAudiences(item.Audiences));
            if (item.Territory != null) content.Add(ReadOnixTerritory(item.Territory).ToXml());
            ArgumentNullException.ThrowIfNull(item.Texts);
            if (item.Texts.Count == 0) throw new ArgumentException("Collateral requires text.", nameof(item.Texts));
            bool shortText = item.Type is BookOnixTextType.ShortDescription or BookOnixTextType.CollectionShortDescription;
            content.Add(BuildOnixCollateralValues(item.Texts, "Text", 65536, shortText, ref textBudget, cancellationToken));
            content.Add(BuildOnixReviewRating(item.ReviewRating, item.Type, ref textBudget, cancellationToken));
            ArgumentNullException.ThrowIfNull(item.Authors);
            if (item.Authors.Count > 16 || item.Authors.Distinct(StringComparer.Ordinal).Count() != item.Authors.Count)
                throw new ArgumentException("Supply at most 16 distinct text authors.", nameof(item.Authors));
            foreach (var author in item.Authors) {
                RequireOnixText(author, nameof(item.Authors));
                content.Add(new XElement(ns + "TextAuthor", author));
            }
            if (item.SourceCorporate != null) {
                RequireOnixText(item.SourceCorporate, nameof(item.SourceCorporate));
                content.Add(new XElement(ns + "TextSourceCorporate", item.SourceCorporate));
            }
            content.Add(BuildOnixCollateralValues(item.SourceTitles, "SourceTitle", 4096, false, ref textBudget, cancellationToken));
            ArgumentNullException.ThrowIfNull(item.SourceLinks);
            if (item.SourceLinks.Count > 16 || item.SourceLinks.Distinct(StringComparer.Ordinal).Count() != item.SourceLinks.Count)
                throw new ArgumentException("Supply at most 16 distinct source links.", nameof(item.SourceLinks));
            foreach (var link in item.SourceLinks) {
                RequireOnixHttpUrl(link, nameof(item.SourceLinks));
                content.Add(new XElement(ns + "TextSourceLink", link));
            }
            content.Add(BuildOnixUsageConstraints(item.UsageConstraints, cancellationToken));
            content.Add(BuildOnixLicenses(item.Licenses, ref textBudget, cancellationToken));
            if (item.UsableFrom > item.UsableUntil) throw new ArgumentException("Collateral usage dates are reversed.", nameof(item));
            void AddDate(string role, DateOnly? date) {
                if (date == null) return;
                content.Add(new XElement(ns + "ContentDate", new XElement(ns + "ContentDateRole", role),
                    new XElement(ns + "Date", new XAttribute("dateformat", "00"), date.Value.ToString("yyyyMMdd", CultureInfo.InvariantCulture))));
            }
            AddDate("01", item.PublishedOn); AddDate("14", item.UsableFrom);
            AddDate("15", item.UsableUntil); AddDate("17", item.UpdatedOn);
            result.Add(content);
        }
        AddOnixSupportingResources(result, resources, contributors, ref textBudget, cancellationToken);
        return result;
    }

    private static IReadOnlyList<XElement> BuildOnixCollateralValues(IReadOnlyList<BookOnixCollateralTextValue> values,
        string elementName, int maximumLength, bool shortText, ref int textBudget, CancellationToken cancellationToken) {
        ArgumentNullException.ThrowIfNull(values);
        if (values.Count > 16) throw new ArgumentException("At most 16 language variants are supported.", nameof(values));
        XNamespace ns = OnixNamespace;
        var result = new List<XElement>();
        var languages = new HashSet<string>(StringComparer.Ordinal);
        foreach (var value in values) {
            cancellationToken.ThrowIfCancellationRequested();
            ArgumentNullException.ThrowIfNull(value);
            ArgumentException.ThrowIfNullOrWhiteSpace(value.Text);
            if (value.Text.Length > maximumLength || value.Text.Length > textBudget)
                throw new ArgumentException("Collateral text exceeds the per-field or aggregate text limit.", nameof(values));
            textBudget -= value.Text.Length;
            XmlConvert.VerifyXmlChars(value.Text);
            if (!Enum.IsDefined(value.Format)) throw new ArgumentOutOfRangeException(nameof(value.Format));
            if (elementName != "Text" && value.Format != BookOnixCollateralTextFormat.PlainText)
                throw new ArgumentException("This ONIX element accepts plain text only.", nameof(values));
            XElement? fragment = value.Format == BookOnixCollateralTextFormat.Xhtml
                ? ReadOnixXhtmlFragment(value.Text, cancellationToken) : null;
            if (shortText) {
                int characters = 0;
                foreach (var rune in (fragment?.Value ?? value.Text).EnumerateRunes()) {
                    if (++characters > 350) throw new ArgumentException("Short descriptions cannot exceed 350 Unicode scalar values.", nameof(values));
                }
            }
            RequireOnixTranslationLanguage(value.LanguageCode, values.Count, elementName);
            if (!languages.Add(value.LanguageCode ?? "")) throw new ArgumentException("Text variant languages must be distinct.", nameof(values));
            var element = new XElement(ns + elementName);
            if (fragment == null) element.Add(value.Text);
            else element.Add(fragment.Nodes());
            if (elementName == "Text") element.Add(new XAttribute("textformat", fragment == null ? "06" : "05"));
            if (value.LanguageCode != null) element.Add(new XAttribute("language", value.LanguageCode));
            result.Add(element);
        }
        return result;
    }
}
