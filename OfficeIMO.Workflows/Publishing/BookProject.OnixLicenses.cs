using System.Globalization;
using System.Xml.Linq;

namespace OfficeIMO.Workflows;

public sealed partial class BookProject {
    private static IReadOnlyList<XElement> BuildOnixLicenses(IReadOnlyList<BookOnixLicense> licenses,
        ref int textBudget, CancellationToken token) {
        ArgumentNullException.ThrowIfNull(licenses);
        if (licenses.Count > 16) throw new ArgumentException("At most 16 licenses per collateral item are supported.", nameof(licenses));
        var result = new List<XElement>();
        XNamespace ns = OnixNamespace;
        foreach (var license in licenses) {
            token.ThrowIfCancellationRequested();
            ArgumentNullException.ThrowIfNull(license);
            ArgumentNullException.ThrowIfNull(license.Names);
            ArgumentNullException.ThrowIfNull(license.Expressions);
            if (license.Names.Count is < 1 or > 16 || license.Expressions.Count > 16)
                throw new ArgumentException("Supply one to 16 license names and at most 16 expressions.", nameof(licenses));
            if (license.ValidFrom > license.ValidUntil)
                throw new ArgumentException("License validity dates are reversed.", nameof(licenses));
            var element = new XElement(ns + "EpubLicense");
            var languages = new HashSet<string>(StringComparer.Ordinal);
            foreach (var name in license.Names) {
                token.ThrowIfCancellationRequested();
                ArgumentNullException.ThrowIfNull(name);
                RequireOnixText(name.Text, nameof(name.Text));
                if (name.Text.Length > 100 || name.Text.Length > textBudget)
                    throw new ArgumentException("License names exceed the field or aggregate collateral text limit.", nameof(licenses));
                textBudget -= name.Text.Length;
                RequireOnixTranslationLanguage(name.LanguageCode, license.Names.Count, nameof(license.Names));
                if (!languages.Add(name.LanguageCode ?? string.Empty))
                    throw new ArgumentException("License name languages must be distinct.", nameof(licenses));
                element.Add(new XElement(ns + "EpubLicenseName", name.LanguageCode != null ? new XAttribute("language", name.LanguageCode) : null, name.Text));
            }
            var expressions = new HashSet<(BookOnixLicenseExpressionType, string)>();
            foreach (var expression in license.Expressions) {
                token.ThrowIfCancellationRequested();
                ArgumentNullException.ThrowIfNull(expression);
                string type = expression.Type switch {
                    BookOnixLicenseExpressionType.HumanReadable => "01", BookOnixLicenseExpressionType.ProfessionalReadable => "02",
                    BookOnixLicenseExpressionType.AdditionalHumanReadable => "03", BookOnixLicenseExpressionType.AdditionalProfessionalReadable => "04",
                    BookOnixLicenseExpressionType.OnixPl => "10", BookOnixLicenseExpressionType.Odrl => "20",
                    BookOnixLicenseExpressionType.AdditionalOdrl => "21", _ => throw new ArgumentOutOfRangeException(nameof(expression.Type))
                };
                RequireOnixHttpUrl(expression.Link, nameof(expression.Link));
                if (!expressions.Add((expression.Type, expression.Link)))
                    throw new ArgumentException("License expressions must have distinct format/link pairs.", nameof(licenses));
                if (expression.Link.Length > textBudget)
                    throw new ArgumentException("License links exceed the aggregate collateral text limit.", nameof(licenses));
                textBudget -= expression.Link.Length;
                element.Add(new XElement(ns + "EpubLicenseExpression", new XElement(ns + "EpubLicenseExpressionType", type),
                    new XElement(ns + "EpubLicenseExpressionLink", expression.Link)));
            }
            void AddDate(string role, DateOnly? date) {
                if (date == null) return;
                element.Add(new XElement(ns + "EpubLicenseDate", new XElement(ns + "EpubLicenseDateRole", role),
                    new XElement(ns + "Date", new XAttribute("dateformat", "00"), date.Value.ToString("yyyyMMdd", CultureInfo.InvariantCulture))));
            }
            AddDate("14", license.ValidFrom); AddDate("15", license.ValidUntil);
            result.Add(element);
        }
        return result;
    }
}
