using System.Xml.Linq;

namespace OfficeIMO.Workflows;

public sealed partial class BookProject {
    private static IReadOnlyList<XElement> BuildOnixContributors(IReadOnlyList<BookOnixContributor> contributors,
        bool noContributors, bool requireDeclaration, CancellationToken cancellationToken) {
        ArgumentNullException.ThrowIfNull(contributors);
        if (contributors.Count > 100 || (noContributors && contributors.Count != 0))
            throw new ArgumentException("Supply at most 100 credits, or explicitly select NoContributors.", nameof(contributors));
        if (requireDeclaration && !noContributors && contributors.Count == 0)
            throw new ArgumentException("Supply 1-100 credits, or explicitly select NoContributors.", nameof(contributors));
        XNamespace ns = OnixNamespace;
        if (noContributors) return [new XElement(ns + "NoContributor")];
        var result = new List<XElement>();
        foreach (var credit in contributors) {
            cancellationToken.ThrowIfCancellationRequested();
            ArgumentNullException.ThrowIfNull(credit);
            RequireOnixText(credit.Name, nameof(credit.Name));
            string role = credit.Role switch {
                BookOnixContributorRole.Author => "A01", BookOnixContributorRole.Editor => "B01",
                BookOnixContributorRole.Translator => "B06", BookOnixContributorRole.Illustrator => "A12",
                BookOnixContributorRole.Other => "Z99", BookOnixContributorRole.SeriesEditor => "B09",
                _ => throw new ArgumentOutOfRangeException(nameof(credit.Role))
            };
            result.Add(new XElement(ns + "Contributor", new XElement(ns + "SequenceNumber", result.Count + 1),
                new XElement(ns + "ContributorRole", role),
                new XElement(ns + (credit.IsOrganization ? "CorporateName" : "PersonName"), credit.Name)));
        }
        return result;
    }
}
