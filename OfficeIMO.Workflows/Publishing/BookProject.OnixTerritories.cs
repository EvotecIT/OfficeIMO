using System.Xml.Linq;

namespace OfficeIMO.Workflows;

public sealed partial class BookProject {
    private sealed record OnixTerritory(bool Worldwide, HashSet<string> Countries, HashSet<string> Excluded) {
        internal bool Includes(string country) => Worldwide ? !Excluded.Contains(country) : Countries.Contains(country);
        internal bool Overlaps(OnixTerritory other) => Worldwide && other.Worldwide ||
            (Worldwide ? other.Countries.Any(Includes) : Countries.Any(other.Includes));
        internal bool Covers(OnixTerritory other) => other.Worldwide
            ? Worldwide && Excluded.IsSubsetOf(other.Excluded)
            : other.Countries.All(Includes);
        internal XElement ToXml() {
            XNamespace ns = OnixNamespace;
            var element = new XElement(ns + "Territory");
            if (Worldwide) {
                element.Add(new XElement(ns + "RegionsIncluded", "WORLD"));
                if (Excluded.Count != 0) element.Add(new XElement(ns + "CountriesExcluded", string.Join(" ", Excluded.Order(StringComparer.Ordinal))));
            } else element.Add(new XElement(ns + "CountriesIncluded", string.Join(" ", Countries.Order(StringComparer.Ordinal))));
            return element;
        }
    }

    private static OnixTerritory ReadOnixTerritory(BookOnixTerritory territory) {
        ArgumentNullException.ThrowIfNull(territory);
        HashSet<string> countries = ReadOnixCountries(territory.Countries), excluded = ReadOnixCountries(territory.ExcludedCountries);
        if (territory.Worldwide ? countries.Count != 0 : countries.Count == 0 || excluded.Count != 0)
            throw new ArgumentException("Use included countries, or Worldwide with optional country exclusions.", nameof(territory));
        return new OnixTerritory(territory.Worldwide, countries, excluded);
    }

    private static HashSet<string> ReadOnixCountries(IReadOnlyList<string> countries) {
        ArgumentNullException.ThrowIfNull(countries);
        if (countries.Count > 250) throw new ArgumentException("A territory country list cannot exceed 250 codes.", nameof(countries));
        var result = new HashSet<string>(StringComparer.Ordinal);
        foreach (string code in countries) {
            if (code == null || code.Length != 2 || code.Any(c => c < 'A' || c > 'Z') || !result.Add(code))
                throw new ArgumentException("Country codes must be unique uppercase two-letter ONIX list 91 values.", nameof(countries));
        }
        return result;
    }

    private static bool OnixGrantsCover(OnixTerritory market, IReadOnlyList<OnixTerritory> grants) {
        if (!market.Worldwide) return market.Countries.All(country => grants.Any(grant => grant.Includes(country)));
        OnixTerritory[] worldwide = grants.Where(grant => grant.Worldwide).ToArray();
        if (worldwide.Length == 0) return false; // A finite list is never silently broadened into WORLD.
        var uncovered = new HashSet<string>(worldwide[0].Excluded, StringComparer.Ordinal);
        foreach (OnixTerritory grant in worldwide.Skip(1)) uncovered.IntersectWith(grant.Excluded);
        foreach (OnixTerritory grant in grants.Where(grant => !grant.Worldwide)) uncovered.ExceptWith(grant.Countries);
        return uncovered.IsSubsetOf(market.Excluded);
    }
}
