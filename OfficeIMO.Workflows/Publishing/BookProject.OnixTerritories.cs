using System.Xml.Linq;

namespace OfficeIMO.Workflows;

public sealed partial class BookProject {
    private sealed record OnixTerritory(bool Worldwide, HashSet<string> Countries, HashSet<string> Excluded,
        HashSet<string> Regions, HashSet<string> ExcludedRegions) {
        internal bool Includes((string? Country, string? Region) atom) {
            if (atom.Country == null) return Worldwide;
            bool country = Worldwide ? !Excluded.Contains(atom.Country) : Countries.Contains(atom.Country);
            return atom.Region == null ? country : (country || Regions.Contains(atom.Region)) && !ExcludedRegions.Contains(atom.Region);
        }
        internal bool Overlaps(OnixTerritory other) => TerritoryAtoms([this, other]).Any(atom => Includes(atom) && other.Includes(atom));
        internal bool Covers(OnixTerritory other) => TerritoryAtoms([this, other]).All(atom => !other.Includes(atom) || Includes(atom));
        internal XElement ToXml() {
            XNamespace ns = OnixNamespace;
            var element = new XElement(ns + "Territory");
            void Add(string name, IEnumerable<string> codes) {
                string value = string.Join(" ", codes.Order(StringComparer.Ordinal));
                if (value.Length != 0) element.Add(new XElement(ns + name, value));
            }
            Add("CountriesIncluded", Countries);
            Add("RegionsIncluded", Worldwide ? ["WORLD"] : Regions);
            Add("CountriesExcluded", Excluded);
            Add("RegionsExcluded", ExcludedRegions);
            return element;
        }
    }

    private static OnixTerritory ReadOnixTerritory(BookOnixTerritory territory) {
        ArgumentNullException.ThrowIfNull(territory);
        HashSet<string> countries = ReadOnixCountries(territory.Countries), excluded = ReadOnixCountries(territory.ExcludedCountries);
        HashSet<string> regions = ReadOnixSubregions(territory.Regions), excludedRegions = ReadOnixSubregions(territory.ExcludedRegions);
        if (territory.Worldwide ? countries.Count + regions.Count != 0 : countries.Count + regions.Count == 0 || excluded.Count != 0)
            throw new ArgumentException("Use included countries/subregions, or Worldwide with optional exclusions.", nameof(territory));
        if (regions.Any(region => countries.Contains(region[..2])))
            throw new ArgumentException("An included subregion cannot repeat its included parent country.", nameof(territory));
        if (excludedRegions.Any(region => territory.Worldwide ? excluded.Contains(region[..2]) : !countries.Contains(region[..2])))
            throw new ArgumentException("An excluded subregion must belong to an included country and cannot repeat an excluded country.", nameof(territory));
        return new(territory.Worldwide, countries, excluded, regions, excludedRegions);
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

    // Partition only at declared boundaries. A residual atom deliberately prevents
    // a finite set of subdivisions from silently becoming a whole-country grant.
    private static IEnumerable<(string? Country, string? Region)> TerritoryAtoms(IReadOnlyList<OnixTerritory> territories) {
        yield return (null, null); // The rest of WORLD outside explicitly mentioned countries.
        string[] regions = territories.SelectMany(t => t.Regions.Concat(t.ExcludedRegions)).Distinct(StringComparer.Ordinal).ToArray();
        IEnumerable<string> countries = territories.SelectMany(t => t.Countries.Concat(t.Excluded))
            .Concat(regions.Select(region => region[..2])).Distinct(StringComparer.Ordinal);
        foreach (string country in countries) {
            yield return (country, null);
            foreach (string region in regions.Where(region => region.StartsWith(country + "-", StringComparison.Ordinal))) yield return (country, region);
        }
    }

    private static bool OnixGrantsCover(OnixTerritory market, IReadOnlyList<OnixTerritory> grants) =>
        TerritoryAtoms(grants.Append(market).ToArray()).All(atom => !market.Includes(atom) || grants.Any(grant => grant.Includes(atom)));
}
