using System.Xml.Linq;

namespace OfficeIMO.Workflows;

public sealed partial class BookProject {
    private static XElement[] BuildOnixMarketSegments(BookOnixSupply supply, OnixTerritory market,
        ref int budget, CancellationToken token) {
        ArgumentNullException.ThrowIfNull(supply.MarketSegments);
        ArgumentNullException.ThrowIfNull(supply.Restrictions);
        if (supply.MarketSegments.Count > 32)
            throw new ArgumentException("At most 32 market segments are supported per supply.", nameof(supply));
        XNamespace ns = OnixNamespace;
        if (supply.MarketSegments.Count == 0)
            return [new XElement(ns + "Market", market.ToXml(), BuildOnixSalesRestrictions(supply.Restrictions, ref budget, token))];
        var territories = new List<OnixTerritory>();
        var elements = new List<XElement>();
        foreach (var segment in supply.MarketSegments) {
            token.ThrowIfCancellationRequested(); ArgumentNullException.ThrowIfNull(segment);
            ArgumentNullException.ThrowIfNull(segment.Restrictions);
            var territory = ReadOnixTerritory(segment.Territory);
            if (!market.Covers(territory) || territories.Any(previous => previous.Overlaps(territory)))
                throw new ArgumentException("Market segments must be disjoint and contained in the supply territory.", nameof(supply));
            if (supply.Restrictions.Count > 32 || segment.Restrictions.Count > 32 - supply.Restrictions.Count)
                throw new ArgumentException("Combined supply and segment restrictions cannot exceed 32.", nameof(supply));
            territories.Add(territory);
            elements.Add(new XElement(ns + "Market", territory.ToXml(), BuildOnixSalesRestrictions(
                supply.Restrictions.Concat(segment.Restrictions).ToArray(), ref budget, token)));
        }
        if (!OnixTerritoriesCover(market, territories))
            throw new ArgumentException("Market segments must cover the entire supply territory.", nameof(supply));
        return elements.ToArray();
    }
}
