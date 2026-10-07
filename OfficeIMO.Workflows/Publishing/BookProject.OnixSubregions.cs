namespace OfficeIMO.Workflows;

public sealed partial class BookProject {
    // ONIX list 49 geographic subdivisions from the retained issue 72 catalog. Airport markets, aliases and overlapping region
    // hierarchies require their own coverage model and are not accepted here.
    private static readonly HashSet<string> OnixSubregions = new(string.Join(" ", new[] {
        "AU-CT AU-NS AU-NT AU-QL AU-SA AU-TS AU-VI AU-WA",
        "BE-BRU BE-VLG BE-WAL",
        "CA-AB CA-BC CA-MB CA-NB CA-NL CA-NS CA-NT CA-NU CA-ON CA-PE CA-QC CA-SK CA-YT",
        "CN-BJ CN-TJ CN-HE CN-SX CN-NM CN-LN CN-JL CN-HL CN-SH CN-JS CN-ZJ CN-AH CN-FJ CN-JX CN-SD CN-HA CN-HB CN-HN CN-GD CN-GX CN-HI CN-CQ CN-SC CN-GZ CN-YN CN-XZ CN-SN CN-GS CN-QH CN-NX CN-XJ",
        "ES-CN",
        "FR-H",
        "GB-ENG GB-NIR GB-SCT GB-WLS",
        "US-AK US-AL US-AR US-AZ US-CA US-CO US-CT US-DC US-DE US-FL US-GA US-HI US-IA US-ID US-IL US-IN US-KS US-KY US-LA US-MA US-MD US-ME US-MI US-MN US-MO US-MS US-MT US-NC US-ND US-NE US-NH US-NJ US-NM US-NV US-NY US-OH US-OK US-OR US-PA US-RI US-SC US-SD US-TN US-TX US-UT US-VA US-VT US-WA US-WI US-WV US-WY"
    }).Split(' ', StringSplitOptions.RemoveEmptyEntries), StringComparer.Ordinal);

    private static HashSet<string> ReadOnixSubregions(IReadOnlyList<string> regions) {
        ArgumentNullException.ThrowIfNull(regions);
        if (regions.Count > 250) throw new ArgumentException("A territory subregion list cannot exceed 250 codes.", nameof(regions));
        var result = new HashSet<string>(StringComparer.Ordinal);
        foreach (string code in regions) {
            if (code == null || !OnixSubregions.Contains(code) || !result.Add(code))
                throw new ArgumentException("Subregions must be unique supported uppercase ONIX list 49 geographic codes.", nameof(regions));
        }
        return result;
    }
}
