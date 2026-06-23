namespace OfficeIMO.Pdf;

/// <summary>Commonly used page sizes.</summary>
public static class PageSizes {
    private static readonly IReadOnlyDictionary<string, PageSize> KnownSizes = new Dictionary<string, PageSize>(StringComparer.OrdinalIgnoreCase) {
        ["A0"] = new PageSize(2384, 3370),
        ["A1"] = new PageSize(1684, 2384),
        ["A2"] = new PageSize(1190, 1684),
        ["A3"] = new PageSize(842, 1190),
        ["A4"] = new PageSize(595, 842),
        ["A5"] = new PageSize(420, 595),
        ["A6"] = new PageSize(298, 420),
        ["A7"] = new PageSize(210, 298),
        ["A8"] = new PageSize(148, 210),
        ["A9"] = new PageSize(105, 148),
        ["A10"] = new PageSize(74, 105),
        ["B0"] = new PageSize(2834, 4008),
        ["B1"] = new PageSize(2004, 2834),
        ["B2"] = new PageSize(1417, 2004),
        ["B3"] = new PageSize(1000, 1417),
        ["B4"] = new PageSize(708, 1000),
        ["B5"] = new PageSize(498, 708),
        ["B6"] = new PageSize(354, 498),
        ["B7"] = new PageSize(249, 354),
        ["B8"] = new PageSize(175, 249),
        ["B9"] = new PageSize(124, 175),
        ["B10"] = new PageSize(88, 124),
        ["Executive"] = new PageSize(522, 756),
        ["Letter"] = new PageSize(612, 792),
        ["Legal"] = new PageSize(612, 1008),
        ["Ledger"] = new PageSize(1224, 792),
        ["Tabloid"] = new PageSize(792, 1224),
        ["LedgerOrTabloid"] = new PageSize(1224, 792)
    };

    /// <summary>ISO A0 size (2384 × 3370 pt).</summary>
    public static PageSize A0 => KnownSizes["A0"];
    /// <summary>ISO A1 size (1684 × 2384 pt).</summary>
    public static PageSize A1 => KnownSizes["A1"];
    /// <summary>ISO A2 size (1190 × 1684 pt).</summary>
    public static PageSize A2 => KnownSizes["A2"];
    /// <summary>ISO A3 size (842 × 1190 pt).</summary>
    public static PageSize A3 => KnownSizes["A3"];
    /// <summary>ISO A5 size (420 × 595 pt).</summary>
    public static PageSize A5 => KnownSizes["A5"];
    /// <summary>ISO A4 size (595 × 842 pt).</summary>
    public static PageSize A4 => KnownSizes["A4"];
    /// <summary>ISO A6 size (298 × 420 pt).</summary>
    public static PageSize A6 => KnownSizes["A6"];
    /// <summary>ISO A7 size (210 × 298 pt).</summary>
    public static PageSize A7 => KnownSizes["A7"];
    /// <summary>ISO A8 size (148 × 210 pt).</summary>
    public static PageSize A8 => KnownSizes["A8"];
    /// <summary>ISO A9 size (105 × 148 pt).</summary>
    public static PageSize A9 => KnownSizes["A9"];
    /// <summary>ISO A10 size (74 × 105 pt).</summary>
    public static PageSize A10 => KnownSizes["A10"];
    /// <summary>ISO B0 size (2834 × 4008 pt).</summary>
    public static PageSize B0 => KnownSizes["B0"];
    /// <summary>ISO B1 size (2004 × 2834 pt).</summary>
    public static PageSize B1 => KnownSizes["B1"];
    /// <summary>ISO B2 size (1417 × 2004 pt).</summary>
    public static PageSize B2 => KnownSizes["B2"];
    /// <summary>ISO B3 size (1000 × 1417 pt).</summary>
    public static PageSize B3 => KnownSizes["B3"];
    /// <summary>ISO B4 size (708 × 1000 pt).</summary>
    public static PageSize B4 => KnownSizes["B4"];
    /// <summary>ISO B5 size (498 × 708 pt).</summary>
    public static PageSize B5 => KnownSizes["B5"];
    /// <summary>ISO B6 size (354 × 498 pt).</summary>
    public static PageSize B6 => KnownSizes["B6"];
    /// <summary>ISO B7 size (249 × 354 pt).</summary>
    public static PageSize B7 => KnownSizes["B7"];
    /// <summary>ISO B8 size (175 × 249 pt).</summary>
    public static PageSize B8 => KnownSizes["B8"];
    /// <summary>ISO B9 size (124 × 175 pt).</summary>
    public static PageSize B9 => KnownSizes["B9"];
    /// <summary>ISO B10 size (88 × 124 pt).</summary>
    public static PageSize B10 => KnownSizes["B10"];
    /// <summary>US Executive size (522 × 756 pt).</summary>
    public static PageSize Executive => KnownSizes["Executive"];
    /// <summary>US Letter size (612 × 792 pt).</summary>
    public static PageSize Letter => KnownSizes["Letter"];
    /// <summary>US Legal size (612 × 1008 pt).</summary>
    public static PageSize Legal => KnownSizes["Legal"];
    /// <summary>US Ledger size (1224 × 792 pt).</summary>
    public static PageSize Ledger => KnownSizes["Ledger"];
    /// <summary>US Tabloid size (792 × 1224 pt).</summary>
    public static PageSize Tabloid => KnownSizes["Tabloid"];
    /// <summary>Alias for US Ledger or Tabloid landscape size (1224 × 792 pt).</summary>
    public static PageSize LedgerOrTabloid => KnownSizes["LedgerOrTabloid"];

    /// <summary>Known standard page-size names.</summary>
    public static IReadOnlyCollection<string> Names => KnownSizes.Keys.ToArray();

    /// <summary>Attempts to resolve a standard page-size name.</summary>
    public static bool TryGet(string? name, out PageSize pageSize) {
        pageSize = default;
        if (string.IsNullOrWhiteSpace(name)) {
            return false;
        }

        string trimmed = name!.Trim();
        return KnownSizes.TryGetValue(trimmed, out pageSize);
    }

    /// <summary>Resolves a standard page-size name.</summary>
    public static PageSize Get(string name) {
        if (TryGet(name, out PageSize pageSize)) {
            return pageSize;
        }

        throw new ArgumentException("Unknown page size name '" + name + "'.", nameof(name));
    }
}
