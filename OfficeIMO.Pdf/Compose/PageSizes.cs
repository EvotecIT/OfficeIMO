namespace OfficeIMO.Pdf;

/// <summary>Commonly used page sizes.</summary>
public static class PdfPageSizes {
    private static readonly IReadOnlyDictionary<string, PdfPageSize> KnownSizes = new Dictionary<string, PdfPageSize>(StringComparer.OrdinalIgnoreCase) {
        ["A0"] = new PdfPageSize(OfficeIMO.OfficeLength.Points(2384),OfficeIMO.OfficeLength.Points(3370)),
        ["A1"] = new PdfPageSize(OfficeIMO.OfficeLength.Points(1684),OfficeIMO.OfficeLength.Points(2384)),
        ["A2"] = new PdfPageSize(OfficeIMO.OfficeLength.Points(1190),OfficeIMO.OfficeLength.Points(1684)),
        ["A3"] = new PdfPageSize(OfficeIMO.OfficeLength.Points(842),OfficeIMO.OfficeLength.Points(1190)),
        ["A4"] = new PdfPageSize(OfficeIMO.OfficeLength.Points(595),OfficeIMO.OfficeLength.Points(842)),
        ["A5"] = new PdfPageSize(OfficeIMO.OfficeLength.Points(420),OfficeIMO.OfficeLength.Points(595)),
        ["A6"] = new PdfPageSize(OfficeIMO.OfficeLength.Points(298),OfficeIMO.OfficeLength.Points(420)),
        ["A7"] = new PdfPageSize(OfficeIMO.OfficeLength.Points(210),OfficeIMO.OfficeLength.Points(298)),
        ["A8"] = new PdfPageSize(OfficeIMO.OfficeLength.Points(148),OfficeIMO.OfficeLength.Points(210)),
        ["A9"] = new PdfPageSize(OfficeIMO.OfficeLength.Points(105),OfficeIMO.OfficeLength.Points(148)),
        ["A10"] = new PdfPageSize(OfficeIMO.OfficeLength.Points(74),OfficeIMO.OfficeLength.Points(105)),
        ["B0"] = new PdfPageSize(OfficeIMO.OfficeLength.Points(2834),OfficeIMO.OfficeLength.Points(4008)),
        ["B1"] = new PdfPageSize(OfficeIMO.OfficeLength.Points(2004),OfficeIMO.OfficeLength.Points(2834)),
        ["B2"] = new PdfPageSize(OfficeIMO.OfficeLength.Points(1417),OfficeIMO.OfficeLength.Points(2004)),
        ["B3"] = new PdfPageSize(OfficeIMO.OfficeLength.Points(1000),OfficeIMO.OfficeLength.Points(1417)),
        ["B4"] = new PdfPageSize(OfficeIMO.OfficeLength.Points(708),OfficeIMO.OfficeLength.Points(1000)),
        ["B5"] = new PdfPageSize(OfficeIMO.OfficeLength.Points(498),OfficeIMO.OfficeLength.Points(708)),
        ["B6"] = new PdfPageSize(OfficeIMO.OfficeLength.Points(354),OfficeIMO.OfficeLength.Points(498)),
        ["B7"] = new PdfPageSize(OfficeIMO.OfficeLength.Points(249),OfficeIMO.OfficeLength.Points(354)),
        ["B8"] = new PdfPageSize(OfficeIMO.OfficeLength.Points(175),OfficeIMO.OfficeLength.Points(249)),
        ["B9"] = new PdfPageSize(OfficeIMO.OfficeLength.Points(124),OfficeIMO.OfficeLength.Points(175)),
        ["B10"] = new PdfPageSize(OfficeIMO.OfficeLength.Points(88),OfficeIMO.OfficeLength.Points(124)),
        ["Executive"] = new PdfPageSize(OfficeIMO.OfficeLength.Points(522),OfficeIMO.OfficeLength.Points(756)),
        ["Letter"] = new PdfPageSize(OfficeIMO.OfficeLength.Points(612),OfficeIMO.OfficeLength.Points(792)),
        ["Legal"] = new PdfPageSize(OfficeIMO.OfficeLength.Points(612),OfficeIMO.OfficeLength.Points(1008)),
        ["Ledger"] = new PdfPageSize(OfficeIMO.OfficeLength.Points(1224),OfficeIMO.OfficeLength.Points(792)),
        ["Tabloid"] = new PdfPageSize(OfficeIMO.OfficeLength.Points(792),OfficeIMO.OfficeLength.Points(1224)),
        ["LedgerOrTabloid"] = new PdfPageSize(OfficeIMO.OfficeLength.Points(1224),OfficeIMO.OfficeLength.Points(792))
    };

    /// <summary>ISO A0 size (2384 × 3370 pt).</summary>
    public static PdfPageSize A0 => KnownSizes["A0"];
    /// <summary>ISO A1 size (1684 × 2384 pt).</summary>
    public static PdfPageSize A1 => KnownSizes["A1"];
    /// <summary>ISO A2 size (1190 × 1684 pt).</summary>
    public static PdfPageSize A2 => KnownSizes["A2"];
    /// <summary>ISO A3 size (842 × 1190 pt).</summary>
    public static PdfPageSize A3 => KnownSizes["A3"];
    /// <summary>ISO A5 size (420 × 595 pt).</summary>
    public static PdfPageSize A5 => KnownSizes["A5"];
    /// <summary>ISO A4 size (595 × 842 pt).</summary>
    public static PdfPageSize A4 => KnownSizes["A4"];
    /// <summary>ISO A6 size (298 × 420 pt).</summary>
    public static PdfPageSize A6 => KnownSizes["A6"];
    /// <summary>ISO A7 size (210 × 298 pt).</summary>
    public static PdfPageSize A7 => KnownSizes["A7"];
    /// <summary>ISO A8 size (148 × 210 pt).</summary>
    public static PdfPageSize A8 => KnownSizes["A8"];
    /// <summary>ISO A9 size (105 × 148 pt).</summary>
    public static PdfPageSize A9 => KnownSizes["A9"];
    /// <summary>ISO A10 size (74 × 105 pt).</summary>
    public static PdfPageSize A10 => KnownSizes["A10"];
    /// <summary>ISO B0 size (2834 × 4008 pt).</summary>
    public static PdfPageSize B0 => KnownSizes["B0"];
    /// <summary>ISO B1 size (2004 × 2834 pt).</summary>
    public static PdfPageSize B1 => KnownSizes["B1"];
    /// <summary>ISO B2 size (1417 × 2004 pt).</summary>
    public static PdfPageSize B2 => KnownSizes["B2"];
    /// <summary>ISO B3 size (1000 × 1417 pt).</summary>
    public static PdfPageSize B3 => KnownSizes["B3"];
    /// <summary>ISO B4 size (708 × 1000 pt).</summary>
    public static PdfPageSize B4 => KnownSizes["B4"];
    /// <summary>ISO B5 size (498 × 708 pt).</summary>
    public static PdfPageSize B5 => KnownSizes["B5"];
    /// <summary>ISO B6 size (354 × 498 pt).</summary>
    public static PdfPageSize B6 => KnownSizes["B6"];
    /// <summary>ISO B7 size (249 × 354 pt).</summary>
    public static PdfPageSize B7 => KnownSizes["B7"];
    /// <summary>ISO B8 size (175 × 249 pt).</summary>
    public static PdfPageSize B8 => KnownSizes["B8"];
    /// <summary>ISO B9 size (124 × 175 pt).</summary>
    public static PdfPageSize B9 => KnownSizes["B9"];
    /// <summary>ISO B10 size (88 × 124 pt).</summary>
    public static PdfPageSize B10 => KnownSizes["B10"];
    /// <summary>US Executive size (522 × 756 pt).</summary>
    public static PdfPageSize Executive => KnownSizes["Executive"];
    /// <summary>US Letter size (612 × 792 pt).</summary>
    public static PdfPageSize Letter => KnownSizes["Letter"];
    /// <summary>US Legal size (612 × 1008 pt).</summary>
    public static PdfPageSize Legal => KnownSizes["Legal"];
    /// <summary>US Ledger size (1224 × 792 pt).</summary>
    public static PdfPageSize Ledger => KnownSizes["Ledger"];
    /// <summary>US Tabloid size (792 × 1224 pt).</summary>
    public static PdfPageSize Tabloid => KnownSizes["Tabloid"];
    /// <summary>Alias for US Ledger or Tabloid landscape size (1224 × 792 pt).</summary>
    public static PdfPageSize LedgerOrTabloid => KnownSizes["LedgerOrTabloid"];

    /// <summary>Known standard page-size names.</summary>
    public static IReadOnlyCollection<string> Names => KnownSizes.Keys.ToArray();

    /// <summary>Attempts to resolve a standard page-size name.</summary>
    public static bool TryGet(string? name, out PdfPageSize pageSize) {
        pageSize = default;
        if (string.IsNullOrWhiteSpace(name)) {
            return false;
        }

        string trimmed = name!.Trim();
        return KnownSizes.TryGetValue(trimmed, out pageSize);
    }

    /// <summary>Resolves a standard page-size name.</summary>
    public static PdfPageSize Get(string name) {
        if (TryGet(name, out PdfPageSize pageSize)) {
            return pageSize;
        }

        throw new ArgumentException("Unknown page size name '" + name + "'.", nameof(name));
    }
}
