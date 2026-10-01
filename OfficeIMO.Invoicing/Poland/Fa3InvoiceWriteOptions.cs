namespace OfficeIMO.Invoicing;

/// <summary>Legal-basis field selected explicitly for a Polish VAT exemption.</summary>
public enum Fa3ExemptionBasisKind {
    /// <summary>P_19A: national provision.</summary>
    NationalProvision,
    /// <summary>P_19B: EU directive provision.</summary>
    EuDirective,
    /// <summary>P_19C: another legal basis.</summary>
    Other
}

/// <summary>Explicit Polish margin procedure.</summary>
public enum Fa3MarginProcedure {
    /// <summary>No margin procedure (P_PMarzyN=1).</summary>
    None,
    /// <summary>Travel services (P_PMarzy_2=1).</summary>
    Travel,
    /// <summary>Used goods (P_PMarzy_3_1=1).</summary>
    UsedGoods,
    /// <summary>Works of art (P_PMarzy_3_2=1).</summary>
    Art,
    /// <summary>Collector's items and antiques (P_PMarzy_3_3=1).</summary>
    CollectorsAndAntiques
}

/// <summary>National declarations supplied by the issuer. They are not inferred from party names, item descriptions or EN type codes.</summary>
/// <remarks>This authoring contract declares no new means of transport. Such transactions require additional national data.</remarks>
public sealed class Fa3InvoiceAnnotations {
    /// <summary>P_16: cash accounting.</summary>
    public bool CashAccounting { get; set; }
    /// <summary>P_17: self-billing.</summary>
    public bool SelfBilling { get; set; }
    /// <summary>P_18: reverse charge.</summary>
    public bool ReverseCharge { get; set; }
    /// <summary>P_18A: split payment.</summary>
    public bool SplitPayment { get; set; }
    /// <summary>P_23: simplified triangular transaction.</summary>
    public bool TriangularTransaction { get; set; }
    /// <summary>Explicit exemption field kind. Supply together with its legal basis.</summary>
    public Fa3ExemptionBasisKind? ExemptionBasisKind { get; set; }
    /// <summary>Literal exemption legal basis; null declares no VAT exemption.</summary>
    public string? ExemptionLegalBasis { get; set; }
    /// <summary>Explicit margin procedure.</summary>
    public Fa3MarginProcedure MarginProcedure { get; set; }
}

/// <summary>Issuer-supplied national totals for corrections, advances, settlements, simplified invoices, foreign currency or margin procedures.</summary>
/// <remarks>These declarations are retained as supplied. They do not establish fiscal correctness and are not recalculated as EN totals.</remarks>
public sealed class Fa3FiscalAmounts {
    /// <summary>Creates immutable national amount declarations with at most thirteen distinct VAT buckets.</summary>
    public Fa3FiscalAmounts(decimal total, IEnumerable<Fa3TaxSummary> taxes) {
        if (taxes == null) throw new ArgumentNullException(nameof(taxes));
        List<Fa3TaxSummary> snapshot = taxes.Take(14).ToList();
        if (snapshot.Count > 13 || snapshot.Any(item => item == null))
            throw new ArgumentException("Supply at most thirteen non-null national VAT buckets.", nameof(taxes));
        Total = total; Taxes = snapshot.AsReadOnly();
    }
    /// <summary>P_15, including its signed or advance/settlement meaning.</summary>
    public decimal Total { get; }
    /// <summary>Explicit P_13/P_14 declarations.</summary>
    public IReadOnlyList<Fa3TaxSummary> Taxes { get; }
}

/// <summary>Reference to a previous advance invoice without inventing a KSeF identifier.</summary>
public sealed class Fa3AdvanceInvoiceReference {
    /// <summary>Creates a reference with an issuer invoice number, a KSeF number, or both.</summary>
    public Fa3AdvanceInvoiceReference(string? number, string? ksefNumber = null) { Number = number; KsefNumber = ksefNumber; }
    /// <summary>Optional issuer-assigned invoice number.</summary>
    public string? Number { get; }
    /// <summary>Optional actual KSeF identifier.</summary>
    public string? KsefNumber { get; }
}

/// <summary>Full advance order rows or explicit advance-correction difference rows, using the shared line model.</summary>
public sealed class Fa3Order {
    /// <summary>Creates an order with explicit gross value and at most 10,000 lines. Line objects remain editable by the caller.</summary>
    public Fa3Order(decimal total, IEnumerable<InvoiceLine> lines) : this(total, lines, null, null) { }

    private Fa3Order(decimal total, IEnumerable<InvoiceLine> lines, decimal? previousTotal, IEnumerable<decimal?>? differenceTaxAmounts) {
        if (lines == null) throw new ArgumentNullException(nameof(lines));
        List<InvoiceLine> snapshot = lines.Take(10_001).ToList();
        if (snapshot.Count == 0 || snapshot.Count > 10_000 || snapshot.Any(item => item == null))
            throw new ArgumentException("Supply one to 10,000 non-null order lines.", nameof(lines));
        Total = total; Lines = snapshot.AsReadOnly(); PreviousTotal = previousTotal;
        if (differenceTaxAmounts != null) {
            List<decimal?> taxes = differenceTaxAmounts.Take(10_001).ToList();
            if (taxes.Count != snapshot.Count) throw new ArgumentException("Supply one explicit tax declaration per order difference row.", nameof(differenceTaxAmounts));
            DifferenceTaxAmounts = taxes.AsReadOnly();
        }
    }
    /// <summary>Creates an advance correction with signed order differences and the original/revised full gross totals.</summary>
    /// <remarks>The differences must reconcile those totals. An unchanged order needs before/after rows and is rejected by this bounded contract.</remarks>
    public static Fa3Order CorrectionDifferences(decimal previousTotal, decimal revisedTotal, IEnumerable<InvoiceLine> differences, IEnumerable<decimal?> differenceTaxAmounts) =>
        new Fa3Order(revisedTotal, differences, previousTotal, differenceTaxAmounts ?? throw new ArgumentNullException(nameof(differenceTaxAmounts)));
    /// <summary>WartoscZamowienia: explicit gross value of the full order.</summary>
    public decimal Total { get; }
    /// <summary>Original gross order value for the difference-row correction contract; absent for a full advance order.</summary>
    public decimal? PreviousTotal { get; }
    /// <summary>Explicit P_11VatZ per difference row in line order, distinct from corrected advance payment amounts. Null entries are permitted only for nontaxable rows.</summary>
    public IReadOnlyList<decimal?>? DifferenceTaxAmounts { get; }
    /// <summary>Full ordered items for an advance, or signed differences for an advance correction.</summary>
    public IReadOnlyList<InvoiceLine> Lines { get; }
}

/// <summary>Explicit FA(3) authoring declarations, separate from EN profiles and releases.</summary>
public sealed class Fa3InvoiceWriteOptions {
    /// <summary>Creates the national authoring contract. The issuer supplies the kind, creation instant and annotations explicitly.</summary>
    public Fa3InvoiceWriteOptions(Fa3InvoiceKind kind, DateTimeOffset createdAt, Fa3InvoiceAnnotations annotations) {
        Kind = kind; CreatedAt = createdAt; Annotations = annotations ?? throw new ArgumentNullException(nameof(annotations));
    }
    /// <summary>National document kind.</summary>
    public Fa3InvoiceKind Kind { get; }
    /// <summary>Creation instant, emitted in UTC.</summary>
    public DateTimeOffset CreatedAt { get; }
    /// <summary>Explicit national annotations.</summary>
    public Fa3InvoiceAnnotations Annotations { get; }
    /// <summary>Generating system name.</summary>
    public string? SystemInfo { get; set; } = "OfficeIMO";
    /// <summary>Optional place of issue.</summary>
    public string? PlaceOfIssue { get; set; }
    /// <summary>Required national totals for kinds that cannot be calculated as an ordinary invoice; also supplies per-bucket PLN tax for foreign currency.</summary>
    public Fa3FiscalAmounts? FiscalAmounts { get; set; }
    /// <summary>Optional correction reason.</summary>
    public string? CorrectionReason { get; set; }
    /// <summary>Optional TypKorekty value 1, 2 or 3, supplied by the issuer.</summary>
    public int? CorrectionTimingCode { get; set; }
    /// <summary>KSeF numbers keyed by the common preceding-invoice reference number. An absent key explicitly identifies a correction reference outside KSeF.</summary>
    public IDictionary<string, string> CorrectionKsefNumbers { get; } = new Dictionary<string, string>(StringComparer.Ordinal);
    /// <summary>Previous advance invoice references for settlements.</summary>
    public IList<Fa3AdvanceInvoiceReference> AdvanceInvoiceReferences { get; } = new List<Fa3AdvanceInvoiceReference>();
    /// <summary>Full order rows for advances; explicit original/revised totals and difference rows for advance corrections.</summary>
    public Fa3Order? Order { get; set; }
    /// <summary>Optional original advance or settlement total for P_15ZK in a correction.</summary>
    public decimal? PreviousAdvanceOrSettlementTotal { get; set; }
    /// <summary>Literal national unit labels keyed by common unit code. Unrecognized codes require an explicit label.</summary>
    public IDictionary<string, string> UnitLabels { get; } = new Dictionary<string, string>(StringComparer.Ordinal);
    /// <summary>Explicit P_12 labels keyed by line identifier, required when a common category cannot distinguish national meanings.</summary>
    public IDictionary<string, string> LineTaxLabels { get; } = new Dictionary<string, string>(StringComparer.Ordinal);
}
