namespace OfficeIMO.Invoicing;

/// <summary>Polish FA(3) document kind. Corrections retain signed differences rather than positive credit-note amounts.</summary>
public enum Fa3InvoiceKind {
    /// <summary>VAT.</summary>
    TaxInvoice,
    /// <summary>KOR.</summary>
    Correction,
    /// <summary>ZAL.</summary>
    Advance,
    /// <summary>ROZ.</summary>
    Settlement,
    /// <summary>UPR.</summary>
    Simplified,
    /// <summary>KOR_ZAL.</summary>
    AdvanceCorrection,
    /// <summary>KOR_ROZ.</summary>
    SettlementCorrection
}

/// <summary>A national VAT bucket. The field suffix identifies the legal bucket, not an inferred EN VAT rate.</summary>
public sealed class Fa3TaxSummary {
    /// <summary>Creates an explicit national declaration. Authoring validates the field suffix, monetary precision and currency context.</summary>
    public Fa3TaxSummary(string suffix, decimal basis, decimal? tax = null, decimal? taxInPln = null) {
        if (suffix == null) throw new ArgumentNullException(nameof(suffix));
        FieldSuffix = suffix; TaxableAmount = basis; TaxAmount = tax; TaxAmountInPln = taxInPln;
    }
    /// <summary>Suffix of P_13, such as 1, 6_2 or 11.</summary>
    public string FieldSuffix { get; }
    /// <summary>Source-declared P_13 amount, including signed correction differences.</summary>
    public decimal TaxableAmount { get; }
    /// <summary>Source-declared P_14 amount; null means absent.</summary>
    public decimal? TaxAmount { get; }
    /// <summary>Source-declared P_14_W amount in PLN; no exchange rate is inferred.</summary>
    public decimal? TaxAmountInPln { get; }
}

/// <summary>National line fields retained without assigning a unit code or fabricating omitted quantities and prices.</summary>
public sealed class Fa3InvoiceLineData {
    internal Fa3InvoiceLineData(string number, string? uuid, string? unit, string? rate,
        string? name, decimal? quantity, decimal? price, decimal? net, bool previous) {
        Number = number; UniqueIdentifier = uuid; UnitLabel = unit; TaxRateLabel = rate;
        Name = name; Quantity = quantity; UnitPrice = price; NetAmount = net; IsPreviousState = previous;
    }
    /// <summary>NrWierszaFa.</summary>
    public string Number { get; }
    /// <summary>UU_ID.</summary>
    public string? UniqueIdentifier { get; }
    /// <summary>Literal P_8A unit text.</summary>
    public string? UnitLabel { get; }
    /// <summary>Literal P_12 national rate label.</summary>
    public string? TaxRateLabel { get; }
    /// <summary>P_7 item name.</summary>
    public string? Name { get; }
    /// <summary>P_8B quantity, or null when omitted.</summary>
    public decimal? Quantity { get; }
    /// <summary>P_9A net unit price, or null when omitted.</summary>
    public decimal? UnitPrice { get; }
    /// <summary>P_11 net amount, or null when omitted.</summary>
    public decimal? NetAmount { get; }
    /// <summary>True for StanPrzed=1: this is a previous-state row, not an additional billed line.</summary>
    public bool IsPreviousState { get; }
}

/// <summary>FA(3) source data and a partial common invoice projection. Parsing does not establish fiscal or KSeF acceptance.</summary>
public sealed class Fa3InvoiceReadResult {
    private readonly byte[] _source;
    internal Fa3InvoiceReadResult(byte[] source, Invoice invoice, Fa3InvoiceKind kind, DateTimeOffset created,
        string? systemInfo, string? place, decimal total, IReadOnlyList<Fa3TaxSummary> taxes,
        IReadOnlyList<Fa3InvoiceLineData> lines, IReadOnlyList<InvoiceDiagnostic> unmapped,
        IReadOnlyList<InvoiceDiagnostic> projection) {
        _source = source; Invoice = invoice; Kind = kind; CreatedAt = created; SystemInfo = systemInfo;
        PlaceOfIssue = place; DeclaredTotal = total; TaxSummaries = taxes; Lines = lines;
        UnmappedData = unmapped; CommonMappingDiagnostics = projection;
    }
    /// <summary>Editable common fields. This projection can be incomplete; inspect both diagnostic collections before using it.</summary>
    public Invoice Invoice { get; }
    /// <summary>Explicit national document kind.</summary>
    public Fa3InvoiceKind Kind { get; }
    /// <summary>DataWytworzeniaFa, including its explicit UTC offset.</summary>
    public DateTimeOffset CreatedAt { get; }
    /// <summary>Optional generating-system identity.</summary>
    public string? SystemInfo { get; }
    /// <summary>Optional P_1M place of issue.</summary>
    public string? PlaceOfIssue { get; }
    /// <summary>Literal P_15 total. For advances, corrections and settlements it is not relabeled as an EN amount due.</summary>
    public decimal DeclaredTotal { get; }
    /// <summary>Source-declared national VAT buckets in schema order.</summary>
    public IReadOnlyList<Fa3TaxSummary> TaxSummaries { get; }
    /// <summary>Every observed FaWiersz, including previous-state and partially specified rows.</summary>
    public IReadOnlyList<Fa3InvoiceLineData> Lines { get; }
    /// <summary>Business fields outside the supported national/common read mapping, with indexed namespace-aware locations.</summary>
    public IReadOnlyList<InvoiceDiagnostic> UnmappedData { get; }
    /// <summary>National semantics that cannot safely be expressed by the common invoice projection.</summary>
    public IReadOnlyList<InvoiceDiagnostic> CommonMappingDiagnostics { get; }
    /// <summary>Returns the original bytes unchanged, even after common-model edits. This is preservation, not a model rewrite.</summary>
    public byte[] GetOriginalBytes() => (byte[])_source.Clone();
    /// <summary>Converts common fields only when no observed business data or national semantics would be discarded.</summary>
    public byte[] WriteCommonInvoice(InvoiceXmlOptions target) {
        if (target == null) throw new ArgumentNullException(nameof(target));
        if (UnmappedData.Count != 0 || CommonMappingDiagnostics.Any(item => item.Severity == InvoiceDiagnosticSeverity.Error))
            throw new InvalidDataException("FA(3) conversion is blocked by unmapped data or national semantics. Inspect both diagnostic collections.");
        return InvoiceSerializer.Write(Invoice, target);
    }
}
