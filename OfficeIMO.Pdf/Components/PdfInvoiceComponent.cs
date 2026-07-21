using System.Globalization;

namespace OfficeIMO.Pdf;

/// <summary>Postal or business identity used by the invoice recipe.</summary>
public sealed class PdfInvoiceParty {
    /// <summary>Creates a named party with optional address/details lines.</summary>
    public PdfInvoiceParty(string name, IEnumerable<string>? details = null) {
        Guard.NotNullOrWhiteSpace(name, nameof(name));
        Name = name;
        Details = details?.Where(static value => !string.IsNullOrWhiteSpace(value)).ToArray() ?? Array.Empty<string>();
    }

    /// <summary>Display name.</summary>
    public string Name { get; }
    /// <summary>Address, tax, or contact lines.</summary>
    public IReadOnlyList<string> Details { get; }
    internal string ToDisplayText() => Details.Count == 0 ? Name : Name + "\n" + string.Join("\n", Details);
}

/// <summary>One quantity/price line in the invoice recipe.</summary>
public sealed class PdfInvoiceLine {
    /// <summary>Creates an invoice line.</summary>
    public PdfInvoiceLine(string description, decimal quantity, decimal unitPrice, decimal taxRate = 0M) {
        Guard.NotNullOrWhiteSpace(description, nameof(description));
#pragma warning disable CA1512 // Cross-target guard supports netstandard2.0 and net472.
        if (quantity < 0M) throw new ArgumentOutOfRangeException(nameof(quantity));
        if (unitPrice < 0M) throw new ArgumentOutOfRangeException(nameof(unitPrice));
        if (taxRate < 0M) throw new ArgumentOutOfRangeException(nameof(taxRate));
#pragma warning restore CA1512
        Description = description;
        Quantity = quantity;
        UnitPrice = unitPrice;
        TaxRate = taxRate;
    }

    /// <summary>Line description.</summary>
    public string Description { get; }
    /// <summary>Quantity.</summary>
    public decimal Quantity { get; }
    /// <summary>Unit price before tax.</summary>
    public decimal UnitPrice { get; }
    /// <summary>Tax rate expressed as a fraction, for example 0.23.</summary>
    public decimal TaxRate { get; }
    /// <summary>Line subtotal before tax.</summary>
    public decimal Subtotal => Quantity * UnitPrice;
    /// <summary>Line tax amount.</summary>
    public decimal Tax => Subtotal * TaxRate;
    /// <summary>Line total including tax.</summary>
    public decimal Total => Subtotal + Tax;
}

/// <summary>A deterministic invoice recipe implemented over flow paragraphs and tables.</summary>
public sealed class PdfInvoiceComponent : IPdfComponent {
    private static readonly string[] InvoiceHeaders = { "Description", "Quantity", "Unit price", "Tax", "Total" };
    private readonly PdfInvoiceLine[] _lines;
    private readonly CultureInfo _culture;

    /// <summary>Creates an invoice recipe.</summary>
    public PdfInvoiceComponent(
        string invoiceNumber,
        DateTime issueDate,
        PdfInvoiceParty seller,
        PdfInvoiceParty customer,
        IEnumerable<PdfInvoiceLine> lines,
        string currencyCode = "USD",
        CultureInfo? culture = null,
        DateTime? dueDate = null) {
        Guard.NotNullOrWhiteSpace(invoiceNumber, nameof(invoiceNumber));
        Guard.NotNull(seller, nameof(seller));
        Guard.NotNull(customer, nameof(customer));
        Guard.NotNull(lines, nameof(lines));
        Guard.NotNullOrWhiteSpace(currencyCode, nameof(currencyCode));
        InvoiceNumber = invoiceNumber;
        IssueDate = issueDate;
        DueDate = dueDate;
        Seller = seller;
        Customer = customer;
        CurrencyCode = currencyCode.ToUpperInvariant();
        _culture = culture ?? CultureInfo.InvariantCulture;
        _lines = lines.ToArray();
        if (_lines.Length == 0) throw new ArgumentException("An invoice requires at least one line.", nameof(lines));
        if (_lines.Any(static line => line == null)) throw new ArgumentException("Invoice lines cannot contain null entries.", nameof(lines));
    }

    /// <summary>Invoice identifier.</summary>
    public string InvoiceNumber { get; }
    /// <summary>Issue date.</summary>
    public DateTime IssueDate { get; }
    /// <summary>Optional due date.</summary>
    public DateTime? DueDate { get; }
    /// <summary>Seller identity.</summary>
    public PdfInvoiceParty Seller { get; }
    /// <summary>Customer identity.</summary>
    public PdfInvoiceParty Customer { get; }
    /// <summary>ISO-style currency code placed beside formatted numbers.</summary>
    public string CurrencyCode { get; }
    /// <summary>Total before tax.</summary>
    public decimal Subtotal => _lines.Sum(static line => line.Subtotal);
    /// <summary>Total tax.</summary>
    public decimal Tax => _lines.Sum(static line => line.Tax);
    /// <summary>Total including tax.</summary>
    public decimal Total => _lines.Sum(static line => line.Total);

    /// <inheritdoc />
    public void Compose(PdfItemCompose content) {
        Guard.NotNull(content, nameof(content));
        content.H1("Invoice " + InvoiceNumber)
            .Table(new[] {
                new[] { "Seller", Seller.ToDisplayText(), "Customer", Customer.ToDisplayText() },
                new[] { "Issued", IssueDate.ToString("d", _culture), "Due", DueDate?.ToString("d", _culture) ?? string.Empty }
            }, style: new PdfTableStyle { HeaderRowCount = 0, RowStripeFill = null, SpacingAfter = 12 });

        var rows = new List<string[]> {
            InvoiceHeaders
        };
        rows.AddRange(_lines.Select(line => new[] {
            line.Description,
            line.Quantity.ToString("0.##", _culture),
            Money(line.UnitPrice),
            line.TaxRate.ToString("P0", _culture),
            Money(line.Total)
        }));
        content.Table(rows, style: new PdfTableStyle { HeaderRowCount = 1, SpacingAfter = 10 });
        content.Paragraph(paragraph => paragraph
            .Bold("Subtotal: ").Text(Money(Subtotal))
            .LineBreak().Bold("Tax: ").Text(Money(Tax))
            .LineBreak().Bold("Total: ").Text(Money(Total)), PdfAlign.Right);
    }

    private string Money(decimal value) => value.ToString("N2", _culture) + " " + CurrencyCode;
}
