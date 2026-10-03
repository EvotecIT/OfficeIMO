namespace OfficeIMO.Invoicing.Tests;

public class Fa3InvoiceWriterTests {
    [Theory]
    [InlineData(true)]
    [InlineData(false)]
    public void Oversized_contact_email_is_rejected_before_schema_pattern_matching(bool seller) {
        Invoice invoice = Fa3InvoiceFixture.Create();
        (seller ? invoice.Seller : invoice.Buyer).Contact = new InvoiceContact {
            Email = new string('@', 8192) + "\n"
        };
        string location = (seller ? "Seller" : "Buyer") + ".Contact.Email";

        IReadOnlyList<InvoiceDiagnostic> diagnostics = Fa3InvoiceWriter.Inspect(invoice, Fa3InvoiceFixture.Options());
        InvoiceDiagnostic diagnostic = Assert.Single(diagnostics, item => item.Location == location);
        Assert.Contains("255 characters", diagnostic.Message, StringComparison.Ordinal);
        Assert.Throws<InvalidDataException>(() => Fa3InvoiceWriter.Write(invoice, Fa3InvoiceFixture.Options()));
    }

    [Fact]
    public void Contact_email_schema_pattern_still_checks_bounded_values() {
        Invoice invoice = Fa3InvoiceFixture.Create();
        invoice.Seller.Contact = new InvoiceContact { Email = "missing-at.example.test" };
        Assert.Contains(Fa3InvoiceWriter.Inspect(invoice, Fa3InvoiceFixture.Options()), diagnostic =>
            diagnostic.Location == "Seller.Contact.Email" &&
            diagnostic.Message.Contains("schema pattern", StringComparison.Ordinal));

        invoice.Seller.Contact.Email = "seller@example.test";
        Assert.Empty(Fa3InvoiceWriter.Inspect(invoice, Fa3InvoiceFixture.Options()));
        Assert.NotEmpty(Fa3InvoiceWriter.Write(invoice, Fa3InvoiceFixture.Options()));
    }

    [Fact]
    public void AdvanceCorrectionOrderDoesNotInferSignedNationalVatFromEnRounding() {
        Invoice invoice = Fa3InvoiceFixture.Create(); invoice.TypeCode = "384";
        invoice.PrecedingInvoices.Add(new InvoiceReference("original", invoice.IssueDate)); invoice.Lines.Clear();
        Fa3InvoiceWriteOptions options = Fa3InvoiceFixture.Options(Fa3InvoiceKind.AdvanceCorrection);
        options.FiscalAmounts = new Fa3FiscalAmounts(-0.50m, new[] { new Fa3TaxSummary("1", -0.40m, -0.10m) });
        var rows = new[] { 23m, 5m, 3m }.Select((rate, index) => new InvoiceLine { Id = (index + 1).ToString(), Name = "Difference", UnitCode = "HUR", Quantity = -1, UnitPrice = 0.50m, Tax = new InvoiceTaxCategory { Code = "S", Rate = rate } }).ToArray();
        options.Order = Fa3Order.CorrectionDifferences(10, 8.33m, rows, new decimal?[] { -0.12m, -0.03m, -0.02m });
        Assert.Empty(Fa3InvoiceWriter.Inspect(invoice, options));
        Assert.Equal("8.33", System.Xml.Linq.XDocument.Parse(Encoding.UTF8.GetString(Fa3InvoiceWriter.Write(invoice, options))).Descendants().Single(element => element.Name.LocalName == "WartoscZamowienia").Value);
        options.Order = Fa3Order.CorrectionDifferences(10, 8.33m, rows, new decimal?[] { -0.11m, -0.02m, -0.01m });
        Assert.Contains(Fa3InvoiceWriter.Inspect(invoice, options), diagnostic => diagnostic.Location == "Order.Total");
    }
    [Fact]
    public void NationalScalarDictionariesMatchThePinnedSchemaEnumerations() {
        string directory = Path.Combine(AppContext.BaseDirectory, "Fixtures", "FA3", "Schemas");
        System.Xml.Linq.XNamespace ns = "http://www.w3.org/2001/XMLSchema";
        string[] Codes(string file, string type) => System.Xml.Linq.XDocument.Load(Path.Combine(directory, file)).Descendants(ns + "simpleType")
            .Single(element => (string?)element.Attribute("name") == type).Descendants(ns + "enumeration").Select(element => element.Attribute("value")!.Value).OrderBy(value => value, StringComparer.Ordinal).ToArray();
        Assert.Equal(Codes("FA3.xsd", "TKodWaluty"), Fa3ScalarContract.Currencies.OrderBy(value => value, StringComparer.Ordinal));
        Assert.Equal(Codes("KodyKrajow_v10-0E.xsd", "TKodKraju"), Fa3ScalarContract.Countries.OrderBy(value => value, StringComparer.Ordinal));
    }
    [Fact]
    public void OrdinaryPlnAmountsCannotOverrideTheCalculatedTaxAndPayableTotal() {
        Invoice invoice = Fa3InvoiceFixture.Create(); Fa3InvoiceWriteOptions options = Fa3InvoiceFixture.Options();
        options.FiscalAmounts = new Fa3FiscalAmounts(0, Array.Empty<Fa3TaxSummary>());
        Assert.Contains(Fa3InvoiceWriter.Inspect(invoice, options), diagnostic => diagnostic.Location == "FiscalAmounts");
        Assert.Throws<InvalidDataException>(() => Fa3InvoiceWriter.Write(invoice, options));
    }
    [Fact]
    public void OrdinaryNativeAuthoringUsesTheCanonicalFinancialCalculationAndExplicitTime() {
        Invoice invoice = Fa3InvoiceFixture.Create(); Fa3InvoiceWriteOptions options = Fa3InvoiceFixture.Options();
        byte[] xml = Fa3InvoiceWriter.Write(invoice, options);
        Fa3InvoiceReadResult parsed = Fa3InvoiceReader.Read(xml);
        Assert.Equal(InvoiceCalculator.Calculate(invoice).TaxInclusiveTotal, parsed.DeclaredTotal);
        Fa3TaxSummary tax = Assert.Single(parsed.TaxSummaries);
        Assert.Equal("1", tax.FieldSuffix); Assert.Equal(100, tax.TaxableAmount); Assert.Equal(23, tax.TaxAmount);
        Assert.Equal("godz.", Assert.Single(parsed.Lines).UnitLabel);
        Assert.Equal(options.CreatedAt, parsed.CreatedAt);
        Assert.Equal(xml, Fa3InvoiceWriter.Write(invoice, options));
        Assert.Null(invoice.DeclaredTotals); Assert.Empty(invoice.DeclaredTaxes);
    }

    [Fact]
    public void SignedAmountsRequireADeclarationAtTheNationalRoundingBoundary() {
        Invoice invoice = Fa3InvoiceFixture.Create(); Fa3InvoiceWriteOptions options = Fa3InvoiceFixture.Options(Fa3InvoiceKind.Correction);
        invoice.TypeCode = "384"; invoice.PrecedingInvoices.Add(new InvoiceReference("original", invoice.IssueDate.AddDays(-1)));
        invoice.Lines[0].Quantity = -1; invoice.Lines[0].UnitPrice = 0.005m;
        options.FiscalAmounts = new Fa3FiscalAmounts(-0.01m, new[] { new Fa3TaxSummary("1", -0.01m, 0) });
        Assert.Contains(Fa3InvoiceWriter.Inspect(invoice, options), diagnostic => diagnostic.Location == "Lines[0].DeclaredNetAmount");
        invoice.Lines[0].DeclaredNetAmount = -0.01m;
        Assert.Equal(-0.01m, Assert.Single(Fa3InvoiceReader.Read(Fa3InvoiceWriter.Write(invoice, options)).Lines).NetAmount);
        invoice.Lines[0].DeclaredNetAmount = -100m;
        Assert.Throws<InvalidDataException>(() => Fa3InvoiceWriter.Write(invoice, options));
    }

    [Theory]
    [InlineData("reference")]
    [InlineData("unknown-unit")]
    [InlineData("wrong-label")]
    [InlineData("sepa")]
    [InlineData("foreign-tax")]
    [InlineData("advance")]
    [InlineData("null-party")]
    public void UnsupportedOrMissingDataBlocksOutputWithAFieldDiagnostic(string trigger) {
        Invoice invoice = Fa3InvoiceFixture.Create(); Fa3InvoiceWriteOptions options = Fa3InvoiceFixture.Options();
        switch (trigger) {
            case "reference": invoice.PaymentReference = "retain-me"; break;
            case "unknown-unit": invoice.Lines[0].UnitCode = "CUSTOM"; break;
            case "wrong-label": options.LineTaxLabels["1"] = "zw"; break;
            case "sepa": invoice.Payments.Add(new InvoicePayment { MeansCode = "58" }); break;
            case "foreign-tax": invoice.Currency = "EUR"; options.FiscalAmounts = new Fa3FiscalAmounts(123, new[] { new Fa3TaxSummary("1", 100, 23) }); break;
            case "advance": invoice.TypeCode = "386"; options = Fa3InvoiceFixture.Options(Fa3InvoiceKind.Advance); break;
            case "null-party": invoice.Seller = null!; break;
        }
        Assert.Contains(Fa3InvoiceWriter.Inspect(invoice, options), diagnostic => diagnostic.Code == "FA3-AUTHORING" && diagnostic.Severity == InvoiceDiagnosticSeverity.Error && diagnostic.Location.Length != 0);
        Assert.Throws<InvalidDataException>(() => Fa3InvoiceWriter.Write(invoice, options));
    }
}
