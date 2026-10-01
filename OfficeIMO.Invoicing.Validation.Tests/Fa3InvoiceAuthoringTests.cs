using OfficeIMO.Invoicing.Tests;

namespace OfficeIMO.Invoicing.Validation.Tests;

public class Fa3InvoiceAuthoringTests {
    [Theory]
    [InlineData("CreatedAt")]
    [InlineData("IssueDate")]
    [InlineData("DueDate")]
    [InlineData("TaxPointDate")]
    [InlineData("Period.Start")]
    [InlineData("Period.End")]
    [InlineData("Seller.Contact.Email")]
    [InlineData("Seller.TaxRegistrations")]
    [InlineData("Seller.Address.CountryCode")]
    [InlineData("Currency")]
    public void AuthoringRejectsSupportedScalarsOutsideThePinnedSchema(string field) {
        Invoice invoice = Fa3InvoiceFixture.Create(); Fa3InvoiceWriteOptions options = Fa3InvoiceFixture.Options();
        var document = System.Xml.Linq.XDocument.Parse(System.Text.Encoding.UTF8.GetString(Fa3InvoiceWriter.Write(invoice, options)));
        var ns = document.Root!.Name.Namespace;
        switch (field) {
            case "CreatedAt": options = new Fa3InvoiceWriteOptions(options.Kind, new DateTimeOffset(2024, 1, 1, 0, 0, 0, TimeSpan.Zero), options.Annotations); document.Descendants(ns + "DataWytworzeniaFa").Single().Value = "2024-01-01T00:00:00Z"; break;
            case "IssueDate": invoice.IssueDate = new DateTime(2005, 1, 1); document.Descendants(ns + "P_1").Single().Value = "2005-01-01"; break;
            case "DueDate": invoice.DueDate = new DateTime(2015, 1, 1); document.Descendants(ns + "Fa").Single().Add(new System.Xml.Linq.XElement(ns + "Platnosc", new System.Xml.Linq.XElement(ns + "TerminPlatnosci", new System.Xml.Linq.XElement(ns + "Termin", "2015-01-01")))); break;
            case "TaxPointDate": invoice.TaxPointDate = new DateTime(2050, 1, 2); document.Descendants(ns + "P_2").Single().AddAfterSelf(new System.Xml.Linq.XElement(ns + "P_6", "2050-01-02")); break;
            case "Period.Start":
            case "Period.End":
                invoice.Period = field == "Period.Start" ? new InvoicePeriod { Start = new DateTime(2005, 1, 1), End = invoice.IssueDate } : new InvoicePeriod { Start = invoice.IssueDate, End = new DateTime(2050, 1, 2) };
                document.Descendants(ns + "P_2").Single().AddAfterSelf(new System.Xml.Linq.XElement(ns + "OkresFa", new System.Xml.Linq.XElement(ns + "P_6_Od", invoice.Period.Start!.Value.ToString("yyyy-MM-dd")), new System.Xml.Linq.XElement(ns + "P_6_Do", invoice.Period.End!.Value.ToString("yyyy-MM-dd")))); break;
            case "Seller.Contact.Email": invoice.Seller.Contact = new InvoiceContact { Email = "x" }; document.Descendants(ns + "Podmiot1").Single().Add(new System.Xml.Linq.XElement(ns + "DaneKontaktowe", new System.Xml.Linq.XElement(ns + "Email", "x"))); break;
            case "Seller.TaxRegistrations": invoice.Seller.TaxRegistrations[0].Identifier = "0000000000"; document.Descendants(ns + "NIP").First().Value = "0000000000"; break;
            case "Seller.Address.CountryCode": invoice.Seller.Address.CountryCode = "ZZ"; document.Descendants(ns + "KodKraju").First().Value = "ZZ"; break;
            case "Currency": invoice.Currency = "ZZZ"; document.Descendants(ns + "KodWaluty").Single().Value = "ZZZ"; options.FiscalAmounts = new Fa3FiscalAmounts(123, new[] { new Fa3TaxSummary("1", 100, 23, 98.90m) }); break;
        }
        var validator = new Fa3SchemaValidator(Fa3SchemaBundle.LoadDirectory(Path.Combine(AppContext.BaseDirectory, "Fixtures", "FA3", "Schemas")));
        Assert.False(validator.Validate(System.Text.Encoding.UTF8.GetBytes(document.ToString())).IsValid);
        Assert.Contains(Fa3InvoiceWriter.Inspect(invoice, options), diagnostic => diagnostic.Location == field);
        Assert.Throws<InvalidDataException>(() => Fa3InvoiceWriter.Write(invoice, options));
    }
    [Theory]
    [InlineData("foreign")]
    [InlineData("exempt")]
    [InlineData("margin")]
    [InlineData("period-payment")]
    [InlineData("signed-order")]
    public void NationalFinancialAndPaymentDeclarationsHaveQualifiedMappings(string scenario) {
        Invoice invoice = Fa3InvoiceFixture.Create(); Fa3InvoiceWriteOptions options = Fa3InvoiceFixture.Options();
        switch (scenario) {
            case "signed-order":
                options = Fa3InvoiceFixture.Options(Fa3InvoiceKind.AdvanceCorrection); invoice.TypeCode = "384";
                invoice.Lines.Clear(); invoice.PrecedingInvoices.Add(new InvoiceReference("original", invoice.IssueDate));
                options.FiscalAmounts = new Fa3FiscalAmounts(-0.50m, new[] { new Fa3TaxSummary("1", -0.40m, -0.10m) });
                var rows = new[] { 23m, 5m, 3m }.Select((rate, index) => new InvoiceLine { Id = (index + 1).ToString(), Name = "Difference", UnitCode = "HUR", Quantity = -1, UnitPrice = 0.50m, Tax = new InvoiceTaxCategory { Code = "S", Rate = rate } }).ToArray();
                options.Order = Fa3Order.CorrectionDifferences(10, 8.33m, rows, new decimal?[] { -0.12m, -0.03m, -0.02m }); break;
            case "foreign":
                invoice.Currency = "EUR"; invoice.TaxCurrency = "PLN"; invoice.TaxAmountInAccountingCurrency = 98.90m;
                options.FiscalAmounts = new Fa3FiscalAmounts(123, new[] { new Fa3TaxSummary("1", 100, 23, 98.90m) }); break;
            case "exempt":
                invoice.Lines[0].Tax = new InvoiceTaxCategory { Code = "E", Rate = 0 };
                options.Annotations.ExemptionBasisKind = Fa3ExemptionBasisKind.NationalProvision;
                options.Annotations.ExemptionLegalBasis = "Explicit issuer-supplied legal basis"; break;
            case "margin":
                invoice.Lines[0].Tax = new InvoiceTaxCategory { Code = "O", Rate = null };
                options.Annotations.MarginProcedure = Fa3MarginProcedure.UsedGoods;
                options.FiscalAmounts = new Fa3FiscalAmounts(100, new[] { new Fa3TaxSummary("11", 100) }); break;
            case "period-payment":
                invoice.Period = new InvoicePeriod { Start = invoice.IssueDate.AddDays(-30), End = invoice.IssueDate };
                invoice.DueDate = invoice.IssueDate.AddDays(30);
                invoice.Payments.Add(new InvoicePayment { MeansCode = "30", Account = new InvoiceBankAccount { Identifier = "PL61109010140000071219812874", ProviderIdentifier = "WBKPPLPP" } }); break;
        }
        byte[] xml = Fa3InvoiceWriter.Write(invoice, options);
        var validator = new Fa3SchemaValidator(Fa3SchemaBundle.LoadDirectory(Path.Combine(AppContext.BaseDirectory, "Fixtures", "FA3", "Schemas")));
        Fa3SchemaValidationResult validation = validator.Validate(xml);
        Assert.True(validation.IsValid, string.Join(Environment.NewLine, validation.Diagnostics.Select(item => item.Message)));
        Fa3InvoiceReadResult parsed = Fa3InvoiceReader.Read(xml);
        Assert.Equal(options.FiscalAmounts?.Total ?? InvoiceCalculator.Calculate(invoice).TaxInclusiveTotal, parsed.DeclaredTotal);
        if (scenario == "foreign") Assert.Equal(98.90m, Assert.Single(parsed.TaxSummaries).TaxAmountInPln);
        if (scenario == "period-payment") Assert.Equal(invoice.DueDate, parsed.Invoice.DueDate);
        if (scenario == "signed-order") {
            var document = System.Xml.Linq.XDocument.Parse(System.Text.Encoding.UTF8.GetString(xml));
            Assert.Equal(new[] { "-0.12", "-0.03", "-0.02" }, document.Descendants().Where(element => element.Name.LocalName == "P_11VatZ").Select(element => element.Value));
        }
    }
    [Theory]
    [InlineData(Fa3InvoiceKind.TaxInvoice)]
    [InlineData(Fa3InvoiceKind.Correction)]
    [InlineData(Fa3InvoiceKind.Advance)]
    [InlineData(Fa3InvoiceKind.Settlement)]
    [InlineData(Fa3InvoiceKind.Simplified)]
    [InlineData(Fa3InvoiceKind.AdvanceCorrection)]
    [InlineData(Fa3InvoiceKind.SettlementCorrection)]
    public void EachNationalKindHasAnExplicitSchemaQualifiedNativeCreationPath(Fa3InvoiceKind kind) {
        Invoice invoice = Fa3InvoiceFixture.Create(); Fa3InvoiceWriteOptions options = Fa3InvoiceFixture.Options(kind);
        bool correction = kind is Fa3InvoiceKind.Correction or Fa3InvoiceKind.AdvanceCorrection or Fa3InvoiceKind.SettlementCorrection;
        if (kind != Fa3InvoiceKind.TaxInvoice) {
            options.FiscalAmounts = new Fa3FiscalAmounts(correction ? -12.30m : 61.50m, new[] { new Fa3TaxSummary("1", correction ? -10m : 50m, correction ? -2.30m : 11.50m) });
        }
        if (correction) {
            invoice.TypeCode = "384"; invoice.PrecedingInvoices.Add(new InvoiceReference("original-FA", new DateTime(2026, 9, 10)));
            options.CorrectionReason = "Explicit price correction"; options.CorrectionTimingCode = 3;
            invoice.Lines[0].Quantity = -0.1m;
        }
        if (kind is Fa3InvoiceKind.Advance or Fa3InvoiceKind.AdvanceCorrection) {
            options.Order = new Fa3Order(123, Fa3InvoiceFixture.Create().Lines);
            if (kind == Fa3InvoiceKind.Advance) { invoice.TypeCode = "386"; invoice.Lines.Clear(); }
            else {
                options.PreviousAdvanceOrSettlementTotal = 61.50m;
                var differenceRows = Fa3InvoiceFixture.Create().Lines; differenceRows[0].Quantity = -0.1m;
                options.Order = Fa3Order.CorrectionDifferences(123m, 110.70m, differenceRows, new decimal?[] { -2.30m });
            }
        }
        if (kind is Fa3InvoiceKind.Settlement or Fa3InvoiceKind.SettlementCorrection)
            options.AdvanceInvoiceReferences.Add(new Fa3AdvanceInvoiceReference("advance-FA"));
        if (kind == Fa3InvoiceKind.Simplified) { invoice.Buyer.Name = string.Empty; invoice.Buyer.Address = new InvoiceAddress(); invoice.Lines.Clear(); }
        byte[] xml = Fa3InvoiceWriter.Write(invoice, options);
        var validator = new Fa3SchemaValidator(Fa3SchemaBundle.LoadDirectory(Path.Combine(AppContext.BaseDirectory, "Fixtures", "FA3", "Schemas")));
        Fa3SchemaValidationResult validation = validator.Validate(xml);
        Assert.True(validation.IsValid, string.Join(Environment.NewLine, validation.Diagnostics.Select(item => item.Message)));
        Fa3InvoiceReadResult parsed = Fa3InvoiceReader.Read(xml);
        Assert.Equal(kind, parsed.Kind); Assert.Equal(invoice.Number, parsed.Invoice.Number);
        Assert.Equal(kind == Fa3InvoiceKind.TaxInvoice ? 123m : options.FiscalAmounts!.Total, parsed.DeclaredTotal);
        if (kind == Fa3InvoiceKind.AdvanceCorrection) {
            var document = System.Xml.Linq.XDocument.Parse(System.Text.Encoding.UTF8.GetString(xml));
            Assert.Equal("110.7", document.Descendants().Single(element => element.Name.LocalName == "WartoscZamowienia").Value);
            Assert.Equal("-10", document.Descendants().Single(element => element.Name.LocalName == "P_11NettoZ").Value);
            options.Order = new Fa3Order(123, Fa3InvoiceFixture.Create().Lines);
            Assert.Contains(Fa3InvoiceWriter.Inspect(invoice, options), diagnostic => diagnostic.Location == "Order");
            options.Order = Fa3Order.CorrectionDifferences(123, 123, new[] { new InvoiceLine { Id = "1", Name = "Unchanged", UnitCode = "HUR", Quantity = 0, UnitPrice = 100, Tax = new InvoiceTaxCategory { Code = "S", Rate = 23 } } }, new decimal?[] { 0 });
            Assert.Throws<InvalidDataException>(() => Fa3InvoiceWriter.Write(invoice, options));
            options.Order = Fa3Order.CorrectionDifferences(123, 100, Fa3InvoiceFixture.Create().Lines, new decimal?[] { -2.30m });
            Assert.Contains(Fa3InvoiceWriter.Inspect(invoice, options), diagnostic => diagnostic.Location == "Order.Total");
        }
        if (correction) {
            invoice.PrecedingInvoices.Clear(); invoice.PrecedingInvoices.Add(new InvoiceReference("original-FA", new DateTime(2005, 1, 1)));
            Assert.Contains(Fa3InvoiceWriter.Inspect(invoice, options), diagnostic => diagnostic.Location == "PrecedingInvoices.IssueDate");
            var document = System.Xml.Linq.XDocument.Parse(System.Text.Encoding.UTF8.GetString(xml));
            document.Descendants().Single(element => element.Name.LocalName == "DataWystFaKorygowanej").Value = "2005-01-01";
            Assert.False(validator.Validate(System.Text.Encoding.UTF8.GetBytes(document.ToString())).IsValid);
        }
    }
}
