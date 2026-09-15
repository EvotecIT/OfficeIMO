using System.Reflection;
using Xunit;

namespace OfficeIMO.Invoicing.Tests;

public class InvoiceSpecificationContractTests {
    [Theory]
    [InlineData(InvoiceProfile.Minimum)]
    [InlineData(InvoiceProfile.BasicWithoutLines)]
    [InlineData(InvoiceProfile.Basic)]
    [InlineData(InvoiceProfile.En16931)]
    [InlineData(InvoiceProfile.Extended)]
    public void FacturXReleaseAcceptsItsFiveCiiProfiles(InvoiceProfile profile) {
        var options = new InvoiceXmlOptions(
            InvoiceSpecificationRelease.FacturX_1_09_2_Zugferd_2_5_2,
            InvoiceSyntax.Cii,
            profile);
        Assert.Equal(profile, options.Profile);
    }

    [Theory]
    [InlineData(InvoiceProfile.Minimum, 0)]
    [InlineData(InvoiceProfile.BasicWithoutLines, 0)]
    [InlineData(InvoiceProfile.Basic, 1)]
    public void ReducedFacturXOutputCanBeReadAsItsRetainedModel(InvoiceProfile profile, int expectedLines) {
        var options = new InvoiceXmlOptions(
            InvoiceSpecificationRelease.FacturX_1_09_2_Zugferd_2_5_2,
            InvoiceSyntax.Cii,
            profile,
            InvoiceProjectionPolicy.AllowProfileDefinedDataLoss);

        InvoiceReadResult result = InvoiceParser.Read(InvoiceSerializer.Write(CreateProjectionSafeInvoice(profile), options));

        Assert.True(result.HasCompleteMapping);
        Assert.Equal(profile, result.Declaration.Profile);
        Assert.Equal(expectedLines, result.Invoice.Lines.Count);
        Assert.Equal(119m, result.Invoice.DeclaredTotals!.PayableAmount);
        byte[] rewritten = result.Write(options);
        Assert.Equal(profile, InvoiceProfileDeclaration.Read(rewritten).Profile);
        Assert.Equal(expectedLines, InvoiceParser.Read(rewritten).Invoice.Lines.Count);
    }

    [Theory]
    [InlineData(InvoiceProfile.Minimum)]
    [InlineData(InvoiceProfile.BasicWithoutLines)]
    [InlineData(InvoiceProfile.Basic)]
    [InlineData(InvoiceProfile.En16931)]
    [InlineData(InvoiceProfile.Extended)]
    public void FacturXProfilesDoNotInventABusinessProcessIdentifier(InvoiceProfile profile) {
        Invoice invoice = CreateProjectionSafeInvoice(profile);
        invoice.BusinessProcessId = null;
        var options = new InvoiceXmlOptions(
            InvoiceSpecificationRelease.FacturX_1_09_2_Zugferd_2_5_2,
            InvoiceSyntax.Cii,
            profile,
            profile is InvoiceProfile.Minimum or InvoiceProfile.BasicWithoutLines or InvoiceProfile.Basic
                ? InvoiceProjectionPolicy.AllowProfileDefinedDataLoss
                : InvoiceProjectionPolicy.RejectDataLoss);

        byte[] xml = InvoiceSerializer.Write(invoice, options);

        Assert.DoesNotContain("BusinessProcessSpecifiedDocumentContextParameter", System.Text.Encoding.UTF8.GetString(xml));
    }

    [Theory]
    [InlineData(InvoiceSyntax.Cii, InvoiceProfile.XRechnung)]
    [InlineData(InvoiceSyntax.Ubl, InvoiceProfile.XRechnung)]
    [InlineData(InvoiceSyntax.Ubl, InvoiceProfile.PeppolBis)]
    public void CiusAndNetworkContractsStillRequireABusinessProcessIdentifier(InvoiceSyntax syntax, InvoiceProfile profile) {
        Invoice invoice = InvoiceFixture.Create();
        invoice.BusinessProcessId = null;
        InvoiceXmlOptions options = InvoiceTestContracts.For(syntax, profile);

        Assert.Contains(InvoiceSerializer.InspectTarget(invoice, options),
            item => item.Location == "BusinessProcessId" && item.Severity == InvoiceDiagnosticSeverity.Error);
        Assert.Throws<InvalidDataException>(() => InvoiceSerializer.Write(invoice, options));
    }

    [Theory]
    [InlineData(InvoiceProfile.Basic)]
    [InlineData(InvoiceProfile.En16931)]
    [InlineData(InvoiceProfile.Extended)]
    public void AggregateOnlyModelCannotBePromotedToALineBearingFacturXProfile(InvoiceProfile profile) {
        Invoice invoice = InvoiceFixture.Create();
        invoice.Lines.Clear();
        InvoiceXmlOptions options = InvoiceTestContracts.FacturX(profile);

        Assert.Contains(InvoiceSerializer.InspectTarget(invoice, options), item => item.Code == "INV-LINES");
        Assert.Throws<InvalidDataException>(() => InvoiceSerializer.Write(invoice, options));
    }

    [Fact]
    public void LowerProfileLineProjectionStatesThatDescriptionsAndTaxClassificationsAreOmitted() {
        Invoice invoice = InvoiceFixture.Create();
        IReadOnlyList<InvoiceDiagnostic> diagnostics = InvoiceSerializer.InspectTarget(invoice,
            InvoiceTestContracts.FacturX(InvoiceProfile.Minimum, InvoiceProjectionPolicy.AllowProfileDefinedDataLoss));

        InvoiceDiagnostic line = Assert.Single(diagnostics, item => item.Location == "Lines");
        Assert.Contains("descriptions and tax classifications", line.Message, StringComparison.Ordinal);
        Assert.DoesNotContain("visible invoice", line.Message, StringComparison.Ordinal);
    }

    [Fact]
    public void AggregateOnlyRewriteRejectsInconsistentDeclaredTotals() {
        var options = InvoiceTestContracts.FacturX(
            InvoiceProfile.BasicWithoutLines,
            InvoiceProjectionPolicy.AllowProfileDefinedDataLoss);
        InvoiceReadResult result = InvoiceParser.Read(InvoiceSerializer.Write(InvoiceFixture.Create(), options));
        result.Invoice.DeclaredTotals!.PayableAmount += 1m;

        IReadOnlyList<InvoiceDiagnostic> diagnostics = InvoiceSerializer.InspectTarget(result.Invoice, options);

        Assert.Contains(diagnostics, item => item.Code == "INV-DECLARED-AMOUNT" && item.Location == "DeclaredTotals.PayableAmount");
        Assert.Throws<InvalidDataException>(() => result.Write(options));
    }

    [Theory]
    [InlineData(InvoiceProfile.Minimum, "PrepaidAmount")]
    [InlineData(InvoiceProfile.Minimum, "RoundingAmount")]
    [InlineData(InvoiceProfile.BasicWithoutLines, "RoundingAmount")]
    [InlineData(InvoiceProfile.Basic, "RoundingAmount")]
    [InlineData(InvoiceProfile.Basic, "PriceBaseQuantity")]
    [InlineData(InvoiceProfile.Basic, "LineAllowancesAndCharges")]
    [InlineData(InvoiceProfile.Basic, "DocumentAllowancesAndCharges")]
    public void ArithmeticChangingLowerProfileProjectionRemainsBlocked(InvoiceProfile profile, string field) {
        Invoice invoice = CreateProjectionSafeInvoice(profile);
        string location = field;
        switch (field) {
            case "PrepaidAmount": invoice.PrepaidAmount = 1m; break;
            case "RoundingAmount": invoice.RoundingAmount = 0.01m; break;
            case "PriceBaseQuantity": invoice.Lines[0].PriceBaseQuantity = 2m; location = "Lines[0].PriceBaseQuantity"; break;
            case "LineAllowancesAndCharges":
                invoice.Lines[0].AllowancesAndCharges.Add(new InvoiceAllowanceCharge { Amount = 1m, Reason = "Discount" });
                location = "Lines[0].AllowancesAndCharges";
                break;
            case "DocumentAllowancesAndCharges":
                invoice.AllowancesAndCharges.Add(new InvoiceAllowanceCharge {
                    Amount = 1m, Reason = "Discount", Tax = new InvoiceTaxCategory { Code = "S", Rate = 19m }
                });
                location = "AllowancesAndCharges";
                break;
        }
        InvoiceCalculator.UpdateDeclaredAmounts(invoice);
        InvoiceXmlOptions options = InvoiceTestContracts.FacturX(profile, InvoiceProjectionPolicy.AllowProfileDefinedDataLoss);

        InvoiceDiagnostic diagnostic = Assert.Single(InvoiceSerializer.InspectTarget(invoice, options),
            item => item.Code == "INV-TARGET-PROJECTION" && item.Location == location);
        Assert.Equal(InvoiceDiagnosticSeverity.Error, diagnostic.Severity);
        Assert.Throws<InvalidDataException>(() => InvoiceSerializer.Write(invoice, options));
    }

    [Fact]
    public void AggregateOnlyModelRejectsDuplicateVatCategoryAndRate() {
        InvoiceXmlOptions options = InvoiceTestContracts.FacturX(
            InvoiceProfile.BasicWithoutLines,
            InvoiceProjectionPolicy.AllowProfileDefinedDataLoss);
        Invoice invoice = InvoiceParser.Read(InvoiceSerializer.Write(CreateProjectionSafeInvoice(InvoiceProfile.BasicWithoutLines), options)).Invoice;
        InvoiceDeclaredTax source = Assert.Single(invoice.DeclaredTaxes);
        invoice.DeclaredTaxes.Add(new InvoiceDeclaredTax {
            Category = new InvoiceTaxCategory { Code = source.Category.Code, Rate = source.Category.Rate },
            TaxableAmount = 0m,
            TaxAmount = 0m
        });

        Assert.Contains(InvoiceSerializer.InspectTarget(invoice, options),
            item => item.Code == "INV-TAX-BREAKDOWN" && item.Location == "DeclaredTaxes[1].Category");
        Assert.Throws<InvalidDataException>(() => InvoiceSerializer.Write(invoice, options));
    }

    [Fact]
    public void AggregateOnlyModelRejectsVatAmountsThatDoNotMatchTheirRate() {
        InvoiceXmlOptions options = InvoiceTestContracts.FacturX(
            InvoiceProfile.BasicWithoutLines,
            InvoiceProjectionPolicy.AllowProfileDefinedDataLoss);
        Invoice invoice = InvoiceParser.Read(InvoiceSerializer.Write(
            CreateProjectionSafeInvoice(InvoiceProfile.BasicWithoutLines), options)).Invoice;
        InvoiceDeclaredTax tax = Assert.Single(invoice.DeclaredTaxes);
        tax.TaxAmount = 1m;
        invoice.DeclaredTotals!.TaxTotal = 1m;
        invoice.DeclaredTotals.TaxInclusiveTotal = invoice.DeclaredTotals.TaxExclusiveTotal + 1m;
        invoice.DeclaredTotals.PayableAmount = invoice.DeclaredTotals.TaxInclusiveTotal;

        Assert.Contains(InvoiceSerializer.InspectTarget(invoice, options),
            item => item.Code == "INV-DECLARED-AMOUNT" && item.Location == "DeclaredTaxes[0].TaxAmount");
        Assert.Throws<InvalidDataException>(() => InvoiceSerializer.Write(invoice, options));
    }

    [Theory]
    [InlineData(InvoiceProfile.Minimum)]
    [InlineData(InvoiceProfile.BasicWithoutLines)]
    [InlineData(InvoiceProfile.Basic)]
    public void LowerProfileProjectionCannotRemoveTheSellersOnlyQualifyingIdentity(InvoiceProfile profile) {
        Invoice invoice = CreateProjectionSafeInvoice(profile);
        invoice.Seller.TaxRegistrations.Clear();
        invoice.Seller.TaxRegistrations.Add(new InvoiceTaxRegistration("local-tax-id", InvoiceTaxRegistration.TaxScheme));
        invoice.Seller.LegalRegistration = null;
        invoice.Seller.Identifiers.Add(new InvoiceIdentifier("seller-business-id"));
        InvoiceXmlOptions options = InvoiceTestContracts.FacturX(profile, InvoiceProjectionPolicy.AllowProfileDefinedDataLoss);

        Assert.True(InvoiceModelValidator.Validate(invoice).IsValid);
        InvoiceDiagnostic diagnostic = Assert.Single(InvoiceSerializer.InspectTarget(invoice, options),
            item => item.Location == "Seller.Identifiers" && item.Severity == InvoiceDiagnosticSeverity.Error);
        Assert.Contains("only qualifying business identity", diagnostic.Message, StringComparison.Ordinal);
        Assert.Throws<InvalidDataException>(() => InvoiceSerializer.Write(invoice, options));
    }

    [Theory]
    [InlineData(InvoiceProfile.BasicWithoutLines)]
    [InlineData(InvoiceProfile.Basic)]
    public void LowerProfileProjectionCannotRemoveDeliveryEvidenceRequiredByCategoryK(InvoiceProfile profile) {
        Invoice invoice = InvoiceFixture.WithTaxCategory("K");
        InvoiceXmlOptions options = InvoiceTestContracts.FacturX(profile, InvoiceProjectionPolicy.AllowProfileDefinedDataLoss);

        Assert.True(InvoiceModelValidator.Validate(invoice).IsValid);
        InvoiceDiagnostic diagnostic = Assert.Single(InvoiceSerializer.InspectTarget(invoice, options),
            item => item.Location == "Delivery" && item.Severity == InvoiceDiagnosticSeverity.Error);
        Assert.Contains("VAT category K", diagnostic.Message, StringComparison.Ordinal);
        Assert.Throws<InvalidDataException>(() => InvoiceSerializer.Write(invoice, options));
    }

    [Theory]
    [InlineData(InvoiceProfile.BasicWithoutLines, "S")]
    [InlineData(InvoiceProfile.BasicWithoutLines, "G")]
    [InlineData(InvoiceProfile.Basic, "S")]
    [InlineData(InvoiceProfile.Basic, "G")]
    public void LowerProfileProjectionCannotRemoveARequiredTaxRepresentative(InvoiceProfile profile, string categoryCode) {
        Invoice invoice = CreateProjectionSafeInvoice(profile);
        invoice.Seller.TaxRegistrations.Clear();
        invoice.Seller.LegalRegistration = new InvoiceIdentifier("seller-register");
        if (categoryCode == "G")
            invoice.Seller.TaxRegistrations.Add(new InvoiceTaxRegistration("seller-tax", InvoiceTaxRegistration.TaxScheme));
        invoice.Lines[0].Tax = new InvoiceTaxCategory {
            Code = categoryCode,
            Rate = categoryCode == "G" ? 0m : 19m,
            ExemptionReason = categoryCode == "G" ? "Export outside the EU" : null
        };
        invoice.TaxRepresentative = new InvoiceParty {
            Name = "Tax representative",
            Address = new InvoiceAddress { CountryCode = "DE" }
        };
        invoice.TaxRepresentative.TaxRegistrations.Add(
            new InvoiceTaxRegistration("DE999999999", InvoiceTaxRegistration.VatScheme));
        InvoiceCalculator.UpdateDeclaredAmounts(invoice);
        InvoiceXmlOptions options = InvoiceTestContracts.FacturX(profile, InvoiceProjectionPolicy.AllowProfileDefinedDataLoss);

        Assert.True(InvoiceModelValidator.Validate(invoice).IsValid);
        InvoiceDiagnostic diagnostic = Assert.Single(InvoiceSerializer.InspectTarget(invoice, options),
            item => item.Location == "TaxRepresentative" && item.Severity == InvoiceDiagnosticSeverity.Error);
        Assert.Contains("required by a retained VAT category", diagnostic.Message, StringComparison.Ordinal);
        Assert.Throws<InvalidDataException>(() => InvoiceSerializer.Write(invoice, options));
    }

    [Theory]
    [InlineData(InvoiceSpecificationRelease.FacturX_1_09_2_Zugferd_2_5_2, InvoiceSyntax.Ubl, InvoiceProfile.En16931)]
    [InlineData(InvoiceSpecificationRelease.En16931_1_3_16, InvoiceSyntax.Cii, InvoiceProfile.XRechnung)]
    [InlineData(InvoiceSpecificationRelease.XRechnung_3_0_2_2026_08_31, InvoiceSyntax.Cii, InvoiceProfile.ExtendedCtcFr)]
    [InlineData(InvoiceSpecificationRelease.PeppolBis_3_0_21, InvoiceSyntax.Cii, InvoiceProfile.PeppolBis)]
    [InlineData(InvoiceSpecificationRelease.PeppolBis_3_0_21, InvoiceSyntax.Ubl, InvoiceProfile.En16931)]
    public void ReleaseSyntaxAndProfileCannotBeMixed(
        InvoiceSpecificationRelease release,
        InvoiceSyntax syntax,
        InvoiceProfile profile) =>
        Assert.Throws<NotSupportedException>(() => new InvoiceXmlOptions(release, syntax, profile));

    [Fact]
    public void XRechnungIsTheOnlyNationalCiusWithAnAuthoringContract() {
        foreach (InvoiceSyntax syntax in new[] { InvoiceSyntax.Cii, InvoiceSyntax.Ubl }) {
            var contract = new InvoiceXmlOptions(
                InvoiceSpecificationRelease.XRechnung_3_0_2_2026_08_31,
                syntax,
                InvoiceProfile.XRechnung);
            Assert.Equal(syntax, contract.Syntax);
        }
        foreach (InvoiceSpecificationRelease release in Enum.GetValues(typeof(InvoiceSpecificationRelease)))
            Assert.Throws<NotSupportedException>(() => new InvoiceXmlOptions(release, InvoiceSyntax.Cii, InvoiceProfile.ExtendedCtcFr));
    }

    [Fact]
    public void XmlAuthoringHasNoImplicitReleaseOverload() {
        MethodInfo[] writes = typeof(InvoiceSerializer).GetMethods(BindingFlags.Public | BindingFlags.Static)
            .Where(method => method.Name == nameof(InvoiceSerializer.Write)).ToArray();
        MethodInfo write = Assert.Single(writes);
        Assert.Equal(new[] { typeof(Invoice), typeof(InvoiceXmlOptions) }, write.GetParameters().Select(parameter => parameter.ParameterType));
    }

    private static Invoice CreateProjectionSafeInvoice(InvoiceProfile profile) {
        Invoice invoice = InvoiceFixture.Create();
        invoice.RoundingAmount = 0m;
        if (profile == InvoiceProfile.Minimum) invoice.PrepaidAmount = 0m;
        if (profile == InvoiceProfile.Basic) {
            invoice.AllowancesAndCharges.Clear();
            foreach (InvoiceLine line in invoice.Lines) {
                line.PriceBaseQuantity = 1m;
                line.AllowancesAndCharges.Clear();
            }
        }
        InvoiceCalculator.UpdateDeclaredAmounts(invoice);
        return invoice;
    }
}
