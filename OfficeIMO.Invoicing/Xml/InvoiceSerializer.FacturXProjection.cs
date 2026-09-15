namespace OfficeIMO.Invoicing;

public static partial class InvoiceSerializer {
    private static bool IsReducedFacturX(InvoiceXmlOptions options) =>
        options.Release == InvoiceSpecificationRelease.FacturX_1_09_2_Zugferd_2_5_2 &&
        options.Profile is InvoiceProfile.Minimum or InvoiceProfile.BasicWithoutLines or InvoiceProfile.Basic;

    private static void CheckFacturXProjection(Invoice invoice, InvoiceCalculation calculation, InvoiceXmlOptions options, Action<string, string> projection,
        Action<string, string> requiredProjection) {
        if (!IsReducedFacturX(options)) return;

        string profile = options.Profile == InvoiceProfile.BasicWithoutLines ? "BASIC WL" : options.Profile.ToString().ToUpperInvariant();
        void Omitted(bool populated, string path, string meaning) {
            if (populated) projection(path, "Factur-X " + profile + " does not carry " + meaning + " in this authoring projection.");
        }

        if (options.Profile is (InvoiceProfile.Minimum or InvoiceProfile.BasicWithoutLines) && invoice.Lines.Count != 0)
            projection("Lines", "Factur-X " + profile + " carries calculated invoice totals without XML invoice-line occurrences; " + invoice.Lines.Count + " source line(s), including their descriptions and tax classifications, are omitted while calculated aggregate amounts are retained.");

        Omitted(invoice.BuyerReference != null, "BuyerReference", "the buyer routing reference");
        bool sellerVat = invoice.Seller.TaxRegistrations.Any(registration => registration != null &&
            registration.Kind == InvoiceTaxRegistrationKind.Vat && !string.IsNullOrWhiteSpace(registration.Identifier));
        bool sellerTax = invoice.Seller.TaxRegistrations.Any(registration => registration != null &&
            registration.Kind is InvoiceTaxRegistrationKind.Vat or InvoiceTaxRegistrationKind.Fiscal &&
            !string.IsNullOrWhiteSpace(registration.Identifier));
        bool retainsTaxCategories = options.Profile is InvoiceProfile.BasicWithoutLines or InvoiceProfile.Basic;
        bool requiresSellerTax = retainsTaxCategories && calculation.Taxes.Any(tax => tax.CategoryCode != "O");
        bool requiresSellerVat = retainsTaxCategories && calculation.Taxes.Any(tax => tax.CategoryCode is "G" or "K");
        bool representativeRequired = invoice.TaxRepresentative != null && options.Profile is InvoiceProfile.BasicWithoutLines or InvoiceProfile.Basic &&
            (requiresSellerVat ? !sellerVat : requiresSellerTax && !sellerTax);
        if (representativeRequired)
            requiredProjection("TaxRepresentative", "The lower Factur-X projection cannot omit the tax representative whose VAT identifier is required by a retained VAT category.");
        else
            Omitted(invoice.TaxRepresentative != null, "TaxRepresentative", "the seller tax representative");
        Omitted(invoice.Payee != null, "Payee", "a distinct payee");
        bool retainsIntraCommunityTax = options.Profile is InvoiceProfile.BasicWithoutLines or InvoiceProfile.Basic &&
            (invoice.Lines.Any(line => line?.Tax?.Code == "K") ||
             invoice.AllowancesAndCharges.Any(item => item?.Tax?.Code == "K") ||
             invoice.DeclaredTaxes.Any(tax => tax?.Category?.Code == "K"));
        if (retainsIntraCommunityTax && invoice.Delivery != null)
            requiredProjection("Delivery", "The selected lower Factur-X profile cannot omit delivery evidence required by retained VAT category K.");
        else
            Omitted(invoice.Delivery != null, "Delivery", "delivery party, location, address, or date data");
        Omitted(invoice.ProjectReference != null, "ProjectReference", "the project reference");
        Omitted(invoice.ContractReference != null, "ContractReference", "the contract reference");
        Omitted(invoice.PurchaseOrderReference != null, "PurchaseOrderReference", "the purchase-order reference");
        Omitted(invoice.SalesOrderReference != null, "SalesOrderReference", "the sales-order reference");
        Omitted(invoice.ReceivingAdviceReference != null, "ReceivingAdviceReference", "the receiving-advice reference");
        Omitted(invoice.DespatchAdviceReference != null, "DespatchAdviceReference", "the despatch-advice reference");
        Omitted(invoice.TenderReference != null, "TenderReference", "the tender reference");
        Omitted(invoice.AccountingReference != null, "AccountingReference", "the accounting reference");
        Omitted(invoice.ObjectIdentifier != null, "ObjectIdentifier", "the invoiced-object identifier");
        Omitted(invoice.SupportingDocuments.Count != 0, "SupportingDocuments", invoice.SupportingDocuments.Count + " supporting document(s)");
        Omitted(invoice.PrecedingInvoices.Count != 0, "PrecedingInvoices", invoice.PrecedingInvoices.Count + " preceding invoice reference(s)");
        Omitted(invoice.Payments.Count != 0, "Payments", invoice.Payments.Count + " payment instruction occurrence(s)");
        Omitted(invoice.PaymentTerms != null, "PaymentTerms", "payment terms text");
        Omitted(invoice.DueDate.HasValue, "DueDate", "the payment due date");
        if (retainsIntraCommunityTax && invoice.Period != null)
            requiredProjection("Period", "The selected lower Factur-X profile cannot omit invoicing-period evidence required by retained VAT category K.");
        else
            Omitted(invoice.Period != null, "Period", "the invoicing period");
        Omitted(invoice.TaxPointDate.HasValue || invoice.TaxPointDateCode != null, "TaxPoint", "the tax point date or code");
        Omitted(invoice.TaxCurrency != null || invoice.TaxAmountInAccountingCurrency.HasValue, "TaxCurrency", "accounting-currency VAT data");
        if (options.Profile == InvoiceProfile.Basic && invoice.AllowancesAndCharges.Count != 0)
            requiredProjection("AllowancesAndCharges", "Factur-X BASIC cannot omit document allowances or charges without changing the reconstructed invoice arithmetic.");
        else
            Omitted(invoice.AllowancesAndCharges.Count != 0, "AllowancesAndCharges", invoice.AllowancesAndCharges.Count + " document allowance or charge occurrence(s)");
        if (invoice.RoundingAmount != 0m)
            requiredProjection("RoundingAmount", "The selected lower Factur-X profile cannot carry the payable rounding adjustment without changing reconstructed totals.");
        if (options.Profile == InvoiceProfile.Minimum) {
            if (invoice.PrepaidAmount != 0m)
                requiredProjection("PrepaidAmount", "Factur-X MINIMUM cannot carry the prepaid amount without changing the reconstructed payable amount.");
            Omitted(invoice.DeclaredTaxes.Count != 0, "DeclaredTaxes", invoice.DeclaredTaxes.Count + " VAT breakdown occurrence(s)");
            Omitted(invoice.DeclaredTotals?.LineNetTotal.HasValue == true, "DeclaredTotals.LineNetTotal", "the declared line-net total");
            Omitted(invoice.DeclaredTotals?.AllowanceTotal.HasValue == true, "DeclaredTotals.AllowanceTotal", "the declared allowance total");
            Omitted(invoice.DeclaredTotals?.ChargeTotal.HasValue == true, "DeclaredTotals.ChargeTotal", "the declared charge total");
        }

        if (!representativeRequired && invoice.TaxRepresentative == null && requiresSellerTax && !sellerTax)
            requiredProjection("Seller.TaxRegistrations", "The lower Factur-X projection cannot omit the seller's only qualifying tax registration while retaining taxable VAT categories.");

        bool sellerIdentifiersRequired = invoice.Seller.Identifiers.Any(identifier => !string.IsNullOrWhiteSpace(identifier?.Value)) &&
            string.IsNullOrWhiteSpace(invoice.Seller.LegalRegistration?.Value) &&
            !invoice.Seller.TaxRegistrations.Any(registration => registration != null &&
                registration.Kind == InvoiceTaxRegistrationKind.Vat && !string.IsNullOrWhiteSpace(registration.Identifier));
        CheckReducedParty(invoice.Seller, "Seller", options.Profile == InvoiceProfile.Minimum, retainMinimumCountry: true,
            sellerIdentifiersRequired, projection, requiredProjection);
        CheckReducedParty(invoice.Buyer, "Buyer", options.Profile == InvoiceProfile.Minimum, retainMinimumCountry: false,
            requiredIdentifiers: false, projection, requiredProjection);

        if (options.Profile == InvoiceProfile.Basic) {
            for (int index = 0; index < invoice.Lines.Count; index++) {
                InvoiceLine line = invoice.Lines[index];
                string path = "Lines[" + index + "]";
                Omitted(line.Description != null, path + ".Description", "the line description");
                Omitted(line.Note != null, path + ".Note", "the line note");
                Omitted(line.OrderLineReference != null, path + ".OrderLineReference", "the order-line reference");
                Omitted(line.AccountingReference != null, path + ".AccountingReference", "the line accounting reference");
                Omitted(line.StandardItemIdentifier != null || line.SellerItemIdentifier != null || line.BuyerItemIdentifier != null,
                    path + ".ItemIdentifiers", "line item identifiers");
                Omitted(line.ObjectIdentifier != null, path + ".ObjectIdentifier", "the line invoiced-object identifier");
                Omitted(line.GrossPrice.HasValue || line.PriceDiscount.HasValue,
                    path + ".Price", "gross price or discount data");
                if (line.PriceBaseQuantity != 1m)
                    requiredProjection(path + ".PriceBaseQuantity", "Factur-X BASIC cannot omit a non-default price base quantity without changing reconstructed line arithmetic.");
                Omitted(line.Period != null, path + ".Period", "the line invoicing period");
                if (line.AllowancesAndCharges.Count != 0)
                    requiredProjection(path + ".AllowancesAndCharges", "Factur-X BASIC cannot omit line allowances or charges without changing reconstructed line arithmetic.");
                Omitted(line.Classifications.Count != 0, path + ".Classifications", "item classifications");
                Omitted(line.Attributes.Count != 0, path + ".Attributes", "item attributes");
                Omitted(line.OriginCountryCode != null, path + ".OriginCountryCode", "the item origin country");
            }
        }
    }

    private static void CheckReducedParty(InvoiceParty party, string path, bool minimum, bool retainMinimumCountry,
        bool requiredIdentifiers, Action<string, string> projection, Action<string, string> requiredProjection) {
        void Omitted(bool populated, string field) {
            if (populated) projection(path + "." + field, "The selected lower Factur-X profile authoring projection does not carry " + field + " for this party.");
        }
        Omitted(party.TradingName != null, "TradingName");
        Omitted(party.LegalInformation != null, "LegalInformation");
        if (requiredIdentifiers)
            requiredProjection(path + ".Identifiers", "The lower Factur-X projection cannot omit the seller's only qualifying business identity.");
        else
            Omitted(party.Identifiers.Count != 0, "Identifiers");
        Omitted(party.ElectronicAddress != null, "ElectronicAddress");
        Omitted(party.Contact != null, "Contact");
        if (minimum) {
            InvoiceAddress address = party.Address;
            Omitted((!retainMinimumCountry && !string.IsNullOrEmpty(address.CountryCode)) || address.PostCode != null || address.Line1 != null || address.Line2 != null ||
                address.Line3 != null || address.City != null || address.Subdivision != null, "Address");
        }
        for (int index = 0; index < party.TaxRegistrations.Count; index++) {
            InvoiceTaxRegistration registration = party.TaxRegistrations[index];
            if (registration.Kind == InvoiceTaxRegistrationKind.Other)
                projection(path + ".TaxRegistrations[" + index + "]", "The lower Factur-X projection cannot carry tax-registration scheme '" + registration.SchemeId + "'.");
        }
    }
}
