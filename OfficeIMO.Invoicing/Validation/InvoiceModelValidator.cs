using System.Text.RegularExpressions;
using System.Xml;

namespace OfficeIMO.Invoicing;

/// <summary>Checks semantic completeness, arithmetic and source-declared amounts before XML emission.</summary>
public static partial class InvoiceModelValidator {
    /// <summary>Validates the model; use pinned external rules separately for authoritative profile compliance.</summary>
    public static InvoiceModelValidationResult Validate(Invoice invoice) {
        if (invoice == null) throw new ArgumentNullException(nameof(invoice));
        var diagnostics = new List<InvoiceDiagnostic>();
        var check = new ModelChecks(diagnostics);
        try { new InvoiceModelLimits().Check(invoice); }
        catch (InvalidDataException exception) {
            check.Error("INV-MODEL-LIMIT", exception.Message, "Invoice");
            return new InvoiceModelValidationResult(diagnostics, null);
        }
        catch (XmlException) {
            check.Error("INV-MODEL-XML", "Invoice text contains a character that XML cannot represent.", "Invoice");
            return new InvoiceModelValidationResult(diagnostics, null);
        }
        check.Required(invoice.Number, "Number");
        check.Date(invoice.IssueDate, "IssueDate");
        check.Code(invoice.TypeCode, "TypeCode", "^[0-9]{3}\\z");
        check.Code(invoice.Currency, "Currency", "^[A-Z]{3}\\z");
        check.Party(invoice.Seller, "Seller");
        check.Party(invoice.Buyer, "Buyer");
        if (invoice.Payee != null) { check.Required(invoice.Payee.Name, "Payee.Name"); check.PartyIdentifiers(invoice.Payee, "Payee"); }
        if (invoice.ObjectIdentifier != null) check.Identifier(invoice.ObjectIdentifier, "ObjectIdentifier", false);
        if (invoice.TaxRepresentative != null) {
            check.Party(invoice.TaxRepresentative, "TaxRepresentative");
            check.Required(invoice.TaxRepresentative.VatIdentifier, "TaxRepresentative.VatIdentifier");
        }
        check.Period(invoice.Period, "Period");
        check.OptionalDate(invoice.DueDate, "DueDate");
        check.OptionalDate(invoice.TaxPointDate, "TaxPointDate");
        if (invoice.TaxPointDate.HasValue && invoice.TaxPointDateCode != null)
            check.Error("INV-TAX-POINT", "Choose the tax point date or its code, not both.", "TaxPointDate");
        if (invoice.Delivery != null) {
            check.OptionalDate(invoice.Delivery.Date, "Delivery.Date");
            if (invoice.Delivery.Address != null) check.Address(invoice.Delivery.Address, "Delivery.Address");
            if (invoice.Delivery.LocationIdentifier != null) check.Identifier(invoice.Delivery.LocationIdentifier, "Delivery.LocationIdentifier", false);
        }
        check.Money(invoice.PrepaidAmount, "PrepaidAmount");
        check.Money(invoice.RoundingAmount, "RoundingAmount");
        if (invoice.PrepaidAmount < 0m) check.Error("INV-PREPAID", "Prepaid amount cannot be negative.", "PrepaidAmount");
        if ((invoice.TaxCurrency != null) != invoice.TaxAmountInAccountingCurrency.HasValue)
            check.Error("INV-TAX-CURRENCY", "Supply both the accounting currency and its VAT amount.", "TaxCurrency");
        if (invoice.TaxCurrency != null) {
            check.Code(invoice.TaxCurrency, "TaxCurrency", "^[A-Z]{3}\\z");
            if (invoice.TaxCurrency == invoice.Currency) check.Error("INV-TAX-CURRENCY", "Accounting currency must differ from invoice currency.", "TaxCurrency");
        }
        if (invoice.TaxAmountInAccountingCurrency.HasValue) check.Money(invoice.TaxAmountInAccountingCurrency.Value, "TaxAmountInAccountingCurrency");
        if (invoice.Lines.Count == 0) check.Error("INV-LINES", "At least one invoice line is required.", "Lines");
        if (invoice.Lines.Count > 10000) check.Error("INV-LIMIT", "At most 10,000 invoice lines are supported.", "Lines");
        var identifiers = new HashSet<string>(StringComparer.Ordinal);
        for (int index = 0; index < invoice.Lines.Count; index++) {
            InvoiceLine? line = invoice.Lines[index];
            string path = "Lines[" + index + "]";
            if (line == null) { check.Error("INV-NULL", "Line is null.", path); continue; }
            check.Required(line.Id, path + ".Id");
            if (!identifiers.Add(line.Id)) check.Error("INV-LINE-ID", "Line identifier is duplicated.", path + ".Id");
            check.Required(line.Name, path + ".Name");
            check.Required(line.UnitCode, path + ".UnitCode");
            check.Tax(line.Tax, path + ".Tax");
            check.Period(line.Period, path + ".Period");
            if (line.StandardItemIdentifier != null) check.Identifier(line.StandardItemIdentifier, path + ".StandardItemIdentifier", true);
            if (line.ObjectIdentifier != null) check.Identifier(line.ObjectIdentifier, path + ".ObjectIdentifier", false);
            foreach (InvoiceItemClassification classification in line.Classifications) {
                if (classification == null) { check.Error("INV-NULL", "Item classification is null.", path + ".Classifications"); continue; }
                check.Required(classification.Value, path + ".Classifications.Value");
                check.Required(classification.ListId, path + ".Classifications.ListId");
                if (classification.ListVersion != null) check.Required(classification.ListVersion, path + ".Classifications.ListVersion");
            }
            foreach (InvoiceItemAttribute attribute in line.Attributes) {
                if (attribute == null) { check.Error("INV-NULL", "Item attribute is null.", path + ".Attributes"); continue; }
                check.Required(attribute.Name, path + ".Attributes.Name");
                check.Required(attribute.Value, path + ".Attributes.Value");
            }
            if (line.PriceBaseQuantity <= 0m) check.Error("INV-BASE-QUANTITY", "Price base quantity must be positive.", path + ".PriceBaseQuantity");
            bool invalidPrice = line.UnitPrice < 0m || line.GrossPrice < 0m || line.PriceDiscount < 0m;
            if (invalidPrice)
                check.Error("INV-PRICE", "Item prices and price discounts cannot be negative.", path + ".UnitPrice");
            if (line.PriceDiscount.HasValue && !line.GrossPrice.HasValue)
                check.Error("INV-PRICE", "A price discount requires its gross price.", path + ".PriceDiscount");
            if (!invalidPrice && line.GrossPrice.HasValue) {
                try {
                    if (InvoiceArithmetic.Add(line.GrossPrice.Value, -(line.PriceDiscount ?? 0m)) != line.UnitPrice)
                        check.Error("INV-PRICE", "Net price must equal gross price minus price discount.", path + ".UnitPrice");
                } catch (OverflowException) { check.Error("INV-OVERFLOW", "Gross price minus discount exceeds decimal precision.", path + ".UnitPrice"); }
            }
            foreach (InvoiceAllowanceCharge adjustment in line.AllowancesAndCharges) {
                check.Adjustment(adjustment, path + ".AllowancesAndCharges", false);
                if (adjustment?.Tax != null) check.Error("INV-LINE-TAX", "Line adjustments inherit the line VAT category; do not supply a separate category.", path);
            }
            if (line.DeclaredNetAmount.HasValue) check.Money(line.DeclaredNetAmount.Value, path + ".DeclaredNetAmount");
        }
        foreach (InvoiceAllowanceCharge adjustment in invoice.AllowancesAndCharges) check.Adjustment(adjustment, "AllowancesAndCharges", true);
        foreach (InvoiceReference reference in invoice.PrecedingInvoices) {
            if (reference == null) { check.Error("INV-NULL", "Invoice reference is null.", "PrecedingInvoices"); continue; }
            check.Required(reference.Number, "PrecedingInvoices.Number");
            check.OptionalDate(reference.IssueDate, "PrecedingInvoices.IssueDate");
        }
        foreach (InvoiceNote note in invoice.Notes) {
            if (note == null) check.Error("INV-NULL", "Note is null.", "Notes");
            else {
                check.Required(note.Text, "Notes.Text");
                if (note.SubjectCode != null) check.Code(note.SubjectCode, "Notes.SubjectCode", "^[A-Z]{3}\\z");
            }
        }
        foreach (InvoiceSupportingDocument document in invoice.SupportingDocuments) {
            if (document == null) { check.Error("INV-NULL", "Supporting document is null.", "SupportingDocuments"); continue; }
            check.Required(document.Reference, "SupportingDocuments.Reference");
            if (document.Data != null) {
                check.Required(document.FileName, "SupportingDocuments.FileName");
                check.Required(document.MimeType, "SupportingDocuments.MimeType");
                if (document.Data.Length == 0 || document.Data.Length > 8 * 1024 * 1024)
                    check.Error("INV-ATTACHMENT-SIZE", "Embedded document must contain between 1 byte and 8 MiB.", "SupportingDocuments.Data");
            } else if (document.FileName != null || document.MimeType != null) {
                check.Error("INV-ATTACHMENT", "File name and media type require embedded bytes.", "SupportingDocuments");
            }
            if (document.ExternalUri != null && (!Uri.TryCreate(document.ExternalUri, UriKind.Absolute, out Uri? uri) || !uri.IsWellFormedOriginalString()))
                check.Error("INV-ATTACHMENT-URI", "Supporting document locations must be well-formed absolute URIs.", "SupportingDocuments.ExternalUri");
        }
        check.Payment(invoice.Payment);
        InvoiceCalculation? calculation = null;
        try {
            calculation = InvoiceCalculator.Calculate(invoice);
            check.DeclaredAmounts(invoice, calculation);
        } catch (Exception exception) when (exception is ArgumentException || exception is OverflowException) {
            check.Error("INV-CALCULATION", exception.Message, "Invoice");
        }
        return new InvoiceModelValidationResult(diagnostics, calculation);
    }

    private sealed partial class ModelChecks {
        private readonly List<InvoiceDiagnostic> _diagnostics;
        internal ModelChecks(List<InvoiceDiagnostic> diagnostics) => _diagnostics = diagnostics;
        internal void Error(string code, string message, string path) => _diagnostics.Add(new InvoiceDiagnostic(code, message, path));
        internal void Required(string? value, string path) {
            if (string.IsNullOrWhiteSpace(value)) Error("INV-REQUIRED", "A non-empty value is required.", path);
            else if (value != value!.Trim()) Error("INV-WHITESPACE", "Remove leading or trailing whitespace.", path);
        }
        internal void Code(string? value, string path, string pattern) {
            if (value == null || !Regex.IsMatch(value, pattern, RegexOptions.CultureInvariant)) Error("INV-CODE", "Code has an invalid lexical form.", path);
        }
        internal void Date(DateTime value, string path) {
            if (value == default || value.TimeOfDay != TimeSpan.Zero) Error("INV-DATE", "Supply a non-default date without a time component.", path);
        }
        internal void OptionalDate(DateTime? value, string path) { if (value.HasValue) Date(value.Value, path); }
        internal void Period(InvoicePeriod? value, string path) {
            if (value == null) return;
            if (!value.Start.HasValue && !value.End.HasValue) Error("INV-PERIOD", "A period requires at least one boundary.", path);
            OptionalDate(value.Start, path + ".Start"); OptionalDate(value.End, path + ".End");
            if (value.Start > value.End) Error("INV-PERIOD", "Period end precedes its start.", path);
        }
        internal void Money(decimal value, string path) {
            if (decimal.Round(value, 2) != value) Error("INV-DECIMALS", "Monetary amount permits at most two decimal places.", path);
        }
        internal void Address(InvoiceAddress? address, string path) {
            if (address == null) Error("INV-REQUIRED", "Postal address is required.", path);
            else Code(address.CountryCode, path + ".CountryCode", "^[A-Z]{2}\\z");
        }
        internal void Party(InvoiceParty? party, string path) {
            if (party == null) { Error("INV-REQUIRED", "Party is required.", path); return; }
            Required(party.Name, path + ".Name"); Address(party.Address, path + ".Address");
            PartyIdentifiers(party, path);
        }
        internal void PartyIdentifiers(InvoiceParty party, string path) {
            foreach (InvoiceIdentifier identifier in party.Identifiers) Identifier(identifier, path + ".Identifiers", false);
            if (party.LegalRegistration != null) Identifier(party.LegalRegistration, path + ".LegalRegistration", false);
            if (party.ElectronicAddress != null) Identifier(party.ElectronicAddress, path + ".ElectronicAddress", true);
        }
        internal void Identifier(InvoiceIdentifier? identifier, string path, bool requireScheme) {
            if (identifier == null) { Error("INV-NULL", "Identifier is null.", path); return; }
            Required(identifier.Value, path + ".Value");
            if (requireScheme || identifier.SchemeId != null) Required(identifier.SchemeId, path + ".SchemeId");
        }
        internal void Tax(InvoiceTaxCategory? tax, string path, bool breakdown = false) {
            if (tax == null) { Error("INV-REQUIRED", "VAT category is required.", path); return; }
            if (!new[] { "S", "Z", "E", "AE", "K", "G", "O", "L", "M" }.Contains(tax.Code))
                Error("INV-VAT-CATEGORY", "Unsupported VAT category.", path + ".Code");
            if (tax.Code == "O") {
                if (tax.Rate.HasValue && !(breakdown && tax.Rate == 0m)) Error("INV-VAT-RATE", "Outside-scope VAT must not declare a rate.", path + ".Rate");
            } else if (!tax.Rate.HasValue || tax.Rate < 0m) {
                Error("INV-VAT-RATE", "This VAT category requires a non-negative rate.", path + ".Rate");
            } else if (tax.Code == "S" && tax.Rate <= 0m || new[] { "Z", "E", "AE", "K", "G" }.Contains(tax.Code) && tax.Rate != 0m) {
                Error("INV-VAT-RATE", "Rate does not match the VAT category.", path + ".Rate");
            }
        }
        internal void Adjustment(InvoiceAllowanceCharge? item, string path, bool documentLevel) {
            if (item == null) { Error("INV-NULL", "Adjustment is null.", path); return; }
            Money(item.Amount, path + ".Amount");
            if (item.Amount < 0m) Error("INV-ADJUSTMENT", "Allowance/charge amount cannot be negative.", path + ".Amount");
            if (string.IsNullOrWhiteSpace(item.Reason) && string.IsNullOrWhiteSpace(item.ReasonCode)) Error("INV-ADJUSTMENT-REASON", "Supply a reason or reason code.", path);
            if (item.BaseAmount.HasValue != item.Percentage.HasValue) Error("INV-ADJUSTMENT-BASE", "Supply percentage and base amount together.", path);
            if (item.BaseAmount < 0m || item.Percentage < 0m) Error("INV-ADJUSTMENT-BASE", "Base and percentage cannot be negative.", path);
            if (item.BaseAmount.HasValue) Money(item.BaseAmount.Value, path + ".BaseAmount");
            if (item.BaseAmount.HasValue && item.Percentage.HasValue) {
                try {
                    if (InvoiceArithmetic.RoundedProduct(item.BaseAmount.Value, item.Percentage.Value, 100m) != item.Amount)
                        Error("INV-ADJUSTMENT-AMOUNT", "Amount differs from the rounded base multiplied by percentage.", path + ".Amount");
                } catch (OverflowException) { Error("INV-OVERFLOW", "Adjustment calculation exceeds decimal capacity.", path); }
            }
            if (documentLevel) Tax(item.Tax, path + ".Tax");
        }
        internal void Payment(InvoicePayment? payment) {
            if (payment == null) return;
            Code(payment.MeansCode, "Payment.MeansCode", "^[0-9]{1,3}\\z");
            foreach (InvoiceBankAccount account in payment.Accounts) {
                if (account == null) Error("INV-NULL", "Bank account is null.", "Payment.Accounts");
                else Required(account.Identifier, "Payment.Accounts.Identifier");
            }
            if (payment.CardNumber != null && !Regex.IsMatch(payment.CardNumber, "^[0-9]{4,6}\\z", RegexOptions.CultureInvariant))
                Error("INV-CARD", "Supply only the last four to six card digits.", "Payment.CardNumber");
            if (payment.CardHolder != null && payment.CardNumber == null) Error("INV-CARD", "Card holder requires masked card digits.", "Payment.CardHolder");
        }
    }
}
