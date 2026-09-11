using System.Text;
using System.Xml;

namespace OfficeIMO.Invoicing;

/// <summary>Bounds model expansion before building an XML tree, including repeated references to the same data.</summary>
internal sealed class InvoiceModelLimits {
    internal const int MaximumCollectionItems = 50000;
    private long _textBytes;
    private long _binaryBytes;
    private int _items;
    internal int Check(Invoice invoice) {
        Text(invoice.Number, invoice.TypeCode, invoice.Currency, invoice.BusinessProcessId, invoice.BuyerReference, invoice.TaxPointDateCode,
            invoice.ProjectReference, invoice.ContractReference, invoice.PurchaseOrderReference, invoice.SalesOrderReference, invoice.ReceivingAdviceReference,
            invoice.DespatchAdviceReference, invoice.TenderReference, invoice.AccountingReference, invoice.PaymentTerms, invoice.TaxCurrency);
        Party(invoice.Seller); Party(invoice.Buyer); Party(invoice.Payee); Party(invoice.TaxRepresentative); Identifier(invoice.ObjectIdentifier);
        if (invoice.Delivery != null) { Text(invoice.Delivery.Name); Address(invoice.Delivery.Address); Identifier(invoice.Delivery.LocationIdentifier); }
        Each(invoice.Notes, note => Text(note.Text, note.SubjectCode));
        Each(invoice.PrecedingInvoices, reference => Text(reference.Number));
        Each(invoice.SupportingDocuments, document => {
            Text(document.Reference, document.Description, document.ExternalUri, document.FileName, document.MimeType);
            _binaryBytes += document.Data?.LongLength ?? 0;
            if (_binaryBytes > 8 * 1024 * 1024) throw new InvalidDataException("Combined embedded invoice documents exceed 8 MiB.");
        });
        Each(invoice.Lines, line => {
            Text(line.Id, line.Name, line.Description, line.Note, line.UnitCode, line.OrderLineReference, line.AccountingReference,
                line.SellerItemIdentifier, line.BuyerItemIdentifier, line.OriginCountryCode);
            Identifier(line.ObjectIdentifier); Identifier(line.StandardItemIdentifier); Tax(line.Tax);
            Each(line.AllowancesAndCharges, Adjustment);
            Each(line.Classifications, value => Text(value.Value, value.ListId, value.ListVersion));
            Each(line.Attributes, value => Text(value.Name, value.Value));
        });
        Each(invoice.AllowancesAndCharges, Adjustment); Each(invoice.DeclaredTaxes, tax => Tax(tax.Category));
        if (invoice.Payment != null) {
            Text(invoice.Payment.MeansCode, invoice.Payment.MeansText, invoice.Payment.Reference, invoice.Payment.CardNumber, invoice.Payment.CardHolder,
                invoice.Payment.MandateReference, invoice.Payment.CreditorIdentifier, invoice.Payment.DebitedAccount);
            Each(invoice.Payment.Accounts, account => Text(account.Identifier, account.Name, account.ProviderIdentifier));
        }
        return _items;
    }
    private void Party(InvoiceParty? party) {
        if (party == null) return;
        Text(party.Name, party.TradingName, party.LegalInformation, party.VatIdentifier, party.TaxRegistration);
        Identifier(party.LegalRegistration); Identifier(party.ElectronicAddress); Each(party.Identifiers, Identifier); Address(party.Address);
        if (party.Contact != null) Text(party.Contact.Name, party.Contact.Telephone, party.Contact.Email);
    }
    private void Address(InvoiceAddress? address) {
        if (address != null) Text(address.Line1, address.Line2, address.Line3, address.City, address.PostCode, address.Subdivision, address.CountryCode);
    }
    private void Tax(InvoiceTaxCategory? tax) { if (tax != null) Text(tax.Code, tax.ExemptionReason, tax.ExemptionReasonCode); }
    private void Identifier(InvoiceIdentifier? identifier) { if (identifier != null) Text(identifier.Value, identifier.SchemeId); }
    private void Adjustment(InvoiceAllowanceCharge adjustment) { Text(adjustment.Reason, adjustment.ReasonCode); Tax(adjustment.Tax); }
    private void Text(params string?[] values) {
        foreach (string? value in values) {
            if (value == null) continue;
            if (value.Length > 1024 * 1024) throw new InvalidDataException("An invoice text value exceeds one million characters.");
            XmlConvert.VerifyXmlChars(value);
            _textBytes += Encoding.UTF8.GetByteCount(value);
            if (_textBytes > 4 * 1024 * 1024) throw new InvalidDataException("Combined invoice text exceeds 4 MiB of UTF-8 data.");
        }
    }
    private void Each<T>(IEnumerable<T> items, Action<T> visit) where T : class {
        foreach (T item in items) {
            if (++_items > MaximumCollectionItems) throw new InvalidDataException("Invoice model exceeds 50,000 collection items.");
            if (item == null) throw new InvalidDataException("Invoice model contains a null collection item.");
            visit(item);
        }
    }
}
