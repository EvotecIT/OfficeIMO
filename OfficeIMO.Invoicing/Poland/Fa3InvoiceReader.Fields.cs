using System.Globalization;
using System.Xml.Linq;

namespace OfficeIMO.Invoicing;

public static partial class Fa3InvoiceReader {
    private static InvoiceParty ReadParty(InvoiceXmlReadContext c, XElement party, bool seller) {
        XElement id = c.Require(party, Ns + "DaneIdentyfikacyjne");
        var result = new InvoiceParty { Name = c.Text(id, Ns + "Nazwa") ?? string.Empty };
        string? nip = c.Text(id, Ns + "NIP");
        if (nip != null) c.AddTo(result.TaxRegistrations, new InvoiceTaxRegistration(nip, "NIP", InvoiceTaxRegistrationKind.Fiscal));
        string? country = c.Text(id, Ns + "KodUE"), vat = c.Text(id, Ns + "NrVatUE");
        if (country != null && vat != null) c.AddTo(result.TaxRegistrations, new InvoiceTaxRegistration(country + vat, InvoiceTaxRegistration.VatScheme));
        string? foreignCountry = c.Text(id, Ns + "KodKraju"), foreignId = c.Text(id, Ns + "NrID");
        if (foreignId != null || foreignCountry != null)
            c.Loss(id, "Foreign national tax identification requires an explicit fiscal-identifier mapping; it is not a buyer business identifier or VAT registration.");
        XElement? noId = c.Child(id, Ns + "BrakID");
        if (noId != null) c.Expected(noId, "1");
        XElement? address = c.Child(party, Ns + "Adres");
        if (address != null) result.Address = new InvoiceAddress {
            CountryCode = c.Required(address, Ns + "KodKraju"), Line1 = c.Text(address, Ns + "AdresL1"), Line2 = c.Text(address, Ns + "AdresL2")
        };
        XElement[] contacts = c.Children(party, Ns + "DaneKontaktowe").ToArray();
        if (contacts.Length == 1) result.Contact = new InvoiceContact {
            Email = c.Text(contacts[0], Ns + "Email"), Telephone = c.Text(contacts[0], Ns + "Telefon")
        };
        else foreach (XElement contact in contacts) c.Loss(contact, "Multiple contacts cannot be collapsed into one common party contact.");
        if (!seller) {
            c.Expected(c.Child(party, Ns + "JST"), "2");
            c.Expected(c.Child(party, Ns + "GV"), "2");
        }
        return result;
    }

    private static void ReadLine(InvoiceXmlReadContext c, XElement element, Invoice invoice, List<Fa3InvoiceLineData> rows, InvoiceDiagnosticBuffer projection, bool commonBilledLine) {
        string number = c.Required(element, Ns + "NrWierszaFa");
        string? uuid = c.Text(element, Ns + "UU_ID"), unit = c.Text(element, Ns + "P_8A"), rate = c.Text(element, Ns + "P_12"), name = c.Text(element, Ns + "P_7");
        decimal? quantity = c.Decimal(element, Ns + "P_8B"), price = c.Decimal(element, Ns + "P_9A"), net = c.Decimal(element, Ns + "P_11");
        XElement? previousElement = c.Child(element, Ns + "StanPrzed");
        bool previous = previousElement != null;
        if (previous) c.Expected(previousElement, "1");
        c.AddTo(rows, new Fa3InvoiceLineData(number, uuid, unit, rate, name, quantity, price, net, previous));
        string location = "Fa.FaWiersz[" + (rows.Count - 1).ToString(CultureInfo.InvariantCulture) + "]";
        string? unitCode = unit switch { "szt." or "szt" => "C62", "kg" => "KGM", "g" => "GRM", "l" => "LTR", "m" => "MTR", "m2" => "MTK", "m3" => "MTQ", "godz." or "h" => "HUR", _ => null };
        InvoiceTaxCategory? tax = rate switch {
            "0 KR" => new InvoiceTaxCategory { Code = "Z", Rate = 0 },
            "0 WDT" => new InvoiceTaxCategory { Code = "K", Rate = 0 },
            "0 EX" => new InvoiceTaxCategory { Code = "G", Rate = 0 },
            "zw" => new InvoiceTaxCategory { Code = "E", Rate = 0 },
            "oo" => new InvoiceTaxCategory { Code = "AE", Rate = 0 },
            "np I" or "np II" => new InvoiceTaxCategory { Code = "O", Rate = null },
            "23" or "22" or "8" or "7" or "5" or "4" or "3" => new InvoiceTaxCategory { Code = "S", Rate = decimal.Parse(rate!, CultureInfo.InvariantCulture) },
            _ => null
        };
        if (!commonBilledLine || previous || unitCode == null || tax == null || name == null || !quantity.HasValue || !price.HasValue || price < 0) {
            projection.Add("FA3-LINE-PROJECTION", "Row is retained nationally but cannot be used as a complete common billed line (national document kind, previous state, missing fields, unsupported unit/rate or negative price).", location);
            return;
        }
        c.AddTo(invoice.Lines, new InvoiceLine {
            Id = number, Name = name, Quantity = quantity.Value, UnitPrice = price.Value,
            UnitCode = unitCode, Tax = tax, DeclaredNetAmount = net, SellerItemIdentifier = c.Text(element, Ns + "Indeks")
        });
        // UUID, literal unit and the two distinct outside-territory rate labels remain national metadata.
        if (uuid != null || rate == "np I" || rate == "np II")
            projection.Add("FA3-LINE-METADATA", "National row identity or outside-territory classification needs an explicit target mapping.", location);
    }

    private static void ReadPayment(InvoiceXmlReadContext c, XElement fa, Invoice invoice, InvoiceDiagnosticBuffer projection) {
        XElement? payment = c.Child(fa, Ns + "Platnosc");
        if (payment == null) return;
        XElement[] due = c.Children(payment, Ns + "TerminPlatnosci").ToArray();
        if (due.Length == 1) invoice.DueDate = c.Date(c.Child(due[0], Ns + "Termin"));
        else foreach (XElement date in due) c.Loss(date, "Multiple payment terms cannot be collapsed into one common due date.");
        string? means = c.Text(payment, Ns + "FormaPlatnosci");
        string? commonMeans = means switch { "1" => "10", "2" => "48", "4" => "20", "6" => "30", _ => null };
        XElement[] accounts = c.Children(payment, Ns + "RachunekBankowy").ToArray();
        if (commonMeans != null) {
            if (accounts.Length == 0) c.AddTo(invoice.Payments, new InvoicePayment { MeansCode = commonMeans });
            foreach (XElement account in accounts) {
                string identifier = c.Required(account, Ns + "NrRB");
                c.AddTo(invoice.Payments, new InvoicePayment {
                    MeansCode = commonMeans,
                    Account = new InvoiceBankAccount {
                        Identifier = identifier, IsIban = InvoiceBankAccountIdentity.IsValidIban(identifier),
                        ProviderIdentifier = c.Text(account, Ns + "SWIFT")
                    }
                });
            }
        } else if (means != null || accounts.Length != 0)
            projection.Add("FA3-PAYMENT-PROJECTION", "National payment form or accounts require an explicit common payment-means mapping.", "Fa.Platnosc");
    }
}
