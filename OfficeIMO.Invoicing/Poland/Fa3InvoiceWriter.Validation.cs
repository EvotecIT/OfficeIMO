using System.Globalization;
using System.Xml;

namespace OfficeIMO.Invoicing;

public static partial class Fa3InvoiceWriter {
    private static Prepared Prepare(Invoice invoice, Fa3InvoiceWriteOptions options) {
        if (invoice == null) throw new ArgumentNullException(nameof(invoice));
        if (options == null) throw new ArgumentNullException(nameof(options));
        var result = new Prepared(); var diagnostics = new InvoiceDiagnosticBuffer();
        void Error(string path, string message) => diagnostics.Add("FA3-AUTHORING", message, path);
        try { new InvoiceModelLimits().Check(invoice); }
        catch (Exception exception) when (exception is InvalidDataException or XmlException) {
            Error("Invoice", exception.Message); result.Diagnostics = diagnostics.ToList().AsReadOnly(); return result;
        }
        if (invoice.Seller == null || invoice.Buyer == null || invoice.Seller.Address == null || invoice.Buyer.Address == null ||
            invoice.Lines.Any(line => line == null || line.Tax == null)) {
            Error("Invoice", "Parties, addresses, lines and line taxes must not be null."); result.Diagnostics = diagnostics.ToList().AsReadOnly(); return result;
        }
        void Text(string path, string? value, int maximum = 256, bool required = false) {
            if (value == null) { if (required) Error(path, "A value is required."); return; }
            if (string.IsNullOrWhiteSpace(value) || value.Length > maximum) { Error(path, "Supply nonblank text of at most " + maximum + " characters."); return; }
            try { XmlConvert.VerifyXmlChars(value); } catch (XmlException) { Error(path, "Text contains characters XML cannot represent."); }
        }
        void Money(string path, decimal value) {
            if (decimal.Round(value, 2) != value || value > 9_999_999_999_999_999.99m || value < -9_999_999_999_999_999.99m)
                Error(path, "National amounts require at most sixteen whole digits and two fractional digits.");
        }
        void Optional(string path, object? value) { if (value != null) Error(path, "The populated field has no supported FA(3) authoring mapping."); }
        if (options.Kind < Fa3InvoiceKind.TaxInvoice || options.Kind > Fa3InvoiceKind.SettlementCorrection) Error("Kind", "Unknown national document kind.");
        if (options.CreatedAt < new DateTimeOffset(2025, 9, 1, 0, 0, 0, TimeSpan.Zero) || options.CreatedAt > new DateTimeOffset(2050, 1, 1, 23, 59, 59, TimeSpan.Zero)) Error("CreatedAt", "Creation time is outside the pinned schema range.");
        Text("Number", invoice.Number, 256, true); Text("Currency", invoice.Currency, 3, true);
        if (invoice.Currency == null || !Fa3ScalarContract.Currencies.Contains(invoice.Currency)) Error("Currency", "Supply a currency code in the pinned national schema dictionary.");
        if (invoice.IssueDate == default) Error("IssueDate", "An issue date is required.");
        Fa3ScalarContract.Date("IssueDate", invoice.IssueDate, 2006, Error);
        Fa3ScalarContract.Date("DueDate", invoice.DueDate, 2016, Error);
        Fa3ScalarContract.Date("TaxPointDate", invoice.TaxPointDate, 2006, Error);
        Fa3ScalarContract.Date("Period.Start", invoice.Period?.Start, 2006, Error);
        Fa3ScalarContract.Date("Period.End", invoice.Period?.End, 2006, Error);
        Text("SystemInfo", options.SystemInfo); Text("PlaceOfIssue", options.PlaceOfIssue);
        Text("CorrectionReason", options.CorrectionReason); Text("Annotations.ExemptionLegalBasis", options.Annotations.ExemptionLegalBasis);
        if (options.Annotations.ExemptionBasisKind.HasValue != (options.Annotations.ExemptionLegalBasis != null)) Error("Annotations.ExemptionLegalBasis", "Supply both the exemption field kind and legal basis.");
        if (options.Annotations.ExemptionBasisKind < Fa3ExemptionBasisKind.NationalProvision || options.Annotations.ExemptionBasisKind > Fa3ExemptionBasisKind.Other ||
            options.Annotations.MarginProcedure < Fa3MarginProcedure.None || options.Annotations.MarginProcedure > Fa3MarginProcedure.CollectorsAndAntiques) Error("Annotations", "Unknown national annotation choice.");
        bool correction = IsCorrection(options.Kind);
        string expectedType = correction ? "384" : options.Kind == Fa3InvoiceKind.Advance ? "386" : options.Annotations.SelfBilling ? "389" : "380";
        if (invoice.TypeCode != expectedType) Error("TypeCode", "The common type code must agree with the explicitly selected national kind and self-billing declaration: " + expectedType + ".");
        if (options.CorrectionTimingCode.HasValue && options.CorrectionTimingCode is not (1 or 2 or 3)) Error("CorrectionTimingCode", "TypKorekty permits 1, 2 or 3.");
        if (correction && invoice.PrecedingInvoices.Count == 0) Error("PrecedingInvoices", "A correction requires an explicitly identified preceding invoice.");
        if (invoice.PrecedingInvoices.Count > 1000 || options.CorrectionKsefNumbers.Count > 1000 || options.AdvanceInvoiceReferences.Count > 100 || options.UnitLabels.Count > 512 || options.LineTaxLabels.Count > 20_000) {
            Error("Options", "National reference or label counts exceed the supported bounds."); result.Diagnostics = diagnostics.ToList().AsReadOnly(); return result;
        }
        if (!correction) {
            if (invoice.PrecedingInvoices.Count != 0 || options.CorrectionKsefNumbers.Count != 0 || options.CorrectionReason != null || options.CorrectionTimingCode.HasValue || options.PreviousAdvanceOrSettlementTotal.HasValue) Error("Correction", "Correction fields require a correction kind.");
        }
        foreach (InvoiceReference reference in invoice.PrecedingInvoices) {
            if (reference == null) { Error("PrecedingInvoices", "Reference is null."); continue; }
            Text("PrecedingInvoices.Number", reference.Number, 256, true);
            if (!reference.IssueDate.HasValue || reference.IssueDate == default(DateTime)) Error("PrecedingInvoices.IssueDate", "The preceding invoice issue date is required.");
            Fa3ScalarContract.Date("PrecedingInvoices.IssueDate", reference.IssueDate, 2006, Error);
        }
        foreach (KeyValuePair<string, string> entry in options.CorrectionKsefNumbers) {
            if (!invoice.PrecedingInvoices.Any(reference => reference != null && reference.Number == entry.Key)) Error("CorrectionKsefNumbers", "A KSeF number has no matching preceding-invoice reference.");
            CheckKsefNumber(entry.Value, "CorrectionKsefNumbers", Error);
        }
        foreach (Fa3AdvanceInvoiceReference reference in options.AdvanceInvoiceReferences) {
            if (reference == null) { Error("AdvanceInvoiceReferences", "Reference is null."); continue; }
            Text("AdvanceInvoiceReferences.Number", reference.Number, 256);
            if (reference.Number == null && reference.KsefNumber == null) Error("AdvanceInvoiceReferences", "Supply the invoice number or actual KSeF number.");
            if (reference.KsefNumber != null) CheckKsefNumber(reference.KsefNumber, "AdvanceInvoiceReferences.KsefNumber", Error);
        }
        if (options.AdvanceInvoiceReferences.Count != 0 && options.Kind is not (Fa3InvoiceKind.Settlement or Fa3InvoiceKind.SettlementCorrection)) Error("AdvanceInvoiceReferences", "Previous advance references belong to a settlement or settlement correction.");
        if (options.Kind is Fa3InvoiceKind.Settlement or Fa3InvoiceKind.SettlementCorrection && options.AdvanceInvoiceReferences.Count == 0) Error("AdvanceInvoiceReferences", "A settlement requires a previous advance invoice reference.");
        if (options.Kind is Fa3InvoiceKind.Advance or Fa3InvoiceKind.AdvanceCorrection && options.Order == null) Error("Order", "An advance requires explicit order information; an advance correction requires difference rows and original/revised order totals.");
        if (options.Order != null && options.Order.PreviousTotal.HasValue != (options.Kind == Fa3InvoiceKind.AdvanceCorrection)) Error("Order", "Full order rows belong to an advance; advance corrections require the explicit difference-row contract. Before/after rows remain unsupported.");
        if (options.Order != null && options.Kind is not (Fa3InvoiceKind.Advance or Fa3InvoiceKind.AdvanceCorrection)) Error("Order", "Order authoring is supported only for advances and advance corrections.");
        if (options.PreviousAdvanceOrSettlementTotal.HasValue) {
            if (options.Kind is not (Fa3InvoiceKind.AdvanceCorrection or Fa3InvoiceKind.SettlementCorrection)) Error("PreviousAdvanceOrSettlementTotal", "P_15ZK applies only to advance or settlement corrections.");
            Money("PreviousAdvanceOrSettlementTotal", options.PreviousAdvanceOrSettlementTotal.Value);
        }
        foreach (KeyValuePair<string, string> entry in options.UnitLabels) { Text("UnitLabels.Code", entry.Key, 16, true); Text("UnitLabels.Value", entry.Value, 256, true); }
        foreach (KeyValuePair<string, string> entry in options.LineTaxLabels) { Text("LineTaxLabels.Id", entry.Key, 10, true); Text("LineTaxLabels.Value", entry.Value, 5, true); }
        var rowIds = new HashSet<string>(invoice.Lines.Select(line => line.Id), StringComparer.Ordinal);
        if (options.Order != null) rowIds.UnionWith(options.Order.Lines.Select(line => line.Id));
        foreach (string id in options.LineTaxLabels.Keys)
            if (!rowIds.Contains(id)) Error("LineTaxLabels", "A national tax label has no matching invoice or order row.");
        CheckUnsupportedInvoiceFields(invoice, Optional, Error);
        CheckParty(invoice.Seller, "Seller", true, options.Kind, Text, Error, Optional);
        CheckParty(invoice.Buyer, "Buyer", false, options.Kind, Text, Error, Optional);
        if (invoice.Period != null && (!invoice.Period.Start.HasValue || !invoice.Period.End.HasValue || invoice.Period.Start > invoice.Period.End || invoice.TaxPointDate.HasValue)) Error("Period", "Supply a complete ordered period or a tax point date, not both.");
        CheckLines(invoice.Lines, "Lines", options, Text, Error, Optional, Money);
        if (invoice.Lines.Count == 0 && options.Kind == Fa3InvoiceKind.TaxInvoice) Error("Lines", "An ordinary invoice requires billed lines.");
        if (options.Order != null) {
            var orderInvoice = new Invoice { Currency = invoice.Currency ?? string.Empty };
            foreach (InvoiceLine line in options.Order.Lines) orderInvoice.Lines.Add(line);
            try { new InvoiceModelLimits().Check(orderInvoice); }
            catch (Exception exception) when (exception is InvalidDataException or XmlException) { Error("Order", exception.Message); }
            Money("Order.Total", options.Order.Total);
            if (options.Order.PreviousTotal.HasValue) Money("Order.PreviousTotal", options.Order.PreviousTotal.Value);
            CheckLines(options.Order.Lines, "Order.Lines", options, Text, Error, Optional, Money);
            if (!diagnostics.HasErrors) {
                try {
                    result.OrderCalculation = InvoiceCalculator.Calculate(orderInvoice);
                    CheckCalculatedLines(options.Order.Lines, result.OrderCalculation, "Order.Lines", Error, Money);
                    decimal orderDifference = result.OrderCalculation.TaxInclusiveTotal;
                    if (options.Order.DifferenceTaxAmounts != null) {
                        for (int index = 0; index < options.Order.Lines.Count; index++) {
                            decimal? tax = options.Order.DifferenceTaxAmounts[index];
                            string location = "Order.DifferenceTaxAmounts[" + index + "]";
                            if (tax.HasValue) Money(location, tax.Value);
                            bool taxable = options.Annotations.MarginProcedure == Fa3MarginProcedure.None && options.Order.Lines[index].Tax.Code == "S";
                            if (taxable && !tax.HasValue || !taxable && tax.GetValueOrDefault() != 0) Error(location, "Supply explicit VAT for taxable difference rows; other rows may declare only zero or no VAT.");
                        }
                        orderDifference = InvoiceArithmetic.Add(result.OrderCalculation.LineNetTotal, InvoiceArithmetic.Sum(options.Order.DifferenceTaxAmounts.Select(tax => tax ?? 0m)));
                    } else if (result.OrderCalculation.Taxes.Any(tax => tax.TaxableAmount < 0 || tax.TaxAmount < 0))
                        Error("Order", "Negative order buckets require the explicit correction-difference contract; signed national VAT is not inferred from EN rounding.");
                    decimal expectedOrderTotal = InvoiceArithmetic.Add(options.Order.PreviousTotal ?? 0m, orderDifference);
                    decimal discrepancy = InvoiceArithmetic.Add(options.Order.Total, -expectedOrderTotal);
                    if (options.Order.DifferenceTaxAmounts != null ? discrepancy != 0 : Math.Abs(discrepancy) > 0.01m) Error("Order.Total", "The order value must reconcile the explicit correction differences, or agree with a full-row calculation within one cent.");
                    if (options.Order.PreviousTotal.HasValue && (options.Order.Total == options.Order.PreviousTotal.Value || orderDifference == 0)) Error("Order", "An unchanged order requires explicit before/after authoring, which this difference-row path cannot represent.");
                }
                catch (Exception exception) when (exception is InvalidDataException or ArgumentException or OverflowException) { Error("Order", exception.Message); }
            }
        }
        if (options.FiscalAmounts == null && (options.Kind != Fa3InvoiceKind.TaxInvoice || options.Annotations.MarginProcedure != Fa3MarginProcedure.None || invoice.Currency != "PLN"))
            Error("FiscalAmounts", "Supply explicit national totals for this document kind, margin procedure or foreign currency. They cannot be inferred from EN totals.");
        if (!diagnostics.HasErrors && invoice.Lines.Count != 0) {
            try {
                result.Calculation = InvoiceCalculator.Calculate(invoice);
                CheckCalculatedLines(invoice.Lines, result.Calculation, "Lines", Error, Money);
            }
            catch (Exception exception) when (exception is InvalidDataException or ArgumentException or OverflowException) { Error("Lines", exception.Message); }
        }
        if (options.FiscalAmounts != null) {
            if (options.Kind == Fa3InvoiceKind.TaxInvoice && invoice.Currency == "PLN" && options.Annotations.MarginProcedure == Fa3MarginProcedure.None)
                Error("FiscalAmounts", "Ordinary PLN amounts use the canonical calculation; an explicit override is not supported.");
            result.Amounts = options.FiscalAmounts;
        }
        else if (!diagnostics.HasErrors && result.Calculation != null) {
            if (result.Calculation.Taxes.Any(tax => tax.TaxableAmount < 0 || tax.TaxAmount < 0))
                Error("FiscalAmounts", "Negative ordinary PLN buckets are outside this calculation contract; select the applicable native document kind rather than overriding ordinary totals.");
            else result.Amounts = CalculateOrdinaryAmounts(invoice, options, result.Calculation);
        }
        if (result.Amounts != null) CheckFiscalAmounts(invoice, options, result.Amounts, Money, Error);
        result.Diagnostics = diagnostics.ToList().AsReadOnly(); return result;
    }

    private static void CheckKsefNumber(string value, string path, Action<string, string> error) {
        if (value == null || value.Length is not (35 or 36) || !System.Text.RegularExpressions.Regex.IsMatch(value, "^[0-9]{10}-[0-9]{8}-[A-F0-9]{6}-?[A-F0-9]{6}-[A-F0-9]{2}$"))
            error(path, "Supply the actual 35- or 36-character KSeF identifier; no identifier is generated by the writer.");
    }

    private static Fa3FiscalAmounts CalculateOrdinaryAmounts(Invoice invoice, Fa3InvoiceWriteOptions options, InvoiceCalculation calculation) {
        var buckets = new List<Fa3TaxSummary>();
        foreach (string suffix in TaxSuffixes) {
            int[] indexes = Enumerable.Range(0, invoice.Lines.Count).Where(index => TaxSuffix(TaxLabel(invoice.Lines[index], options)) == suffix).ToArray();
            if (indexes.Length == 0) continue;
            decimal basis = InvoiceArithmetic.Sum(indexes.Select(index => calculation.Lines[index].NetAmount));
            decimal? tax = suffix is "1" or "2" or "3" or "4" ? InvoiceArithmetic.Sum(calculation.Taxes.Where(item =>
                indexes.Any(index => invoice.Lines[index].Tax.Code == item.CategoryCode && invoice.Lines[index].Tax.Rate == item.Rate)).Select(item => item.TaxAmount)) : null;
            buckets.Add(new Fa3TaxSummary(suffix, basis, tax));
        }
        return new Fa3FiscalAmounts(calculation.TaxInclusiveTotal, buckets);
    }

    private static void CheckFiscalAmounts(Invoice invoice, Fa3InvoiceWriteOptions options, Fa3FiscalAmounts amounts, Action<string, decimal> money, Action<string, string> error) {
        money("FiscalAmounts.Total", amounts.Total);
        var seen = new HashSet<string>(StringComparer.Ordinal);
        foreach (Fa3TaxSummary tax in amounts.Taxes) {
            if (!TaxSuffixes.Contains(tax.FieldSuffix) || !seen.Add(tax.FieldSuffix)) error("FiscalAmounts.Taxes", "Supply distinct supported national VAT buckets.");
            money("FiscalAmounts.Taxes.P_13_" + tax.FieldSuffix, tax.TaxableAmount);
            bool taxable = tax.FieldSuffix is "1" or "2" or "3" or "4";
            if (options.Annotations.MarginProcedure != Fa3MarginProcedure.None && tax.FieldSuffix != "11" ||
                options.Annotations.MarginProcedure == Fa3MarginProcedure.None && tax.FieldSuffix == "11") error("FiscalAmounts.Taxes", "Margin declarations require the margin annotation and only the national margin bucket.");
            if (tax.FieldSuffix == "7" && options.Annotations.ExemptionLegalBasis == null) error("Annotations.ExemptionLegalBasis", "An exempt national bucket requires an explicit legal basis.");
            if (tax.FieldSuffix == "10" && !options.Annotations.ReverseCharge) error("Annotations.ReverseCharge", "A domestic reverse-charge bucket requires the national annotation.");
            if (taxable && !tax.TaxAmount.HasValue || !taxable && tax.FieldSuffix != "5" && tax.TaxAmount.HasValue) error("FiscalAmounts.Taxes.P_14_" + tax.FieldSuffix, "The tax declaration does not match the selected national bucket.");
            if (tax.TaxAmount.HasValue) money("FiscalAmounts.Taxes.P_14_" + tax.FieldSuffix, tax.TaxAmount.Value);
            if (tax.TaxAmountInPln.HasValue) {
                money("FiscalAmounts.Taxes.P_14_" + tax.FieldSuffix + "W", tax.TaxAmountInPln.Value);
                if (!taxable || invoice.Currency == "PLN") error("FiscalAmounts.Taxes.P_14_" + tax.FieldSuffix + "W", "PLN tax conversion requires a taxable bucket and a foreign invoice currency.");
            }
            if (taxable && invoice.Currency != "PLN" && tax.TaxAmount != 0 && !tax.TaxAmountInPln.HasValue) error("FiscalAmounts.Taxes.P_14_" + tax.FieldSuffix + "W", "Supply the per-bucket PLN tax amount explicitly for foreign currency.");
        }
        if (invoice.TaxCurrency != null || invoice.TaxAmountInAccountingCurrency.HasValue) {
            try {
                if (invoice.TaxCurrency != "PLN" || invoice.Currency == "PLN" || !invoice.TaxAmountInAccountingCurrency.HasValue ||
                    InvoiceArithmetic.Sum(amounts.Taxes.Select(tax => tax.TaxAmountInPln ?? 0m)) != invoice.TaxAmountInAccountingCurrency.Value)
                    error("TaxAmountInAccountingCurrency", "Accounting-currency VAT must match the explicit national per-bucket PLN declarations.");
            } catch (OverflowException) { error("TaxAmountInAccountingCurrency", "The national PLN tax sum exceeds decimal precision."); }
        }
    }

    private static void CheckCalculatedLines(IEnumerable<InvoiceLine> lines, InvoiceCalculation calculation, string path, Action<string, string> error, Action<string, decimal> money) {
        int index = 0;
        foreach (InvoiceLine line in lines) {
            InvoiceCalculatedLine amount = calculation.Lines[index];
            string location = path + "[" + index++ + "].DeclaredNetAmount";
            money(location, amount.NetAmount);
            if (amount.FormulaNetAmount < 0 && decimal.Round(amount.FormulaNetAmount, 2) != amount.FormulaNetAmount && !line.DeclaredNetAmount.HasValue)
                error(location, "Supply a signed national line amount explicitly when a negative formula requires rounding; EN negative-half rounding is not inferred.");
            if (line.DeclaredNetAmount.HasValue && Math.Abs(InvoiceArithmetic.Add(amount.NetAmount, -amount.FormulaNetAmount)) > 0.01m)
                error(location, "The declared line amount differs from the supported quantity and unit-price formula by more than one cent; unrepresented adjustments cannot be discarded.");
        }
    }
}
