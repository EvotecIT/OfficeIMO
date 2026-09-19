namespace OfficeIMO.Invoicing;

public static partial class InvoiceSerializer {
    private static void CheckExemptionConflicts(Invoice invoice, Action<string, string> unsupported) {
        var reasons = new Dictionary<(string Code, decimal? Rate, bool IsCode), (string Value, string Path)>();

        void Check(InvoiceTaxCategory? category, string path) {
            if (category == null) return;
            Compare(category.ExemptionReason, false, path + ".ExemptionReason");
            Compare(category.ExemptionReasonCode, true, path + ".ExemptionReasonCode");

            void Compare(string? value, bool isCode, string valuePath) {
                if (value == null) return;
                var key = (category.Code, InvoiceCalculator.NormalizeRate(category), isCode);
                if (!reasons.TryGetValue(key, out (string Value, string Path) first)) {
                    reasons.Add(key, (value, valuePath));
                } else if (!string.Equals(first.Value, value, StringComparison.Ordinal)) {
                    string field = isCode ? "exemption reason code" : "exemption reason";
                    unsupported(valuePath, "The target has one " + field + " for VAT category '" + category.Code +
                        "' and rate '" + (category.Rate?.ToString(System.Globalization.CultureInfo.InvariantCulture) ?? "absent") +
                        "'; value '" + value + "' conflicts with '" + first.Value + "' from " + first.Path + ".");
                }
            }
        }

        for (int index = 0; index < invoice.Lines.Count; index++) Check(invoice.Lines[index].Tax, "Lines[" + index + "].Tax");
        for (int index = 0; index < invoice.AllowancesAndCharges.Count; index++)
            Check(invoice.AllowancesAndCharges[index].Tax, "AllowancesAndCharges[" + index + "].Tax");
        for (int index = 0; index < invoice.DeclaredTaxes.Count; index++) Check(invoice.DeclaredTaxes[index].Category, "DeclaredTaxes[" + index + "].Category");
    }
}
