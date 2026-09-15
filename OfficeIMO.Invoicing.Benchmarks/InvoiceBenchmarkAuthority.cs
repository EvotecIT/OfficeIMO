using OfficeIMO.Invoicing.Validation;

namespace OfficeIMO.Invoicing.Benchmarks;

internal static class InvoiceBenchmarkAuthority {
    internal static InvoiceValidator CreateValidator() {
        string xRechnung = Required("OFFICEIMO_INVOICE_RULE_BUNDLE");
        string facturX = Required("OFFICEIMO_INVOICE_FACTURX_RULE_BUNDLE");
        string saxon = Required("OFFICEIMO_INVOICE_SAXON_JAR");
        string java = Environment.GetEnvironmentVariable("OFFICEIMO_INVOICE_JAVA") ?? "java";
        InvoiceRuleBundle bundle = InvoiceRuleBundle.Load(xRechnung, facturXArchivePath: facturX);
        return new InvoiceValidator(bundle, new SaxonInvoiceRulesRunner(saxon, java));
    }

    private static string Required(string name) {
        string? value = Environment.GetEnvironmentVariable(name);
        if (string.IsNullOrWhiteSpace(value)) throw new InvalidOperationException(name + " must identify the pinned local benchmark authority artifact.");
        return value;
    }
}
