namespace OfficeIMO.Invoicing.Validation.Tests;

internal static class InvoiceTestContracts {
    internal static InvoiceXmlOptions En16931(InvoiceSyntax syntax = InvoiceSyntax.Cii) =>
        new(InvoiceSpecificationRelease.En16931_1_3_16, syntax, InvoiceProfile.En16931);
    internal static InvoiceXmlOptions XRechnung(InvoiceSyntax syntax) =>
        new(InvoiceSpecificationRelease.XRechnung_3_0_2_2026_08_31, syntax, InvoiceProfile.XRechnung);
    internal static InvoiceXmlOptions Peppol() =>
        new(InvoiceSpecificationRelease.PeppolBis_3_0_21, InvoiceSyntax.Ubl, InvoiceProfile.PeppolBis);
    internal static InvoiceXmlOptions FacturX(InvoiceProfile profile, InvoiceProjectionPolicy policy = InvoiceProjectionPolicy.RejectDataLoss) =>
        new(InvoiceSpecificationRelease.FacturX_1_09_2_Zugferd_2_5_2, InvoiceSyntax.Cii, profile, policy);
    internal static InvoiceXmlOptions For(InvoiceSyntax syntax, InvoiceProfile profile) => profile switch {
        InvoiceProfile.En16931 => En16931(syntax),
        InvoiceProfile.XRechnung => XRechnung(syntax),
        InvoiceProfile.PeppolBis when syntax == InvoiceSyntax.Ubl => Peppol(),
        _ => FacturX(profile)
    };
}
