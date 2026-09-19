namespace OfficeIMO.Invoicing;

/// <summary>Explicit, immutable specification releases used for invoice authoring and validation.</summary>
/// <remarks>There is intentionally no moving <c>Latest</c> value.</remarks>
public enum InvoiceSpecificationRelease {
    /// <summary>EN 16931 validation artifacts 1.3.16, using CII D16B or UBL 2.1.</summary>
    En16931_1_3_16,
    /// <summary>Factur-X 1.09.2 / ZUGFeRD 2.5.2, using CII D22B.</summary>
    FacturX_1_09_2_Zugferd_2_5_2,
    /// <summary>XRechnung 3.0.2 validator configuration dated 2026-08-31.</summary>
    XRechnung_3_0_2_2026_08_31,
    /// <summary>Peppol BIS Billing 3.0.21 published in May 2026, using UBL 2.1.</summary>
    PeppolBis_3_0_21
}

internal static class InvoiceSpecificationContracts {
    internal static void Validate(InvoiceSpecificationRelease release, InvoiceSyntax syntax, InvoiceProfile profile) {
        if (release < InvoiceSpecificationRelease.En16931_1_3_16 || release > InvoiceSpecificationRelease.PeppolBis_3_0_21)
            throw new ArgumentOutOfRangeException(nameof(release));
        if (syntax < InvoiceSyntax.Cii || syntax > InvoiceSyntax.Ubl) throw new ArgumentOutOfRangeException(nameof(syntax));
        if (profile < InvoiceProfile.Minimum || profile > InvoiceProfile.ExtendedCtcFr) throw new ArgumentOutOfRangeException(nameof(profile));

        bool supported = release switch {
            InvoiceSpecificationRelease.En16931_1_3_16 => profile == InvoiceProfile.En16931,
            InvoiceSpecificationRelease.FacturX_1_09_2_Zugferd_2_5_2 => syntax == InvoiceSyntax.Cii &&
                profile is InvoiceProfile.Minimum or InvoiceProfile.BasicWithoutLines or InvoiceProfile.Basic or InvoiceProfile.En16931 or InvoiceProfile.Extended,
            InvoiceSpecificationRelease.XRechnung_3_0_2_2026_08_31 => profile == InvoiceProfile.XRechnung,
            InvoiceSpecificationRelease.PeppolBis_3_0_21 => syntax == InvoiceSyntax.Ubl && profile == InvoiceProfile.PeppolBis,
            _ => false
        };
        if (!supported)
            throw new NotSupportedException($"Release '{release}' does not define the '{profile}' profile in the '{syntax}' syntax.");
    }
}
