using Shared = OfficeIMO.Internal.Invoicing;

namespace OfficeIMO.Invoicing;

/// <summary>Canonical invoice identifiers. Profile inspection and PDF metadata use the same internal catalogue.</summary>
public static class InvoiceProfiles {
    /// <summary>EN 16931 guideline/customization identifier.</summary>
    public const string En16931Guideline = Shared.InvoiceProfiles.En16931Guideline;
    /// <summary>XRechnung 3.0 guideline/customization identifier.</summary>
    public const string XRechnung30Guideline = Shared.InvoiceProfiles.XRechnung30Guideline;
    /// <summary>Peppol BIS Billing 3.0 customization identifier.</summary>
    public const string PeppolBis30Guideline = Shared.InvoiceProfiles.PeppolBis30Guideline;
    /// <summary>French EXTENDED-CTC-FR guideline/customization identifier.</summary>
    public const string ExtendedCtcFrGuideline = Shared.InvoiceProfiles.ExtendedCtcFrGuideline;

    /// <summary>Returns the canonical XML identifier. Release-specific rule validation is separate.</summary>
    public static string GetGuidelineId(InvoiceProfile profile) => Shared.InvoiceProfiles.GetGuidelineId((Shared.InvoiceProfile)profile);

    /// <summary>Resolves supported identifiers without accepting arbitrary suffixes.</summary>
    public static bool TryFromGuidelineId(string? identifier, out InvoiceProfile profile) {
        bool recognized = Shared.InvoiceProfiles.TryFromGuidelineId(identifier, out Shared.InvoiceProfile shared);
        profile = (InvoiceProfile)shared;
        return recognized;
    }

    /// <summary>Returns canonical Factur-X XMP spelling. Peppol is an XML-only profile.</summary>
    public static string GetXmpConformanceLevel(InvoiceProfile profile) => Shared.InvoiceProfiles.GetXmpConformanceLevel((Shared.InvoiceProfile)profile);

    /// <summary>Parses user labels; callers must emit the canonical spelling returned by <see cref="GetXmpConformanceLevel"/>.</summary>
    public static bool TryFromXmpConformanceLevel(string? value, out InvoiceProfile profile) {
        bool recognized = Shared.InvoiceProfiles.TryFromXmpConformanceLevel(value, out Shared.InvoiceProfile shared);
        profile = (InvoiceProfile)shared;
        return recognized;
    }
}
