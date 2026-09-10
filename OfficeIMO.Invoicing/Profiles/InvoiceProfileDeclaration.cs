using System.Xml.Linq;
using Shared = OfficeIMO.Internal.Invoicing;

namespace OfficeIMO.Invoicing;

/// <summary>The namespace-aware invoice declaration. Identification does not establish schema or business-rule validation.</summary>
public sealed class InvoiceProfileDeclaration {
    /// <summary>Maximum input size accepted by the invoice XML reader.</summary>
    public const int MaximumXmlBytes = Shared.InvoiceProfileDeclaration.MaximumXmlBytes;

    private InvoiceProfileDeclaration(Shared.InvoiceProfileDeclaration declaration) {
        Syntax = (InvoiceSyntax)declaration.Syntax;
        GuidelineId = declaration.GuidelineId;
        Profile = declaration.Profile.HasValue ? (InvoiceProfile)declaration.Profile.Value : null;
    }

    /// <summary>Identified XML syntax.</summary>
    public InvoiceSyntax Syntax { get; }
    /// <summary>Exact trimmed guideline/customization identifier.</summary>
    public string GuidelineId { get; }
    /// <summary>Recognized profile, or null for an unsupported identifier.</summary>
    public InvoiceProfile? Profile { get; }

    /// <summary>Reads one unambiguous declaration, rejecting malformed, oversized, DTD-bearing, and ambiguous XML.</summary>
    public static InvoiceProfileDeclaration Read(byte[] xml) => new InvoiceProfileDeclaration(Shared.InvoiceProfileDeclaration.Read(xml));

    internal static InvoiceProfileDeclaration Read(XDocument document) => new InvoiceProfileDeclaration(Shared.InvoiceProfileDeclaration.Read(document));
}
