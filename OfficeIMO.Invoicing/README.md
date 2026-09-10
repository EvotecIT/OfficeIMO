# OfficeIMO.Invoicing

Shared electronic invoice contracts for .NET. Profile inspection reads the declared
CII guideline or UBL customization identifier with namespace-aware, bounded XML
parsing. It does not establish schema or business-rule compliance.

```csharp
using OfficeIMO.Invoicing;

var declaration = InvoiceProfileDeclaration.Read(File.ReadAllBytes("invoice.xml"));
Console.WriteLine(declaration.GuidelineId);
if (declaration.Profile is InvoiceProfile profile && profile != InvoiceProfile.PeppolBis)
    Console.WriteLine(InvoiceProfiles.GetXmpConformanceLevel(profile));
```

The catalogue distinguishes MINIMUM, BASIC WL, BASIC, EN 16931, EXTENDED,
XRechnung 3.0, and Peppol BIS Billing 3.0. Unknown identifiers are returned with a
null profile; missing, ambiguous, malformed, or oversized input is rejected.
Peppol BIS does not define Factur-X XMP metadata.
