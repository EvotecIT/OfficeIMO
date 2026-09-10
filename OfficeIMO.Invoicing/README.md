# OfficeIMO.Invoicing

Shared electronic invoice contracts for .NET. Profile inspection reads the declared
CII guideline or UBL customization identifier with namespace-aware, bounded XML
parsing. It does not establish schema or business-rule compliance.

## Build and install locally

From the repository root, pack the project into a local feed, then add that
package to your application:

```powershell
dotnet pack OfficeIMO.Invoicing/OfficeIMO.Invoicing.csproj -c Release -o artifacts/invoice-feed
dotnet add path/to/Application.csproj package OfficeIMO.Invoicing --source artifacts/invoice-feed
```

This package does not depend on a PDF engine or an external validation runtime.

## Inspect a declaration

```csharp
using OfficeIMO.Invoicing;

var declaration = InvoiceProfileDeclaration.Read(File.ReadAllBytes("invoice.xml"));
Console.WriteLine(declaration.GuidelineId);
if (declaration.Profile is InvoiceProfile profile && profile != InvoiceProfile.PeppolBis)
    Console.WriteLine(InvoiceProfiles.GetXmpConformanceLevel(profile));
```

The catalogue distinguishes MINIMUM, BASIC WL, BASIC, EN 16931, EXTENDED,
XRechnung 3.0, EXTENDED-CTC-FR, and Peppol BIS Billing 3.0. Unknown identifiers are returned with a
null profile; missing, ambiguous, malformed, or oversized input is rejected.
Peppol BIS does not define Factur-X XMP metadata.

The French profile identifier follows [AFNOR XP Z12-012](https://www.impots.gouv.fr/sites/default/files/media/1_metier/2_professionnel/EV/2_gestion/290_facturation_electronique/specification_externes_b2b/afnor/norme-afnor-factures.pdf).
Recognition does not establish French business-rule coverage.
