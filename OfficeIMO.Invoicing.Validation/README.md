# OfficeIMO.Invoicing.Validation

Validate exact CII or UBL invoice bytes against an explicit schema and business-rule
release. The package runs on .NET 8 and .NET 10. It makes no network requests.

The configured stages are:

| Release | Syntax | Schema and rules |
| --- | --- | --- |
| EN 16931 1.3.16 | CII or UBL | D16B CII / UBL 2.1 XSD and EN 16931 Schematron |
| XRechnung 3.0.2, configuration 2026-08-31 | CII or UBL | XSD, EN 16931, XRechnung and the configuration's severity overrides |
| Peppol BIS Billing 3.0.21 | UBL | UBL 2.1 XSD, EN 16931 and May 2026 Peppol rules |

Download the official [KoSIT rule archive](https://github.com/itplr-kosit/validator-configuration-xrechnung/releases/download/v2026-08-31/xrechnung-3.0.2-validator-configuration-2026-08-31.zip)
and extract the official [SaxonJ-HE 12.10 distribution](https://downloads.saxonica.com/SaxonJ/HE/12/SaxonHE12-10J.zip)
to a local tools directory. Keep Saxon's `lib` directory next to its main JAR and
provide a Java runtime. These external artifacts retain their upstream licenses;
OfficeIMO does not redistribute the authority rules or Java runtime.

```csharp
using OfficeIMO.Invoicing.Validation;

var bundle = InvoiceRuleBundle.Load("xrechnung-3.0.2-validator-configuration-2026-08-31.zip");
var runner = new SaxonInvoiceRulesRunner("tools/saxon/saxon-he-12.10.jar");
var validator = new InvoiceValidator(bundle, runner);
var report = await validator.ValidateAsync(
    File.ReadAllBytes("invoice.xml"),
    InvoiceRulesRelease.XRechnung_3_0_2_2026_08_31);

Console.WriteLine($"{report.IsValid}: {report.Sha256}");
foreach (var diagnostic in report.Diagnostics)
    Console.WriteLine($"{diagnostic.Code}: {diagnostic.Message}");
```

For Peppol, also supply the official
[3.0.21 Schematron source](https://docs.peppol.eu/poacc/billing/3.0/files/PEPPOL-EN16931-UBL.sch)
as the second argument to `InvoiceRuleBundle.Load`. The source is verified by
SHA-256 and compiled locally with the included MIT-licensed Schematron compiler.
If the upstream download changes to a newer release, the pinned hash rejects it.

`IsValid` is true only when both XSD and business rules passed. Without a runner,
business rules are `NotRun`. Engine failures are `Failed`, and content failures
are `Invalid`. Reports include exact input length, SHA-256, release, authority
artifact hashes and runner identity. Revalidate after any edit to the XML.

This validates XML. Factur-X/ZUGFeRD PDF/A, XMP, attachment relationships and visible
invoice content require the separate PDF artifact checks. This package does not
certify delivery over the Peppol network or the legality of a transaction.

## Run the standards checks

With Java and the .NET 10 SDK available, run:

```powershell
./Build/Test-InvoicingStandards.ps1
```

The command downloads hash-pinned authority artifacts and Saxon into
`artifacts/invoicing-standards`, then validates generated documents and
independent fixtures. `-OutputPath` selects a different local artifact directory;
`-JavaExecutable` selects an existing Java installation. The artifacts remain
there for reuse and inspection.
