# OfficeIMO.Invoicing.Validation

Validate exact CII or UBL invoice bytes against an explicit schema and business-rule
release. The package runs on .NET 8 and .NET 10. It makes no network requests.

## Install

```powershell
dotnet add package OfficeIMO.Invoicing.Validation --version 3.4.4
```

## Build from source

From the repository root, pack both projects, then add the validation package to
your application. NuGet resolves its invoice-model dependency from the same feed:

```powershell
dotnet pack OfficeIMO.Invoicing/OfficeIMO.Invoicing.csproj -c Release -o artifacts/invoice-feed
dotnet pack OfficeIMO.Invoicing.Validation/OfficeIMO.Invoicing.Validation.csproj -c Release -o artifacts/invoice-feed
dotnet add path/to/Application.csproj package OfficeIMO.Invoicing.Validation --source artifacts/invoice-feed
```

PDF generation is independent of this package. Java and Saxon are required only
when running the configured business-rule stage.

## Validate Polish FA(3)

`Fa3SchemaValidator` validates exact FA(3) bytes against the official schema
published on 25 June 2025, version 1-0E. Download these four files into one local
directory, naming the main schema `FA3.xsd`:

| File | Official source |
| --- | --- |
| FA3.xsd | [FA(3) schema](https://crd.gov.pl/wzor/2025/06/25/13775/schemat.xsd) |
| StrukturyDanych_v10-0E.xsd | [Data structures](https://crd.gov.pl/xml/schematy/dziedzinowe/mf/2022/01/05/eD/DefinicjeTypy/StrukturyDanych_v10-0E.xsd) |
| ElementarneTypyDanych_v10-0E.xsd | [Elementary data types](https://crd.gov.pl/xml/schematy/dziedzinowe/mf/2022/01/05/eD/DefinicjeTypy/ElementarneTypyDanych_v10-0E.xsd) |
| KodyKrajow_v10-0E.xsd | [Country codes](https://crd.gov.pl/xml/schematy/dziedzinowe/mf/2022/01/05/eD/DefinicjeTypy/KodyKrajow_v10-0E.xsd) |

```csharp
var schemas = Fa3SchemaBundle.LoadDirectory("schemas/fa3");
var fa3Validator = new Fa3SchemaValidator(schemas);
var result = fa3Validator.Validate(File.ReadAllBytes("polish-invoice.xml"));
Console.WriteLine($"{result.Status}: {result.InputSha256}");
foreach (var diagnostic in result.Diagnostics)
    Console.WriteLine($"{diagnostic.Code}: {diagnostic.Message}");
```

The bundle verifies each SHA-256 and snapshots the files before validation. The
main schema digest is
`B646B6B525F51ADF1BB2545F111FC8CA6E7AA6DD2F98948F1667D3695C06D958`.
Imports resolve only to these verified local snapshots; document-supplied schema
locations and DTDs cannot trigger downloads. Input is limited to 16 MiB and depth
128, with at most 1,000 reported diagnostics. Reports identify the input length,
input hash and schema hash. No Java or Saxon runtime is required for this lane.

The checked-in corpus contains all 26 independent examples from the
[official FA(3) sample archive](https://ksef.podatki.gov.pl/media/e5cia0ey/przykladowe-pliki-dla-struktury-logicznej-e-faktury-fa-3.zip),
including corrections, advances, settlements, simplified invoices and foreign
currencies. XSD success establishes structure only. It does not certify Polish
tax treatment, authorize submission or establish KSeF acceptance. This national
format has its own schema contract rather than an EN 16931 profile declaration.

## Validate invoice XML

The configured stages are:

| Release | Syntax | Schema and rules |
| --- | --- | --- |
| EN 16931 1.3.16 | CII or UBL | D16B CII / UBL 2.1 XSD and EN 16931 Schematron |
| Factur-X 1.09.2 / ZUGFeRD 2.5.2 | CII D22B | Profile-specific XSD and rules for MINIMUM, BASIC WL, BASIC, EN 16931 and EXTENDED |
| XRechnung 3.0.2, configuration 2026-08-31 | CII or UBL | XSD, EN 16931, XRechnung and the configuration's severity overrides |
| Peppol BIS Billing 3.0.21 | UBL | UBL 2.1 XSD, EN 16931 and May 2026 Peppol rules |

Download the official [KoSIT rule archive](https://github.com/itplr-kosit/validator-configuration-xrechnung/releases/download/v2026-08-31/xrechnung-3.0.2-validator-configuration-2026-08-31.zip)
and the pinned Apache-2.0
[Factur-X/ZUGFeRD validator configuration](https://github.com/LandrixSoftware/validator-configuration-zugferd/tree/8399f5459df34b2af9cae0d613870649b3f17a58),
and extract the official [SaxonJ-HE 12.10 distribution](https://downloads.saxonica.com/SaxonJ/HE/12/SaxonHE12-10J.zip)
to a local tools directory. Keep Saxon's `lib` directory next to its main JAR and
provide a Java runtime. These external artifacts retain their upstream licenses;
OfficeIMO does not redistribute the authority rules or Java runtime.

The runner verifies the main JAR and its `xmlresolver-5.3.3.jar`,
`xmlresolver-5.3.3-data.jar` and `jline-2.14.6.jar` companions at construction and
before execution. Each invocation runs from a private copy of those verified
bytes, excluding installation-directory aliases and loose classes. Reports
identify all four JAR hashes. The configured Java executable remains supplied
and managed by the caller.

```csharp
using OfficeIMO.Invoicing.Validation;

var bundle = InvoiceRuleBundle.Load(
    "xrechnung-3.0.2-validator-configuration-2026-08-31.zip",
    facturXArchivePath: "validator-configuration-zugferd.zip");
var runner = new SaxonInvoiceRulesRunner("tools/saxon/saxon-he-12.10.jar");
var validator = new InvoiceValidator(bundle, runner);
var report = await validator.ValidateAsync(
    File.ReadAllBytes("invoice.xml"),
    InvoiceSpecificationRelease.XRechnung_3_0_2_2026_08_31);

Console.WriteLine($"{report.IsValid}: {report.Sha256}");
foreach (var diagnostic in report.Diagnostics)
    Console.WriteLine($"{diagnostic.Code}: {diagnostic.Message}");
```

For Peppol, also supply the official
[3.0.21 Schematron source](https://raw.githubusercontent.com/OpenPEPPOL/peppol-bis-invoice-3/806866bd2bd91d7e9623b68f08164e8fbe9e67a0/rules/sch/PEPPOL-EN16931-UBL.sch)
as the second argument to `InvoiceRuleBundle.Load`. The source is verified by
SHA-256 and compiled locally with the included MIT-licensed Schematron compiler.
The Peppol download names an immutable official source commit; a changed artifact is rejected by its pinned hash.

Factur-X authoring and validation use release `FacturX_1_09_2_Zugferd_2_5_2`.
The validator archive is pinned by commit and SHA-256. Independent producer
fixtures from AdVitam/facturx cover all five profiles. The checked-in corpus
manifest also pins KoSIT XRechnung CII and UBL fixture hashes. XRechnung is the
only national CIUS exposed by the validator; recognizing another guideline does
not create a validation contract without its own syntax, rules and corpus.

`IsValid` is true only when both XSD and business rules passed. Without a runner,
business rules are `NotRun`. Engine failures are `Failed`, and content failures
are `Invalid`. Reports include exact input length, SHA-256, release, authority
artifact hash, immutable source, source commit when applicable, and runner identity.
Factur-X reports identify the Factur-X bundle rather than the KoSIT bundle used by
the EN 16931/XRechnung lanes. Revalidate after any edit to the XML.
Runner identity is recorded after an invoice-rule process starts, including when
that process fails. Startup failures and compiler-only executions leave it empty.
Reports contain at most 999 detailed diagnostics plus a summary of excess
diagnostics with their highest severity, preserving the
distinction between invalid invoice content and a failed validator.

The pinned EN 16931 artifacts differ for zero-rated IGIC (category `L`): CII
rules BR-AF-05/06/07 require a positive rate, while the UBL rules permit zero.
The semantic model accepts a non-negative IGIC rate; the standards report exposes
the selected syntax's rule result. A zero-rate CII invoice therefore cannot receive
a passing report from this pinned release.

This validates XML. Factur-X/ZUGFeRD PDF/A, XMP, attachment relationships and visible
invoice content require the separate PDF artifact checks. This package does not
certify delivery over the Peppol network or the legality of a transaction.

Each rule-engine invocation uses a separate temporary workspace. On Unix, its
permissions restrict access to the current user before invoice content is written.
The workspace is removed after completion, failure, cancellation, or timeout.

## Run the standards checks

With Java and the .NET 10 SDK available, run:

```powershell
./Build/Test-InvoicingStandards.ps1
```

The command downloads hash-pinned authority artifacts, the independent producer
corpus and Saxon into
`artifacts/invoicing-standards`, then validates generated documents and
independent fixtures. It validates all five generated and independent Factur-X
profiles plus the declared XRechnung and Peppol lanes. `-OutputPath` selects a different local artifact directory;
`-JavaExecutable` selects an existing Java installation. The artifacts remain
there for reuse and inspection.
