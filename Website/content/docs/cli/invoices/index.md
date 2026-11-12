---
title: "Electronic Invoices from the CLI"
description: "Inspect, validate, edit, convert and render CII and UBL electronic invoices with officeimo invoice, including Factur-X and ZUGFeRD hybrid PDFs, batches and pinned standards validation."
layout: docs
---

`officeimo invoice` works with CII and UBL XML invoices: EN 16931, Factur-X/ZUGFeRD, XRechnung and Peppol BIS. Every command returns versioned JSON that keeps model, mapping, target and standards results separate.

## Inspect and validate

```shell
officeimo invoice inspect invoice.xml
officeimo invoice validate invoice.xml
```

`inspect` can succeed while it reports invalid model data. `validate` returns exit code `1` when the invoice has errors, which makes it the command to gate a pipeline on. Schema and business-rule stages stay `NotRun` unless you configure [standards validation](#standards-validation).

## Convert to another syntax or profile

Conversion and rendering need an explicit target: release, syntax and profile.

```shell
officeimo invoice convert invoice.xml --output invoice.ubl.xml --release En16931_1_3_16 --syntax Ubl --profile En16931
```

| Option | Values |
|---|---|
| `--release` | `En16931_1_3_16`, `FacturX_1_09_2_Zugferd_2_5_2`, `XRechnung_3_0_2_2026_08_31`, `PeppolBis_3_0_21` |
| `--syntax` | `Cii`, `Ubl` |
| `--profile` | `Minimum`, `BasicWithoutLines`, `Basic`, `En16931`, `Extended`, `XRechnung`, `PeppolBis` |

Unmapped source data and target fields that cannot be represented block the conversion. `--allow-profile-loss` accepts only the documented reductions of lower Factur-X profiles and reports them as warnings.

## Render a PDF or a hybrid invoice

`render` creates a separate presentation PDF. `hybrid` creates a Factur-X/ZUGFeRD PDF with the captured CII XML embedded. Both require CII:

```shell
officeimo invoice render invoice.xml --output invoice.pdf --release En16931_1_3_16 --syntax Cii --profile En16931
officeimo invoice hybrid invoice.xml --output invoice.pdf --release FacturX_1_09_2_Zugferd_2_5_2 --syntax Cii --profile En16931 --language pl-PL --font ./fonts/InvoiceFont.ttf --columns Item,Quantity,Unit,NetPrice,Vat,NetAmount --unit-display Description --payment-display Description --modern --compact-details --page-identity
```

`--language` sets the label language, `--font` embeds a TrueType or OpenType font, and `--columns` picks two to eight line-item columns, including `Item` and `NetAmount`. `--unit-display` and `--payment-display` take `Code`, `Description` or `CodeAndDescription`. XML standards validation does not certify the PDF's archival conformance.

## Edit header fields

```shell
officeimo invoice edit invoice.xml --output invoice.edited.xml --number INV-002 --issue-date 2026-09-30 --buyer-reference BUYER-002
```

`edit` replaces existing, unique plaintext fields and keeps the rest of the XML, including extensions. It accepts `--number`, `--issue-date`, `--due-date`, `--buyer-reference` and `--payment-reference`, with dates as `yyyy-MM-dd`. Signed XML is refused. The result can still report model or mapping errors, so validate the edited file before sending it.

## Batches

```shell
officeimo invoice batch validate first.xml second.xml --max-items 100 --stop-on-failure
officeimo invoice batch convert *.xml --output-directory ubl --release En16931_1_3_16 --syntax Ubl --profile En16931
officeimo invoice batch edit *.xml --output-directory edited --payment-reference PAY-2026-11
```

Batch output is named `<input-stem>.invoice.xml` or `<input-stem>.invoice.pdf`. All inputs and destinations are checked before any output is written, and colliding or existing destinations fail the batch up front. Each item is published on its own. Defaults allow 256 inputs and 64 MiB of combined input and returned artifacts; change the totals with `--max-items`, `--max-input-bytes` and `--max-output-bytes`. Each XML input is limited to 16 MiB.

Invoice outputs are created atomically and never replace an existing file; there is no `--force` for this area.

## Standards validation

Run schema and business rules against the exact source bytes with pinned rule artifacts and a Saxon runner:

```shell
officeimo invoice validate invoice.xml --standards-release En16931_1_3_16 --rule-bundle ./rules/xrechnung.zip --saxon-jar ./saxon/saxon-he-12.10.jar
```

`--peppol-rules` and `--facturx-rules` supply the other pinned rule sets, and `--java` selects the Java executable. The [standards validator guide](https://github.com/EvotecIT/OfficeIMO/blob/master/OfficeIMO.Invoicing.Validation/README.md) lists the pinned downloads and runtime requirements.

## Related

- [Electronic invoices](/docs/invoicing/) for the .NET API
- [Scripting and CI](/docs/cli/scripting-and-ci/)
