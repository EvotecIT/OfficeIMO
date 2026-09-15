# OfficeIMO.Invoicing benchmarks

This permanent suite measures four separate operations over the same deterministic 25-line Factur-X EN 16931 invoice:

- XML read parses existing bytes into the semantic model.
- XML write serializes the existing semantic model into new bytes.
- Rules validation runs the pinned Factur-X schema and business rules through Saxon.
- PDF generation renders an already captured invoice snapshot, including its exact XML attachment and multilingual fonts.

Every setup performs correctness checks outside measurement. XML writes must reproduce the canonical corpus bytes, and parsed models must serialize back to those same bytes so seller, buyer, payment, line, VAT, and total semantics are all covered. Validation must pass the pinned schema and business rules, while PDF output must contain the exact XML attachment and representative visible mixed-script text. The rules lane intentionally includes the external Java process in elapsed time; managed allocation measures the .NET benchmark process and does not claim Java heap usage.

## BenchmarkDotNet

Configure the same pinned authority paths used by `Build/Test-InvoicingStandards.ps1`, build Release, and start with a dry run:

```powershell
dotnet run --project OfficeIMO.Invoicing.Benchmarks -c Release -f net10.0 -- --filter "*" --job Dry --noOverwrite
```

Run one operation at a time for measured evidence:

```powershell
dotnet run --project OfficeIMO.Invoicing.Benchmarks -c Release -f net10.0 -- --filter "*InvoiceXmlReadBenchmarks*" --noOverwrite
dotnet run --project OfficeIMO.Invoicing.Benchmarks -c Release -f net10.0 -- --filter "*InvoiceXmlWriteBenchmarks*" --noOverwrite
dotnet run --project OfficeIMO.Invoicing.Benchmarks -c Release -f net10.0 -- --filter "*InvoiceRulesValidationBenchmarks*" --noOverwrite
dotnet run --project OfficeIMO.Invoicing.Benchmarks -c Release -f net10.0 -- --filter "*InvoicePdfGenerationBenchmarks*" --noOverwrite
```

## Cross-platform budgets

`invoice-performance-budgets.json` owns separate Windows, Linux, and macOS elapsed/allocation ceilings. The deterministic verifier measures each operation independently, with one warmup and five measured samples by default:

```powershell
dotnet run --project OfficeIMO.Invoicing.Benchmarks -c Release -f net10.0 -- --verify-budgets --output artifacts/invoicing-benchmarks/windows.json
```

Run the same command on each target OS. Treat the ceilings as regression guards, not universal speed claims: compare complete environment metadata and BenchmarkDotNet distributions before drawing performance conclusions.
