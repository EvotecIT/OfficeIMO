# Bibliography validation

Run the normal correctness suite with `dotnet test OfficeIMO.Bibliography.Tests -c Release -f net8.0`. It exercises parsing, serialization, source fidelity, encoding, resource limits, and cancellation. The project also supports .NET 10 and Windows .NET Framework 4.7.2.

Seven resource tests use small inputs in normal runs while retaining their rejection, output, or offset assertions. To run the original large workloads with managed-allocation budgets, rebuild and run the selected cases explicitly:

```powershell
dotnet test OfficeIMO.Bibliography.Tests -c Release -f net8.0 `
    -p:BibliographyPerformanceEvidence=true --filter "Category=ResourcePerformanceEvidence"
```

The Bibliography Resource Performance Evidence workflow runs that command on Windows, Linux, and macOS. Do not pass `--no-build` when changing the property: the workload sizes and allocation assertions are selected at compilation. Return to the ordinary command without the property to rebuild the correctness suite. Compare allocation results under the same runtime and workload; they do not replace the normal correctness assertions.
