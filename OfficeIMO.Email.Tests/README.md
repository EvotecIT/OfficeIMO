# Email validation

Run ordinary correctness tests with:

```powershell
dotnet test OfficeIMO.Email.Tests -c Release -f net8.0 --filter "Category!=Performance"
```

The MIME, mbox, MSG resource-limit, PST reopen, conversion, and deduplication tests run in the normal suite. The project also supports .NET 10 and Windows .NET Framework 4.7.2.

Run the large MIME/mbox/MSG allocation workloads and PST scale measurements explicitly:

```powershell
dotnet test OfficeIMO.Email.Tests -c Release -f net8.0 --filter "Category=Performance"
```

The Email Performance Evidence workflow runs that measurement lane on Windows, Linux, and macOS and retains its results. It checks observed elapsed time and managed-memory envelopes; compare failures under the same workload and runtime. An unfiltered local `dotnet test` includes both lanes.
