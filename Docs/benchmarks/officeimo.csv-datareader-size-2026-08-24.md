# CSV DataReader output-size evidence

This 2026-08-24 evidence measures UTF-8 output size for the existing validated
SQL-shaped `IDataReader` write comparison between OfficeIMO.CSV and
Sylvan.Data.Csv 1.4.4. Both writers receive fresh readers over the same typed
rows. Every generated output is parsed and checked field by field before its
size is accepted.

The run came from clean source commit
`a2e8c72a845e4b579bc5126356a3ba5dad22545e` on Windows 11, .NET 10.0.11 x64,
and the AMD Ryzen 9 9950X3D2 host used by the current timing evidence. Ratios
are OfficeIMO divided by Sylvan, so lower is smaller.

| Rows | Shape | OfficeIMO UTF-8 bytes | Sylvan UTF-8 bytes | Size ratio |
| ---: | --- | ---: | ---: | ---: |
| 25,000 | Mixed | 2,805,195 | 2,830,195 | 0.991x |
| 25,000 | Quoted | 3,632,912 | 3,657,912 | 0.993x |
| 25,000 | Multiline | 3,283,455 | 3,308,455 | 0.992x |
| 100,000 | Mixed | 11,289,498 | 11,389,498 | 0.991x |
| 100,000 | Quoted | 14,600,363 | 14,700,363 | 0.993x |
| 100,000 | Multiline | 13,202,538 | 13,302,538 | 0.992x |

OfficeIMO output is slightly smaller in every lane. Both writers emit the same
line-feed counts with no carriage returns, so this is not a CRLF/LF artifact.
The difference is a valid typed-value representation choice accepted by the
semantic validator, not compression. OfficeIMO's sequential and parallel
writers also produced identical UTF-8 byte counts and SHA-256 hashes in every
lane.

Reproduce with:

```powershell
dotnet run -c Release -f net10.0 --project .\OfficeIMO.CSV.Benchmarks -- --datareader-write-size-evidence --rows 25000,100000 --json .benchmark-artifacts\csv\datareader-write-size.json
```

Raw JSON remains local and ignored; the runner records commit, dirty-tree
state, runtime, OS, architecture, processor count, character and UTF-8 byte
counts, newline counts, and hashes.
