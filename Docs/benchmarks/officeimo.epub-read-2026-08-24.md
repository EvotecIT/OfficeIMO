# EPUB open/read comparison evidence

This 2026-08-24 evidence compares complete EPUB 3 open/read workflows through
OfficeIMO.Epub and VersOne.Epub 3.3.6 with HtmlAgilityPack 1.12.4. Both readers
receive the same deterministic in-memory package, load metadata and content,
extract normalized visible text, and enumerate every chapter in spine order.

The performance acceptance bands used for this work are:

- at or below 2x the comparison implementation in elapsed time, managed
  allocation, and measured memory: contender range, though further improvement
  may still be useful;
- above 2x through 5x in any dimension: material remediation gap;
- above 5x in any dimension: unacceptable unless the observable contracts
  differ and the result is reported as a non-equivalent diagnostic.

A 40x ratio is an incident threshold, never a success boundary.

## Equivalent-work result

The table is a BenchmarkDotNet default job from clean source commit
`81412ef012198297c5a6b8a31cdf0b16cc91288a` on Windows 11
10.0.26200.9168, an AMD Ryzen 9 9950X3D2 with 16 physical cores, .NET SDK
10.0.111, and .NET 8.0.30 x64. Ratios are OfficeIMO divided by the comparison
implementation, so lower is better.

| Scale | Shared input | Chapters | OfficeIMO mean | VersOne mean | Time ratio | OfficeIMO allocation | VersOne allocation | Allocation ratio |
| --- | ---: | ---: | ---: | ---: | ---: | ---: | ---: | ---: |
| Small | 6,442 bytes | 8 | 153.1 us | 581.5 us | 0.26x | 471.26 KB | 1,676.32 KB | 0.28x |
| Normal | 55,829 bytes | 48 | 2,141.7 us | 5,790.5 us | 0.37x | 5,457.64 KB | 22,967.03 KB | 0.24x |

OfficeIMO is 3.80x faster on the small package and 2.70x faster on the normal
package. The comparison allocates 3.56x and 4.21x as much managed memory,
respectively. The VersOne timing distribution was multimodal on this host, but
its fastest observed operations remained slower than OfficeIMO's means; the
conclusion does not depend on a narrow margin.

## Retained and peak memory

An independent process-isolated run from the same clean commit used three
fresh child processes per reader and scale. The retention phase keeps
equivalent metadata, raw XHTML, normalized text, and ordered chapter
projections alive. The parent rejects input-size or semantic-fingerprint
differences before reporting ratios.

| Scale | Retained managed-heap ratio | Sampled managed-heap peak ratio | Absolute process-peak ratio |
| --- | ---: | ---: | ---: |
| Small | 0.94x | 0.28x | 0.92x |
| Normal | 0.99x | 0.44x | 0.94x |

The retained semantic result is nearly the same size for both readers, while
OfficeIMO's sampled managed-heap peak is materially lower. Absolute process
peak is the working-set comparison metric because independently launched
runtimes do not begin with identical working sets.

The isolated runner also reported OfficeIMO/VersOne median elapsed ratios of
0.63x for Small and 0.87x for Normal, and allocation ratios of 0.28x and 0.24x.
BenchmarkDotNet remains the primary timing evidence because the shorter child
process batches contain more tiering, GC, and host noise.

## Output equivalence and size

Validation requires identical title, creator, language, chapter count and
order, boundary paths, raw-XHTML character count and hash, normalized visible
text character count and hash, and path hash. The validated packages contain:

| Scale | Raw XHTML | Normalized text |
| --- | ---: | ---: |
| Small | 39,148 characters | 25,764 characters |
| Normal | 737,307 characters | 498,564 characters |

OfficeIMO.Epub exposes reading and extraction rather than EPUB package
creation, so both lanes consume the same input bytes and there is no honest
output-file-size ratio to report. No runtime optimization was made for this
slice: the measured canonical owner already outperforms the equivalent
comparison and remains below the 2x ceiling on every measured dimension.

Reproduce the validation and measurements with:

```powershell
dotnet run -c Release -f net8.0 --project .\OfficeIMO.Epub.Benchmarks.Comparisons -- validate --json .benchmark-artifacts\epub\validation.json
dotnet run -c Release -f net8.0 --project .\OfficeIMO.Epub.Benchmarks.Comparisons -- --filter '*EpubReadComparisonBenchmarks*' --noOverwrite
dotnet run -c Release -f net8.0 --project .\OfficeIMO.Epub.Benchmarks.Comparisons -- --evidence --repeat 3 --json .benchmark-artifacts\epub\comparison-evidence.json
```

Raw BenchmarkDotNet and process-runner output remains local and excluded from
the repository's small committed evidence set.
