# H4/advanced-held-out static-rendering budget calibration

These reports calibrate the H4/advanced-held-out OfficeIMO static-rendering gate on Windows,
Ubuntu Linux, and macOS. All three runs used clean source at commit
`74d80d8ea32a65836bd35543264ac0a6620b7b66`, corpus manifest SHA-256
`496f78d459bfd7836987541925f3d6f4b26c87512cd03f319ca99bc4057f67a8`, and
three warmed iterations.

Each iteration rendered all eight held-out cases into 76 outputs: screen PNG/SVG,
print PDF/PNG/SVG, and screen-to-page PDF/PNG/SVG. The worker also verified expected
print page counts, declared text markers, the absence of error diagnostics,
within-run output determinism, and prompt cancellation through an asynchronous
resource boundary.

| Platform | Runtime and architecture | Cold process | Worst warm iteration | Cold allocation | Worst warm allocation | Peak working set | Output per iteration | Cancellation |
| --- | --- | ---: | ---: | ---: | ---: | ---: | ---: | ---: |
| Windows 10.0.26200 | .NET 10.0.12, x64 | 8,279.3 ms | 2,469.6 ms | 2,016,045,712 B | 957,406,888 B | 1,103,282,176 B | 2,361,301 B | 3.019 ms |
| Ubuntu 24.04.3 LTS | .NET 10.0.12, x64 | 7,294.0 ms | 2,678.5 ms | 841,351,488 B | 771,755,392 B | 295,354,368 B | 2,013,013 B | 5.692 ms |
| macOS 27.0.0 | .NET 10.0.12, arm64 | 5,468.8 ms | 1,821.5 ms | 988,698,720 B | 829,428,704 B | 460,210,176 B | 2,117,654 B | 1.787 ms |

The checked-in ceilings round each observed value upward with room for ordinary
host variation while remaining close enough to detect material regressions. Output
size has a shared 3 MB ceiling because installed font coverage changes the embedded
font subset between platforms. Cancellation has a 250 ms ceiling to tolerate host
scheduling noise while preserving a prompt-cancellation contract.

The raw reports are [Windows](windows.json), [Linux](linux.json), and
[macOS](macos.json). Run the command in
[`Build/HtmlStaticBudget`](../../../../HtmlStaticBudget/README.md) to reproduce the gate.
