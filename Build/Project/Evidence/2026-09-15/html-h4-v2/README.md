# H4/v2 static-rendering budget calibration

These reports calibrate the H4/v2 OfficeIMO static-rendering gate on Windows,
Ubuntu Linux, and macOS. All three runs used clean source at commit
`3808783dccaaeda9d0af34511e90c5e8af4e959c`, corpus manifest SHA-256
`60af2aab907af324bd1820a247a92b2b26f3e8445f235e06d50b10ac244375c7`, and
three warmed iterations.

Each iteration rendered all eight held-out cases into 76 outputs: screen PNG/SVG,
print PDF/PNG/SVG, and screen-to-page PDF/PNG/SVG. The worker also verified expected
print page counts, declared text markers, the absence of error diagnostics,
within-run output determinism, and prompt cancellation through an asynchronous
resource boundary.

| Platform | Runtime and architecture | Cold process | Worst warm iteration | Cold allocation | Worst warm allocation | Peak working set | Output per iteration | Cancellation |
| --- | --- | ---: | ---: | ---: | ---: | ---: | ---: | ---: |
| Windows 10.0.26200 | .NET 10.0.12, x64 | 7,897.8 ms | 2,664.3 ms | 2,015,442,784 B | 957,124,080 B | 1,092,034,560 B | 2,361,301 B | 4.511 ms |
| Ubuntu 24.04.3 LTS | .NET 10.0.12, x64 | 7,502.1 ms | 2,568.7 ms | 842,072,096 B | 771,512,640 B | 294,420,480 B | 2,013,013 B | 4.489 ms |
| macOS 27.0.0 | .NET 10.0.11, arm64 | 4,870.6 ms | 1,861.0 ms | 988,410,368 B | 829,178,512 B | 462,356,480 B | 2,117,654 B | 1.701 ms |

The checked-in ceilings round each observed value upward with room for ordinary
host variation while remaining close enough to detect material regressions. Output
size has a shared 3 MB ceiling because installed font coverage changes the embedded
font subset between platforms. Cancellation has a 250 ms ceiling to tolerate host
scheduling noise while preserving a prompt-cancellation contract.

The raw reports are [Windows](windows.json), [Linux](linux.json), and
[macOS](macos.json). Run the command in
[`Build/HtmlStaticBudget`](../../../../HtmlStaticBudget/README.md) to reproduce the gate.
