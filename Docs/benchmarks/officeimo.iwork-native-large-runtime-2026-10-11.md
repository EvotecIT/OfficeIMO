# Apple-produced Keynote scale observations — 2026-10-11

All 18 measured samples pass complete 512-slide validation: twelve standard
reading/conversion samples and six separate memory-sampling observations. This
run covers one Apple Keynote 15.4 source with two repeated text layouts. Concurrent
Swift builds and uncontrolled processor placement limit timing conclusions.

The [raw receipt](iwork-native-large-runtime-2026-10-11.json) retains every sample,
PowerForge summary and metadata, artifact hashes, source/assembly identities,
correctness results and the review finding's reproduction and fix.

## Workload and validation

The [native reference](../../OfficeIMO.TestAssets/Documents/IWorkCorpus/native-exports/keynote-scale-512-v15.4.json)
records Keynote duplication from two to 512 slides, native save/close/reopen,
inspection of slide 512 without repair, and complete Apple PPTX/PDF exports.
The 1,286,730-byte source contains 9,178 inventoried native messages. Independent
schemas initialize the selected slide/text/note/frame records and verify every
ordered slide identity, text block, note, background, frame and visible status
against Apple's PPTX. Complete text and canvas checks cover all 512 PDF pages;
six representative Cairo renders are inspected.

`LoadProject` opens the immutable input and reads its Keynote projection.
`ConvertSave` opens the source, requires complete editable reconstruction, and
saves through the PowerPoint owner. Input preparation and complete validation
remain outside timing. Every saved slide passes text, notes, background,
visibility, total shape count, position/size, zero rotation and no-flip checks.
The complete destination passes Open XML validation. Independent XML reading
also verifies a retained reconstructed PPTX.

Both .NET 8 and .NET 10 builds have zero warnings/errors, and all 18 deterministic
operations pass under each target. Four source corruption probes reject missing
slides, reordered native identities, missing notes and changed frames. Review
exposed a saved-output false pass for rotation and non-text overlays; the corrected
validator rejects both and both flips while unchanged output passes all 512
slides. One full read-only review and one targeted confirmation accept that fix.

## Measurement conditions and observations

Source commit: `da9c291c692bfa0b9033a0009a8ec2c944346241`.
PowerForge source: `54a8a91706b9406529884f28ab63897399c28be4`.
Host: macOS 27.0.1 build 26A434, arm64 Apple M4, four performance and six efficiency
cores, AC power, low power mode off. Priority and affinity are inherited.
PowerShell 7.5.4 executes the net8.0 workload on .NET 9.0.10; the .NET 8/10 commands
above provide separate correctness proof.

Each fresh host performs one warmup and three rotated measured iterations per
operation, with collection before each iteration and no outlier removal.
Standard runs are `20261010-233653-d9845ef2` and `20261010-233717-11512066`.
The separate 10 ms sampling run is `20261010-233754-91dab743`.

- Reading median: 216.7 ms, range 187.7–835.8 ms; allocation 161.3–162.1 MB.
- Conversion/save median: 1,203.4 ms, range 725.1–2,852.4 ms; allocation
  391.3–392.1 MB. Saved output ranges from 905,482 to 905,525 bytes.
- Sampled process managed-heap maxima range from 89.9–93.2 MB while reading and
  120.8–123.3 MB while converting. Resident maxima range from 264.4–280.9 MB and
  281.3–289.3 MB respectively. Every instrumented sample has valid observations.
- Post-release collected deltas range from 51,848–284,944 bytes for reading and
  49,568–419,416 bytes for conversion. Relative to each host's first collected
  baseline, the respective ranges are 213,376–2,181,728 and 363,048–1,348,832 bytes.

MB uses decimal units. Allocation includes host invocation; collected deltas
include validation, caches and host effects. Instrumented timing stays separate
from standard timing. Sampled maxima are observed lower bounds, with observer
effects, rather than exact peaks or attributed native allocations. These short
runs establish neither a leak-free long workload nor portable resource budgets.

## Reproduction and limits

Use the [runtime wrapper](../../OfficeIMO.IWork.Benchmarks/README.md) with
`-Scale NativeLarge -Kind Keynote -Operation LoadProject,ConvertSave`. Add
`-MeasureRetainedMemory` for collected baselines and run
`-MemorySamplingIntervalMilliseconds 10` separately for sampled observations.
The source remains hash-pinned and every successful sample verifies 512 slides.

Two repeated layouts do not qualify diverse large presentations, large Pages or
Numbers sources, opaque native records, native rendering of the reconstructed
PPTX, or other platforms and runtimes. Poppler bbox extraction misplaces this
PDF's `fi` ligatures, and its Splash renderer produces blank later-page text;
pypdf and Cairo supply the recorded bounded checks. pypdf reports two
wrong-pointing objects in the Apple PDF, so PDF conformance is unqualified.
Cancellation phases, idle-host budgets, exact peaks, attributed native retention,
sandbox/device and distribution acceptance remain in [I5](../ROADMAP.md#i5-runtime-and-apple-host-acceptance).
