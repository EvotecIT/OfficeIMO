# Archive qualification consumer

`PackageConsumer` is an opt-in executable outside the solution. It restores the
packed Archives and Reader ZIP APIs and checks the independently produced ZIP
fixtures against their manifest hashes, names and member bytes. It exercises
inspection, grouped reads, verification, verified copying, strict memory reads,
filesystem extraction and document ingestion. Producer tools are validation-only;
OfficeIMO does not install, ship or invoke them.

Use the candidate version and an isolated NuGet package cache when testing a local
feed. Keep the feed, cache, build artifacts and output directory under the configured
scratch root. An unconfigured restore is inert so repository-wide dependency discovery can scan this opt-in project. Building or running it still requires `OfficeIMOQualificationVersion`; no public default or placeholder package version is used.

The fourth version segment used by local qualification is private
build identity; it is not a public dependency pin or release version.

```sh
dotnet restore Build/Archives/PackageConsumer \
  --property:OfficeIMOQualificationVersion="$candidate_version" \
  --source "$candidate_feed" --source https://api.nuget.org/v3/index.json \
  --packages "$qualification_cache" --artifacts-path "$scratch/artifacts"
dotnet run --no-restore --project Build/Archives/PackageConsumer -c Release -f net10.0 \
  --property:OfficeIMOQualificationVersion="$candidate_version" \
  --property:RestorePackagesPath="$qualification_cache" \
  --artifacts-path "$scratch/artifacts" -- \
  OfficeIMO.Archives.Tests/Fixtures/IndependentZip \
  "$scratch/output-net10" "$scratch/package-net10.json"
```

Run the same consumer with `-f net8.0` and fresh output paths. The candidate feed
must contain the matching `OfficeIMO.Core`, `OfficeIMO.Archives`,
`OfficeIMO.Reader.Core` and `OfficeIMO.Reader.Zip` packages. Standard project
`dotnet pack` builds are sufficient for unsigned local qualification; release
packaging and signing remain owned by the repository's PowerForge release lane.
Do not publish qualification packages to the ordinary cache or public feeds.

Additional optional positional arguments are:

```text
<ordinary MSI> <independently extracted payload directory> <makecab CAB> <original member>
```

The MSI route compares the complete multiset of SHA-256 payload hashes from
inspection and published extraction with an independent extraction. The CAB route
compares a complete makecab-produced member with its original file. No installer,
custom action or extracted executable is run. Use both the package hash and the
producer/tool identity when recording an external qualification run.

The small ZIP fixtures contain original public-domain text. Their
`manifest.json` records producer commands, exact input hashes, expected member
hashes and explicit encodings. The macOS 7-Zip `-mcl=on -mcp=1250` run actually
produces unflagged UTF-8 bytes; it is labelled and read as code page 65001.
CP1250 qualification uses CPython's ZIP writer with an explicit filename encoder.
The two streaming fixtures use nonseekable CPython output, with and without
forced ZIP64 member descriptors. None was written by OfficeIMO.

Native Windows Explorer creation of Polish filenames, newer MSI producer
families and Windows runtime execution remain distinct qualification outcomes.
Compilation for `net472`, macOS execution, a package signature and green managed
fixtures must not be reported as those outcomes. Keep current coverage in
`OfficeIMO.Archives/SUPPORT.md` and open acceptance work in `Docs/ROADMAP.md`.

The [2026-10-17 qualification record](Evidence/2026-10-17.json) records the
independent input identities, private package hashes, executed runtimes and
remaining acceptance gaps. The [verified-copy benchmark](../../OfficeIMO.Archives.Benchmarks/README.md#verified-member-copying)
is opt-in and checks equal payload hashes before reporting allocations or time.

## Large ZIP input evidence

The package consumer also provides an opt-in allocation probe. It creates a ZIP
slightly above 64 MiB if the input path is absent, then checks inspection,
metadata traversal and Reader ingestion of the same small text member. It
records per-thread managed allocations separately for each workflow. Synthetic
source I/O failure and cancellation check private snapshot cleanup, source
position restoration and caller ownership.

After restoring the chosen candidate as above, replace the consumer's positional
arguments with:

```text
--large-input-evidence <task-owned ZIP path> <JSON report>
```

Reuse the identical ZIP input when comparing candidates. The report retains its
SHA-256, runtime and assembly version. The probe is outside normal correctness
CI; allocation totals are neither peak/RSS measurements nor timing guarantees.

To measure Reader's live temporary storage, use:

```text
--reader-snapshot-evidence <same large ZIP> <JSON report> 1
```

Run this probe in a process with an isolated `TMPDIR` (or `TEMP` on Windows),
without parallel consumers. A custom text handler observes private snapshot files
while the outer ZIP inspection is alive. The report checks one input-sized
snapshot for path and stream dispatch, caller stream ownership and position,
and cleanup after completion, early disposal, cancellation and handler failure.
The final argument is the expected number of input copies, allowing the same
probe to reproduce an earlier candidate with two copies. This is opt-in disk
usage evidence outside correctness CI.

## Solid-folder read evidence

Use `--cab-read-evidence <task-owned CAB/MSI path> <JSON report>` after the same
isolated restore. If the input is absent, the runner creates a reproducible
32 MiB synthetic LZX folder with two small members, including a range across
frames. It compares separate `OpenRead` calls with one grouped `ReadMembers`
call and checks that their complete member hashes agree. An existing native CAB
or MSI can be supplied instead; its two smallest supported nonempty members are
selected without executing content.

Reuse the identical input for before/after candidate comparisons. Input inventory
is outside the measured region. Reports retain the input hash, runtime, package
identity, member hashes, decoded-byte/work charges and per-thread read allocations.
This is opt-in allocation evidence, not a timing or peak-memory benchmark. The
synthetic producer does not establish native CAB compatibility.
