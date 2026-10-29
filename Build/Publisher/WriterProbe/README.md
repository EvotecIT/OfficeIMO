# Publisher writer feasibility probe

This opt-in contributor tool inventories native publications and synthesizes
small, template-free PUB candidates. It stays outside the normal solution and
does not add a production API, runtime dependency or qualified writer profile.
The native bytes remain a hypothesis until Microsoft Publisher opens and resaves
them without repair and produces the expected reference output.

Build with the repository's pinned SDK:

```powershell
$env:MSBUILDDISABLENODEREUSE = '1'
$env:DOTNET_CLI_USE_MSBUILD_SERVER = '0'
dotnet build Build/Publisher/WriterProbe/WriterProbe.csproj -c Release
```

The assembly uses the existing `OfficeIMO.Publisher.Tests` friend identity to
access the Publisher and shared compound codecs. Its separate output directory
keeps it apart from the test assembly. It has no test-framework reference and
requires no new production visibility hook.

Invoke the built assembly with an input and unused output paths:

```powershell
$probe = 'Build/Publisher/WriterProbe/bin/Release/net10.0/OfficeIMO.Publisher.Tests.dll'
dotnet $probe inventory ./reference.pub ./reference.inventory.json
dotnet $probe diff ./native-text.pub ./native-textlong.pub ./text-change.diff.json
dotnet $probe create blank ./candidate-blank.pub
dotnet $probe create text ./candidate-text.pub 'Template-free publication'
dotnet $probe create shape ./candidate-shape.pub
dotnet $probe create image ./candidate-image.pub ./picture.png
dotnet $probe render ./reference.pub ./managed-comparison
```

Output parent directories must already exist. Existing outputs are never
overwritten. Inventory input is bounded to 64 MiB; probe text to 4,096 characters
and images to 1 MiB. Reports can contain complete recovered text: use synthetic
fixtures.

`inventory` records compound metadata and hashes, known Contents fields, Quill
directories, OfficeArt envelopes and a separate result from the OfficeIMO reader.
Unknown encodings and incomplete decoding remain visible. `diff` compares these
inventories by stream name, native chunk ID and Quill tag/ID. It separates changed
payload identities and known field values from record relocation. Repeated fields
retain occurrence indices. Opaque payloads remain opaque; matching inspected
values do not establish semantic equivalence.

`create` never reads a PUB, copies an existing compound stream or invokes
Publisher. It synthesizes a current experimental profile with identifiers from
256, typed Contents header properties, a minimal class catalog, an explicit page
role record and generated root CompObj identification. The canonical compound
writer serializes a fresh directory model carrying the independently observed
Publisher root and QuillSub storage class IDs. No source publication supplies
that directory or its stream bytes. The generated directory includes an empty
root `Objects` storage: the captured native preliminary version reader opens it
before reaching format validation, including for a blank publication.
The root `\u0003Internal` stream is encoded as a 12-byte typed header: version
`0x00010000`, a zero reserved WORD, zero context flags at unaligned offset 6 and
a zero following-byte count at offset 10. The preliminary native reader requires a complete
12-byte read and that version. The reserved zeros follow the native writer's initialization;
zero context flags describe this plain, inert candidate.

The Quill envelope writes the native checked version values 4 and 7, byte
`+0x0F` equal to 3, a first payload position of 512 and the truthful stream size
at `+0x14`. Descriptors retain their generated payloads and point to their actual
serialized positions. Text-format range endpoints follow the relocated text.
The stream has at least 1,024 bytes when descriptors are present. Header
`+0x0C` declares the total descriptor count; the single directory declares its
own count separately. Both counts follow the actual generated records.

The current Quill defaults are generated libraries: `INK ` version words
`0x74/0x74`, an indexed `FONT` library with a generated Arial entry, one `STSH` name entry selecting
the native built-in default and two paired empty character/paragraph OPLs in
`STSH` ID1. Indexed body sizes, entry counts and offsets follow the serialized
entries. The font entry supplies index zero used by native default character
mapping. Auxiliary style ID2 is absent. These records follow the recovered
loader contracts; native acceptance remains separate. No native text-style
payload is copied.

The minimal class descriptors spell out observed version and numeric traits;
their behavioral meanings and the sufficiency of a subset catalog remain
unqualified. The generated page graph has six ordered references: one master,
one physical page and four utility pages, with `DpgMaster=1` and `DpgSpecial=5`.
The recovered physical interval is `[1, 6-5]`. This satisfies the native special
boundary range of 4 through 30 and provides the ordinary loader's reserved page
at index `N-4`. The non-master pages reference the generated master. Optional
auxiliary page records remain absent; the inspected validator permits their zero
references, and this does not establish that the ordinary native loader needs no
others. Document field `0x20` points to a generated Quill document (`0x5B`), with
decoded references to empty `0x61` and `0x65` children and its observed options
field. Those used classes receive explicit catalog descriptors. A successful
managed read establishes only the managed
reader result. Native open, native resave and content fidelity are separate
qualification requirements.

Document field 3 points to a generated print block (`0x4B`) containing zero
printer profiles and an empty profile list. Those values meet the recovered
structural count/list constraints. Optional printer settings and flags remain
absent for native defaults; no device-specific printer payload is copied.
Normal native decoding of this minimal print block remains unqualified.

The earlier named hypothesis commands and their exact source are retained in
commit `a69b7b8b56`. Retained raw candidates and their native rejection records
remain negative controls; the current command does not reproduce those older
variants.

`render` creates a fresh output directory containing `managed.pdf`, per-page SVG
and `managed.json` with normalized page, frame, story, table and diagnostic
evidence plus artifact hashes. The PDF header is checked. Its reports explicitly
leave native acceptance and visual fidelity unassessed. Compare these artifacts
against the separately authored native reference and its expected content.

For each candidate retain its SHA-256, the native application version, native
open/repair result, native resave result, exported reference PDF and inspection
of the intended content. Keep these evidence classes separate from corpus files
authored by Publisher.
