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

dotnet $probe create styled-text ./candidate-styled-text.pub
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

Text candidates carry generated 512-byte `FDPC` and `FDPP` formatting pages,
allocated on 512-byte boundaries. Each page has one range, auxiliary WORD 1,
a zero reserved DWORD and a calculated page-relative OPL offset. Its range end
is the absolute end of the `TEXT` bytes. `BTEC` and `BTEP` use one four-byte
record, flags zero, bounds `[0, absolute TEXT end]` and the calculated absolute
offset of the matching formatting page. Stream padding and descriptor lengths
follow that layout. These carriers address the observed missing-index load
failure; their text semantics still require native qualification.

`styled-text` uses one shared `OfficeRichTextParagraph` containing three Arial
runs: `Plain 12 | ` at 12 points/plain, `Bold 18` at 18 points/bold, and
` | Plain 12` at 12 points/plain. A final native CR follows the last run.
Each run explicitly selects generated font index zero, its size and its bold
state. The 512-byte character page has three parallel endpoints and relative
style offsets; it includes the reset to plain formatting. Existing plain text,
rectangle, blank and image cases retain their qualified wire encodings.

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
references to `0x61` and `0x65` children and its observed options field. Text-bearing
candidates give each catalog one generated row: the frame row binds story key 1
to the shape, and the story row binds that same key with observed `IdLayout`
(field 7) value 1. This field is separate from the frame reference count. Other
profiles retain empty catalogs. The one-story `SYID` uses first DWORD 1, count 1
and key 1; no lower
identifier limit is inferred from controls that use key 2. Those used classes
receive explicit catalog descriptors. A successful
managed read establishes only the managed
reader result. Native open, native resave and content fidelity are separate
qualification requirements.

Every profile references a generated `OplSccm` color scheme (`0x5C`) through
document field `0x22`, `OhColorScheme`. `Cecp=8` agrees with the eight `Rgecp`
rows: automatic black uses the native compressed row, followed by seven
explicit RGB colors. The RGB values are constructed from named color components;
no native row payload is copied. Optional scheme identity, name and extended
color properties retain qualified native defaults. This metadata supplies the
publication context used to resolve indexed colors, including automatic black
from the text engine. Its addition changes every profile's publication bytes;
native lifecycle results must therefore be qualified again for those outputs.

The master alone references a generated margin-guide record (`0x4C`) through
field 3, `OhMgPage`. Four typed guide rows form the full-page rectangle with
coordinates `0,width,0,height`; page dimensions are explicit, with two column
guides, two row guides and zero gutters. Optional gutter flags retain their native
false defaults. Physical pages use their existing master relation for the native
fallback. This supplies the calculated geometry used by the render consumer;
native export qualification remains a separate check.

Document field 3 points to a generated print block (`0x4B`) containing one empty
typed `OplDnb` profile. The ordinary native open path accesses row zero of that
array; an empty list is insufficient even when it passes structural validation.
Profile fields remain absent for native defaults; the recovered entry validator
accepts those defaults. Native materialization and use of that minimal row remain
separate qualification requirements. No printer name, device-specific payload or
DEVMODE data is copied.

The drawing stream contains a calculated `F006` drawing-group record with one
identifier cluster, followed by one `F002` drawing container. Even the blank
candidate has a root group at SPID 1024; content shapes use SPID 1025. Saved
shape/drawing counts, last shape ID and cluster values follow those generated
objects. The next-SPID convention follows the native controls and existing PPT
writer. No native property table or extra utility drawing is copied. Image
candidates use the shared delayed FBSE and standalone BLIP writers, placing the
image record in `EscherDelayStm` and preserving picture shape type 75. This
matches the independently observed native storage layout and the independent
reader's delayed image path. The common blank publication bytes are unchanged.

Text-bearing candidates add a calculated `F00D` ClientTextbox OPL with story
key 1 in field 1/type `0x20`. Its record header follows the native Text,
Rectangle and Newsletter controls. Text boxes use native FSP type 202. The
standalone rectangle retains type 1 and owns a one-CR story, with matching
Contents bindings, frame/story catalog rows and ordinary Quill formatting
carriers. A usable native frame text range, save/reopen and export remain
separate qualification checks.

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
