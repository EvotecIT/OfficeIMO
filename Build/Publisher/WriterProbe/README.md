# Publisher writer feasibility probe

This opt-in contributor tool inventories native publications and synthesizes
small, template-free PUB candidates. It stays outside the normal solution and
does not add a production API, runtime dependency or qualified writer profile.
The native bytes are a hypothesis until Microsoft Publisher opens and resaves
them without repair and produces the expected reference output.

Build with the repository's pinned SDK:

```powershell
dotnet build Build/Publisher/WriterProbe/WriterProbe.csproj -c Release
```

The assembly uses the existing `OfficeIMO.Publisher.Tests` friend identity to
access the Publisher and shared compound codecs. Its separate output directory
keeps it apart from the test assembly. It has no test-framework reference and
requires no new production visibility hook.

Invoke the built assembly with the input and an unused report filename:

```powershell
$probe = 'Build/Publisher/WriterProbe/bin/Release/net10.0/OfficeIMO.Publisher.Tests.dll'
dotnet $probe inventory ./reference.pub ./reference.inventory.json
dotnet $probe diff ./native-text.pub ./native-textlong.pub ./text-change.diff.json
dotnet $probe create blank ./candidate-blank.pub
dotnet $probe create-native-id blank ./candidate-blank-native-id.pub
dotnet $probe create-native-header blank ./header-blank.pub
dotnet $probe create-bootstrap blank ./bootstrap-blank.pub
dotnet $probe create-native-schema blank ./schema-blank.pub
dotnet $probe create-native-schema-roles blank ./schema-role-blank.pub
dotnet $probe create text ./candidate-text.pub 'Template-free publication'
dotnet $probe create shape ./candidate-shape.pub
dotnet $probe create image ./candidate-image.pub ./picture.png
dotnet $probe render ./reference.pub ./managed-comparison
```

Output parent directories must already exist. Existing outputs are never
overwritten. Input is bounded to 64 MiB; probe text to 4,096 characters and
images to 1 MiB. The inventory records compound metadata and hashes, known
Contents fields, Quill directories, OfficeArt envelopes and a separate result
from the OfficeIMO reader. Unknown encodings and incomplete decoding remain
visible. Reports can contain complete recovered text: use synthetic fixtures.

`diff` compares the same inventories by stream name, native chunk ID and Quill
tag/ID. It separates changed payload identities and known field values from
record relocation. Repeated fields retain occurrence indices. Opaque payloads
remain opaque; matching inspected values do not establish semantic equivalence.

Candidate creation never reads a PUB, copies an existing compound stream, or
invokes Publisher. The minimal stream graph deliberately exposes unresolved
native defaults rather than hiding them in an embedded seed. Its output is
not a supported document-generation contract. Header defaults, document/page
utility records, class metadata and auxiliary Quill structures require native
qualification. Own-reader acceptance alone does not close that requirement.

`create` leaves the compound root class ID empty. `create-native-id` changes only
that metadata to `00021201-0000-0000-00c0-000000000046`, the value independently
observed in the retained Publisher-authored fixture. Keep both variants when
checking whether that metadata changes native acceptance.

`create-native-header` is a separately named header/record-type hypothesis. It
also synthesizes the observed Contents version, declared stream length, header
marker, trailer scalar identifiers, directory field identifier, two-byte chunk
kind and typed offset/parent descriptors. Page/object lists use the observed
array type. It retains the small model and original identifier allocation; it
does not copy auxiliary bootstrap records, styles or stream payloads from a
publication. Those unresolved requirements still prevent a native support claim.

`create-bootstrap` retains the header hypothesis, reserves directory slots
0 through 255, allocates model records from 256 and updates their native
references. It synthesizes the inert root CompObj identification through the
shared OLE writer: observed user type `Microsoft Publisher 3.0`, clipboard format
`MSPublisher.3`, the native class ID and an empty program ID. It does not copy a
native CompObj stream or Quill style payload. This identifies the experiment;
it does not establish that Publisher accepts the remaining record graph.

`create-native-schema` adds a typed property chunk beginning at offset 44 and
a minimal class catalog for the page, document and the selected shape kind.
Lengths and document/catalog/trailer pointers are calculated from the generated
records. The class descriptors spell out observed version and numeric traits;
their behavioral meanings and the sufficiency of a subset catalog remain
unqualified. The producer's 143-entry catalog and native text-style payloads
are not copied. The relative identification filename is `WriterProbe.pub`.

`create-native-schema-roles` adds the explicit document role fields to a new
hypothesis: field 1 declares one ordered page reference, `0x2C` is 1 and `0x2D`
is 0. The recovered physical interval is the single reference at index 0.
This exercises the native role metadata contract rather than inferring pages
from a content count or a publication-specific identifier. Native acceptance
of a publication with this minimal page/master/utility graph remains unqualified.
Earlier commands keep their original role omissions as negative controls.

`render` creates a fresh output directory containing `managed.pdf`, per-page SVG
and `managed.json` with normalized page, frame, story, table and diagnostic
evidence plus artifact hashes. The PDF header is checked. Its reports explicitly
leave native acceptance and visual fidelity unassessed. Compare these artifacts
against the separately authored native reference and its expected content.

For each candidate retain its SHA-256, the native application version, native
open/repair result, native resave result, exported reference PDF and inspection
of the intended content. Keep these evidence classes separate from corpus
files authored by Publisher and from managed record mutations.
