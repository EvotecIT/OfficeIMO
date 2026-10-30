# Publisher native fixtures

The seven corpus publications come from Apache POI's
[`test-data/publisher`](https://github.com/apache/poi/tree/66109187d12b72696ea02101cec765d77d9bbdc9/test-data/publisher)
corpus. Source revision and file hashes are recorded in `provenance.json`.
`LICENSE-APACHE-POI.txt` and `NOTICE-APACHE-POI.txt` retain the upstream terms.

Simple, Sample and Sample_2010 exercise positioned text, styles and tables.
SampleBrochure and SampleNewsletter exercise grouped artwork, GIF/JPEG/WMF
resources, page order and linked stories. Sample2000 and Sample98 protect the
explicit earlier-generation rejection contract.

The fixtures are test inputs, not native output authored by OfficeIMO. Expected
simple/table geometry and style values are checked against libmspub 0.1.5's
`pub2raw` output. Native Publisher PDF renderings are not supplied by this corpus.

`native-authored.pub` and its reference PDF were created in Microsoft Publisher
16.0 using [the opt-in Windows producer script](../../Build/Publisher/New-NativeQualificationFixture.ps1).
The fixture contains a text frame, two-color gradient, decorated dashed line,
rotated diamond and 2-by-2 table on a 612-by-792-point page. Its producer metadata
and SHA-256 hashes are in `native-authored.json`. These authored fixture files
use OfficeIMO's MIT license; the Apache POI terms apply to the upstream corpus.

## Produce native comparison controls

[New-NativeComparisonCorpus.ps1](../../Build/Publisher/New-NativeComparisonCorpus.ps1)
creates synthetic publications directly in Microsoft Publisher and exports their
native PDF references. Run it in PowerShell 7 on Windows with Publisher installed and no existing
Publisher session. Choose an empty task-owned output directory:

```powershell
./Build/Publisher/New-NativeComparisonCorpus.ps1 -OutputDirectory ./native-controls `
    -Case Blank,Text,TextLong,Rectangle,Image
```

The minimal cases isolate a blank publication, one text frame, a longer version
of the same story, a solid rectangle and an embedded synthetic picture. BlankPages
has two empty physical pages; TextBlank has a text page followed by an empty page.
They distinguish blank document pages from utility and master definitions. Text and
TextLong use identical page geometry, frame geometry and Arial formatting; their
only authored difference is the text. TextParagraphs has one Arial 12-point story
with an empty paragraph between two text paragraphs. Its retained character
formatting covers that empty paragraph's CR; the final story CR has a separate
native character range. TextStories has three separate stories, with an empty
paragraph inside its middle story. TextFormattingPages retains 96 paragraphs,
including overflow beyond its sole text frame, across two native paragraph-formatting
pages. ZeroAnchors places each client-anchor coordinate at native zero: native
storage omits the corresponding property instead of writing an explicit zero.
Its four text frames qualify recovery of those omitted coordinates.
Newsletter adds a two-page linked story
and a picture. TableLabel adds a custom page size, styled 2-by-2 table, picture
and Polish/Greek characters. Brochure adds two 792-by-612-point landscape sides,
each with three 264-point panels, explicit text frames and panel markers.
MultipleMasters applies two single-page masters in A/B/A order across three
physical pages. TwoPageMaster retains the default first-page master and applies
the distinct members of a native two-page master to the next left and right
pages. Distinct marker text and geometry identify each master in the reference
PDF. PictureOwnership places two physical pictures on its first page and one
picture on a master shared by both physical pages. The second physical picture
uses Publisher's [Shape.Duplicate method](https://learn.microsoft.com/en-us/office/vba/api/publisher.shape.duplicate).
Its saved native image-store references must be inspected separately; duplicating
a shape does not establish that Publisher shares the embedded image asset.
The retained PictureOwnership PUB has three placements referencing one native
image-store entry: one master picture and two pictures on the first physical page.
Omitting `-Case` selects every available case.

Each case has a PUB, a Publisher-exported PDF and JSON that separates its authored
contract from the observed native page/shape inventory. The root manifest records
the case contracts and file hashes. Master controls also record authored
assignment IDs and the observed shapes on each native master definition.
PictureOwnership also records its authored zero crop values and native
observations of picture cropping, embedding and page ownership. Numeric crop
values follow [Publisher's original-image point units](https://learn.microsoft.com/en-us/office/vba/api/publisher.pictureformat.cropleft).
Text fonts, margins and object colors are explicit; other native publication defaults
remain recorded as defaults. Producer
metadata records the Publisher executable version, PowerShell version, culture
and Arial font-file hash when available. Native file bytes can vary between runs;
hashes identify a particular produced reference rather than promise deterministic
Publisher serialization.

Master definition geometry uses that master's local coordinate system. Applied
marker placement on physical pages is checked separately against text origins
in the native reference PDF; local definition coordinates do not establish
their applied positions.

The master controls use unique single-character abbreviations. On the qualified
Publisher 16.0 build, the paired-master abbreviation `P` succeeds and `TP` returns
an invalid-argument error; the API documentation does not specify a maximum
abbreviation length. The producer keeps that observed constraint within its
synthetic controls.

The shared [native operation controller](../../Build/Publisher/Invoke-NativePublisherOperation.ps1)
also validates an OfficeIMO-written candidate:

```powershell
./Build/Publisher/Invoke-NativePublisherOperation.ps1 -Operation Validate `
    -InputPath ./candidate.pub -OutputDirectory ./candidate-native
```

Validation takes a byte-identical task-owned copy, opens it read-only without
adding it to recent files, and captures its inventory and native PDF before any
resave. It then opens only the task-owned copy writable, saves a separate PUB and
reopens that result. The original candidate's hash must remain unchanged.
Frame inventories read the visible `TextFrame.TextRange` prefix; they do not
establish that overflow was retained. Complete overflowing stories require a
separate check of the native-resaved Quill text and story-length tables. COM
Open success does not expose whether Publisher repaired a document, so the report
keeps repair status unassessed. Native UI observation is separate evidence.

The controller runs a hidden STA worker with a bounded deadline. It disables
macros for its owned application and records the Publisher process identified by
the COM window and exact process start time. Cleanup never stops processes by
name alone. A startup failure before that identity is recorded leaves uncertain
processes untouched and reports the gap. `.run` contains task-local process
receipts and logs; it is not a fixture input. An explicit completed-operation
receipt separates completed evidence from COM host teardown. The controller can
stop its exact worker after that receipt and records this disposition separately;
an ordinary timeout without a completion receipt remains a failed operation.

These controls are independent producer evidence. They do not establish native
writing, rendering parity or supported editing until the corresponding library
operation is checked against the exact retained files. Promote only reference
files used by executable tests or a maintained qualification contract; keep
intermediate runs and process logs outside the source fixture directory.
