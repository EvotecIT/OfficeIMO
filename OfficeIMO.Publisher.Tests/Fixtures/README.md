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
native PDF references. Run it on Windows with Publisher installed and no existing
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
only authored difference is the text. Newsletter adds a two-page linked story
and a picture. TableLabel adds a custom page size, styled 2-by-2 table, picture
and Polish/Greek characters. Omitting `-Case` selects all nine controls.

Each case has a PUB, a Publisher-exported PDF and JSON that separates its authored
contract from the observed native page/shape inventory. The root manifest records
the case contracts and file hashes. Text fonts, margins and object colors are
explicit; other native publication defaults remain recorded as defaults. Producer
metadata records the Publisher executable version, PowerShell version, culture
and Arial font-file hash when available. Native file bytes can vary between runs;
hashes identify a particular produced reference rather than promise deterministic
Publisher serialization.

The shared [native operation controller](../../Build/Publisher/Invoke-NativePublisherOperation.ps1)
also validates an OfficeIMO-written candidate:

```powershell
./Build/Publisher/Invoke-NativePublisherOperation.ps1 -Operation Validate `
    -InputPath ./candidate.pub -OutputDirectory ./candidate-native
```

Validation takes a byte-identical task-owned copy, opens it read-only without
adding it to recent files, and captures its inventory and native PDF before any
resave. It then opens only the task-owned copy writable, saves a separate PUB and
reopens that result. The original candidate's hash must remain unchanged. COM
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
