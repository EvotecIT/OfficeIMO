# EPUB validation evidence

`validate_epub.py` runs an explicitly supplied EPUBCheck distribution over captured
publication bytes. It is an opt-in validation tool, outside normal builds and product
runtime dependencies. It does not install Java, EPUBCheck or any accessibility tool.

Download an official [EPUBCheck distribution](https://www.w3.org/publishing/epubcheck/)
and retain its adjacent library directory. Supply its JAR and a new output directory:

```sh
python3 Build/Epub/validate_epub.py book.epub edited-book.epub \
  --epubcheck-jar /path/to/epubcheck/epubcheck.jar \
  --output /path/to/task-evidence/epubcheck
```

The runner records the tool version and JAR hash, captures each input, hashes the
captured bytes, and retains the validator's log and JSON report. `summary.json`
identifies each input and outcome. A failed validator, timeout or missing report
produces a nonzero exit. Existing output directories are rejected to prevent mixing
reports from different runs. Inputs are limited to 256 publications, 128 MiB each;
`--timeout` controls the per-process limit in seconds.

To also capture automated accessibility checks, supply `--ace /path/to/ace-puppeteer`
from an independently installed [DAISY Ace](https://daisy.github.io/ace/) distribution.
Its browser runtime must already be available. The runner records its version, log,
JSON report, exit code, and reported outcome for the same captured EPUB bytes. Missing,
malformed, failing, or timed-out evidence produces a nonzero exit. Without `--ace`,
automated accessibility is recorded as unchecked.
On interruption or timeout, the runner terminates Ace's process tree, including detached
browser groups on POSIX. Cleanup has a bounded grace period beyond the requested audit
timeout. Run the POSIX timeout contract with
`python3 -m unittest discover -s Build/Epub`; it exercises ordinary and detached children.

A passing result establishes only the checks performed by the recorded tool versions.
Review warnings and retain the versions with the evidence. The summary explicitly
leaves comprehensive accessibility assessment and reader presentation unchecked even
when Ace passes. Complete human accessibility review, then inspect the same publication
bytes in the chosen reading systems. Never infer those results from validator exit codes.

Use a task-owned output location. Retain compact reports and decisive publication
fixtures deliberately; remove superseded captured publications and downloaded tools
when they are no longer needed. Source fixtures need producer, version and license
provenance before entering the maintained corpus.

## Independent media-overlay source

The W3C EPUB sample collection's
[Moby-Dick media-overlay source](https://github.com/IDPF/epub3-samples/tree/7651e2002b631e6577fadf7e9e0692fa6efb8746/30/moby-dick-mo)
provides an independent preservation and narration-inspection case. Pin revision
`7651e2002b631e6577fadf7e9e0692fa6efb8746`, retain its copyright and licensing notices,
and package the source with `mimetype` first and uncompressed. Keep the source
payloads unchanged and record the resulting archive hash with the validation reports.

The checked sample has two overlays. It passes EPUBCheck 5.4.0 and the native
`media-overlays` preflight check, and unchanged load/save preserves its archive bytes.
Its package lacks modern accessibility discovery metadata, so the overall native
preflight reports that gap. These results do not establish audio decoding, reader
synchronization, assistive-technology behavior or comprehensive accessibility.
The sample stays outside shipped packages and is not a runtime dependency.

## ONIX fixtures

The opt-in `OnixFixtureGenerator` exercises the workflow's bibliographic ONIX export
with early, advance and confirmed notifications, person and organization credits,
explicit no-contributor metadata, a subtitle, publication date and English, Polish
and French language codes. It writes each ONIX record with its exact EPUB, records
their hashes and the three schema hashes, and checks that the schema rejects an
invalid language code. Commercial fixtures cover worldwide exclusions, distinct
Polish and British markets, all seven supported price bases, decimal scale, effective
dates, free and unannounced pricing, an unknown supply date, and withdrawal after
rights are lost. The schema must also reject unknown country and currency codes.
It is outside normal builds and shipped packages.

Obtain the ONIX 3.1 reference XSD from [EDItEUR](https://www.editeur.org/93/Release-3.0-and-3.1-Downloads/)
and retain its unchanged adjacent code-list and XHTML schemas and their license
notices. The local schema directory must contain `ONIX_BookProduct_3.1_reference.xsd`,
`ONIX_BookProduct_CodeLists.xsd` and `ONIX_XHTML_Subset.xsd`. The runner's resolver
permits only those local files; it does not fetch network dependencies.

```sh
dotnet run --project Build/Epub/Onix/OnixFixtureGenerator.csproj -- \
  /path/to/onix-schema /path/to/new-onix-evidence
xmllint --nonet --noout --schema /path/to/onix-schema/ONIX_BookProduct_3.1_reference.xsd \
  /path/to/new-onix-evidence/early.onix \
  /path/to/new-onix-evidence/advance.onix \
  /path/to/new-onix-evidence/confirmed.onix \
  /path/to/new-onix-evidence/priced.onix \
  /path/to/new-onix-evidence/withdrawn.onix
```

Retain the independent validator version, exit code and log alongside `evidence.json`.
The generator records independent validation and retailer acceptance as unperformed;
its own successful schema check does not stand in for either. Run the generated EPUBs
through the EPUB validation runner separately. ONIX schema conformance is narrower
than trade business-rule validation or recipient acceptance. When schema files come
from a mirror, retain its immutable revision and the original schema headers and
distinguish that provenance from a fresh official download.

## Typography fixtures

The opt-in generator uses the current EPUB owner to create Basic, Prose, and
Technical publications from one manuscript containing tables, code, long links,
an accessible SVG, Arabic, and Japanese. It also creates `glossary.epub` with two
terms, repeated references and localized return links, and `bibliography.epub` with
two entries formatted and title-sorted by `OfficeIMO.Bibliography`, three citations,
and return links. The bibliography fixture checks that order, escaped title text,
and italic formatting survive EPUB reopening. Its CSL style is an illustrative
local style, not a claim of qualification for every publisher's citation style.
It has no third-party test dependency and
is outside the normal solution and shipped packages.

`index.epub` exercises nested terms, multiple locators and an intra-index
cross-reference. Its labels target sections and chapters rather than inferred
screen page numbers. `renamed-index.epub` moves both chapters and the navigation
document into new directories, exercising locator, cross-reference and TOC repair.
Run both artifacts through the independent validators to compare their conformance.

`split-chapter.epub` divides a sectioned chapter into consecutive reading positions,
with forward, return and incoming links repaired across the new resource boundary.
Use it to check TOC order and link destinations in independent readers as well as
with the automated validators.

`fixed-layout-ltr.epub` and `fixed-layout-rtl.epub` contain an 800×600 landscape
canvas and a 600×800 portrait canvas. Typed regions position headers, sections and
footers while preserving the semantic document order. `fixed-layout-escaped-id.epub`
exercises fractional coordinates and an identifier containing punctuation and a
supplementary Unicode character. Verify its region rectangle in a browser as well
as checking its EPUB package. The page fixtures exercise typed orientation, spread and
page-side requests, logical DOM order and semantic links with a package-level
fixed-layout default. `fixed-layout-item-overrides.epub` deliberately omits that
default to expose readers that do not honor item-only fixed-layout declarations.
Validate their exact EPUB
bytes with EPUBCheck, then inspect spread placement, rotation and scaling in target
readers. Extracted XHTML can verify CSS canvas geometry in a browser; it does not
qualify a reader's spine interpretation or assistive-technology behavior.

`read-aloud.epub` pairs two text paragraphs with recorded narration, explicit SMIL
clip intervals, package durations and active-text styling. The test-only MP3 and
its source/encoding notes live in `Fixtures/Assets`. Its accessibility summary
identifies the unnarrated heading and unqualified reader playback. Verify cue
synchronization, highlighting, seeking and pause/resume in a media-overlay reader;
successful decoding and EPUBCheck/Ace results do not establish those behaviors.

`merged-chapters.epub` merges those split reading positions back into one resource,
retaining both TOC entries and repairing the incoming reference. Check its chapter
order, section links and return links alongside the split fixture.

```sh
dotnet run --project Build/Epub/Fixtures/EpubFixtureGenerator.csproj -- \
  /path/to/task-evidence/typography
```

Supply a new directory. The generator writes EPUB files, a manifest with SHA-256
hashes and native preflight results, and expanded content with standards-mode HTML
previews. A native preflight error fails the run; unchecked scopes remain explicit
in the manifest. Fixture accessibility metadata describes the actual content and
known limits without declaring certification. The previews keep
the generated content and profile stylesheet but add simulated reader CSS for
light, dark, and enlarged text. Serve this directory on loopback when inspecting
the preview paths listed in `manifest.json`; for example:

```sh
python3 -m http.server 8766 --bind 127.0.0.1 \
  --directory /path/to/task-evidence/typography
```

Check narrow and wide widths, wrapping, direction, text enlargement, and inherited
colors. Run the EPUB files through `validate_epub.py`, and use the same captured
bytes for native reader checks. Browser previews do not reproduce EPUB pagination,
reader preferences, font handling, or assistive-technology behavior. Retain their
results separately from native reading-system acceptance.
