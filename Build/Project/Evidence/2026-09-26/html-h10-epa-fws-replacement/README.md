# EPA and FWS held-out replay

The [H10 selection manifest](../../../../../OfficeIMO.Pdf.Benchmarks.Comparisons/Corpus/html-h10-page-selection.json) named the EPA drinking-water article and FWS wetland article, their source classes, PDF intents, editable targets, and inspection criteria before either page was captured. Their third-party page bytes and rendered excerpts remain in ignored local evidence, not in this repository. Both archives were acquired at clean source `fee44e59f`; the EPA archive is 1,196,649 bytes (SHA-256 `12ac44c94520a28c5797c1c764df47ecb833db932e2772be868a25e9b7062dc3`) and the FWS archive is 7,971,709 bytes (SHA-256 `644ccacfe10dde7484549e46642be88727c34f9ba98fb36407cb6fca9f1f2fd4`). EPA loaded 4/4 images and FWS loaded 27/27. The selection is held out from selection by OfficeIMO output, but EPA's table-caption defect has now informed implementation; use it as a frozen regression from this point onward.

## EPA: three distinct PDF intents

At clean source `64f1e80749d169b99934c622729a55b05b58ba5d`, the opt-in `html-mhtml-evidence` runner replayed the unchanged EPA archive offline with Chromium 151.0.7922.34, PeachPDF 0.9.19, and OfficeIMO. All 13 PDF operations completed. The current replay's 60 dpi rasters are pixel-identical, page for page, to the earlier capture for the five main Chromium and OfficeIMO intents; PDF binary metadata changed on the Chromium side. Every printed page was inspected in contact sheets.

| Intent | Chromium | OfficeIMO | Observed difference |
| --- | ---: | ---: | --- |
| Print reflow | 4 pages | 4 pages | OfficeIMO's table has column rules but lacks the browser's complete cell grid; the section, table and footer break at different places. PeachPDF prints 3 pages on the same archive, so page count alone is not a fidelity verdict. |
| Screen-media pagination | 3 pages | 4 pages | OfficeIMO's table spans two sheets and the dark footer continues onto a mostly empty fourth sheet; Chromium completes it on three. |
| Screen snapshot pagination | Frozen screen reference retained separately | 5 pages | The fifth OfficeIMO slice contains only a dark footer strip. This is not equivalent to browser screen-media print; compare viewport slices and final padding on their own terms. |

The EPA PDF conversion reports are `Degraded`: default print records 84 unavailable resources, 25 unavailable font faces, an SVG-content warning, and a stylesheet URL resource warning; screen intents also report unsupported SVG and positioning cases. These counts include repeated references, not distinct missing files. The output contains readable article text and contaminant values, but the resource and geometry losses prevent qualification as browser-faithful PDF. Exact PDF page counts, sizes, SHA-256 values, operation timings and warnings are in the local JSON report.

## EPA: six editable targets

The same frozen archive was converted with its captured image and stylesheet context through Word, Excel, PowerPoint, OneNote, RTF and Markdown at `64f1e8074`. All six artifacts saved and reopened, and each has an operation-level report. The first table-caption fix preserved the full authored caption in Word, RTF, OneNote, PowerPoint and Markdown. Excel follow-up `ed2dbd9db` keeps the native two-column table at A1:B16 and writes the full EPA caption in A18, after the grid; the worksheet tab still has Excel's 31-character limit. A separate linked-caption regression proves that a single link covering the caption can survive save/reopen. The EPA caption itself is plain text. This is an explicitly reported placement approximation, not a caption omission. The source's `National Secondary Drinking Water Regulations` section heading remains on a second sheet.

| Target | Saved/reopened | Declared marker result | Material remaining gap |
| --- | --- | --- | --- |
| Word | Yes | 5/5 | Imported styling and document page flow remain visually unqualified. |
| Excel | Yes | 5/5 on the follow-up | Full EPA caption survives after the table without shifting typed data coordinates; resource/image omissions remain. |
| PowerPoint | Yes | 5/5 | Final exact-head LibreOffice render has 20 slides, many with only one or a few table cells. Caption and first row stay together, but this is not a readable slide grouping. |
| OneNote | Yes | 5/5 | Semantic content survives; whole-page appearance and re-export fidelity remain unverified. |
| RTF | Yes | 5/5 | Table/content selection and whole-document appearance need visual qualification. |
| Markdown | Yes | 5/5 | Semantic table and link projection are present; styling and image behavior are format-specific losses. |

Each target reports loss. Common diagnostics include 82 unavailable resource references and stylesheet-resource warnings; some adapters additionally report 29 pending or skipped external stylesheet references. These are explicit resource limits, not proof that all visible resources failed. The first one-off sequential Mac probe measured roughly 3.6–8.0 seconds and 1.0–1.1 GB cumulative allocation per conversion. The Excel follow-up took 10.4 seconds under a different host load; neither run measured peak memory or satisfies supported-platform budgets. A LibreOffice PDF export of the follow-up workbook remains nine pages, the same as the baseline, and the first sheet visibly shows the complete caption below the native contaminant table. The earlier placement beside the table produced an extra page and was discarded.

The linked-caption sibling sweep and table-placement regressions passed in `OfficeIMO.Html.Tests`: 3,537/3,537 on each of .NET 8 and .NET 10 at the first caption head, then 3,537/3,537 on .NET 10 and a complete rerun of 3,537/3,537 on .NET 8 for the Excel follow-up. The first .NET 8 run had one detached-projection concurrency failure outside the Excel path; that test passed in isolation and the full rerun passed. The touched adapter projects built for netstandard2.0 with no warnings. Exact clean-head H4/advanced-held-out passed 8/8 at the first caption head; the follow-up still needs its exact-head H4 replay. These gates protect the current selected contracts; they do not establish unfamiliar-page visual equivalence.

## FWS: bounded rejection

The FWS archive captured 46 resources and all 27 page images without a pending image. Offline Chromium prints seven pages and screen-media prints eight; PeachPDF prints seven. All OfficeIMO PDF intents stop at the default CSS rule limit (`actual=10001`, `limit=10000`). An exploratory replay with a 20,000-rule limit also stops at `actual=20001`; it was not adopted as a new untrusted-input policy. No OfficeIMO FWS PDF or six-target editable acceptance is claimed. Determine the effective CSS scope and measure time, allocation and peak memory on supported platforms before changing the cap or qualifying this source.

## Reproduction and next checks

The clean-head EPA PDF report and PDFs are under `Ignore/HtmlUnknownPageQualification/h10-epa-replay-clean-64f1e8074/`; the first six editable reports and artifacts are under `h10-epa-editable-clean-64f1e8074/`, with the Excel follow-up under `h10-epa-excel-caption-candidate/` until its clean-head replay. The FWS archive and report are under `h10-fws-wetland-capture-fee44e59f/`; the 20,000-rule report remains under `h10-fws-wetland-20k-clean-fee44e59f/` without duplicated source and PDFs. These local paths are evidence retention locations, not redistributable fixtures. The opt-in runner command is documented in [the comparison README](../../../../../OfficeIMO.Pdf.Benchmarks.Comparisons/README.md); replay with `--mhtml <frozen source.mhtml> --replay-browser --require-clean-source` and a new output directory. The editable evidence used a task-owned probe in `Ignore/HtmlUnknownPageQualification/h10-epa-editable-probe-fee44e59f/`, which records each target's source hash, saved artifact and report. It is not yet a portable corpus runner.

Next work is to close the EPA table-grid and page-flow gaps in the owning HTML/CSS/PDF path and make the PowerPoint table compact and readable. FWS needs a measured CSS-scope decision before conversion. The WAI, MDN, NASA and TodoMVC gaps and cross-platform budgets remain in the single [H10 roadmap](../../../../../Docs/ROADMAP.md#unfamiliar-page-conversion-qualification).
