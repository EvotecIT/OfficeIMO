# Isolated document navigation evidence

This packet qualifies four retained application workflows through the public-page
host broker and network-disabled renderer: a form with six structured actions, a
long table, an external classic/module/JSON graph with two actions, and a document
navigation workflow with fifteen actions. Each case was acquired through the same
host authorization and public-address checks, captured inside a fresh rootless
Podman container, and rendered as screen PNG, print PDF, and screen-to-page PDF.

The exact source candidate was `5a396758bc44cc272becdfe7f149eac51487eedb`.
The image was built from the pinned .NET runtime base and pinned Fontconfig package
and has immutable ID
`sha256:7532fc839b64e3e3312c4526bec95c100978dfe669e143952975c7c58e61ea9b`.
The [image inspection](oci-image-inspect.json) retains its OCI configuration and
layer identity. The [whole-workflow summary](controlled-summary.json) binds that
image to these published payloads:

| Payload | SHA-256 |
| --- | --- |
| Renderer entry DLL | `2a9881e11cf614ffb95bb0ffcc3353ce356fffbd5b4159ace2c144747bfd5e3b` |
| Renderer directory | `d629f59919f76c433f29e2e90f870f9a3b09db30ee773628a3436d412173ad7c` |
| Script worker entry DLL | `9ce49ffaf07ca5388564cf4a3749b307f2e160764bd2e4834175dfadc51be9af` |
| Script worker directory | `ab4d098b83c4f579a23979089b7007709d8378e55c02c7c12943902ba54f776f` |

The run passed all 22 deterministic render cases, all six controlled acquisition
cases, and all four application cases. Every created container was removed,
including the hostile-script and bounded-error cases. Live acquisition was not
requested, so mutable public DNS or service state is absent from this gate.

## Navigation contract

The navigation application starts at an acquired index document, follows an exact
top-level navigation transcript, changes history state, traverses backward and
forward, reloads the selected entry, and finishes at
`http://navigation.officeimo.test/report/approved`. The summary retains the exact
initiator, requested and final URLs, reduced referrer, final origin, redirect-taint
state, method, occurrence, history index, and replacement decision for each of the
four document loads. It also records seven host requests, seven validated public
connections, the fifteen structured actions, lifecycle observations, viewport
restoration, and the final capture manifest.

Dynamic requests retain their exact initiating document origin. Host authorization
happens before a connection is opened, including after a cross-document navigation.
Redirect referrers apply the qualified `strict-origin-when-cross-origin` default and
supported response `Referrer-Policy` reductions without reintroducing a suppressed
value. Navigation and dynamic-origin grants remain exact scheme/host/port origins.

## Portable font package

The host and worker independently accepted font package
`officeimo-browser-compact-2026.08`. Its manifest SHA-256 is
`0334abe014199558a69e9678daa6a5f572b278342d73497fbd53cdb5875051d1`;
the complete declared font/license package digest is
`befe8f7ee4c57493032846ede2816abe8eb008e6c850b813fd5073ffbe4d2bbe`.
The retained [manifest and license files](font-package/) bind each of the seven font
files to its SHA-256 and license. The image installs exactly those seven fonts after
removing the base image's DejaVu files.

## Application observations

| Case | Resources | Actions | Retained outputs | Visual observation |
| --- | ---: | ---: | --- | --- |
| Form | 3 | 6 | [screen](applications/forms/screen.png), [print](applications/forms/print.pdf), [screen-to-page](applications/forms/screen-to-page.pdf) | Values, checkbox, disabled policy field, button, and captured-state panel remain aligned and legible. |
| Table | 3 | 0 | [screen](applications/tables/screen.png), [print](applications/tables/print.pdf), [screen-to-page](applications/tables/screen-to-page.pdf) | All 42 rows, the amount column, and qualified total `7333` remain visible without clipping or overlap. |
| External graph | 7 | 2 | [screen](applications/external-graph/screen.png), [print](applications/external-graph/print.pdf), [screen-to-page](applications/external-graph/screen-to-page.pdf) | Classic script, module graph, JSON data, action state, and total `51` render with consistent portable typography. |
| Navigation | 7 | 15 | [screen](applications/navigation/screen.png), [print](applications/navigation/print.pdf), [screen-to-page](applications/navigation/screen-to-page.pdf) | The approved route and controls remain intact. Screen capture restores the viewport marker; print omits the screen-only marker; screen-to-page preserves screen styling at page size. |

The navigation outputs were inspected at original resolution after the final run.
This evidence qualifies these retained workflows and output contracts. It does not
claim general browser equivalence, arbitrary navigation, additional browsing
contexts, or cross-origin frame execution.

## Static-renderer budget

The unchanged eight-case H4 selection passed all 76 outputs per iteration from
clean, commit-addressable source. The [Windows report](budget-windows-passed.json)
records a 7.19-second cold process, a 2.40-second slowest warm iteration, 1.107 GB
peak working set, matching cold/warm fingerprints, and cancellation in 2.98 ms. All
checked-in ceilings remained unchanged.
