# Isolated application and portable-font evidence

This packet qualifies three retained application workflows through the public-page
host broker and the network-disabled renderer: a form with six structured actions,
a long table, and an external classic/module/JSON graph with two actions. Each case
was acquired through the same host authorization and public-address checks, captured
inside a fresh rootless Podman container, and rendered as screen PNG, print PDF, and
screen-to-page PDF.

The exact source candidate was `52cdd759e4e23f323c62836d7b7a14d3a0cb2f0e`.
The image was built from the pinned .NET runtime base and pinned Fontconfig package
and has immutable ID
`sha256:74ef1d7acbf513426f8b1511fd1118c30a40ca3c41cc2828cd8f942191c7a4df`.
The [image inspection](oci-image-inspect.json) retains its OCI configuration and
layer identity. The [whole-workflow summary](controlled-summary.json) binds that
image to these published payloads:

| Payload | SHA-256 |
| --- | --- |
| Renderer entry DLL | `f8407c4a8a500b9b27695b8b97ab2a7ed8236853ac8f5dcefe68dacc045fdf18` |
| Renderer directory | `f22294ac8d16537a2f189fb6bb10c3150b6884b02f0b4cc64736607810abd68e` |
| Script worker entry DLL | `762a042715e5373f56f8e6d745d45046cadf2c42e6e1957ae4052dada2b8c4d7` |
| Script worker directory | `958c40a10e2ed8eb174980942c630f4bc05e9d47cb60352de5c1e756466db6c2` |

The run passed all 22 deterministic render cases, all six controlled acquisition
cases, all three application cases, and both optional live acquisition observations.
The controlled cases include same-host and approved cross-host redirects, DNS
rebinding rejection, declared-response size rejection, a dynamic POST redirect, and
a dynamic CORS exchange. The live observations recorded the then-current
`httpbingo.org` redirect and OPTIONS/POST preflight path; they are evidence of that
run rather than a deterministic service guarantee. Every created container was
removed, including the hostile-script and bounded-error cases.

## Portable font package

The host and worker independently accepted font package
`officeimo-browser-compact-2026.08`. Its manifest SHA-256 is
`0dcbc5ee021736d9526647d307ca52886cd74898b3da410247b13440f6281e46`;
the complete declared font/license package digest is
`ebc659487e86bcb221c802849038c645bec4907756e0b28dd0e5904685aa34e7`.
The retained [manifest and license files](font-package/) bind each of the seven font
files to its SHA-256 and license. The image installs exactly those seven fonts after
removing the base image's DejaVu files. OfficeIMO activates the same reusable
`HtmlPortableBrowserFontProfile` for layout, screen output, and embedded PDF fonts.

## Application observations

| Case | Acquired resources | Actions | Retained outputs | Visual observation |
| --- | ---: | ---: | --- | --- |
| Form | 3 | 6 | [screen](applications/forms/screen.png), [print](applications/forms/print.pdf), [screen-to-page](applications/forms/screen-to-page.pdf) | Values, checkbox, disabled policy field, button, and captured-state panel are aligned and legible. |
| Table | 3 | 0 | [screen](applications/tables/screen.png), [print](applications/tables/print.pdf), [screen-to-page](applications/tables/screen-to-page.pdf) | All 42 rows, the amount column, and qualified total `7333` remain visible without clipping or overlap. |
| External graph | 7 | 2 | [screen](applications/external-graph/screen.png), [print](applications/external-graph/print.pdf), [screen-to-page](applications/external-graph/screen-to-page.pdf) | Classic script, module graph, JSON data, action state, and total `51` render with consistent portable typography. |

The three screenshots above were inspected at original resolution after the final
run. This evidence qualifies these retained workflows and output contracts. It does
not claim general browser equivalence, arbitrary navigation, or cross-origin frame
execution.

## Static-renderer budget

The unchanged eight-case H4 selection passed all 76 outputs per iteration from clean,
commit-addressable source. The [Windows report](budget-windows-passed.json) records
an 8.15-second cold process, a 3.27-second slowest warm iteration, 1.096 GB peak
working set, matching cold/warm fingerprints, and cancellation in 4.54 ms. All
checked-in ceilings remained unchanged.
