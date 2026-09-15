# H4 advanced held-out visual acceptance

Status: **Passed**  
Acceptance configuration SHA-256: `b347bb8e73ee67bd7130a460de8ee5bd4d3e336ff15cdc9f2756e02a3632a8ed`

Chromium provides screen and print reference observations, not a universal oracle. OfficeIMO screen-to-page is compared with the completed OfficeIMO screen display list. Every selected case and capability must pass each applicable intent; intentional sheet, clipping, and font differences are admitted only through explicit bounded criteria.

## Case results

| Case | Screen | Print | Screen-to-page |
| --- | --- | --- | --- |
| `subgrid-operations` | Passed (`browser-reference-tolerance`) | Passed (`browser-reference-tolerance`) | Passed (`profile-consistency`) |
| `nested-fragmentation-report` | Passed (`browser-reference-tolerance`) | Passed (`browser-reference-tolerance`) | Passed (`profile-consistency`) |
| `columns-footnotes-brief` | Passed (`browser-reference-tolerance`) | Passed (`browser-reference-tolerance`) | Passed (`profile-consistency`) |
| `named-pages-brochure` | Passed (`browser-reference-tolerance`) | Passed (`browser-reference-tolerance`) | Passed (`profile-consistency`) |
| `stacking-clipping-board` | Passed (`browser-reference-tolerance`) | Passed (`browser-reference-tolerance`) | Passed (`profile-consistency`) |
| `typography-specimen` | Passed (`bounded-managed-layout-deviation`) | Passed (`browser-text-extraction-variance-and-bounded-managed-layout-deviation`) | Passed (`profile-consistency`) |
| `professional-print-catalog` | Passed (`browser-reference-tolerance`) | Passed (`intentional-sheet-expansion`) | Passed (`profile-consistency`) |
| `legacy-portal` | Passed (`browser-reference-tolerance`) | Passed (`browser-reference-tolerance`) | Passed (`profile-consistency`) |

## Capability results

| Profile | Capability | Intents | Selected cases | Status |
| --- | --- | --- | ---: | --- |
| `paged-print-v1` | `bidi-text` | print-paged-v1 | 1 | Passed |
| `paged-print-v1` | `css-backgrounds` | print-paged-v1 | 2 | Passed |
| `paged-print-v1` | `css-borders-effects` | print-paged-v1 | 1 | Passed |
| `paged-print-v1` | `generated-content` | print-paged-v1 | 1 | Passed |
| `paged-print-v1` | `html-source-decoding` | print-paged-v1 | 1 | Passed |
| `paged-print-v1` | `layout-columns` | print-paged-v1 | 2 | Passed |
| `paged-print-v1` | `layout-grid` | print-paged-v1 | 3 | Passed |
| `paged-print-v1` | `layout-positioning` | print-paged-v1 | 2 | Passed |
| `paged-print-v1` | `layout-tables` | print-paged-v1 | 1 | Passed |
| `paged-print-v1` | `paged-footnotes` | print-paged-v1 | 1 | Passed |
| `paged-print-v1` | `paged-fragmentation` | print-paged-v1 | 3 | Passed |
| `paged-print-v1` | `paged-page-rules` | print-paged-v1 | 2 | Passed |
| `paged-print-v1` | `text-flow` | print-paged-v1 | 2 | Passed |
| `paged-print-v1` | `vertical-text` | print-paged-v1 | 1 | Passed |
| `static-screen-v1` | `bidi-text` | screen-full-page-v1, screen-snapshot-paged-v1 | 1 | Passed |
| `static-screen-v1` | `css-backgrounds` | screen-full-page-v1, screen-snapshot-paged-v1 | 2 | Passed |
| `static-screen-v1` | `css-borders-effects` | screen-full-page-v1, screen-snapshot-paged-v1 | 1 | Passed |
| `static-screen-v1` | `generated-content` | screen-full-page-v1, screen-snapshot-paged-v1 | 1 | Passed |
| `static-screen-v1` | `html-source-decoding` | screen-full-page-v1, screen-snapshot-paged-v1 | 1 | Passed |
| `static-screen-v1` | `layout-columns` | screen-full-page-v1, screen-snapshot-paged-v1 | 2 | Passed |
| `static-screen-v1` | `layout-grid` | screen-full-page-v1, screen-snapshot-paged-v1 | 3 | Passed |
| `static-screen-v1` | `layout-positioning` | screen-full-page-v1, screen-snapshot-paged-v1 | 2 | Passed |
| `static-screen-v1` | `layout-tables` | screen-full-page-v1, screen-snapshot-paged-v1 | 1 | Passed |
| `static-screen-v1` | `text-flow` | screen-full-page-v1, screen-snapshot-paged-v1 | 2 | Passed |
| `static-screen-v1` | `vertical-text` | screen-full-page-v1, screen-snapshot-paged-v1 | 1 | Passed |

## Visual review artifacts

### subgrid-operations

- Screen: [OfficeIMO](subgrid-operations/officeimo-screen.png), [Chromium](subgrid-operations/chromium-screen.png), [difference](subgrid-operations/screen-difference.png)
- Screen-to-page: [difference](subgrid-operations/screen-to-page-difference.png)
- Print page 1: [OfficeIMO](subgrid-operations/officeimo-print-page-1.png), [Chromium](subgrid-operations/chromium-print-1.png), [difference](subgrid-operations/print-chromium-page-1-difference.png)

### nested-fragmentation-report

- Screen: [OfficeIMO](nested-fragmentation-report/officeimo-screen.png), [Chromium](nested-fragmentation-report/chromium-screen.png), [difference](nested-fragmentation-report/screen-difference.png)
- Screen-to-page: [difference](nested-fragmentation-report/screen-to-page-difference.png)
- Print page 1: [OfficeIMO](nested-fragmentation-report/officeimo-print-page-1.png), [Chromium](nested-fragmentation-report/chromium-print-1.png), [difference](nested-fragmentation-report/print-chromium-page-1-difference.png)
- Print page 2: [OfficeIMO](nested-fragmentation-report/officeimo-print-page-2.png), [Chromium](nested-fragmentation-report/chromium-print-2.png), [difference](nested-fragmentation-report/print-chromium-page-2-difference.png)
- Print page 3: [OfficeIMO](nested-fragmentation-report/officeimo-print-page-3.png), [Chromium](nested-fragmentation-report/chromium-print-3.png), [difference](nested-fragmentation-report/print-chromium-page-3-difference.png)

### columns-footnotes-brief

- Screen: [OfficeIMO](columns-footnotes-brief/officeimo-screen.png), [Chromium](columns-footnotes-brief/chromium-screen.png), [difference](columns-footnotes-brief/screen-difference.png)
- Screen-to-page: [difference](columns-footnotes-brief/screen-to-page-difference.png)
- Print page 1: [OfficeIMO](columns-footnotes-brief/officeimo-print-page-1.png), [Chromium](columns-footnotes-brief/chromium-print-1.png), [difference](columns-footnotes-brief/print-chromium-page-1-difference.png)

### named-pages-brochure

- Screen: [OfficeIMO](named-pages-brochure/officeimo-screen.png), [Chromium](named-pages-brochure/chromium-screen.png), [difference](named-pages-brochure/screen-difference.png)
- Screen-to-page: [difference](named-pages-brochure/screen-to-page-difference.png)
- Print page 1: [OfficeIMO](named-pages-brochure/officeimo-print-page-1.png), [Chromium](named-pages-brochure/chromium-print-1.png), [difference](named-pages-brochure/print-chromium-page-1-difference.png)
- Print page 2: [OfficeIMO](named-pages-brochure/officeimo-print-page-2.png), [Chromium](named-pages-brochure/chromium-print-2.png), [difference](named-pages-brochure/print-chromium-page-2-difference.png)
- Print page 3: [OfficeIMO](named-pages-brochure/officeimo-print-page-3.png), [Chromium](named-pages-brochure/chromium-print-3.png), [difference](named-pages-brochure/print-chromium-page-3-difference.png)

### stacking-clipping-board

- Screen: [OfficeIMO](stacking-clipping-board/officeimo-screen.png), [Chromium](stacking-clipping-board/chromium-screen.png), [difference](stacking-clipping-board/screen-difference.png)
- Screen-to-page: [difference](stacking-clipping-board/screen-to-page-difference.png)
- Print page 1: [OfficeIMO](stacking-clipping-board/officeimo-print-page-1.png), [Chromium](stacking-clipping-board/chromium-print-1.png), [difference](stacking-clipping-board/print-chromium-page-1-difference.png)

### typography-specimen

- Screen: [OfficeIMO](typography-specimen/officeimo-screen.png), [Chromium](typography-specimen/chromium-screen.png), [difference](typography-specimen/screen-difference.png)
- Screen-to-page: [difference](typography-specimen/screen-to-page-difference.png)
- Print page 1: [OfficeIMO](typography-specimen/officeimo-print-page-1.png), [Chromium](typography-specimen/chromium-print-1.png), [difference](typography-specimen/print-chromium-page-1-difference.png)

### professional-print-catalog

- Screen: [OfficeIMO](professional-print-catalog/officeimo-screen.png), [Chromium](professional-print-catalog/chromium-screen.png), [difference](professional-print-catalog/screen-difference.png)
- Screen-to-page: [difference](professional-print-catalog/screen-to-page-difference.png)
- Print page 1: [OfficeIMO](professional-print-catalog/officeimo-print-page-1.png), [Chromium](professional-print-catalog/chromium-print-1.png), [difference](professional-print-catalog/print-chromium-page-1-difference.png)
- Print page 2: [OfficeIMO](professional-print-catalog/officeimo-print-page-2.png), [Chromium](professional-print-catalog/chromium-print-2.png), [difference](professional-print-catalog/print-chromium-page-2-difference.png)

### legacy-portal

- Screen: [OfficeIMO](legacy-portal/officeimo-screen.png), [Chromium](legacy-portal/chromium-screen.png), [difference](legacy-portal/screen-difference.png)
- Screen-to-page: [difference](legacy-portal/screen-to-page-difference.png)
- Print page 1: [OfficeIMO](legacy-portal/officeimo-print-page-1.png), [Chromium](legacy-portal/chromium-print-1.png), [difference](legacy-portal/print-chromium-page-1-difference.png)

