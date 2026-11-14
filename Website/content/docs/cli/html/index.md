---
title: "Render HTML to PDF and Images"
description: "Render HTML, MHTML and ZIP site bundles to PDF or PNG and SVG page images with officeimo html, with explicit resources and limits."
layout: docs
---

`officeimo html` renders a web page, an MHTML archive or a ZIP site bundle with the OfficeIMO HTML engine. `convert` writes a PDF; `render` writes page images in a ZIP archive with a manifest. Neither command starts a browser or reads files and URLs the page refers to unless you supply them.

## Convert HTML to PDF

```shell
officeimo html convert invoice.html --output invoice.pdf --base-uri https://example.com/billing/
officeimo html convert invoice.html --base-uri https://example.com/billing/   # writes invoice.pdf next to the source
officeimo html convert archive.mhtml --output archive.pdf
officeimo html convert site.zip --output site.pdf
```

Without `--output` the PDF is written next to the source with a `.pdf` extension. An existing output is refused unless you pass `--force`. Pass `--output -` to write the PDF to standard output.

`--base-uri` takes an absolute `http`, `https` or `file` URI. Relative links in the page are resolved against it, so a link such as `terms.html` becomes a working link in the PDF. Pass it whenever the page links to other pages with relative addresses; a conversion that cannot resolve a relative link fails with exit code `5`.

Read the page from standard input with `-` and name its format:

```shell
cat invoice.html | officeimo html convert - --input-format html --base-uri https://example.com/billing/ --output - > invoice.pdf
```

## Render page images

```shell
officeimo html render invoice.html --encoder png --output invoice-pages.zip --base-uri https://example.com/billing/
officeimo html render invoice.html --profile screen-full-page --encoder svg --viewport-width 1280 --output invoice-screen.zip --base-uri https://example.com/billing/
```

The archive holds one image per page under `pages/`, such as `pages/page-0001.png`, and a `manifest.json` that describes the pages. `--encoder` is `png` (the default) or `svg`. `--scale <factor>` scales the rendered output.

## Profiles and pages

`--profile` chooses how the page is laid out:

| Profile | Layout | `convert` | `render` |
|---|---|:---:|:---:|
| `print-paged` (default) | Print media, paginated like a printed page | Yes | Yes |
| `screen-media-paged` | Screen media, paginated | Yes | Yes |
| `screen-snapshot-paged` | Screen layout sliced into pages | Yes | Yes |
| `screen-viewport` | One viewport-sized image | | Yes |
| `screen-full-page` | The whole page as one image | | Yes |
| `continuous-vector` | One continuous vector page | | Yes |

`--viewport-width` and `--viewport-height` set the CSS viewport in pixels for the screen profiles. `--pages` selects `all`, one page such as `2`, a range such as `1-3`, or `stitched`, which places every page on one continuous surface.

`officeimo html capabilities` lists the profiles and the CSS and HTML features the engine supports, with their qualification status. Add `--format json` for a machine-readable report.

## Supply stylesheets and fonts

Local and remote resource reads are disabled. Data URIs and resources inside an MHTML or ZIP input are used; everything else you pass explicitly:

```shell
officeimo html convert invoice.html --output invoice.pdf --base-uri https://example.com/billing/ --stylesheet brand.css
officeimo html convert invoice.html --output invoice.pdf --base-uri https://example.com/billing/ \
  --font-family Body --font-regular fonts/Body-Regular.ttf --font-bold fonts/Body-Bold.ttf
```

`--stylesheet` can be repeated up to 16 times. `--font-family` names a family that the page's CSS uses, and `--font-regular`, `--font-bold`, `--font-italic` and `--font-bold-italic` supply its faces. A stylesheet or font cannot be the output path.

## ZIP site bundles

A ZIP archive is read as a small static site. The tool opens the root `index.html`, then `index.htm`, then the only HTML file in the archive; name another entry with `--entry-path`. It resolves links and resources between the archive's files and never reaches the network. For ZIP input, `--base-uri` must be an `http` or `https` directory URI ending in `/`; it becomes the virtual address of the archive.

`--max-bundle-entries` (2,048 by default), `--max-bundle-entry-bytes` (16 MiB) and `--max-bundle-decoded-bytes` (128 MiB) bound what the archive may expand to. `--max-input-bytes` bounds the ZIP file itself.

## PDF/UA groundwork

```shell
officeimo html convert report.html --output report.pdf --base-uri https://example.com/docs/ \
  --font-family Body --font-regular fonts/Body-Regular.ttf --pdf-ua-language en-US
```

`--pdf-ua-language` sets the document language and turns on PDF/UA tagging. It needs embedded fonts with full Unicode mappings, so supply the faces with the `--font-*` options; without them the conversion fails and explains which requirement is missing. The tool reports PDF/UA readiness on standard error. It does not claim conformance without passing external validator evidence.

## Limits

| Limit | Default | Option |
|---|---|---|
| Input size | 64 MiB | `--max-input-bytes` |
| Pages | 10,000 | `--max-pages`, up to 100,000 |
| Image archive size (`render`) | 1 GiB | `--max-archive-bytes` |
| Manifest size (`render`) | 4 MiB | `--max-manifest-bytes` |

`officeimo html --help` prints the complete option list for your version.

## Related

- [Convert documents](/docs/cli/convert/) for HTML as one of many conversion routes
- [HTML rendering and conversion](/docs/html/render-and-convert/)
- [Scripting and CI](/docs/cli/scripting-and-ci/)
