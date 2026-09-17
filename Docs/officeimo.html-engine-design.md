# Independent HTML engine design

This architecture extends OfficeIMO's HTML engine into a reusable web-document platform. It describes ownership and acceptance contracts; package READMEs and the generated [HTML support matrix](officeimo.html-support-matrix.md) define current supported behavior. Remaining implementation work and its delivery order belong in the [roadmap](ROADMAP.md#independent-html-engine).

The implemented document foundation is documented in the [HTML package README](../OfficeIMO.Html/README.md#inspect-and-edit-owned-html): owned document/node and contextual-fragment snapshots, explicit edits, public DOM migration, and parser/charset provider contracts. `HtmlDocumentEngine` is the provider-neutral entry point, `OfficeIMO.Html.Core` is the dependency-free contract leaf, and `OfficeIMO.Html.AngleSharp` is the current HTML parser provider. The optional runtime adds persistent scripted sessions, bounded resources, fetch and asynchronous XMLHttpRequest, storage, modules, lifecycle events, mutation delivery, locators, DOM actions, waits and independent capture. CSS/layout and the runtime still use replaceable retained providers. Static-renderer qualification retains the established H4/representative corpus. The independently authored H4/advanced-held-out selection now has a checked-in acceptance manifest for screen, print, and screen-to-page output. It requires source-marker preservation and evaluates each selected capability against declared per-case geometry, page, text, pixel, clipping, padding, and alignment criteria on Windows, Linux, and macOS, with [retained visual evidence](../Build/Project/Evidence/2026-09-16/html-h4-visual-acceptance/README.md). Named classifications preserve accepted differences without turning Chromium into a universal oracle. The same workload has enforced cross-platform elapsed, allocation, process-tree memory, output-size, cold-to-warm determinism and cancellation ceilings with [retained calibration evidence](../Build/Project/Evidence/2026-09-15/html-h4-advanced-held-out/README.md). Wider standards coverage, further application qualification and dependency retirement remain open.

The objective is to parse, inspect, query, edit, style, lay out, render, and convert HTML through OfficeIMO-owned contracts. OfficeIMO converters, HtmlTinkerX, document readers, and preview hosts should consume the same implementation. Runtime dependencies can supply difficult algorithms initially; each must have an explicit replacement boundary and evidence for its eventual removal.

Delivery prioritizes usable components. AngleSharp and AngleSharp.Css are temporary implementation providers, not permanent parts of the target engine. Keep them, HarfBuzz, CodePages and other effective providers for as long as they help deliver usable behavior, but isolate them behind OfficeIMO-owned contracts and remove them from the default package graph when their replacements pass the declared gates. Dependency removal is a later qualification track, not a prerequisite for releasing useful rendering or runtime capabilities. A retired provider may remain in an optional adapter package for migration and differential testing; it must not shape the public model.

## Product boundaries

There are four separately useful products, with different acceptance criteria:

| Product | User outcome | Completion boundary |
| --- | --- | --- |
| HTML document platform | Parse and edit documents, query content, analyze CSS, normalize HTML, extract structured data, and convert to editable document formats | Published DOM, syntax, query, serialization, semantic and resource contracts work without launching a browser |
| Static web renderer | Render HTML and CSS to continuous previews, paged PDF, SVG and raster output | Declared static-web compatibility profile passes its frozen corpus without Chromium; script-dependent content is identified as unexecuted |
| Web runtime | Load pages, execute scripts, maintain browser state and produce live document revisions without an external browser | Declared language, web API, lifecycle, navigation, storage and isolation profiles pass their independent suites |
| Programmatic automation | Inspect structured page state, find actionable elements, interact, wait, extract, trace and capture through code or optional agent tools | Provider-neutral context, page, locator, observation, action and event contracts pass deterministic workflows on every advertised runtime provider |

A static page can be complex and still require no JavaScript. A simple-looking page can require scripts to create all its content. Classify inputs by required behavior and captured state, not their appearance or URL.

Playwright is an automation layer over Chromium, Firefox and WebKit. Rendering replacement and automation replacement therefore have different owners and tests. A new engine also cannot prove how a website behaves in those other browsers; cross-browser testing remains a valid development dependency even after the runtime is independent. See [Playwright browser support](https://playwright.dev/docs/browsers).

No stage promises all websites or an identical Playwright API. The browser stage is part of the long-term design, with its own executable scope rather than an indefinite promise hidden inside HTML-to-PDF. The default managed provider ultimately owns HTML, CSS, JavaScript and the selected web platform without third-party runtime packages or a browser binary; optional external providers remain explicit compatibility choices.

## Public adoption boundary

The HTML platform is public, MIT-licensed and usable independently of Word, Excel, PowerPoint and PDF. Office document packages are consumers of the same engine; they do not define its API or require a user to adopt the rest of OfficeIMO. Package names may retain the OfficeIMO family identity, but package descriptions, examples and dependency graphs must make standalone use clear.

External consumers can adopt the platform at the smallest layer that owns their outcome:

| Adoption mode | Consumer outcome | Required package boundary |
| --- | --- | --- |
| Web document | Parse, query, inspect, edit and serialize HTML without graphics, PDF, networking or a browser | `OfficeIMO.Html.Core`, using the selected parser provider until the managed parser becomes the default |
| CSS analysis | Tokenize and inspect stylesheets, match selectors, explain cascade results and report parsed-but-unsupported features | Owned syntax and style contracts; no layout or output encoder required |
| Static rendering | Resolve resources, compute styles, create continuous or paged layout and return a reusable display list | `OfficeIMO.Html` over Core and shared drawing primitives |
| Output encoding | Write PNG and other raster images, SVG, PDF, previews, geometry maps or hit-test data from one render result | A target adapter over the owned display list; target-specific policy stays explicit |
| Semantic conversion | Convert HTML structure into Word, Excel, PowerPoint, Markdown, RTF or another editable model with loss reporting | Thin format adapter over the owned DOM and semantic projection |
| Trusted application runtime | Execute a declared JavaScript and web-API profile, interact with the live document, then capture an owned snapshot or render result | Optional runtime package behind the same document, resource, readiness and capture contracts |
| Provider interoperability | Compare, import or temporarily execute through AngleSharp, a browser or another provider without exposing its objects to normal consumers | Optional adapter selected explicitly and identified in diagnostics |

The target is an independent alternative with its own coherent API, not an AngleSharp-compatible API clone. Existing AngleSharp consumers should receive a deliberate migration guide and, only where real demand exists, narrow import/export adapters. Source compatibility with provider-specific types would preserve the dependency in the public contract and is therefore outside the design.

Public adoption requires more than publishing assemblies. Each advertised layer needs a focused README, complete API documentation, a compatibility and versioning policy, runnable small examples, packed-package consumers, supported-target checks, trimming/AOT and browser-WASM evidence where claimed, package-size and transitive-dependency evidence, and diagnostics that identify the selected provider. A feature is advertised only at the layer whose qualification gate it passes.

## Existing foundation

The package and implementation inventory was refreshed against OfficeIMO source commit `c94e8b97d1e035c58e94bfa8952cb8f5fbf34344` and HtmlTinkerX source commit `04f77ddee2dd5d33bb381e5f921b440710492aad`. These identify the inspected source; they do not establish public-package or runtime qualification.

| Capability | Existing owner and evidence | Design implication |
| --- | --- | --- |
| Shared conversion input | [`HtmlConversionDocument`](../OfficeIMO.Html/Engine/HtmlConversionDocument.cs) retains source and lazily builds policy, logical, semantic, style and resource projections | Evolve this facade; preserve explicit trust and conversion profiles |
| HTML parser | [`HtmlDocumentEngine`](../OfficeIMO.Html/Engine/HtmlDocumentEngine.cs) exposes owned documents and contextual fragments; [`AngleSharpHtmlParser`](../OfficeIMO.Html.AngleSharp/AngleSharpHtmlParser.cs) implements the replaceable provider contract, including isolated foreign and ancestor-form context reconstruction | Qualify replacement parsers against the same owned document, fragment, recovery and resource-limit fixtures |
| CSS processing | [`HtmlComputedStyleEngine.Rules`](../OfficeIMO.Html/Styles/HtmlComputedStyleEngine.Rules.cs) combines AngleSharp.Css with raw-rule recovery and protected declaration transformations | Own lossless CSS syntax before adding more parser-specific recovery layers |
| Layout | [`HtmlRenderLayoutEngine`](../OfficeIMO.Html/Rendering/HtmlRenderLayoutEngine.cs) already has block, inline, table, flex, grid, positioning, pagination and related implementations | Audit and adapt existing algorithms behind clearer intermediate representations; avoid a second permanent renderer |
| Render result | [`HtmlRenderDocument`](../OfficeIMO.Html/Rendering/HtmlRenderDocument.cs) carries shared pages, fonts, text and diagnostics | Extend its scene and semantic mapping rather than add a PDF-specific layout model |
| Shared graphics and fonts | [`OfficeIMO.Core`](../OfficeIMO.Core/OfficeIMO.Core.csproj) contains the `OfficeIMO.Drawing` namespace; [`OfficeManagedTextShapingProvider`](../OfficeIMO.Core/OfficeManagedTextShapingProvider.cs) explicitly declines some scripts and features | Keep typography and codecs shared across formats; absence of an external package does not imply full shaping coverage |
| Optional shaping | [`OfficeIMO.Drawing.HarfBuzz`](../OfficeIMO.Drawing.HarfBuzz/OfficeIMO.Drawing.HarfBuzz.csproj) supplies native shaping dependencies | Retain the existing provider contract while expanding managed shaping |
| PDF | [`OfficeIMO.Html.Pdf`](../OfficeIMO.Html.Pdf/OfficeIMO.Html.Pdf.csproj) depends on Html, Core and Pdf | HTML owns layout; PDF owns PDF serialization, fonts, tags, security and conformance |
| Browser bridge | [`OfficeIMO.Html.Pdf.Browser`](../OfficeIMO.Html.Pdf.Browser/OfficeIMO.Html.Pdf.Browser.csproj) references HtmlTinkerX | Keep external browser execution in this optional direction; never make Html depend on its browser bridge |
| HtmlTinkerX | Its inspected project already references OfficeIMO.Html and OfficeIMO.Markdown.Html, alongside AngleSharp packages, Jint, Playwright and other tools | Migrate workflows incrementally through the existing consumer relationship; HTML-engine independence does not automatically remove every HtmlTinkerX dependency |

Current code is a substantial starting point. This inspection was not a new rendering qualification run. Existing feature names, catalog entries and passing historical tests must be classified as implementation, conformance evidence, or remaining proof gaps before setting a baseline.

## Dependency direction and packages

The target package graph is acyclic:

```mermaid
flowchart BT
    HC[OfficeIMO.Html.Core: existing DOM leaf, syntax extensions planned]
    OC[OfficeIMO.Core: existing drawing and document primitives]
    H[OfficeIMO.Html: styles, resources, semantics and layout] --> HC
    H --> OC
    A[OfficeIMO.Html.AngleSharp: existing transitional provider] --> HC
    P[OfficeIMO.Html.Pdf: existing adapter] --> H
    P --> PDF[OfficeIMO.Pdf]
    PDF --> OC
    T[HtmlTinkerX: extraction and automation workflows] --> H
    R[Optional web runtime and programmatic automation] --> H
    G[Optional agent, CLI and tool adapters] --> R
    T --> R
    B[OfficeIMO.Html.Pdf.Browser: existing external bridge] --> T
```

The target graph differs deliberately from the current packaged graph:

| Package | Current role | Target role |
| --- | --- | --- |
| `OfficeIMO.Html.Core` | Publishable MIT contract leaf with no package or project dependencies | Standalone DOM, HTML/CSS syntax, selector and managed parser foundation |
| `OfficeIMO.Html.AngleSharp` | Publishable temporary parser provider over Core | Optional migration and differential-testing adapter, absent from default dependencies |
| `OfficeIMO.Html` | Publishable renderer that currently references AngleSharp, AngleSharp.Css, Core and shared OfficeIMO primitives | Owned CSS, resources, semantics, layout and display list over Core and shared drawing primitives |
| `OfficeIMO.Html.Pdf` | PDF output adapter over Html and OfficeIMO.Pdf | Same thin output direction, consuming an explicit render profile |
| `OfficeIMO.Html.Pdf.Browser` | Optional HtmlTinkerX/browser bridge | Explicit external-browser provider outside the independent static profile |
| Optional runtime | Trusted scripted sessions and automation through retained providers | Provider-neutral context, page, observation and action contracts over an owned HTML/CSS/JavaScript runtime |
| Optional agent/tool adapters | Not yet a separately productized HTML-runtime surface | Thin model, CLI or MCP adapters over structured observations and actions, with no LLM dependency in the runtime |

`OfficeIMO.Html.Core` is the existing lightweight leaf for owned source buffers, DOM, serialization and provider contracts; HTML/CSS syntax and selector ownership continue to move into it as their implementations mature. It must not depend on graphics, PDF, Office file formats, networking sessions, JavaScript or browser installers. This gives HtmlTinkerX and other parsing or extraction consumers a small dependency graph. Namespaces can distinguish `Dom`, `Css.Syntax` and `Selectors` without creating a package for every folder.

Keep computed styles, semantic projection, resource orchestration and layout in `OfficeIMO.Html` initially. Extract a separate CSS package only if a measured consumer requirement justifies it. Existing graphics primitives remain in `OfficeIMO.Core`; do not create another font or image implementation under Html. Create optional runtime/provider packages only when an executable consumer is ready to use them.

The Html composition layer currently references the AngleSharp provider to supply the default parser. The provider translates into owned nodes and tokens; it does not control the render pipeline. It can remain the implementation across multiple usable releases while the managed parser is built and qualified. The default graph then drops the provider. `OfficeIMO.Html.AngleSharp` may remain as an optional compatibility and differential-testing adapter, but no default package or public consumer contract may require it. Add interop only for a demonstrated consumer requirement; do not preserve provider-specific public APIs by default.

HtmlTinkerX retains website workflows, extraction recipes, sessions, authentication integration and PowerShell/CLI surfaces. Shared DOM/CSS/rendering algorithms move to OfficeIMO. Inspect each remaining HtmlTinkerX dependency by capability: replacing HTML rendering alone does not replace JavaScript minification, DOM diffing, email inlining or every legacy parser surface.

The existing HtmlTinkerX package is an aggregate with browser dependencies. Merely calling a browserless method does not make its package graph lightweight. A future parsing/extraction package can consume the new leaf while an optional browser package supplies execution; retain the existing aggregate facade during a documented packaging migration. Validate this separately from the OfficeIMO engine graph.

HtmlForgeX remains the typed HTML authoring owner. It can generate fixtures and consume previews; the HTML engine must also handle independently authored pages. Do not optimize the architecture around HTML emitted only by Evotec libraries.

## One document, several representations

```mermaid
flowchart TD
    I[Bytes, text, URI or resource bundle] --> S[Source and resource session]
    S --> D[Owned DOM and CSS syntax]
    D --> Q[Query, edit, inspect and serialize]
    D --> E[Semantic projection]
    E --> N[Word, Excel, PowerPoint, RTF, Markdown and readers]
    D --> C[Computed styles]
    C --> L[Formatting boxes and intrinsic sizing]
    L --> F[Layout fragments: continuous or paged]
    F --> V[Display list plus semantic and source mapping]
    V --> O[PDF, SVG, raster, preview and hit testing]
    R[Optional runtime: DOM mutation and event loop] --> D
    R --> C
```

The DOM, semantic model, formatting boxes, layout fragments and display list serve different purposes. A DOM node can create no box, several boxes, or fragments on several pages. CSS explicitly defines a box tree separate from the element tree; these identities must not be conflated. See [CSS Display](https://www.w3.org/TR/css-display-3/).

Extend the existing logical, semantic and render models where their contracts fit. New formatting and fragment representations address missing responsibilities; they must not become competing versions of the same semantic document.

### Public operation shape

Preserve `HtmlConversionDocument` as the familiar high-level facade while replacing provider-specific inputs and outputs with owned contracts. The following operations describe the proposed API responsibilities, not new method names available today:

| Operation | Required public contract |
| --- | --- |
| Parse/load | Explicit text/bytes/fragment/URI source, encoding/base URI, immutable options, limits and parse report; asynchronous resource I/O stays explicit |
| Query/inspect | Owned node handles, selector context, source/metadata/resource evidence and diagnostics; no layout or network work unless requested |
| Edit | Scoped mutation session, revision and source-preservation rules; commit returns a new conversion snapshot |
| Prepare/render | Explicit media, viewport/page environment, fonts, resources, renderer/provider and loss policy; returns an owned reusable render result |
| Convert/export | Target-specific semantic, visual or hybrid profile; caller-owned stream/path contract and output report |
| Execute/interact | Optional runtime/session selection, isolated execution capabilities, deadlines, readiness condition and capture-state provenance |

Results distinguish native handling, approximation, omission, unsupported behavior, policy blocking and failure. Include selected provider versions, input/environment identity and resource outcomes. Keep capability inspection available before execution, but distinguish static preflight from runtime discovery. A successful parse must not imply that rendering or a chosen target will succeed.

Separate CPU work from asynchronous I/O; expose cooperative cancellation on both. Define reusable-result disposal and thread-safety explicitly. Use interfaces at real substitution boundaries such as parser providers, resource fetchers, text shaping and output sinks; do not make every internal layout node an extensibility point.

### Source and DOM contract

Own node identity, namespaces, attributes, text, comments, doctype, document mode, template contents, fragments and parse diagnostics. A fragment parse accepts an explicit context element. HTML parsing and XML/XHTML parsing are separate modes; do not treat malformed HTML as XML with permissive flags. HTML tokenizer states and tree construction follow the [HTML parsing standard](https://html.spec.whatwg.org/multipage/parsing.html).

Separate structural preservation from conversion policy. Parsing can retain script elements without executing them. Sanitization creates a policy-governed projection; it must not silently erase the source used for inspection. Normalized HTML is a deliberate export, not a trusted substitute for original bytes.

Use stable node IDs within a document revision and explicit source spans. Mark implied nodes and generated content as such. Retain original source optionally for exact unchanged writing and diagnostics; normal DOM serialization need not preserve quote style, whitespace or malformed source byte for byte. Editing invalidates affected source ranges. Define decoded-character and original-byte offsets separately when an encoding transform occurred.

An immutable document snapshot supports concurrent conversions. An editable document/session has one mutation owner, batches changes, increments a revision and invalidates dependent projections. Start with correct full recomputation per revision; add incremental invalidation after measured need. Node handles from another document or obsolete snapshot are rejected or explicitly remapped. Returned output retains the resource/font lifetime it requires and does not reference disposed pooled buffers.

Preserve tree semantics needed for later event dispatch, ranges and shadow trees without claiming those browser APIs already exist. The [DOM Standard](https://dom.spec.whatwg.org/) is the semantic reference. Do not implement the entire browser DOM interface surface just to replace one parser dependency.

### HTML and CSS parsing

The transitional HTML provider parses once and translates into owned structures. The current adapter retains its native tree alongside an owned document when owned nodes are requested so selectors and conversion can reuse the parsed representation; conversion-only parsing keeps the owned projection lazy. Measure and reduce that retained dual-graph cost before changing providers. Rendering and conversion must not repeatedly serialize and reparse provider DOMs. Keep any old API clone/export path explicit and outside the new fast path.

The managed HTML parser needs the full selected tree-construction contract: malformed tables and foster parenting, misnested formatting, implied elements, raw text/RCDATA, character references, templates, foreign SVG/MathML content, fragments, quirks mode, and streaming chunk boundaries. A tokenizer alone is insufficient. Decoding must handle the selected BOM, transport, label, replacement and HTML encoding rules, with limits enforced during parsing. Use the [Encoding Standard](https://encoding.spec.whatwg.org/) and HTML parsing rules as references.

Own CSS tokenization and a syntax tree that preserves unknown declarations, nested component values and source spans. Then implement declaration/property grammar, selector matching, cascade and computed values separately. This permits parsing an unsupported feature without treating it as rendered. Use [CSS Syntax](https://www.w3.org/TR/css-syntax-3/) for tokenization and error recovery.

Prioritize CSS replacement when the baseline proves that current recovery transformations block correctness or duplicate substantial grammar. Keep AngleSharp for HTML while owned CSS syntax replaces sentinel rewriting and raw-rule reconciliation. Remove old recovery code only after its regression fixtures pass through the new path.

### Styles and environment

Represent specified, cascaded, computed and used values separately. Preserve unresolved percentages, `auto`, intrinsic sizes and deferred calculations until their containing context is known. Support explicit user-agent styles, author/user origins, importance, specificity, layers, inheritance, custom properties, pseudo-elements and source order. Cascade precedence comes from [CSS Cascade](https://www.w3.org/TR/css-cascade-5/), with unsupported parts visible in the capability contract.

Use one selector implementation for queries and style matching, with different calling contracts where standards require them. Include document mode, namespace/context, pseudo-state and mutation revision in relevant cache keys. Preserve every declaration required for proper fallback; do not pick a value by source order before knowing whether its grammar is supported.

The render environment explicitly fixes screen/print media, viewport, device scale, page geometry, language, direction, font set and relevant preference/media values. Container queries require style/layout dependency handling and bounded convergence; they cannot be finalized entirely before layout. Freeze animation/time for static output using an explicit policy. Do not invent hover, scroll or animation state to make a page appear complete.

### Rendering intent and output profiles

Rendering intent is a public input, not an output-format side effect. A caller selecting PDF must not silently select print CSS, and a caller selecting PNG must not silently choose a viewport height. Define a render request through independent axes:

1. **Document state** selects original static source, an edited snapshot, an imported captured state or a live-runtime snapshot.
2. **CSS environment** selects media type, viewport, device scale, user preferences, language, direction, time and pseudo-state.
3. **Layout surface** selects a bounded viewport, a continuous canvas or paged sheets.
4. **Pagination policy** selects reflow, fragmentation, fixed-canvas slicing or element-aware placement.
5. **Encoder** selects a retained display list, raster image set, SVG, PDF or another target without changing the prior choices.

Named profiles provide useful defaults while leaving these axes visible. The first qualified set should cover:

The public request exposes independent media, surface, and pagination overrides.
A request that differs from its named profile remains executable when the
combination is coherent, but it reports unqualified coverage and does not inherit
the profile's declared provider evidence. Viewport and continuous surfaces require
no pagination; paged surfaces require fragmented reflow or fixed-canvas slicing.

| Profile | CSS and layout behavior | Typical result | Reference and acceptance evidence |
| --- | --- | --- | --- |
| Screen viewport | `screen` media at an exact viewport, clipped to its bounds | One raster/SVG surface, preview or hit-test map | Browser viewport screenshot plus geometry, text and overflow checks |
| Screen full page | `screen` media at a fixed width and content-driven continuous height | Full-page image, SVG or continuous display list | Browser full-page screenshot plus geometry and resource checks |
| Print paged | `print` media with page size, margins and fragmentation | Searchable multipage PDF, SVG/raster page set or retained page scenes | Browser print-to-PDF, PDF structure/readback and all-page visual checks |
| Screen media paged | `screen` media recomputed in a paged layout environment | PDF or page set that preserves screen styling while allowing page reflow | Browser PDF after explicit screen-media emulation plus pagination checks |
| Screen snapshot paged | One continuous screen layout frozen before fixed-canvas slicing or element-aware placement | PDF or page set matching the screen composition | Full-page browser screenshot, slice manifest and page-boundary checks |
| Continuous vector | Declared media on an unbounded vertical canvas | SVG, drawing scene, geometry map or downstream preview | Continuous browser capture plus vector/text/source-map checks |

These profiles serve different contracts. Print paged may change navigation, hide controls and apply `@page`; screen media paged preserves screen cascade but still reflows content across pages; screen snapshot paged preserves one screen layout and then places or slices it. No API or command should call all three simply “HTML to PDF.”

Multi-surface output is explicit. A raster or SVG request declares separate pages, a stitched canvas, one selected page or a bounded range. Archive-plus-manifest packaging is a separate output operation over that resolved page set, so storage does not alter layout or selection semantics. The report records surface dimensions, order, clipping, scale, background, pagination mode, provider identity and any fallback. Output encoders consume the same qualified display list and cannot rerun layout with private defaults.

Fixed-canvas slicing selects requested page indices before materializing projected
scenes. Projection is cancellable, removes nonintersecting nodes, assigns logical
text and navigation evidence to one retained slice, and is bounded by page, visual,
and surface limits. Stitched results retain a placement record for every source
page or slice. Parsing, resources, layout, projection, and encoding consume one
operation deadline.

### Text, sizing and fragmentation

Keep one shaping contract in the existing shared graphics owner. Layout receives glyph IDs, advances, offsets, clusters, logical-text mapping, fallback fonts and break constraints. Measurement and painting consume the same font data and shaping result. Start with optional HarfBuzz where needed, while improving the managed provider by script and font feature. [HarfBuzz shaping concepts](https://harfbuzz.github.io/shaping-concepts.html) and [cluster mapping](https://harfbuzz.github.io/clusters.html) explain the relevant input/output distinction.

Track bidirectional resolution, grapheme boundaries, line-break opportunities, shaping and final line fitting as distinct responsibilities. Re-shape at breaks where required. Include Latin, combining sequences, Arabic/Hebrew, Indic, CJK, vertical text and emoji in qualification; do not equate character count with width. Pin Unicode data and redistributable test fonts. Use [Unicode bidi](https://www.unicode.org/reports/tr9/) and [line breaking](https://www.unicode.org/reports/tr14/) tests for their respective algorithms.

Build formatting contexts around intrinsic minimum/maximum sizing, containing blocks, percentage resolution, margin behavior, anonymous boxes and logical axes. Adapt existing block, inline, table, flex and grid implementations one context at a time. Positioning, floats, replaced elements and stacking must agree on the same geometry. Unsupported combinations deserve interaction fixtures, not an ever larger collection of special cases in adapters.

Pagination is a layout constraint. Each layout operation accepts available inline/block space and a fragmentation context, then returns fragments plus an explicit continuation/break token. Page and column boundaries can change available space and child placement; cropping a tall screenshot cannot implement this. Define repeated headers, rowspans, widows/orphans, unbreakable oversized content, fixed/running elements, footnotes, page counters and named-page geometry. Iterative page-dependent content must converge within a declared limit or produce a diagnostic. See [CSS Fragmentation](https://www.w3.org/TR/css-break-3/).

### Painting and output contracts

Extend the existing render model into an ordered display list with text, paths, images, clipping, transforms, opacity groups, links and semantic/source IDs. Share stacking decisions across output backends. Retain logical reading order independently from paint order. Specify CSS pixels, device pixels and PDF-point conversion at the boundary, including the 96 CSS px to 72 PDF pt relationship at scale 1.

PDF writes searchable text, link geometry, embedded font data and supported structure tags from that scene. SVG and raster backends reuse it; SVG export must define whether text remains text or becomes paths, including portability consequences. Backend-specific unsupported effects produce a loss record or explicit rasterization policy with bounds and resolution. PDF/A, PDF/UA and PDF/X claims still require their own output validation.

Editable formats consume semantics first. A Word paragraph, Excel table or PowerPoint shape needs target-specific structure. Visual and hybrid profiles may use layout geometry, but must report approximated layout, rasterized content and reduced editability. Do not force every target through PDF or claim exact editable reconstruction from a display list.

### Use-case catalog and qualification

The engine supports several product families. Each has its own result, oracle and release gate; success in one family does not imply the others.

| Use case | Primary result | Required engine slice | Qualification basis |
| --- | --- | --- | --- |
| Parse, query and edit | Owned document snapshot, edits and serialized HTML | Source/encoding, HTML parser, DOM, selectors and serializer | Standards fixtures, source/DOM assertions, mutation invariants and round trips |
| Extraction and analysis | Structured values, matched nodes, links, metadata, style evidence and diagnostics | DOM, selectors, optional CSS and resource metadata | Independently authored corpora, deterministic results and bounded traversal |
| CSS tooling | Tokens, syntax tree, selector matches, cascade trace and computed values | CSS syntax, selectors, cascade and environment | Selected CSS suites, explainable winning declarations and lossless unsupported syntax |
| Sanitization and normalization | Policy-governed document plus exact change report | Source DOM, policy and serializer | Adversarial fixtures, explicit trust policy and structural reopen checks |
| Thumbnail and screenshot | Viewport or full-page raster/SVG output | Static resources, styles, layout and painting | Browser capture, geometry/text assertions and declared pixel tolerances |
| Browser-style print | Paged PDF or page images using print CSS | Print profile, fragmentation and output adapter | Browser print-to-PDF, independent PDF readback and all-page comparison |
| Screen-to-PDF | Paged PDF from screen-media reflow or frozen screen composition | Chosen screen-paged profile and PDF adapter | Explicit browser-media reference or screenshot/slice reference, never an implicit print comparison |
| Visual regression | Deterministic scenes, images and difference evidence | Frozen environment, rendering profiles and artifact manifest | Stable fixtures, controlled fonts/resources and explainable geometry/pixel deltas |
| Semantic document conversion | Editable Office, Markdown, RTF or other target with loss report | DOM and semantic projection, optionally layout geometry | Target-native reopen, structural assertions and declared approximation |
| Preview and editor host | Retained scene, hit-test map, source/semantic mapping and incremental revisions | Document/edit contracts, layout and display list | Interaction, lifetime, revision invalidation and accessibility/reading-order checks |
| Local or bundled site rendering | One or more outputs plus resource and navigation report | Resource session, URL policy, selected static/runtime profile | Deterministic bundles, resource limits, base-URI tests and capture provenance |
| Trusted scripted application | Live session, post-interaction snapshot and any qualified output | Optional runtime, selected web APIs, readiness and automation | Language/API suites and representative application workflows |

Email inlining, DOM diffing, JavaScript minification, web crawling policy and product-specific extraction recipes can consume the document platform, but do not automatically become HTML-engine responsibilities. Admit shared behavior only when its contract belongs at this layer and at least one real consumer exercises it.

Preview and automation hit testing can consume the same fragment geometry. Accessibility needs roles, names, language, relationships and reading order; bounding boxes alone do not provide an accessibility tree. Reuse and extend the existing accessibility owner within Html.

## Resource and execution contract

All entry points share one resource session and aggregate job budget. Pure parse/query operations perform no network access. Resource requests carry their kind, initiating node or rule, resolved URL, origin, allowed credentials, content identity and cancellation. Resolve stylesheets, imports, fonts, responsive images and nested SVG resources through that same session.

The request lifecycle is:

1. Snapshot options, source identity, policy, environment and deadlines.
2. Decode and parse under byte, node, depth and work limits.
3. Discover and resolve permitted resource dependencies under aggregate limits.
4. Resolve styles and fonts; compute semantics and/or layout as requested.
5. Validate the requested support/loss policy and construct the output.
6. Publish a completed artifact atomically where the destination supports it, then release job-owned resources.

Network access is explicit. Revalidate redirects and connected endpoints, including private-address policy and DNS changes. Define scheme, credentials, cookie, MIME, decompressed-byte, image-pixel, font and archive limits. A trusted document does not automatically authorize every network or local-file destination. Treat cancellation as cooperative inside the library; hostile hosted workloads also need externally enforced process memory/time limits.

Centralize web URL interpretation and use [WHATWG URL](https://url.spec.whatwg.org/) fixtures before claiming browser equivalence. A .NET URI object alone is not proof of identical web parsing semantics. Resource sessions partition credentialed caches and preserve caller stream ownership. No ambient authenticated browser profile is used implicitly.

Represent a captured page as a versioned resource bundle: root bytes or DOM state, URLs, response metadata, resource hashes, capture environment, fonts and resource failures. Snapshotting `outerHTML` alone misses relevant state such as canvas pixels, runtime form values and shadow content. Declare which state was retained and which was not.

Distinguish two paths: native layout of captured HTML/CSS/assets, and importing externally measured/computed browser state. The second can be useful for migration but does not prove the native CSS/layout engine. Any external-browser fallback is caller-selected and reported with provider identity; it must not silently grant network or execution permissions. HTML inspection cannot reliably detect every possible script dependency.

## Browser runtime evolution

The static engine is inert even when the input contains scripts. An optional OfficeIMO runtime owns the public session, execution, readiness and capture contracts. Its implementation supplies realms, DOM bindings, event dispatch, tasks/microtasks, timers, navigation, fetch, storage, origins and invalidation as each profile is qualified. A session has one authoritative live document. A transitional provider may retain that document in its native representation and produce independent OfficeIMO snapshots at an explicit capture boundary; do not maintain two independently mutable DOMs or expose provider nodes as public session results. A future owned live DOM replaces the provider implementation without changing the session contract.

The scripted-document workflow reaches an explicitly ready document, owned interaction and independent rendered or semantic output. `WebApplicationV1` adds bounded modules, fetch, asynchronous XMLHttpRequest, storage, observers, cross-document navigation, reload, history restoration, forms, lifecycle behavior and layout-aware automation while retaining session resource limits and locator handles. Same-origin frame documents run qualified classic scripts in isolated globals with shared session accounting, same-origin parent/top/frame access and bounded JSON-compatible messaging. Cross-origin or opaque sandbox frames, sandbox-blocked scripts and child module graphs remain inert. Frame documents are captured as a separate immutable tree and projected into clipped static screen, print and screen-to-page output. The named profile accepts explicitly trusted content and declares its unsupported credential, isolation, layout and input behavior. Retain effective HTML, CSS and JavaScript providers while advancing static CSS and rendering qualification; owned parser and interpreter replacement remain separate tracks. HtmlTinkerX can reuse the engine later, but its current helper APIs, package graph and workflows do not define the runtime's boundary or acceptance criteria.

A JavaScript interpreter can execute language code but does not supply a web platform. Evaluate the Jint dependency already present in HtmlTinkerX for the first bounded runtime, preserving a provider boundary. Language testing uses [Test262](https://github.com/tc39/test262); page behavior needs additional DOM and web API tests. Event scheduling follows the [HTML event-loop model](https://html.spec.whatwg.org/multipage/webappapis.html#event-loops).

Admit runtime features through executable profiles:

| Profile | Required behavior and evidence |
| --- | --- |
| Scripted local document | Selected language features, DOM mutation, events, timers, microtasks and deterministic capture; no unrestricted network or CLR exposure |
| `WebApplicationV1` | Modules and chosen fetch/XMLHttpRequest/storage/history/observer APIs, bounded document navigation, origin behavior, forms, lifecycle and layout invalidation; representative framework application fixtures |
| Interactive automation | Selected input dispatch, focus, scrolling, hit testing, locator semantics, actionability, explicit waits and cancellation; representative standalone OfficeIMO workflows pass on each advertised provider |
| Broader browser compatibility | Separately qualified shadow DOM/custom elements, canvas, workers, service workers, media, WebGL and other APIs as product scope expands |

Freeze time and disable unrelated animations in reproducibility tests; separately test real scheduling and animation behavior when those become supported. Define readiness using explicit lifecycle events, application predicates, resource/font completion and layout stability under a deadline. Network-idle alone cannot establish that a page is complete.

Any future untrusted-script profile must execute in an isolated worker with OS-enforced bounds. Interpreter constraints supplement process isolation; they are not the isolation boundary. Host objects are explicit capabilities and never expose arbitrary CLR access. Browser origins, cross-origin resource rules, credentials and navigation require dedicated security tests.

The public-page pilot has a separate admission path and provider identity. It may
reuse the owned host/page/action/capture contracts and the application-to-document
workflow, but it must never call the trusted worker launcher for public scripts.
Before admitting a page, the host must verify an isolation mechanism on the
current OS and return a report containing the mechanism, worker and policy
identities, enforced memory/CPU/process limits, filesystem and network scope,
and termination outcome. An unavailable or failed check rejects the job; there
is no best-effort fallback to `WebApplicationV1`.

For the first corpus, the host acquires bounded HTTP(S) input and resources and
passes immutable bytes to a worker with no direct network access. The acquisition
broker owns DNS/IP and redirect checks, response-size and time limits, origin
policy, content hashes and retained-source provenance. The worker can request
only broker-authorized resources; until that request channel is qualified, a
pilot page is limited to resources captured before execution. Credentials,
cookies and persistent storage are excluded from the first public profile.
Probe the isolation boundary with denied file reads and writes, denied network
calls, process spawning, memory pressure, CPU loops and cancellation before a
real public URL can enter the corpus. Each output names its exact page, inputs,
unsupported APIs and isolation report. The chosen OS mechanism is replaceable
behind the worker-launch boundary; its presence alone does not imply broad
browser compatibility.

Removing a third-party JavaScript engine is a later language-runtime project covering parsing, evaluation, modules, promises, built-ins, memory management, internationalization and performance. Start with an interpreter only if the stage has its own funded scope and Test262 acceptance. A JIT is not required to begin; its absence also does not establish adequate performance for real applications. The static product must remain useful and independently releasable throughout.

### Programmatic browser control and agent integration

The browser-control product is an API over the runtime, not an AI agent embedded in the engine. `OfficeIMO.Html.Runtime` already supplies the first provider-neutral foundation: persistent sessions, navigation, evaluation, locators, structured actions and failures, layout-backed actionability, waits and independent captures. Extend those owners instead of creating a separate automation DOM.

The complete public control model has six layers:

| Layer | Contract |
| --- | --- |
| Runtime host | Starts the managed or explicitly selected provider, reports capabilities and versions, enforces process/resource policy and owns shutdown |
| Browser context | Partitions origins, storage, cookies where qualified, permissions, credentials, downloads, resource caches and recording policy |
| Page | Owns one live document, URL/history, viewport, lifecycle, navigation, script evaluation, capture and page-scoped events |
| Locator and action | Re-resolves semantic/CSS/text/role queries against the current revision and performs bounded click, fill, select, focus, keyboard, pointer, scroll, drag and wait operations |
| Observation | Returns a revision-bound, size-limited page state containing URL/title, visible and actionable semantic tree, accessible names/roles, values, selected text, geometry, scroll state, diagnostics and optional screenshot references |
| Trace and result | Records navigation, resources, console/script failures, observations, actions, downloads, captures, timing, policy decisions and final structured output with redaction controls |

Context and page types are future API responsibilities, not names promised by the current package. Keep one-page sessions valid while introducing multiple pages and isolated contexts. Locator queries remain durable descriptions; element references emitted in an observation are bound to its document revision and fail as stale rather than acting on a different element after navigation or mutation.

Programmatic use is the primary contract. A .NET application can navigate, observe, query, interact, extract and capture with typed requests and results without a model. Declarative workflow and replay support build on the same commands and events. CLI, PowerShell, MCP and other hosts serialize those contracts without inventing their own locator, readiness, security or error semantics.

An optional agent adapter adds a replaceable planning loop:

```text
goal + policy
    -> bounded observation
    -> planner proposes one typed action
    -> runtime validates capability, authority and current revision
    -> action executes and emits events
    -> new observation or structured result
```

The planner may be an LLM, rules engine or application callback. It receives only the selected observation and tool schemas; it does not gain direct access to interpreter objects, arbitrary CLR APIs, credentials or unrestricted network functions. The runtime package has no dependency on a model SDK. Optional model adapters map provider responses into the same action union and preserve model, prompt, token, cost and decision provenance where the caller requests it.

The structured tool surface should include navigation, page observation, locator inspection, click/fill/select/check/focus, keyboard/pointer/scroll, explicit waits, script evaluation when policy permits it, extraction, screenshot/render capture, downloads and final-result submission. Every tool declares required capabilities and returns stable status codes. Whole-page observation must support semantic-only, visual-only and combined modes so a caller can trade cost against fidelity without changing action semantics.

The same automation contract supports three providers during migration:

1. The OfficeIMO managed runtime, which becomes the dependency-free default after H6 and H9.
2. An explicitly selected external-browser adapter for websites outside the managed compatibility profile.
3. A consumer-supplied provider implementing the published host/page/action contracts.

Provider switching is never silent. Capability inspection occurs before a workflow starts, every observation and trace records the provider, and unsupported behavior returns a typed result. Cross-provider tests exercise equivalent outcomes rather than assuming identical screenshots, timing or hidden browser behavior.

A local programmatic library is the first product boundary. Hosted browser fleets, residential proxies, stealth behavior, CAPTCHA services, account/profile synchronization and recurring task infrastructure are separate deployment products. They can consume the runtime and automation contracts later but do not belong in the HTML/CSS/JavaScript engine or determine when it is independent.

## Independent parser and adoption strategy

OfficeIMO can compete as a web-document and rendering platform before it replaces every provider. It cannot claim to be an independent parser/DOM alternative while the default parsing path still requires AngleSharp. Use four explicit adoption stages:

| Stage | Honest product claim | Exit gate |
| --- | --- | --- |
| Owned contract over a provider | Consumers use OfficeIMO document, query, edit and conversion contracts while diagnostics identify AngleSharp as the parser | No provider objects or exceptions escape normal APIs; packed consumers prove the owned surface |
| Standalone web-document product | Non-Office applications can adopt the small public package and complete supported parse/query/edit/serialize workflows | Public docs, examples, compatibility policy, target/platform proof and package evidence are complete |
| Managed parser by default | OfficeIMO owns tokenization, tree construction, contextual fragments, decoding and serialization in the default path | Pinned conformance, recovery, hostile-input, streaming, performance and differential gates pass |
| Independent static distribution | The default static renderer has no third-party runtime packages or browser binaries | CSS, parsing, encoding, typography, resource, layout and output profiles pass through the packed graph |

The managed parser gate covers the selected WHATWG behavior as an integrated system: tokenizer states, tree construction and repair, fragments with context, templates, quirks/document modes, SVG and MathML foreign content, character references, encoding detection and restart policy, streaming chunk boundaries, serialization, source locations and resource limits. Passing a tokenizer suite alone is insufficient.

The public document product also needs declared selector and DOM scope. Publish the supported selector levels and pseudo-classes, mutation and snapshot rules, collection behavior, namespace handling and source-preservation guarantees. Browser-only APIs such as layout properties, event dispatch, custom elements and shadow DOM belong to separately qualified runtime profiles; they must not appear as empty compatibility members just to resemble another DOM API.

Use standards fixtures as the primary contract and differential testing as a diagnostic tool. Compare the managed parser with the temporary provider and independent browser trees, minimize disagreements and resolve them against the chosen standards scope. Add grammar-aware fuzzing, round-trip and mutation fuzzing, allocation and retained-graph measurements, deep/wide/adversarial limits, cancellation latency, trimming/AOT checks and supported-platform runs. Benchmark equivalent parse, query, edit and serialization workloads without changing the requested outcome for one implementation.

Provider retirement is a packaging change only after the public contract is stable. Keep `OfficeIMO.Html.AngleSharp` as an optional migration, comparison or fallback adapter if users need it, version it separately where practical, and require explicit provider selection. Do not silently fall back to it when the managed parser rejects or limits input; return a typed unsupported or failure result so users can make the policy decision.

## Compatibility profiles and specification governance

The engine does not target “twenty years of browsers” as one compatibility mode. Modern HTML parsing rules already define interoperable recovery for much old and malformed markup; remaining historical behavior is admitted by observed content classes and explicit profiles. Do not emulate a browser release, engine brand or collection of undocumented quirks unless a separately funded compatibility provider owns that exact outcome.

Every released profile pins a standards-and-evidence manifest. Living standards and test suites continue to change, so a bare link to their latest page is insufficient for reproducible support claims. The manifest records the specification URL and immutable revision or published snapshot, selected sections/features, upstream test repository commit, included and excluded test paths, OfficeIMO fixture revision, required providers, platforms and output profiles.

Use five compatibility bands:

| Band | Scope | Admission rule |
| --- | --- | --- |
| Standards-required parsing and recovery | Tokenization, tree construction, quirks/document mode, DOM basics, URL and encoding behavior required to interpret ordinary HTML | Required by the selected document profile and gated by pinned conformance fixtures |
| Stable deployed web | Interoperable HTML, CSS, DOM, JavaScript and web APIs needed by representative current documents and applications | Add only with standards references, multi-engine evidence where useful and an adopted consumer workflow |
| Legacy content | Obsolete elements/attributes, presentational hints, common historical encodings, malformed tables and document patterns still found in archived or enterprise content | Preserve or map through the current HTML processing model and a provenance-tracked, legally usable legacy corpus; diagnose unsupported behavior |
| Optional modern capability | Newer, specialized or expensive CSS and web APIs such as advanced containment, web components, canvas, media or workers | Separate opt-in profile until specifications, implementation and workload evidence are stable |
| Proprietary or platform behavior | Vendor-prefixed features, browser extensions, plugins, ActiveX, browser chrome, operating-system integration and undocumented quirks | Explicitly unsupported by the managed profile or supplied by an optional compatibility provider |

The first manifest family separates the following claims; final public identifiers should follow the existing profile naming conventions:

The `v1` suffixes below version observable public behavior after release. They are not corpus generations or preview labels. Qualification corpora use descriptive names such as `baseline-report`, `representative` and `advanced-held-out`; a profile version changes only when its compatibility contract requires it.

| Profile contract | Includes | Does not imply |
| --- | --- | --- |
| Web document v1 | Decoding, HTML parsing/recovery, owned DOM, selected selectors, mutation and serialization | CSS rendering, networking or script execution |
| Static screen v1 | Web document plus declared CSS cascade, screen layout, painting, resources and screen output profiles | Print pagination, live JavaScript or browser automation |
| Paged print v1 | Web document plus print cascade, paged layout, fragmentation and qualified paged outputs | Screen appearance or editable target reconstruction |
| Scripted document v1 | One trusted document, selected JavaScript, DOM bindings, events, tasks and deterministic capture | Navigation, credentials or general web-application compatibility |
| Web application v1 | Selected navigation, URL/origin, modules, fetch, asynchronous XMLHttpRequest, storage, history, forms, observers, lifecycle and layout invalidation | Every browser API, hostile-script safety or cross-browser equivalence |
| Programmatic automation v1 | Versioned observation, locator, action, wait, trace and capture contracts over a declared runtime profile | Autonomous planning or a bundled model provider |

Support is recorded per processing stage. Recognizing or preserving syntax is different from computing it, and computing a value is different from laying it out or making it interactive. Each capability declares the applicable outcomes from this set:

1. Source and decoding: bytes are decoded under a named encoding policy and source evidence is retained.
2. Parse and preserve: syntax is recognized, recovered and serializable without implying behavior.
3. DOM and query: owned nodes, namespaces, mutations and selectors have the declared semantics.
4. Cascade and compute: applicable CSS declarations produce declared computed values.
5. Layout: boxes, intrinsic sizes, lines, fragmentation and geometry follow the selected profile.
6. Paint and output: the display list and each named encoder preserve the declared visual or semantic result.
7. Runtime and interaction: script, events, navigation, APIs and automation produce the declared state transitions.

Evolve the existing `HtmlRenderCapabilityCatalog` into the package-wide executable source of truth instead of creating another registry. A capability record needs a stable ID, owning profile/version, specification references and pinned revisions, applicable processing stages, exact supported subset, support outcome, implementation provider, test-manifest evidence, platform/output scope, limits, fallbacks and diagnostics. Generate the package support matrix, runtime capability inspection and website documentation from that owner.

Do not overload one support label with several meanings. Track coverage (`Qualified`, `Partial`, `Unsupported`, `Unqualified`), handling (`Native`, `Preserved`, `Fallback`, `Ignored`, `Rejected`) and maturity (`Required`, `Optional`, `Experimental`) independently. Provider ownership is another field: provider-backed behavior can be qualified while still remaining scheduled for managed replacement.

The initial standards families are:

| Family | Versioning authority | Qualification source |
| --- | --- | --- |
| HTML parsing and elements | Pinned revision of the [WHATWG HTML Living Standard](https://html.spec.whatwg.org/) | html5lib parser fixtures, selected WPT, owned recovery/streaming/bounds cases and independent DOM comparison |
| DOM and events | Pinned [WHATWG DOM](https://dom.spec.whatwg.org/) revision plus explicitly selected HTML-defined interfaces | Selected WPT, mutation/event ordering fixtures and owned snapshot/edit invariants |
| URL, origin and encoding | Pinned WHATWG [URL](https://url.spec.whatwg.org/) and [Encoding](https://encoding.spec.whatwg.org/) revisions with declared Unicode/IDNA data | Upstream fixture data, redirect/origin/resource-policy cases and byte-level decoding evidence |
| CSS | A chosen [W3C CSS Snapshot](https://www.w3.org/TR/css/all/) as an index, with every implemented module and section pinned separately | Selected WPT/reftests, computed-value assertions, geometry fixtures and output comparisons |
| JavaScript language | One pinned ECMA-262 edition or revision; ECMA-402 is separate | Pinned [Test262](https://github.com/tc39/test262) manifest, resource limits, host-integration cases and explicit exclusions |
| Web APIs | One specification and selected interface/algorithm set per API | Selected WPT plus integrated application, lifecycle, security and cancellation cases |
| Text and fonts | Pinned Unicode algorithms/data and declared OpenType/font behavior | Unicode conformance data, redistributable font corpus, glyph/cluster/geometry checks and cross-platform output |

CSS snapshots organize many independently versioned modules and do not establish browser adoption by themselves. A profile therefore names exact modules and sections rather than claiming “CSS3” or all of a snapshot. JavaScript conformance similarly reports the pinned Test262 totals by required, passed, failed, excluded and untested feature; a passing selected subset is not advertised as the entire language.

Use this oracle order when evidence disagrees:

1. The pinned normative specification and applicable errata.
2. A pinned conformance test whose assumptions match the OfficeIMO profile.
3. A minimized independently authored fixture compared across current browser engines.
4. A documented compatibility decision for deployed content, with its scope and diagnostic.

No single browser screenshot, provider result or passing test suite defines the whole contract. Browser references remain essential evidence for deployed layout and application behavior, while standards determine whether matching one browser would preserve or reproduce a bug.

Legacy qualification starts with common content, not old engine emulation: HTML4 and XHTML-shaped markup delivered as HTML, obsolete presentational elements and attributes, standards/limited-quirks/quirks document modes, common legacy encodings, malformed tables, early CSS box/table patterns and inline event handlers required by selected runtime profiles. Actual XML/XHTML parsing, namespaces, well-formedness and MIME-type behavior require a separate qualified profile. Maintain a separate corpus with provenance, usage rights and expected preservation, semantic, layout and diagnostic outcomes. Frames, plugins, ActiveX, VBScript, browser toolbars, IE document modes and proprietary DOM/CSS behavior remain unsupported until an explicit provider/profile decision admits them.

Standards upgrades are deliberate releases. Refresh upstream revisions in a dedicated change, review changed tests and algorithms, classify new failures, update the capability manifest and preserve the preceding profile for users who require reproducibility. Additive capability may extend a profile version; changed observable behavior or removed fallback requires a new profile or documented breaking release according to the public compatibility policy.

## Incremental delivery and promotion

The complete managed browser is not the unit of delivery. A release unit is one useful vertical capability slice through the owned contracts: for example contextual fragment parsing, a selector family, one CSS module through computed values and layout, a fragmentation rule across every output, a navigation lifecycle, or one automation action. Provider-backed implementation is acceptable at any stage when the provider is isolated, reported and replaceable. This lets OfficeIMO improve current document, rendering and automation workflows while managed replacements mature independently.

Use a promotion ladder for each capability and profile combination:

| State | User availability | Admission rule |
| --- | --- | --- |
| Incubating | Integration builds and internal differential runs | Owned contract and diagnostics exist; behavior may change and is not advertised as supported |
| Experimental opt-in | Explicit preview profile, option or provider selection | Bounded implementation, representative fixtures, known limitations and no effect on stable defaults |
| Qualified opt-in | Released but caller-selected profile or capability | Pinned manifest, declared platforms/outputs, regression and resource gates, package/docs evidence and typed unsupported behavior pass |
| Stable default | Selected by the normal supported profile | Existing stable-profile regressions pass, migration impact is resolved, cross-platform evidence is complete where applicable and the former provider remains explicitly selectable when a compatibility need exists |
| Retired or superseded | Versioned compatibility path where justified | Replacement and migration policy are published; removal follows the normal breaking-release contract |

Promotion changes the catalog record and generated documentation; it does not require redesigning the public document, render, runtime or automation contracts. Never use an environment-dependent silent fallback to make an experimental implementation appear stable. The caller either selects the provider/profile or receives the provider and fallback decision in the result.

Every slice must protect both the established behavior and its new claim. Run the existing stable-profile corpus first, then the slice's normative and regression fixtures, relevant provider-differential checks, supported output/platform checks, resource and cancellation budgets, packed-consumer checks and generated capability-matrix verification. Visual work compares semantic and geometry invariants alongside pixels so harmless raster differences do not block delivery and structural regressions do not hide inside a pixel threshold.

Breaking cleanup remains allowed, but batch it around durable owned contracts and record user action in `MIGRATION.md`. Once a named profile is published as stable, ordinary additions must preserve its qualified behavior. Observable standards changes, removed fallbacks and incompatible default changes use a new profile version or a documented breaking release rather than quietly changing an old manifest.

The long-running engine branch is an integration and full-stack qualification line, not a release gate. Keep it current with the default branch and use it to prove interactions among dependent slices. When a slice reaches its admission gate, replay it on the current default branch as a focused, reviewable change and merge it through the normal OfficeIMO path. Bring the merged result back into the integration line. Do not wait for H6 or H9, and do not merge the accumulated integration branch wholesale unless its remaining diff is itself cohesive and qualified.

## Dependency retirement

### Retained-provider decision baseline

The September 2026 provider baseline and the subsequent owned-document budget gate make three separate decisions. They apply to the retained-provider implementation and must be revisited when a replacement passes the corresponding removal gate.

| Provider area | Current decision | Evidence and next trigger |
| --- | --- | --- |
| AngleSharp HTML parser and static adapter | Retain as the default parser behind `HtmlDocumentEngine`; do not start an HTML parser rewrite without a measured product or qualification trigger | Weak provider projections reduced directly comparable Windows owned-document retention from 638.9 KiB to 204.3 KiB and conversion native-plus-owned retention from 641.9 KiB to 207.0 KiB. Parse, query, edit, serialization and conversion budgets pass at 10, 100 and 1,000 rows on Windows, Linux and macOS; separate in-flight cancellation workloads cover 10,000, 25,000 and 100,000 rows. Reconsider the default only when an owned parser passes the selected HTML recovery, fragment, encoding, hostile-input and performance gates. |
| AngleSharp.Css syntax provider | Retain for selectors, cascade and computed values; use the owned syntax tree for source-preserving syntax consumers and future migration slices | The first owned lane preserves exact source, trivia, unknown rules and declarations, nested component values, source spans and invalid-input recovery. On the 105-rule Windows workload it takes a 5.400 ms median versus 15.070 ms for AngleSharp.Css syntax, while its richer retained graph is 2,241.2 KiB versus 691.0 KiB. Add property grammar and migrate selected selector/cascade slices before changing the default style path. |
| Retained AngleSharp DOM runtime fork | Retain for the qualified trusted runtime while treating it as a separate maintenance liability | The fork is substantially larger than the static adapter and is pinned independently for runtime hooks. Track upstream version lag, local patches, conformance and rebase cost. Prefer upstreamable hooks or a replaceable runtime DOM before expanding the fork; static-parser replacement alone does not retire it. |

Package and source evidence accompanies the baseline in `Build/Project/Evidence/2026-09-15/html-provider-decision`. The enforced current budgets and provider delta are in `Build/Project/Evidence/2026-09-15/html-owned-document-budgets`. Dependency updates remain ordinary qualified maintenance. A newer upstream release is a reason to test and update the retained provider, not by itself a reason to replace or freeze it.

### Replaceable provider contracts

Replaceability means that the engine can use another implementation without changing its public document model or rewriting consumers. The algorithm behind that boundary can still be substantial work.

| Dependency | Boundary to establish now | What remains owned by OfficeIMO |
| --- | --- | --- |
| AngleSharp | Parse source/fragment into owned nodes and parse diagnostics | Document identity, mutation, serialization policy, query contract and consumer-facing types |
| CSS parser | The owned lossless syntax boundary is established; add declaration/property grammar and isolate temporary selector/cascade implementations behind the next owned contracts | Source retention, cascade decisions, computed-value contract and capability reporting |
| HarfBuzz | Existing shared shaping request/result contract | Fonts, glyph/cluster/logical-text mapping, lifetime, fallback policy and layout consumption |
| CodePages | Central encoding resolution/decoding boundary using owned metadata and suitable BCL types | Label policy, BOM/HTML precedence, errors, streaming behavior and limits |
| Script interpreter | Runtime provider bound to owned host/DOM contracts | Web API behavior, origins, scheduling policy, resource access and process isolation |

Do not expose provider nodes, native handles or provider exception types as normal public results. Translate errors and define cancellation, disposal and resource ownership at the boundary. Keep provider setup and configuration in composition/provider code; do not scatter encoding registrations or library-version checks through consumers. Existing .NET/BCL contracts do not need replacement wrappers merely to remove a NuGet dependency.

Use capability-specific contracts rather than mirroring every third-party method. Prove the first real implementation through contract fixtures; later run replacement providers against the same fixtures. A fake second parser does not establish replaceability. For CodePages, test actual legacy byte decoding and unavailable-encoding behavior; for shaping, test real glyph/cluster output and lifetime. Package separation follows real deployment and consumer needs, not one package per interface.

### Removal gates

Here, independence means no third-party runtime packages or browser binaries in the advertised default profile. The .NET runtime/BCL, OS services and clearly declared font/Unicode/data resources remain. Test tools and optional providers are separate. If a stronger goal excludes OS text, TLS, codecs or all external data, it requires a different platform scope.

| Dependency | Short-term role | Replacement/removal gate |
| --- | --- | --- |
| AngleSharp.Css | Current selector, cascade and computed-style implementation while owned syntax and grammar mature | Selected declaration/property grammar, selector, cascade-trace and computed-value fixtures pass through the owned path; existing protected-token/raw-recovery paths are removed |
| AngleSharp HTML/DOM | Production parser behind owned contracts | Selected full HTML parsing corpus, fragment/encoding tests, resource bounds and performance gates pass through a qualified replacement |
| Encoding.CodePages | Current legacy encoding support | Explicit encoding compatibility matrix passes without it, or it remains only in a separately selected encoding provider; never silently narrow accepted inputs |
| HarfBuzz/native typography | Optional complex shaping | Managed script/font/feature corpus meets text, cluster, glyph and geometry requirements on supported platforms |
| Playwright/browser binaries | Explicit external execution/capture provider and independent test reference | Static runtime graph excludes them first; later selected interactive workloads pass on the owned runtime |
| Jint or another script provider | Optional runtime implementation | Owned language runtime passes its versioned language suite and integrated web-runtime corpus within resource/performance budgets |
| Other HtmlTinkerX tools | Formatting, inlining, alternate parsers, diffing and related workflows | Separate owner and replacement proof for each capability; no blanket removal based on rendering progress |

Check complete packed transitive graphs, runtime native assets, browser downloads and data requirements for each supported target. `PrivateAssets` does not remove a runtime need. Do not describe a package as dependency-free just because its project file contains only project references.

Breaking API changes are accepted for this program. Replace public `IHtmlDocument`/`IElement` contracts during the foundation milestone once owned alternatives have end-to-end proof, while AngleSharp remains the parser implementation. Land each coherent break when its replacement is usable and migrate all in-repository consumers in the same change; do not hold every cleanup for one final engine rewrite. Remove superseded paths and do not carry duplicate APIs, obsolete overloads or compatibility wrappers solely to avoid an approved break. Preserve existing document/conversion behavior unless an intentional behavior change is separately documented and validated. Record actual upgrade actions in `MIGRATION.md` and use an appropriate breaking release version; removing the parser dependency itself can happen much later without another public API change.

## Qualification and failure analysis

Expand the existing Html tests, capability gallery, real-world corpus tooling and HTML/PDF artifact gates. Do not build a competing evidence system. Every fixture must identify its input/resource hashes, source/license, environment, engine/provider versions, tested output contract and known exceptions. Pin standards/test-suite revisions for qualification; evaluate upstream changes deliberately.

| Evidence layer | What it proves | Acceptance rule |
| --- | --- | --- |
| HTML parser tests | Token/tree construction, fragments and recovery | All tests in the declared supported subset pass; exclusions are named and counted |
| CSS/DOM tests | Syntax, queries, cascade, computed values and mutation semantics | Expected structures/values match independently defined fixtures |
| Layout fixtures | Box sizes, intrinsic constraints, line breaks, fragmentation and stacking | Geometry and text assertions pass across the selected environment matrix |
| Reference rendering | Supported pages look correct | Per-fixture tolerances plus inspection at normal reading size; no hidden clipping, missing content or accepted reference regression |
| Output round trips | Artifacts contain usable content | Reopen with independent readers; check text, links, resources, tags and target editability as applicable |
| Adversarial tests | Bounds and trust behavior | Budget exhaustion/cancellation is controlled; output is not partially published; no unauthorized access |
| Consumer packages | The contract is usable beyond project references | Real packed consumers restore, build and run on every claimed target; check transitive dependencies |
| Runtime/browser tests | Language and interaction behavior | Explicit language/API subsets and representative workflows pass; unsupported features remain visible |

Use [html5lib parser fixtures](https://github.com/html5lib/html5lib-tests) and selected [Web Platform Tests](https://web-platform-tests.org/writing-tests/) as independent inputs, respecting licenses and harness requirements. Many WPT tests require JavaScript or a browser environment. Adapt static fixtures with preserved provenance where appropriate and label them as adapted; do not report adapted tests as the complete official WPT run. Rendering comparisons can follow [WPT reftest](https://web-platform-tests.org/writing-tests/reftests.html) relationships and declared fuzzy bounds.

Use pinned Chromium references for common deployed behavior and another engine where ambiguity matters. Resolve disagreements through the specification and a minimized case; Chromium output is not the definition of every print feature. Keep browser-rendered references separate from outputs produced by the native engine and from browser-state imports.

Capture real pages with their assets and fonts for deterministic CI. Sample independently authored documentation, articles, shops/catalog pages, email, invoices, tables, dashboards, forms and multilingual documents. Script-heavy samples belong in both an unexecuted-input lane and a captured-state/runtime lane, with different expectations. Track interaction coverage such as nested grid/table pagination; isolated feature counts do not establish combination coverage.

Report passed, failed, excluded and untested separately. Maintain a held-out corpus to detect overfitting. A proposed static-release threshold is 100% of the mandatory regression/conformance subset plus at least 95% of the frozen representative-page corpus accepted against predeclared semantic and visual criteria. This is a planning target, not measured coverage or a claim that 95% of the web works. Name every remaining failed case. Adjust the target only as an explicit scope decision, never by quietly removing inconvenient fixtures.

For visual qualification, fix font bytes, viewport, device scale, media, language, page size, background and capture state. Combine image comparison with geometry and content checks: high image similarity can hide missing text, and harmless antialiasing can lower a pixel score. Inspect every page for small acceptance fixtures and use declared sampling plus automated all-page checks for large documents.

The H4/advanced-held-out acceptance gate applies that model to three distinct intents. `screen-full-page-v1` compares the completed screen surface with Chromium full-page capture. `print-paged-v1` compares OfficeIMO's fragmented print sheets with Chromium print-to-PDF. `screen-snapshot-paged-v1` compares fixed-size pages with the completed OfficeIMO screen display list, because browsers do not expose that product contract directly. That gate requires sequential page numbers, uniform fixed canvas dimensions, the expected page count and exact final-page padding before comparing stitched pixels. The fixed-slicing profile may split elements, clip horizontal overflow to the canvas, and pad the last page. Screen markers require an exact normalized phrase; PDF extraction uses an explicitly reported ordered-token policy so table reading order can interleave cells without accepting words in a different order. The typography fixture also records a bounded managed-layout deviation: a vertical float currently reserves its full float-only formatting context before later block flow instead of wrapping that flow beside the float. Cross-block float wrapping remains open work.

Measure cold/warm time, per-stage elapsed time, allocation, peak/retained memory, output size, cancellation latency and scaling. Include parse-only, query-only, semantics-only, continuous layout and long paged documents. Calibrate hard budgets from the initial baseline on named machines; do not invent a throughput promise before measuring. Cache by document revision, CSS, resource/font identity, environment and provider version, with bounded retention.

Diagnostics should locate the earliest divergent stage: source/token, DOM, winning declaration, computed value, box, line, fragment or paint operation. Extend the existing report with stable node/source identity, requested feature, selected provider, outcome and affected target. No-warning output means only that no known problem was reported; it is not an independent fidelity certificate.

## Current implementation direction

The owned document boundary is the stable consumer surface. Continue widening real document and runtime workflows through it rather than exposing provider nodes or creating product-specific DOMs. Keep HtmlTinkerX and other hosts thin; migrate them only when the owning OfficeIMO component is usable and available through an explicit package or source relationship.

Strengthen style, layout, conversion and runtime components through OfficeIMO-owned contracts while the temporary providers remain behind them. The lossless CSS syntax tree is now the owned input for future declaration/property grammar and selector/cascade slices; it is not yet the renderer's default style path. Move HTML parser replacement ahead when measured recovery or capability failures obstruct a required workflow. Advance managed typography in the shared graphics owner. Treat static rendering, the selected application profile and dependency retirement as separately qualified outcomes.

The owned render request, six named profiles, explicit page selection and retained
result now form the common static output boundary. Existing continuous image calls
map to screen-full-page, paged image calls and existing PDF calls map to print-paged,
and output adapters cannot privately change CSS media or pagination. Executable
surface views now provide Drawing previews, source-coordinate mapping and bounded
clip-aware hit testing. PNG/SVG page archives package those same resolved surfaces
with deterministic names, hashes, source placements, provider identity and
diagnostics. The manifest separates retained HTML diagnostics from per-page encoder
diagnostics, includes source-to-target provenance, and treats bounded scale, font,
or codec fallback as archive loss. Container adapters such as MHTML attach their
input-boundary diagnostics to the retained result before packaging.

The table formatting context now carries fixed-width visible descendant contributions into auto track sizing while leaving percentage descendants dependent on the resolved cell width. It classifies header and footer repetition from computed table-group display, removes page boundaries crossed by rowspans or row and group avoidance, admits authored row breaks, and finds common line boundaries for oversized multi-cell rows. Table, row, and cell structure keys survive slicing and repetition, so retained scenes and tagged PDF represent one logical structure across page fragments. An oversized atomic or avoided row uses the shared forced-fragment diagnostic rather than stalling pagination.

Deliver the remaining work in reviewable vertical slices:

1. Add selected declaration/property grammar over the owned CSS syntax tree, then migrate selector, cascade-trace and computed-value slices with exact conformance manifests.
2. Close remaining H4 layout, fragmentation and output gaps profile by profile against frozen browser, geometry, semantic and artifact references.
3. Complete H5 as an integrated managed parser and serializer, switch the default only after conformance, bounds and performance gates pass, and retain the adapter only where real migration demand exists.
4. Productize the existing H7-H8 runtime and locator foundation as typed contexts, pages, observations, actions, events, traces and optional agent tools while retained providers remain effective.
5. Complete H6 by removing AngleSharp, AngleSharp.Css and other third-party runtime dependencies from the advertised static graph, with packed transitive-graph proof on every supported target.
6. Complete H9 by replacing the JavaScript and remaining runtime providers over the H6 engine. The resulting managed HTML/CSS/JavaScript runtime has no third-party runtime package or browser binary; broader browser compatibility continues through explicit profiles.

Preserve explicit bounds, provider identity and unsupported results throughout. Competitive claims attach to the completed profile or adoption stage, never to the repository as a whole.

The open milestones and their readiness gates are maintained once in the [roadmap](ROADMAP.md#independent-html-engine). A static-engine release, an AngleSharp-free package and an interactive-runtime release are separate reviewable outcomes. Calendar forecasts should follow the first slice and the measured remaining failure classes. Broad browser compatibility remains the largest and least predictable part of the program.
