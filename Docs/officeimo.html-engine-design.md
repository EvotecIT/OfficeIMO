# Independent HTML engine design

This architecture extends OfficeIMO's HTML engine into a reusable web-document platform. It describes ownership and acceptance contracts; package READMEs and the generated [HTML support matrix](officeimo.html-support-matrix.md) define current supported behavior. Remaining implementation work and its delivery order belong in the [roadmap](ROADMAP.md#independent-html-engine).

The implemented document foundation is documented in the [HTML package README](../OfficeIMO.Html/README.md#inspect-and-edit-owned-html): owned document/node snapshots, explicit edits, public DOM migration, and parser/charset provider contracts. `OfficeIMO.Html.Core` is the dependency-free contract leaf and `OfficeIMO.Html.AngleSharp` is the current HTML parser provider. The optional runtime adds persistent scripted sessions, bounded resources, fetch, storage, modules, lifecycle events, mutation delivery, locators, DOM actions, waits and independent capture. CSS/layout and the runtime still use replaceable retained providers. Static-renderer qualification now includes the versioned H4/v1 document corpus and immutable source-browser references in conversion-consistency bundles. Contextual fragments, wider conformance coverage, further application qualification and dependency retirement remain open.

The objective is to parse, inspect, query, edit, style, lay out, render, and convert HTML through OfficeIMO-owned contracts. OfficeIMO converters, HtmlTinkerX, document readers, and preview hosts should consume the same implementation. Runtime dependencies can supply difficult algorithms initially; each must have an explicit replacement boundary and evidence for its eventual removal.

Delivery prioritizes usable components. AngleSharp and AngleSharp.Css are temporary implementation providers, not permanent parts of the target engine. Keep them, HarfBuzz, CodePages and other effective providers for as long as they help deliver usable behavior, but isolate them behind OfficeIMO-owned contracts and remove them from the default package graph when their replacements pass the declared gates. Dependency removal is a later qualification track, not a prerequisite for releasing useful rendering or runtime capabilities. A retired provider may remain in an optional adapter package for migration and differential testing; it must not shape the public model.

## Product boundaries

There are three separately useful products, with different acceptance criteria:

| Product | User outcome | Completion boundary |
| --- | --- | --- |
| HTML document platform | Parse and edit documents, query content, analyze CSS, normalize HTML, extract structured data, and convert to editable document formats | Published DOM, syntax, query, serialization, semantic and resource contracts work without launching a browser |
| Static web renderer | Render HTML and CSS to continuous previews, paged PDF, SVG and raster output | Declared static-web compatibility profile passes its frozen corpus without Chromium; script-dependent content is identified as unexecuted |
| Browser runtime and automation | Load pages, execute scripts, interact, wait for specified states, inspect results, and capture output | Declared web API and interaction profile passes independent runtime tests and representative site workflows without an external browser |

A static page can be complex and still require no JavaScript. A simple-looking page can require scripts to create all its content. Classify inputs by required behavior and captured state, not their appearance or URL.

Playwright is an automation layer over Chromium, Firefox and WebKit. Rendering replacement and automation replacement therefore have different owners and tests. A new engine also cannot prove how a website behaves in those other browsers; cross-browser testing remains a valid development dependency even after the runtime is independent. See [Playwright browser support](https://playwright.dev/docs/browsers).

No stage promises all websites or an identical Playwright API. The browser stage is part of the long-term design, with its own executable scope rather than an indefinite promise hidden inside HTML-to-PDF.

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
| HTML parser | [`AngleSharpHtmlParser`](../OfficeIMO.Html.AngleSharp/AngleSharpHtmlParser.cs) implements the owned `IHtmlParserProvider` contract and returns `HtmlDocument`; native parsing remains an internal provider helper | Qualify replacement parsers against the owned document contract and the same recovery fixtures |
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
    R[Optional web runtime] --> H
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
| Optional runtime | Trusted scripted sessions and automation through retained providers | Separately qualified runtime whose providers can be retired without changing the static graph |

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

The transitional HTML provider parses once, translates into owned structures, then releases its temporary graph. Measure the peak cost of having both graphs during conversion. Rendering and conversion must not repeatedly serialize and reparse provider DOMs. Keep any old API clone/export path explicit and outside the new fast path.

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

| Profile | CSS and layout behavior | Typical result | Reference and acceptance evidence |
| --- | --- | --- | --- |
| Screen viewport | `screen` media at an exact viewport, clipped to its bounds | One raster/SVG surface, preview or hit-test map | Browser viewport screenshot plus geometry, text and overflow checks |
| Screen full page | `screen` media at a fixed width and content-driven continuous height | Full-page image, SVG or continuous display list | Browser full-page screenshot plus geometry and resource checks |
| Print paged | `print` media with page size, margins and fragmentation | Searchable multipage PDF, SVG/raster page set or retained page scenes | Browser print-to-PDF, PDF structure/readback and all-page visual checks |
| Screen media paged | `screen` media recomputed in a paged layout environment | PDF or page set that preserves screen styling while allowing page reflow | Browser PDF after explicit screen-media emulation plus pagination checks |
| Screen snapshot paged | One continuous screen layout frozen before fixed-canvas slicing or element-aware placement | PDF or page set matching the screen composition | Full-page browser screenshot, slice manifest and page-boundary checks |
| Continuous vector | Declared media on an unbounded vertical canvas | SVG, drawing scene, geometry map or downstream preview | Continuous browser capture plus vector/text/source-map checks |

These profiles serve different contracts. Print paged may change navigation, hide controls and apply `@page`; screen media paged preserves screen cascade but still reflows content across pages; screen snapshot paged preserves one screen layout and then places or slices it. No API or command should call all three simply “HTML to PDF.”

Multi-surface output is explicit. A raster or SVG request declares separate pages, a stitched canvas, a selected page or an archive plus manifest. The report records surface dimensions, order, clipping, scale, background, pagination mode and any rasterized fallback. Output encoders consume the same qualified display list and cannot rerun layout with private defaults.

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

The scripted-document workflow reaches an explicitly ready document, owned interaction and independent rendered or semantic output. `WebApplicationV1` adds bounded modules, fetch, storage, observers, cross-document navigation, reload, history restoration, forms, lifecycle behavior and layout-aware automation while retaining session resource limits and locator handles. The named profile accepts explicitly trusted content and declares its unsupported credential, isolation, layout and input behavior. Retain effective HTML, CSS and JavaScript providers while advancing static CSS and rendering qualification; owned parser and interpreter replacement remain separate tracks. HtmlTinkerX can reuse the engine later, but its current helper APIs, package graph and workflows do not define the runtime's boundary or acceptance criteria.

A JavaScript interpreter can execute language code but does not supply a web platform. Evaluate the Jint dependency already present in HtmlTinkerX for the first bounded runtime, preserving a provider boundary. Language testing uses [Test262](https://github.com/tc39/test262); page behavior needs additional DOM and web API tests. Event scheduling follows the [HTML event-loop model](https://html.spec.whatwg.org/multipage/webappapis.html#event-loops).

Admit runtime features through executable profiles:

| Profile | Required behavior and evidence |
| --- | --- |
| Scripted local document | Selected language features, DOM mutation, events, timers, microtasks and deterministic capture; no unrestricted network or CLR exposure |
| `WebApplicationV1` | Modules and chosen fetch/storage/history/observer APIs, bounded document navigation, origin behavior, forms, lifecycle and layout invalidation; representative framework application fixtures |
| Interactive automation | Selected input dispatch, focus, scrolling, hit testing, locator semantics, actionability, explicit waits and cancellation; representative standalone OfficeIMO workflows pass on each advertised provider |
| Broader browser compatibility | Separately qualified shadow DOM/custom elements, canvas, workers, service workers, media, WebGL and other APIs as product scope expands |

Freeze time and disable unrelated animations in reproducibility tests; separately test real scheduling and animation behavior when those become supported. Define readiness using explicit lifecycle events, application predicates, resource/font completion and layout stability under a deadline. Network-idle alone cannot establish that a page is complete.

Any future untrusted-script profile must execute in an isolated worker with OS-enforced bounds. Interpreter constraints supplement process isolation; they are not the isolation boundary. Host objects are explicit capabilities and never expose arbitrary CLR access. Browser origins, cross-origin resource rules, credentials and navigation require dedicated security tests.

Removing a third-party JavaScript engine is a later language-runtime project covering parsing, evaluation, modules, promises, built-ins, memory management, internationalization and performance. Start with an interpreter only if the stage has its own funded scope and Test262 acceptance. A JIT is not required to begin; its absence also does not establish adequate performance for real applications. The static product must remain useful and independently releasable throughout.

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

## Dependency retirement

### Replaceable provider contracts

Replaceability means that the engine can use another implementation without changing its public document model or rewriting consumers. The algorithm behind that boundary can still be substantial work.

| Dependency | Boundary to establish now | What remains owned by OfficeIMO |
| --- | --- | --- |
| AngleSharp | Parse source/fragment into owned nodes and parse diagnostics | Document identity, mutation, serialization policy, query contract and consumer-facing types |
| CSS parser | Parse into owned syntax/declaration data; isolate any selector implementation used temporarily | Source retention, cascade decisions, computed-value contract and capability reporting |
| HarfBuzz | Existing shared shaping request/result contract | Fonts, glyph/cluster/logical-text mapping, lifetime, fallback policy and layout consumption |
| CodePages | Central encoding resolution/decoding boundary using owned metadata and suitable BCL types | Label policy, BOM/HTML precedence, errors, streaming behavior and limits |
| Script interpreter | Runtime provider bound to owned host/DOM contracts | Web API behavior, origins, scheduling policy, resource access and process isolation |

Do not expose provider nodes, native handles or provider exception types as normal public results. Translate errors and define cancellation, disposal and resource ownership at the boundary. Keep provider setup and configuration in composition/provider code; do not scatter encoding registrations or library-version checks through consumers. Existing .NET/BCL contracts do not need replacement wrappers merely to remove a NuGet dependency.

Use capability-specific contracts rather than mirroring every third-party method. Prove the first real implementation through contract fixtures; later run replacement providers against the same fixtures. A fake second parser does not establish replaceability. For CodePages, test actual legacy byte decoding and unavailable-encoding behavior; for shaping, test real glyph/cluster output and lifetime. Package separation follows real deployment and consumer needs, not one package per interface.

### Removal gates

Here, independence means no third-party runtime packages or browser binaries in the advertised default profile. The .NET runtime/BCL, OS services and clearly declared font/Unicode/data resources remain. Test tools and optional providers are separate. If a stronger goal excludes OS text, TLS, codecs or all external data, it requires a different platform scope.

| Dependency | Short-term role | Replacement/removal gate |
| --- | --- | --- |
| AngleSharp.Css | Current CSS parser while owned syntax is established | CSS grammar/recovery and cascade fixtures pass; existing protected-token/raw-recovery paths are removed |
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

Measure cold/warm time, per-stage elapsed time, allocation, peak/retained memory, output size, cancellation latency and scaling. Include parse-only, query-only, semantics-only, continuous layout and long paged documents. Calibrate hard budgets from the initial baseline on named machines; do not invent a throughput promise before measuring. Cache by document revision, CSS, resource/font identity, environment and provider version, with bounded retention.

Diagnostics should locate the earliest divergent stage: source/token, DOM, winning declaration, computed value, box, line, fragment or paint operation. Extend the existing report with stable node/source identity, requested feature, selected provider, outcome and affected target. No-warning output means only that no known problem was reported; it is not an independent fidelity certificate.

## Current implementation direction

The owned document boundary is the stable consumer surface. Continue widening real document and runtime workflows through it rather than exposing provider nodes or creating product-specific DOMs. Keep HtmlTinkerX and other hosts thin; migrate them only when the owning OfficeIMO component is usable and available through an explicit package or source relationship.

Strengthen style, layout, conversion and runtime components through OfficeIMO-owned contracts while the temporary providers remain behind them. Move CSS or parser replacement ahead when measured recovery or capability failures obstruct a required workflow. Advance managed typography in the shared graphics owner. Treat static rendering, the selected application profile and dependency retirement as separately qualified outcomes.

Deliver the next work in reviewable vertical slices:

1. Make the render request, named profiles, page-set behavior and shared result contract executable without changing layout algorithms merely to rename them.
2. Productize `OfficeIMO.Html.Core` for a small non-Office consumer, including packed-package, public API, lifecycle and compatibility proof while the temporary parser is identified in diagnostics.
3. Close H3-H4 CSS, layout, fragmentation and output gaps profile by profile against frozen browser, geometry, semantic and artifact references.
4. Complete H5 as an integrated managed parser and serializer, switch the default only after conformance, bounds and performance gates pass, and retain the adapter only where real migration demand exists.
5. Complete H6 by removing AngleSharp, AngleSharp.Css and other third-party runtime dependencies from the advertised static graph, with packed transitive-graph proof on every supported target.
6. Expand trusted runtime and automation profiles independently when real application workflows require them; broader browser compatibility and an owned language engine retain their own gates.

Preserve explicit bounds, provider identity and unsupported results throughout. Competitive claims attach to the completed profile or adoption stage, never to the repository as a whole.

The open milestones and their readiness gates are maintained once in the [roadmap](ROADMAP.md#independent-html-engine). A static-engine release, an AngleSharp-free package and an interactive-runtime release are separate reviewable outcomes. Calendar forecasts should follow the first slice and the measured remaining failure classes. Broad browser compatibility remains the largest and least predictable part of the program.
