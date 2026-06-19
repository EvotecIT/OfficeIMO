# OfficeIMO.Word Market Readiness Roadmap

Date: 2026-06-19
Branch/worktree: `codex/word-market-readiness-no-pdf` at `C:\Support\GitHub\_worktrees\OfficeIMO-word-market-readiness-no-pdf`
Scope: `OfficeIMO.Word`, `OfficeIMO.Word.Html`, `OfficeIMO.Word.Markdown`, docs, examples, and future thin PSWriteOffice exposure.

## Purpose

This roadmap tracks the non-PDF work needed to make `OfficeIMO.Word` feel like the best practical open-source Word automation choice for .NET and PowerShell.

Word-to-PDF fidelity is intentionally out of scope for this roadmap because it has a separate active engine effort. The work here focuses on the durable Word value proposition:

- document assembly and template automation
- review, redline, and diff workflows
- HTML and Markdown conversion fidelity with proof artifacts
- public docs and showcases based on real generated documents

The product position should stay honest: OfficeIMO.Word is not a commercial-suite clone and not a raw Open XML helper. The winning position is a source-available, MIT, COM-free, service-safe document automation engine with approachable APIs, actionable diagnostics, and strong proof for real workflows.

## Market Assessment Summary

| Area | What we have | What we do not have yet | Best-market move |
|------|--------------|-------------------------|------------------|
| Document assembly | Merge fields, repeated rows, grouped rows, repeated blocks, content-control conditionals, inspection, validation, and batch output. | A public workflow guide and proof gallery that starts from real templates and data. | Lead with status-report, contract, invoice, and evidence-pack scenarios backed by source inputs and generated documents. |
| Review and redline | Comments, revisions, visible markup helpers, comparison primitives, and initial structured findings through `CompareStructure(...)`. | Rich run-level diffing, review reports, redline generation from structured findings, and complete review metadata readback. | Make structured review automation the differentiator for contracts, policies, proposals, and audit documents. |
| HTML and Markdown | Dedicated converter packages, bidirectional conversion, diagnostics, support matrix, resource policy, and visual fallback modes. | Broader real-world fixture coverage and generated public artifacts. | Publish conversion proof with source HTML/Markdown, generated DOCX, round-trip output, diagnostics, and validation status. |
| Public story | Broad feature coverage, examples, tests, website docs, and package pages. | A fast evaluator path that connects claims to real artifacts and known limitations. | Replace generic feature-list marketing with workflow proof and honest support boundaries. |

## Ownership Boundaries

- `OfficeIMO.Word` owns the Word document model, template/mail-merge behavior, comments, revisions, compare/diff primitives, content controls, and feature inspection.
- `OfficeIMO.Word.Html` owns Word/HTML conversion, HTML import/export diagnostics, resource policy, CSS mapping, and HTML fixture galleries.
- `OfficeIMO.Word.Markdown` owns Word/Markdown conversion, semantic Markdown fallback, sidecar visual fallback, and Markdown round-trip behavior.
- `PSWriteOffice` should expose stable workflows only after reusable behavior exists in OfficeIMO. It should stay a thin PowerShell-facing layer.
- Website and docs should describe validated workflows and link to proof artifacts; they should not promise unsupported Word or PDF fidelity.

## Current State

### Document Assembly

Existing engine pieces are already substantial:

- merge-field replacement with simple and complex field support
- repeated table-row regions
- grouped table-row regions with group/detail templates
- repeated block regions for paragraph and table content
- nested repeated block regions
- content-control conditional template blocks in body, headers, footers, table cells, and blocks
- template inspection and validation diagnostics
- batch output from template files
- Custom XML-bound content-control fill/update workflows
- content-control form-map fill/extraction and preflight validation

The gap is productization. Users should not need to discover these pieces by reading tests or internals.

### Review And Redline

Existing engine pieces include:

- `WordDocumentComparer`
- structured comparison results through `WordDocumentComparer.CompareStructure(...)`
- comment authoring and traversal
- threaded comment examples
- revision settings and inserted/deleted run helpers
- accept/reject APIs
- visible-markup conversion paths
- tests for compare, comments, and revisions

The gap is a cohesive review workflow: structured diffs, redline documents, comment/revision metadata readback, and review reports.

### HTML And Markdown Conversion

Existing converter work is strong:

- bidirectional Word/HTML conversion
- bidirectional Word/Markdown conversion
- HTML diagnostics for skipped/degraded content and unsupported CSS
- resource limits and URI/content-type policy for HTML import
- support matrix and HTML roadmap
- Markdown visual fallback modes and semantic unsupported-content handling

The gap is public proof and corpus breadth: users should be able to inspect representative source HTML/Markdown, generated DOCX files, round-trip output, diagnostics, and validation status.

### Docs And Showcase

The current Word docs explain classes and basic creation. The product page lists many features. The missing piece is a workflow-first story built around:

- assembling documents from templates
- reviewing and redlining documents
- converting to and from HTML/Markdown
- proving output against real scenarios

## Priority 1: Template And Document Assembly Polish

Goal: make OfficeIMO.Word the obvious open-source choice for report, contract, policy, invoice, and evidence-pack assembly.

### Work Items

1. Add a workflow-level template guide that starts from a real `.docx` template and data model.
2. Document the marker vocabulary for merge fields, repeated rows, grouped rows, repeated blocks, conditionals, and content-control bindings.
3. Add a small scenario gallery:
   - status report
   - contract clause pack
   - invoice or order confirmation
   - audit evidence pack
4. Add a template preflight API story to the docs:
   - missing bindings
   - unused data
   - unsupported/nested region shape
   - ambiguous content-control tag/alias
   - formatting-preservation warnings
5. Fill the remaining engine gaps only when driven by scenario proof:
   - richer section regions
   - broader nested region shapes
   - deeper SDT binding scenarios
   - more formatting-preservation cases

### Acceptance Criteria

- A user can understand the template model from docs without reading tests.
- Every documented marker or content-control binding shape has a focused test.
- Gallery outputs include source template, data input, generated DOCX, diagnostics, and validation status.
- PSWriteOffice exposes only thin commands over the stable OfficeIMO engine behavior.

## Priority 2: Review, Redline, And Diff Workflows

Goal: make OfficeIMO.Word useful not only for creating documents, but for understanding and governing changes in documents.

This means Word Review-tab style automation:

- compare two `.docx` files
- produce structured differences
- generate a redline document
- preserve/read comments and review metadata
- report risky edits for contracts, policies, proposals, and audit documents
- accept or reject changes programmatically where the structure is supported

### Work Items

1. Define a stable structured diff model:
   - document metadata changes
   - paragraph insert/delete/update (initial path, text, table, and image findings exist through `CompareStructure(...)`)
   - run-level text/style changes
   - table row/cell changes
   - image and relationship changes
   - field/content-control changes
2. Add redline output helpers that turn the diff model into a reviewable `.docx`.
3. Add a review report output that can be consumed as JSON, Markdown, or Word.
4. Improve comments and revisions:
   - comment replies
   - resolved/reopened state where available
   - author/date metadata
   - revision readback
   - safer accept/reject behavior
5. Add scenario tests:
   - contract clause edits
   - policy text updates
   - table evidence changes
   - image replacement
   - comment resolution

### Acceptance Criteria

- `WordDocumentComparer` can produce a structured result, not only a document-level comparison artifact.
- Redline output is deterministic enough for tests and useful enough for human review.
- Unsupported diff areas surface diagnostics instead of silently disappearing.
- Docs clearly separate editable review metadata from preserve-only metadata.

## Priority 3: HTML And Markdown Fidelity With Proof Gallery

Goal: make conversion claims visible and reviewable through artifacts, not just prose.

### Work Items

1. Build a Word conversion proof gallery covering:
   - browser copy/paste HTML
   - Word-exported HTML
   - Outlook/email-like HTML
   - CMS article HTML
   - code-heavy documentation
   - Markdown tables/lists/images/front matter
   - document comments and notes where supported
2. For each scenario, emit:
   - source HTML or Markdown
   - generated DOCX
   - round-trip HTML or Markdown where meaningful
   - diagnostics
   - Open XML validation status
3. Keep the support matrix current as the gallery grows.
4. Expand tests only around product contracts:
   - diagnostics for skipped/degraded content
   - table/list/image round trips
   - metadata and accessibility preservation
   - unsupported-content placeholders
5. Document conversion profiles:
   - strict archival conversion
   - trusted document conversion
   - untrusted HTML ingestion
   - forgiving editor/import conversion

### Acceptance Criteria

- The gallery can be generated locally and in CI.
- Every gallery scenario has machine-readable diagnostics.
- The public docs link from claims to artifacts or support-matrix rows.
- Unsupported conversion behavior is deliberate, documented, and test-covered.

### Initial Implementation Slice

`MarketReadinessProofGallery.Example_GenerateWordMarketReadinessProof(...)` now emits a local proof gallery with template assembly, review/diff, HTML conversion, and Markdown conversion scenarios. The gallery writes source inputs, generated DOCX files, diagnostics, round-trip converter output where meaningful, and Open XML validation summaries. Run it with `dotnet run --project OfficeIMO.Examples -- --word-market-readiness`.

## Priority 4: Real-Document Docs And Showcase

Goal: replace feature-list marketing with workflow proof.

### Work Items

1. Update Word docs around workflows:
   - create from code
   - assemble from templates
   - review/redline/diff
   - convert through HTML/Markdown
   - inspect unknown documents before editing
2. Add showcase entries based on real generated documents:
   - report
   - contract/policy
   - invoice/order
   - review/redline
   - web-content import
3. For each showcase, include:
   - short scenario
   - source inputs
   - generated output
   - validation/proof status
   - known limitations
4. Keep the product page honest:
   - lead with COM-free Word automation
   - emphasize source-visible diagnostics and workflow ownership
   - keep PDF out of this story until the separate PDF effort is ready

### Acceptance Criteria

- A first-time evaluator can see the four main Word workflows in under five minutes.
- Docs point to real APIs and generated artifacts.
- The roadmap, support matrix, and website do not contradict each other.
- The public story is stronger because it is specific, not broader.

## Suggested First Implementation Slices

1. Template guide plus one generated status-report gallery scenario.
2. Structured diff model sketch plus paragraph/table/comment diff tests.
3. HTML/Markdown proof gallery script with two HTML and two Markdown scenarios. Initial example implementation exists in `OfficeIMO.Examples/Word/MarketReadiness/MarketReadinessProofGallery.cs`.
4. Website Word docs refresh linking the four workflows and current support matrix.

Each slice should end with source, test, generated artifact, and docs proof. If a feature is only partially supported, document it as partial and return diagnostics from the engine.

## Validation Gates

Use focused validation for each slice:

- `dotnet build OfficeIMO.Word/OfficeIMO.Word.csproj -c Debug`
- `dotnet test OfficeIMO.Tests/OfficeIMO.Tests.csproj -c Debug --filter "FullyQualifiedName~Word.MailMerge|FullyQualifiedName~Word.ContentControl"`
- `dotnet test OfficeIMO.Tests/OfficeIMO.Tests.csproj -c Debug --filter "FullyQualifiedName~Word.CompareDocuments|FullyQualifiedName~Word.Comments|FullyQualifiedName~Word.Revisions"`
- `dotnet test OfficeIMO.Tests/OfficeIMO.Tests.csproj -c Debug --filter "FullyQualifiedName~Html|FullyQualifiedName~Markdown"`

Broader release candidates should also run Open XML validation over generated gallery documents and package validation for any public surface additions.
