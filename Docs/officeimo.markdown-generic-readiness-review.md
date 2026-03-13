# OfficeIMO.Markdown Generic Readiness Review

This document starts the cleanup track for turning `OfficeIMO.Markdown` into a generic-first markdown package that can still power `IntelligenceX.Chat`, future `OfficeIMO.Markdown.Html`, and other hosts.

## Current Read

The good news:

- `OfficeIMO.Markdown` already has the foundations of a real parser instead of a pure string transformer.
- The reader has a typed document model, a syntax tree with source spans, query helpers, and a configurable block parser pipeline.
- `IntelligenceX.Chat` already tries to separate transcript normalization, runtime probing, and export behavior in its own contracts/docs.

The main blockers:

- the renderer extension story is still HTML-string postprocessing instead of AST-driven rendering
- IX-specific fenced block concepts currently live in the generic renderer package
- chat/transcript repair logic is spread across `IntelligenceX`, `OfficeIMO.Markdown`, and `OfficeIMO.MarkdownRenderer`
- the public AST is present, but it is not yet the primary extension surface for transforms/renderers

## What Needs To Change

### 1. Make the AST the center of the package

Target direction:

- parser produces a stable public AST with source spans
- HTML rendering walks AST nodes instead of re-parsing emitted HTML with regex
- future transforms such as HTML-to-markdown can target the same AST instead of host-specific HTML conventions

Why this matters:

- today diagrams/charts/networks/dataviews are mostly discovered after HTML generation
- that means the markdown model cannot express those semantics cleanly
- it also makes round-trip and alternate renderer work harder than it should be

### 2. Move product-specific visuals out of the generic renderer core

`OfficeIMO.MarkdownRenderer` should stay generic and host-oriented.

These should become opt-in extensions instead of built-ins in the generic package:

- `ix-dataview`
- `ix-chart`
- `ix-network`

Good end state:

- generic renderer knows how to host fenced-block extensions
- IX ships its own renderer extension package or registration helper
- aliases and richer behavior for IX stay available without leaking IX contracts into the base package

### 3. Separate markdown semantics from transcript repair

There are three different concerns and they should stay separate:

1. generic markdown parsing/rendering
2. transcript/chat repair for malformed model output
3. host/runtime shell behavior

Desired ownership:

- `OfficeIMO.Markdown`: parser, AST, generic renderer, neutral normalization primitives
- `OfficeIMO.MarkdownRenderer`: shell/document host behavior, incremental update script, generic fenced-block extension hooks
- `IntelligenceX.Chat`: transcript repair policy, visual preference policy, runtime enablement

### 4. Add a neutral preset story

Right now the renderer preset surface is chat-first. We also need generic presets such as:

- strict generic
- portable generic
- docs generic
- host shell minimal

Chat presets can stay, but they should sit on top of neutral presets instead of defining the whole public mental model.

## Immediate Findings

### Finding A: renderer extensions are post-HTML regex conversions

Current shape:

- markdown parses into `MarkdownDoc`
- HTML is generated
- fenced block replacements then scan HTML strings for `<pre><code class="language-*">...`

This is the largest architectural issue because it keeps visual extensions outside the AST.

### Finding B: generic renderer is IX-aware today

Current built-ins include IX-specific languages and payload conventions.

That is useful short term, but it makes the package harder to position as a generic alternative and harder to reuse for non-IX hosts.

### Finding C: normalization is duplicated across layers

Some repair behavior is available:

- in `IntelligenceX.Chat` transcript preparation
- in `OfficeIMO.Markdown` input normalization
- in `OfficeIMO.MarkdownRenderer` chat presets and preprocessing

This creates drift risk and makes it unclear which layer owns a given repair.

### Finding D: the syntax tree exists, but renderer/plugin APIs do not really flow through it yet

The syntax tree is useful for diagnostics and span lookup, but not yet the main composition surface for transforms, renderers, or custom fenced block semantics.

## Proposed Phases

### Phase 0: lock behavior with corpus tests

Before deeper refactors:

- add a markdown corpus for generic docs cases
- add a chat corpus for malformed transcript cases
- add a visual fence corpus for mermaid/chart/network/dataview behavior
- keep paired tests in `OfficeIMO` and `IntelligenceX` for contract-sensitive cases

### Phase 1: promote fenced code blocks to a first-class extension seam

Add a real fenced-block extension pipeline that operates before HTML generation:

- parse fenced blocks into typed AST/block nodes
- allow extension handlers to claim languages/aliases
- let HTML renderers render those nodes without regex backtracking over HTML

This phase gives the biggest payoff while keeping current behavior reachable.

### Phase 2: de-IX the generic package

After the extension seam exists:

- keep Mermaid generic in base renderer
- move `ix-dataview`, `ix-chart`, and `ix-network` registrations into an IX-specific adapter/extension package
- keep backward compatibility through an IX registration helper in the chat app during transition

### Phase 3: align normalization ownership

- keep `MarkdownInputNormalizationPresets` as generic opt-in presets
- stop duplicating equivalent repair passes in the renderer when the reader already owns them
- keep IX transcript repair limited to transcript-specific policy and persisted-history cleanup

### Phase 4: build the HTML-to-markdown bridge on top of the AST

For `OfficeIMO.Markdown.Html`:

- map HTML into the same markdown AST/block model
- reuse markdown rendering/serialization from the core package
- avoid creating a second independent markdown representation

## First Practical Backlog

1. Add a neutral renderer preset family that is not chat-branded.
2. Remove double-applied normalization paths so one layer clearly owns each repair.
3. Introduce an AST-level fenced block extension contract.
4. Move IX-specific built-in fenced block registrations behind an optional adapter.
5. Add round-trip tests for markdown -> AST -> HTML and HTML -> AST -> markdown.
6. Publish a compatibility matrix that clearly marks generic features vs host-specific extensions.

## Recommendation

Do not try to replace Markdig feature-for-feature in one jump.

The better path is:

1. make the AST and extension model correct
2. make the generic package neutral again
3. keep IX working through adapters and compatibility tests
4. then expand standards coverage and import/export breadth

That order gives us a package that is easier to reason about, easier to extend, and much safer to use as the base for `OfficeIMO.Markdown.Html`.
