# OfficeIMO Markdown Duplicated Node Shapes Inventory (2026-03-21)

This note is the B1 inventory for duplicated mutable node shapes in `OfficeIMO.Markdown`.

It is not a complaint about convenience APIs. Multiple views are fine when one view is clearly primary and the others are derived, read-only, and cheap to recompute.

The risk starts when a node keeps multiple structural representations that can drift, or when helper code has to guess which representation is authoritative.

## Classification rules

Low risk:

- one primary structural representation
- secondary views are read-only or obviously derived
- no meaningful ambiguity about ownership

Medium risk:

- multiple structural views exist
- drift is possible but constrained
- the cleanup pattern looks localized

High risk:

- multiple mutable representations overlap
- syntax/rendering/helpers choose between them dynamically
- source-span or AST ownership can become ambiguous

## Inventory

### 1. `ListItem`

File:

- `OfficeIMO.Markdown/Blocks/ListItem.cs`

Current shape:

- primary-ish inline content: `Content`
- additional paragraph inlines: `AdditionalParagraphs`
- synthesized paragraph blocks: `_leadParagraphBlock`, `_additionalParagraphBlocks`, `ParagraphBlocks`
- nested block children: `Children`
- mixed block projection: `BlockChildren`
- cached syntax children: `SyntaxChildren`

Why it is duplicated:

- the same logical list-item body exists as both inline sequences and paragraph blocks
- nested content exists both as direct children and as the mixed `BlockChildren` projection
- syntax generation special-cases paragraph blocks but still routes ownership through the inline sequence path

Risk:

- High

Why:

- several APIs expose overlapping structure
- some of those views are mutable (`AdditionalParagraphs`, `Children`)
- syntax ownership already showed correctness issues here

Suggested canonical direction:

- make paragraph blocks the clear primary block-level representation for list-item paragraph content
- keep inline access as a convenience view, not as a competing source of truth

### 2. `DefinitionListBlock`

File:

- `OfficeIMO.Markdown/Blocks/DefinitionListBlock.cs`

Current shape:

- semantic grouped model: `Groups`
- typed entry model: `Entries`
- legacy mutable tuple view: `Items`
- derived inline view: `InlineItems`
- syntax cache: `SyntaxItems`
- parser context for recovery: `ReaderOptions`, `ReaderState`

Why it is duplicated:

- there are at least three distinct representations of the same definition-list content:
  - grouped semantic wrappers
  - flat typed entries
  - legacy tuple items
- syntax and inline recovery paths can derive additional transient structures from those

Risk:

- High

Why:

- the grouped and entry-based models overlap semantically
- legacy compatibility adds another mutable surface
- syntax ownership and fallback recovery are already ambiguous here

Suggested canonical direction:

- choose one semantic representation as primary
- derive the others explicitly
- treat legacy tuple access as an adapter layer, not a peer data model

### 3. `TableBlock`

File:

- `OfficeIMO.Markdown/Blocks/TableBlock.cs`

Current shape:

- raw markdown cells: `Headers`, `Rows`
- parsed inline cells: `HeaderInlines`, `RowInlines`, `ParsedHeaders`, `ParsedRows`
- structured cells: `HeaderCells`, `RowCells`, `StructuredHeaders`, `StructuredRows`
- several signatures and realized-cell caches

Why it is duplicated:

- the same table content can exist as raw strings, parsed inlines, structured cells, and realized cached cells
- rendering and syntax paths can switch between these views depending on signatures and parse context

Risk:

- High

Why:

- this is the broadest representation matrix in the package
- some duplication is justified for performance and compatibility
- but the ownership model is difficult to reason about without reading multiple helper paths

Suggested canonical direction:

- keep one canonical semantic cell model
- make raw-string and parsed-inline views adapters or caches around it
- reserve signatures/caches for performance, not semantics

### 4. `FootnoteDefinitionBlock`

File:

- `OfficeIMO.Markdown/Blocks/FootnoteDefinitionBlock.cs`

Current shape:

- fallback text: `_fallbackText`
- block view: `Blocks`
- inline paragraph view: `Paragraphs`
- paragraph block view: `ParagraphBlocks`
- syntax cache/input: `SyntaxChildren`

Why it is duplicated:

- the footnote body can be carried as text, block children, paragraph inlines, and paragraph blocks
- several constructors normalize different inputs into internal content views

Risk:

- Medium

Why:

- there is real duplication, but the internal `FootnoteContentViews` wrapper already moves the type toward a central normalization point
- the cleanup pattern is local enough to be a good refactor template

Suggested canonical direction:

- make `Blocks` the primary structural representation
- derive paragraph-only helpers from `Blocks`
- keep raw text only as an input fallback, not as a peer semantic shape

Recommended B2 candidate:

- Yes

Reason:

- medium complexity
- already partially normalized
- useful proving ground for “single primary representation, derived convenience views”

### 5. `CalloutBlock`

File:

- `OfficeIMO.Markdown/Blocks/CalloutBlock.cs`

Current shape:

- title text convenience view: `Title`
- title inline structure: `TitleInlines`
- fallback raw body: `_fallbackBody`
- parsed body blocks: `ChildBlocks`
- syntax cache/input: `SyntaxChildren`

Why it is duplicated:

- both title and body have dual representations:
  - plain text / raw fallback
  - structured parsed content

Risk:

- Medium

Why:

- the duplication is understandable, but syntax ownership for the title is already incomplete
- still narrower and easier to reason about than `DefinitionListBlock` or `TableBlock`

Suggested canonical direction:

- keep `TitleInlines` and `ChildBlocks` as the semantic shape
- treat `Title` and raw fallback body as derived or builder-only inputs

### 6. `DetailsBlock`

File:

- `OfficeIMO.Markdown/Blocks/DetailsBlock.cs`

Current shape:

- structured summary block: `Summary`
- mutable child block list: `Children`
- read-only child projection: `ChildBlocks`
- syntax cache/input: `SyntaxChildren`

Why it is duplicated:

- mostly because of the usual mutable list plus read-only view plus syntax cache pattern

Risk:

- Low

Why:

- the primary structure is fairly clear
- `ChildBlocks` is just the public read-only projection of `Children`
- this does not currently look like a major semantic duplication hotspot

Suggested canonical direction:

- leave mostly as-is unless a broader container policy standardizes syntax-cache handling

## Recommended order

1. `FootnoteDefinitionBlock`
   - best B2 candidate
   - medium complexity with a visible cleanup pattern
2. `ListItem`
   - high-value correctness cleanup after the pattern is proven
3. `DefinitionListBlock`
   - larger semantic cleanup once the pattern is established
4. `TableBlock`
   - broadest and riskiest; do after ownership rules are explicit

## Working rule going forward

For each node type, we should be able to answer three questions without hesitation:

1. What is the primary semantic structure?
2. Which views are derived from it?
3. Which object should own the syntax node for each emitted AST wrapper?

If those answers are not obvious, the node shape is still too ambiguous.
