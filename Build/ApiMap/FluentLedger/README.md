# 4.0 fluent symbol ledger

The ledger lists every exported member of every authoring receiver, one row per overload, and gives each its 4.0
form under the [#3344](https://github.com/EvotecIT/OfficeIMO/issues/3344) language and migration maps. It is read
from the built assemblies, not grepped from source, so it can be regenerated once the Open XML SDK removal has
changed the engines. Its output is [`Build/ApiMap/fluent`](../fluent/README.md): one JSON file per package and a
generated summary. Edit the rules or `fluent-ledger.json`, never the output.

## Regenerate

Build the packages listed in `fluent-ledger.json` for every framework it names, then run the tool:

```sh
for p in $(python3 -c "import json;print(' '.join(json.load(open('Build/ApiMap/FluentLedger/fluent-ledger.json'))['packages']))"); do
  dotnet build $p/$p.csproj -c Release -f net10.0
  dotnet build $p/$p.csproj -c Release -f netstandard2.0
done
dotnet run --project Build/ApiMap/FluentLedger -c Release -- .
```

`--framework <tfm>` (repeatable) overrides the configured frameworks, `--config` and `--output` override the paths.
The configured set is `net10.0` and `netstandard2.0`; a row present in only some of them lists them in
`frameworks`. Add `net8.0`, or `net472` on Windows, to cover those builds as well.

The tool reuses the surface inventory's metadata loader and source matcher (`../SurfaceAnalysis`). Source is read
only to find each member's file and line and what its body returns.

## What is in scope

A public class, struct or interface is an authoring receiver when it has a public settable property, or a public
instance method that is `Set...`, `With...`, `Add...` or returns the receiver, or when a public extension method
targets it that way. Exceptions and attributes are left out. Every public member of a receiver gets a row:
methods, extension methods (listed under the type they extend, with `declaredIn` naming the static class),
constructors, properties, indexers, fields and events. A static class appears only when `fluent-ledger.json` has a
decision or retirement for it. A member inherited from a base type appears once, under the type that declares it.

## Row fields

| Field | Meaning |
| --- | --- |
| `id` | Documentation comment ID: the stable key for codemods (`DocumentationCommentId.GetFirstSymbolForDeclarationId`). |
| `receiver`, `declaredIn` | The configured type and, for an extension method, the static class declaring it. |
| `kind`, `member`, `signature`, `returns` | The current exported shape, including parameter names and defaults. |
| `treatment` | The 4.0 form (below). |
| `proposed` | The 4.0 call for `rename`, `generate*`, `retire` and decisions that already have a proposal. |
| `powerForge` | What PowerForge.Generators produces for the member once its receiver is `[Fluent]` (below). |
| `semantics` | What the body returns: `returns-receiver` (mutation), `returns-new` (copy), `mixed`, `delegates`, `void` or `no-source`. |
| `reason` | Why, when the treatment alone does not say. |
| `flags` | Extra facts and decision codes (below). |
| `frameworks` | Present only when the member is missing from some configured framework builds. |
| `source` | File and line of the declaration. |

## Treatments

| Treatment | 4.0 form |
| --- | --- |
| `rename` | A `Set...` configuration method becomes the property-named extension in `proposed` (`SetBold` to `Bold()` / `Bold(false)`). Also property renames from `fluent-ledger.json`. |
| `generate` | A public settable property gets a generated `X(value)` extension; `bool` and `bool?` get `X(bool value = true)`, `IsX` becomes `X()`. Assignment keeps working. |
| `generate-nested` | A read-only property holding an option object gets `X(o => ...)`, configured in place. |
| `generate-verb` | A delegate-valued property gets a method named by a verb (`[Fluent("FormatLabels")]`), never a same-named call, which would invoke it. |
| `keep-action` | An explicit editing action keeps its verb (`ApplyTheme`, `ReplaceAll`, `RemoveX`), including keyed `Set...` writes. |
| `keep-copy` | A genuine `With...` copy keeps its name and copy semantics. |
| `keep-add` | `Add...` keeps returning the created native child. |
| `keep-lifecycle` | Constructors, static factories, `Create`/`Load`/`Save` and disposal. |
| `keep-conversion` | `To...`, `SaveAs...`, `Convert...`, export and import adapters. |
| `keep-selector` | Getters, indexers and `Get`/`Find`/`Is...` reads. |
| `keep-property` | Init-only or non-public setters: initializer or read-only state. |
| `keep-fluent` | Already returns the receiver under a property-named or content-scope name (`Margins(...)`, `Paragraph(...)`). |
| `keep` | Object contracts, operators, overrides and other domain members. |
| `retire` | Removed in 4.0; `proposed` is the replacement (the special-purpose chart methods). |
| `decision` | Needs a decision before the rename; the code is in `flags`. |

## Flags

Decision codes: `name-collision` (the 4.0 name already names a method, static factory, nested type or event on
the receiver or a derived receiver), `returns-other` (a `Set...` that returns something other than the receiver
or `void`), `with-mutates` (a `With...` that returns the same mutated receiver), `use-prefix` (`Use...`
configuration, which #3344 has no family for), `bare-set` (`Set(...)` with nothing to name), `static-set`,
`set-returns-copy`, `generic-set` (a generic `Set...<T>`, whose type parameters a generated mutator call does not carry),
`getter-wraps` (a read-only option property whose getter builds a new object through a constructor that may copy
it, or from state whose type is not known to be a reference), `struct-property` (a setter on a value type, which an extension would apply to a copy),
`public-field`, and the open #3344 decisions configured in `fluent-ledger.json` (`pdf-run-factories`,
`markdown-shorthand`, `word-run-type`, `chart-entry-point`, `conversion-request`).

Facts: `void-setter` (the 4.0 extension returns the receiver where the method returned nothing), `no-property`
(no property of that name, so the extension must call a mutator), `multi-value` (a `Set...` with several values
and no matching property, renamed with its arguments), `keyed-set` (the first argument addresses what is written,
so it stays an action), `merges-set` (a property whose `Set...` method maps to the same extension: one generated
method), `is-prefix`, `property-rename`, `struct-receiver` (a value-type handle: the extension must call a handle
mutator), `interface-member`, `init-only`, `non-public-setter`, `set-method`, `collection`, `returns-parent`,
`returns-void`, `getter-computed` (a collection getter that builds the list each time, so adding to it is lost),
`collection-collision` (a collection's one-item or range method name is already a method), `getter-copies` (a read-only option property whose getter returns a copy or a detached new object, so it gets no in-place
callback), `getter-view` (the getter returns a new wrapper over the receiver or a reference it holds, so the callback
edits the receiver), `copy-unverified` (a `With...` on a mutable type whose body was not proven to copy),
`matches-property` and `callback`.

## PowerForge.Generators

The extensions come from [PowerForge.Generators](https://github.com/EvotecIT/PSPublishModule/tree/main/PowerForge.Generators),
the build-time generator ConsoleMonster and the other Evotec libraries already use, referenced with
`PrivateAssets="all"` so it adds no runtime dependency. A receiver opts in with `[PowerForge.Fluent.Fluent]`; every
public settable property then gets `X(value)` (`X(bool value = true)` for a `bool`, `IsX` becoming `X`), a settings
object gets `X(o => ...)` configured in place, and an `ICollection<T>` gets one-item and range methods named after
the singular. `[Fluent("Name")]` renames a method, `[Child("Name")]` renames a collection's one-item methods and
`[NoFluent]` leaves a member out; a collision is build error PFF002.

`powerForge` records the outcome per row:

| Value | Meaning |
| --- | --- |
| `generated` | A property setter, or a `Set...` method whose property has a public setter (the method then retires). |
| `generated-configure` | `X(o => ...)` on a read-only settings property, including a live wrapper (`getter-view`). |
| `generated-collection` | One-item and range methods on a stored collection. |
| `[Fluent("Verb")]` | A delegate property: PFF001 until it names a verb. |
| `[Fluent("Name")]` | The generated name collides (`name-collision`): PFF002 until it is renamed or left out. Also a configured property rename (`property-rename`), whose 4.0 name differs from the one the generator derives. |
| `[Child("Name")]` | A collection whose one-item or range name is taken (`collection-collision`). |
| `[NoFluent]` | Must be left out: a getter that copies (`getter-copies`), or a collection that is computed (`getter-computed`) or read-only, where generated methods would edit a throwaway object or throw. |
| `gap:...` | Something OfficeIMO needs that the generator does not support yet (below). |

### What PowerForge.Generators does not support yet

These are generator changes, made in PSPublishModule as their own release. The summary in
[`Build/ApiMap/fluent`](../fluent/README.md#powerforgegenerators) counts the rows for each.

- **`gap:mutator`: a canonical write path without a public setter.** The generator only assigns public setters.
  `WordComment.SetResolved(bool)` writes `Resolved`, whose setter is not public; `WordChart.SetRadialLayout(...)` has
  no property at all. `[FluentForward]` reaches a static helper with one value, so each would need a handwritten
  static wrapper. Needed: an attribute on an internal instance method (`[FluentMutator("Resolved")]`) that the
  generated `Resolved(bool value = true)` calls.
- **`gap:multi-value`: a configuration call with several values.** `PdfHeaderFooter.SetZones(left, center, right)`,
  `PowerPointChart.SetAccessibility(alternativeText, dataSummary, includeDataSummary)`. Forwarding takes one value.
  Needed: the mutator above with its full parameter list and defaults.
- **`gap:value-handle`: value-type handles.** `ExcelCell` is a struct over the sheet, so `ExcelCell.SetBold(...)` and
  the other cell setters cannot become `Bold()`: the generator only writes for classes. Needed: by-value
  extensions on a struct (`this ExcelCell cell`) that call a handle mutator and return the same handle, without
  allocating a wrapper.
- **`gap:nullable-bool`: `bool?` properties.** `PdfViewerPreferencesOptions.CenterWindow` and
  `WordParagraph.AvoidWidowAndOrphanOverride` get `X(bool? value)` with no default, so `X()` does not compile, and
  `IsX` keeps its prefix. Needed: the `bool` rule (`X(bool? value = true)`, `IsX` to `X`) for `bool?`.
- **`gap:receiver-shape`: nested option types.** `PowerPointExtractionExtensions.PowerPointExtractOptions`,
  `TocBlock.Entry` and `ExcelDocument.ExcelTabularRowWriter` are nested, which the generator skips. Either they
  move to the namespace or the generator learns nested types.
- **`gap:settings-shape`: abstract settings types.** `PowerPointConnectionShape.StartShape` returns an abstract
  `PowerPointShape`, which is not configured in place because the generator requires a concrete type. Only two rows;
  configuring in place does not need to create the object.

Not gaps, but worth folding into the generator: it configures a read-only settings property in place without looking
at the getter and generates adders for any `ICollection<T>`, so the `[NoFluent]` rows must be marked by hand
(`WordDocument.Charts` returns a new list on every read, `HtmlRenderRequest.Options` returns a clone).

## Migrating callers

The rename runs after the SDK removal, against a regenerated ledger. It is mechanical because every row is keyed
by `id` and carries its `proposed` call.

1. **Add before removing.** Reference PowerForge.Generators, mark the receivers `[Fluent]`, add the attributes the
   `powerForge` field names, and close the `gap:` rows in the generator first. The generated `Bold()` coexists with
   the public `SetBold()` (different names), so every caller still compiles.
2. **Rewrite C# callers with a Roslyn codemod.** It opens each project with `MSBuildWorkspace` (Microsoft.Build.Locator
   and Microsoft.CodeAnalysis.Workspaces.MSBuild, build-time only), resolves every `rename` and `retire` row from its
   `id` in the engine compilation, and finds references with `SymbolFinder`. A `rename` reference becomes the
   `proposed` name with the same arguments; a named argument is renamed when the proposed parameter name differs
   (`isBold:` to `value:`), and a renamed property (`Strike` to `Strikethrough`) is rewritten where it is accessed.
   A `retire` or unresolved `decision` reference is not rewritten: the codemod lists it with the replacement and
   fails until none remain. Projects, in this order:
   - engines and converters that call each other (`OfficeIMO.*`), so internal callers move first;
   - tests (`*.Tests`, `*.Verification`, `*.AotSmoke`, benchmarks);
   - `OfficeIMO.Examples`, Studio, `OfficeIMO.Tool`, `OfficeIMO.Workflows` and the other hosts;
   - PSWriteOffice's C# cmdlets and services (`PowerShell/PSWriteOffice/Sources`). Cmdlet names, parameters and
     output stay as they are.
   - `Build/DocumentationExamples/*` (they are projects), then `powerforge docs sync --config Build/docs-sync.json --apply`
     regenerates the README and guide blocks from them.
3. **Rewrite code that is not a project by syntax.** Fenced `csharp` blocks in `Docs/` and `Website/content/docs`
   and PowerShell scripts that call .NET members (`$paragraph.SetBold()`, found with the PowerShell AST) have no
   semantic model. Their calls are renamed only when the member name maps to one 4.0 name for every receiver in
   the ledger; any other call is reported. `Build/Test-DocsCodeExamples.ps1` then compiles the website blocks and the
   Pester tests cover the scripts.
4. **Remove the old surface.** Make the `Set...` methods internal (they stay the canonical write paths) or
   delete them where a property setter already is the write path, rebuild, and regenerate the ledger: no public
   `rename` rows may remain.
5. **Generate the migration notes.** Write the applied rows into `Build/ApiMap/csharp/{package}.json` as `rename`,
   `change` and `remove` entries; `New-MigrationGuide.ps1` builds the 4.0 section of `MIGRATION.md` from them.

`.SetBold(` is a useful progress indicator while this runs, not a measure of completeness: the ledger is.
