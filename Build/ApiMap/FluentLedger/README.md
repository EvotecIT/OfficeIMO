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
| `generate-configure` | A delegate-valued property is assigned through `Configure(x => x.Callback = ...)`, never by a same-named call. |
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
`getter-wraps` (a read-only option property whose getter builds a new object through a constructor that may copy it), `struct-property` (a setter on a value type, which an extension would apply to a copy),
`public-field`, and the open #3344 decisions configured in `fluent-ledger.json` (`pdf-run-factories`,
`markdown-shorthand`, `word-run-type`, `chart-entry-point`, `conversion-request`).

Facts: `void-setter` (the 4.0 extension returns the receiver where the method returned nothing), `no-property`
(no property of that name, so the extension must call a mutator), `multi-value` (a `Set...` with several values
and no matching property, renamed with its arguments), `keyed-set` (the first argument addresses what is written,
so it stays an action), `merges-set` (a property whose `Set...` method maps to the same extension: one generated
method), `is-prefix`, `property-rename`, `struct-receiver` (a value-type handle: the extension must call a handle
mutator), `interface-member`, `init-only`, `non-public-setter`, `set-method`, `collection`, `returns-parent`,
`returns-void`, `getter-copies` (a read-only option property whose getter returns a copy or a detached new object, so it gets no in-place
callback), `getter-view` (the getter returns a new wrapper over the receiver or a reference it holds, so the callback
edits the receiver), `copy-unverified` (a `With...` on a mutable type whose body was not proven to copy),
`matches-property` and `callback`.

## Generating the extensions

`OfficeIMO.Fluent.Generators` is the prototype that turns `rename` and `generate*` rows into code. A receiver opts
in with `[Fluent]`; a canonical write path that has no public setter (a `Set...` method becoming internal, or a
value-type handle such as `ExcelCell`) is marked `[FluentMutator("Bold")]`, and the generated `Bold(...)` calls it
and returns the receiver. Reference types get `TSelf` extensions so derived receivers keep their type; value-type
handles return the same handle without allocating. Collisions, setters on value types and two members mapping to
one method are build errors. `OfficeIMO.Fluent.Generators.Tests` proves it on a small receiver.

PowerForge.Generators already generates fluent methods for ConsoleMonster and the other Evotec libraries from the
same kind of attributes. It cannot yet call an instance mutator, generate for structs, or give `bool?` a
`X()` that means true, which the engines need. Either it gains those three and OfficeIMO consumes it as a
build-time dependency, or this generator stays package-local; #3344 tracks that decision.

## Migrating callers

The rename runs after the SDK removal, against a regenerated ledger. It is mechanical because every row is keyed
by `id` and carries its `proposed` call.

1. **Add before removing.** Mark the receivers `[Fluent]` and the canonical write paths `[FluentMutator]`. The
   generated `Bold()` coexists with the public `SetBold()` (different names), so every caller still compiles.
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
4. **Remove the old surface.** Make the `Set...` methods internal (they stay the `[FluentMutator]` write paths) or
   delete them where a property setter already is the write path, rebuild, and regenerate the ledger: no public
   `rename` rows may remain.
5. **Generate the migration notes.** Write the applied rows into `Build/ApiMap/csharp/{package}.json` as `rename`,
   `change` and `remove` entries; `New-MigrationGuide.ps1` builds the 4.0 section of `MIGRATION.md` from them.

`.SetBold(` is a useful progress indicator while this runs, not a measure of completeness: the ledger is.
