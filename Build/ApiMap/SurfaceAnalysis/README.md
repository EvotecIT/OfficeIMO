# Public surface inventory

This tool proposes accessibility candidates for the Tier 1 libraries. It reads
built Release assemblies, user guides, package READMEs and C# examples. One model
signature closure spans the selected packages, so a documented Word API can retain
its Core options and result models. Public interfaces, overrides, operators and
exposed model data retain their contracts.

Build the selected packages for `net10.0` before analysis, then run:

```sh
dotnet run --project Build/ApiMap/SurfaceAnalysis -- \
    . /path/to/surface-inventory.json
```

The default set is Core, Word, Excel, PowerPoint, Pdf, Markdown and OfficeIMO.
Additional arguments replace that package set. The tool uses the repository SDK's
Roslyn assemblies and adds no runtime package dependency.

The output is a review list, not an automatic accessibility edit. Check each
candidate's callers, public inheritance and signatures in adapters before changing
it. Constructor evidence is separate from type references, so an editable image
returned by `AddImage` does not make its package-internal construction public.
Compiler-generated constructors can have no explicit source location. Overload
locations are omitted when the source signature cannot be matched unambiguously.
Generated API references and design documents are excluded as evidence; otherwise
an automatically listed API would keep itself public.

After reviewing the surface, use the public API analyzer's shipped baseline to
track deliberate changes. [The analyzer documentation](https://github.com/dotnet/roslyn/blob/main/src/RoslynAnalyzers/PublicApiAnalyzers/PublicApiAnalyzers.Help.md)
explains nullability and separate baseline files for conditional framework APIs.
