# OfficeIMO.Latex

`OfficeIMO.Latex` is a dependency-free, source-preserving LaTeX2e interoperability engine. It implements the bounded `OfficeIMO` document profile; it is not a TeX compiler and never executes commands, loads packages, or invokes an external runtime.

The native parser retains every input character, exposes tokens, groups, commands, environments, math, comments, and common document semantics, and writes unchanged input exactly.

```csharp
using OfficeIMO.Latex;

LatexParseResult result = LatexDocument.Parse(source);
LatexDocument document = result.Document;

LatexHeading first = document.Headings[0];
first.Command.GetRequiredArgument(0)!.Content = "Updated section";

string updated = document.ToLatex();
```

The profile recognizes article/report/book structure, paragraphs, lists, figures, tabular data, labels/references, citations, theorem-like environments, and inline/display math. Unknown commands and environments remain source-backed instead of disappearing.

Simple document-local macros can be expanded only when explicitly enabled:

```csharp
LatexDocument document = LatexDocument.Parse(
    source,
    new LatexParseOptions {
        MacroExpansion = LatexMacroExpansion.SafeSimpleDefinitions
    }).Document;

LatexMacroExpansionResult expansion = document.ExpandSimpleMacros(@"\project{OfficeIMO}");
```

This is deliberately not general TeX expansion. Replacement control words must be another transitively safe document-local simple macro or an explicitly allow-listed formatting/reference command; file I/O, packages, shell escape, dynamic control sequences, category-code changes, bibliography tools, and TeX typesetting remain outside the product. The operation is bounded string substitution, not a sanitizer: invocation arguments and expanded output must still be treated as untrusted if another system will compile the TeX.

See the [LaTeX support matrix](https://github.com/EvotecIT/OfficeIMO/blob/master/Docs/officeimo.latex-support-matrix.md) for the exact boundary.

Targets: `netstandard2.0`, `net8.0`, `net10.0`, and `net472` on Windows.
