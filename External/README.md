# HTML runtime providers

The interactive worker builds the AngleSharp and AngleSharp.Js forks from the exact
Git revisions recorded by this repository. The submodules own the dependency source;
OfficeIMO owns rendering, resource policy, session budgets, module resolution,
automation, and the adapters that connect those capabilities.

Initialize the sources before building the worker or the full solution:

```sh
git submodule update --init --recursive
dotnet build OfficeIMO.Html.Runtime.Worker/OfficeIMO.Html.Runtime.Worker.csproj -f net10.0
dotnet test OfficeIMO.Html.Runtime.Tests/OfficeIMO.Html.Runtime.Tests.csproj -f net10.0
```

Normal `git submodule update` uses the committed revisions. Do not use `--remote`
for reproducible builds. A clone intended only for the static HTML packages does
not need the runtime providers.

| Provider | Source and patch notes | OfficeIMO consumer |
| --- | --- | --- |
| AngleSharp DOM | [EvotecIT/AngleSharp](https://github.com/EvotecIT/AngleSharp), [fork notes](AngleSharp/FORK.md) | Runtime worker |
| AngleSharp.Js bindings | [EvotecIT/AngleSharp.Js](https://github.com/EvotecIT/AngleSharp.Js), [fork notes](AngleSharp.Js/FORK.md) | Runtime worker |
| Jint | Unmodified NuGet dependency | Runtime worker and AngleSharp.Js |

Implement reusable fixes and their standalone regression tests in the owning fork.
Keep upstream history intact. Push the fork commit before updating its OfficeIMO
submodule revision, then validate the consuming runtime and build/publish artifacts.
Include both fork revisions and the OfficeIMO commit in validation evidence.

The worker copies each dependency's license into its output. Its isolated directory
contains the fork AngleSharp assembly; ordinary static consumers retain the official
package. Do not combine those assemblies in one application load context.

Source revision pins make development builds reproducible without publishing a
package. A later package release needs a distinct fork package identity, artifact
validation and an explicitly chosen publication destination. Fork assemblies must
not be published as official upstream packages.
