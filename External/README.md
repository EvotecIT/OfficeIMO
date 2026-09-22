# HTML runtime providers

The interactive worker builds the AngleSharp, AngleSharp.Js and Jint forks from the exact
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
| Jint interpreter | [EvotecIT/jint](https://github.com/EvotecIT/jint), [fork notes](Jint/FORK.md) | Runtime worker and AngleSharp.Js |

Implement reusable fixes and their standalone regression tests in the owning fork.
Keep upstream history intact. Push the fork commit before updating its OfficeIMO
submodule revision, then validate the consuming runtime and build/publish artifacts.
Include both fork revisions and the OfficeIMO commit in validation evidence.

The worker copies each dependency's license into its output. Its isolated directory
contains the fork AngleSharp assembly; ordinary static consumers retain the official
package. Do not combine those assemblies in one application load context.

With both submodules present, the AngleSharp.Js integration build selects the
adjacent AngleSharp core project so its native listener-removal hook comes from
the pinned source. A standalone AngleSharp.Js checkout needs the matching core
source path passed as `AngleSharpTestProject`; its current branch is not a
standalone package against the published AngleSharp 1.8.0 API.

Source revision pins make development builds reproducible without publishing a
package. A later package release needs a distinct fork package identity, artifact
validation and an explicitly chosen publication destination. Fork assemblies must
not be published as official upstream packages.

## Contributing while retaining a usable fork

Prepare each upstream candidate from its upstream base with one reusable correction
and its regression proof. Keep ongoing integration on the fork's runtime branch;
OfficeIMO consumes qualified revision pins independently of upstream PR decisions.

If a patch is accepted, qualify the upstream release before removing the fork delta.
If it is declined, retain the smallest tested patch needed by the supported contract
and record the maintainer's reason with the patch's upstream link. Rework it when a
better extension point is agreed. Host-specific policy stays in OfficeIMO. A declined
PR alone is not a reason to replace a working dependency.

The combined binding suite can use the pinned core source:

```sh
dotnet test External/AngleSharp.Js/src/AngleSharp.Js.Tests/AngleSharp.Js.Tests.csproj \
  -f net10.0 -p:AngleSharpTestProject="$PWD/External/AngleSharp/src/AngleSharp/AngleSharp.Core.csproj"
```
