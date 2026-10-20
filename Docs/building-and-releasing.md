# Building OfficeIMO and PSWriteOffice

OfficeIMO contains the document engines and the PSWriteOffice PowerShell module.
The module source, tests, examples, and command documentation live in
[`PowerShell/PSWriteOffice`](../PowerShell/PSWriteOffice/README.MD).
Builds use the SDK selected by `global.json` and PSPublishModule 3.0.170 or newer.

## Local validation

The PowerShell projects reference the OfficeIMO projects in this checkout. A
normal solution build and the PowerShell CI workflows therefore exercise engine
and cmdlet changes together.

For a local unsigned module build without packaging the complete NuGet family:

```powershell
./Build/Build-Module.ps1 -IncludePackages:$false
```

This produces the native module artifacts without installing the module. On
Windows, the module includes the PowerShell 7 and Windows PowerShell 5.1 assemblies.
Unix validation builds the PowerShell 7 assembly. A public module release is built
on Windows so both runtime payloads are present.

Regenerate command Markdown and external help through the native documentation
lane after editing cmdlet XML comments:

```powershell
./Build/Build-Module.ps1 -RunMode Documentation -IncludePackages:$false
```

This lane does not sign, package, publish, install, or advance release versions.

For the combined package and module build:

```powershell
./Build/Build-Release.ps1
```

Build mode does not publish or use signing keys. PowerForge stages the OfficeIMO
packages first, uses `OfficeIMO.Word` as the version anchor, and synchronizes
PSWriteOffice to that package version. The module uses the same versioned engine
source. `Build/project.build.json` owns package selection, version bindings,
and NuGet configuration.

## Coordinated publication

Publication uses the local credential files configured in
`Build/project.build.json` and the module build configuration. Manifest and
unsigned build modes do not read these files; the shared publisher resolves
them during publication.

PowerForge generates both packaged loaders and the source development
bootstrapper. Development imports prefer Debug, then Release;
`PSWRITEOFFICE_DEVELOPMENT_CONFIGURATION` selects a configuration explicitly.
The build synchronizes the compiled module's project version with the manifest
and derives command exports and help from the binary and XML comments.

After the candidate has passed its required tests and reviews, an authorized
release operator runs:

```powershell
./Build/Build-Release.ps1 -RunMode Publish -SignModule:$true
```

PowerForge owns package and module construction, signing, artifact selection,
credentials, and publication. The configured order is NuGet, PowerShell Gallery,
then GitHub. NuGet packages and PSWriteOffice share the resolved three-part public
version. The module ZIP joins the package assets in the same
`OfficeIMO-v<version>` GitHub release. The existing release workflows build the
agent plugin and refresh the website from that source revision.

`Build/Build-Project.ps1 -PublishNuget:$true` and its `PublishGitHub` option delegate
to the coordinated release. They no longer publish a separate engine-only release.
Unsigned builds and module-only builds cannot enter Publish mode.

Registries do not share a transaction. If publication stops after a destination
has completed, inspect the failure and resume the native checkpoint:

```powershell
./Build/Build-Release.ps1 -RunMode Publish -SignModule:$true -ResumeIncompleteRelease
```

Retain the source candidate and the PowerForge checkpoint until release verification
finishes. Verify the actual NuGet packages, downloaded PSGallery module, GitHub
assets, and deployed website before treating delivery as complete.

## JavaScript npm delivery

The `Publish coordinated npm package` workflow delivers `@evotecit/officeimo`
from the same normal OfficeIMO GitHub release. It creates no separate npm release
tag or GitHub release. `Build/project.build.json` binds the source package and
lockfile versions to `OfficeIMO.JavaScript.Assets` when the native publisher updates versions.

The workflow reads the actual version from the unique signed
`OfficeIMO.JavaScript.Assets.<version>.nupkg` attached to the release. The release tag may
contain the native publisher's earlier version metadata, so the runner stamps
the package and lockfile before packing. Native publication builds the
operator-selected working tree. Committing version metadata before that native
publication remains an operator choice.
The workflow waits up to ten minutes for the JavaScript asset because native
publication creates the GitHub release before uploading its package assets.

Before publication, the workflow runs the JavaScript contracts and isolated
strict TypeScript, Node and worker consumers. It compares every bundled JavaScript
and module asset with the signed JavaScript asset assembly's embedded resources. This
proves runtime-byte equivalence between the npm and NuGet JavaScript bundles; it
does not claim that every type source or native build input is identical. A
runtime mismatch stops npm delivery and leaves the completed native publication
intact.

The shared PowerForge npm publisher uploads the same qualified archive and
verifies registry identity, integrity and public-source provenance. This separate
workflow has its own retries; it is outside the native PowerForge release
checkpoint. Retry it on the original tag after diagnosing a delivery failure:

```powershell
gh workflow run npm-release.yml --ref OfficeIMO-v<version> -f publish=true
```

The original fourteen-digit `OfficeIMO-v<timestamp>` tags are also accepted.
A successful master push `Test .NET Libraries` run must qualify the exact tagged commit.
An existing npm version succeeds only when its bytes and provenance match the
qualified archive; conflicting bytes stop publication rather than replacing it.

Configure the `npm-release` GitHub environment with one deployment policy: type
`tag`, name `OfficeIMO-v*`. Register npm Trusted Publishing for organization
`EvotecIT`, repository `OfficeIMO`, calling workflow `npm-release.yml`, and that
environment. Ordinary releases use OIDC and no persistent npm token. The first
publication requires an operator-created environment secret named
`NPM_BOOTSTRAP_TOKEN` and an explicit tag dispatch with `publish=true` and
`bootstrap=true`. Remove that secret after verified bootstrap and configure the
package's trusted publisher. The shared
[npm publication guide](https://github.com/EvotecIT/PSPublishModule/blob/main/Docs/PowerForge.NpmPublishing.md)
owns credential setup and registry recovery details.

## Website documentation

The website reads the module from this checkout. Its PowerShell API step builds an
unsigned module, exports help and command metadata, verifies the source revision
and component fingerprints, and imports the command catalog, guides, and examples.
No separate PSWriteOffice clone or repository release pin is required.

Keep `/api/powershell/`, `/docs/pswriteoffice/`, and `/products/pswriteoffice/` as
the public routes. Update compiled cmdlet help and the module's maintained guide
sources, then run the website pipeline; do not hand-edit generated command pages
or API metadata.
