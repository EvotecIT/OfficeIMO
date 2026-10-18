# Building OfficeIMO and PSWriteOffice

OfficeIMO contains the document engines and the PSWriteOffice PowerShell module.
The module source, tests, examples, and command documentation live in
[`PowerShell/PSWriteOffice`](../PowerShell/PSWriteOffice/README.MD).
Builds use the SDK selected by `global.json` and PSPublishModule 3.0.168 or newer.

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

## Website documentation

The website reads the module from this checkout. Its PowerShell API step builds an
unsigned module, exports help and command metadata, verifies the source revision
and component fingerprints, and imports the command catalog, guides, and examples.
No separate PSWriteOffice clone or repository release pin is required.

Keep `/api/powershell/`, `/docs/pswriteoffice/`, and `/products/pswriteoffice/` as
the public routes. Update compiled cmdlet help and the module's maintained guide
sources, then run the website pipeline; do not hand-edit generated command pages
or API metadata.
