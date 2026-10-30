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

## README examples

The Word and Excel quick starts are generated from the executable programs under
`Build/DocumentationExamples`. Edit those source files, then refresh the marked
README blocks with PowerForge:

The standalone CLI comes from the `PowerForge.Build` .NET tool, separately from
the PSPublishModule PowerShell module. To use the same implementation as CI,
check out the PowerForge revision pinned in
`.github/workflows/documentation-examples.yml`, then run from its checkout:

```powershell
dotnet build PowerForge.Cli/PowerForge.Cli.csproj -c Release -f net10.0
```

From the OfficeIMO repository root, invoke that built CLI with
`dotnet <PowerForge checkout>/PowerForge.Cli/bin/Release/net10.0/PowerForge.Cli.dll`
followed by `docs sync --config Build/docs-sync.json`, adding `--apply` to refresh.
When using an installed CLI with this command, the equivalent calls are:

```powershell
powerforge docs sync --config Build/docs-sync.json --apply
powerforge docs sync --config Build/docs-sync.json
```

The check command returns a nonzero exit code when a block is stale. The
documentation examples workflow runs the same check and executes both programs
against their pinned published NuGet packages. Update
`OfficeIMODocumentationPackageVersion` in the example projects when qualifying a
new public version. Source synchronization and published-package compatibility
are checked separately.

To run an example locally from the repository root, use `dotnet run --project
Build/DocumentationExamples/Word/Word.csproj -c Release -f net8.0` or the matching
Excel project. The programs write `report.docx` and `report.xlsx` into the current
directory. From a scratch working directory, pass an absolute project path.

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

`Publish npm package` delivers `@evotecit/officeimo` after `Test .NET Libraries`
succeeds for a push to `master` in this repository. The workflow checks out that
exact qualified commit. It compares the package inputs with the source tag of
the highest stable version actually published to npm. Changes to TypeScript
sources, bundles, package metadata, package documentation, build scripts or the
npm release workflow qualify. Merges containing only .NET or JavaScript tests
are skipped when no unreleased package-input changes remain.
A pending version reservation is not treated as a published baseline. Missing
source tags or diverged history stop the release rather than silently skipping
changes. The path-gate contracts run in the browser-export CI lane.

The shared PowerForge `npm-release-version` action selects the next free patch
and reserves its `npm-v<version>` tag without committing to `master`. Reruns reuse
the reservation, and newer reserved source prevents an older upload from
becoming latest. npm versions are independent of NuGet/PSWriteOffice versions;
the native publisher does not stamp the npm manifests. Major and minor releases
use the manual workflow's exact `version` input. A manifest minimum that raises
the major or minor line stops automatic publication until that manual release.

The package job compiles TypeScript, checks the committed bundles, runs the
JavaScript contracts and qualifies an isolated archive consumer. It also builds
the small `OfficeIMO.JavaScript.Assets` NuGet package locally and compares every
embedded runtime bundle with the npm bundle bytes. This unsigned local package
is an integrity check, not a public .NET release or signing requirement.
`Get-NpmReleaseVersion.ps1` remains available for operators checking a signed
native-release package; its signature verification is unchanged.

The shared PowerForge publisher receives the exact qualified archive, its
SHA256 and the successful quality run. It publishes with lifecycle scripts
disabled and verifies registry identity, SHA512 integrity and public-source
provenance. GitHub's OIDC source identity must match the qualified commit. If a
newer merge advances that identity, the older upload stops; the newer commit's
successful quality run can deliver the cumulative package changes.

### Owner setup and first publication

The package owner creates the `npm-release` GitHub environment with a custom
deployment policy allowing exactly one **branch**, `master`. Do not add tag
policies or wildcard patterns. The shared publisher verifies this policy.
Agents do not create the environment or its credentials.

For the first publication, the owner creates a short-lived npm granular token
with permission to create/publish `@evotecit/officeimo` in the `@evotecit` scope,
package-write permission and bypass 2FA for unattended publishing. Store its
value only as the **environment secret** `NPM_BOOTSTRAP_TOKEN` in `npm-release`;
there is no repository-level copy. Do not put the value in workflow inputs,
commands, issues or chat. Automatic runs skip an unpublished package and reserve
no version, so they cannot create a release environment or attempt OIDC bootstrap.

After this workflow PR is merged and `master` quality has succeeded for the
current source, dispatch on `master`:

```powershell
gh workflow run npm-release.yml --repo EvotecIT/OfficeIMO --ref master `
  -f publish=true -f bootstrap=true -f quality-run-id=<successful-master-run-id>
```

The optional `version` input selects an explicit three-part version; otherwise
the shared action uses the manifest minimum for the initial publication and the
next free patch thereafter. A manual dispatch with `publish=false` performs the
normal qualified dry run. Bootstrap requires `publish=true` and an authenticated
registry lookup confirming the package does not exist.

After verifying the public version, archive integrity and a clean installed
consumer, the owner configures npm Trusted Publishing for organization
`EvotecIT`, repository `OfficeIMO`, calling workflow **`npm-release.yml`** and
environment **`npm-release`**. Register the calling workflow, not the shared
`powerforge-npm-publish.yml`. Remove the bootstrap environment secret after the
trust configuration is ready. Ordinary qualifying merges then publish through
OIDC with `bootstrap=false`. A direct .NET/Gallery/GitHub release remains a
separate operation and does not trigger or gate npm delivery.

The shared [npm publication guide](https://github.com/EvotecIT/PSPublishModule/blob/main/Docs/PowerForge.NpmPublishing.md)
owns credential setup and registry recovery. If an upload succeeds but readback
fails, inspect the public version and provenance before retrying; immutable
versions are never replaced with another archive.

## Website documentation

The website reads the module from this checkout. Its PowerShell API step builds an
unsigned module, exports help and command metadata, verifies the source revision
and component fingerprints, and imports the command catalog, guides, and examples.
No separate PSWriteOffice clone or repository release pin is required.

Keep `/api/powershell/`, `/docs/pswriteoffice/`, and `/products/pswriteoffice/` as
the public routes. Update compiled cmdlet help and the module's maintained guide
sources, then run the website pipeline; do not hand-edit generated command pages
or API metadata.
