# Production-built application acceptance slice

An independently authored report app was built from pinned React 18.3.1 and
esbuild 0.25.12 sources into minified ES modules with a shared chunk and a
dynamically loaded review chunk. The [fixture](../../../../../OfficeIMO.Html.Runtime.Tests/Fixtures/ReactBuild/README.md)
records the exact inputs, build command, package lock, license and output
hashes. The test supplies those assets offline to `WebApplicationV1`, performs
accessible-name form actions, changes route, captures the review screen, ends
the worker session, and renders the retained document in all three output
intents. React and esbuild remain test-only dependencies.

Run the selected acceptance case with:

```sh
dotnet test OfficeIMO.Html.Runtime.Tests/OfficeIMO.Html.Runtime.Tests.csproj -f net10.0 --filter FullyQualifiedName~RuntimeReactBuildApplicationTests
```

On Windows x64, the selected test passed under .NET 8 and .NET 10. At commit
`ed980db7a4`, the complete HTML suite passed 3,148/3,148 and the complete runtime
suite passed 354/354 on each framework. The same commit passed both suites under
.NET 10 on Linux x64 and macOS Arm64. The first macOS HTML run had one failure in
`ConcurrentDetachedProjectionAndCallbackSnapshotsRetainNodeIdentity`; that case
passed alone, and a full rerun passed 3,148/3,148 without a code change. The
captured route was
`/review`, the selected region was South, the controlled title was Quarterly,
and the adjusted total was 21. Both PDFs reopened as tagged, searchable files
with the selected title and total. The screen PNG decoded and retained the
authored `#f3f6fa` background. Independent Chromium 152.0.0.0 on Linux x64,
driven with Playwright CLI 0.1.20 against the same checked-in assets, reached
the same route and semantic values after the same actions.

| Intent | OfficeIMO | Browser reference |
| --- | --- | --- |
| Screen at 816 CSS px | [PNG](officeimo-screen.png) | [Chromium screenshot](chromium-screen.png) |
| Print | [PDF](officeimo-print.pdf), [preview](officeimo-print-preview.png) | [Chromium PDF](chromium-print.pdf), [preview](chromium-print-preview.png) |
| Screen-to-page PDF | [PDF](officeimo-screen-to-page.pdf), [preview](officeimo-screen-to-page-preview.png) | Separate OfficeIMO output intent |

The visual comparison exposed generic static-renderer gaps: default table
captions were left-aligned, header cells were not bold, and normal-flow block
auto margins did not absorb free width. Those defaults and auto margins now have
focused tests and produce a closer screen and print layout. The retained images
still show small text metric and vertical-spacing differences; this is a named
application workflow, not a pixel-parity or arbitrary-website claim. The PDF
previews were rasterized with Poppler 26.07.0. This slice has runtime proof on
Windows, Linux and macOS, plus the separate Linux browser reference.

The session uses explicitly trusted inputs. Public-site scripts still require
the separate isolation profile tracked in the roadmap.
