# Browser reference for application fixtures

This test-only harness serves the four OfficeIMO runtime fixtures on localhost,
performs the same user actions as `RuntimeApplicationDocumentWorkflowTests`,
checks the resulting route and visible values, and saves Chromium screen and
print outputs. It uses Playwright 1.62.1 from the adjacent lockfile. No browser
package is referenced by an OfficeIMO runtime project.

From this folder, run:

```sh
npm ci
npx playwright install chromium
npm run capture -- ./output
```

The output directory contains `chromium-manifest.json` with the browser version,
viewport, final states, resource response hashes and errors. The harness serves
the checked-in [vanilla](../Fixtures/StandaloneApplication/README.md),
[React build](../Fixtures/ReactBuild/README.md),
[Preact](../Fixtures/Preact/README.md) and
[legacy](../Fixtures/LegacyApplication/index.html) inputs. It supplies the
same Preact startup call and data as the runtime test, returns the vanilla
health response as HTTP 503, and serves the vanilla review route as a second
document. Re-run the OfficeIMO output test with
`OFFICEIMO_APPLICATION_EVIDENCE_DIR` set to an output directory to retain its
three selected outputs per case. This harness is for a named fixture comparison,
not a general website test.
