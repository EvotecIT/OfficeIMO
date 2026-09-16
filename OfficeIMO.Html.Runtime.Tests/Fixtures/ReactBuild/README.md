# Production-built React application fixture

This test-only application is independently authored for OfficeIMO. Its checked-in
`dist` was built from `src` with the exact versions in `package-lock.json`:
React and React DOM 18.3.1, and esbuild 0.25.12. The source is intentionally a
normal JSX/ES-module application. esbuild creates a production, minified entry,
a shared chunk and a dynamically loaded review-screen chunk. A normal Node.js
installation can rebuild it with `npm ci && npm run build` from this folder.

The app fetches report data, updates controlled form fields, changes route with
the History API, loads the second screen with dynamic import, and renders both
screens with React. The runtime test performs real locator actions, freezes an
independent review document, and produces screen, print and screen-to-page
outputs from retained resources. All runtime assets are supplied offline. React
and esbuild are not OfficeIMO runtime dependencies.

`LICENSE` retains the upstream React/React DOM MIT license. The exact upstream
package tarball integrities and the original React bundle provenance are also
recorded in the adjacent `React18` fixture. `package-lock.json` records the
resolved package URLs and integrity values for this build. The checked-in outputs
have these SHA-256 digests:

| Output | SHA-256 |
| --- | --- |
| `dist/app.js` | `c61c4452acf7ac4b0ad1f755bc8d8648fb0dc4d34d29f30f00f42d029fd6e5aa` |
| `dist/chunk-chunk-YDHNFGZD.js` | `6cabbc8d3b4c384f657366aaad563f3786eaf09d898be56685fe00607fa527fc` |
| `dist/chunk-review-7VOCLQDE.js` | `bf9519f21a60e566bfac76fe55e10cce38b27a6ca1527e8fd1b771e7b18808eb` |

Passing this case qualifies the named build and interactions. It does not imply
general React, hydration, routing-library, or browser compatibility.
