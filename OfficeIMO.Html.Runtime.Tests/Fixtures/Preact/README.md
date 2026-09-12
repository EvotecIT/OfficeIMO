# Preact application fixture

The unmodified Preact 10.29.8 UMD bundles exercise a real application's mounting,
hooks, fetch effects, controlled input, event updates, storage and unmount/remount
behavior. `report.js` is the OfficeIMO test application. These assets are test-only;
OfficeIMO runtime packages do not depend on Preact. Tests load all assets offline.

Upstream: [Preact](https://github.com/preactjs/preact), MIT license (retained in
`LICENSE`). The bundles were extracted from the
[10.29.8 npm tarball](https://registry.npmjs.org/preact/-/preact-10.29.8.tgz), whose
SHA-512 integrity was verified before extraction:

```text
sha512-ej2aVZ+vZ8WO7tvlQWRM9N63A0KzF9q4mWJfDUHgYaIofWY9hu74QdnQrjoPMmZi2/nZ5gN0bJCQF49xQqx09Q==
```

| Fixture | Original archive entry |
| --- | --- |
| `preact.umd.js` | `package/dist/preact.umd.js` |
| `hooks.umd.js` | `package/hooks/dist/hooks.umd.js` |
| `LICENSE` | `package/LICENSE` |

Passing this fixture establishes these application paths, not general Preact or
browser compatibility. Module loading, navigation, layout-driven interaction,
live form-state capture and combined mutation/promise ordering need separate
qualification.
