# React 18 application fixture

This test-only fixture pins the unmodified React and React DOM 18.3.1 production
UMD bundles. `app.js`, `index.html`, `style.css`, and `data.json` are independently
authored OfficeIMO acceptance inputs. The app mounts through `createRoot`, loads
data with `fetch`, updates hooks and controlled form state through real actions,
and prepares a semantic report for capture and rendering. All assets load offline;
no React package enters an OfficeIMO runtime dependency graph.

The bundles came from the official npm tarballs with their published SHA-512
integrities verified before extraction:

| Package | npm tarball | SHA-512 integrity | Fixture entry | SHA-256 |
| --- | --- | --- | --- | --- |
| React 18.3.1 | [react-18.3.1.tgz](https://registry.npmjs.org/react/-/react-18.3.1.tgz) | `sha512-wS+hAgJShR0KhEvPJArfuPVN1+Hz1t0Y6n5jLrGQbkb4urgPE/0Rve+1kMB1v/oWgHgm4WIcV+i7F2pTVj+2iQ==` | `package/umd/react.production.min.js` | `d949f1c3687aedadcedac85261865f29b17cd273997e7f6b2bfc53b2f9d4c4dd` |
| React DOM 18.3.1 | [react-dom-18.3.1.tgz](https://registry.npmjs.org/react-dom/-/react-dom-18.3.1.tgz) | `sha512-5m4nQKp+rZRb09LNH59GM4BxTh9251/ylbKIbpe7TpGxfJ+9kv6BLkLBXIjjspbgbnIBNqlI23tRnTWT0snUIw==` | `package/umd/react-dom.production.min.js` | `35f4f974f4b2bcd44da73963347f8952e341f83909e4498227d4e26b98f66f0d` |

Both packages use the MIT license. Their identical `package/LICENSE` is retained
once as `LICENSE` (SHA-256 `52412d7bc7ce4157ea628bbaacb8829e0a9cb3c58f57f99176126bc8cf2bfc85`).

Passing this fixture qualifies the selected application path. It does not claim
general React, React Router, hydration, concurrent scheduling, browser API, or
hostile-script compatibility.
