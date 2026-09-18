# Child-frame module and JSON import evidence

This evidence qualifies same-origin child-frame module graphs and JSON import
attributes through the complete `NetworklessRootlessOciV1` acquisition, execution,
capture, screen, print, and screen-to-page pipeline.

The Windows host used WSL2 Ubuntu 24.04.3 LTS, rootless Podman 4.9.3, and .NET
10.0.112. The renderer ran from immutable image
`sha256:b54f04cec83c05211a6188a4bf95bbdb665e77990ed203293dc625cf6ae24da1`
with network disabled, a read-only root filesystem, seccomp and cgroup limits,
UID 65532, dropped capabilities, and no mounts. `controlled-summary.json` records
16 passing render cases, four passing acquisition cases, exact payload hashes,
and confirmed removal of every case container.

The `frame-module-json` case discovered five ordered rounds:

1. the child HTML document;
2. its external module root;
3. its JavaScript dependency;
4. a statically imported `application/json` module; and
5. a dynamically imported `application/vnd.officeimo+json` module.

The child used its own import map, awaited the dynamic JSON import, updated its
document, and sent the computed value to the parent. The parent and child both
rendered `Frame module graph ready 42`. The PNG was visually inspected: both
lines are readable and the child document remains clipped to its light-blue frame
area. Both PDFs reopen as one tagged A4 page and contain the computed text.

The `json-module-rejections` case acquired a syntactically invalid media type and
malformed JSON in one bounded static discovery round, then imported both resources
inside the networkless container. It rendered `JSON rejected TypeError/SyntaxError`,
proving the public workflow's MIME-mismatch and parse-failure distinction.

The output SHA-256 digests are:

| Artifact | SHA-256 |
| --- | --- |
| `screen.png` | `d2e170e3cb53a7b3292d495acef079fe19393129a7276ea52931ec731c456749` |
| `print.pdf` | `6e1ffd49aeb958641b23fd7c3e85ba24f680825772cd8641242f9aabfdbf42dc` |
| `screen-to-page.pdf` | `82687ecc2042a165d96e8f53c906ac0ffaa12cc701a2134108cff85101cf4537` |

The rejection output SHA-256 digests are:

| Artifact | SHA-256 |
| --- | --- |
| `json-module-rejections/screen.png` | `9f3598510f3b398d4f652031db97544d8fd4756a1d249eadb3e5c47dfda99601` |
| `json-module-rejections/print.pdf` | `5c12a7a83b30b6aaa463b32f7b49fadc228d19064a4fc53aff93a00a435dd478` |
| `json-module-rejections/screen-to-page.pdf` | `5c156460e6c4655af5cb713d9378252b1c9b1f379eb671c9c36af6a77e1c0063` |

Managed validation passed 468 runtime and rendering tests on both .NET 8 and
.NET 10. The broader HTML suite passed 3,164 tests on each framework. These
results qualify the selected same-origin and JSON-module contract; they do not
claim cross-origin frame execution or module types beyond JSON.
