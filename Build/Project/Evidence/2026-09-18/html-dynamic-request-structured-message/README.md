# Dynamic request replay and structured message evidence

This evidence qualifies exact same-origin dynamic request replay through the
complete `NetworklessRootlessOciV1` parsing, scripting, capture, screen, print,
and screen-to-page pipeline. It also records the managed cross-realm structured
message contract.

The Windows host used WSL2 Ubuntu with rootless Podman. The renderer ran from
immutable image
`sha256:abc78f0ce749b1f988cf6a27d0d2bac2a84e07187dbd8458e84a47890e44153a`
with network disabled, a read-only root filesystem, seccomp and cgroup limits,
UID 65532, dropped capabilities, and no mounts. `controlled-summary.json`
records 17 passing render cases, four passing acquisition cases, exact payload
hashes, ordered discovery, and confirmed removal of every case container.

The `dynamic-xhr-headered-get` case discovered one exact request:

1. `GET https://fixture.officeimo.invalid/api/header-varying.json #1`

The fixture required the `X-Variant: private` request header before returning
the response. The networkless rerun rendered `Headered XHR ready 42`.

The `dynamic-post-occurrences` case used the same URL, headers, and JSON body
twice. Discovery retained two occurrences in separate rounds:

1. `POST https://fixture.officeimo.invalid/api/submit #1`
2. `POST https://fixture.officeimo.invalid/api/submit #2`

The host fixture supplied `first` and `second` independently. The final
networkless run consumed both exact responses and rendered
`Dynamic POST ready first/second`. This proves that discovery restarts do not
repeat an already acquired POST and do not collapse repeated identical requests.

Both PNGs were visually inspected at 816 x 720. Their expected blue text is
readable without clipping or fallback content. Both PDF modes reopened during
the probe and contained the same script-produced text.

The output SHA-256 digests are:

| Case | Screen PNG | Print PDF | Screen-to-page PDF |
| --- | --- | --- | --- |
| Headered XHR | `463c282f3f0ddab9029616ac33222fc52304747c319a78d8a9b7869b5298e643` | `97e6ee7477661344dbbad8a16e2e9a3391577fd3691b9a5b6b9330e0961e8644` | `b190887457cdf38fae9916c694456f4aaefe741684dd5ce4e973f49d5fc3fabe` |
| Repeated POST | `f662d5d5f8b7139bbae18ac767b296bba8ebcf7c127fa14a76986d11aecd5b49` | `e70a1d09997c426e3034d6787b457fd683c918e8436c2a810cf676164be6428a` | `9abb5e8d4d619e25e744a47a8858423d26e66d80905440ba018866705958210c` |

Managed frame tests qualify synchronous source-realm snapshots and independent
target-realm reconstruction for cycles, repeated references, Map, Set, Date,
RegExp, BigInt, undefined, special numeric values, Error with cause,
ArrayBuffer, DataView, and typed arrays. Functions, symbols, exotic host
objects, transfer lists, message ports, and the `postMessage(message, options)`
overload remain outside this contract.

Managed validation passed 485 runtime and rendering tests on both .NET 8 and
.NET 10. The broader HTML suite passed 3,164 tests on each framework. Dynamic
live acquisition remains same-origin and credentialless. GET and HEAD are
enabled by default; non-GET methods require explicit caller authority. Dynamic
redirects and cross-origin requests remain unsupported.
