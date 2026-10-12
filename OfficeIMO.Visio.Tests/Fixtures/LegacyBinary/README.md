# Independent binary Visio fixtures

The PRONOM corpus supplies unmodified Microsoft Visio drawing, stencil and
template files for binary version 11, together with their paired XML exports.
`producer-manifest.json` records source revision, URLs and SHA-256 hashes;
`pronom-NOTICE.txt` records the upstream CC0 declaration. The older version 5
file qualifies explicit rejection rather than older-generation support.

`VisioLegacyBinaryTests` checks shape identities and nesting, master references,
cached transforms and labels against the XML model, modern package reopening,
SVG/PDF generation, text-only master children, cached fill transparency, Reader dispatch,
input ownership and resource limits. Native
binary input is also readable by the independent libvisio 0.1.11 tools used for
opt-in qualification; those tools are not required by ordinary tests or shipped
packages.

This corpus covers one diagram family. It does not establish broad binary
compatibility, native field/formula editing, binary save-back or Microsoft Visio
application acceptance. The import report describes the selected cached profile
and omitted records.
