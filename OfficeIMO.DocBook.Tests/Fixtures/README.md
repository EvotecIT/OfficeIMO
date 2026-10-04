# DocBook fixture provenance

The small DocBook documents in `DocBookDocumentTests` are repository-authored XML fixtures derived from the public OASIS DocBook XML 4.5 DTD identifiers and DocBook 5.2 standard profile. They cover articles, books, common structure, 4.5 and 5.2 creation, schema-identifier reporting, namespaced extensions, comments, DTD/entity policy, resource limits, shared-model conversion, byte-exact unchanged-source output, and reopen validation.

`pandoc-3.12-common-structure.docbook` was generated on 2026-10-03 by [Pandoc 3.12](https://github.com/jgm/pandoc/releases/tag/3.12) from the repository-authored `pandoc-common-structure.md`. The input and generated fixture are covered by this repository's MIT license; the producer executable uses GPL licensing and is not redistributed or required by OfficeIMO.

Reproduce it with:

```sh
pandoc --from markdown --to docbook5 --standalone pandoc-common-structure.md --output pandoc-3.12-common-structure.docbook
```

The source declares DocBook 5.0 with the DocBook namespace. It covers metadata, a section, lists, emphasis, literal text, an external link, a code block and a table. Native tests check unchanged byte preservation and common-content projection. The opt-in schema runner separately checks it against the official 5.2 RELAX NG schema.

- Producer archive: `pandoc-3.12-arm64-macOS.zip`, SHA-256 `f148ca09c9f36594db527a9fc988ad736290ce428f79594c50208cd1ec58b3c0`.
- Generated fixture SHA-256: `7e31ba064dc1398d7a9621fc29463c010494226ebd8df10d56326128de16186d`.

No producer executable or OASIS schema files are downloaded at runtime. OfficeIMO validation remains the bounded common-structure profile, and `IsOfficialSchemaValidated` remains false.
