# Atlassian ADF schema

OfficeIMO.Adf embeds the unmodified `dist/json-schema/v1/full.json` from
Atlassian's `@atlaskit/adf-schema` version **57.6.21**.

- Source: [Atlassian's published npm package](https://registry.npmjs.org/@atlaskit/adf-schema/-/adf-schema-57.6.21.tgz)
- Schema SHA-256: `5128562B75278C8A83E7E3619A570205BC80D59696985EC31A7A7883CFF66FBE`
- Copyright 2019 Atlassian Pty Ltd
- License: Apache License, Version 2.0

The upstream copyright notice is retained in `Schema/UPSTREAM-LICENSE.txt`.
The full license is in `Schema/LICENSE-APACHE-2.0.txt`. Both are included in
the NuGet package. The OfficeIMO validation code is independently implemented
and uses the repository's MIT license.

The pinned schema describes ADF syntax. Acceptance by an Atlassian product
also depends on that product and the destination field.
