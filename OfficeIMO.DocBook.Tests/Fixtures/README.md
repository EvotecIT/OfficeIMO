# DocBook fixture provenance

The small DocBook documents in `DocBookDocumentTests` are repository-authored XML fixtures derived from the public OASIS DocBook XML 4.5 DTD identifiers and DocBook 5.2 standard profile. They cover articles, books, common structure, 4.5 and 5.2 creation, schema-identifier reporting, namespaced extensions, comments, DTD/entity policy, resource limits, shared-model conversion, byte-exact unchanged-source output, and reopen validation.

No OASIS schema files or third-party fixture bytes are redistributed or downloaded at runtime. OfficeIMO validation is intentionally the bounded common-structure profile; `IsOfficialSchemaValidated` remains false. Add independently produced files here when a producer-specific behavior becomes part of the compatibility contract, recording the producer, version, date, source, license, exact declared DocBook profile, and stable hash.
