# Structured format conformance

This opt-in runner checks generated ADF JSON against Atlassian's full schema and DocBook output against OASIS DocBook 5.2 RELAX NG. It includes valid output, deliberately invalid shapes, empty native arrays, block marks, typed content ordering, and the independently produced DocBook fixture. Each schema download must match its checked-in SHA-256 pin; changed schemas require review.

Use Python with `jsonschema` and `lxml` in an isolated validation environment, plus the repository's pinned .NET SDK:

```sh
python verify_schemas.py --output /path/to/task-owned/conformance-output
```

The runner writes a manifest, schema snapshots, native outputs and `results.json` to the supplied directory. A mismatch fails the command. Remove that directory when its evidence is no longer needed.

The project is outside the normal solution and is not shipped. These schema validators are independent test tools; OfficeIMO does not download schemas or invoke them at runtime. `DocBookValidationResult.IsOfficialSchemaValidated` continues to describe the native bounded validator, which does not run the official schema.
