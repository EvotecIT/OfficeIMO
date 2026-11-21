# JSON Schema validation design

OfficeIMO validates configuration files against JSON Schema without a runtime dependency. Kubernetes manifests, GitHub
Actions workflows, docker-compose files and `pyproject.toml` are YAML or TOML, but their schemas are JSON Schema, and
editors and CI check them against those schemas. The validator answers the same question for OfficeIMO documents and
points each finding back at the line and column of the source file.

## Shape

One format-neutral validator in `OfficeIMO.Core`, namespace `OfficeIMO.Json`, and a thin adapter per format.

| Type | Job |
|---|---|
| `JsonData` | Immutable JSON data model value: null, boolean, number, string, array, object. Each value carries a `JsonDataLocation` (line, column, source node). `JsonData.Parse` reads JSON text (RFC 8259) with positions. |
| `JsonSchema` | A loaded schema. `Parse`/`Load`, then `Validate(JsonData)` for findings or `IsValid(JsonData)` to stop at the first one. |
| `JsonSchemaOptions` | Dialect, base URI, reference resolver, format assertion and the limits. |
| `JsonSchemaValidationResult` | `IsValid`, `Findings`, `IsTruncated`, `UncheckedFormats`. |
| `JsonSchemaFinding` | `InstancePath` and `SchemaPath` (JSON Pointers), `Keyword`, `Message`, `Line`, `Column`, `Source`. |

The adapters convert their model to `JsonData` and add `Validate(JsonSchema)`:

- `YamlNode.ToJsonData()`, `YamlDocumentPart.ToJsonData()`/`Validate()`, `YamlDocument.ToJsonData()`/`Validate()` (first document; validate `Documents` one by one for streams).
- `TomlNode.ToJsonData()`, `TomlDocument.ToJsonData()`/`Validate()`.

A converted tree rather than a node interface keeps the evaluator simple and fast: kinds, numbers and property
lookups are computed once, and a third format plugs in by building `JsonData` with locations. Schemas themselves are
`JsonData` too, so a schema can be written in JSON or YAML (`JsonSchema.Load(yaml.ToJsonData())`).

## Format mapping

**YAML** follows the document's schema, as typed reads do: under YAML 1.1 `yes` is a boolean and `010` is octal 8;
under the core schema `yes` is a string. Timestamps, custom-tagged scalars (`!Ref`) and scalars a tag claims but cannot
hold stay strings, so `!!int 1.5` is the string `1.5`. `<<` merge keys apply where the document's schema applies them: the mapping's own keys win, then
merged mappings in order. Aliases share their anchored value, scalars included, so an alias-heavy document converts in
linear time and memory; a finding at an alias reports the anchored node's line and column. Collection keys become their flow text. Duplicate keys follow the load's `DuplicateKeys` policy.

**TOML** tables become objects; arrays and arrays of tables become arrays; date-times, dates and times become their
RFC 3339 text. Offset date-times satisfy `date-time` and local dates `date`; local times and local date-times carry no
offset, so `time` and `date-time` reject them rather than assume a timezone.

**Numbers** keep their exact decimal value at any precision (an internal mantissa and exponent), so `0.1`,
`multipleOf: 0.01` and integers beyond `decimal` range compare exactly. Doubles from YAML and TOML keep the exact value
of their shortest round-trip text. Only infinities and NaN are inexact, and a finite number never equals one. Number
text with an exponent beyond nine digits is rejected rather than rounded to infinity or zero.

## Dialects and keywords

Draft 2020-12 is the default; `$schema` naming draft 7 or 6, or `JsonSchemaOptions.Draft`, selects draft 7 rules
(`definitions`, array `items` with `additionalItems`, `dependencies`, `$ref` overriding its siblings, `$id` fragments as
anchors). Draft 2020-12 also reads `dependencies` when `dependentRequired` and `dependentSchemas` are absent.

Supported: `type`, `enum`, `const`, the numeric and string keywords, `pattern`, `format`, `prefixItems`, `items`,
`contains` with `minContains`/`maxContains`, `uniqueItems`, `unevaluatedItems`, `properties`, `patternProperties`,
`additionalProperties`, `propertyNames`, `required`, `minProperties`/`maxProperties`, `dependentRequired`,
`dependentSchemas`, `unevaluatedProperties`, `allOf`, `anyOf`, `oneOf`, `not`, `if`/`then`/`else`, `$ref`,
`$dynamicRef`, `$defs`, `$id`, `$anchor` and `$dynamicAnchor`. Annotations for the `unevaluated*` keywords are collected
per instance location and dropped with failing subschemas, as the specification requires.

Not supported: custom meta-schemas and `$vocabulary` (every keyword is always active), switching dialect inside a
referenced resource, and assertion of `contentEncoding`/`contentMediaType` (annotations, as draft 2019-09 and later
require).

## References

`$ref` and `$dynamicRef` resolve against `$id` base URIs within the loaded schema, including embedded resources,
anchors and JSON Pointer fragments. Anything else goes to `JsonSchemaOptions.ReferenceResolver`, once per URI, with the
resolved URI without its fragment. Nothing is fetched from the network or the file system otherwise. Without
`BaseUri` or an absolute `$id`, relative references reach the resolver as written, so a caller can map them to files
next to the schema. A query-only reference such as `?version=2` keeps the base's path. Relative references drop `.` and `..` segments, so `./child.json` finds an embedded
`child.json`. A `BaseUri` with a scheme must be a valid absolute URI. A schema value reused under two resources, as a
YAML alias can do, resolves its references against whichever resource it is reached through. An unresolved reference
is a finding at the `$ref`, not an exception; inside `not`, `anyOf` and other branch checks it stops validation,
because a broken reference cannot count as a mismatch that turns into a pass.

## Formats

`ValidateFormats` (on by default) asserts `date-time`, `date`, `time`, `duration`, `email`, `idn-email`, `hostname`,
`idn-hostname`, `ipv4`, `ipv6`, `uri`, `uri-reference`, `iri`, `iri-reference`, `uri-template`, `uuid`, `regex`,
`json-pointer` and `relative-json-pointer`. Other format values are annotations, listed in `UncheckedFormats`.
Host names check A-labels by decoding Punycode; IDNA2008 code point validity, contextual rules, the Bidi rule and the
UTS 46 mapping are approximated from Unicode categories and script ranges rather than the full IDNA tables.

## Patterns

Patterns are ECMA-262 regular expressions, translated to .NET: `\d`, `\w`, `\s` and `\b` keep their ECMA-262 sets
(`\d` matches only ASCII digits), `$` matches only at the end, `.` and literal astral characters match whole code
points even when backtracking, captures keep their ECMA-262 numbers when named and unnamed groups mix or a name repeats across alternatives, a
backreference to a group that captured nothing matches empty, `(?m:...)` and
`(?s:...)` groups change `^`, `$` and `.` as ECMA-262 defines, and `\p{Letter}`-style property names map to .NET categories (`Script=` maps to the .NET block of that name). Syntax
ECMA-262 lacks, such as `\a`, `\z` or `(?#...)`, makes the schema invalid, and so does an astral character
inside a character class, which .NET would match one UTF-16 unit at a time. Every match runs with
`JsonSchemaOptions.RegexTimeout`.

## Limits

Schemas and documents can be untrusted. Every limit reached is a finding, so a result is never silently incomplete.

| Limit | Default | Bounds |
|---|---|---|
| `MaxErrors` | 100 | Findings per validation; `IsTruncated` reports the stop |
| `MaxDepth` | 256 | Nested schema evaluation, including followed references; catches `$ref` cycles |
| `MaxEvaluationSteps` | 10,000,000 | Total subschema evaluations and `uniqueItems` comparisons of items with equal hashes; catches exponential `anyOf`/`oneOf` nesting |
| `RegexTimeout` | 1 second | Each `pattern` and `patternProperties` match |

`JsonData` values nest at most 1024 levels however they are built, which bounds comparison, hashing, writing and
schema compilation. `JsonData.Parse` and the YAML and TOML conversions stop earlier, at 512 levels; YAML aliases
count the full depth of the value they reuse. `multipleOf` stays exact at any exponent gap. Values that share subtrees,
as YAML aliases do, are compared, hashed for `uniqueItems` and previewed in findings without expanding them.

## Errors

A schema whose keyword has the wrong type (`"minLength": "two"`) or whose pattern is not valid ECMA-262 throws
`JsonSchemaException` at load, with the schema line and column when the schema came from text. A schema that the
`ReferenceResolver` returns is compiled when a validation first reaches it; if it is invalid, that validation and every
later one that reaches it throw the same way. Validation findings otherwise never throw.

## Qualification

`OfficeIMO.Yaml.Tests/JsonSchema/JsonSchemaTestSuiteTests` runs the official
[JSON-Schema-Test-Suite](https://github.com/json-schema-org/JSON-Schema-Test-Suite) (MIT), vendored at a pinned commit in
`OfficeIMO.TestAssets/Json/JSON-Schema-Test-Suite` by `Update-JsonSchemaTestSuite.ps1`, for draft 7 and draft 2020-12
with the suite's compatibility rules; format-assertion cases run with `ValidateFormats` on and format-annotation cases
with it off. Known failures are listed with a reason, and a listed case that starts passing fails the run.

| Dialect | Required | Optional |
|---|---|---|
| Draft 2020-12 | 2081 / 2082 | 204 / 206 |
| Draft 7 | 1667 / 1667 | 146 / 152 |

`JsonDataParseTests` runs `JsonData.Parse` against JSONTestSuite's accept and reject cases.
