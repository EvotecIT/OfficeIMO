---
title: "Lint and format YAML from the CLI"
description: "Check YAML files against yamllint-style rules, apply safe fixes that change only the affected lines, and format files in one layout with officeimo yaml."
layout: docs
---

`officeimo yaml` checks and formats YAML files with OfficeIMO's own YAML 1.2 engine, the same `OfficeIMO.Yaml` package .NET code uses. It reads Kubernetes manifests, Helm values, GitHub Actions and Azure Pipelines workflows, Compose files and any other YAML, keeps comments and spellings, and needs no Python or Node.js.

## Lint

```shell
officeimo yaml lint .github/workflows
officeimo yaml lint values.yaml deployment.yaml --strict
```

Each finding is one line with the file, line, column, severity, message and rule, the layout yamllint prints:

```text
values.yaml:12:10: warning Truthy value "yes" is a string here, but YAML 1.1 readers such as PyYAML, go-yaml v2 and Ruby Psych read it as a boolean; quote it. (truthy)
values.yaml:18:1: error The mapping key 'image' at 18:1 repeats the key at 3:1. (key-duplicates)
```

A path can be a file or a folder; folders are searched for `*.yaml` and `*.yml` files. `-` reads standard input. The command returns exit code `1` when a file has errors, or warnings as well with `--strict`, so it can gate a pull request.

| Rule | Default | Reports |
|---|---|---|
| `syntax` | always, error | Text that is not valid YAML. The line rules still run on it. |
| `key-duplicates` | on, error | A key that repeats an earlier key of the same mapping. |
| `indentation` | on | Nested mappings indented by a different width than the file uses first, and sequences mixing `key:\n- item` with `key:\n  - item`. |
| `truthy` | on | `yes`, `no`, `on`, `off` and their capitalized forms, which YAML 1.1 readers take as booleans and YAML 1.2 readers as strings. Keys are not checked, so a GitHub Actions `on:` is fine. |
| `octal-values` | on | Integers such as `0755`, an octal 493 to YAML 1.1 readers and a decimal 755 to YAML 1.2 readers. |
| `trailing-spaces` | on | Spaces or tabs at the end of a line. |
| `line-length` | on, 120 | Longer lines, except one long word or `key: long-word` such as a URL. |
| `comments` | on | `#comment` without a space, and comments closer than two spaces to the content before them. |
| `tabs` | on | Tab characters in indentation. |
| `new-line-at-end-of-file` | on | A file that does not end with a line break. |
| `empty-lines` | on | More than two blank lines in a row, or blank lines at the start or end of the file. |
| `document-start` | off | A missing `---` at the start of the file. |
| `empty-values` | off | `key:` or `-` without a value, which reads as null. |
| `quoted-strings` | off | String values that are not quoted. |
| `key-ordering` | off | Keys that are not in ascending order. |

Turn rules on or off with `--enable` and `--disable`, set the indentation width with `--indent`, and the line limit with `--max-line-length`:

```shell
officeimo yaml lint config --enable document-start,key-ordering --disable line-length --indent 2
```

`--format json` writes one report with the schema `officeimo.yaml.lint.v1`: totals of errors, warnings and fixes, and for each file its findings with `rule`, `severity`, `line`, `column`, `message` and whether the finding is `fixable`.

## Fix

```shell
officeimo yaml lint values.yaml --fix
```

`--fix` rewrites each file with the repairs that cannot change a value: trailing spaces and extra blank lines removed, a space added after `#` or before an inline comment, the final line break added, `yes` quoted as `'yes'` where it is a string (or written `true` in a `%YAML 1.1` document, where it is a boolean), and the `---` marker added when `document-start` is enabled. Only the lines a fix repairs change; comments, quoting and indentation elsewhere stay byte for byte. The fixed text is loaded again and compared with the original, and a file whose values would differ is left as it was. Findings that have no safe repair, such as an implicit octal or a duplicate key, are still reported. Files keep their encoding and byte-order mark.

## Format

```shell
officeimo yaml format deployment.yaml --indent 2 --sequences indented
officeimo yaml format charts --write --indent 2
officeimo yaml format . --check
```

`format` writes the whole file anew in one layout. One file, or standard input, goes to standard output; `--write` rewrites files in place, and `--check` lists the files that would change and returns exit code `1` if there are any, without writing. A folder or several files need `--write` or `--check`.

| Option | Values | Default |
|---|---|---|
| `--indent` | 1 to 9 | The width the file uses |
| `--sequences` | `preserve`, `compact` (`key:\n- item`), `indented` (`key:\n  - item`) | The style the file uses |
| `--quotes` | `preserve`, `single`, `double` | Each string keeps its quotes |
| `--line-ending` | `preserve`, `lf`, `crlf` | The break most lines use |
| `--no-final-newline` | | The file ends with a line break |

Formatting keeps every value, scalar spelling (`TRUE`, `0x1F`, `~`), anchor, alias and tag, every comment outside flow collections, and the blank lines above entries. Strings that would read as another type stay quoted, and `--quotes single` keeps double quotes where a value needs an escape. Flow collections are written on one line, multi-line plain and folded strings are laid out again, and comments inside flow collections are dropped. Formatting a formatted file returns it unchanged.

Use `lint --fix` to repair a file while leaving its layout alone, and `format` to normalize the whole file.

## In CI

```yaml
- name: Check YAML
  run: |
    dotnet tool install --global OfficeIMO.Tool
    officeimo yaml lint . --strict --disable line-length
    officeimo yaml format . --check --indent 2
```
