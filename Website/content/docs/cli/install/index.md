---
title: "Install and Pin the OfficeIMO CLI"
description: "Install officeimo as a global .NET tool or pin it per repository with a tool manifest, so scripts and CI jobs run the same version."
layout: docs
---

`OfficeIMO.Tool` is a .NET tool on NuGet. Installing it needs the .NET SDK; Microsoft Office is not required. The command it installs is `officeimo`.

## Install globally

Use a global installation when `officeimo` should be available from any directory on your machine:

```shell
dotnet tool install --global OfficeIMO.Tool
officeimo --version
officeimo help
```

Update or remove it with the standard .NET tool commands:

```shell
dotnet tool update --global OfficeIMO.Tool
dotnet tool uninstall --global OfficeIMO.Tool
```

## Pin it to a repository

For scripts and CI, I recommend a local tool manifest. The manifest records the exact package version, so every clone and build agent runs the same `officeimo`:

```shell
dotnet new tool-manifest
dotnet tool install OfficeIMO.Tool
```

Commit `.config/dotnet-tools.json`. On a fresh clone or build agent, restore the pinned version and run it through the SDK:

```shell
dotnet tool restore
dotnet tool run officeimo --version
dotnet officeimo help
```

`dotnet officeimo` is the SDK shorthand for `dotnet tool run officeimo`. Move the pin forward deliberately with `dotnet tool update OfficeIMO.Tool`, and commit the changed manifest with the scripts that depend on it.

A GitHub Actions step looks like this:

```yaml
- uses: actions/setup-dotnet@v4
  with:
    dotnet-version: '10.0.x'
- run: dotnet tool restore
- run: dotnet officeimo convert docs/report.docx artifacts/report.pdf
```

## Check what your version supports

The commands print their contracts from the installed version, which is the source to trust when scripts move between versions:

```shell
officeimo --version
officeimo convert --help
officeimo reader capabilities --format json
officeimo html capabilities --format json
officeimo provenance capabilities
```

`officeimo convert --help` ends with the list of conversion routes the installed tool runs. `reader capabilities` lists the extraction handlers and their extensions.

## Optional components

The tool does not include an OCR engine or model. OCR commands load a provider assembly you supply with `--ocr-provider-assembly`; see [PDF workflows](/docs/cli/pdf/#searchable-scans-with-ocr). Invoice standards validation needs pinned rule artifacts and a Saxon runner you provide; see [Invoices](/docs/cli/invoices/#standards-validation). Printing on macOS and Linux uses the system CUPS tools.

## Next steps

- [Convert documents](/docs/cli/convert/)
- [Scripting and CI](/docs/cli/scripting-and-ci/)
- [MCP and agents](/docs/cli/mcp-and-agents/)
