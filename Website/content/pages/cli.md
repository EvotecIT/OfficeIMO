---
title: "CLI"
description: "Convert, extract, inspect and clean documents from your terminal with the officeimo CLI, in scripts, CI jobs, batches and MCP clients."
meta.seo_title: "OfficeIMO CLI: convert and inspect documents"
layout: application
meta.application: "cli"
---

## One command for document work in scripts

`officeimo` runs the OfficeIMO engines from a shell, a CI job, a scheduled task or an AI agent, without Microsoft Office. Results go to standard output as JSON, diagnostics go to standard error, failures return documented exit codes, and existing files are protected unless you pass `--force`.

## Install

```shell
dotnet tool install --global OfficeIMO.Tool
officeimo --version
officeimo help
```

For scripts and CI, pin the version in the repository with a tool manifest and run it as `dotnet officeimo`. [Install and pin the tool](/docs/cli/install/) shows both.

Every command also prints its own reference with `officeimo <area> --help`.
