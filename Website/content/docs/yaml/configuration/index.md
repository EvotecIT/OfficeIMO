---
title: YAML Configuration Files in .NET
description: "Load appsettings.yaml into Microsoft.Extensions.Configuration with AddYamlFile: reload on change, merge keys, aliases and options binding."
order: 55
verified: 2026-11-15
packages:
  - OfficeIMO.Yaml.Configuration
limits: One YAML document per file; a stream with several documents is rejected. Duplicate keys, including keys that differ only in case, are errors.
search_terms:
  - appsettings.yaml
  - AddYamlFile
  - yaml configuration provider
  - IConfiguration yaml
  - NetEscapades.Configuration.Yaml
---

`OfficeIMO.Yaml.Configuration` adds YAML files to `Microsoft.Extensions.Configuration`, so ASP.NET Core, the generic host and plain `ConfigurationBuilder` code read `appsettings.yaml` the way they read `appsettings.json`. It is a separate package because it depends on `Microsoft.Extensions.Configuration.FileExtensions`; `OfficeIMO.Yaml` itself stays dependency-free.

```sh
dotnet add package OfficeIMO.Yaml.Configuration
```

## Add a YAML file

### C#

```csharp
using Microsoft.Extensions.Configuration;

var builder = WebApplication.CreateBuilder(args);
builder.Configuration
    .AddYamlFile("appsettings.yaml", optional: false, reloadOnChange: true)
    .AddYamlFile($"appsettings.{builder.Environment.EnvironmentName}.yaml", optional: true, reloadOnChange: true);
builder.Services.Configure<SmtpOptions>(builder.Configuration.GetSection("Smtp"));
```

`AddYamlFile` has the overloads of `AddJsonFile`: a path, `optional`, `reloadOnChange`, an `IFileProvider`, or an `Action<YamlConfigurationSource>` for every file-source setting. `AddYamlStream(stream)` reads YAML from a stream.

## How keys and values map

```yaml
defaults: &defaults
  Timeout: 30
  Retries: 3
Smtp:
  <<: *defaults
  Host: smtp.example.com
  UseTls: TRUE
  Recipients: [ops@example.com, audit@example.com]
  ReplyTo: ~
```

| Key | Value |
|---|---|
| `Smtp:Host` | `smtp.example.com` |
| `Smtp:Timeout` | `30`, from the merged `defaults` |
| `Smtp:UseTls` | `TRUE`, the file's spelling, which binds to `bool` |
| `Smtp:Recipients:1` | `audit@example.com` |
| `Smtp:ReplyTo` | empty string, as JSON `null` gives |

Keys are case-insensitive and joined with `:`; sequence items use their zero-based index. Aliases and `<<` merge keys resolve as in the document's resolved view, and the `<<` key never appears. Empty collections are keys with a `null` value, as in the JSON provider.

## Errors

Malformed YAML, duplicate keys, a root that is not a mapping and a stream with more than one document throw `FormatException` with the line; a file source wraps it in an `InvalidDataException` naming the file. The [package README](https://github.com/EvotecIT/OfficeIMO/blob/master/OfficeIMO.Yaml.Configuration/README.md) lists every rule and how to migrate from NetEscapades.Configuration.Yaml.
