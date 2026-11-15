---
title: TOML Configuration Files in .NET
description: "Load appsettings.toml into Microsoft.Extensions.Configuration with AddTomlFile: reload on change, arrays of tables and ISO 8601 date-times."
order: 60
verified: 2026-11-15
packages:
  - OfficeIMO.Toml.Configuration
limits: Keys that differ only in case are rejected, because configuration keys are case-insensitive while TOML keys are not.
search_terms:
  - appsettings.toml
  - AddTomlFile
  - toml configuration provider
  - IConfiguration toml
  - Tomlyn.Extensions.Configuration
---

`OfficeIMO.Toml.Configuration` adds TOML 1.1 files to `Microsoft.Extensions.Configuration`, so ASP.NET Core, the generic host and plain `ConfigurationBuilder` code read `appsettings.toml` the way they read `appsettings.json`. It is a separate package because it depends on `Microsoft.Extensions.Configuration.FileExtensions`; `OfficeIMO.Toml` itself stays dependency-free.

```sh
dotnet add package OfficeIMO.Toml.Configuration
```

## Add a TOML file

### C#

```csharp
using Microsoft.Extensions.Configuration;

var builder = WebApplication.CreateBuilder(args);
builder.Configuration
    .AddTomlFile("appsettings.toml", optional: false, reloadOnChange: true)
    .AddTomlFile($"appsettings.{builder.Environment.EnvironmentName}.toml", optional: true, reloadOnChange: true);
builder.Services.Configure<SmtpOptions>(builder.Configuration.GetSection("Smtp"));
```

`AddTomlFile` has the overloads of `AddJsonFile`: a path, `optional`, `reloadOnChange`, an `IFileProvider`, or an `Action<TomlConfigurationSource>` for every file-source setting. `AddTomlStream(stream)` reads TOML from a stream.

## How keys and values map

```toml
[Smtp]
Host = "smtp.example.com"
Port = 0x24B
Retry = { Count = 3, Delay = "00:00:05" }
Since = 2024-05-01 08:30:00Z

[[Smtp.Relays]]
Host = "relay-1.example.com"
```

| Key | Value |
|---|---|
| `Smtp:Host` | `smtp.example.com` |
| `Smtp:Port` | `587`: hexadecimal, octal, binary and `_`-grouped integers become decimals |
| `Smtp:Retry:Count` | `3`: inline tables and dotted keys flatten like `[tables]` |
| `Smtp:Since` | `2024-05-01T08:30:00Z`, ISO 8601, which binds to `DateTimeOffset` |
| `Smtp:Relays:0:Host` | `relay-1.example.com`: arrays of tables use indexes |

Keys are case-insensitive and joined with `:`. Floats keep their spelling, and `inf` and `nan` become `Infinity` and `NaN`. Empty arrays and tables are keys with a `null` value, as in the JSON provider.

## Errors

Malformed TOML and duplicate keys throw `FormatException` with the line and column; keys that differ only in case are rejected with their line. A file source wraps the exception in an `InvalidDataException` naming the file. The [package README](https://github.com/EvotecIT/OfficeIMO/blob/master/OfficeIMO.Toml.Configuration/README.md) lists every rule and how to migrate from Tomlyn.Extensions.Configuration.
