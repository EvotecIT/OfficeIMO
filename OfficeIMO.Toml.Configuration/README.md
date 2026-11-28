# OfficeIMO.Toml.Configuration

A TOML configuration provider for `Microsoft.Extensions.Configuration`, built on the [OfficeIMO.Toml](https://github.com/EvotecIT/OfficeIMO/blob/master/OfficeIMO.Toml/README.md) engine. It reads TOML 1.1, targets .NET Standard 2.0, .NET Framework 4.7.2, .NET 8 and .NET 10, works in trimmed and NativeAOT applications, and depends only on `OfficeIMO.Toml` and `Microsoft.Extensions.Configuration.FileExtensions`.

## Install

```powershell
dotnet add package OfficeIMO.Toml.Configuration
```

## Usage

```csharp
using Microsoft.Extensions.Configuration;

IConfiguration configuration = new ConfigurationBuilder()
    .SetBasePath(AppContext.BaseDirectory)
    .AddTomlFile("appsettings.toml", optional: false, reloadOnChange: true)
    .AddTomlFile($"appsettings.{environment}.toml", optional: true)
    .Build();

int port = int.Parse(configuration["Smtp:Port"]!);
```

In ASP.NET Core and the generic host, add the file to the host's builder and bind sections as usual:

```csharp
var builder = WebApplication.CreateBuilder(args);
builder.Configuration.AddTomlFile("appsettings.toml", optional: true, reloadOnChange: true);
builder.Services.Configure<SmtpOptions>(builder.Configuration.GetSection("Smtp"));
```

```toml
[Logging.LogLevel]
Default = "Information"
"Microsoft.AspNetCore" = "Warning"

[Smtp]
Host = "smtp.example.com"
Port = 587
Recipients = ["ops@example.com", "audit@example.com"]

[[Smtp.Relays]]
Host = "relay-1.example.com"

[[Smtp.Relays]]
Host = "relay-2.example.com"
```

The methods match the JSON provider's:

| Method | Reads |
| --- | --- |
| `AddTomlFile(path)`, `AddTomlFile(path, optional)`, `AddTomlFile(path, optional, reloadOnChange)` | A file relative to the builder's base path |
| `AddTomlFile(IFileProvider?, path, optional, reloadOnChange)` | A file from a file provider |
| `AddTomlFile(Action<TomlConfigurationSource>)` | A file, with every `FileConfigurationSource` setting (`ReloadDelay`, `OnLoadException`, ...) |
| `AddTomlStream(stream)` | A stream, read once when the configuration is built |

`TomlConfigurationSource`, `TomlConfigurationProvider`, `TomlStreamConfigurationSource` and `TomlStreamConfigurationProvider` are in the `OfficeIMO.Toml.Configuration` namespace; the extension methods are in `Microsoft.Extensions.Configuration`.

## How TOML becomes configuration

- **Keys** are flattened with `:`, and the items of arrays and arrays of tables use their zero-based index: `Smtp:Recipients:1` and `Smtp:Relays:0:Host`. Keys are case-insensitive.
- **Tables** flatten the same way however the file writes them: `[a.b]` headers, dotted keys (`a.b.c = 1`) and inline tables (`a = { b = { c = 1 } }`) all give `a:b:c`.
- **Strings** keep their decoded text, from any of the four quoting styles.
- **Integers** keep their decimal spelling; hexadecimal, octal, binary, `+`-signed and `_`-grouped integers are written as invariant decimals (`0xFF` is `255`, `1_000` is `1000`), so they bind to numeric types.
- **Floats** keep their spelling (`3.0`, `6.626e-34`) unless it has `_` or a leading `+`, which are written as invariant round-trip text. `inf`, `-inf` and `nan` become `Infinity`, `-Infinity` and `NaN`, which `double` parses.
- **Booleans** are `true` and `false`.
- **Date-times** are ISO 8601 with a `T`, whatever the file's separator: `1979-05-27T07:32:00.5-07:00` binds to `DateTimeOffset` and `DateTime`, a local date `1979-05-27` to `DateTime`, and a local time `07:32:00` to `TimeSpan`.
- **Empty arrays, tables and array-of-tables items** are keys with a `null` value, as in the JSON provider.
- A file that is empty or holds only comments adds no keys.

## Errors

The provider throws `FormatException` for input it cannot map; a file provider reports it as an `InvalidDataException` naming the file, with the `FormatException` as its inner exception, as the JSON provider does.

- Malformed TOML and duplicate keys (which TOML forbids): the message gives the line and column, and the inner `TomlException` carries `Line`, `Column` and a stable `Code` such as `Toml.Load.DuplicateKey`.
- Keys that differ only in case, which TOML allows but configuration would merge: `A duplicate key 'server:port' was found at line 3.`

Files are loaded with the engine's default limits (nesting depth 256, 1,000,000 nodes, 16 MiB strings), which do not reject normal configuration.

## Migrating from Tomlyn.Extensions.Configuration

1. Replace the package: `dotnet remove package Tomlyn.Extensions.Configuration`, then `dotnet add package OfficeIMO.Toml.Configuration`.
2. Calls to `AddTomlFile` keep compiling: the overloads follow the JSON provider and live in the `Microsoft.Extensions.Configuration` namespace. Code that names the source or provider types directly switches its `using` to `OfficeIMO.Toml.Configuration`.
3. Values bind the same way to typed options. Code that compares raw configuration strings should expect the spellings above, such as `true` for booleans and ISO 8601 for date-times.
4. The provider reads TOML 1.1: newlines and trailing commas in inline tables, `\e` and `\xHH` escapes, and times without seconds are accepted.

## NativeAOT and trimming

The provider uses no reflection. `OfficeIMO.Toml.AotSmoke` reads configuration through it in a NativeAOT executable. Binding to options classes goes through `Microsoft.Extensions.Configuration.Binder`, whose source generator (`EnableConfigurationBindingGenerator`) covers trimmed and NativeAOT applications.
