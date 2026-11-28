# OfficeIMO.Yaml.Configuration

A YAML configuration provider for `Microsoft.Extensions.Configuration`, built on the [OfficeIMO.Yaml](https://github.com/EvotecIT/OfficeIMO/blob/master/OfficeIMO.Yaml/README.md) engine. It targets .NET Standard 2.0, .NET Framework 4.7.2, .NET 8 and .NET 10, works in trimmed and NativeAOT applications, and depends only on `OfficeIMO.Yaml` and `Microsoft.Extensions.Configuration.FileExtensions`.

## Install

```powershell
dotnet add package OfficeIMO.Yaml.Configuration
```

## Usage

```csharp
using Microsoft.Extensions.Configuration;

IConfiguration configuration = new ConfigurationBuilder()
    .SetBasePath(AppContext.BaseDirectory)
    .AddYamlFile("appsettings.yaml", optional: false, reloadOnChange: true)
    .AddYamlFile($"appsettings.{environment}.yaml", optional: true)
    .Build();

string? level = configuration["Logging:LogLevel:Default"];
```

In ASP.NET Core and the generic host, add the file to the host's builder and bind sections as usual:

```csharp
var builder = WebApplication.CreateBuilder(args);
builder.Configuration.AddYamlFile("appsettings.yaml", optional: true, reloadOnChange: true);
builder.Services.Configure<SmtpOptions>(builder.Configuration.GetSection("Smtp"));
```

```yaml
Logging:
  LogLevel:
    Default: Information
    Microsoft.AspNetCore: Warning
Smtp:
  Host: smtp.example.com
  Port: 587
  Recipients:
    - ops@example.com
    - audit@example.com
```

The methods match the JSON provider's:

| Method | Reads |
| --- | --- |
| `AddYamlFile(path)`, `AddYamlFile(path, optional)`, `AddYamlFile(path, optional, reloadOnChange)` | A file relative to the builder's base path |
| `AddYamlFile(IFileProvider?, path, optional, reloadOnChange)` | A file from a file provider |
| `AddYamlFile(Action<YamlConfigurationSource>)` | A file, with every `FileConfigurationSource` setting (`ReloadDelay`, `OnLoadException`, ...) |
| `AddYamlStream(stream)` | A stream, read once when the configuration is built |

`YamlConfigurationSource`, `YamlConfigurationProvider`, `YamlStreamConfigurationSource` and `YamlStreamConfigurationProvider` are in the `OfficeIMO.Yaml.Configuration` namespace; the extension methods are in `Microsoft.Extensions.Configuration`.

## How YAML becomes configuration

- **Keys** are flattened with `:`, and sequence items use their zero-based index, as the JSON provider does: `Smtp:Recipients:1` is `audit@example.com`. Keys are case-insensitive.
- **Scalars** keep their decoded text, whatever their quoting or block style. A file's `TRUE` stays `TRUE` and `0x1F` stays `0x1F`; the binder reads `TRUE` as a `bool`. Timestamps such as `2024-05-01T08:30:00Z` bind to `DateTime` and `DateTimeOffset`.
- **Nulls** (`null`, `~` or no value) become empty strings, as JSON `null` does. A quoted `"null"` is the text `null`.
- **Empty collections** (`[]`, `{}`) are keys with a `null` value, as in the JSON provider.
- **Aliases** read the anchored node, and **merge keys** (`<<: *defaults`, `<<: [*a, *b]`) apply as in the document's resolved view: a mapping's own keys win, then the merged mappings in order. The `<<` key itself never appears. Merge keys work under the default Core schema and in `%YAML 1.1` documents.
- **One document per file.** A stream with several documents (`---` between them) is rejected with a `FormatException` that gives the line of the second document, because configuration has one set of keys and silently dropping the rest would hide settings. A single document may still start with `---` or end with `...`. A file that is empty or holds only comments adds no keys.
- **The root** must be a mapping. A root sequence or scalar is a `FormatException`.

## Errors

The provider throws `FormatException` for input it cannot map; a file provider reports it as an `InvalidDataException` naming the file, with the `FormatException` as its inner exception, as the JSON provider does.

- Malformed YAML: the message gives the line and column, and the inner `YamlException` carries `Line`, `Column` and a stable `Code`.
- Duplicate keys, including keys that differ only in case and paths such as `a:b` written both as a key and as nested mappings: `A duplicate key 'Server:Port' was found at line 3.`
- Collection keys (`[a, b]: x`) and merge keys whose value is not a mapping, with their line.

Files are loaded with the engine's default limits (nesting depth 256, 100,000 alias-expanded nodes per document), which bound alias expansion without rejecting normal configuration.

## Migrating from NetEscapades.Configuration.Yaml

1. Replace the package: `dotnet remove package NetEscapades.Configuration.Yaml`, then `dotnet add package OfficeIMO.Yaml.Configuration`.
2. Calls to `AddYamlFile` keep compiling: the overloads and the `Microsoft.Extensions.Configuration` namespace are the same. Code that names `YamlConfigurationSource` or `YamlConfigurationProvider` directly changes `using NetEscapades.Configuration.Yaml;` to `using OfficeIMO.Yaml.Configuration;`.
3. Check the files against the rules above:
   - YAML nulls are empty strings here, matching the JSON provider. Code that tested a value for `null` to detect `key: ~` sees `""`.
   - A file with several documents is an error, rather than only its first document being read.
   - `<<` merge keys are resolved instead of appearing as a key named `<<`.
   - Keys that differ only in case in one mapping are an error.

## NativeAOT and trimming

The provider uses no reflection. `OfficeIMO.Yaml.AotSmoke` reads configuration through it in a NativeAOT executable. Binding to options classes goes through `Microsoft.Extensions.Configuration.Binder`, whose source generator (`EnableConfigurationBindingGenerator`) covers trimmed and NativeAOT applications.
