using OfficeIMO.Excel;
using OfficeIMO.Markdown;
using OfficeIMO.Pdf;
using OfficeIMO.PowerPoint;
using OfficeIMO.Word;
using OfficeIMO.Word.Markdown;
using DocumentFormat.OpenXml.Packaging;
using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Runtime.ExceptionServices;
using System.Security.Cryptography;
using System.Text;
using System.Text.Json;
using System.Threading;

namespace OfficeIMO.Reader;

/// <summary>
/// Unified, read-only document extraction facade intended for AI ingestion.
/// </summary>
/// <remarks>
/// This facade is intentionally dependency-free and deterministic.
/// It normalizes extraction into <see cref="ReaderChunk"/> instances with stable IDs and location metadata.
/// The API is thread-safe as it does not use shared mutable state.
/// </remarks>
public static partial class DocumentReader {
    private static readonly string[] DefaultFolderExtensions = {
        ".docx", ".docm",
        ".xlsx", ".xlsm",
        ".pptx", ".pptm",
        ".md", ".markdown",
        ".pdf",
        ".txt", ".log", ".csv", ".tsv", ".json", ".xml", ".yml", ".yaml"
    };

    private static readonly ReaderHandlerCapability[] BuiltInCapabilities = {
        new ReaderHandlerCapability {
            Id = "officeimo.reader.word",
            DisplayName = "Word Reader",
            Description = "Built-in Word (.docx/.docm) chunk extractor.",
            Kind = ReaderInputKind.Word,
            Extensions = new[] { ".docx", ".docm" },
            IsBuiltIn = true,
            SupportsPath = true,
            SupportsStream = true
        },
        new ReaderHandlerCapability {
            Id = "officeimo.reader.excel",
            DisplayName = "Excel Reader",
            Description = "Built-in Excel (.xlsx/.xlsm) table and markdown extractor.",
            Kind = ReaderInputKind.Excel,
            Extensions = new[] { ".xlsx", ".xlsm" },
            IsBuiltIn = true,
            SupportsPath = true,
            SupportsStream = true
        },
        new ReaderHandlerCapability {
            Id = "officeimo.reader.powerpoint",
            DisplayName = "PowerPoint Reader",
            Description = "Built-in PowerPoint (.pptx/.pptm) slide extractor.",
            Kind = ReaderInputKind.PowerPoint,
            Extensions = new[] { ".pptx", ".pptm" },
            IsBuiltIn = true,
            SupportsPath = true,
            SupportsStream = true
        },
        new ReaderHandlerCapability {
            Id = "officeimo.reader.markdown",
            DisplayName = "Markdown Reader",
            Description = "Built-in Markdown chunk extractor.",
            Kind = ReaderInputKind.Markdown,
            Extensions = new[] { ".md", ".markdown" },
            IsBuiltIn = true,
            SupportsPath = true,
            SupportsStream = true
        },
        new ReaderHandlerCapability {
            Id = "officeimo.reader.pdf",
            DisplayName = "PDF Reader",
            Description = "Built-in PDF logical page and markdown extractor.",
            Kind = ReaderInputKind.Pdf,
            Extensions = new[] { ".pdf" },
            IsBuiltIn = true,
            SupportsPath = true,
            SupportsStream = true
        },
        new ReaderHandlerCapability {
            Id = "officeimo.reader.text",
            DisplayName = "Text Reader",
            Description = "Built-in plain text reader for text-like formats.",
            Kind = ReaderInputKind.Text,
            Extensions = new[] { ".txt", ".log", ".csv", ".tsv", ".json", ".xml", ".yml", ".yaml" },
            IsBuiltIn = true,
            SupportsPath = true,
            SupportsStream = true
        }
    };

    private static readonly HashSet<string> BuiltInExtensions = BuildBuiltInExtensionSet();
    private static readonly object HandlerRegistrySync = new object();
    private static readonly Dictionary<string, CustomReaderHandler> CustomHandlersById = new Dictionary<string, CustomReaderHandler>(StringComparer.OrdinalIgnoreCase);
    private static readonly Dictionary<string, string> CustomHandlerIdByExtension = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);

    private static string? TryGetExtension(string path) {
        if (path == null) return null;
        try {
            return Path.GetExtension(path);
        } catch (ArgumentException) {
            return null;
        } catch (NotSupportedException) {
            return null;
        }
    }

    /// <summary>
    /// Registers a custom handler for one or more file extensions.
    /// </summary>
    /// <param name="registration">Custom handler registration.</param>
    /// <param name="replaceExisting">
    /// When true, removes conflicting custom handlers and allows built-in extension overrides.
    /// </param>
    public static void RegisterHandler(ReaderHandlerRegistration registration, bool replaceExisting = false) {
        if (registration == null) throw new ArgumentNullException(nameof(registration));

        var id = (registration.Id ?? string.Empty).Trim();
        if (id.Length == 0) throw new ArgumentException("Handler Id cannot be empty.", nameof(registration));

        if (registration.ReadPath == null && registration.ReadStream == null) {
            throw new ArgumentException("Handler must define ReadPath and/or ReadStream.", nameof(registration));
        }
        if (registration.DefaultMaxInputBytes.HasValue && registration.DefaultMaxInputBytes.Value < 1) {
            throw new ArgumentException("DefaultMaxInputBytes must be greater than 0 when specified.", nameof(registration));
        }

        var normalizedExtensions = NormalizeRegistrationExtensions(registration.Extensions);
        if (normalizedExtensions.Count == 0) {
            throw new ArgumentException("Handler must define at least one extension.", nameof(registration));
        }

        lock (HandlerRegistrySync) {
            if (!replaceExisting) {
                if (CustomHandlersById.ContainsKey(id)) {
                    throw new InvalidOperationException($"Handler '{id}' is already registered.");
                }

                foreach (var ext in normalizedExtensions) {
                    if (BuiltInExtensions.Contains(ext)) {
                        throw new InvalidOperationException($"Extension '{ext}' is handled by a built-in reader. Use replaceExisting=true to override.");
                    }
                    if (CustomHandlerIdByExtension.ContainsKey(ext)) {
                        throw new InvalidOperationException($"Extension '{ext}' is already handled by a custom reader.");
                    }
                }
            } else {
                var toRemove = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
                if (CustomHandlersById.ContainsKey(id)) {
                    toRemove.Add(id);
                }
                foreach (var ext in normalizedExtensions) {
                    if (CustomHandlerIdByExtension.TryGetValue(ext, out var existing)) {
                        toRemove.Add(existing);
                    }
                }
                foreach (var existingId in toRemove) {
                    RemoveCustomHandlerUnsafe(existingId);
                }
            }

            var custom = new CustomReaderHandler(
                id: id,
                displayName: string.IsNullOrWhiteSpace(registration.DisplayName) ? id : registration.DisplayName!.Trim(),
                description: registration.Description,
                kind: registration.Kind,
                extensions: normalizedExtensions.ToArray(),
                defaultMaxInputBytes: registration.DefaultMaxInputBytes,
                warningBehavior: registration.WarningBehavior,
                deterministicOutput: registration.DeterministicOutput,
                readPath: registration.ReadPath,
                readStream: registration.ReadStream);

            CustomHandlersById[id] = custom;
            foreach (var ext in custom.Extensions) {
                CustomHandlerIdByExtension[ext] = custom.Id;
            }
        }
    }

    /// <summary>
    /// Unregisters a custom handler by identifier.
    /// </summary>
    public static bool UnregisterHandler(string handlerId) {
        if (string.IsNullOrWhiteSpace(handlerId)) return false;
        lock (HandlerRegistrySync) {
            return RemoveCustomHandlerUnsafe(handlerId.Trim());
        }
    }

    /// <summary>
    /// Lists built-in and custom reader capabilities for host discovery.
    /// </summary>
    public static IReadOnlyList<ReaderHandlerCapability> GetCapabilities(bool includeBuiltIn = true, bool includeCustom = true) {
        var list = new List<ReaderHandlerCapability>();

        if (includeBuiltIn) {
            list.AddRange(BuiltInCapabilities.Select(CloneCapability));
        }

        if (includeCustom) {
            lock (HandlerRegistrySync) {
                foreach (var custom in CustomHandlersById.Values.OrderBy(static c => c.Id, StringComparer.Ordinal)) {
                    list.Add(custom.ToCapability());
                }
            }
        }

        return list
            .OrderBy(static c => c.IsBuiltIn ? 0 : 1)
            .ThenBy(static c => c.Id, StringComparer.Ordinal)
            .ToArray();
    }

    /// <summary>
    /// Builds a machine-readable capability manifest for host auto-discovery.
    /// </summary>
    public static ReaderCapabilityManifest GetCapabilityManifest(bool includeBuiltIn = true, bool includeCustom = true) {
        var handlers = GetCapabilities(includeBuiltIn, includeCustom)
            .Select(CloneCapability)
            .ToArray();

        return new ReaderCapabilityManifest {
            SchemaId = ReaderCapabilitySchema.Id,
            SchemaVersion = ReaderCapabilitySchema.Version,
            Handlers = handlers
        };
    }

    /// <summary>
    /// Builds a JSON capability manifest payload for host auto-discovery.
    /// </summary>
    public static string GetCapabilityManifestJson(bool includeBuiltIn = true, bool includeCustom = true, bool indented = false) {
        var manifest = GetCapabilityManifest(includeBuiltIn, includeCustom);
        return ReaderCapabilityManifestJson.Serialize(manifest, indented);
    }

    /// <summary>
    /// Discovers modular registrar methods in the provided assemblies.
    /// </summary>
    public static IReadOnlyList<ReaderHandlerRegistrarDescriptor> DiscoverHandlerRegistrars(IEnumerable<Assembly> assemblies) {
        var candidates = DiscoverHandlerRegistrarsCore(assemblies);
        return candidates
            .Select(static c => CloneRegistrarDescriptor(c.Descriptor))
            .ToArray();
    }

    /// <summary>
    /// Discovers modular registrar methods in the provided assemblies.
    /// </summary>
    public static IReadOnlyList<ReaderHandlerRegistrarDescriptor> DiscoverHandlerRegistrars(params Assembly[] assemblies) {
        return DiscoverHandlerRegistrars((IEnumerable<Assembly>)assemblies);
    }

    /// <summary>
    /// Discovers modular registrar methods from currently loaded assemblies
    /// whose simple name starts with <paramref name="assemblyNamePrefix"/>.
    /// </summary>
    /// <param name="assemblyNamePrefix">
    /// Simple assembly-name prefix filter. Default: <c>OfficeIMO.Reader.</c>.
    /// </param>
    public static IReadOnlyList<ReaderHandlerRegistrarDescriptor> DiscoverHandlerRegistrarsFromLoadedAssemblies(string assemblyNamePrefix = "OfficeIMO.Reader.") {
        var assemblies = GetLoadedAssembliesByPrefix(assemblyNamePrefix);
        return DiscoverHandlerRegistrars(assemblies);
    }

    /// <summary>
    /// Registers modular handlers discovered in the provided assemblies.
    /// </summary>
    /// <param name="assemblies">Assemblies to scan for registrar methods.</param>
    /// <param name="replaceExisting">
    /// Passed to discovered registrar methods via their <c>replaceExisting</c> parameter when present.
    /// </param>
    public static IReadOnlyList<ReaderHandlerRegistrarDescriptor> RegisterHandlersFromAssemblies(IEnumerable<Assembly> assemblies, bool replaceExisting = true) {
        var candidates = DiscoverHandlerRegistrarsCore(assemblies);
        var registered = new List<ReaderHandlerRegistrarDescriptor>(candidates.Count);

        foreach (var candidate in candidates) {
            var parameters = candidate.Method.GetParameters();
            var args = new object?[parameters.Length];
            for (int i = 0; i < parameters.Length; i++) {
                var parameter = parameters[i];
                if (parameter.ParameterType == typeof(bool) &&
                    string.Equals(parameter.Name, "replaceExisting", StringComparison.OrdinalIgnoreCase)) {
                    args[i] = replaceExisting;
                } else if (parameter.IsOptional) {
                    args[i] = Type.Missing;
                } else {
                    throw new InvalidOperationException(
                        $"Registrar method '{candidate.Method.DeclaringType?.FullName}.{candidate.Method.Name}' has unsupported non-optional parameter '{parameter.Name}'.");
                }
            }

            try {
                candidate.Method.Invoke(obj: null, parameters: args);
            } catch (TargetInvocationException ex) when (ex.InnerException != null) {
                ExceptionDispatchInfo.Capture(ex.InnerException).Throw();
                throw;
            }

            registered.Add(CloneRegistrarDescriptor(candidate.Descriptor));
        }

        return registered.ToArray();
    }

    /// <summary>
    /// Registers modular handlers discovered in the provided assemblies.
    /// </summary>
    public static IReadOnlyList<ReaderHandlerRegistrarDescriptor> RegisterHandlersFromAssemblies(bool replaceExisting = true, params Assembly[] assemblies) {
        return RegisterHandlersFromAssemblies((IEnumerable<Assembly>)assemblies, replaceExisting);
    }

    /// <summary>
    /// Registers modular handlers discovered from currently loaded assemblies
    /// whose simple name starts with <paramref name="assemblyNamePrefix"/>.
    /// </summary>
    /// <param name="replaceExisting">
    /// Passed to discovered registrar methods via their <c>replaceExisting</c> parameter when present.
    /// </param>
    /// <param name="assemblyNamePrefix">
    /// Simple assembly-name prefix filter. Default: <c>OfficeIMO.Reader.</c>.
    /// </param>
    public static IReadOnlyList<ReaderHandlerRegistrarDescriptor> RegisterHandlersFromLoadedAssemblies(bool replaceExisting = true, string assemblyNamePrefix = "OfficeIMO.Reader.") {
        var assemblies = GetLoadedAssembliesByPrefix(assemblyNamePrefix);
        return RegisterHandlersFromAssemblies(assemblies, replaceExisting);
    }

    /// <summary>
    /// Host bootstrap helper that registers modular handlers from the provided assemblies
    /// and returns both typed and JSON capability manifests in one payload.
    /// </summary>
    /// <param name="assemblies">Assemblies to scan for registrar methods.</param>
    /// <param name="options">Bootstrap options. When null, defaults are used.</param>
    public static ReaderHostBootstrapResult BootstrapHostFromAssemblies(IEnumerable<Assembly> assemblies, ReaderHostBootstrapOptions? options = null) {
        var normalizedOptions = NormalizeHostBootstrapOptions(options);
        var registered = RegisterHandlersFromAssemblies(assemblies, replaceExisting: normalizedOptions.ReplaceExistingHandlers);
        var manifest = GetCapabilityManifest(
            includeBuiltIn: normalizedOptions.IncludeBuiltInCapabilities,
            includeCustom: normalizedOptions.IncludeCustomCapabilities);

        return new ReaderHostBootstrapResult {
            ReplaceExistingHandlers = normalizedOptions.ReplaceExistingHandlers,
            RegisteredHandlers = registered
                .Select(static r => CloneRegistrarDescriptor(r))
                .ToArray(),
            Manifest = manifest,
            ManifestJson = ReaderCapabilityManifestJson.Serialize(manifest, normalizedOptions.IndentedManifestJson)
        };
    }

    /// <summary>
    /// Host bootstrap helper that applies a preset profile, registers modular handlers from the provided
    /// assemblies, and returns both typed and JSON capability manifests in one payload.
    /// </summary>
    /// <param name="assemblies">Assemblies to scan for registrar methods.</param>
    /// <param name="profile">Bootstrap profile preset.</param>
    /// <param name="indentedManifestJson">When true, indents the returned manifest JSON payload.</param>
    public static ReaderHostBootstrapResult BootstrapHostFromAssemblies(
        IEnumerable<Assembly> assemblies,
        ReaderHostBootstrapProfile profile,
        bool indentedManifestJson = false) {
        var options = CreateHostBootstrapOptions(profile, indentedManifestJson);
        var result = BootstrapHostFromAssemblies(assemblies, options);
        result.Profile = profile;
        return result;
    }

    /// <summary>
    /// Host bootstrap helper that discovers and registers modular handlers from currently loaded assemblies
    /// whose simple name starts with <paramref name="assemblyNamePrefix"/>, then returns both typed and
    /// JSON capability manifests in one payload.
    /// </summary>
    /// <param name="assemblyNamePrefix">
    /// Simple assembly-name prefix filter. Default: <c>OfficeIMO.Reader.</c>.
    /// </param>
    /// <param name="options">Bootstrap options. When null, defaults are used.</param>
    public static ReaderHostBootstrapResult BootstrapHostFromLoadedAssemblies(
        string assemblyNamePrefix = "OfficeIMO.Reader.",
        ReaderHostBootstrapOptions? options = null) {
        if (string.IsNullOrWhiteSpace(assemblyNamePrefix)) {
            throw new ArgumentException("Assembly name prefix cannot be empty.", nameof(assemblyNamePrefix));
        }

        var normalizedOptions = NormalizeHostBootstrapOptions(options);
        var registered = RegisterHandlersFromLoadedAssemblies(
            replaceExisting: normalizedOptions.ReplaceExistingHandlers,
            assemblyNamePrefix: assemblyNamePrefix);
        var manifest = GetCapabilityManifest(
            includeBuiltIn: normalizedOptions.IncludeBuiltInCapabilities,
            includeCustom: normalizedOptions.IncludeCustomCapabilities);

        return new ReaderHostBootstrapResult {
            AssemblyNamePrefix = assemblyNamePrefix.Trim(),
            ReplaceExistingHandlers = normalizedOptions.ReplaceExistingHandlers,
            RegisteredHandlers = registered
                .Select(static r => CloneRegistrarDescriptor(r))
                .ToArray(),
            Manifest = manifest,
            ManifestJson = ReaderCapabilityManifestJson.Serialize(manifest, normalizedOptions.IndentedManifestJson)
        };
    }

    /// <summary>
    /// Host bootstrap helper that applies a preset profile, discovers and registers modular handlers from loaded
    /// assemblies, and returns both typed and JSON capability manifests in one payload.
    /// </summary>
    /// <param name="profile">Bootstrap profile preset.</param>
    /// <param name="assemblyNamePrefix">
    /// Simple assembly-name prefix filter. Default: <c>OfficeIMO.Reader.</c>.
    /// </param>
    /// <param name="indentedManifestJson">When true, indents the returned manifest JSON payload.</param>
    public static ReaderHostBootstrapResult BootstrapHostFromLoadedAssemblies(
        ReaderHostBootstrapProfile profile,
        string assemblyNamePrefix = "OfficeIMO.Reader.",
        bool indentedManifestJson = false) {
        var options = CreateHostBootstrapOptions(profile, indentedManifestJson);
        var result = BootstrapHostFromLoadedAssemblies(assemblyNamePrefix, options);
        result.Profile = profile;
        return result;
    }

    /// <summary>
    /// Detects the input kind based on file extension.
    /// </summary>
    /// <param name="path">Source file path.</param>
    public static ReaderInputKind DetectKind(string path) {
        if (path == null) throw new ArgumentNullException(nameof(path));
        if (path.Length == 0) throw new ArgumentException("Path cannot be empty.", nameof(path));

        var extLower = NormalizeExtension(TryGetExtension(path));
        if (extLower.Length > 0 && TryResolveCustomHandlerByExtension(extLower, out var custom)) {
            return custom.Kind;
        }
        if (extLower.Length == 0) return ReaderInputKind.Unknown;
        return extLower switch {
            ".docx" or ".docm" => ReaderInputKind.Word,
            ".xlsx" or ".xlsm" => ReaderInputKind.Excel,
            ".pptx" or ".pptm" => ReaderInputKind.PowerPoint,
            ".md" or ".markdown" => ReaderInputKind.Markdown,
            ".pdf" => ReaderInputKind.Pdf,
            ".txt" or ".log" or ".csv" or ".tsv" or ".json" or ".xml" or ".yml" or ".yaml" => ReaderInputKind.Text,
            ".doc" or ".xls" or ".ppt" => ReaderInputKind.Unknown, // Legacy binary formats are not supported.
            _ => ReaderInputKind.Unknown
        };
    }

}
