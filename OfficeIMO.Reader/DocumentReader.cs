using OfficeIMO.Excel;
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
public static class DocumentReader {
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
            Description = "Built-in PDF page extractor.",
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

    /// <summary>
    /// Reads a supported document file and emits normalized extraction chunks.
    /// </summary>
    /// <param name="path">Source file path.</param>
    /// <param name="options">Extraction options.</param>
    /// <param name="cancellationToken">Cancellation token.</param>
    public static IEnumerable<ReaderChunk> Read(string path, ReaderOptions? options = null, CancellationToken cancellationToken = default) {
        if (path == null) throw new ArgumentNullException(nameof(path));
        if (Directory.Exists(path)) {
            // Keep Read(file) semantics intact; require explicit folder method for directories.
            throw new IOException($"'{path}' is a directory. Use {nameof(ReadFolder)}(...) to ingest directories.");
        }
        if (!File.Exists(path)) throw new FileNotFoundException($"File '{path}' doesn't exist.", path);

        var opt = NormalizeOptions(options);
        EnforceFileSize(path, opt.MaxInputBytes);
        var source = BuildSourceInfoFromPath(path, opt.ComputeHashes);

        IEnumerable<ReaderChunk> raw;
        if (TryResolveCustomHandlerByPath(path, out var customPathHandler) && customPathHandler.ReadPath != null) {
            raw = customPathHandler.ReadPath(path, opt, cancellationToken);
        } else {
            var kind = DetectKind(path);
            raw = kind switch {
                ReaderInputKind.Word => ReadWord(path, opt, cancellationToken),
                ReaderInputKind.Excel => ReadExcel(path, opt, cancellationToken),
                ReaderInputKind.PowerPoint => ReadPowerPoint(path, opt, cancellationToken),
                ReaderInputKind.Markdown => ReadMarkdown(path, opt, cancellationToken),
                ReaderInputKind.Pdf => ReadPdf(path, opt, cancellationToken),
                ReaderInputKind.Text => ReadText(path, opt, cancellationToken),
                _ => ReadUnknown(path, opt, cancellationToken)
            };
        }

        foreach (var chunk in raw) {
            cancellationToken.ThrowIfCancellationRequested();
            yield return EnrichChunk(chunk, source, opt.ComputeHashes);
        }
    }

    /// <summary>
    /// Enumerates a folder and ingests all supported files (best-effort), emitting warning chunks for skipped files.
    /// </summary>
    /// <param name="folderPath">Folder path.</param>
    /// <param name="folderOptions">Folder enumeration options.</param>
    /// <param name="options">Extraction options.</param>
    /// <param name="cancellationToken">Cancellation token.</param>
    public static IEnumerable<ReaderChunk> ReadFolder(string folderPath, ReaderFolderOptions? folderOptions = null, ReaderOptions? options = null, CancellationToken cancellationToken = default) {
        foreach (var chunk in ReadFolder(folderPath, folderOptions, options, onProgress: null, cancellationToken))
            yield return chunk;
    }

    /// <summary>
    /// Enumerates a folder and ingests all supported files (best-effort), emitting warning chunks for skipped files.
    /// </summary>
    /// <param name="folderPath">Folder path.</param>
    /// <param name="folderOptions">Folder enumeration options.</param>
    /// <param name="options">Extraction options.</param>
    /// <param name="onProgress">Optional progress callback for file-level lifecycle and aggregate counts.</param>
    /// <param name="cancellationToken">Cancellation token.</param>
    public static IEnumerable<ReaderChunk> ReadFolder(
        string folderPath,
        ReaderFolderOptions? folderOptions,
        ReaderOptions? options,
        Action<ReaderProgress>? onProgress,
        CancellationToken cancellationToken = default) {
        if (folderPath == null) throw new ArgumentNullException(nameof(folderPath));
        if (!Directory.Exists(folderPath)) throw new DirectoryNotFoundException($"Folder '{folderPath}' doesn't exist.");

        var fo = NormalizeFolderOptions(folderOptions);
        var opt = NormalizeOptions(options);
        var fileReadOptions = CloneOptions(opt, computeHashes: false);
        var allowedExt = NormalizeExtensions(fo.Extensions);
        long total = 0;
        int warningIndex = 0;
        var state = new FolderIngestState();

        foreach (var file in EnumerateFilesSafeDeterministic(folderPath, fo, cancellationToken)) {
            cancellationToken.ThrowIfCancellationRequested();

            var ext = TryGetExtension(file);
            if (string.IsNullOrEmpty(ext)) continue;
            if (!allowedExt.Contains(ext!)) continue;
            if (state.FilesScanned >= fo.MaxFiles) break;

            state.FilesScanned++;
            var source = BuildSourceInfoFromPath(file, opt.ComputeHashes);
            NotifyProgress(onProgress, ReaderProgressEventKind.FileStarted, state, source, null, fileChunkCount: null);

            string? statWarning = null;
            var length = source.LengthBytes;
            if (!length.HasValue) {
                statWarning = "Skipped file because metadata could not be read.";
            }
            if (statWarning != null) {
                state.FilesSkipped++;
                NotifyProgress(onProgress, ReaderProgressEventKind.FileSkipped, state, source, statWarning, fileChunkCount: null);
                yield return EnrichChunk(BuildFolderWarningChunk(file, warningIndex++, statWarning), source, opt.ComputeHashes);
                continue;
            }

            var lengthValue = length.GetValueOrDefault();
            if (fo.MaxTotalBytes.HasValue) {
                if ((total + lengthValue) > fo.MaxTotalBytes.Value) {
                    state.FilesSkipped++;
                    var limitWarning = $"Stopped folder ingestion after reaching MaxTotalBytes ({fo.MaxTotalBytes.Value.ToString(CultureInfo.InvariantCulture)}).";
                    NotifyProgress(onProgress, ReaderProgressEventKind.FileSkipped, state, source, limitWarning, fileChunkCount: null);
                    yield return EnrichChunk(
                        BuildFolderWarningChunk(
                        file,
                        warningIndex++,
                        limitWarning),
                        source,
                        opt.ComputeHashes);
                    yield break;
                }
            }
            total += lengthValue;

            if (opt.MaxInputBytes.HasValue && lengthValue > opt.MaxInputBytes.Value) {
                // Skip too-large files rather than failing the whole folder.
                state.FilesSkipped++;
                var warning = $"Skipped file because it exceeds MaxInputBytes ({lengthValue.ToString(CultureInfo.InvariantCulture)} > {opt.MaxInputBytes.Value.ToString(CultureInfo.InvariantCulture)}).";
                NotifyProgress(onProgress, ReaderProgressEventKind.FileSkipped, state, source, warning, fileChunkCount: null);
                yield return EnrichChunk(
                    BuildFolderWarningChunk(
                    file,
                    warningIndex++,
                    warning),
                    source,
                    opt.ComputeHashes);
                continue;
            }

            List<ReaderChunk>? fileChunks = null;
            string? readWarning = null;
            try {
                fileChunks = Read(file, fileReadOptions, cancellationToken)
                    .Select(c => EnrichChunk(c, source, opt.ComputeHashes))
                    .ToList();
            } catch (OperationCanceledException) {
                throw;
            } catch (Exception ex) {
                // Keep folder ingestion best-effort; skip files that fail parsing.
                readWarning = $"Skipped file due read error: {ex.GetType().Name}.";
            }
            if (readWarning != null) {
                state.FilesSkipped++;
                NotifyProgress(onProgress, ReaderProgressEventKind.FileSkipped, state, source, readWarning, fileChunkCount: null);
                yield return EnrichChunk(BuildFolderWarningChunk(file, warningIndex++, readWarning), source, opt.ComputeHashes);
                continue;
            }

            foreach (var chunk in fileChunks!) {
                cancellationToken.ThrowIfCancellationRequested();
                yield return chunk;
            }

            state.FilesParsed++;
            state.BytesRead += lengthValue;
            state.ChunksProduced += fileChunks!.Count;
            NotifyProgress(onProgress, ReaderProgressEventKind.FileCompleted, state, source, null, fileChunks.Count);
        }

        NotifyProgress(onProgress, ReaderProgressEventKind.Completed, state, source: null, message: null, fileChunkCount: null);
    }

    /// <summary>
    /// Enumerates a folder and emits one source-level payload per file, ready for direct DB upserts.
    /// </summary>
    /// <param name="folderPath">Folder path.</param>
    /// <param name="folderOptions">Folder enumeration options.</param>
    /// <param name="options">Extraction options.</param>
    /// <param name="onProgress">Optional progress callback for file-level lifecycle and aggregate counts.</param>
    /// <param name="cancellationToken">Cancellation token.</param>
    public static IEnumerable<ReaderSourceDocument> ReadFolderDocuments(
        string folderPath,
        ReaderFolderOptions? folderOptions = null,
        ReaderOptions? options = null,
        Action<ReaderProgress>? onProgress = null,
        CancellationToken cancellationToken = default) {
        if (folderPath == null) throw new ArgumentNullException(nameof(folderPath));
        if (!Directory.Exists(folderPath)) throw new DirectoryNotFoundException($"Folder '{folderPath}' doesn't exist.");

        var fo = NormalizeFolderOptions(folderOptions);
        var opt = NormalizeOptions(options);
        var fileReadOptions = CloneOptions(opt, computeHashes: false);
        var allowedExt = NormalizeExtensions(fo.Extensions);
        long total = 0;
        var state = new FolderIngestState();

        foreach (var file in EnumerateFilesSafeDeterministic(folderPath, fo, cancellationToken)) {
            cancellationToken.ThrowIfCancellationRequested();

            var ext = TryGetExtension(file);
            if (string.IsNullOrEmpty(ext)) continue;
            if (!allowedExt.Contains(ext!)) continue;
            if (state.FilesScanned >= fo.MaxFiles) break;

            state.FilesScanned++;
            var source = BuildSourceInfoFromPath(file, opt.ComputeHashes);
            NotifyProgress(onProgress, ReaderProgressEventKind.FileStarted, state, source, null, fileChunkCount: null);

            var length = source.LengthBytes;
            if (!length.HasValue) {
                var warning = "Skipped file because metadata could not be read.";
                state.FilesSkipped++;
                NotifyProgress(onProgress, ReaderProgressEventKind.FileSkipped, state, source, warning, fileChunkCount: null);
                yield return BuildSourceDocument(source, parsed: false, chunks: null, sourceWarnings: new[] { warning });
                continue;
            }

            var lengthValue = length.GetValueOrDefault();
            if (fo.MaxTotalBytes.HasValue && (total + lengthValue) > fo.MaxTotalBytes.Value) {
                var limitWarning = $"Stopped folder ingestion after reaching MaxTotalBytes ({fo.MaxTotalBytes.Value.ToString(CultureInfo.InvariantCulture)}).";
                state.FilesSkipped++;
                NotifyProgress(onProgress, ReaderProgressEventKind.FileSkipped, state, source, limitWarning, fileChunkCount: null);
                yield return BuildSourceDocument(source, parsed: false, chunks: null, sourceWarnings: new[] { limitWarning });
                yield break;
            }
            total += lengthValue;

            if (opt.MaxInputBytes.HasValue && lengthValue > opt.MaxInputBytes.Value) {
                var warning = $"Skipped file because it exceeds MaxInputBytes ({lengthValue.ToString(CultureInfo.InvariantCulture)} > {opt.MaxInputBytes.Value.ToString(CultureInfo.InvariantCulture)}).";
                state.FilesSkipped++;
                NotifyProgress(onProgress, ReaderProgressEventKind.FileSkipped, state, source, warning, fileChunkCount: null);
                yield return BuildSourceDocument(source, parsed: false, chunks: null, sourceWarnings: new[] { warning });
                continue;
            }

            List<ReaderChunk>? fileChunks = null;
            string? readWarning = null;
            try {
                fileChunks = Read(file, fileReadOptions, cancellationToken)
                    .Select(c => EnrichChunk(c, source, opt.ComputeHashes))
                    .ToList();
            } catch (OperationCanceledException) {
                throw;
            } catch (Exception ex) {
                // Keep folder ingestion best-effort; skip files that fail parsing.
                readWarning = $"Skipped file due read error: {ex.GetType().Name}.";
            }

            if (readWarning != null) {
                state.FilesSkipped++;
                NotifyProgress(onProgress, ReaderProgressEventKind.FileSkipped, state, source, readWarning, fileChunkCount: null);
                yield return BuildSourceDocument(source, parsed: false, chunks: null, sourceWarnings: new[] { readWarning });
                continue;
            }

            state.FilesParsed++;
            state.BytesRead += lengthValue;
            state.ChunksProduced += fileChunks!.Count;
            NotifyProgress(onProgress, ReaderProgressEventKind.FileCompleted, state, source, null, fileChunks.Count);

            yield return BuildSourceDocument(source, parsed: true, chunks: fileChunks, sourceWarnings: null);
        }

        NotifyProgress(onProgress, ReaderProgressEventKind.Completed, state, source: null, message: null, fileChunkCount: null);
    }

    /// <summary>
    /// Reads a folder and returns ingestion-ready summary/counts with optional chunk materialization.
    /// </summary>
    /// <param name="folderPath">Folder path.</param>
    /// <param name="folderOptions">Folder enumeration options.</param>
    /// <param name="options">Extraction options.</param>
    /// <param name="includeChunks">When true, materializes chunks in the result object.</param>
    /// <param name="onProgress">Optional progress callback.</param>
    /// <param name="cancellationToken">Cancellation token.</param>
    public static ReaderIngestResult ReadFolderDetailed(
        string folderPath,
        ReaderFolderOptions? folderOptions = null,
        ReaderOptions? options = null,
        bool includeChunks = true,
        Action<ReaderProgress>? onProgress = null,
        CancellationToken cancellationToken = default) {
        var chunks = includeChunks ? new List<ReaderChunk>() : null;
        var files = new Dictionary<string, ReaderIngestFileResult>(IsWindows() ? StringComparer.OrdinalIgnoreCase : StringComparer.Ordinal);
        var warnings = new List<string>();
        ReaderProgress? completed = null;

        void HandleProgress(ReaderProgress progress) {
            onProgress?.Invoke(progress);
            if (progress.Kind == ReaderProgressEventKind.Completed) {
                completed = progress;
                return;
            }
            var progressPath = progress.Path;
            if (string.IsNullOrWhiteSpace(progressPath)) return;
            var filePath = progressPath!;

            if (!files.TryGetValue(filePath, out var file)) {
                file = new ReaderIngestFileResult {
                    Path = filePath
                };
                files[filePath] = file;
            }

            file.SourceId = progress.SourceId ?? file.SourceId;
            file.SourceHash = progress.SourceHash ?? file.SourceHash;
            file.SourceLengthBytes = progress.CurrentFileBytes ?? file.SourceLengthBytes;
            file.SourceLastWriteUtc = progress.CurrentFileLastWriteUtc ?? file.SourceLastWriteUtc;

            if (progress.Kind == ReaderProgressEventKind.FileCompleted) {
                file.Parsed = true;
                file.ChunksProduced = progress.CurrentFileChunks ?? file.ChunksProduced;
            } else if (progress.Kind == ReaderProgressEventKind.FileSkipped) {
                file.Parsed = false;
                if (!string.IsNullOrWhiteSpace(progress.Message)) {
                    var list = file.Warnings?.ToList() ?? new List<string>();
                    list.Add(progress.Message!);
                    file.Warnings = list;
                    warnings.Add(progress.Message!);
                }
            }
        }

        foreach (var chunk in ReadFolder(folderPath, folderOptions, options, HandleProgress, cancellationToken)) {
            cancellationToken.ThrowIfCancellationRequested();
            if (includeChunks) {
                chunks!.Add(chunk);
            }
            if (chunk.Warnings != null && chunk.Warnings.Count > 0) {
                warnings.AddRange(chunk.Warnings);
            }
        }

        var snapshot = completed ?? new ReaderProgress {
            Kind = ReaderProgressEventKind.Completed,
            FilesScanned = files.Count,
            FilesParsed = files.Values.Count(f => f.Parsed),
            FilesSkipped = files.Values.Count(f => !f.Parsed),
            BytesRead = files.Values.Where(f => f.Parsed).Sum(f => f.SourceLengthBytes ?? 0),
            ChunksProduced = includeChunks ? chunks!.Count : files.Values.Sum(f => f.ChunksProduced)
        };

        return new ReaderIngestResult {
            Files = files.Values
                .OrderBy(static f => f.Path, StringComparer.Ordinal)
                .ToArray(),
            Chunks = includeChunks ? chunks! : Array.Empty<ReaderChunk>(),
            FilesScanned = snapshot.FilesScanned,
            FilesParsed = snapshot.FilesParsed,
            FilesSkipped = snapshot.FilesSkipped,
            BytesRead = snapshot.BytesRead,
            ChunksProduced = snapshot.ChunksProduced,
            Warnings = warnings.Count > 0 ? warnings.ToArray() : null
        };
    }

    /// <summary>
    /// Reads a supported document from a stream and emits normalized extraction chunks.
    /// </summary>
    /// <param name="stream">Source stream. This method does not close the stream.</param>
    /// <param name="sourceName">
    /// Optional source name used for kind detection (via extension) and citations/IDs.
    /// For example: "Policy.docx" or "Workbook.xlsx".
    /// </param>
    /// <param name="options">Extraction options.</param>
    /// <param name="cancellationToken">Cancellation token.</param>
    public static IEnumerable<ReaderChunk> Read(Stream stream, string? sourceName = null, ReaderOptions? options = null, CancellationToken cancellationToken = default) {
        if (stream == null) throw new ArgumentNullException(nameof(stream));
        if (!stream.CanRead) throw new ArgumentException("Stream must be readable.", nameof(stream));

        var opt = NormalizeOptions(options);
        EnforceStreamSize(stream, opt.MaxInputBytes);
        var source = BuildSourceInfoFromStream(stream, sourceName, opt.ComputeHashes);

        IEnumerable<ReaderChunk> raw;
        if (TryResolveCustomHandlerBySourceName(sourceName, out var customStreamHandler) && customStreamHandler.ReadStream != null) {
            raw = customStreamHandler.ReadStream(stream, sourceName, opt, cancellationToken);
        } else {
            var kind = string.IsNullOrWhiteSpace(sourceName) ? ReaderInputKind.Unknown : DetectKind(sourceName!);
            raw = kind switch {
                ReaderInputKind.Word => ReadWord(stream, sourceName, opt, cancellationToken),
                ReaderInputKind.Excel => ReadExcel(stream, sourceName, opt, cancellationToken),
                ReaderInputKind.PowerPoint => ReadPowerPoint(stream, sourceName, opt, cancellationToken),
                ReaderInputKind.Markdown => ReadMarkdown(stream, sourceName, opt, cancellationToken),
                ReaderInputKind.Pdf => ReadPdf(stream, sourceName, opt, cancellationToken),
                ReaderInputKind.Text => ReadText(stream, sourceName, opt, cancellationToken),
                _ => ReadUnknown(stream, sourceName, opt, cancellationToken)
            };
        }

        foreach (var chunk in raw) {
            cancellationToken.ThrowIfCancellationRequested();
            yield return EnrichChunk(chunk, source, opt.ComputeHashes);
        }
    }

    /// <summary>
    /// Reads a supported document from bytes and emits normalized extraction chunks.
    /// </summary>
    /// <param name="bytes">Source bytes.</param>
    /// <param name="sourceName">
    /// Optional source name used for kind detection (via extension) and citations/IDs.
    /// For example: "Policy.docx" or "Workbook.xlsx".
    /// </param>
    /// <param name="options">Extraction options.</param>
    /// <param name="cancellationToken">Cancellation token.</param>
    public static IEnumerable<ReaderChunk> Read(byte[] bytes, string? sourceName = null, ReaderOptions? options = null, CancellationToken cancellationToken = default) {
        if (bytes == null) throw new ArgumentNullException(nameof(bytes));
        using var ms = new MemoryStream(bytes, writable: false);
        foreach (var c in Read(ms, sourceName, options, cancellationToken))
            yield return c;
    }

    private static IEnumerable<ReaderChunk> ReadWord(string path, ReaderOptions opt, CancellationToken ct) {
        using var doc = WordDocument.Load(path, readOnly: true, autoSave: false, openSettings: CreateOpenSettings(opt));
        var chunks = doc.ExtractMarkdownChunks(
            markdownOptions: new WordToMarkdownOptions(),
            chunking: new WordMarkdownChunkingOptions { MaxChars = opt.MaxChars, IncludeFootnotes = opt.IncludeWordFootnotes },
            sourcePath: path,
            cancellationToken: ct);

        int outIndex = 0;
        foreach (var c in chunks) {
            ct.ThrowIfCancellationRequested();
            yield return new ReaderChunk {
                Id = c.Id,
                Kind = ReaderInputKind.Word,
                Location = new ReaderLocation {
                    Path = c.Location.Path,
                    BlockIndex = outIndex,
                    SourceBlockIndex = c.Location.BlockIndex,
                    HeadingPath = c.Location.HeadingPath
                },
                Text = c.Text,
                Markdown = c.Markdown,
                Warnings = c.Warnings
            };
            outIndex++;
        }
    }

    private static IEnumerable<ReaderChunk> ReadWord(Stream stream, string? sourceName, ReaderOptions opt, CancellationToken ct) {
        // Copy input so we can open read-only without affecting caller's stream.
        using var ms = CopyToMemory(stream, ct);
        using var doc = WordDocument.Load(ms, readOnly: true, autoSave: false, openSettings: CreateOpenSettings(opt));

        var chunks = doc.ExtractMarkdownChunks(
            markdownOptions: new WordToMarkdownOptions(),
            chunking: new WordMarkdownChunkingOptions { MaxChars = opt.MaxChars, IncludeFootnotes = opt.IncludeWordFootnotes },
            sourcePath: sourceName,
            cancellationToken: ct);

        int outIndex = 0;
        foreach (var c in chunks) {
            ct.ThrowIfCancellationRequested();
            yield return new ReaderChunk {
                Id = c.Id,
                Kind = ReaderInputKind.Word,
                Location = new ReaderLocation {
                    Path = sourceName,
                    BlockIndex = outIndex,
                    SourceBlockIndex = c.Location.BlockIndex,
                    HeadingPath = c.Location.HeadingPath
                },
                Text = c.Text,
                Markdown = c.Markdown,
                Warnings = c.Warnings
            };
            outIndex++;
        }
    }

    private static IEnumerable<ReaderChunk> ReadExcel(string path, ReaderOptions opt, CancellationToken ct) {
        // Use OpenSettings for basic OpenXML hardening (best-effort) and open from stream to avoid file handle collisions.
        using var fs = new FileStream(path, FileMode.Open, FileAccess.Read, FileShare.ReadWrite | FileShare.Delete);
        var openSettings = CreateOpenSettings(opt);
        using var openXml = openSettings == null
            ? SpreadsheetDocument.Open(fs, false)
            : SpreadsheetDocument.Open(fs, false, openSettings);
        using var reader = ExcelDocumentReader.Wrap(openXml);
        var sheets = ResolveSheetNames(reader, opt.ExcelSheetName);

        int outIndex = 0;
        foreach (var sheet in sheets) {
            ct.ThrowIfCancellationRequested();

            var chunks = reader.ExtractChunks(
                sheetName: sheet,
                a1Range: opt.ExcelA1Range,
                extract: new ExcelExtractionExtensions.ExcelExtractOptions {
                    HeadersInFirstRow = opt.ExcelHeadersInFirstRow,
                    ChunkRows = opt.ExcelChunkRows,
                    EmitMarkdownTable = true
                },
                chunking: new ExcelExtractChunkingOptions { MaxChars = opt.MaxChars, MaxTableRows = opt.MaxTableRows },
                sourcePath: path,
                cancellationToken: ct);

            foreach (var c in chunks) {
                ct.ThrowIfCancellationRequested();

                IReadOnlyList<ReaderTable>? tables = null;
                if (c.Tables != null && c.Tables.Count > 0) {
                    tables = c.Tables.Select(MapTable).ToArray();
                }

                yield return new ReaderChunk {
                    Id = c.Id,
                    Kind = ReaderInputKind.Excel,
                    Location = new ReaderLocation {
                        Path = c.Location.Path,
                        Sheet = c.Location.Sheet,
                        A1Range = c.Location.A1Range,
                        BlockIndex = outIndex,
                        SourceBlockIndex = c.Location.BlockIndex
                    },
                    Text = c.Text,
                    Markdown = c.Markdown,
                    Tables = tables,
                    Warnings = c.Warnings
                };
                outIndex++;
            }
        }
    }

    private static IEnumerable<ReaderChunk> ReadExcel(Stream stream, string? sourceName, ReaderOptions opt, CancellationToken ct) {
        // Avoid exposing OpenXml types in the public API surface; internally we can wrap.
        using var ms = CopyToMemory(stream, ct);
        var openSettings = CreateOpenSettings(opt);
        using var openXml = openSettings == null
            ? SpreadsheetDocument.Open(ms, false)
            : SpreadsheetDocument.Open(ms, false, openSettings);
        using var reader = ExcelDocumentReader.Wrap(openXml);

        var sheets = ResolveSheetNames(reader, opt.ExcelSheetName);

        int outIndex = 0;
        foreach (var sheet in sheets) {
            ct.ThrowIfCancellationRequested();

            var chunks = reader.ExtractChunks(
                sheetName: sheet,
                a1Range: opt.ExcelA1Range,
                extract: new ExcelExtractionExtensions.ExcelExtractOptions {
                    HeadersInFirstRow = opt.ExcelHeadersInFirstRow,
                    ChunkRows = opt.ExcelChunkRows,
                    EmitMarkdownTable = true
                },
                chunking: new ExcelExtractChunkingOptions { MaxChars = opt.MaxChars, MaxTableRows = opt.MaxTableRows },
                sourcePath: sourceName,
                cancellationToken: ct);

            foreach (var c in chunks) {
                ct.ThrowIfCancellationRequested();

                IReadOnlyList<ReaderTable>? tables = null;
                if (c.Tables != null && c.Tables.Count > 0) {
                    tables = c.Tables.Select(MapTable).ToArray();
                }

                yield return new ReaderChunk {
                    Id = c.Id,
                    Kind = ReaderInputKind.Excel,
                    Location = new ReaderLocation {
                        Path = sourceName,
                        Sheet = c.Location.Sheet,
                        A1Range = c.Location.A1Range,
                        BlockIndex = outIndex,
                        SourceBlockIndex = c.Location.BlockIndex
                    },
                    Text = c.Text,
                    Markdown = c.Markdown,
                    Tables = tables,
                    Warnings = c.Warnings
                };
                outIndex++;
            }
        }
    }

    private static IEnumerable<ReaderChunk> ReadPowerPoint(string path, ReaderOptions opt, CancellationToken ct) {
        using var presentation = PowerPointPresentation.OpenRead(path);
        var chunks = presentation.ExtractMarkdownChunks(
            extract: new PowerPointExtractionExtensions.PowerPointExtractOptions { IncludeNotes = opt.IncludePowerPointNotes },
            chunking: new PowerPointExtractChunkingOptions { MaxChars = opt.MaxChars },
            sourcePath: path,
            cancellationToken: ct);

        int outIndex = 0;
        foreach (var c in chunks) {
            ct.ThrowIfCancellationRequested();
            yield return new ReaderChunk {
                Id = c.Id,
                Kind = ReaderInputKind.PowerPoint,
                Location = new ReaderLocation {
                    Path = c.Location.Path,
                    Slide = c.Location.Slide,
                    BlockIndex = outIndex,
                    SourceBlockIndex = c.Location.BlockIndex
                },
                Text = c.Text,
                Markdown = c.Markdown,
                Warnings = c.Warnings
            };
            outIndex++;
        }
    }

    private static IEnumerable<ReaderChunk> ReadPowerPoint(Stream stream, string? sourceName, ReaderOptions opt, CancellationToken ct) {
        // PowerPointPresentation.Open(stream, readOnly:true) already copies to an internal stream for safety.
        using var presentation = PowerPointPresentation.Open(stream, readOnly: true, autoSave: false);
        var chunks = presentation.ExtractMarkdownChunks(
            extract: new PowerPointExtractionExtensions.PowerPointExtractOptions { IncludeNotes = opt.IncludePowerPointNotes },
            chunking: new PowerPointExtractChunkingOptions { MaxChars = opt.MaxChars },
            sourcePath: sourceName,
            cancellationToken: ct);

        int outIndex = 0;
        foreach (var c in chunks) {
            ct.ThrowIfCancellationRequested();
            yield return new ReaderChunk {
                Id = c.Id,
                Kind = ReaderInputKind.PowerPoint,
                Location = new ReaderLocation {
                    Path = sourceName,
                    Slide = c.Location.Slide,
                    BlockIndex = outIndex,
                    SourceBlockIndex = c.Location.BlockIndex
                },
                Text = c.Text,
                Markdown = c.Markdown,
                Warnings = c.Warnings
            };
            outIndex++;
        }
    }

    private static IEnumerable<ReaderChunk> ReadPdf(string path, ReaderOptions opt, CancellationToken ct) {
        var fileName = Path.GetFileName(path);
        var doc = PdfReadDocument.Load(path);
        int outIndex = 0;

        for (int pageIndex = 0; pageIndex < doc.Pages.Count; pageIndex++) {
            ct.ThrowIfCancellationRequested();

            var pageNumber = pageIndex + 1;
            var pageText = doc.Pages[pageIndex].ExtractText();
            if (string.IsNullOrWhiteSpace(pageText)) {
                yield return BuildPdfEmptyChunk(path, fileName, pageNumber, outIndex);
                outIndex++;
                continue;
            }

            var pageChunks = ChunkPdfText(path, fileName, pageNumber, pageText, opt, outIndex, ct, out var nextIndex);
            outIndex = nextIndex;
            foreach (var chunk in pageChunks) {
                yield return chunk;
            }
        }
    }

    private static IEnumerable<ReaderChunk> ReadPdf(Stream stream, string? sourceName, ReaderOptions opt, CancellationToken ct) {
        using var ms = CopyToMemory(stream, ct);
        var fileName = string.IsNullOrWhiteSpace(sourceName) ? "memory.pdf" : Path.GetFileName(sourceName!.Trim());
        var doc = PdfReadDocument.Load(ms.ToArray());
        int outIndex = 0;

        for (int pageIndex = 0; pageIndex < doc.Pages.Count; pageIndex++) {
            ct.ThrowIfCancellationRequested();

            var pageNumber = pageIndex + 1;
            var pageText = doc.Pages[pageIndex].ExtractText();
            if (string.IsNullOrWhiteSpace(pageText)) {
                yield return BuildPdfEmptyChunk(sourceName ?? fileName, fileName, pageNumber, outIndex);
                outIndex++;
                continue;
            }

            var pageChunks = ChunkPdfText(sourceName ?? fileName, fileName, pageNumber, pageText, opt, outIndex, ct, out var nextIndex);
            outIndex = nextIndex;
            foreach (var chunk in pageChunks) {
                yield return chunk;
            }
        }
    }

    private static IEnumerable<ReaderChunk> ReadMarkdown(string path, ReaderOptions opt, CancellationToken ct) {
        // Keep it simple: chunk by headings (ATX, best-effort), with size cap.
        if (!opt.MarkdownChunkByHeadings) {
            foreach (var c in ChunkPlainTextByParagraphs(path, opt, ReaderInputKind.Markdown, ct, treatAsMarkdown: true))
                yield return c;
            yield break;
        }

        var fileName = Path.GetFileName(path);
        var headingStack = new List<(int Level, string Text)>();

        var current = new StringBuilder(capacity: Math.Min(opt.MaxChars, 16_384));
        int chunkIndex = 0;
        int? firstLine = null;
        string? firstHeadingPath = null;
        var warnings = new List<string>(capacity: 2);

        int lineNo = 0;
        foreach (var line in File.ReadLines(path)) {
            ct.ThrowIfCancellationRequested();
            lineNo++;

            if (TryParseAtxHeading(line, out var level, out var headingText)) {
                // Flush current section before starting a new heading section.
                if (current.Length > 0) {
                    yield return BuildMarkdownChunk(path, fileName, chunkIndex, firstLine, firstHeadingPath, current.ToString().TrimEnd(), warnings);
                    chunkIndex++;
                    current.Clear();
                    warnings.Clear();
                    firstLine = null;
                    firstHeadingPath = null;
                }

                UpdateHeadingStack(headingStack, level, headingText);
                var headingPath = BuildHeadingPath(headingStack);
                firstHeadingPath = headingPath;
                firstLine = lineNo;

                // Keep the heading line as part of the new chunk content.
                AppendLineCapped(opt, current, line, warnings);
                continue;
            }

            if (firstLine == null) firstLine = lineNo;
            if (firstHeadingPath == null) firstHeadingPath = BuildHeadingPath(headingStack);

            // If adding this line would exceed MaxChars, flush a chunk boundary.
            if (WouldExceed(opt, current, line)) {
                yield return BuildMarkdownChunk(path, fileName, chunkIndex, firstLine, firstHeadingPath, current.ToString().TrimEnd(), warnings);
                chunkIndex++;
                current.Clear();
                warnings.Clear();
                firstLine = lineNo;
                firstHeadingPath = BuildHeadingPath(headingStack);
            }

            AppendLineCapped(opt, current, line, warnings);
        }

        if (current.Length > 0) {
            yield return BuildMarkdownChunk(path, fileName, chunkIndex, firstLine, firstHeadingPath, current.ToString().TrimEnd(), warnings);
        }
    }

    private static IEnumerable<ReaderChunk> ReadMarkdown(Stream stream, string? sourceName, ReaderOptions opt, CancellationToken ct) {
        var fileName = string.IsNullOrWhiteSpace(sourceName) ? "memory.md" : Path.GetFileName(sourceName!.Trim());
        var text = ReadAllText(stream, ct);
        foreach (var c in ChunkMarkdownFromText(text, sourceName, fileName, opt, ct))
            yield return c;
    }

    private static IEnumerable<ReaderChunk> ReadText(string path, ReaderOptions opt, CancellationToken ct) {
        foreach (var c in ChunkPlainTextByParagraphs(path, opt, ReaderInputKind.Text, ct, treatAsMarkdown: false))
            yield return c;
    }

    private static IEnumerable<ReaderChunk> ReadText(Stream stream, string? sourceName, ReaderOptions opt, CancellationToken ct) {
        var fileName = string.IsNullOrWhiteSpace(sourceName) ? "memory.txt" : Path.GetFileName(sourceName!.Trim());
        var text = ReadAllText(stream, ct);
        foreach (var c in ChunkPlainTextFromText(text, sourceName, fileName, opt, ReaderInputKind.Text, ct, treatAsMarkdown: false))
            yield return c;
    }

    private static IEnumerable<ReaderChunk> ReadUnknown(string path, ReaderOptions opt, CancellationToken ct) {
        var extLower = (TryGetExtension(path) ?? string.Empty).ToLowerInvariant();
        if (extLower is ".doc" or ".xls" or ".ppt") {
            throw new NotSupportedException($"Legacy binary format '{extLower}' is not supported. Convert to OpenXML (.docx/.xlsx/.pptx) first.");
        }

        // Try plain text; if it fails (binary), the caller can decide how to handle it.
        foreach (var c in ChunkPlainTextByParagraphs(path, opt, ReaderInputKind.Unknown, ct, treatAsMarkdown: false))
            yield return c;
    }

    private static IEnumerable<ReaderChunk> ReadUnknown(Stream stream, string? sourceName, ReaderOptions opt, CancellationToken ct) {
        // When we can't detect kind, treat as plain text.
        var fileName = string.IsNullOrWhiteSpace(sourceName) ? "memory" : Path.GetFileName(sourceName!.Trim());
        var text = ReadAllText(stream, ct);
        foreach (var c in ChunkPlainTextFromText(text, sourceName, fileName, opt, ReaderInputKind.Unknown, ct, treatAsMarkdown: false))
            yield return c;
    }

    private static IReadOnlyList<string> ResolveSheetNames(ExcelDocumentReader reader, string? singleSheet) {
        if (reader == null) throw new ArgumentNullException(nameof(reader));
        if (!string.IsNullOrWhiteSpace(singleSheet)) return new[] { singleSheet!.Trim() };
        return reader.GetSheetNames();
    }

    private static ReaderTable MapTable(ExcelExtractTable t) {
        return new ReaderTable {
            Title = t.Title,
            Columns = t.Columns,
            Rows = t.Rows,
            TotalRowCount = t.TotalRowCount,
            Truncated = t.Truncated
        };
    }

    private static IEnumerable<ReaderChunk> ChunkPlainTextByParagraphs(
        string path,
        ReaderOptions opt,
        ReaderInputKind kind,
        CancellationToken ct,
        bool treatAsMarkdown) {
        var fileName = Path.GetFileName(path);
        var current = new StringBuilder(capacity: Math.Min(opt.MaxChars, 16_384));
        int chunkIndex = 0;
        int? firstLine = null;
        var warnings = new List<string>(capacity: 2);
        int lineNo = 0;

        foreach (var line in File.ReadLines(path)) {
            ct.ThrowIfCancellationRequested();
            lineNo++;

            if (firstLine == null) firstLine = lineNo;

            // Prefer splitting at empty lines when close to the cap.
            if (current.Length > 0 && current.Length >= (opt.MaxChars - 256) && string.IsNullOrWhiteSpace(line)) {
                yield return BuildTextChunk(path, fileName, kind, chunkIndex, firstLine, current.ToString().TrimEnd(), warnings, treatAsMarkdown);
                chunkIndex++;
                current.Clear();
                warnings.Clear();
                firstLine = null;
                continue;
            }

            if (WouldExceed(opt, current, line)) {
                yield return BuildTextChunk(path, fileName, kind, chunkIndex, firstLine, current.ToString().TrimEnd(), warnings, treatAsMarkdown);
                chunkIndex++;
                current.Clear();
                warnings.Clear();
                firstLine = lineNo;
            }

            AppendLineCapped(opt, current, line, warnings);
        }

        if (current.Length > 0) {
            yield return BuildTextChunk(path, fileName, kind, chunkIndex, firstLine, current.ToString().TrimEnd(), warnings, treatAsMarkdown);
        }
    }

    private static IEnumerable<ReaderChunk> ChunkPlainTextFromText(
        string text,
        string? sourceName,
        string fileName,
        ReaderOptions opt,
        ReaderInputKind kind,
        CancellationToken ct,
        bool treatAsMarkdown) {
        var current = new StringBuilder(capacity: Math.Min(opt.MaxChars, 16_384));
        int chunkIndex = 0;
        int? firstLine = null;
        var warnings = new List<string>(capacity: 2);
        int lineNo = 0;

        using var sr = new StringReader(text ?? string.Empty);
        string? line;
        while ((line = sr.ReadLine()) != null) {
            ct.ThrowIfCancellationRequested();
            lineNo++;

            if (firstLine == null) firstLine = lineNo;

            if (current.Length > 0 && current.Length >= (opt.MaxChars - 256) && string.IsNullOrWhiteSpace(line)) {
                yield return BuildTextChunk(sourceName ?? fileName, fileName, kind, chunkIndex, firstLine, current.ToString().TrimEnd(), warnings, treatAsMarkdown);
                chunkIndex++;
                current.Clear();
                warnings.Clear();
                firstLine = null;
                continue;
            }

            if (WouldExceed(opt, current, line)) {
                yield return BuildTextChunk(sourceName ?? fileName, fileName, kind, chunkIndex, firstLine, current.ToString().TrimEnd(), warnings, treatAsMarkdown);
                chunkIndex++;
                current.Clear();
                warnings.Clear();
                firstLine = lineNo;
            }

            AppendLineCapped(opt, current, line, warnings);
        }

        if (current.Length > 0) {
            yield return BuildTextChunk(sourceName ?? fileName, fileName, kind, chunkIndex, firstLine, current.ToString().TrimEnd(), warnings, treatAsMarkdown);
        }
    }

    private static IEnumerable<ReaderChunk> ChunkMarkdownFromText(string text, string? sourceName, string fileName, ReaderOptions opt, CancellationToken ct) {
        if (!opt.MarkdownChunkByHeadings) {
            foreach (var c in ChunkPlainTextFromText(text, sourceName, fileName, opt, ReaderInputKind.Markdown, ct, treatAsMarkdown: true))
                yield return c;
            yield break;
        }

        var headingStack = new List<(int Level, string Text)>();
        var current = new StringBuilder(capacity: Math.Min(opt.MaxChars, 16_384));
        int chunkIndex = 0;
        int? firstLine = null;
        string? firstHeadingPath = null;
        var warnings = new List<string>(capacity: 2);

        int lineNo = 0;
        using var sr = new StringReader(text ?? string.Empty);
        string? line;
        while ((line = sr.ReadLine()) != null) {
            ct.ThrowIfCancellationRequested();
            lineNo++;

            if (TryParseAtxHeading(line, out var level, out var headingText)) {
                if (current.Length > 0) {
                    yield return BuildMarkdownChunk(sourceName ?? fileName, fileName, chunkIndex, firstLine, firstHeadingPath, current.ToString().TrimEnd(), warnings);
                    chunkIndex++;
                    current.Clear();
                    warnings.Clear();
                    firstLine = null;
                    firstHeadingPath = null;
                }

                UpdateHeadingStack(headingStack, level, headingText);
                var headingPath = BuildHeadingPath(headingStack);
                firstHeadingPath = headingPath;
                firstLine = lineNo;

                AppendLineCapped(opt, current, line, warnings);
                continue;
            }

            if (firstLine == null) firstLine = lineNo;
            if (firstHeadingPath == null) firstHeadingPath = BuildHeadingPath(headingStack);

            if (WouldExceed(opt, current, line)) {
                yield return BuildMarkdownChunk(sourceName ?? fileName, fileName, chunkIndex, firstLine, firstHeadingPath, current.ToString().TrimEnd(), warnings);
                chunkIndex++;
                current.Clear();
                warnings.Clear();
                firstLine = lineNo;
                firstHeadingPath = BuildHeadingPath(headingStack);
            }

            AppendLineCapped(opt, current, line, warnings);
        }

        if (current.Length > 0) {
            yield return BuildMarkdownChunk(sourceName ?? fileName, fileName, chunkIndex, firstLine, firstHeadingPath, current.ToString().TrimEnd(), warnings);
        }
    }

    private static List<ReaderChunk> ChunkPdfText(
        string path,
        string fileName,
        int pageNumber,
        string text,
        ReaderOptions opt,
        int startChunkIndex,
        CancellationToken ct,
        out int nextChunkIndex) {
        var list = new List<ReaderChunk>();
        var current = new StringBuilder(capacity: Math.Min(opt.MaxChars, 16_384));
        var outIndex = startChunkIndex;
        int? firstLine = null;
        var warnings = new List<string>(capacity: 2);
        int lineNo = 0;

        using var sr = new StringReader(text ?? string.Empty);
        string? line;
        while ((line = sr.ReadLine()) != null) {
            ct.ThrowIfCancellationRequested();
            lineNo++;

            if (firstLine == null) firstLine = lineNo;

            if (current.Length > 0 && current.Length >= (opt.MaxChars - 256) && string.IsNullOrWhiteSpace(line)) {
                list.Add(BuildPdfChunk(path, fileName, pageNumber, outIndex, firstLine, current.ToString().TrimEnd(), warnings));
                outIndex++;
                current.Clear();
                warnings.Clear();
                firstLine = null;
                continue;
            }

            if (WouldExceed(opt, current, line)) {
                list.Add(BuildPdfChunk(path, fileName, pageNumber, outIndex, firstLine, current.ToString().TrimEnd(), warnings));
                outIndex++;
                current.Clear();
                warnings.Clear();
                firstLine = lineNo;
            }

            AppendLineCapped(opt, current, line, warnings);
        }

        if (current.Length > 0) {
            list.Add(BuildPdfChunk(path, fileName, pageNumber, outIndex, firstLine, current.ToString().TrimEnd(), warnings));
            outIndex++;
        }
        nextChunkIndex = outIndex;
        return list;
    }

    private static ReaderChunk BuildMarkdownChunk(
        string path,
        string fileName,
        int chunkIndex,
        int? firstLine,
        string? headingPath,
        string markdown,
        List<string> warnings) {
        var id = BuildStableId("md", fileName, chunkIndex, firstLine);
        return new ReaderChunk {
            Id = id,
            Kind = ReaderInputKind.Markdown,
            Location = new ReaderLocation {
                Path = path,
                BlockIndex = chunkIndex,
                StartLine = firstLine,
                HeadingPath = headingPath
            },
            Text = markdown,
            Markdown = markdown,
            Warnings = warnings.Count > 0 ? warnings.ToArray() : null
        };
    }

    private static ReaderChunk BuildTextChunk(
        string path,
        string fileName,
        ReaderInputKind kind,
        int chunkIndex,
        int? firstLine,
        string text,
        List<string> warnings,
        bool treatAsMarkdown) {
        var id = BuildStableId(kind == ReaderInputKind.Text ? "text" : "unknown", fileName, chunkIndex, firstLine);
        return new ReaderChunk {
            Id = id,
            Kind = kind,
            Location = new ReaderLocation {
                Path = path,
                BlockIndex = chunkIndex,
                StartLine = firstLine
            },
            Text = text,
            Markdown = treatAsMarkdown ? text : null,
            Warnings = warnings.Count > 0 ? warnings.ToArray() : null
        };
    }

    private static ReaderChunk BuildPdfChunk(
        string path,
        string fileName,
        int pageNumber,
        int chunkIndex,
        int? firstLine,
        string text,
        List<string> warnings) {
        var id = BuildStableId("pdf", fileName, chunkIndex, firstLine);
        return new ReaderChunk {
            Id = id,
            Kind = ReaderInputKind.Pdf,
            Location = new ReaderLocation {
                Path = path,
                Page = pageNumber,
                BlockIndex = chunkIndex,
                SourceBlockIndex = pageNumber - 1,
                StartLine = firstLine
            },
            Text = text,
            Markdown = null,
            Warnings = warnings.Count > 0 ? warnings.ToArray() : null
        };
    }

    private static ReaderChunk BuildPdfEmptyChunk(
        string path,
        string fileName,
        int pageNumber,
        int chunkIndex) {
        var id = BuildStableId("pdf", fileName, chunkIndex, null);
        return new ReaderChunk {
            Id = id,
            Kind = ReaderInputKind.Pdf,
            Location = new ReaderLocation {
                Path = path,
                Page = pageNumber,
                BlockIndex = chunkIndex,
                SourceBlockIndex = pageNumber - 1
            },
            Text = string.Empty,
            Markdown = null,
            Warnings = new[] { "No extractable text found on this PDF page." }
        };
    }

    private static ReaderChunk BuildFolderWarningChunk(string path, int warningIndex, string warning) {
        var fileName = Path.GetFileName(path);
        if (string.IsNullOrWhiteSpace(fileName)) fileName = "folder";

        return new ReaderChunk {
            Id = BuildStableId("warn", fileName, warningIndex, null),
            Kind = ReaderInputKind.Unknown,
            Location = new ReaderLocation {
                Path = path,
                BlockIndex = warningIndex
            },
            Text = string.Empty,
            Markdown = null,
            Warnings = new[] { warning }
        };
    }

    private static void NotifyProgress(
        Action<ReaderProgress>? onProgress,
        ReaderProgressEventKind kind,
        FolderIngestState state,
        SourceInfo? source,
        string? message,
        int? fileChunkCount) {
        if (onProgress == null) return;

        onProgress(new ReaderProgress {
            Kind = kind,
            Path = source?.Path,
            SourceId = source?.SourceId,
            SourceHash = source?.SourceHash,
            FilesScanned = state.FilesScanned,
            FilesParsed = state.FilesParsed,
            FilesSkipped = state.FilesSkipped,
            BytesRead = state.BytesRead,
            ChunksProduced = state.ChunksProduced,
            Message = message,
            CurrentFileBytes = source?.LengthBytes,
            CurrentFileChunks = fileChunkCount,
            CurrentFileLastWriteUtc = source?.LastWriteUtc
        });
    }

    private static ReaderSourceDocument BuildSourceDocument(
        SourceInfo source,
        bool parsed,
        IReadOnlyList<ReaderChunk>? chunks,
        IReadOnlyList<string>? sourceWarnings) {
        var chunkList = chunks ?? Array.Empty<ReaderChunk>();
        var warnings = new List<string>();
        if (sourceWarnings != null) {
            for (int i = 0; i < sourceWarnings.Count; i++) {
                if (!string.IsNullOrWhiteSpace(sourceWarnings[i])) {
                    warnings.Add(sourceWarnings[i]!);
                }
            }
        }

        int tokenEstimateTotal = 0;
        for (int i = 0; i < chunkList.Count; i++) {
            var chunk = chunkList[i];
            tokenEstimateTotal += chunk.TokenEstimate ?? EstimateTokenCount(chunk.Text);

            if (chunk.Warnings == null) continue;
            for (int j = 0; j < chunk.Warnings.Count; j++) {
                var warning = chunk.Warnings[j];
                if (!string.IsNullOrWhiteSpace(warning)) {
                    warnings.Add(warning!);
                }
            }
        }

        return new ReaderSourceDocument {
            Path = source.Path,
            SourceId = source.SourceId,
            SourceHash = source.SourceHash,
            SourceLastWriteUtc = source.LastWriteUtc,
            SourceLengthBytes = source.LengthBytes,
            Parsed = parsed,
            ChunksProduced = chunkList.Count,
            TokenEstimateTotal = tokenEstimateTotal,
            Warnings = warnings.Count > 0 ? warnings : null,
            Chunks = chunkList
        };
    }

    private static ReaderChunk EnrichChunk(ReaderChunk chunk, SourceInfo source, bool computeHashes) {
        if (chunk == null) throw new ArgumentNullException(nameof(chunk));
        if (source == null) throw new ArgumentNullException(nameof(source));

        chunk.SourceId ??= source.SourceId;
        chunk.SourceHash ??= source.SourceHash;
        chunk.SourceLastWriteUtc ??= source.LastWriteUtc;
        chunk.SourceLengthBytes ??= source.LengthBytes;
        if (!chunk.TokenEstimate.HasValue) {
            chunk.TokenEstimate = EstimateTokenCount(chunk.Markdown ?? chunk.Text);
        }
        if (computeHashes && string.IsNullOrWhiteSpace(chunk.ChunkHash)) {
            chunk.ChunkHash = ComputeChunkHash(chunk);
        }
        return chunk;
    }

    private static int EstimateTokenCount(string? text) {
        var safeText = text ?? string.Empty;
        if (safeText.Length == 0) return 0;
        // Heuristic: roughly 4 characters per token for mixed English/code.
        return Math.Max(1, (safeText.Length + 3) / 4);
    }

    private static string ComputeChunkHash(ReaderChunk chunk) {
        var data = string.Join("|",
            chunk.Kind.ToString(),
            chunk.SourceId ?? string.Empty,
            chunk.Location.Path ?? string.Empty,
            chunk.Location.HeadingPath ?? string.Empty,
            chunk.Location.Sheet ?? string.Empty,
            chunk.Location.A1Range ?? string.Empty,
            chunk.Location.Page?.ToString(CultureInfo.InvariantCulture) ?? string.Empty,
            chunk.Location.Slide?.ToString(CultureInfo.InvariantCulture) ?? string.Empty,
            chunk.Location.StartLine?.ToString(CultureInfo.InvariantCulture) ?? string.Empty,
            chunk.Text ?? string.Empty,
            chunk.Markdown ?? string.Empty);
        return ComputeSha256Hex(data);
    }

    private static SourceInfo BuildSourceInfoFromPath(string path, bool computeHash) {
        string normalizedPath = NormalizePathForId(path);
        string sourceId = BuildSourceId(normalizedPath);

        DateTime? lastWriteUtc = null;
        long? lengthBytes = null;
        try {
            var fi = new FileInfo(path);
            if (fi.Exists) {
                lastWriteUtc = fi.LastWriteTimeUtc;
                lengthBytes = fi.Length;
            }
        } catch {
            // Best-effort metadata; leave null on failure.
        }

        string? sourceHash = null;
        if (computeHash) {
            sourceHash = TryComputeFileSha256(path);
        }

        return new SourceInfo {
            Path = path,
            SourceId = sourceId,
            SourceHash = sourceHash,
            LastWriteUtc = lastWriteUtc,
            LengthBytes = lengthBytes
        };
    }

    private static SourceInfo BuildSourceInfoFromStream(Stream stream, string? sourceName, bool computeHash) {
        string logicalName = "memory";
        if (!string.IsNullOrWhiteSpace(sourceName)) {
            logicalName = sourceName!.Trim();
        }
        string sourceId = BuildSourceId(logicalName);

        long? lengthBytes = null;
        try {
            if (stream.CanSeek) {
                lengthBytes = stream.Length;
            }
        } catch {
            // Best-effort metadata.
        }

        string? sourceHash = null;
        if (computeHash) {
            sourceHash = TryComputeStreamSha256(stream);
        }

        return new SourceInfo {
            Path = logicalName,
            SourceId = sourceId,
            SourceHash = sourceHash,
            LastWriteUtc = null,
            LengthBytes = lengthBytes
        };
    }

    private static string? TryComputeFileSha256(string path) {
        try {
            using var fs = new FileStream(path, FileMode.Open, FileAccess.Read, FileShare.ReadWrite | FileShare.Delete);
            return ComputeSha256Hex(fs);
        } catch {
            return null;
        }
    }

    private static string? TryComputeStreamSha256(Stream stream) {
        if (stream == null || !stream.CanSeek) return null;
        long position;
        try {
            position = stream.Position;
        } catch {
            return null;
        }

        try {
            stream.Position = 0;
            var hash = ComputeSha256Hex(stream);
            stream.Position = position;
            return hash;
        } catch {
            try {
                stream.Position = position;
            } catch {
                // ignore
            }
            return null;
        }
    }

    private static string ComputeSha256Hex(string value) {
        using var sha = SHA256.Create();
        var bytes = Encoding.UTF8.GetBytes(value ?? string.Empty);
        var hash = sha.ComputeHash(bytes);
        return ConvertToHexLower(hash);
    }

    private static string ComputeSha256Hex(Stream stream) {
        using var sha = SHA256.Create();
        var hash = sha.ComputeHash(stream);
        return ConvertToHexLower(hash);
    }

    private static string ConvertToHexLower(byte[] bytes) {
        var sb = new StringBuilder(bytes.Length * 2);
        for (int i = 0; i < bytes.Length; i++) {
            sb.Append(bytes[i].ToString("x2", CultureInfo.InvariantCulture));
        }
        return sb.ToString();
    }

    private static string BuildSourceId(string sourceKey) {
        var normalized = sourceKey ?? string.Empty;
        if (IsWindows()) {
            normalized = normalized.ToLowerInvariant();
        }
        return "src:" + ComputeSha256Hex(normalized);
    }

    private static bool IsWindows() {
        return Path.DirectorySeparatorChar == '\\';
    }

    private static string NormalizePathForId(string path) {
        if (string.IsNullOrWhiteSpace(path)) return string.Empty;
        string full;
        try {
            full = Path.GetFullPath(path);
        } catch {
            full = path;
        }
        return full.Replace('\\', '/');
    }

    private static bool TryParseAtxHeading(string line, out int level, out string text) {
        level = 0;
        text = string.Empty;
        if (line == null) return false;

        int i = 0;
        while (i < line.Length && line[i] == '#') i++;
        if (i < 1 || i > 6) return false;
        if (i >= line.Length) return false;
        if (line[i] != ' ' && line[i] != '\t') return false;

        level = i;
        text = line.Substring(i).Trim();
        if (text.Length == 0) text = $"Heading {level}";
        return true;
    }

    private static void UpdateHeadingStack(List<(int Level, string Text)> stack, int level, string text) {
        if (level < 1) return;
        if (string.IsNullOrWhiteSpace(text)) text = $"Heading {level}";

        for (int i = stack.Count - 1; i >= 0; i--) {
            if (stack[i].Level >= level) stack.RemoveAt(i);
        }
        stack.Add((level, CollapseWhitespace(text)));
    }

    private static string? BuildHeadingPath(List<(int Level, string Text)> stack) {
        if (stack.Count == 0) return null;
        var sb = new StringBuilder();
        for (int i = 0; i < stack.Count; i++) {
            if (i > 0) sb.Append(" > ");
            sb.Append(stack[i].Text);
        }
        var s = sb.ToString().Trim();
        return s.Length == 0 ? null : s;
    }

    private static bool WouldExceed(ReaderOptions opt, StringBuilder current, string nextLine) {
        // +1 for newline to keep final chunk shape similar to file.
        int nextLen = nextLine?.Length ?? 0;
        int extra = (current.Length == 0 ? 0 : 1) + nextLen;
        return current.Length > 0 && (current.Length + extra) > opt.MaxChars;
    }

    private static void AppendLineCapped(ReaderOptions opt, StringBuilder sb, string line, List<string> warnings) {
        if (sb.Length > 0) sb.AppendLine();

        var s = line ?? string.Empty;
        // Hard-cap pathological single lines so callers don't accidentally ingest megabytes in one chunk.
        if (s.Length > opt.MaxChars) {
            s = s.Substring(0, opt.MaxChars) + " <!-- truncated -->";
            warnings.Add("A single line exceeded MaxChars and was truncated.");
        }
        sb.Append(s);
    }

    private static string CollapseWhitespace(string text) {
        if (string.IsNullOrEmpty(text)) return string.Empty;
        var sb = new StringBuilder(text.Length);
        bool prevWs = false;
        for (int i = 0; i < text.Length; i++) {
            char c = text[i];
            bool ws = char.IsWhiteSpace(c);
            if (ws) {
                if (!prevWs) sb.Append(' ');
                prevWs = true;
            } else {
                sb.Append(c);
                prevWs = false;
            }
        }
        return sb.ToString().Trim();
    }

    private static string BuildStableId(string kind, string fileName, int chunkIndex, int? blockIndex) {
        // Keep IDs short, stable and ASCII-only; do not leak full paths.
        var l = blockIndex.HasValue ? blockIndex.Value.ToString(CultureInfo.InvariantCulture) : "na";
        return $"{kind}:{fileName}:c{chunkIndex}:l{l}";
    }

    private static MemoryStream CopyToMemory(Stream stream, CancellationToken ct) {
        ct.ThrowIfCancellationRequested();
        var ms = new MemoryStream();
        var buffer = new byte[64 * 1024];
        int read;
        while ((read = stream.Read(buffer, 0, buffer.Length)) > 0) {
            ct.ThrowIfCancellationRequested();
            ms.Write(buffer, 0, read);
        }
        ms.Position = 0;
        return ms;
    }

    private static ReaderHandlerCapability CloneCapability(ReaderHandlerCapability capability) {
        return new ReaderHandlerCapability {
            Id = capability.Id,
            DisplayName = capability.DisplayName,
            Description = capability.Description,
            Kind = capability.Kind,
            Extensions = capability.Extensions.ToArray(),
            IsBuiltIn = capability.IsBuiltIn,
            SupportsPath = capability.SupportsPath,
            SupportsStream = capability.SupportsStream,
            SchemaId = capability.SchemaId,
            SchemaVersion = capability.SchemaVersion,
            DefaultMaxInputBytes = capability.DefaultMaxInputBytes,
            WarningBehavior = capability.WarningBehavior,
            DeterministicOutput = capability.DeterministicOutput
        };
    }

    private static ReaderHandlerRegistrarDescriptor CloneRegistrarDescriptor(ReaderHandlerRegistrarDescriptor descriptor) {
        return new ReaderHandlerRegistrarDescriptor {
            HandlerId = descriptor.HandlerId,
            AssemblyName = descriptor.AssemblyName,
            TypeName = descriptor.TypeName,
            MethodName = descriptor.MethodName
        };
    }

    private static ReaderHostBootstrapOptions NormalizeHostBootstrapOptions(ReaderHostBootstrapOptions? options) {
        if (options == null) {
            return new ReaderHostBootstrapOptions();
        }

        return new ReaderHostBootstrapOptions {
            ReplaceExistingHandlers = options.ReplaceExistingHandlers,
            IncludeBuiltInCapabilities = options.IncludeBuiltInCapabilities,
            IncludeCustomCapabilities = options.IncludeCustomCapabilities,
            IndentedManifestJson = options.IndentedManifestJson
        };
    }

    private static List<RegistrarCandidate> DiscoverHandlerRegistrarsCore(IEnumerable<Assembly> assemblies) {
        if (assemblies == null) throw new ArgumentNullException(nameof(assemblies));

        var candidates = new List<RegistrarCandidate>();
        var uniqueAssemblies = new Dictionary<string, Assembly>(StringComparer.Ordinal);
        foreach (var assembly in assemblies) {
            if (assembly == null) continue;
            var key = assembly.FullName ?? assembly.GetName().Name ?? assembly.ManifestModule.Name;
            if (!uniqueAssemblies.ContainsKey(key)) {
                uniqueAssemblies.Add(key, assembly);
            }
        }

        var dedupe = new HashSet<string>(StringComparer.Ordinal);
        foreach (var assembly in uniqueAssemblies.Values) {
            foreach (var type in EnumerateLoadableTypes(assembly)) {
                if (type == null) continue;
                if (!type.IsClass || !type.IsAbstract || !type.IsSealed) continue; // static class

                foreach (var method in type.GetMethods(BindingFlags.Public | BindingFlags.Static | BindingFlags.DeclaredOnly)) {
                    if (!IsRegistrarMethod(method, out var handlerId)) continue;

                    var descriptor = new ReaderHandlerRegistrarDescriptor {
                        HandlerId = handlerId,
                        AssemblyName = assembly.GetName().Name ?? string.Empty,
                        TypeName = type.FullName ?? type.Name,
                        MethodName = method.Name
                    };

                    var key = string.Concat(
                        descriptor.AssemblyName, "|",
                        descriptor.TypeName, "|",
                        descriptor.MethodName, "|",
                        descriptor.HandlerId);
                    if (!dedupe.Add(key)) continue;

                    candidates.Add(new RegistrarCandidate(method, descriptor));
                }
            }
        }

        candidates.Sort(static (a, b) => {
            int cmp = string.CompareOrdinal(a.Descriptor.HandlerId, b.Descriptor.HandlerId);
            if (cmp != 0) return cmp;
            cmp = string.CompareOrdinal(a.Descriptor.AssemblyName, b.Descriptor.AssemblyName);
            if (cmp != 0) return cmp;
            cmp = string.CompareOrdinal(a.Descriptor.TypeName, b.Descriptor.TypeName);
            if (cmp != 0) return cmp;
            return string.CompareOrdinal(a.Descriptor.MethodName, b.Descriptor.MethodName);
        });

        return candidates;
    }

    private static IEnumerable<Type> EnumerateLoadableTypes(Assembly assembly) {
        try {
            return assembly.GetTypes();
        } catch (ReflectionTypeLoadException ex) {
            return ex.Types.Where(static t => t != null)!;
        } catch {
            return Array.Empty<Type>();
        }
    }

    private static IReadOnlyList<Assembly> GetLoadedAssembliesByPrefix(string assemblyNamePrefix) {
        if (assemblyNamePrefix == null) throw new ArgumentNullException(nameof(assemblyNamePrefix));

        var prefix = assemblyNamePrefix.Trim();
        if (prefix.Length == 0) {
            throw new ArgumentException("Assembly name prefix cannot be empty.", nameof(assemblyNamePrefix));
        }

        return AppDomain.CurrentDomain.GetAssemblies()
            .Where(static assembly => !assembly.IsDynamic)
            .Where(assembly => (assembly.GetName().Name ?? string.Empty).StartsWith(prefix, StringComparison.Ordinal))
            .OrderBy(static assembly => assembly.GetName().Name ?? string.Empty, StringComparer.Ordinal)
            .ToArray();
    }

    private static bool IsRegistrarMethod(MethodInfo method, out string handlerId) {
        handlerId = string.Empty;
        if (method == null) return false;
        if (method.IsGenericMethodDefinition) return false;
        if (method.ReturnType != typeof(void)) return false;

        var attribute = method.GetCustomAttribute<ReaderHandlerRegistrarAttribute>(inherit: false);
        if (attribute == null) return false;

        handlerId = (attribute.HandlerId ?? string.Empty).Trim();
        if (handlerId.Length == 0) return false;

        bool hasReplaceExisting = false;
        foreach (var parameter in method.GetParameters()) {
            if (parameter.ParameterType == typeof(bool) &&
                string.Equals(parameter.Name, "replaceExisting", StringComparison.OrdinalIgnoreCase)) {
                hasReplaceExisting = true;
                continue;
            }

            if (!parameter.IsOptional) {
                return false;
            }
        }

        return hasReplaceExisting;
    }

    private static string NormalizeExtension(string? extension) {
        var value = extension ?? string.Empty;
        if (string.IsNullOrWhiteSpace(value)) return string.Empty;
        var ext = value.Trim();
        if (!ext.StartsWith(".", StringComparison.Ordinal)) {
            ext = "." + ext;
        }
        return ext.ToLowerInvariant();
    }

    private static List<string> NormalizeRegistrationExtensions(IReadOnlyList<string>? extensions) {
        var list = new List<string>();
        if (extensions == null) return list;

        var set = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
        foreach (var ext in extensions) {
            var normalized = NormalizeExtension(ext);
            if (normalized.Length == 0) continue;
            if (set.Add(normalized)) {
                list.Add(normalized);
            }
        }

        list.Sort(StringComparer.Ordinal);
        return list;
    }

    private static HashSet<string> BuildBuiltInExtensionSet() {
        var set = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
        foreach (var capability in BuiltInCapabilities) {
            foreach (var ext in capability.Extensions) {
                var normalized = NormalizeExtension(ext);
                if (normalized.Length == 0) continue;
                set.Add(normalized);
            }
        }
        return set;
    }

    private static bool TryResolveCustomHandlerByPath(string path, out CustomReaderHandler handler) {
        var ext = NormalizeExtension(TryGetExtension(path));
        return TryResolveCustomHandlerByExtension(ext, out handler);
    }

    private static bool TryResolveCustomHandlerBySourceName(string? sourceName, out CustomReaderHandler handler) {
        var ext = NormalizeExtension(TryGetExtension(sourceName ?? string.Empty));
        return TryResolveCustomHandlerByExtension(ext, out handler);
    }

    private static bool TryResolveCustomHandlerByExtension(string ext, out CustomReaderHandler handler) {
        handler = null!;
        if (string.IsNullOrWhiteSpace(ext)) return false;

        lock (HandlerRegistrySync) {
            if (!CustomHandlerIdByExtension.TryGetValue(ext, out var handlerId)) {
                return false;
            }
            if (!CustomHandlersById.TryGetValue(handlerId, out var resolved) || resolved == null) {
                return false;
            }
            handler = resolved;
            return true;
        }
    }

    private static bool RemoveCustomHandlerUnsafe(string handlerId) {
        if (!CustomHandlersById.TryGetValue(handlerId, out var existing)) return false;

        CustomHandlersById.Remove(handlerId);
        foreach (var ext in existing.Extensions) {
            if (CustomHandlerIdByExtension.TryGetValue(ext, out var current) &&
                string.Equals(current, handlerId, StringComparison.OrdinalIgnoreCase)) {
                CustomHandlerIdByExtension.Remove(ext);
            }
        }

        return true;
    }

    private static ReaderOptions NormalizeOptions(ReaderOptions? options) {
        // Avoid mutating a caller-provided options instance.
        var o = options;
        var clone = new ReaderOptions {
            MaxInputBytes = o?.MaxInputBytes,
            OpenXmlMaxCharactersInPart = o?.OpenXmlMaxCharactersInPart,
            MaxChars = o?.MaxChars ?? 8_000,
            MaxTableRows = o?.MaxTableRows ?? 200,
            IncludeWordFootnotes = o?.IncludeWordFootnotes ?? true,
            IncludePowerPointNotes = o?.IncludePowerPointNotes ?? true,
            ExcelHeadersInFirstRow = o?.ExcelHeadersInFirstRow ?? true,
            ExcelChunkRows = o?.ExcelChunkRows ?? 200,
            ExcelSheetName = o?.ExcelSheetName,
            ExcelA1Range = o?.ExcelA1Range,
            MarkdownChunkByHeadings = o?.MarkdownChunkByHeadings ?? true,
            ComputeHashes = o?.ComputeHashes ?? true
        };

        if (clone.MaxChars < 256) clone.MaxChars = 256;
        if (clone.MaxTableRows < 1) clone.MaxTableRows = 1;
        if (clone.ExcelChunkRows < 1) clone.ExcelChunkRows = 1;
        if (clone.OpenXmlMaxCharactersInPart.HasValue && clone.OpenXmlMaxCharactersInPart.Value < 1) clone.OpenXmlMaxCharactersInPart = null;

        return clone;
    }

    private static ReaderOptions CloneOptions(ReaderOptions options, bool? computeHashes = null) {
        if (options == null) throw new ArgumentNullException(nameof(options));
        return new ReaderOptions {
            MaxInputBytes = options.MaxInputBytes,
            OpenXmlMaxCharactersInPart = options.OpenXmlMaxCharactersInPart,
            MaxChars = options.MaxChars,
            MaxTableRows = options.MaxTableRows,
            IncludeWordFootnotes = options.IncludeWordFootnotes,
            IncludePowerPointNotes = options.IncludePowerPointNotes,
            ExcelHeadersInFirstRow = options.ExcelHeadersInFirstRow,
            ExcelChunkRows = options.ExcelChunkRows,
            ExcelSheetName = options.ExcelSheetName,
            ExcelA1Range = options.ExcelA1Range,
            MarkdownChunkByHeadings = options.MarkdownChunkByHeadings,
            ComputeHashes = computeHashes ?? options.ComputeHashes
        };
    }

    private static ReaderFolderOptions NormalizeFolderOptions(ReaderFolderOptions? folderOptions) {
        var o = folderOptions;
        var clone = new ReaderFolderOptions {
            Recurse = o?.Recurse ?? true,
            MaxFiles = o?.MaxFiles ?? 500,
            MaxTotalBytes = o?.MaxTotalBytes,
            Extensions = (o?.Extensions == null || o.Extensions.Count == 0) ? null : o.Extensions.ToArray(),
            SkipReparsePoints = o?.SkipReparsePoints ?? true,
            DeterministicOrder = o?.DeterministicOrder ?? true
        };

        if (clone.MaxFiles < 1) clone.MaxFiles = 1;
        if (clone.MaxTotalBytes.HasValue && clone.MaxTotalBytes.Value < 1) clone.MaxTotalBytes = 1;
        return clone;
    }

    private static HashSet<string> NormalizeExtensions(IReadOnlyList<string>? configuredExtensions) {
        var allowedExt = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
        var source = (configuredExtensions == null || configuredExtensions.Count == 0)
            ? DefaultFolderExtensions
            : configuredExtensions;

        foreach (var e in source) {
            if (string.IsNullOrWhiteSpace(e)) continue;
            var normalized = e.StartsWith(".", StringComparison.Ordinal) ? e.Trim() : "." + e.Trim();
            if (normalized.Length > 1) allowedExt.Add(normalized);
        }

        return allowedExt;
    }

    private static IEnumerable<string> EnumerateFilesSafeDeterministic(string folderPath, ReaderFolderOptions options, CancellationToken cancellationToken) {
        var dirs = new Queue<string>();
        dirs.Enqueue(folderPath);

        while (dirs.Count > 0) {
            cancellationToken.ThrowIfCancellationRequested();
            var dir = dirs.Dequeue();

            IEnumerable<string> entries;
            try {
                entries = Directory.EnumerateFileSystemEntries(dir);
            } catch {
                // Best-effort traversal: unreadable directories are ignored.
                continue;
            }

            var ordered = options.DeterministicOrder
                ? entries.OrderBy(static x => x, StringComparer.Ordinal).ToArray()
                : entries.ToArray();

            foreach (var entry in ordered) {
                cancellationToken.ThrowIfCancellationRequested();

                FileAttributes attrs;
                try {
                    attrs = File.GetAttributes(entry);
                } catch {
                    continue;
                }

                var isDirectory = (attrs & FileAttributes.Directory) == FileAttributes.Directory;
                if (isDirectory) {
                    if (!options.Recurse) continue;

                    if (options.SkipReparsePoints && (attrs & FileAttributes.ReparsePoint) == FileAttributes.ReparsePoint) {
                        continue;
                    }

                    dirs.Enqueue(entry);
                    continue;
                }

                yield return entry;
            }
        }
    }

    private static OpenSettings? CreateOpenSettings(ReaderOptions opt) {
        if (opt == null) return null;
        if (!opt.OpenXmlMaxCharactersInPart.HasValue) return null;
        return new OpenSettings {
            MaxCharactersInPart = opt.OpenXmlMaxCharactersInPart.Value
        };
    }

    private static void EnforceFileSize(string path, long? maxBytes) {
        if (!maxBytes.HasValue) return;
        try {
            var fi = new FileInfo(path);
            if (fi.Length > maxBytes.Value) {
                throw new IOException($"Input exceeds MaxInputBytes ({fi.Length.ToString(CultureInfo.InvariantCulture)} > {maxBytes.Value.ToString(CultureInfo.InvariantCulture)}).");
            }
        } catch (IOException) {
            throw;
        } catch {
            // If we can't stat, don't block reads.
        }
    }

    private static void EnforceStreamSize(Stream stream, long? maxBytes) {
        if (!maxBytes.HasValue) return;
        if (!stream.CanSeek) return;
        try {
            if (stream.Length > maxBytes.Value) {
                throw new IOException($"Input exceeds MaxInputBytes ({stream.Length.ToString(CultureInfo.InvariantCulture)} > {maxBytes.Value.ToString(CultureInfo.InvariantCulture)}).");
            }
        } catch (NotSupportedException) {
            // ignore
        }
    }

    private static string ReadAllText(Stream stream, CancellationToken ct) {
        ct.ThrowIfCancellationRequested();
        using var reader = new StreamReader(stream, Encoding.UTF8, detectEncodingFromByteOrderMarks: true, bufferSize: 16 * 1024, leaveOpen: true);
        var sb = new StringBuilder();
        var buffer = new char[16 * 1024];
        const int HardCapChars = 50_000_000; // Defensive: avoid runaway memory usage on huge "text" streams.
        int read;
        while ((read = reader.Read(buffer, 0, buffer.Length)) > 0) {
            ct.ThrowIfCancellationRequested();
            sb.Append(buffer, 0, read);
            if (sb.Length >= HardCapChars) break;
        }
        return sb.ToString();
    }

    private sealed class RegistrarCandidate {
        public RegistrarCandidate(MethodInfo method, ReaderHandlerRegistrarDescriptor descriptor) {
            Method = method ?? throw new ArgumentNullException(nameof(method));
            Descriptor = descriptor ?? throw new ArgumentNullException(nameof(descriptor));
        }

        public MethodInfo Method { get; }
        public ReaderHandlerRegistrarDescriptor Descriptor { get; }
    }

    private sealed class CustomReaderHandler {
        public CustomReaderHandler(
            string id,
            string displayName,
            string? description,
            ReaderInputKind kind,
            IReadOnlyList<string> extensions,
            long? defaultMaxInputBytes,
            ReaderWarningBehavior warningBehavior,
            bool deterministicOutput,
            Func<string, ReaderOptions, CancellationToken, IEnumerable<ReaderChunk>>? readPath,
            Func<Stream, string?, ReaderOptions, CancellationToken, IEnumerable<ReaderChunk>>? readStream) {
            Id = id;
            DisplayName = displayName;
            Description = description;
            Kind = kind;
            Extensions = extensions;
            DefaultMaxInputBytes = defaultMaxInputBytes;
            WarningBehavior = warningBehavior;
            DeterministicOutput = deterministicOutput;
            ReadPath = readPath;
            ReadStream = readStream;
        }

        public string Id { get; }
        public string DisplayName { get; }
        public string? Description { get; }
        public ReaderInputKind Kind { get; }
        public IReadOnlyList<string> Extensions { get; }
        public long? DefaultMaxInputBytes { get; }
        public ReaderWarningBehavior WarningBehavior { get; }
        public bool DeterministicOutput { get; }
        public Func<string, ReaderOptions, CancellationToken, IEnumerable<ReaderChunk>>? ReadPath { get; }
        public Func<Stream, string?, ReaderOptions, CancellationToken, IEnumerable<ReaderChunk>>? ReadStream { get; }

        public ReaderHandlerCapability ToCapability() {
            return new ReaderHandlerCapability {
                Id = Id,
                DisplayName = DisplayName,
                Description = Description,
                Kind = Kind,
                Extensions = Extensions.ToArray(),
                IsBuiltIn = false,
                SupportsPath = ReadPath != null,
                SupportsStream = ReadStream != null,
                SchemaId = ReaderCapabilitySchema.Id,
                SchemaVersion = ReaderCapabilitySchema.Version,
                DefaultMaxInputBytes = DefaultMaxInputBytes,
                WarningBehavior = WarningBehavior,
                DeterministicOutput = DeterministicOutput
            };
        }
    }

    private sealed class FolderIngestState {
        public int FilesScanned { get; set; }
        public int FilesParsed { get; set; }
        public int FilesSkipped { get; set; }
        public long BytesRead { get; set; }
        public int ChunksProduced { get; set; }
    }

    private sealed class SourceInfo {
        public string Path { get; set; } = string.Empty;
        public string SourceId { get; set; } = string.Empty;
        public string? SourceHash { get; set; }
        public DateTime? LastWriteUtc { get; set; }
        public long? LengthBytes { get; set; }
    }
}
