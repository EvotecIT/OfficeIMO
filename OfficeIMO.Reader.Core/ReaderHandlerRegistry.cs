using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Threading;
using System.Threading.Tasks;

namespace OfficeIMO.Reader;

internal sealed class ReaderHandlerRegistry {
    private readonly Dictionary<string, ReaderHandlerDescriptor> _handlersById = new Dictionary<string, ReaderHandlerDescriptor>(StringComparer.OrdinalIgnoreCase);
    private readonly Dictionary<string, string> _handlerIdByExtension = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);
    private readonly object _sync = new object();

    public void Register(ReaderHandlerRegistration registration, bool replaceExisting) {
        ReaderHandlerDescriptor handler = ReaderHandlerDescriptor.Create(registration);

        lock (_sync) {
            if (!replaceExisting) {
                ValidateNoConflicts(handler.Id, handler.Extensions);
            } else {
                RemoveConflicts(handler.Id, handler.Extensions);
            }

            _handlersById[handler.Id] = handler;
            foreach (string extension in handler.Extensions) {
                _handlerIdByExtension[extension] = handler.Id;
            }

        }
    }

    public ReaderHandlerRegistrySnapshot CaptureSnapshot() {
        lock (_sync) {
            return new ReaderHandlerRegistrySnapshot(
                new Dictionary<string, ReaderHandlerDescriptor>(_handlersById, StringComparer.OrdinalIgnoreCase),
                new Dictionary<string, string>(_handlerIdByExtension, StringComparer.OrdinalIgnoreCase));
        }
    }

    private void ValidateNoConflicts(string handlerId, IReadOnlyList<string> extensions) {
        if (_handlersById.ContainsKey(handlerId)) {
            throw new InvalidOperationException($"Handler '{handlerId}' is already registered.");
        }

        foreach (string extension in extensions) {
            if (_handlerIdByExtension.ContainsKey(extension)) {
                throw new InvalidOperationException($"Extension '{extension}' is already handled by a configured reader.");
            }
        }
    }

    private void RemoveConflicts(string handlerId, IReadOnlyList<string> extensions) {
        var toRemove = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
        if (_handlersById.ContainsKey(handlerId)) {
            toRemove.Add(handlerId);
        }

        foreach (string extension in extensions) {
            if (_handlerIdByExtension.TryGetValue(extension, out string? existingHandlerId)) {
                toRemove.Add(existingHandlerId);
            }
        }

        foreach (string existingHandlerId in toRemove) {
            RemoveUnsafe(existingHandlerId);
        }
    }

    private bool RemoveUnsafe(string handlerId) {
        if (!_handlersById.TryGetValue(handlerId, out ReaderHandlerDescriptor? existing)) {
            return false;
        }

        _handlersById.Remove(handlerId);
        foreach (string extension in existing.Extensions) {
            if (_handlerIdByExtension.TryGetValue(extension, out string? current) &&
                string.Equals(current, handlerId, StringComparison.OrdinalIgnoreCase)) {
                _handlerIdByExtension.Remove(extension);
            }
        }

        return true;
    }
}

internal sealed class ReaderHandlerRegistrySnapshot {
    private readonly IReadOnlyDictionary<string, ReaderHandlerDescriptor> _handlersById;
    private readonly IReadOnlyDictionary<string, string> _handlerIdByExtension;

    public ReaderHandlerRegistrySnapshot(
        IReadOnlyDictionary<string, ReaderHandlerDescriptor> handlersById,
        IReadOnlyDictionary<string, string> handlerIdByExtension) {
        _handlersById = handlersById;
        _handlerIdByExtension = handlerIdByExtension;
    }

    public IReadOnlyList<ReaderHandlerDescriptor> Handlers => _handlersById.Values
        .OrderBy(static handler => handler.Id, StringComparer.Ordinal)
        .ToArray();

    public IReadOnlyList<string> Extensions => _handlerIdByExtension.Keys
        .OrderBy(static extension => extension, StringComparer.Ordinal)
        .ToArray();

    public bool TryResolve(string extension, out ReaderHandlerDescriptor handler) {
        handler = null!;
        if (string.IsNullOrWhiteSpace(extension) ||
            !_handlerIdByExtension.TryGetValue(extension, out string? handlerId) ||
            !_handlersById.TryGetValue(handlerId, out ReaderHandlerDescriptor? resolved)) {
            return false;
        }

        handler = resolved;
        return true;
    }

    public bool TryResolveByKind(ReaderInputKind kind, bool pathInput, out ReaderHandlerDescriptor handler) {
        handler = null!;
        ReaderHandlerDescriptor? match = null;
        foreach (ReaderHandlerDescriptor candidate in _handlersById.Values.OrderBy(static item => item.Id, StringComparer.Ordinal)) {
            if (!candidate.UseDetectedKindFallback || candidate.Kind != kind) continue;
            bool supportsInput = pathInput ? candidate.SupportsPathInput : candidate.SupportsStreamInput;
            if (!supportsInput) continue;
            if (match != null) {
                return false;
            }
            match = candidate;
        }

        if (match == null) return false;
        handler = match;
        return true;
    }

    public bool TryProbeStream(
        Stream stream,
        string? sourceName,
        ReaderOptions options,
        CancellationToken cancellationToken,
        out ReaderHandlerDescriptor handler) {
        handler = null!;
        ReaderHandlerDescriptor? match = null;
        long position = stream.Position;
        foreach (ReaderHandlerDescriptor candidate in _handlersById.Values
                     .Where(static item => item.ProbeStream != null && item.SupportsStreamInput)
                     .OrderBy(static item => item.Id, StringComparer.Ordinal)) {
            cancellationToken.ThrowIfCancellationRequested();
            stream.Position = position;
            bool accepted = candidate.ProbeStream!(stream, sourceName, options, cancellationToken);
            if (!accepted) continue;
            if (match != null) {
                stream.Position = position;
                return false;
            }
            match = candidate;
        }
        stream.Position = position;
        if (match == null) return false;
        handler = match;
        return true;
    }
}

internal sealed class ReaderHandlerDescriptor {
    private ReaderHandlerDescriptor(
        string id,
        string displayName,
        string? description,
        ReaderHandlerOrigin origin,
        ReaderInputKind kind,
        bool useDetectedKindFallback,
        IReadOnlyList<string> extensions,
        long? defaultMaxInputBytes,
        ReaderSourceHashBehavior sourceHashBehavior,
        ReaderWarningBehavior warningBehavior,
        bool deterministicOutput,
        Func<string, ReaderOptions, CancellationToken, IEnumerable<ReaderChunk>>? readPath,
        Func<Stream, string?, ReaderOptions, CancellationToken, IEnumerable<ReaderChunk>>? readStream,
        Func<string, ReaderOptions, CancellationToken, OfficeDocumentReadResult>? readDocumentPath,
        Func<Stream, string?, ReaderOptions, CancellationToken, OfficeDocumentReadResult>? readDocumentStream,
        Func<string, ReaderOptions, CancellationToken, Task<OfficeDocumentReadResult>>? readDocumentPathAsync,
        Func<Stream, string?, ReaderOptions, CancellationToken, Task<OfficeDocumentReadResult>>? readDocumentStreamAsync,
        Func<Stream, string?, ReaderOptions, CancellationToken, bool>? probeStream) {
        Id = id;
        DisplayName = displayName;
        Description = description;
        Origin = origin;
        Kind = kind;
        UseDetectedKindFallback = useDetectedKindFallback;
        Extensions = extensions;
        DefaultMaxInputBytes = defaultMaxInputBytes;
        SourceHashBehavior = sourceHashBehavior;
        WarningBehavior = warningBehavior;
        DeterministicOutput = deterministicOutput;
        ReadPath = readPath;
        ReadStream = readStream;
        ReadDocumentPath = readDocumentPath;
        ReadDocumentStream = readDocumentStream;
        ReadDocumentPathAsync = readDocumentPathAsync;
        ReadDocumentStreamAsync = readDocumentStreamAsync;
        ProbeStream = probeStream;
    }

    public string Id { get; }
    public string DisplayName { get; }
    public string? Description { get; }
    public ReaderHandlerOrigin Origin { get; }
    public ReaderInputKind Kind { get; }
    public bool UseDetectedKindFallback { get; }
    public IReadOnlyList<string> Extensions { get; }
    public long? DefaultMaxInputBytes { get; }
    public ReaderSourceHashBehavior SourceHashBehavior { get; }
    public ReaderWarningBehavior WarningBehavior { get; }
    public bool DeterministicOutput { get; }
    public Func<string, ReaderOptions, CancellationToken, IEnumerable<ReaderChunk>>? ReadPath { get; }
    public Func<Stream, string?, ReaderOptions, CancellationToken, IEnumerable<ReaderChunk>>? ReadStream { get; }
    public Func<string, ReaderOptions, CancellationToken, OfficeDocumentReadResult>? ReadDocumentPath { get; }
    public Func<Stream, string?, ReaderOptions, CancellationToken, OfficeDocumentReadResult>? ReadDocumentStream { get; }
    public Func<string, ReaderOptions, CancellationToken, Task<OfficeDocumentReadResult>>? ReadDocumentPathAsync { get; }
    public Func<Stream, string?, ReaderOptions, CancellationToken, Task<OfficeDocumentReadResult>>? ReadDocumentStreamAsync { get; }
    public Func<Stream, string?, ReaderOptions, CancellationToken, bool>? ProbeStream { get; }
    public bool SupportsPathInput => ReadPath != null || ReadDocumentPath != null || ReadDocumentPathAsync != null;
    public bool SupportsStreamInput => ReadStream != null || ReadDocumentStream != null || ReadDocumentStreamAsync != null;

    public static ReaderHandlerDescriptor Create(ReaderHandlerRegistration registration) {
        if (registration == null) throw new ArgumentNullException(nameof(registration));

        string id = (registration.Id ?? string.Empty).Trim();
        if (id.Length == 0) throw new ArgumentException("Handler Id cannot be empty.", nameof(registration));
        if (registration.ReadPath == null &&
            registration.ReadStream == null &&
            registration.ReadDocumentPath == null &&
            registration.ReadDocumentStream == null &&
            registration.ReadDocumentPathAsync == null &&
            registration.ReadDocumentStreamAsync == null) {
            throw new ArgumentException(
                "Handler must define a synchronous or asynchronous reader for paths and/or streams.",
                nameof(registration));
        }

        if (registration.DefaultMaxInputBytes.HasValue && registration.DefaultMaxInputBytes.Value < 1) {
            throw new ArgumentException("DefaultMaxInputBytes must be greater than 0 when specified.", nameof(registration));
        }

        IReadOnlyList<string> extensions = NormalizeExtensions(registration.Extensions);
        if (extensions.Count == 0 && !registration.UseDetectedKindFallback) {
            throw new ArgumentException("Handler must define at least one extension unless it is a detected-kind fallback.", nameof(registration));
        }

        return new ReaderHandlerDescriptor(
            id,
            string.IsNullOrWhiteSpace(registration.DisplayName) ? id : registration.DisplayName!.Trim(),
            registration.Description,
            registration.Origin,
            registration.Kind,
            registration.UseDetectedKindFallback,
            extensions,
            registration.DefaultMaxInputBytes,
            registration.SourceHashBehavior,
            registration.WarningBehavior,
            registration.DeterministicOutput,
            registration.ReadPath,
            registration.ReadStream,
            registration.ReadDocumentPath,
            registration.ReadDocumentStream,
            registration.ReadDocumentPathAsync,
            registration.ReadDocumentStreamAsync,
            registration.ProbeStream);
    }

    public ReaderHandlerCapability ToCapability() {
        return new ReaderHandlerCapability {
            Id = Id,
            DisplayName = DisplayName,
            Description = Description,
            Origin = Origin,
            Kind = Kind,
            Extensions = Extensions.ToArray(),
            SupportsPath = SupportsPathInput,
            SupportsStream = SupportsStreamInput,
            SupportsDocumentPath = ReadDocumentPath != null || ReadDocumentPathAsync != null,
            SupportsDocumentStream = ReadDocumentStream != null || ReadDocumentStreamAsync != null,
            SupportsAsyncPath = ReadDocumentPathAsync != null,
            SupportsAsyncStream = ReadDocumentStreamAsync != null,
            SchemaId = ReaderCapabilitySchema.Id,
            SchemaVersion = ReaderCapabilitySchema.Version,
            DefaultMaxInputBytes = DefaultMaxInputBytes,
            WarningBehavior = WarningBehavior,
            DeterministicOutput = DeterministicOutput
        };
    }

    private static IReadOnlyList<string> NormalizeExtensions(IReadOnlyList<string>? extensions) {
        var normalized = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
        if (extensions != null) {
            foreach (string extension in extensions) {
                string value = DocumentReaderEngine.NormalizeExtension(extension);
                if (value.Length > 0) {
                    normalized.Add(value);
                }
            }
        }

        return normalized.OrderBy(static extension => extension, StringComparer.Ordinal).ToArray();
    }
}
