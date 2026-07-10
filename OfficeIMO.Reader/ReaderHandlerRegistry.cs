using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Threading;

namespace OfficeIMO.Reader;

internal sealed class ReaderHandlerRegistry {
    private readonly HashSet<string> _builtInExtensions;
    private readonly Dictionary<string, ReaderHandlerDescriptor> _handlersById = new Dictionary<string, ReaderHandlerDescriptor>(StringComparer.OrdinalIgnoreCase);
    private readonly Dictionary<string, string> _handlerIdByExtension = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);
    private readonly object _sync = new object();

    public ReaderHandlerRegistry(IEnumerable<string> builtInExtensions) {
        _builtInExtensions = new HashSet<string>(builtInExtensions ?? Array.Empty<string>(), StringComparer.OrdinalIgnoreCase);
    }

    public IReadOnlyList<string> Register(ReaderHandlerRegistration registration, bool replaceExisting, bool preserveExistingCustomExtensions) {
        ReaderHandlerDescriptor handler = ReaderHandlerDescriptor.Create(registration);

        lock (_sync) {
            IReadOnlyList<string> effectiveExtensions = handler.Extensions;
            if (preserveExistingCustomExtensions) {
                var preserved = new List<string>(handler.Extensions.Count);
                foreach (string extension in handler.Extensions) {
                    if (_handlerIdByExtension.TryGetValue(extension, out string? existing) &&
                        !string.Equals(existing, handler.Id, StringComparison.OrdinalIgnoreCase)) {
                        continue;
                    }

                    preserved.Add(extension);
                }

                if (preserved.Count == 0) {
                    return Array.Empty<string>();
                }

                effectiveExtensions = preserved.ToArray();
                handler = handler.WithExtensions(effectiveExtensions);
            }

            if (!replaceExisting) {
                ValidateNoConflicts(handler.Id, effectiveExtensions);
            } else {
                RemoveConflicts(handler.Id, handler.Extensions, preserveExistingCustomExtensions);
            }

            _handlersById[handler.Id] = handler;
            foreach (string extension in handler.Extensions) {
                _handlerIdByExtension[extension] = handler.Id;
            }

            return handler.Extensions.ToArray();
        }
    }

    public bool Unregister(string handlerId) {
        if (string.IsNullOrWhiteSpace(handlerId)) {
            return false;
        }

        lock (_sync) {
            return RemoveUnsafe(handlerId.Trim());
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
            if (_builtInExtensions.Contains(extension)) {
                throw new InvalidOperationException($"Extension '{extension}' is handled by a built-in reader. Use replaceExisting=true to override.");
            }

            if (_handlerIdByExtension.ContainsKey(extension)) {
                throw new InvalidOperationException($"Extension '{extension}' is already handled by a custom reader.");
            }
        }
    }

    private void RemoveConflicts(string handlerId, IReadOnlyList<string> extensions, bool preserveExistingCustomExtensions) {
        var toRemove = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
        if (_handlersById.ContainsKey(handlerId)) {
            toRemove.Add(handlerId);
        }

        foreach (string extension in extensions) {
            if (preserveExistingCustomExtensions &&
                _handlerIdByExtension.TryGetValue(extension, out string? preservedHandlerId) &&
                !string.Equals(preservedHandlerId, handlerId, StringComparison.OrdinalIgnoreCase)) {
                continue;
            }

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
}

internal sealed class ReaderHandlerDescriptor {
    private ReaderHandlerDescriptor(
        string id,
        string displayName,
        string? description,
        ReaderInputKind kind,
        IReadOnlyList<string> extensions,
        long? defaultMaxInputBytes,
        ReaderWarningBehavior warningBehavior,
        bool deterministicOutput,
        Func<string, ReaderOptions, CancellationToken, IEnumerable<ReaderChunk>>? readPath,
        Func<Stream, string?, ReaderOptions, CancellationToken, IEnumerable<ReaderChunk>>? readStream,
        Func<string, ReaderOptions, CancellationToken, OfficeDocumentReadResult>? readDocumentPath,
        Func<Stream, string?, ReaderOptions, CancellationToken, OfficeDocumentReadResult>? readDocumentStream) {
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
        ReadDocumentPath = readDocumentPath;
        ReadDocumentStream = readDocumentStream;
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
    public Func<string, ReaderOptions, CancellationToken, OfficeDocumentReadResult>? ReadDocumentPath { get; }
    public Func<Stream, string?, ReaderOptions, CancellationToken, OfficeDocumentReadResult>? ReadDocumentStream { get; }

    public static ReaderHandlerDescriptor Create(ReaderHandlerRegistration registration) {
        if (registration == null) throw new ArgumentNullException(nameof(registration));

        string id = (registration.Id ?? string.Empty).Trim();
        if (id.Length == 0) throw new ArgumentException("Handler Id cannot be empty.", nameof(registration));
        if (registration.ReadPath == null &&
            registration.ReadStream == null &&
            registration.ReadDocumentPath == null &&
            registration.ReadDocumentStream == null) {
            throw new ArgumentException(
                "Handler must define a chunk reader or rich document reader for paths and/or streams.",
                nameof(registration));
        }

        if (registration.DefaultMaxInputBytes.HasValue && registration.DefaultMaxInputBytes.Value < 1) {
            throw new ArgumentException("DefaultMaxInputBytes must be greater than 0 when specified.", nameof(registration));
        }

        IReadOnlyList<string> extensions = NormalizeExtensions(registration.Extensions);
        if (extensions.Count == 0) {
            throw new ArgumentException("Handler must define at least one extension.", nameof(registration));
        }

        return new ReaderHandlerDescriptor(
            id,
            string.IsNullOrWhiteSpace(registration.DisplayName) ? id : registration.DisplayName!.Trim(),
            registration.Description,
            registration.Kind,
            extensions,
            registration.DefaultMaxInputBytes,
            registration.WarningBehavior,
            registration.DeterministicOutput,
            registration.ReadPath,
            registration.ReadStream,
            registration.ReadDocumentPath,
            registration.ReadDocumentStream);
    }

    public ReaderHandlerDescriptor WithExtensions(IReadOnlyList<string> extensions) {
        return new ReaderHandlerDescriptor(
            Id,
            DisplayName,
            Description,
            Kind,
            extensions.ToArray(),
            DefaultMaxInputBytes,
            WarningBehavior,
            DeterministicOutput,
            ReadPath,
            ReadStream,
            ReadDocumentPath,
            ReadDocumentStream);
    }

    public ReaderHandlerCapability ToCapability() {
        return new ReaderHandlerCapability {
            Id = Id,
            DisplayName = DisplayName,
            Description = Description,
            Kind = Kind,
            Extensions = Extensions.ToArray(),
            IsBuiltIn = false,
            SupportsPath = ReadPath != null || ReadDocumentPath != null,
            SupportsStream = ReadStream != null || ReadDocumentStream != null,
            SupportsDocumentPath = ReadDocumentPath != null,
            SupportsDocumentStream = ReadDocumentStream != null,
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
                string value = DocumentReader.NormalizeExtension(extension);
                if (value.Length > 0) {
                    normalized.Add(value);
                }
            }
        }

        return normalized.OrderBy(static extension => extension, StringComparer.Ordinal).ToArray();
    }
}
