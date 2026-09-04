using System;
using System.Collections.Generic;
using System.Linq;

namespace OfficeIMO.Ocr;

/// <summary>Describes an optional OCR provider without creating its runtime engine.</summary>
public sealed class OcrEngineDescriptor {
    internal OcrEngineDescriptor(string id, string displayName, OcrEngineCapabilities capabilities) {
        Id = id;
        DisplayName = displayName;
        Capabilities = capabilities.Clone();
    }

    /// <summary>Stable provider identifier selected by a host or CLI.</summary>
    public string Id { get; }
    /// <summary>Human-readable provider name.</summary>
    public string DisplayName { get; }
    /// <summary>Provider capability snapshot.</summary>
    public OcrEngineCapabilities Capabilities { get; }
}

/// <summary>Factory implemented by optional OCR packages or host integrations.</summary>
public interface IOcrEngineProvider {
    /// <summary>Stable provider identifier.</summary>
    string Id { get; }
    /// <summary>Human-readable provider name.</summary>
    string DisplayName { get; }
    /// <summary>Capabilities available from engines created by this provider.</summary>
    OcrEngineCapabilities Capabilities { get; }
    /// <summary>Creates one configured engine from bounded scalar options.</summary>
    IOcrEngine Create(IReadOnlyDictionary<string, string> options);
}

/// <summary>
/// Explicit registry for optional OCR providers. The core package performs no ambient scanning and carries no
/// provider runtime dependency; hosts decide which provider instances or assemblies are trusted.
/// </summary>
public sealed class OcrEngineCatalog {
    private const int MaximumProviders = 128;
    private const int MaximumOptions = 128;
    private const int MaximumOptionCharacters = 64 * 1024;
    private readonly object _sync = new object();
    private readonly Dictionary<string, IOcrEngineProvider> _providers =
        new Dictionary<string, IOcrEngineProvider>(StringComparer.OrdinalIgnoreCase);

    /// <summary>Registers a provider. Identifiers are unique without regard to case.</summary>
    public OcrEngineCatalog Register(IOcrEngineProvider provider) {
        if (provider == null) throw new ArgumentNullException(nameof(provider));
        string id = ValidateIdentity(provider.Id, nameof(provider));
        string displayName = ValidateIdentity(provider.DisplayName, nameof(provider));
        OcrEngineCapabilities capabilities = (provider.Capabilities ?? new OcrEngineCapabilities()).Clone();
        lock (_sync) {
            if (_providers.Count >= MaximumProviders && !_providers.ContainsKey(id)) {
                throw new InvalidOperationException("The OCR provider catalog cannot contain more than " + MaximumProviders + " providers.");
            }
            if (_providers.ContainsKey(id)) throw new ArgumentException("OCR provider '" + id + "' is already registered.", nameof(provider));
            _providers.Add(id, new ValidatedOcrEngineProvider(provider, id, displayName, capabilities));
        }
        return this;
    }

    /// <summary>Returns immutable descriptors ordered by stable identifier.</summary>
    public IReadOnlyList<OcrEngineDescriptor> Discover() {
        lock (_sync) {
            return _providers.Values
                .Select(static provider => new OcrEngineDescriptor(provider.Id, provider.DisplayName, provider.Capabilities))
                .OrderBy(static descriptor => descriptor.Id, StringComparer.OrdinalIgnoreCase)
                .ToArray();
        }
    }

    /// <summary>Creates a selected provider engine from caller-owned scalar options.</summary>
    public IOcrEngine Create(string id, IReadOnlyDictionary<string, string>? options = null) {
        string normalizedId = ValidateIdentity(id, nameof(id));
        IOcrEngineProvider provider;
        lock (_sync) {
            if (!_providers.TryGetValue(normalizedId, out provider!)) {
                throw new KeyNotFoundException("OCR provider '" + normalizedId + "' is not registered.");
            }
        }
        IReadOnlyDictionary<string, string> snapshot = SnapshotOptions(options);
        IOcrEngine engine = provider.Create(snapshot) ?? throw new InvalidOperationException("OCR provider '" + normalizedId + "' returned a null engine.");
        string engineId = OcrEngineRunner.GetValidatedEngineId(engine);
        if (!string.Equals(engineId, provider.Id, StringComparison.OrdinalIgnoreCase)) {
            throw new InvalidOperationException("OCR provider '" + provider.Id + "' created engine '" + engineId + "'. Provider and engine identifiers must match.");
        }
        return engine;
    }

    private static IReadOnlyDictionary<string, string> SnapshotOptions(IReadOnlyDictionary<string, string>? options) {
        if (options == null || options.Count == 0) return new Dictionary<string, string>(StringComparer.Ordinal);
        if (options.Count > MaximumOptions) throw new ArgumentException("OCR provider options cannot exceed " + MaximumOptions + " entries.", nameof(options));
        long characters = 0;
        var snapshot = new Dictionary<string, string>(StringComparer.Ordinal);
        foreach (KeyValuePair<string, string> option in options) {
            if (string.IsNullOrWhiteSpace(option.Key)) throw new ArgumentException("OCR provider option keys cannot be empty.", nameof(options));
            if (option.Value == null) throw new ArgumentException("OCR provider option values cannot be null.", nameof(options));
            characters += (long)option.Key.Length + option.Value.Length;
            if (characters > MaximumOptionCharacters) throw new ArgumentException("OCR provider options exceed the aggregate character limit.", nameof(options));
            if (snapshot.ContainsKey(option.Key)) throw new ArgumentException("OCR provider option keys must be unique.", nameof(options));
            snapshot.Add(option.Key, option.Value);
        }
        return snapshot;
    }

    private static string ValidateIdentity(string? value, string parameterName) {
        if (string.IsNullOrWhiteSpace(value)) throw new ArgumentException("OCR provider identity cannot be empty.", parameterName);
        if (value!.Length > OcrEngineRunner.MaximumEngineIdCharacters) throw new ArgumentException("OCR provider identity is too long.", parameterName);
        return value.Trim();
    }

    private sealed class ValidatedOcrEngineProvider : IOcrEngineProvider {
        private readonly IOcrEngineProvider _inner;
        internal ValidatedOcrEngineProvider(IOcrEngineProvider inner, string id, string displayName, OcrEngineCapabilities capabilities) {
            _inner = inner; Id = id; DisplayName = displayName; Capabilities = capabilities;
        }
        public string Id { get; }
        public string DisplayName { get; }
        public OcrEngineCapabilities Capabilities { get; }
        public IOcrEngine Create(IReadOnlyDictionary<string, string> options) => _inner.Create(options);
    }
}
