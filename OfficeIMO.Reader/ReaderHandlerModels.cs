using OfficeIMO.Markdown;
using System;
using System.Collections.Generic;
using System.IO;
using System.Threading;
using System.Threading.Tasks;

namespace OfficeIMO.Reader;

/// <summary>
/// Custom handler registration model for configuring an <see cref="OfficeDocumentReaderBuilder"/>
/// without hard dependencies.
/// </summary>
public sealed class ReaderHandlerRegistration {
    /// <summary>
    /// Stable unique identifier for this handler (for example: "officeimo.reader.epub").
    /// </summary>
    public string Id { get; set; } = string.Empty;

    /// <summary>
    /// Optional display name shown in capability listings.
    /// </summary>
    public string? DisplayName { get; set; }

    /// <summary>
    /// Optional handler description shown in capability listings.
    /// </summary>
    public string? Description { get; set; }

    /// <summary>
    /// Kind advertised by this handler for detect/capability workflows.
    /// </summary>
    public ReaderInputKind Kind { get; set; } = ReaderInputKind.Unknown;

    /// <summary>
    /// File extensions handled by this registration (for example: ".epub", ".zip").
    /// </summary>
    public IReadOnlyList<string> Extensions { get; set; } = Array.Empty<string>();

    /// <summary>
    /// Path-based reader delegate.
    /// </summary>
    public Func<string, ReaderOptions, CancellationToken, IEnumerable<ReaderChunk>>? ReadPath { get; set; }

    /// <summary>
    /// Stream-based reader delegate.
    /// </summary>
    public Func<Stream, string?, ReaderOptions, CancellationToken, IEnumerable<ReaderChunk>>? ReadStream { get; set; }

    /// <summary>
    /// Optional path-based rich document reader delegate. When present,
    /// <see cref="OfficeDocumentReader.ReadDocument(string, ReaderOptions?, CancellationToken)"/>
    /// dispatches directly to this delegate instead of rebuilding a generic result from chunks.
    /// </summary>
    public Func<string, ReaderOptions, CancellationToken, OfficeDocumentReadResult>? ReadDocumentPath { get; set; }

    /// <summary>
    /// Optional stream-based rich document reader delegate. The delegate must not close the caller-owned stream.
    /// When present, <see cref="OfficeDocumentReader.ReadDocument(Stream, string?, ReaderOptions?, CancellationToken)"/>
    /// dispatches directly to this delegate instead of rebuilding a generic result from chunks.
    /// </summary>
    public Func<Stream, string?, ReaderOptions, CancellationToken, OfficeDocumentReadResult>? ReadDocumentStream { get; set; }

    /// <summary>
    /// Optional native asynchronous path-based rich document reader delegate.
    /// </summary>
    public Func<string, ReaderOptions, CancellationToken, Task<OfficeDocumentReadResult>>? ReadDocumentPathAsync { get; set; }

    /// <summary>
    /// Optional native asynchronous stream-based rich document reader delegate. The delegate must not close the caller-owned stream.
    /// </summary>
    public Func<Stream, string?, ReaderOptions, CancellationToken, Task<OfficeDocumentReadResult>>? ReadDocumentStreamAsync { get; set; }

    /// <summary>
    /// Optional advertised default max input bytes for this handler.
    /// Null means "no handler-specific default advertised".
    /// </summary>
    public long? DefaultMaxInputBytes { get; set; }

    /// <summary>
    /// Advertised warning model for this handler.
    /// </summary>
    public ReaderWarningBehavior WarningBehavior { get; set; } = ReaderWarningBehavior.Mixed;

    /// <summary>
    /// True when this handler advertises deterministic chunk ordering/output for identical input.
    /// </summary>
    public bool DeterministicOutput { get; set; } = true;
}

/// <summary>
/// Immutable capability descriptor for built-in and registered handlers.
/// </summary>
public sealed class ReaderHandlerCapability {
    /// <summary>
    /// Stable unique handler identifier.
    /// </summary>
    public string Id { get; set; } = string.Empty;

    /// <summary>
    /// Human-readable name.
    /// </summary>
    public string DisplayName { get; set; } = string.Empty;

    /// <summary>
    /// Optional handler description.
    /// </summary>
    public string? Description { get; set; }

    /// <summary>
    /// Advertised input kind.
    /// </summary>
    public ReaderInputKind Kind { get; set; }

    /// <summary>
    /// Extensions served by this handler.
    /// </summary>
    public IReadOnlyList<string> Extensions { get; set; } = Array.Empty<string>();

    /// <summary>
    /// True for built-in reader handlers.
    /// </summary>
    public bool IsBuiltIn { get; set; }

    /// <summary>
    /// True when path-based read delegate is available.
    /// </summary>
    public bool SupportsPath { get; set; }

    /// <summary>
    /// True when stream-based read delegate is available.
    /// </summary>
    public bool SupportsStream { get; set; }

    /// <summary>
    /// True when the handler supplies a native path-based <see cref="OfficeDocumentReadResult"/> projection.
    /// </summary>
    public bool SupportsDocumentPath { get; set; }

    /// <summary>
    /// True when the handler supplies a native stream-based <see cref="OfficeDocumentReadResult"/> projection.
    /// </summary>
    public bool SupportsDocumentStream { get; set; }

    /// <summary>
    /// True when the handler supplies a native asynchronous path reader.
    /// </summary>
    public bool SupportsAsyncPath { get; set; }

    /// <summary>
    /// True when the handler supplies a native asynchronous stream reader.
    /// </summary>
    public bool SupportsAsyncStream { get; set; }

    /// <summary>
    /// Capability schema identifier for host integration contracts.
    /// </summary>
    public string SchemaId { get; set; } = ReaderCapabilitySchema.Id;

    /// <summary>
    /// Capability schema version for host integration contracts.
    /// </summary>
    public int SchemaVersion { get; set; } = ReaderCapabilitySchema.Version;

    /// <summary>
    /// Optional advertised default max input bytes for this handler.
    /// Null means no handler-specific default is advertised.
    /// </summary>
    public long? DefaultMaxInputBytes { get; set; }

    /// <summary>
    /// Advertised warning model for this handler.
    /// </summary>
    public ReaderWarningBehavior WarningBehavior { get; set; } = ReaderWarningBehavior.Mixed;

    /// <summary>
    /// True when this handler advertises deterministic chunk ordering/output for identical input.
    /// </summary>
    public bool DeterministicOutput { get; set; } = true;
}

/// <summary>
/// Stable capability schema contract values exposed by the Reader capability APIs.
/// </summary>
public static class ReaderCapabilitySchema {
    /// <summary>
    /// Stable schema identifier.
    /// </summary>
    public const string Id = "officeimo.reader.capability";

    /// <summary>
    /// Current schema version.
    /// </summary>
    public const int Version = 3;
}

/// <summary>
/// Advertised warning behavior model for reader handlers.
/// </summary>
public enum ReaderWarningBehavior {
    /// <summary>
    /// Handler may both emit warning chunks and throw exceptions, depending on scenario.
    /// </summary>
    Mixed = 0,
    /// <summary>
    /// Handler prefers warning chunks over throwing for recoverable issues.
    /// </summary>
    WarningChunksOnly = 1,
    /// <summary>
    /// Handler prefers exception-based signaling for issues.
    /// </summary>
    ExceptionsOnly = 2
}

/// <summary>
/// Machine-readable capability manifest for host discovery/integration.
/// </summary>
public sealed class ReaderCapabilityManifest {
    /// <summary>
    /// Capability schema identifier.
    /// </summary>
    public string SchemaId { get; set; } = ReaderCapabilitySchema.Id;

    /// <summary>
    /// Capability schema version.
    /// </summary>
    public int SchemaVersion { get; set; } = ReaderCapabilitySchema.Version;

    /// <summary>
    /// Discovered handler capabilities included in this manifest.
    /// </summary>
    public IReadOnlyList<ReaderHandlerCapability> Handlers { get; set; } = Array.Empty<ReaderHandlerCapability>();
}
