using System;
using System.Collections.Generic;
using System.IO;
using System.Threading;
using OfficeIMO.Core.Internal;
using OfficeIMO.Core.Internal.Cabinet;

namespace OfficeIMO.Archives;

/// <summary>An immutable passive inventory with serialized, bounded member reads. Dispose releases its input and decoder indexes.</summary>
public sealed partial class ArchiveInspection : IDisposable {
    private readonly object _gate = new();
    private readonly ArchiveLimits _limits;
    private readonly int _directoryRecords;
    private readonly Dictionary<string, OfficeCabinetArchiveReader.Inspection> _cabinets = new(StringComparer.Ordinal);
    private MsiPackageReader? _database;
    private MemoryStream? _databaseSource;
    private long _readBytes;
    private bool _disposed;

    internal ArchiveInspection(byte[] input, ArchiveLimits limits, int directoryRecords, CancellationToken token) {
        _limits = limits; _directoryRecords = directoryRecords; InputLength = input.LongLength;
        using var source = new MemoryStream(input, writable: false);
        Format = ArchiveExtractor.Detect(source, token);
        StorageStreams = Array.Empty<ArchiveStorageStream>();
        try {
            var budget = new OfficeCabinetDirectoryBudget(directoryRecords);
            if (Format == ArchiveFormat.Cab) {
                var cabinet = OfficeCabinetArchiveReader.Inspect(input, limits.Expanded, limits.Entry, limits.Entries, budget, token);
                _cabinets.Add(string.Empty, cabinet);
                var entries = new List<ArchiveInspectionEntry>(cabinet.Members.Count);
                foreach (OfficeCabinetMemberInfo member in cabinet.Members) {
                    token.ThrowIfCancellationRequested();
                    entries.Add(Entry(member, member.Ordinal, null, null, null, null));
                }
                Entries = entries.AsReadOnly();
            } else Entries = ReadMsiInventory(input, budget, directoryRecords, token);
            token.ThrowIfCancellationRequested();
        } catch (OfficeCabinetException exception) { Dispose(); throw _limits.Translate(exception, directoryMaximum: _directoryRecords); }
        catch { Dispose(); throw; }
    }

    /// <summary>The detected container format.</summary>
    public ArchiveFormat Format { get; }
    /// <summary>Complete input snapshot size.</summary>
    public long InputLength { get; }
    /// <summary>Authored CAB records or MSI File rows in source order, including aliases and unsupported source profiles.</summary>
    public IReadOnlyList<ArchiveInspectionEntry> Entries { get; }
    /// <summary>Inert raw MSI storage streams in compound directory traversal order; empty for CAB.</summary>
    public IReadOnlyList<ArchiveStorageStream> StorageStreams { get; private set; }
    /// <summary>Cumulative charged expanded folder and raw stream bytes, including attempted reads that fail.</summary>
    public long ReadBytes { get { lock (_gate) return _readBytes; } }
    /// <summary>Cumulative CAB frame-processing work charged across member reads, including failed attempts after charging.
    /// Units bound frame setup and processing rather than exact CPU time or expanded payload size.</summary>
    public long DecompressionWorkBytes { get { lock (_gate) return _limits.DecompressionWork.ConsumedBytes; } }

    /// <summary>Returns a read-only stream containing one complete length-validated member. Every nonzero CAB frame checksum
    /// in its containing folder is checked; absent checksums do not establish authenticity. The source ordinal selects duplicate names
    /// independently. The entire containing CAB folder is decoded and charged before returning; no prefix or partial bytes escape.
    /// The returned stream is caller-owned and remains valid after this inspection is disposed.</summary>
    /// <exception cref="NotSupportedException">The entry's declared source or compression cannot be read.</exception>
    /// <exception cref="OfficeInspectionResourceLimitException">A configured or fixed managed admission ceiling is exhausted.</exception>
    /// <exception cref="InvalidDataException">Decoded content is corrupt.</exception>
    public Stream OpenRead(int ordinal, CancellationToken cancellationToken = default) {
        lock (_gate) {
            ThrowIfDisposed(); cancellationToken.ThrowIfCancellationRequested();
            if (ordinal < 0 || ordinal >= Entries.Count) throw new ArgumentOutOfRangeException(nameof(ordinal));
            ArchiveInspectionEntry entry = Entries[ordinal];
            if (entry.Diagnostic != null) throw new NotSupportedException(entry.Diagnostic.Code + ": " + entry.Diagnostic.Message);
            string key = entry.CabinetName == null ? string.Empty : entry.CabinetName.Substring(1);
            OfficeCabinetArchiveReader.Inspection cabinet = _cabinets[key];
            long bytes = entry.FolderExpandedBytes!.Value;
            long remaining = _limits.Expanded - _readBytes;
            // Reserve the full attempted decode, including gaps and sibling files, before invoking the decoder.
            ArchiveLimits.Add(ref _readBytes, bytes, _limits.Expanded, "MaxExpandedBytes");
            try {
                byte[] payload = cabinet.ReadMember(entry.CabinetOrdinal!.Value, remaining, _limits.DecompressionWork, cancellationToken);
                if (payload.LongLength != entry.Length) throw new InvalidDataException("Member length differs from its declared complete size.");
                return new MemoryStream(payload, writable: false);
            } catch (OfficeCabinetException exception) { throw _limits.Translate(exception, remaining, directoryMaximum: _directoryRecords); }
        }
    }

    /// <summary>Returns all bytes of one inert raw MSI stream, without interpreting or executing its content.
    /// Per-member and cumulative read limits apply, and the returned caller-owned stream survives inspection disposal.</summary>
    public Stream OpenStorageStream(int ordinal, CancellationToken cancellationToken = default) {
        lock (_gate) {
            ThrowIfDisposed(); cancellationToken.ThrowIfCancellationRequested();
            if (ordinal < 0 || ordinal >= StorageStreams.Count) throw new ArgumentOutOfRangeException(nameof(ordinal));
            ArchiveStorageStream entry = StorageStreams[ordinal];
            _limits.CheckEntry(entry.Length);
            ArchiveLimits.Add(ref _readBytes, entry.Length, _limits.Expanded, "MaxExpandedBytes");
            using Stream source = _database!.OpenRawStream(entry.RawName, cancellationToken);
            ArchiveLimits.Check(entry.Length, Math.Min(_limits.Entry, int.MaxValue), "MaxEntryBytes");
            byte[] payload = OfficeArchiveSafety.ReadEntryBytes(source, entry.Length, _limits.Entry, cancellationToken);
            if (payload.LongLength != entry.Length) throw new InvalidDataException("MSI raw stream length differs from its declared complete size.");
            cancellationToken.ThrowIfCancellationRequested();
            return new MemoryStream(payload, writable: false);
        }
    }

    /// <summary>Releases the input snapshot and indexes. Already returned member streams and immutable inventory metadata remain usable.</summary>
    public void Dispose() {
        lock (_gate) {
            if (_disposed) return;
            _disposed = true; _database?.Dispose(); _databaseSource?.Dispose();
            _database = null; _databaseSource = null; _cabinets.Clear();
        }
    }
    private void ThrowIfDisposed() { if (_disposed) throw new ObjectDisposedException(nameof(ArchiveInspection)); }
    private static ArchiveInspectionEntry Entry(OfficeCabinetMemberInfo member, int ordinal, string? rawFileName,
        string? path, string? cabinet, int? sequence) {
        int method = member.Compression & 15;
        ArchiveCompression compression = method == 0 ? ArchiveCompression.Store : method == 1 ? ArchiveCompression.MsZip
            : method == 2 ? ArchiveCompression.Quantum : method == 3 ? ArchiveCompression.Lzx : ArchiveCompression.Unknown;
        var diagnostic = member.UnsupportedCode == null ? null : new ArchiveInspectionDiagnostic(member.UnsupportedCode,
            "The declared CAB folder compression or its parameters are unsupported.");
        return new ArchiveInspectionEntry(ordinal, member.Name, rawFileName, path, member.Length, cabinet,
            member.Ordinal, sequence, compression, member.Compression, member.FolderOrdinal,
            member.FolderCompressedBytes, member.FolderExpandedBytes, member.FolderFrameCount, member.FolderChecksumFrameCount, diagnostic);
    }
}
