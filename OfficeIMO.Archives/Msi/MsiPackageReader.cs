using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Threading;
using OfficeIMO.Core.Internal;

namespace OfficeIMO.Archives;

/// <summary>Reads inert MSI metadata and embedded streams; never invokes Windows Installer.</summary>
internal sealed partial class MsiPackageReader : IDisposable {
    private static readonly Guid InstallerClass = new("000c1084-0000-0000-c000-000000000046");
    private static readonly char[] InvalidCompoundNameCharacters = { '/', '\\', ':', '!' };
    private readonly ArchiveLimits _limits;
    private readonly CancellationToken _token;
    private readonly OfficeCompoundFileReader.ReadSession _compound;
    private readonly Dictionary<string, string> _streams;
    private long _metadataBytes;
    private string?[] _strings = Array.Empty<string?>();
    private int _stringWidth;
    private Dictionary<string, MsiColumn[]> _schemas = new(StringComparer.OrdinalIgnoreCase);

    internal MsiPackageReader(Stream source, ArchiveLimits limits, CancellationToken token, int maxDirectoryEntries = 65536) {
        _limits = limits; _token = token;
        var compoundOptions = new OfficeCompoundReadOptions(maxDirectoryEntries: maxDirectoryEntries, maxStreamBytes: limits.Input, maxTotalStreamBytes: limits.Input);
        source.Position = 0;
        _compound = new OfficeCompoundFileReader.ReadSession(source, compoundOptions, leaveOpen: true, token,
            requireCanonicalHeader: true);
        _streams = new Dictionary<string, string>(StringComparer.Ordinal);
        try {
            if (!_compound.HasCanonicalRootDirectory || _compound.HasInvalidDirectoryTree)
                throw new InvalidDataException("MSI compound directory has an invalid root, tree links or unreachable active entries.");
            if (_compound.RootEntry.ClassId != InstallerClass) throw new NotSupportedException("The compound container is not an ordinary Windows Installer database. Patches and transforms are unsupported.");
            foreach (OfficeCompoundFileEntry entry in _compound.Entries) {
                token.ThrowIfCancellationRequested();
                if (entry.Name.IndexOfAny(InvalidCompoundNameCharacters) >= 0)
                    throw new InvalidDataException("MSI compound entry has an invalid raw directory name.");
            }
            // Inspect storages before stream aliases: the compound reader may also
            // expose unqualified fallback aliases for streams inside a valid storage.
            foreach (OfficeCompoundFileEntry entry in _compound.Entries.Where(entry => entry.ObjectType == 1)) {
                token.ThrowIfCancellationRequested();
                if (entry.IsFallback)
                    throw new InvalidDataException("MSI contains unreachable compound entries.");
                throw new NotSupportedException("MSI embedded substorage, including customization transforms, is unsupported.");
            }
            foreach (OfficeCompoundFileEntry entry in _compound.Entries.Where(entry => entry.IsStream)) {
                token.ThrowIfCancellationRequested();
                if (entry.IsFallback || _streams.ContainsKey(entry.Path))
                    throw new InvalidDataException("MSI contains unreachable or ambiguous stream names.");
                _streams.Add(entry.Path, entry.Path);
            }
            _compound.ValidateAllocation();
            LoadStrings();
            LoadSchemas();
        } catch { _compound.Dispose(); throw; }
    }
    internal IReadOnlyList<OfficeCompoundFileEntry> StorageStreams => _compound.Entries.Where(entry => entry.IsStream).OrderBy(entry => entry.DirectoryOrder).ToArray();
    internal Stream OpenRawStream(string rawName, CancellationToken token) => _compound.OpenStream(rawName, token);
    internal Stream OpenCabinet(string name) => Open(name, table: false);
    internal long GetCabinetLength(string name) => _compound.GetStreamLength(StoredName(name, table: false));
    private Stream Open(string name, bool table = true) => _compound.OpenStream(StoredName(name, table));
    private string StoredName(string name, bool table) {
        _token.ThrowIfCancellationRequested();
        string encoded = EncodeStreamName(name, table);
        if (!_streams.TryGetValue(encoded, out string? stored)) throw new InvalidDataException("Required MSI stream is missing: " + name);
        return stored;
    }
    public void Dispose() => _compound.Dispose();
    private byte[] ReadMetadata(string name) {
        using Stream payload = Open(name);
        ArchiveLimits.Add(ref _metadataBytes, payload.Length, _limits.Metadata, "MSI metadata");
        return OfficeArchiveSafety.ReadEntryBytes(payload, payload.Length, _limits.Metadata, _token);
    }
    internal static string EncodeStreamName(string name, bool table) {
        const string alphabet = "0123456789ABCDEFGHIJKLMNOPQRSTUVWXYZabcdefghijklmnopqrstuvwxyz._";
        var result = new System.Text.StringBuilder();
        if (table) result.Append('\u4840');
        for (int index = 0; index < name.Length; index++) {
            int first = alphabet.IndexOf(name[index]);
            int next = index + 1 < name.Length ? alphabet.IndexOf(name[index + 1]) : -1;
            if (first < 0) result.Append(name[index]);
            else if (next < 0) result.Append((char)(0x4800 + first));
            else { result.Append((char)(0x3800 + first + (next << 6))); index++; }
        }
        if (result.Length > 31) throw new InvalidDataException("MSI encoded stream name exceeds compound-storage bounds.");
        return result.ToString();
    }
    private static ushort UInt16(byte[] bytes, int offset) {
        if (offset < 0 || offset > bytes.Length - 2) throw new InvalidDataException("MSI metadata is truncated.");
        return (ushort)(bytes[offset] | bytes[offset + 1] << 8);
    }
    private static uint UInt32(byte[] bytes, int offset) => (uint)(UInt16(bytes, offset) | UInt16(bytes, offset + 2) << 16);
    private sealed class MsiColumn {
        internal MsiColumn(string name, char kind, int position, int stringLength = 0) {
            Name = name; Kind = kind; Position = position; StringLength = stringLength;
        }
        internal string Name { get; }
        internal char Kind { get; }
        internal int Position { get; }
        internal int StringLength { get; }
    }
}
