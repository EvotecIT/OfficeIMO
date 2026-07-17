using OfficeIMO.Email;
using System.Security.Cryptography;

namespace OfficeIMO.Email.Store;

internal sealed partial class PstStoreWriterCore {
    private static readonly byte[] CheckpointMagic = Encoding.ASCII.GetBytes("OIMOPSTC");
    private const int MinimumCheckpointVersion = 1;
    private const int CheckpointVersion = 2;
    private const long MaxCheckpointPayloadBytes = 128L * 1024 * 1024;

    internal string? CheckpointPath => _options.CheckpointPath;

    internal static PstStoreWriterCore Resume(string checkpointPath,
        IProgress<EmailStorePstWriteProgress>? progress) {
        string fullCheckpointPath = Path.GetFullPath(checkpointPath);
        WriterCheckpointState state = ReadCheckpointFile(fullCheckpointPath);
        state.ValidateOwnership(fullCheckpointPath);
        CleanupCheckpointCommitFiles(fullCheckpointPath);
        EmailStorePstWriterOptions options = state.CreateOptions(
            fullCheckpointPath, progress);
        if (File.Exists(state.DestinationPath) && !options.OverwriteExisting) {
            throw new IOException("The checkpoint destination now exists and overwrite is disabled.");
        }
        return new PstStoreWriterCore(state, options);
    }

    internal void Checkpoint() {
        ThrowIfUnavailable();
        CheckpointCore();
    }

    internal void Abandon() {
        ThrowIfUnavailable();
        _abandon = true;
    }

    internal static void DeleteCheckpoint(string checkpointPath) {
        if (string.IsNullOrWhiteSpace(checkpointPath)) {
            throw new ArgumentException("A checkpoint path is required.", nameof(checkpointPath));
        }
        string fullPath = Path.GetFullPath(checkpointPath);
        if (!File.Exists(fullPath)) return;
        WriterCheckpointState state = ReadCheckpointFile(fullPath);
        state.ValidateOwnership(fullPath);
        CleanupWorkingFiles(state.TemporaryPath, state.DestinationPath);
        CleanupCheckpointCommitFiles(fullPath);
        TryDelete(fullPath);
    }

    private void CheckpointCore() {
        string? checkpointPath = _options.CheckpointPath;
        if (checkpointPath == null) {
            throw new InvalidOperationException("This PST writer was not configured with a checkpoint path.");
        }
        ReportProgress(EmailStorePstWriteStage.Checkpointing);
        PstWriterFileCheckpoint fileState = _file.CaptureCheckpoint();
        _nodes.Flush(durable: true);
        _items.Flush(durable: true);

        string fullPath = Path.GetFullPath(checkpointPath);
        string? directory = Path.GetDirectoryName(fullPath);
        if (string.IsNullOrEmpty(directory)) directory = Directory.GetCurrentDirectory();
        Directory.CreateDirectory(directory);
        CleanupCheckpointCommitFiles(fullPath);
        string temporary = Path.Combine(directory, string.Concat(".", Path.GetFileName(fullPath),
            ".", Guid.NewGuid().ToString("N"), ".tmp"));
        try {
            using (var stream = new FileStream(temporary, FileMode.CreateNew, FileAccess.ReadWrite,
                FileShare.Read, 64 * 1024, FileOptions.SequentialScan))
            using (var writer = new BinaryWriter(stream, Encoding.UTF8, leaveOpen: true)) {
                writer.Write(CheckpointMagic);
                writer.Write(CheckpointVersion);
                long lengthPosition = stream.Position;
                writer.Write(0L);
                writer.Flush();
                long payloadPosition = stream.Position;
                using (var bounded = new EmailBoundedWriteStream(stream,
                           checked(payloadPosition + MaxCheckpointPayloadBytes)))
                using (var payloadWriter = new BinaryWriter(bounded, Encoding.UTF8, leaveOpen: true)) {
                    WriteCheckpointPayload(payloadWriter, fileState);
                    payloadWriter.Flush();
                }
                long payloadLength = checked(stream.Position - payloadPosition);
                byte[] digest = ComputeCheckpointDigest(stream, payloadPosition, payloadLength);
                stream.Position = lengthPosition;
                writer.Write(payloadLength);
                stream.Position = checked(payloadPosition + payloadLength);
                writer.Write(digest);
                writer.Flush();
                stream.Flush(flushToDisk: true);
                Array.Clear(digest, 0, digest.Length);
            }
            if (File.Exists(fullPath)) File.Replace(temporary, fullPath, null);
            else File.Move(temporary, fullPath);
        } finally {
            TryDelete(temporary);
        }
        _lastCheckpointItemCount = _itemCount;
        ReportProgress(EmailStorePstWriteStage.WritingItems);
    }

    private void WriteCheckpointPayload(BinaryWriter writer, PstWriterFileCheckpoint file) {
        writer.Write(_destinationPath);
        writer.Write(_temporaryPath);
        writer.Write(_options.DisplayName);
        writer.Write(_options.OverwriteExisting);
        writer.Write(_options.FailOnDataLoss);
        writer.Write(_options.MaxFolderCount);
        writer.Write(_options.MaxItemCount);
        writer.Write(_options.MaxNestedMessageDepth);
        writer.Write(_options.CheckpointIntervalItems);
        writer.Write(_options.MaxIndexRecordsInMemory);
        writer.Write(_options.RetainCheckpointOnDispose);
        writer.Write(_options.MaxDiagnostics);
        writer.Write(_providerUid.ToByteArray());
        writer.Write(_nextFolderIndex);
        writer.Write(_nextMessageIndex);
        writer.Write(_userFolderCount);
        writer.Write(_itemCount);
        writer.Write(file.StreamLength);
        writer.Write(file.NextOffset);
        writer.Write(file.NextBlockBid);
        writer.Write(file.NextPageBid);
        writer.Write(file.BlockCount);
        writer.Write((long)_nodes.Count);
        writer.Write((long)_items.Count);
        writer.Write(_items.PayloadLength);
        writer.Write(_diagnosticsTruncated);
        _namedProperties.WriteCheckpoint(writer);

        writer.Write(_folders.Count);
        foreach (FolderState folder in _folders.Values.OrderBy(item => item.Nid)) {
            writer.Write(folder.Nid);
            writer.Write(folder.ParentNid);
            writer.Write(folder.Name);
            WriteNullableString(writer, folder.ContainerClass);
            writer.Write(folder.IsSearchFolder);
            writer.Write((int)folder.SpecialFolderKind);
            writer.Write(folder.NormalItemCount);
            writer.Write(folder.AssociatedItemCount);
            writer.Write(folder.UnreadItemCount);
        }
        writer.Write(_diagnostics.Count);
        foreach (EmailStoreDiagnostic diagnostic in _diagnostics) {
            writer.Write(diagnostic.Code);
            writer.Write(diagnostic.Message);
            writer.Write((int)diagnostic.Severity);
            WriteNullableString(writer, diagnostic.Location);
        }
    }

    private static WriterCheckpointState ReadCheckpointFile(string checkpointPath) {
        string fullPath = Path.GetFullPath(checkpointPath);
        using (var stream = new FileStream(fullPath, FileMode.Open, FileAccess.Read,
            FileShare.Read, 64 * 1024, FileOptions.SequentialScan))
        using (var reader = new BinaryReader(stream, Encoding.UTF8, leaveOpen: false)) {
            byte[] magic = reader.ReadBytes(CheckpointMagic.Length);
            if (!magic.SequenceEqual(CheckpointMagic)) {
                throw new InvalidDataException("The file is not an OfficeIMO PST writer checkpoint.");
            }
            int version = reader.ReadInt32();
            if (version < MinimumCheckpointVersion || version > CheckpointVersion) {
                throw new NotSupportedException("The PST writer checkpoint version is not supported.");
            }
            long payloadLength = reader.ReadInt64();
            if (payloadLength < 0 || payloadLength > MaxCheckpointPayloadBytes) {
                throw new InvalidDataException("The PST writer checkpoint payload length is invalid.");
            }
            byte[] payload = reader.ReadBytes(checked((int)payloadLength));
            byte[] expected = reader.ReadBytes(32);
            if (payload.LongLength != payloadLength || expected.Length != 32 || stream.Position != stream.Length) {
                throw new InvalidDataException("The PST writer checkpoint is truncated.");
            }
            byte[] actual;
            using (SHA256 sha = SHA256.Create()) actual = sha.ComputeHash(payload);
            if (!FixedTimeEquals(expected, actual)) {
                throw new InvalidDataException("The PST writer checkpoint failed its integrity check.");
            }
            try {
                using (var payloadStream = new MemoryStream(payload, writable: false))
                using (var payloadReader = new BinaryReader(payloadStream, Encoding.UTF8, leaveOpen: false)) {
                    WriterCheckpointState state = ReadCheckpointPayload(payloadReader, version);
                    if (payloadStream.Position != payloadStream.Length) {
                        throw new InvalidDataException("The PST writer checkpoint contains trailing state.");
                    }
                    return state;
                }
            } finally {
                Array.Clear(payload, 0, payload.Length);
                Array.Clear(expected, 0, expected.Length);
                Array.Clear(actual, 0, actual.Length);
            }
        }
    }

    private static byte[] ComputeCheckpointDigest(Stream stream, long position, long length) {
        stream.Position = position;
        using (SHA256 sha = SHA256.Create()) {
            var buffer = new byte[64 * 1024];
            long remaining = length;
            while (remaining > 0) {
                int requested = (int)Math.Min(buffer.Length, remaining);
                int read = stream.Read(buffer, 0, requested);
                if (read == 0) throw new EndOfStreamException("The PST checkpoint payload is truncated.");
                sha.TransformBlock(buffer, 0, read, buffer, 0);
                remaining -= read;
            }
            sha.TransformFinalBlock(Array.Empty<byte>(), 0, 0);
            return sha.Hash ?? throw new InvalidDataException("The PST checkpoint digest could not be computed.");
        }
    }

    private static WriterCheckpointState ReadCheckpointPayload(BinaryReader reader, int version) {
        var state = new WriterCheckpointState {
            DestinationPath = Path.GetFullPath(reader.ReadString()),
            TemporaryPath = Path.GetFullPath(reader.ReadString()),
            DisplayName = reader.ReadString(),
            OverwriteExisting = reader.ReadBoolean(),
            FailOnDataLoss = reader.ReadBoolean(),
            MaxFolderCount = reader.ReadInt32(),
            MaxItemCount = reader.ReadInt32(),
            MaxNestedMessageDepth = reader.ReadInt32(),
            CheckpointIntervalItems = reader.ReadInt32(),
            MaxIndexRecordsInMemory = reader.ReadInt32(),
            RetainCheckpointOnDispose = reader.ReadBoolean(),
            MaxDiagnostics = reader.ReadInt32()
        };
        byte[] provider = reader.ReadBytes(16);
        if (provider.Length != 16) throw new EndOfStreamException("The PST checkpoint provider UID is truncated.");
        state.ProviderUid = new Guid(provider);
        state.NextFolderIndex = reader.ReadUInt32();
        state.NextMessageIndex = reader.ReadUInt32();
        state.UserFolderCount = reader.ReadInt32();
        state.ItemCount = reader.ReadInt32();
        state.File = new PstWriterFileCheckpoint(reader.ReadInt64(), reader.ReadInt64(),
            reader.ReadUInt64(), reader.ReadUInt64(), reader.ReadInt64());
        state.NodeCount = reader.ReadInt64();
        state.ItemJournalCount = reader.ReadInt64();
        state.ItemPayloadLength = reader.ReadInt64();
        state.DiagnosticsTruncated = reader.ReadBoolean();
        state.NamedProperties = PstNamedPropertyWriter.ReadCheckpoint(reader);

        int folderCount = reader.ReadInt32();
        if (folderCount < 5 || folderCount > state.MaxFolderCount + 5) {
            throw new InvalidDataException("The PST writer checkpoint folder count is invalid.");
        }
        for (int index = 0; index < folderCount; index++) {
            uint nid = reader.ReadUInt32();
            uint parentNid = reader.ReadUInt32();
            string name = reader.ReadString();
            string? containerClass = ReadNullableString(reader);
            bool isSearchFolder = reader.ReadBoolean();
            EmailStoreSpecialFolderKind specialFolderKind = version >= 2
                ? ReadSpecialFolderKind(reader)
                : InferLegacySpecialFolderKind(nid);
            var folder = new FolderState(nid, parentNid, name, containerClass,
                isSearchFolder, specialFolderKind) {
                NormalItemCount = reader.ReadInt32(),
                AssociatedItemCount = reader.ReadInt32(),
                UnreadItemCount = reader.ReadInt32()
            };
            state.Folders.Add(folder);
        }
        int diagnosticCount = reader.ReadInt32();
        if (diagnosticCount < 0 || diagnosticCount > 1_000_000) {
            throw new InvalidDataException("The PST writer checkpoint diagnostic count is invalid.");
        }
        for (int index = 0; index < diagnosticCount; index++) {
            state.Diagnostics.Add(new EmailStoreDiagnostic(reader.ReadString(), reader.ReadString(),
                (EmailStoreDiagnosticSeverity)reader.ReadInt32(), ReadNullableString(reader)));
        }
        state.Validate();
        return state;
    }

    private static EmailStoreSpecialFolderKind InferLegacySpecialFolderKind(uint nid) {
        if (nid == RootFolderNid) return EmailStoreSpecialFolderKind.Root;
        if (nid == IpmSubtreeNid) return EmailStoreSpecialFolderKind.IpmSubtree;
        if (nid == SearchRootNid) return EmailStoreSpecialFolderKind.SearchRoot;
        if (nid == DeletedItemsNid) return EmailStoreSpecialFolderKind.DeletedItems;
        return EmailStoreSpecialFolderKind.Unknown;
    }

    private static EmailStoreSpecialFolderKind ReadSpecialFolderKind(BinaryReader reader) {
        int value = reader.ReadInt32();
        if (!Enum.IsDefined(typeof(EmailStoreSpecialFolderKind), value)) {
            throw new InvalidDataException("The PST writer checkpoint contains an invalid special-folder role.");
        }
        return (EmailStoreSpecialFolderKind)value;
    }

    private void ReportProgress(EmailStorePstWriteStage stage) {
        if (_options.Progress == null) return;
        long bytes = stage == EmailStorePstWriteStage.Completed && File.Exists(_destinationPath)
            ? new FileInfo(_destinationPath).Length
            : _file.Length;
        _options.Progress.Report(new EmailStorePstWriteProgress(stage,
            _userFolderCount, _itemCount, bytes, _options.CheckpointPath));
    }

    private void DeleteCheckpointFile() {
        if (_options.CheckpointPath != null) TryDelete(_options.CheckpointPath);
    }

    private static void CleanupWorkingFiles(string temporaryPath, string destinationPath) {
        string fullPath = Path.GetFullPath(temporaryPath);
        ValidateTemporaryPath(destinationPath, fullPath);
        string? directory = Path.GetDirectoryName(fullPath);
        if (string.IsNullOrEmpty(directory) || !Directory.Exists(directory)) return;
        foreach (string candidate in Directory.EnumerateFiles(directory)) {
            string candidateFull = Path.GetFullPath(candidate);
            if (!IsWriterOwnedWorkingFile(fullPath, candidateFull)) continue;
            TryDelete(candidateFull);
        }
    }

    private static void ValidateTemporaryPath(string destinationPath, string temporaryPath) {
        string destination = Path.GetFullPath(destinationPath);
        string temporary = Path.GetFullPath(temporaryPath);
        string? destinationDirectory = Path.GetDirectoryName(destination);
        string? temporaryDirectory = Path.GetDirectoryName(temporary);
        if (destinationDirectory == null || temporaryDirectory == null ||
            !EmailStorePathIdentity.AreEquivalent(destinationDirectory, temporaryDirectory)) {
            throw new InvalidDataException("The PST checkpoint working path is outside the destination directory.");
        }
        string expectedPrefix = string.Concat(".", Path.GetFileName(destination), ".");
        string temporaryName = Path.GetFileName(temporary);
        if (!temporaryName.StartsWith(expectedPrefix, StringComparison.Ordinal) ||
            !temporaryName.EndsWith(".tmp", StringComparison.Ordinal)) {
            throw new InvalidDataException("The PST checkpoint working path is not writer-owned.");
        }
        string token = temporaryName.Substring(expectedPrefix.Length,
            temporaryName.Length - expectedPrefix.Length - 4);
        if (!Guid.TryParseExact(token, "N", out _)) {
            throw new InvalidDataException("The PST checkpoint working path has an invalid ownership token.");
        }
    }

    private static bool IsWriterOwnedWorkingFile(string temporaryPath, string candidatePath) {
        if (string.Equals(candidatePath, temporaryPath, StringComparison.Ordinal)) return true;
        if (!candidatePath.StartsWith(temporaryPath, StringComparison.Ordinal)) return false;
        string suffix = candidatePath.Substring(temporaryPath.Length);
        if (suffix == ".blocks" || suffix == ".nodes" || suffix == ".items" ||
            suffix == ".item-data" || suffix == ".amap") return true;
        return HasGuidSuffix(suffix, ".nodes.sort.") || HasGuidSuffix(suffix, ".items.sort.") ||
            HasGuidSuffix(suffix, ".btree.") || HasGuidSuffix(suffix, ".datatree.") ||
            HasGuidSuffix(suffix, ".table-matrix.") ||
            HasGuidSuffix(suffix, ".table-row-index.") ||
            HasGuidSuffix(suffix, ".table-subnodes.");
    }

    private static bool HasGuidSuffix(string value, string prefix) =>
        value.StartsWith(prefix, StringComparison.Ordinal) &&
        Guid.TryParseExact(value.Substring(prefix.Length), "N", out _);

    private static void CleanupCheckpointCommitFiles(string checkpointPath) {
        string fullPath = Path.GetFullPath(checkpointPath);
        string? directory = Path.GetDirectoryName(fullPath);
        if (string.IsNullOrEmpty(directory) || !Directory.Exists(directory)) return;
        string prefix = string.Concat(".", Path.GetFileName(fullPath), ".");
        foreach (string candidate in Directory.EnumerateFiles(directory)) {
            string name = Path.GetFileName(candidate);
            if (!name.StartsWith(prefix, StringComparison.Ordinal) ||
                !name.EndsWith(".tmp", StringComparison.Ordinal) ||
                name.Length <= prefix.Length + 4) continue;
            string token = name.Substring(prefix.Length, name.Length - prefix.Length - 4);
            if (Guid.TryParseExact(token, "N", out _)) TryDelete(candidate);
        }
    }

    private static void WriteNullableString(BinaryWriter writer, string? value) {
        writer.Write(value != null);
        if (value != null) writer.Write(value);
    }

    private static string? ReadNullableString(BinaryReader reader) =>
        reader.ReadBoolean() ? reader.ReadString() : null;

    private static bool FixedTimeEquals(byte[] left, byte[] right) {
        if (left.Length != right.Length) return false;
        int difference = 0;
        for (int index = 0; index < left.Length; index++) difference |= left[index] ^ right[index];
        return difference == 0;
    }

    private sealed class WriterCheckpointState {
        internal string DestinationPath { get; set; } = string.Empty;
        internal string TemporaryPath { get; set; } = string.Empty;
        internal string DisplayName { get; set; } = string.Empty;
        internal bool OverwriteExisting { get; set; }
        internal bool FailOnDataLoss { get; set; }
        internal int MaxFolderCount { get; set; }
        internal int MaxItemCount { get; set; }
        internal int MaxNestedMessageDepth { get; set; }
        internal int CheckpointIntervalItems { get; set; }
        internal int MaxIndexRecordsInMemory { get; set; }
        internal bool RetainCheckpointOnDispose { get; set; }
        internal int MaxDiagnostics { get; set; }
        internal Guid ProviderUid { get; set; }
        internal uint NextFolderIndex { get; set; }
        internal uint NextMessageIndex { get; set; }
        internal int UserFolderCount { get; set; }
        internal int ItemCount { get; set; }
        internal PstWriterFileCheckpoint File { get; set; }
        internal long NodeCount { get; set; }
        internal long ItemJournalCount { get; set; }
        internal long ItemPayloadLength { get; set; }
        internal bool DiagnosticsTruncated { get; set; }
        internal PstNamedPropertyWriter NamedProperties { get; set; } = null!;
        internal List<FolderState> Folders { get; } = new List<FolderState>();
        internal List<EmailStoreDiagnostic> Diagnostics { get; } = new List<EmailStoreDiagnostic>();

        internal EmailStorePstWriterOptions CreateOptions(string checkpointPath,
            IProgress<EmailStorePstWriteProgress>? progress) => new EmailStorePstWriterOptions(
                DisplayName, OverwriteExisting, FailOnDataLoss, MaxFolderCount,
                MaxItemCount, MaxNestedMessageDepth, checkpointPath,
                CheckpointIntervalItems, MaxIndexRecordsInMemory,
                RetainCheckpointOnDispose, MaxDiagnostics, progress);

        internal void Validate() {
            if (string.IsNullOrWhiteSpace(DestinationPath) || string.IsNullOrWhiteSpace(TemporaryPath) ||
                string.IsNullOrWhiteSpace(DisplayName) || MaxFolderCount <= 0 || MaxItemCount <= 0 ||
                MaxNestedMessageDepth < 0 || CheckpointIntervalItems <= 0 ||
                MaxIndexRecordsInMemory <= 0 || MaxDiagnostics <= 0 ||
                UserFolderCount < 0 || ItemCount < 0 ||
                UserFolderCount > MaxFolderCount || ItemCount > MaxItemCount ||
                NodeCount < 0 || ItemJournalCount < 0 || ItemPayloadLength < 0 ||
                NodeCount > int.MaxValue ||
                NamedProperties == null || ItemJournalCount != ItemCount ||
                Folders.Count != UserFolderCount + 5 || Diagnostics.Count > MaxDiagnostics ||
                File.StreamLength < 0x4800 || File.NextOffset < 0x4800 ||
                File.NextOffset > File.StreamLength || File.BlockCount < 0 ||
                File.NextBlockBid < 0x100 || File.NextPageBid < 0x100 ||
                Folders.Select(folder => folder.Nid).Distinct().Count() != Folders.Count ||
                Folders.Any(folder => folder.NormalItemCount < 0 ||
                    folder.AssociatedItemCount < 0 || folder.UnreadItemCount < 0)) {
                throw new InvalidDataException("The PST writer checkpoint state is inconsistent.");
            }
            if (!System.IO.File.Exists(TemporaryPath) ||
                !System.IO.File.Exists(string.Concat(TemporaryPath, ".blocks")) ||
                !System.IO.File.Exists(string.Concat(TemporaryPath, ".nodes")) ||
                !System.IO.File.Exists(string.Concat(TemporaryPath, ".items")) ||
                !System.IO.File.Exists(string.Concat(TemporaryPath, ".item-data"))) {
                throw new InvalidDataException("One or more checkpoint working files are missing.");
            }
            if (new FileInfo(TemporaryPath).Length < File.StreamLength ||
                new FileInfo(string.Concat(TemporaryPath, ".blocks")).Length <
                    checked(File.BlockCount * 24) ||
                new FileInfo(string.Concat(TemporaryPath, ".nodes")).Length <
                    checked(NodeCount * 24) ||
                new FileInfo(string.Concat(TemporaryPath, ".items")).Length <
                    checked(ItemJournalCount * 32) ||
                new FileInfo(string.Concat(TemporaryPath, ".item-data")).Length < ItemPayloadLength) {
                throw new InvalidDataException("One or more checkpoint working files are truncated.");
            }
        }

        internal void ValidateOwnership(string checkpointPath) {
            string fullCheckpointPath = Path.GetFullPath(checkpointPath);
            if (EmailStorePathIdentity.AreEquivalent(fullCheckpointPath, DestinationPath) ||
                EmailStorePathIdentity.AreEquivalent(fullCheckpointPath, TemporaryPath)) {
                throw new InvalidDataException(
                    "The PST checkpoint path collides with a destination or working file.");
            }
            ValidateTemporaryPath(DestinationPath, TemporaryPath);
        }
    }
}
