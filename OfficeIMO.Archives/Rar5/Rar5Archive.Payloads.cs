using OfficeIMO.Drawing;

namespace OfficeIMO.Archives.Rar5;

internal sealed partial class Rar5Archive {
    private void FinishMember(Member member) {
        if (member.Kind == ArchiveEntryKind.Directory && (member.Fields.Size != 0 || member.PackedLength != 0))
            throw new InvalidDataException("A RAR5 directory cannot contain file data.");
        if (member.Extras.Redirection is { Kind: Rar5RedirectionKind.FileCopy } copy) {
            if (copy.TargetIsDirectory || member.PackedLength != 0 || !_fileSources.TryGetValue(copy.Target, out Member? target) ||
                target.Fields.Size != member.Fields.Size)
                throw new InvalidDataException("The RAR5 file copy has no matching earlier regular file.");
            member.CopySource = target;
        } else if (member.Kind == ArchiveEntryKind.File && member.Extras.Redirection == null) {
            if (member.Fields.Method == 0 && member.Fields.Size != member.PackedLength)
                throw new InvalidDataException("The RAR5 member has inconsistent stored sizes.");
            _fileSources.Add(member.Path, member);
        }
    }

    internal void Read(Action<string, byte[]> consume, CancellationToken token, Action<string>? directory = null) {
        var selected = new List<Member>();
        long total = 0;
        foreach (Member member in _members) {
            token.ThrowIfCancellationRequested();
            if (member.Kind == ArchiveEntryKind.Directory || !_limits.Select(member.Path)) continue;
            Member source = member.CopySource ?? member;
            if (source.Fields.Method != 0)
                throw new Rar5CompressionNotSupportedException(member.Fields.Name, source.Fields.Method, source.Fields.Dictionary);
            if (member.UnsupportedReason is { } reason) throw new NotSupportedException(reason + " " + member.Fields.Name);
            ArchiveLimits.Check(member.Fields.Size, Math.Min(_limits.Entry, int.MaxValue), nameof(ArchiveExtractionOptions.MaxEntryBytes));
            ArchiveLimits.Add(ref total, member.Fields.Size, _limits.Expanded, nameof(ArchiveExtractionOptions.MaxExpandedBytes));
            selected.Add(member);
        }
        token.ThrowIfCancellationRequested();
        _limits.Operation.Select(total);
        foreach (Member member in selected) {
            token.ThrowIfCancellationRequested();
            if (member.CopySource != null) {
                Part alias = member.Parts[0];
                _budget.Charge(0, alias.Integrity.Hash == null ? 64 : 4096);
                Verify(Array.Empty<byte>(), 0, 0, alias.Integrity, member.Path, alias.Offset, token);
            }
            consume(member.Path, ReadPayload(member.CopySource ?? member, token));
        }
        if (directory != null && _limits.Include == null) {
            foreach (Member member in _members) {
                token.ThrowIfCancellationRequested();
                if (member.Kind == ArchiveEntryKind.Directory) directory(member.Path);
            }
        }
        token.ThrowIfCancellationRequested();
    }

    private byte[] ReadPayload(Member member, CancellationToken token) {
        long size = member.Fields.Size;
        // Allow independently for copying, piece checks, and the final checksum over the whole member.
        _budget.Charge(size, 3 * Math.Max(64L, size) + 4096L * member.Parts.Count(part => part.Integrity.Hash != null));
        var bytes = new byte[(int)size];
        int position = 0;
        for (int index = 0; index < member.Parts.Count; index++) {
            Part part = member.Parts[index];
            int length = (int)part.Length;
            part.Source.Position = part.Offset;
            Rar5HeaderReader.Fill(part.Source, bytes, length, token, position);
            bool last = index == member.Parts.Count - 1;
            Verify(bytes, last ? 0 : position, last ? bytes.Length : length, part.Integrity, member.Path, part.Offset, token);
            position += length;
        }
        return bytes;
    }

    private static void Verify(byte[] bytes, int start, int length, Integrity check, string name, long offset, CancellationToken token) {
        if (check.Crc is uint expected) {
            uint crc = OfficePngCrc32.Begin();
            for (int position = 0; position < length;) {
                token.ThrowIfCancellationRequested();
                int count = Math.Min(65536, length - position);
                crc = OfficePngCrc32.Append(crc, bytes, start + position, count);
                position += count;
            }
            if (OfficePngCrc32.Complete(crc) != expected)
                throw new InvalidDataException(FormattableString.Invariant($"The RAR5 file checksum (CRC-32) does not match for '{name}' at byte offset {offset}."));
        }
        if (check.Hash != null && !Blake2s.TreeDigest(bytes, start, length, token).SequenceEqual(check.Hash))
            throw new InvalidDataException(FormattableString.Invariant($"The RAR5 BLAKE2sp checksum does not match for '{name}' at byte offset {offset}."));
    }
}
