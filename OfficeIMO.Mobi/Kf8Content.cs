using System.Globalization;
using static OfficeIMO.Mobi.PalmDatabase;

namespace OfficeIMO.Mobi;

internal sealed class Kf8Content {
    private readonly (int Part, int Insertion)[] _positions;
    private readonly (int Start, int End)[] _partRanges;

    internal Kf8Content(PalmDatabase database, MobiHeader header, byte[] raw, CancellationToken token) {
        int markUpEnd = raw.Length;
        if (header.FlowIndex is { } flowRecord) {
            ReadOnlySpan<byte> flow = database.Record(header.Record + flowRecord);
            if (!HasSignature(flow, "FDST") || Number(flow, 4) != 12)
                throw new InvalidDataException("The Kindle FDST flow record is malformed.");
            int count = Number(flow, 8);
            if (count != header.FlowCount || count > 256 || Number(flow, 12) != 0)
                throw new InvalidDataException("The Kindle FDST flow count or start offset is invalid.");
            markUpEnd = Number(flow, 16);
            int previous = 0;
            for (int i = 0; i < count; i++) {
                int start = Number(flow, 12 + i * 8), end = Number(flow, 16 + i * 8);
                if (start != previous || end < start || end > raw.Length)
                    throw new InvalidDataException("Kindle flow ranges must be ordered and contained in the decompressed text.");
                previous = end;
            }
            if (previous != raw.Length) throw new InvalidDataException("Kindle FDST flows do not cover the decompressed text.");
            HasSeparateFlows = count > 1;
        }

        if (header.SkeletonIndex is not { } skeletonRecord || header.FragmentIndex is not { } fragmentRecord)
            throw new NotSupportedException("The Kindle Format 8 book does not provide supported skeleton and fragment indexes.");

        var skeletons = new MobiIndex(database, header.Record + skeletonRecord, header.Encoding, token).Rows;
        var fragments = new MobiIndex(database, header.Record + fragmentRecord, header.Encoding, token).Rows;
        if (skeletons.Count is < 1 or > 1_000) throw new InvalidDataException("The Kindle skeleton count exceeds its bounds.");
        _positions = new (int, int)[fragments.Count];
        Parts = new byte[skeletons.Count][];
        _partRanges = new (int, int)[skeletons.Count];
        ReadOnlySpan<byte> markup = raw.AsSpan(0, markUpEnd);
        int fragment = 0;
        int previousEnd = 0;
        for (int i = 0; i < skeletons.Count; i++) {
            token.ThrowIfCancellationRequested();
            var skeleton = skeletons[i];
            int count = skeleton.Require(1);
            int start = skeleton.Require(6);
            int length = skeleton.Require(6, 1);
            if (count > fragments.Count - fragment || start < previousEnd)
                throw new InvalidDataException("Kindle skeletons overlap or reference unavailable fragments.");
            ReadOnlySpan<byte> template = Slice(markup, start, length);
            using var part = new MemoryStream();
            int consumed = 0;
            int fragmentBytes = 0;
            for (int j = 0; j < count; j++, fragment++) {
                token.ThrowIfCancellationRequested();
                var row = fragments[fragment];
                int file = row.Require(3);
                int offset = row.Require(6);
                int size = row.Require(6, 1);
                if (file != i || offset != fragmentBytes ||
                    !int.TryParse(row.Name, NumberStyles.None, CultureInfo.InvariantCulture, out int absoluteInsertion))
                    throw new InvalidDataException("A Kindle fragment does not match its skeleton or source range.");
                int insertion = absoluteInsertion - start;
                int split = insertion - fragmentBytes;
                if (split < consumed || split > template.Length)
                    throw new InvalidDataException("A Kindle fragment points outside its reconstructed skeleton.");
                byte[] before = template.Slice(consumed, split - consumed).ToArray();
                part.Write(before, 0, before.Length);
                byte[] bytes = Slice(markup, start + length + fragmentBytes, size).ToArray();
                part.Write(bytes, 0, bytes.Length);
                _positions[fragment] = (i, insertion);
                consumed = split;
                fragmentBytes += size;
            }
            byte[] rest = template.Slice(consumed).ToArray();
            part.Write(rest, 0, rest.Length);
            Parts[i] = part.ToArray();
            previousEnd = start + length + fragmentBytes;
            _partRanges[i] = (start, previousEnd);
        }
        if (fragment != fragments.Count)
            throw new InvalidDataException("The Kindle fragment index contains unassigned fragments.");
    }

    internal byte[][] Parts { get; }
    internal bool HasSeparateFlows { get; }

    internal int? ResolveAbsolutePart(int offset) {
        int start = 0, end = _partRanges.Length - 1;
        while (start <= end) {
            int index = start + (end - start) / 2;
            var range = _partRanges[index];
            if (offset < range.Start) end = index - 1;
            else if (offset >= range.End) start = index + 1;
            else return index;
        }
        return null;
    }

    internal (int Part, int Offset)? Resolve(int fragment, int offset) {
        if (fragment < 0 || fragment >= _positions.Length || offset < 0) return null;
        var entry = _positions[fragment];
        return offset <= Parts[entry.Part].Length - entry.Insertion ? (entry.Part, entry.Insertion + offset) : null;
    }
}
