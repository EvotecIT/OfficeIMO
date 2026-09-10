using OfficeIMO.Core.Internal;

namespace OfficeIMO.Project;

/// <summary>Producer-supplied field storage mapping; native positions never enter the public model.</summary>
internal sealed class ProjectNativeField {
    internal uint Id, Mask;
    internal int Offset, Source, Type, Size, Position;
    internal bool Secondary;
}

/// <summary>Bounded fixed/variable table access for a qualified MPP14 container.</summary>
internal sealed class ProjectNativeTable {
    internal readonly Dictionary<uint, ProjectNativeField> Fields = new Dictionary<uint, ProjectNativeField>();
    internal readonly List<ProjectNativeRecord> Records = new List<ProjectNativeRecord>();
    internal readonly Dictionary<int, Dictionary<uint, ProjectNativeValue>> Variable = new Dictionary<int, Dictionary<uint, ProjectNativeValue>>();
    private readonly Dictionary<string, byte[]> _streams;
    private readonly string _prefix;

    internal ProjectNativeTable(OfficeCompoundFile file, string table, ProjectNativeValue primaryMap, ProjectNativeValue? extendedMap,
        int maxRecords, CancellationToken token) {
        _streams = file.Streams.ToDictionary(s => s.Key, s => s.Value, StringComparer.Ordinal);
        _prefix = "   114/" + table + "/";
        ReadMap(primaryMap, false, token);
        if (extendedMap.HasValue) ReadMap(extendedMap.Value, true, token);
        ReadVariable(maxRecords, token);
        byte[] metaBytes = Stream("FixedMeta"), fixedBytes = Stream("FixedData");
        var meta = new ProjectNativeValue(metaBytes, 0, metaBytes.Length);
        if (meta.Length < 16 || meta.UInt32() != 0xfadfadba) throw new InvalidDataException("Unrecognized native fixed metadata header.");
        int count = meta.Int32(8);
        if (count < 0 || count > maxRecords || (count == 0 ? meta.Length != 16 : (meta.Length - 16) % count != 0))
            throw new InvalidDataException("Invalid native fixed record count.");
        int stride = count == 0 ? 0 : (meta.Length - 16) / count;
        if (count > 0 && (stride < 8 || stride > 4096)) throw new InvalidDataException("Native fixed metadata width is outside the qualified bounds.");
        byte[]? secondMetaBytes = OptionalStream("Fixed2Meta"), secondBytes = OptionalStream("Fixed2Data");
        var secondMeta = new ProjectNativeValue(secondMetaBytes ?? Array.Empty<byte>(), 0, secondMetaBytes?.Length ?? 0);
        int secondStride = count > 0 && secondMeta.Length >= 16 && (secondMeta.Length - 16) % count == 0 ? (secondMeta.Length - 16) / count : 0;
        if (secondMetaBytes != null && (secondMeta.Length < 16 || secondMeta.UInt32() != 0xfadfadba || secondMeta.Int32(8) != count || secondBytes == null ||
            (count == 0 ? secondMeta.Length != 16 : secondStride < 8 || secondStride > 4096)))
            throw new InvalidDataException("Secondary record metadata does not match the primary table.");
        int mainSize = Fields.Values.Where(f => f.Source == 10 && !f.Secondary).Select(f => checked(f.Offset + f.Size)).DefaultIfEmpty(0).Max();
        int secondarySize = Fields.Values.Where(f => f.Source == 10 && f.Secondary).Select(f => checked(f.Offset + f.Size)).DefaultIfEmpty(0).Max();
        for (int index = 0; index < count; index++) {
            token.ThrowIfCancellationRequested();
            var recordMeta = meta.Slice(16 + index * stride, stride);
            // Deleted and reserved rows have short sentinel records, never semantic entities.
            if ((recordMeta.UInt16() & 6) != 0) continue;
            int offset = recordMeta.Int32(4);
            var data = new ProjectNativeValue(fixedBytes, offset, mainSize);
            ProjectNativeValue? second = null, secondFlags = null;
            if (secondBytes != null && secondStride >= 8) {
                var part = secondMeta.Slice(16 + index * secondStride, secondStride);
                secondFlags = part;
                second = new ProjectNativeValue(secondBytes, part.Int32(4), secondarySize);
            }
            Records.Add(new ProjectNativeRecord(this, index, data, recordMeta, second, secondFlags));
        }
    }

    private void ReadMap(ProjectNativeValue map, bool secondary, CancellationToken token) {
        if (map.Length % 28 != 0 || map.Length / 28 > 8192) throw new InvalidDataException("Unrecognized or excessive native field table width.");
        int primaryCount = Fields.Count;
        for (int i = 0; i < map.Length; i += 28) {
            token.ThrowIfCancellationRequested();
            uint id = map.UInt32(i + 12);
            if (Fields.TryGetValue(id, out var existing)) {
                if (!secondary || existing.Position != i / 28 || existing.Secondary)
                    throw new InvalidDataException("Duplicate or reordered native field mapping.");
                if (existing.Mask != map.UInt32(i) || existing.Offset != map.Int32(i + 4) || existing.Source != map.Int32(i + 8) ||
                    existing.Type != map.UInt16(i + 20) || existing.Size != map.UInt16(i + 22))
                    throw new InvalidDataException("The extended native map changes a primary field definition.");
                continue;
            }
            if (secondary && i / 28 < primaryCount) throw new InvalidDataException("Secondary field map does not extend the primary map.");
            int source = map.Int32(i + 8);
            var field = new ProjectNativeField { Id = id, Mask = map.UInt32(i), Offset = map.Int32(i + 4), Source = source,
                Type = map.UInt16(i + 20), Size = map.UInt16(i + 22), Secondary = secondary,
                Position = i / 28 - (secondary ? primaryCount : 0) };
            if (source == 10 && (field.Offset < 0 || field.Offset > 1024 * 1024 || field.Size == 0))
                throw new InvalidDataException("Native fixed field position exceeds supported limits.");
            Fields.Add(id, field);
        }
    }

    private void ReadVariable(int maxRecords, CancellationToken token) {
        var bytes = Stream("VarMeta");
        var meta = new ProjectNativeValue(bytes, 0, bytes.Length);
        var dataBytes = OptionalStream("Var2Data") ?? Array.Empty<byte>();
        var data = new ProjectNativeValue(dataBytes, 0, dataBytes.Length);
        if (meta.Length < 24 || meta.UInt32() != 0xfadfadba || (meta.Length - 24) % 12 != 0 || meta.Int32(8) != (meta.Length - 24) / 12)
            throw new InvalidDataException("Unrecognized native variable metadata header.");
        for (int i = 24; i < meta.Length; i += 12) {
            token.ThrowIfCancellationRequested();
            int uid = meta.Int32(i), offset = meta.Int32(i + 4);
            uint fieldId = meta.UInt32(i + 8);
            int length = data.Int32(offset);
            var value = data.Slice(checked(offset + 4), length);
            if (!Variable.TryGetValue(uid, out var fields)) {
                if (Variable.Count >= maxRecords) throw new InvalidDataException("Native variable record budget exceeded.");
                Variable.Add(uid, fields = new Dictionary<uint, ProjectNativeValue>());
            }
            if (fields.ContainsKey(fieldId)) throw new InvalidDataException("Duplicate native variable field.");
            fields.Add(fieldId, value);
        }
    }
    private byte[] Stream(string name) => OptionalStream(name) ?? throw new InvalidDataException("Required native stream is missing: " + _prefix + name);
    private byte[]? OptionalStream(string name) => _streams.TryGetValue(_prefix + name, out var value) ? value : null;
}

internal sealed class ProjectNativeRecord {
    private readonly ProjectNativeTable _table;
    private readonly ProjectNativeValue _data, _metadata;
    private readonly ProjectNativeValue? _secondary, _secondaryMetadata;
    internal int Uid { get; set; }
    internal int MetadataIndex { get; }
    internal ProjectNativeRecord(ProjectNativeTable table, int metadataIndex, ProjectNativeValue data, ProjectNativeValue metadata,
        ProjectNativeValue? secondary, ProjectNativeValue? secondaryMetadata) {
        _table = table; MetadataIndex = metadataIndex; _data = data; _metadata = metadata; _secondary = secondary; _secondaryMetadata = secondaryMetadata;
    }
    internal ProjectNativeValue? Value(uint id) {
        if (_table.Fields.TryGetValue(id, out var field) && field.Source == 10) {
            if (!Present(field)) return null;
            var data = field.Secondary ? _secondary : _data;
            return data?.Slice(field.Offset, field.Size);
        }
        return _table.Variable.TryGetValue(Uid, out var values) && values.TryGetValue(id, out var value) ? value : (ProjectNativeValue?)null;
    }
    private bool Present(ProjectNativeField field) {
        var metadata = field.Secondary ? _secondaryMetadata : _metadata;
        return metadata.HasValue && (metadata.Value.Byte(8 + field.Position / 8) & (1 << (field.Position % 8))) != 0;
    }
    internal bool? Boolean(uint id) => _table.Fields.TryGetValue(id, out var field) && (field.Source == 18 || field.Source == 19 || field.Source == 22)
        ? Present(field) : (bool?)null;
    internal int? Integer(uint id) { var value = Value(id); return value.HasValue ? value.Value.Length == 2 ? value.Value.Int16() : value.Value.Int32() : (int?)null; }
    internal string? Text(uint id) => Value(id)?.Unicode();
    internal DateTime? Date(uint id) => Value(id)?.Date();
    internal decimal? Number(uint id) {
        var value = Value(id);
        if (!value.HasValue) return null;
        double number = value.Value.Double();
        if (double.IsNaN(number) || double.IsInfinity(number) || number > (double)decimal.MaxValue || number < (double)decimal.MinValue)
            throw new InvalidDataException("Non-finite or out-of-range native numeric value.");
        return (decimal)number;
    }
}
