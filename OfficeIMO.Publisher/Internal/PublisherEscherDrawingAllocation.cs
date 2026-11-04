namespace OfficeIMO.Publisher.Internal;

/// <summary>Retains ordered OfficeArt identifier history and allocates only fresh slots in the selected drawing.</summary>
internal sealed class PublisherEscherDrawingAllocation {
    private readonly List<(uint Owner, uint Used)> _clusters = new();
    private readonly int _maximumClusters;
    private int _currentCluster;
    private uint _drawingId;
    internal uint Maximum { get; private set; }
    internal uint LastSpid { get; private set; }

    private PublisherEscherDrawingAllocation(int maximumClusters) => _maximumClusters = maximumClusters;

    /// <summary>Reserves preceding cluster positions for a standalone fresh drawing encoder.</summary>
    internal static PublisherEscherDrawingAllocation Fresh(uint firstCluster, CancellationToken token) {
        var allocation = new PublisherEscherDrawingAllocation((int)(uint.MaxValue >> 10));
        for (uint cluster = 1; cluster < firstCluster; cluster++) {
            token.ThrowIfCancellationRequested(); allocation._clusters.Add((0, 0));
        }
        return allocation;
    }

    internal static PublisherEscherDrawingAllocation? Read(PublisherBinaryData data, PublisherEscherRecord record,
        PublisherParseContext context) {
        if (record.Initial != 0 || record.Length < 16) return null;
        uint count = data.U32(record.Offset + 4);
        if (count < 2 || count - 1 > uint.MaxValue >> 10 || count - 1 > context.Options.Limits.MaxRecords
            || 16L + (count - 1L) * 8 != record.Length) return null;
        var allocation = new PublisherEscherDrawingAllocation(Math.Min(context.Options.Limits.MaxRecords, (int)(uint.MaxValue >> 10))) {
            Maximum = data.U32(record.Offset)
        };
        if (allocation.Maximum > (ulong)count * 1024) return null;
        for (int index = 0; index < count - 1; index++) {
            context.Record(); int slot = checked(record.Offset + 16 + index * 8);
            uint owner = data.U32(slot), used = data.U32(slot + 4);
            ulong next = ((ulong)index + 1) * 1024 + used;
            // Zero-owner entries can retain consumed identifiers after native
            // temporary drawings disappear. Their positions and cursors survive.
            if (owner > 4095 || used > 1024 || used != 0 && next > allocation.Maximum) return null;
            allocation._clusters.Add((owner, used));
        }
        return allocation;
    }

    /// <summary>Adds a drawing at the global tail without reclaiming another drawing's retained history.</summary>
    internal bool StartDrawing(uint drawingId, uint used, CancellationToken token) {
        token.ThrowIfCancellationRequested();
        if (drawingId is 0 or > 4095 || used > 1024 || _clusters.Count >= _maximumClusters) return false;
        foreach (var entry in _clusters) {
            token.ThrowIfCancellationRequested(); if (entry.Owner == drawingId) return false;
        }
        ulong next = ((ulong)_clusters.Count + 1) * 1024 + used;
        if (next > uint.MaxValue) return false;
        _clusters.Add((drawingId, used)); _currentCluster = _clusters.Count; _drawingId = drawingId;
        LastSpid = used == 0 ? 0 : (uint)next - 1;
        if (used != 0) Maximum = Math.Max(Maximum, (uint)next);
        return true;
    }

    internal bool SelectDrawing(uint drawingId, uint cursor, CancellationToken token) {
        token.ThrowIfCancellationRequested(); uint cluster = cursor >> 10;
        if (drawingId is 0 or > 4095 || cluster == 0 || cluster > _clusters.Count) return false;
        var entry = _clusters[(int)cluster - 1];
        if (entry.Owner != drawingId || entry.Used != (cursor & 1023) + 1 || cursor >= Maximum) return false;
        for (int index = (int)cluster; index < _clusters.Count; index++) {
            token.ThrowIfCancellationRequested();
            if (_clusters[index].Owner == drawingId && _clusters[index].Used != 0) return false;
        }
        _currentCluster = (int)cluster; _drawingId = drawingId; LastSpid = cursor;
        return true;
    }

    internal bool Owns(uint drawingId, uint spid) {
        uint cluster = spid >> 10;
        return cluster != 0 && cluster <= _clusters.Count && _clusters[(int)cluster - 1].Owner == drawingId
            && (spid & 1023) < _clusters[(int)cluster - 1].Used && spid < Maximum;
    }

    /// <summary>Uses the current cluster's unused suffix, then appends a new cluster; retired slots are never reused.</summary>
    internal bool Allocate(CancellationToken token, out uint spid) {
        token.ThrowIfCancellationRequested(); spid = 0;
        if (_currentCluster == 0) return false;
        var entry = _clusters[_currentCluster - 1];
        if (entry.Used == 1024) {
            if (_clusters.Count >= _maximumClusters || ((ulong)_clusters.Count + 1) * 1024 >= uint.MaxValue) return false;
            _clusters.Add((_drawingId, 0)); _currentCluster = _clusters.Count; entry = (_drawingId, 0);
        }
        ulong value = (ulong)_currentCluster * 1024 + entry.Used;
        if (value >= uint.MaxValue) return false;
        spid = (uint)value; _clusters[_currentCluster - 1] = (_drawingId, entry.Used + 1);
        LastSpid = spid; Maximum = Math.Max(Maximum, spid + 1);
        return true;
    }

    /// <summary>Owns the complete FDGG replacement so scalar patches cannot overwrite allocation metadata.</summary>
    internal byte[] Record(uint savedShapes, uint savedDrawings, CancellationToken token) {
        token.ThrowIfCancellationRequested();
        var result = new byte[checked(24 + _clusters.Count * 8)];
        Write16(0, 0); Write16(2, 0xF006); Write(4, (uint)result.Length - 8);
        Write(8, Maximum); Write(12, (uint)_clusters.Count + 1); Write(16, savedShapes); Write(20, savedDrawings);
        for (int index = 0; index < _clusters.Count; index++) {
            token.ThrowIfCancellationRequested(); Write(24 + index * 8, _clusters[index].Owner); Write(28 + index * 8, _clusters[index].Used);
        }
        return result;

        void Write(int offset, uint value) => BitConverter.GetBytes(value).CopyTo(result, offset);
        void Write16(int offset, ushort value) => BitConverter.GetBytes(value).CopyTo(result, offset);
    }
}
