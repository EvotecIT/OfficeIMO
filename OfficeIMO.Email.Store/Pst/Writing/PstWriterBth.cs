namespace OfficeIMO.Email.Store;

internal static class PstWriterBth {
    private const int MaximumAllocationBytes = 8000;

    internal static void Complete(PstWriterHeap heap, byte[] header,
        int keySize, int valueSize, byte[] leafRecords) {
        if (heap == null) throw new ArgumentNullException(nameof(heap));
        if (header == null || header.Length < 8) throw new ArgumentException("A BTH header requires eight bytes.", nameof(header));
        int recordSize = checked(keySize + valueSize);
        if (recordSize <= 0 || leafRecords.Length % recordSize != 0) {
            throw new ArgumentException("BTH leaf records are not aligned to their declared size.", nameof(leafRecords));
        }
        header[0] = 0xB5;
        header[1] = checked((byte)keySize);
        header[2] = checked((byte)valueSize);
        if (leafRecords.Length == 0) {
            header[3] = 0;
            PstBinary.WriteUInt32(header, 4, 0);
            return;
        }

        int recordsPerLeaf = Math.Max(1, MaximumAllocationBytes / recordSize);
        var level = new List<BthReference>();
        int recordCount = leafRecords.Length / recordSize;
        for (int recordIndex = 0; recordIndex < recordCount; recordIndex += recordsPerLeaf) {
            int count = Math.Min(recordsPerLeaf, recordCount - recordIndex);
            var allocation = new byte[count * recordSize];
            Buffer.BlockCopy(leafRecords, recordIndex * recordSize,
                allocation, 0, allocation.Length);
            uint hid = heap.Add(allocation);
            var firstKey = new byte[keySize];
            Buffer.BlockCopy(allocation, 0, firstKey, 0, keySize);
            level.Add(new BthReference(firstKey, hid));
        }

        int indexLevels = 0;
        int indexRecordSize = checked(keySize + 4);
        int recordsPerIndex = Math.Max(1, MaximumAllocationBytes / indexRecordSize);
        while (level.Count > 1) {
            var parent = new List<BthReference>();
            for (int offset = 0; offset < level.Count; offset += recordsPerIndex) {
                int count = Math.Min(recordsPerIndex, level.Count - offset);
                var allocation = new byte[count * indexRecordSize];
                for (int index = 0; index < count; index++) {
                    BthReference child = level[offset + index];
                    int target = index * indexRecordSize;
                    Buffer.BlockCopy(child.FirstKey, 0, allocation, target, keySize);
                    PstBinary.WriteUInt32(allocation, target + keySize, child.Hid);
                }
                uint hid = heap.Add(allocation);
                parent.Add(new BthReference(level[offset].FirstKey, hid));
            }
            level = parent;
            indexLevels++;
            if (indexLevels > byte.MaxValue) throw new NotSupportedException("The PST BTH is too deep.");
        }
        header[3] = checked((byte)indexLevels);
        PstBinary.WriteUInt32(header, 4, level[0].Hid);
    }

    private sealed class BthReference {
        internal BthReference(byte[] firstKey, uint hid) {
            FirstKey = firstKey;
            Hid = hid;
        }
        internal byte[] FirstKey { get; }
        internal uint Hid { get; }
    }
}
