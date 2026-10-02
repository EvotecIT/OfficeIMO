using System;

namespace OfficeIMO.Drawing;

internal sealed partial class OfficeAvifContainerReader {
    private void ReadPrimary(Box box) {
        int p = FullBox(box, 1, out uint flags);
        Require(flags == 0);
        var cursor = new Cursor(this, p, box.End);
        _primaryId = cursor.Id(_bytes[box.Start] == 1);
        Require(_primaryId != 0);
        cursor.End();
    }

    private void ReadItemTypes(Box box) {
        int p = FullBox(box, 1, out uint flags);
        Require(flags == 0);
        var cursor = new Cursor(this, p, box.End);
        int count = checked((int)cursor.Integer(_bytes[box.Start] == 0 ? 2 : 4));
        Require(count > 0 && count <= MaximumItems);
        int found = 0;
        foreach (Box item in Boxes(cursor.Position, box.End)) {
            Require(item.Type == "infe" && ++found <= count);
            int itemStart = FullBox(item, 3, out uint itemFlags);
            Require(_bytes[item.Start] >= 2 && (itemFlags & ~1U) == 0);
            var entry = new Cursor(this, itemStart, item.End);
            uint id = entry.Id(_bytes[item.Start] == 3);
            Require(id != 0 && entry.Integer(2) == 0 && !_types.ContainsKey(id));
            string type = entry.Type();
            entry.SkipTerminatedString(); // Item name must terminate within its own box, without an unused allocation.
            if (type == "av01") entry.End();
            _types.Add(id, type);
        }
        Require(found == count);
    }

    /// <summary>Retains whole in-file extents only; external data and derived/stitched items require another decode path.</summary>
    private void ReadLocations(Box box) {
        int p = FullBox(box, 2, out uint flags);
        Require(flags == 0);
        int version = _bytes[box.Start];
        var cursor = new Cursor(this, p, box.End);
        int sizes = (int)cursor.Integer(1);
        int sizes2 = (int)cursor.Integer(1);
        int offsetSize = sizes >> 4, lengthSize = sizes & 15;
        int baseSize = sizes2 >> 4, indexSize = sizes2 & 15;
        Require(IntegerSize(offsetSize) && IntegerSize(lengthSize) && IntegerSize(baseSize));
        Require(version == 0 ? indexSize == 0 : IntegerSize(indexSize));
        int count = checked((int)cursor.Integer(version == 2 ? 4 : 2));
        Require(count > 0 && count <= MaximumItems);
        for (int i = 0; i < count; i++) {
            _options.CancellationToken.ThrowIfCancellationRequested();
            uint id = cursor.Id(version == 2);
            Require(id != 0 && !_locations.ContainsKey(id));
            if (version != 0) Require(cursor.Integer(2) == 0); // In-file construction method 0, reserved bits zero.
            Require(cursor.Integer(2) == 0); // No external data reference.
            ulong baseOffset = cursor.Integer(baseSize);
            Require(cursor.Integer(2) == 1); // One complete extent, no concatenation or item derivation.
            if (version != 0 && indexSize > 0) Require(cursor.Integer(indexSize) == 0);
            ulong offset = checked(baseOffset + cursor.Integer(offsetSize));
            ulong length = cursor.Integer(lengthSize);
            Require(length > 0 && offset <= (ulong)_bytes.Length && length <= (ulong)_bytes.Length - offset);
            int start = checked((int)offset), end = checked(start + (int)length);
            bool inMediaData = false;
            foreach (Box data in _dataBoxes) inMediaData |= start >= data.Start && end <= data.End;
            Require(inMediaData);
            _locations.Add(id, (start, (int)length));
        }
        cursor.End();
    }

    private static bool IntegerSize(int size) => size is 0 or 4 or 8;

    private void ReadReferences(Box box) {
        int p = FullBox(box, 1, out uint flags);
        Require(flags == 0);
        bool large = _bytes[box.Start] == 1;
        foreach (Box reference in Boxes(p, box.End)) {
            var cursor = new Cursor(this, reference.Start, reference.End);
            uint from = cursor.Id(large);
            int count = (int)cursor.Integer(2);
            Require(from != 0 && count > 0 && count <= MaximumItems);
            for (int i = 0; i < count; i++) {
                _options.CancellationToken.ThrowIfCancellationRequested();
                uint to = cursor.Id(large);
                Require(to != 0);
                if (reference.Type == "auxl") {
                    Require(_alphaReferences.Count < MaximumItems);
                    _alphaReferences.Add((from, to));
                }
                if (reference.Type == "prem") _premultipliedItems.Add(from);
            }
            cursor.End();
        }
    }
}
