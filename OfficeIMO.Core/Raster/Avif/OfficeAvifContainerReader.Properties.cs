using System;
using System.Collections.Generic;

namespace OfficeIMO.Drawing;

internal sealed partial class OfficeAvifContainerReader {
    private void ReadProperties(Box box) {
        var associations = new List<Box>();
        bool hasContainer = false;
        foreach (Box child in Boxes(box.Start, box.End)) {
            if (child.Type == "ipco") {
                Require(!hasContainer);
                hasContainer = true;
                foreach (Box property in Boxes(child.Start, child.End)) {
                    Require(_properties.Count < MaximumItems);
                    _properties.Add(property);
                }
            } else if (child.Type == "ipma") {
                associations.Add(child);
            }
        }
        Require(hasContainer && associations.Count > 0);
        foreach (Box association in associations) ReadAssociations(association);
    }

    private void ReadAssociations(Box box) {
        int p = FullBox(box, 1, out uint flags);
        Require((flags & ~1U) == 0);
        var cursor = new Cursor(this, p, box.End);
        int count = checked((int)cursor.Integer(4));
        Require(count <= MaximumItems);
        for (int i = 0; i < count; i++) {
            _options.CancellationToken.ThrowIfCancellationRequested();
            uint id = cursor.Id(_bytes[box.Start] == 1);
            Require(id != 0 && !_associations.ContainsKey(id) && _associations.Count < MaximumItems);
            int entries = (int)cursor.Integer(1);
            var values = new List<(int Index, bool Essential)>(entries);
            var seen = new HashSet<int>();
            for (int j = 0; j < entries; j++) {
                int encoded = (int)cursor.Integer(flags == 0 ? 1 : 2);
                int essentialBit = flags == 0 ? 128 : 32768;
                int index = encoded & (essentialBit - 1);
                bool essential = (encoded & essentialBit) != 0;
                Require(index <= _properties.Count);
                if (index == 0) { Require(!essential); continue; }
                Require(seen.Add(index));
                values.Add((index, essential));
            }
            _associations.Add(id, values);
        }
        cursor.End();
    }

    private OfficeAvifImageItem ResolveItem(uint id, bool alpha) {
        Require(_types.TryGetValue(id, out string? type) && type == "av01");
        Require(_locations.TryGetValue(id, out var location));
        Require(_associations.TryGetValue(id, out var properties));
        int width = 0, height = 0, pixelChannels = 0, pixelDepth = 0;
        byte[]? configuration = null;
        OfficeAvifColorDescription? color = null;
        bool hasAuxiliaryType = false;
        var seen = new HashSet<string>(StringComparer.Ordinal);
        foreach (var association in properties!) {
            _options.CancellationToken.ThrowIfCancellationRequested();
            Box property = _properties[association.Index - 1];
            // Never silently omit crop/rotation or layer selection, even when incorrectly marked non-essential.
            Require(property.Type is not ("clap" or "irot" or "imir" or "lsel" or "a1op" or "a1lx"));
            switch (property.Type) {
                case "ispe": {
                    Require(seen.Add(property.Type));
                    int p = FullBox(property, 0, out uint flags);
                    Require(flags == 0 && property.End - p == 8);
                    width = checked((int)U32(p)); height = checked((int)U32(p + 4));
                    Require(OfficeRasterGuards.TryEnsurePixelCount(width, height, _options.MaximumDecodedPixels, out _));
                    break;
                }
                case "av1C":
                    Require(seen.Add(property.Type) && property.Length == 4 && _bytes[property.Start] == 0x81);
                    configuration = new byte[4];
                    Buffer.BlockCopy(_bytes, property.Start, configuration, 0, 4);
                    // Main profile permits 8/10-bit declarations. Pixel reconstruction is a separate gate.
                    // Tier and twelve_bit remain zero; extra configuration OBUs are outside this path.
                    Require((configuration[1] >> 5) == 0 && (configuration[2] & 0xA0) == 0 && configuration[3] == 0);
                    Require((configuration[2] & 0x0C) == 0x0C);
                    bool monochrome = (configuration[2] & 0x10) != 0;
                    Require(!alpha || monochrome);
                    if (monochrome) Require((configuration[2] & 3) == 0);
                    break;
                case "pixi": {
                    Require(seen.Add(property.Type));
                    int p = FullBox(property, 0, out uint flags);
                    Require(flags == 0);
                    var cursor = new Cursor(this, p, property.End);
                    int channels = (int)cursor.Integer(1);
                    Require(channels is 1 or 3);
                    pixelChannels = channels;
                    pixelDepth = (int)cursor.Integer(1);
                    Require(pixelDepth is 8 or 10);
                    for (int i = 1; i < channels; i++) Require(cursor.Integer(1) == (uint)pixelDepth);
                    cursor.End();
                    break;
                }
                case "colr":
                    Require(seen.Add(property.Type));
                    if (!alpha) color = ReadColorDescription(property);
                    break;
                case "auxC": {
                    Require(seen.Add(property.Type) && alpha);
                    int p = FullBox(property, 0, out uint flags);
                    Require(flags == 0);
                    var cursor = new Cursor(this, p, property.End);
                    Require(cursor.TerminatedString() == "urn:mpeg:mpegB:cicp:systems:auxiliary:alpha");
                    // Remaining aux_subtype bytes are opaque and bounded by this property box.
                    hasAuxiliaryType = true;
                    break;
                }
                default:
                    Require(!association.Essential);
                    break;
            }
        }
        Require(width > 0 && height > 0 && configuration != null && (!alpha || hasAuxiliaryType));
        // pixi may precede av1C in the association list; validate their agreement after reading both.
        Require(pixelChannels == 0 || pixelChannels == ((configuration![2] & 0x10) != 0 ? 1 : 3));
        Require(pixelDepth == 0 || pixelDepth == ((configuration![2] & 0x40) != 0 ? 10 : 8));
        return new OfficeAvifImageItem(id, width, height, location.Offset, location.Length, configuration!, alpha, color);
    }

    private OfficeAvifColorDescription ReadColorDescription(Box box) {
        var cursor = new Cursor(this, box.Start, box.End);
        Require(cursor.Type() == "nclx"); // ICC color processing must be wired deliberately, not discarded.
        int primaries = (int)cursor.Integer(2), transfer = (int)cursor.Integer(2), matrix = (int)cursor.Integer(2);
        int range = (int)cursor.Integer(1);
        Require((range & 127) == 0);
        cursor.End();
        return new OfficeAvifColorDescription(primaries, transfer, matrix, (range & 128) != 0);
    }
}
