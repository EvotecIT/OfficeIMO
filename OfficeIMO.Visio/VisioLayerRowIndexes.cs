using System.Collections.Generic;
using System.Linq;

namespace OfficeIMO.Visio;

/// <summary>Keeps native layer row identities separate from ordinal layer membership.</summary>
internal static class VisioLayerRowIndexes {
    internal static int[] Create(IReadOnlyList<VisioLayer> layers) {
        var reserved = new HashSet<int>(layers.Where(layer => layer.SourceIndex.HasValue).Select(layer => layer.SourceIndex!.Value));
        var used = new HashSet<int>();
        var indexes = new int[layers.Count];
        for (int i = 0; i < layers.Count; i++) {
            if (layers[i].SourceIndex is int original && used.Add(original)) {
                indexes[i] = original;
                continue;
            }
            int index = i;
            while (reserved.Contains(index) || used.Contains(index)) index++;
            indexes[i] = index;
            used.Add(index);
        }
        return indexes;
    }
}
