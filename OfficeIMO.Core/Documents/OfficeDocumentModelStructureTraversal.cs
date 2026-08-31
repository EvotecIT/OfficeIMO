using System;
using System.Collections.Generic;
using System.IO;
using System.Runtime.CompilerServices;

namespace OfficeIMO.Core.Internal;

/// <summary>Validates caller-created recursive document structures without recursive traversal.</summary>
internal static class OfficeDocumentModelStructureTraversal {
    internal const int MaximumSupportedDepth = 256;

    internal static IReadOnlyList<OfficeDocumentModelNode> ValidateAndFlatten(
        IReadOnlyList<OfficeDocumentModelNode> roots,
        int maxDepth,
        int maxNodes) {
        if (roots == null) throw new ArgumentNullException(nameof(roots));
        if (maxDepth < 1 || maxDepth > MaximumSupportedDepth) throw new ArgumentOutOfRangeException(nameof(maxDepth));
        if (maxNodes < 1) throw new ArgumentOutOfRangeException(nameof(maxNodes));

        var flattened = new List<OfficeDocumentModelNode>(Math.Min(maxNodes, 4_096));
        var ancestry = new HashSet<OfficeDocumentModelNode>(ReferenceComparer.Instance);
        var stack = new Stack<TraversalFrame>();
        stack.Push(new TraversalFrame(roots, 0, 1, owner: null));

        while (stack.Count > 0) {
            TraversalFrame frame = stack.Pop();
            if (frame.Index >= frame.Nodes.Count) {
                if (frame.Owner != null) ancestry.Remove(frame.Owner);
                continue;
            }
            if (frame.Depth > maxDepth) {
                throw new InvalidDataException($"The shared document structure exceeds MaxStructureDepth ({maxDepth}).");
            }
            if (flattened.Count >= maxNodes) {
                throw new InvalidDataException($"The shared document structure exceeds MaxStructureNodes ({maxNodes}).");
            }
            OfficeDocumentModelNode? node = frame.Nodes[frame.Index];
            stack.Push(new TraversalFrame(frame.Nodes, frame.Index + 1, frame.Depth, frame.Owner));
            if (node == null) throw new InvalidDataException("The shared document structure contains a null node.");
            if (!ancestry.Add(node)) {
                throw new InvalidDataException("The shared document structure contains a reference cycle.");
            }
            flattened.Add(node);
            IReadOnlyList<OfficeDocumentModelNode>? children = node.Children;
            if (children == null) throw new InvalidDataException("A shared document structure node has a null Children collection.");
            if (children.Count == 0) ancestry.Remove(node);
            else stack.Push(new TraversalFrame(children, 0, frame.Depth + 1, node));
        }
        return flattened;
    }

    private readonly struct TraversalFrame {
        internal TraversalFrame(IReadOnlyList<OfficeDocumentModelNode> nodes, int index, int depth, OfficeDocumentModelNode? owner) {
            Nodes = nodes;
            Index = index;
            Depth = depth;
            Owner = owner;
        }

        internal IReadOnlyList<OfficeDocumentModelNode> Nodes { get; }
        internal int Index { get; }
        internal int Depth { get; }
        internal OfficeDocumentModelNode? Owner { get; }
    }

    private sealed class ReferenceComparer : IEqualityComparer<OfficeDocumentModelNode> {
        internal static ReferenceComparer Instance { get; } = new ReferenceComparer();

        public bool Equals(OfficeDocumentModelNode? left, OfficeDocumentModelNode? right) => ReferenceEquals(left, right);

        public int GetHashCode(OfficeDocumentModelNode value) => RuntimeHelpers.GetHashCode(value);
    }
}
