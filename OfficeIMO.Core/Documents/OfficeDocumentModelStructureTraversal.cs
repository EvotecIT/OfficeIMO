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
        for (int index = roots.Count - 1; index >= 0; index--) {
            stack.Push(new TraversalFrame(roots[index], 1, isExit: false));
        }

        while (stack.Count > 0) {
            TraversalFrame frame = stack.Pop();
            OfficeDocumentModelNode? node = frame.Node;
            if (node == null) throw new InvalidDataException("The shared document structure contains a null node.");
            if (frame.IsExit) {
                ancestry.Remove(node);
                continue;
            }
            if (frame.Depth > maxDepth) {
                throw new InvalidDataException($"The shared document structure exceeds MaxStructureDepth ({maxDepth}).");
            }
            if (!ancestry.Add(node)) {
                throw new InvalidDataException("The shared document structure contains a reference cycle.");
            }
            if (flattened.Count >= maxNodes) {
                throw new InvalidDataException($"The shared document structure exceeds MaxStructureNodes ({maxNodes}).");
            }
            flattened.Add(node);
            stack.Push(new TraversalFrame(node, frame.Depth, isExit: true));
            IReadOnlyList<OfficeDocumentModelNode>? children = node.Children;
            if (children == null) throw new InvalidDataException("A shared document structure node has a null Children collection.");
            for (int index = children.Count - 1; index >= 0; index--) {
                stack.Push(new TraversalFrame(children[index], frame.Depth + 1, isExit: false));
            }
        }
        return flattened;
    }

    private readonly struct TraversalFrame {
        internal TraversalFrame(OfficeDocumentModelNode? node, int depth, bool isExit) {
            Node = node;
            Depth = depth;
            IsExit = isExit;
        }

        internal OfficeDocumentModelNode? Node { get; }
        internal int Depth { get; }
        internal bool IsExit { get; }
    }

    private sealed class ReferenceComparer : IEqualityComparer<OfficeDocumentModelNode> {
        internal static ReferenceComparer Instance { get; } = new ReferenceComparer();

        public bool Equals(OfficeDocumentModelNode? left, OfficeDocumentModelNode? right) => ReferenceEquals(left, right);

        public int GetHashCode(OfficeDocumentModelNode value) => RuntimeHelpers.GetHashCode(value);
    }
}
