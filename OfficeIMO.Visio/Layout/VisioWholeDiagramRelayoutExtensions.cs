using System;
using System.Collections.Generic;
using System.Linq;

namespace OfficeIMO.Visio {
    /// <summary>Topology-aware whole-page relayout for dense imported diagrams.</summary>
    public static class VisioWholeDiagramRelayoutExtensions {
        /// <summary>
        /// Relayouts eligible top-level shapes using connector topology. Cycles are kept
        /// together as strongly connected components while downstream nodes retain later
        /// layers; semantic surfaces remain fixed unless explicitly included.
        /// </summary>
        public static VisioPage RelayoutDiagram(this VisioPage page,
            VisioWholeDiagramRelayoutOptions? options = null) {
            if (page == null) throw new ArgumentNullException(nameof(page));
            VisioWholeDiagramRelayoutOptions resolved = options ?? new VisioWholeDiagramRelayoutOptions();
            Validate(resolved);

            List<VisioShape> nodes = page.Shapes.Where(shape =>
                    (resolved.IncludeContainers || (!shape.IsContainer && !shape.IsBackgroundSurface)) &&
                    (resolved.IncludeGroups || shape.Children.Count == 0) &&
                    !shape.IsDiagramAdornment)
                .OrderBy(shape => shape.Id, StringComparer.OrdinalIgnoreCase).ToList();
            if (nodes.Count == 0) return page;

            var nodeSet = new HashSet<VisioShape>(nodes);
            var outgoing = nodes.ToDictionary(node => node, _ => new HashSet<VisioShape>());
            foreach (VisioConnector connector in page.Connectors) {
                if (connector.From == null || connector.To == null ||
                    connector.From == connector.To || !nodeSet.Contains(connector.From) ||
                    !nodeSet.Contains(connector.To)) continue;
                outgoing[connector.From].Add(connector.To);
            }

            Dictionary<VisioShape, int> layers = BuildTopologyLayers(nodes,
                outgoing);

            double originX = nodes.Min(node => node.PinX);
            double originY = nodes.Max(node => node.PinY);
            List<IGrouping<int, VisioShape>> orderedLayers = nodes.GroupBy(node => layers[node])
                .OrderBy(group => group.Key).ToList();
            double layerCursor = resolved.LeftToRight ? originX : originY;
            foreach (IGrouping<int, VisioShape> layer in orderedLayers) {
                double layerSize = resolved.LeftToRight
                    ? layer.Max(node => node.Width)
                    : layer.Max(node => node.Height);
                double layerCenter = resolved.LeftToRight
                    ? layerCursor + layerSize / 2D
                    : layerCursor - layerSize / 2D;
                double cursor = originY;
                foreach (VisioShape node in layer.OrderBy(node => node.Id, StringComparer.OrdinalIgnoreCase)) {
                    if (resolved.LeftToRight) {
                        node.PinX = layerCenter;
                        node.PinY = cursor - node.Height / 2D;
                        cursor -= node.Height + resolved.NodeSpacing;
                    } else {
                        node.PinX = cursor + node.Width / 2D;
                        node.PinY = layerCenter;
                        cursor += node.Width + resolved.NodeSpacing;
                    }
                }
                layerCursor += resolved.LeftToRight
                    ? layerSize + resolved.LayerSpacing
                    : -(layerSize + resolved.LayerSpacing);
            }

            if (resolved.RouteConnectors) {
                foreach (VisioConnector connector in page.Connectors.Where(connector =>
                    connector.From != null && connector.To != null &&
                    nodeSet.Contains(connector.From) && nodeSet.Contains(connector.To))) {
                    connector.RouteOrthogonal();
                }
            }
            if (resolved.PolishAfterLayout) page.PolishDiagram(resolved.PolishOptions);
            return page;
        }

        /// <summary>
        /// Computes deterministic topology layers without consuming call-stack depth
        /// proportional to the graph depth.
        /// </summary>
        internal static Dictionary<VisioShape, int> BuildTopologyLayers(
            IReadOnlyList<VisioShape> nodes,
            IReadOnlyDictionary<VisioShape, HashSet<VisioShape>> outgoing) {
            int nextIndex = 0;
            var indices = new Dictionary<VisioShape, int>();
            var lowLinks = new Dictionary<VisioShape, int>();
            var stack = new Stack<VisioShape>();
            var onStack = new HashSet<VisioShape>();
            var components = new List<List<VisioShape>>();
            var traversal = new Stack<TraversalFrame>();

            void BeginVisit(VisioShape node) {
                indices[node] = nextIndex;
                lowLinks[node] = nextIndex++;
                stack.Push(node);
                onStack.Add(node);
                traversal.Push(new TraversalFrame(node, outgoing[node]
                    .OrderBy(item => item.Id, StringComparer.OrdinalIgnoreCase)
                    .ToArray()));
            }

            foreach (VisioShape node in nodes) {
                if (indices.ContainsKey(node)) continue;
                BeginVisit(node);
                while (traversal.Count > 0) {
                    TraversalFrame frame = traversal.Peek();
                    if (frame.NextTargetIndex < frame.Targets.Length) {
                        VisioShape target = frame.Targets[frame.NextTargetIndex++];
                        if (!indices.ContainsKey(target)) {
                            BeginVisit(target);
                        } else if (onStack.Contains(target)) {
                            lowLinks[frame.Node] = Math.Min(lowLinks[frame.Node],
                                indices[target]);
                        }
                        continue;
                    }

                    traversal.Pop();
                    if (lowLinks[frame.Node] == indices[frame.Node]) {
                        var component = new List<VisioShape>();
                        VisioShape member;
                        do {
                            member = stack.Pop();
                            onStack.Remove(member);
                            component.Add(member);
                        } while (!ReferenceEquals(member, frame.Node));
                        components.Add(component);
                    }

                    if (traversal.Count > 0) {
                        VisioShape parent = traversal.Peek().Node;
                        lowLinks[parent] = Math.Min(lowLinks[parent],
                            lowLinks[frame.Node]);
                    }
                }
            }

            var componentByNode = new Dictionary<VisioShape, int>();
            for (int index = 0; index < components.Count; index++) {
                foreach (VisioShape node in components[index])
                    componentByNode[node] = index;
            }
            var componentOutgoing = Enumerable.Range(0, components.Count)
                .ToDictionary(index => index, _ => new HashSet<int>());
            var incoming = Enumerable.Range(0, components.Count)
                .ToDictionary(index => index, _ => 0);
            foreach (VisioShape source in nodes) {
                int sourceComponent = componentByNode[source];
                foreach (VisioShape target in outgoing[source]) {
                    int targetComponent = componentByNode[target];
                    if (sourceComponent != targetComponent &&
                        componentOutgoing[sourceComponent].Add(targetComponent))
                        incoming[targetComponent]++;
                }
            }
            string[] keys = components.Select(component => component
                .Select(node => node.Id)
                .OrderBy(id => id, StringComparer.OrdinalIgnoreCase)
                .FirstOrDefault() ?? string.Empty)
                .ToArray();
            var ready = new SortedSet<int>(Comparer<int>.Create((left, right) => {
                int result = string.Compare(keys[left], keys[right],
                    StringComparison.OrdinalIgnoreCase);
                return result != 0 ? result : left.CompareTo(right);
            }));
            foreach (int component in incoming.Keys.Where(index =>
                         incoming[index] == 0)) ready.Add(component);
            var componentLayers = Enumerable.Range(0, components.Count)
                .ToDictionary(index => index, _ => 0);
            while (ready.Count > 0) {
                int component = ready.Min;
                ready.Remove(component);
                foreach (int target in componentOutgoing[component]
                             .OrderBy(index => keys[index],
                                 StringComparer.OrdinalIgnoreCase)) {
                    componentLayers[target] = Math.Max(componentLayers[target],
                        componentLayers[component] + 1);
                    if (--incoming[target] == 0) ready.Add(target);
                }
            }
            return nodes.ToDictionary(node => node,
                node => componentLayers[componentByNode[node]]);
        }

        private sealed class TraversalFrame {
            internal TraversalFrame(VisioShape node, VisioShape[] targets) {
                Node = node;
                Targets = targets;
            }

            internal VisioShape Node { get; }

            internal VisioShape[] Targets { get; }

            internal int NextTargetIndex { get; set; }
        }

        private static void Validate(VisioWholeDiagramRelayoutOptions options) {
            if (options.LayerSpacing <= 0D || double.IsNaN(options.LayerSpacing) || double.IsInfinity(options.LayerSpacing))
                throw new ArgumentOutOfRangeException(nameof(options), "Layer spacing must be positive and finite.");
            if (options.NodeSpacing < 0D || double.IsNaN(options.NodeSpacing) || double.IsInfinity(options.NodeSpacing))
                throw new ArgumentOutOfRangeException(nameof(options), "Node spacing must be non-negative and finite.");
            if (options.PolishOptions == null) throw new ArgumentException("Polish options cannot be null.", nameof(options));
        }
    }
}
