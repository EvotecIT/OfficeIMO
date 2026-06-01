using System;
using System.Collections.Generic;
using OfficeIMO.Visio.Stencils;

namespace OfficeIMO.Visio.Diagrams {
    /// <summary>
    /// Simple data record used to import graph nodes into <see cref="VisioGraphDiagramBuilder"/>.
    /// </summary>
    public sealed class VisioGraphNodeRecord {
        /// <summary>
        /// Initializes a graph node import record.
        /// </summary>
        public VisioGraphNodeRecord(string id, string text) {
            if (string.IsNullOrWhiteSpace(id)) throw new ArgumentException("Node id cannot be null or whitespace.", nameof(id));
            Id = id;
            Text = text ?? string.Empty;
        }

        /// <summary>Stable node id used in the generated Visio page.</summary>
        public string Id { get; }

        /// <summary>Node text.</summary>
        public string Text { get; }

        /// <summary>Fallback graph node kind when no stencil is supplied.</summary>
        public VisioGraphNodeKind Kind { get; set; } = VisioGraphNodeKind.Process;

        /// <summary>Whether this node should be treated as a graph root.</summary>
        public bool IsRoot { get; set; }

        /// <summary>Optional concrete stencil shape for this node.</summary>
        public VisioStencilShape? Stencil { get; set; }

        /// <summary>Optional stencil catalog used with <see cref="StencilQueries"/>.</summary>
        public VisioStencilCatalog? StencilCatalog { get; set; }

        /// <summary>Ordered stencil lookup/search queries used when <see cref="StencilCatalog"/> is set.</summary>
        public IList<string> StencilQueries { get; } = new List<string>();

        /// <summary>Shape Data rows to apply to the generated node.</summary>
        public IDictionary<string, string?> ShapeData { get; } = new Dictionary<string, string?>(StringComparer.OrdinalIgnoreCase);

        /// <summary>Optional hyperlink address to attach to the generated node.</summary>
        public string? HyperlinkAddress { get; set; }

        /// <summary>Optional hyperlink description.</summary>
        public string? HyperlinkDescription { get; set; }

        /// <summary>Optional hyperlink sub-address.</summary>
        public string? HyperlinkSubAddress { get; set; }
    }

    /// <summary>
    /// Simple data record used to import graph edges into <see cref="VisioGraphDiagramBuilder"/>.
    /// </summary>
    public sealed class VisioGraphEdgeRecord {
        /// <summary>
        /// Initializes a graph edge import record.
        /// </summary>
        public VisioGraphEdgeRecord(string fromId, string toId) : this(null, fromId, toId) {
        }

        /// <summary>
        /// Initializes a named graph edge import record.
        /// </summary>
        public VisioGraphEdgeRecord(string? id, string fromId, string toId) {
            if (string.IsNullOrWhiteSpace(fromId)) throw new ArgumentException("From node id cannot be null or whitespace.", nameof(fromId));
            if (string.IsNullOrWhiteSpace(toId)) throw new ArgumentException("To node id cannot be null or whitespace.", nameof(toId));
            Id = string.IsNullOrWhiteSpace(id) ? null : id;
            FromId = fromId;
            ToId = toId;
        }

        /// <summary>Stable edge id. When null, the builder derives a diff-friendly id from endpoints and kind.</summary>
        public string? Id { get; }

        /// <summary>Source node id.</summary>
        public string FromId { get; }

        /// <summary>Target node id.</summary>
        public string ToId { get; }

        /// <summary>Visual edge kind.</summary>
        public VisioGraphConnectorKind Kind { get; set; } = VisioGraphConnectorKind.Standard;

        /// <summary>Optional connector label.</summary>
        public string? Label { get; set; }

        /// <summary>Whether the generated edge should be directed.</summary>
        public bool Directed { get; set; } = true;

        /// <summary>Shape Data rows to apply to the generated connector.</summary>
        public IDictionary<string, string?> ShapeData { get; } = new Dictionary<string, string?>(StringComparer.OrdinalIgnoreCase);

        /// <summary>Optional hyperlink address to attach to the generated connector.</summary>
        public string? HyperlinkAddress { get; set; }

        /// <summary>Optional hyperlink description.</summary>
        public string? HyperlinkDescription { get; set; }

        /// <summary>Optional hyperlink sub-address.</summary>
        public string? HyperlinkSubAddress { get; set; }
    }

    /// <summary>
    /// Simple data record used to import graph clusters into <see cref="VisioGraphDiagramBuilder"/>.
    /// </summary>
    public sealed class VisioGraphClusterRecord {
        /// <summary>
        /// Initializes a graph cluster import record.
        /// </summary>
        public VisioGraphClusterRecord(string id, string text, IEnumerable<string> nodeIds) {
            if (string.IsNullOrWhiteSpace(id)) throw new ArgumentException("Cluster id cannot be null or whitespace.", nameof(id));
            if (nodeIds == null) throw new ArgumentNullException(nameof(nodeIds));

            Id = id;
            Text = text ?? string.Empty;
            foreach (string nodeId in nodeIds) {
                if (string.IsNullOrWhiteSpace(nodeId)) {
                    throw new ArgumentException("Cluster node ids cannot contain null or whitespace values.", nameof(nodeIds));
                }

                NodeIds.Add(nodeId);
            }

            if (NodeIds.Count == 0) {
                throw new ArgumentException("A graph cluster requires at least one node id.", nameof(nodeIds));
            }
        }

        /// <summary>Stable cluster id used for the generated background shape.</summary>
        public string Id { get; }

        /// <summary>Cluster caption text.</summary>
        public string Text { get; }

        /// <summary>Node ids contained by the cluster.</summary>
        public IList<string> NodeIds { get; } = new List<string>();

        /// <summary>Shape Data rows to apply to the generated cluster background.</summary>
        public IDictionary<string, string?> ShapeData { get; } = new Dictionary<string, string?>(StringComparer.OrdinalIgnoreCase);

        /// <summary>Optional hyperlink address to attach to the generated cluster background.</summary>
        public string? HyperlinkAddress { get; set; }

        /// <summary>Optional hyperlink description.</summary>
        public string? HyperlinkDescription { get; set; }

        /// <summary>Optional hyperlink sub-address.</summary>
        public string? HyperlinkSubAddress { get; set; }
    }
}
