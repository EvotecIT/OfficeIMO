using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.Globalization;
using System.Linq;
using System.Xml.Linq;
using OfficeIMO.Drawing;

namespace OfficeIMO.PowerPoint {
    /// <summary>One editable node in an imported SmartArt topology.</summary>
    public sealed class PowerPointSmartArtNode {
        /// <summary>Creates a topology node.</summary>
        public PowerPointSmartArtNode(string id, string text, string? parentId,
            uint order) {
            Id = id ?? throw new ArgumentNullException(nameof(id));
            Text = text ?? throw new ArgumentNullException(nameof(text));
            ParentId = parentId;
            Order = order;
        }

        /// <summary>Stable producer model identifier.</summary>
        public string Id { get; }
        /// <summary>Editable node text.</summary>
        public string Text { get; set; }
        /// <summary>Parent node identifier, or null for a semantic root.</summary>
        public string? ParentId { get; set; }
        /// <summary>Zero-based sibling order.</summary>
        public uint Order { get; set; }
    }

    /// <summary>Typed topology snapshot for a safely editable SmartArt family.</summary>
    public sealed class PowerPointSmartArtTopology {
        internal PowerPointSmartArtTopology(OfficeDiagramKind kind,
            IReadOnlyList<PowerPointSmartArtNode> nodes) {
            Kind = kind;
            Nodes = new ReadOnlyCollection<PowerPointSmartArtNode>(nodes.ToList());
        }
        /// <summary>Semantic diagram family.</summary>
        public OfficeDiagramKind Kind { get; }
        /// <summary>Nodes in semantic traversal order.</summary>
        public IReadOnlyList<PowerPointSmartArtNode> Nodes { get; }
    }

    public partial class PowerPointSmartArt {
        /// <summary>
        /// Tries to project imported layout and parent/child connections into a
        /// topology that OfficeIMO can edit without changing diagram meaning.
        /// </summary>
        public bool TryGetTopology(out PowerPointSmartArtTopology topology,
            out string diagnostic) {
            try {
                var (xdoc, ns, textBodies, _) = LoadNodeTextBodiesWithPart();
                XElement? properties = xdoc.Descendants(ns.dgm + "prSet")
                    .FirstOrDefault(element => element.Attribute("loCatId") != null ||
                                               element.Attribute("loTypeId") != null);
                string category = ((string?)properties?.Attribute("loCatId") ??
                    (string?)properties?.Attribute("loTypeId") ?? string.Empty)
                    .ToLowerInvariant();
                if (!TryResolveDiagramKind(category, out OfficeDiagramKind kind)) {
                    topology = null!;
                    diagnostic = "The producer-specific SmartArt layout is preservation-only.";
                    return false;
                }
                if (!HasRepresentableLayoutDefinition(kind, textBodies.Count)) {
                    topology = null!;
                    diagnostic = "The producer SmartArt layout definition is not the canonical editable layout for this semantic family and remains preservation-only.";
                    return false;
                }
                if (!TryCreateSemanticNodeMap(xdoc, ns, textBodies,
                        out Dictionary<string, (int Index, XElement TextBody)> nodeById,
                        out HashSet<string> documentIds,
                        out Dictionary<string, string> parentByNode,
                        out Dictionary<string, uint> sourceOrderByNode) ||
                    documentIds.Count != 1 ||
                    parentByNode.Count != nodeById.Count) {
                    topology = null!;
                    diagnostic = "The imported SmartArt connection topology cannot be represented safely and was left unchanged.";
                    return false;
                }
                List<KeyValuePair<string, (int Index, XElement TextBody)>> ordered =
                    OrderSemanticNodes(nodeById, parentByNode, sourceOrderByNode,
                        documentIds).ToList();
                var nodes = ordered.Select(node => new PowerPointSmartArtNode(
                    node.Key, ReadNodeText(node.Value.TextBody, ns.a),
                    documentIds.Contains(parentByNode[node.Key]) ? null : parentByNode[node.Key],
                    sourceOrderByNode.TryGetValue(node.Key, out uint order) ? order : 0U)).ToList();
                ValidateTopology(kind, nodes);
                topology = new PowerPointSmartArtTopology(kind, nodes);
                diagnostic = string.Empty;
                return true;
            } catch (Exception ex) when (ex is InvalidOperationException ||
                                         ex is FormatException) {
                topology = null!;
                diagnostic = ex.Message;
                return false;
            }
        }

        /// <summary>
        /// Updates text, sibling order, and parent connections in place. Node creation,
        /// deletion, unsupported layouts, cycles, and meaning-changing topologies are rejected.
        /// </summary>
        public PowerPointSmartArt UpdateTopology(
            IEnumerable<PowerPointSmartArtNode> nodes) {
            if (nodes == null) throw new ArgumentNullException(nameof(nodes));
            if (!TryGetTopology(out PowerPointSmartArtTopology current,
                    out string diagnostic)) throw new NotSupportedException(diagnostic);
            List<PowerPointSmartArtNode> requested = nodes.ToList();
            if (requested.Select(node => node.Id).Distinct(StringComparer.Ordinal).Count() != requested.Count)
                throw new ArgumentException("SmartArt node identifiers must be unique.", nameof(nodes));
            if (!new HashSet<string>(requested.Select(node => node.Id), StringComparer.Ordinal)
                .SetEquals(current.Nodes.Select(node => node.Id)))
                throw new NotSupportedException("Adding or removing imported SmartArt nodes is not supported because it can change diagram meaning.");
            ValidateTopology(current.Kind, requested);

            var (xdoc, ns, textBodies, dataPart) = LoadNodeTextBodiesWithPart();
            if (!TryCreateSemanticNodeMap(xdoc, ns, textBodies,
                    out Dictionary<string, (int Index, XElement TextBody)> nodeById,
                    out HashSet<string> documentIds,
                    out _, out _))
                throw new InvalidOperationException("The imported SmartArt topology changed while it was being edited.");
            string documentId = documentIds.Single();
            Dictionary<string, XElement> connections = xdoc.Descendants(ns.dgm + "cxn")
                .Where(connection => {
                    string? destination = (string?)connection.Attribute("destId");
                    return destination != null && nodeById.ContainsKey(destination);
                }).ToDictionary(connection => (string)connection.Attribute("destId")!,
                    StringComparer.Ordinal);
            foreach (PowerPointSmartArtNode node in requested) {
                PowerPointXmlValueValidator.ValidateCharacters(node.Text,
                    nameof(nodes), "SmartArt node text");
                if (string.IsNullOrWhiteSpace(node.Text))
                    throw new ArgumentException("SmartArt node text cannot be empty.", nameof(nodes));
                XElement textBody = nodeById[node.Id].TextBody;
                string currentText = ReadNodeText(textBody, ns.a);
                if (!string.Equals(currentText, node.Text, StringComparison.Ordinal)) {
                    XElement[] textRuns = textBody.Descendants(ns.a + "t").ToArray();
                    bool containsBreaks = textBody.Descendants(ns.a + "br").Any();
                    if (textRuns.Length != 1 || containsBreaks ||
                        textBody.Elements(ns.a + "p").Count() != 1) {
                        throw new NotSupportedException(
                            $"SmartArt node '{node.Id}' contains rich producer text. Topology changes are supported, but replacing that text would discard formatting or paragraph meaning.");
                    }
                    textRuns[0].Value = node.Text;
                }
                XElement connection = connections[node.Id];
                connection.SetAttributeValue("srcId", node.ParentId ?? documentId);
                connection.SetAttributeValue("srcOrd", node.Order.ToString(CultureInfo.InvariantCulture));
            }
            SaveDiagramData(dataPart, xdoc);
            return this;
        }

        private static void ValidateTopology(OfficeDiagramKind kind,
            IReadOnlyList<PowerPointSmartArtNode> nodes) {
            if (nodes.Count == 0) throw new InvalidOperationException("SmartArt must contain at least one semantic node.");
            var byId = nodes.ToDictionary(node => node.Id, StringComparer.Ordinal);
            foreach (PowerPointSmartArtNode node in nodes) {
                if (node.ParentId != null && !byId.ContainsKey(node.ParentId))
                    throw new InvalidOperationException($"SmartArt parent '{node.ParentId}' does not exist.");
                var seen = new HashSet<string>(StringComparer.Ordinal) { node.Id };
                string? parent = node.ParentId;
                while (parent != null) {
                    if (!seen.Add(parent)) throw new InvalidOperationException("SmartArt parent connections contain a cycle.");
                    parent = byId[parent].ParentId;
                }
            }
            PowerPointSmartArtNode[] roots = nodes.Where(node => node.ParentId == null).ToArray();
            if (kind == OfficeDiagramKind.Hierarchy) {
                if (roots.Length != 1) throw new InvalidOperationException("Hierarchy SmartArt requires exactly one root.");
            } else if (kind == OfficeDiagramKind.Relationship) {
                if (roots.Length != 1 || nodes.Any(node => node.ParentId != null && node.ParentId != roots[0].Id))
                    throw new InvalidOperationException("Relationship SmartArt requires one center with direct children.");
            } else if (nodes.Any(node => node.ParentId != null)) {
                throw new InvalidOperationException($"{kind} SmartArt supports ordered root nodes but not parent/child nesting.");
            }
            foreach (IGrouping<string?, PowerPointSmartArtNode> siblings in nodes.GroupBy(node => node.ParentId, StringComparer.Ordinal)) {
                if (siblings.Select(node => node.Order).Distinct().Count() != siblings.Count())
                    throw new InvalidOperationException("SmartArt sibling order values must be unique.");
            }
        }
    }
}
