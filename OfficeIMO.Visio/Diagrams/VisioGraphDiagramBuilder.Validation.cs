using System;
using System.Collections.Generic;
using System.Linq;
using OfficeIMO.Visio.Stencils;
using Color = OfficeIMO.Drawing.OfficeColor;

namespace OfficeIMO.Visio.Diagrams {
    /// <summary>
    /// High-level builder for generic graph diagrams where OfficeIMO lays out arbitrary nodes and edges.
    /// </summary>
    public sealed partial class VisioGraphDiagramBuilder {
        private void AddRoot(string id) {
            if (_rootIdSet.Add(id)) {
                _rootIds.Add(id);
            }
        }

        private bool IsIdInUse(string id) {
            if (!string.IsNullOrWhiteSpace(_titleText) && string.Equals(_titleId, id, StringComparison.Ordinal)) {
                return true;
            }

            if (_nodesById.ContainsKey(id) || _zoneIds.Contains(id) || _edgeIds.Contains(id)) {
                return true;
            }

            return false;
        }

        private string CreateGeneratedId(string baseId) {
            string id = baseId;
            int index = 2;
            while (IsIdInUse(id) || _generatedIds.Contains(id)) {
                id = baseId + "-" + index;
                index++;
            }

            _generatedIds.Add(id);
            return id;
        }

        private void EnsureKnownNode(string id, string parameterName) {
            if (!_nodesById.ContainsKey(id)) {
                throw new ArgumentException($"Unknown graph node id '{id}'.", parameterName);
            }
        }

        private NodeItem GetKnownNode(string id, string parameterName) {
            if (_nodesById.TryGetValue(id, out NodeItem? node)) {
                return node;
            }

            throw new ArgumentException($"Unknown graph node id '{id}'.", parameterName);
        }

        private EdgeItem GetKnownEdge(string id, string parameterName) {
            if (_edgesById.TryGetValue(id, out EdgeItem? edge)) {
                return edge;
            }

            throw new ArgumentException($"Unknown graph edge id '{id}'.", parameterName);
        }

        private ZoneItem GetKnownZone(string id, string parameterName) {
            if (_zonesById.TryGetValue(id, out ZoneItem? zone)) {
                return zone;
            }

            throw new ArgumentException($"Unknown graph zone id '{id}'.", parameterName);
        }

        private static IReadOnlyList<string> NormalizeZoneNodeIds(IEnumerable<string> nodeIds, string parameterName, string label) {
            if (nodeIds == null) throw new ArgumentNullException(nameof(nodeIds));
            List<string> normalizedNodeIds = new();
            HashSet<string> seen = new(StringComparer.Ordinal);
            foreach (string nodeId in nodeIds) {
                string normalizedId = RequireId(nodeId, parameterName, label);
                if (seen.Add(normalizedId)) {
                    normalizedNodeIds.Add(normalizedId);
                }
            }

            if (normalizedNodeIds.Count == 0) {
                throw new ArgumentException("A graph zone or cluster requires at least one node id.", parameterName);
            }

            return normalizedNodeIds.AsReadOnly();
        }

        private static string RequireId(string id, string parameterName, string label) {
            if (string.IsNullOrWhiteSpace(id)) {
                throw new ArgumentException(label + " cannot be null or whitespace.", parameterName);
            }

            return id.Trim();
        }

        private static string SlugId(string value) {
            if (string.IsNullOrWhiteSpace(value)) {
                return "item";
            }

            char[] characters = value.Trim().Select(ch => char.IsLetterOrDigit(ch) ? char.ToLowerInvariant(ch) : '-').ToArray();
            string slug = new(characters);
            while (slug.Contains("--")) {
                slug = slug.Replace("--", "-");
            }

            slug = slug.Trim('-');
            return string.IsNullOrWhiteSpace(slug) ? "item" : slug;
        }

        private static void ValidatePositive(double value, string parameterName) {
            if (double.IsNaN(value) || double.IsInfinity(value) || value <= 0D) {
                throw new ArgumentOutOfRangeException(parameterName, "Value must be positive.");
            }
        }

        private static void ValidateNonNegative(double value, string parameterName) {
            if (double.IsNaN(value) || double.IsInfinity(value) || value < 0D) {
                throw new ArgumentOutOfRangeException(parameterName, "Value must be zero or greater.");
            }
        }
    }
}
