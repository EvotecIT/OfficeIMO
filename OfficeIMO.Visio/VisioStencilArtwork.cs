using System;
using System.Collections.Generic;
using System.Text.RegularExpressions;
using Color = OfficeIMO.Drawing.OfficeColor;

namespace OfficeIMO.Visio {
    internal static class VisioStencilArtwork {
        private static readonly Regex TokenPattern = new("[a-z0-9]+", RegexOptions.Compiled | RegexOptions.CultureInvariant);

        internal static string? GetKey(VisioShape shape) {
            if (shape == null) {
                throw new ArgumentNullException(nameof(shape));
            }

            if (ShouldSkip(shape)) {
                return null;
            }

            string metadata = ((shape.GetUserCellValue(VisioSemanticUserCells.StencilId) ?? string.Empty) + " " +
                               (shape.GetUserCellValue(VisioSemanticUserCells.StencilName) ?? string.Empty) + " " +
                               (shape.GetUserCellValue(VisioSemanticUserCells.StencilCategory) ?? string.Empty) + " " +
                               (shape.GetUserCellValue(VisioSemanticUserCells.StencilAliases) ?? string.Empty) + " " +
                               (shape.GetUserCellValue(VisioSemanticUserCells.StencilTags) ?? string.Empty) + " " +
                               (shape.MasterNameU ?? shape.NameU ?? string.Empty)).ToLowerInvariant();

            if (string.IsNullOrWhiteSpace(metadata)) {
                return null;
            }

            HashSet<string> tokens = Tokenize(metadata);
            if (ContainsAny(tokens, "person", "user", "actor", "client", "customer", "principal")) return "person";
            if (ContainsAny(tokens, "database", "sql", "data", "lake", "warehouse", "storage", "catalog", "audit", "document", "record")) return "data";
            if (ContainsAny(tokens, "security", "policy", "firewall", "identity", "secret", "key", "trust", "risk", "waf", "vault")) return "security";
            if (ContainsAny(tokens, "cloud", "subscription", "tenant")) return "cloud";
            if (ContainsAny(tokens, "k8s", "kubernetes", "pod", "container", "cluster", "namespace")) return "container";
            if (ContainsAny(tokens, "queue", "stream", "event", "pipeline", "bus")) return "event";
            if (ContainsAny(tokens, "monitor", "observability", "metrics", "telemetry", "logs")) return "monitoring";
            if (ContainsAny(tokens, "server", "serverless", "compute", "service", "api", "app", "application", "function", "worker", "gateway", "endpoint", "ingress", "appliance", "workstation", "module")) return "compute";
            return null;
        }

        internal static Color ResolveColor(VisioShape shape, byte alpha) {
            double brightness = ((shape.FillColor.R * 299D) + (shape.FillColor.G * 587D) + (shape.FillColor.B * 114D)) / 1000D;
            if (shape.FillPattern != 0 && shape.FillColor.A > 0 && brightness < 120D) {
                return Color.FromRgba(255, 255, 255, alpha);
            }

            Color line = shape.LineColor.A > 0 ? shape.LineColor : Color.FromRgb(31, 41, 55);
            return Color.FromRgba(line.R, line.G, line.B, alpha);
        }

        private static bool ShouldSkip(VisioShape shape) {
            string? kind = shape.GetUserCellValue(VisioSemanticUserCells.Kind);
            return string.Equals(kind, VisioSemanticUserCells.SequenceFragmentKind, StringComparison.OrdinalIgnoreCase) ||
                   string.Equals(kind, VisioSemanticUserCells.BackgroundSurfaceKind, StringComparison.OrdinalIgnoreCase) ||
                   string.Equals(kind, VisioSemanticUserCells.DiagramAdornmentKind, StringComparison.OrdinalIgnoreCase);
        }

        private static HashSet<string> Tokenize(string value) {
            HashSet<string> tokens = new(StringComparer.OrdinalIgnoreCase);
            foreach (Match match in TokenPattern.Matches(value)) {
                if (match.Value.Length > 0) {
                    tokens.Add(match.Value);
                }
            }

            return tokens;
        }

        private static bool ContainsAny(HashSet<string> tokens, params string[] terms) {
            for (int i = 0; i < terms.Length; i++) {
                if (tokens.Contains(terms[i])) {
                    return true;
                }
            }

            return false;
        }
    }
}
