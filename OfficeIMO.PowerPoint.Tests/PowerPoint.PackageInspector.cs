using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using DocumentFormat.OpenXml.Packaging;

namespace OfficeIMO.Tests {
    /// <summary>
    /// Lightweight diagnostics helper that writes a human-friendly manifest of
    /// parts, relationships, and content types for a PPTX. Useful when comparing
    /// our generated files against expected package shapes without needing
    /// PowerPoint on the host.
    /// </summary>
    internal static class PowerPointPackageInspector {
        public static void WriteManifest(string pptxPath, string outputDir) {
            Directory.CreateDirectory(outputDir);
            var manifestPath = Path.Combine(outputDir, Path.GetFileNameWithoutExtension(pptxPath) + ".manifest.txt");

            using var doc = PresentationDocument.Open(pptxPath, false);
            var sb = new StringBuilder();

            sb.AppendLine($"File: {pptxPath}");
            sb.AppendLine($"Slides: {doc.PresentationPart?.SlideParts.Count() ?? 0}");
            sb.AppendLine("Content Types (from package parts):");
            foreach (var pair in doc.Parts) {
                var part = pair.OpenXmlPart;
                sb.AppendLine($"  {part.ContentType} -> {part.Uri}");
            }

            sb.AppendLine("Parts and Relationships:");
            var visited = new HashSet<string>();
            foreach (var part in doc.Parts) {
                AppendPart(sb, part.OpenXmlPart, doc.GetIdOfPart(part.OpenXmlPart), indent:"  ", visited);
            }

            File.WriteAllText(manifestPath, sb.ToString());
        }

        private static void AppendPart(StringBuilder sb, OpenXmlPart part, string relId, string indent, HashSet<string> visited) {
            string key = part.Uri.ToString();
            if (!visited.Add(key)) {
                sb.AppendLine($"{indent}- {relId}: {part.Uri} ({part.ContentType}) [already listed]");
                return;
            }
            sb.AppendLine($"{indent}- {relId}: {part.Uri} ({part.ContentType})");
            foreach (var rel in part.Parts) {
                AppendPart(sb, rel.OpenXmlPart, part.GetIdOfPart(rel.OpenXmlPart), indent + "  ", visited);
            }
        }
    }
}
