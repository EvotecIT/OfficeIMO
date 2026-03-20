using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;

namespace OfficeIMO.Excel {
    public partial class ExcelSheet {
        internal void CleanupHyperlinkArtifacts() {
            var worksheet = _worksheetPart.Worksheet;
            var hyperlinks = worksheet.Elements<Hyperlinks>().FirstOrDefault();
            if (hyperlinks == null) {
                CleanupUnreferencedHyperlinkRelationships(Array.Empty<string>());
                return;
            }

            var referenceGroups = hyperlinks.Elements<Hyperlink>()
                .Where(hyperlink => !string.IsNullOrWhiteSpace(hyperlink.Reference?.Value))
                .GroupBy(hyperlink => hyperlink.Reference!.Value!, StringComparer.OrdinalIgnoreCase)
                .ToList();

            foreach (var group in referenceGroups) {
                foreach (var duplicate in group.Take(Math.Max(0, group.Count() - 1)).ToList()) {
                    RemoveHyperlinkAndUnusedRelationship(hyperlinks, duplicate);
                }
            }

            foreach (var hyperlink in hyperlinks.Elements<Hyperlink>().ToList()) {
                if (!IsValidHyperlinkReference(hyperlink.Reference?.Value) || !IsValidHyperlinkTarget(hyperlink)) {
                    RemoveHyperlinkAndUnusedRelationship(hyperlinks, hyperlink);
                }
            }

            var referencedRelationshipIds = hyperlinks.Elements<Hyperlink>()
                .Select(hyperlink => hyperlink.Id?.Value)
                .Where(id => !string.IsNullOrWhiteSpace(id))
                .Select(id => id!)
                .Distinct(StringComparer.OrdinalIgnoreCase)
                .ToList();

            CleanupUnreferencedHyperlinkRelationships(referencedRelationshipIds);
        }

        private static bool IsValidHyperlinkReference(string? reference) {
            if (string.IsNullOrWhiteSpace(reference)) {
                return false;
            }

            string candidate = reference!.Trim();
            if (A1.TryParseRange(candidate, out _, out _, out _, out _)) {
                return true;
            }

            var cell = A1.ParseCellRef(candidate);
            return cell.Row > 0 && cell.Col > 0;
        }

        private bool IsValidHyperlinkTarget(Hyperlink hyperlink) {
            string? relationshipId = hyperlink.Id?.Value;
            if (!string.IsNullOrWhiteSpace(relationshipId)) {
                return _worksheetPart.HyperlinkRelationships.Any(relationship => string.Equals(relationship.Id, relationshipId, StringComparison.OrdinalIgnoreCase));
            }

            return !string.IsNullOrWhiteSpace(hyperlink.Location?.Value);
        }

        private void RemoveHyperlinkAndUnusedRelationship(Hyperlinks hyperlinks, Hyperlink hyperlink) {
            string? relationshipId = hyperlink.Id?.Value;
            hyperlink.Remove();

            if (string.IsNullOrWhiteSpace(relationshipId)) {
                return;
            }

            if (hyperlinks.Elements<Hyperlink>().Any(existing => string.Equals(existing.Id?.Value, relationshipId, StringComparison.OrdinalIgnoreCase))) {
                return;
            }

            var relationship = _worksheetPart.HyperlinkRelationships
                .FirstOrDefault(existing => string.Equals(existing.Id, relationshipId, StringComparison.OrdinalIgnoreCase));
            if (relationship != null) {
                _worksheetPart.DeleteReferenceRelationship(relationship);
            }
        }

        private void CleanupUnreferencedHyperlinkRelationships(IEnumerable<string> referencedRelationshipIds) {
            var referenced = new HashSet<string>(referencedRelationshipIds, StringComparer.OrdinalIgnoreCase);
            foreach (var relationship in _worksheetPart.HyperlinkRelationships.ToList()) {
                if (!referenced.Contains(relationship.Id)) {
                    _worksheetPart.DeleteReferenceRelationship(relationship);
                }
            }
        }
    }
}
