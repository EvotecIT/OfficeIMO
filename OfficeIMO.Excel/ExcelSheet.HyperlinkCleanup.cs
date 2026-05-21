using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;

namespace OfficeIMO.Excel {
    public partial class ExcelSheet {
        internal void CleanupHyperlinkArtifacts() {
            var worksheet = WorksheetRoot;
            var hyperlinks = worksheet.Elements<Hyperlinks>().FirstOrDefault();
            if (hyperlinks == null) {
                CleanupUnreferencedHyperlinkRelationships(new HashSet<string>(StringComparer.OrdinalIgnoreCase));
                return;
            }

            RemoveDuplicateHyperlinkReferences(hyperlinks);

            var relationshipIds = BuildHyperlinkRelationshipIdSet();

            foreach (var hyperlink in hyperlinks.Elements<Hyperlink>().ToList()) {
                if (!IsValidHyperlinkReference(hyperlink.Reference?.Value) || !IsValidHyperlinkTarget(hyperlink, relationshipIds)) {
                    hyperlink.Remove();
                }
            }

            CleanupUnreferencedHyperlinkRelationships(BuildReferencedRelationshipIdSet(hyperlinks));
        }

        private void RemoveDuplicateHyperlinkReferences(Hyperlinks hyperlinks) {
            var byReference = new Dictionary<string, Hyperlink>(StringComparer.OrdinalIgnoreCase);
            List<Hyperlink>? duplicates = null;

            foreach (var hyperlink in hyperlinks.Elements<Hyperlink>()) {
                string? reference = hyperlink.Reference?.Value;
                if (string.IsNullOrWhiteSpace(reference)) {
                    continue;
                }

                if (byReference.TryGetValue(reference, out var previous)) {
                    (duplicates ??= new List<Hyperlink>()).Add(previous);
                }

                byReference[reference] = hyperlink;
            }

            if (duplicates == null) {
                return;
            }

            foreach (var duplicate in duplicates) {
                duplicate.Remove();
            }
        }

        private HashSet<string> BuildHyperlinkRelationshipIdSet() {
            var relationshipIds = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
            foreach (var relationship in _worksheetPart.HyperlinkRelationships) {
                relationshipIds.Add(relationship.Id);
            }

            return relationshipIds;
        }

        private static HashSet<string> BuildReferencedRelationshipIdSet(Hyperlinks hyperlinks) {
            var referencedRelationshipIds = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
            foreach (var hyperlink in hyperlinks.Elements<Hyperlink>()) {
                string? id = hyperlink.Id?.Value;
                if (!string.IsNullOrWhiteSpace(id)) {
                    referencedRelationshipIds.Add(id);
                }
            }

            return referencedRelationshipIds;
        }

        private static bool IsValidHyperlinkReference(string? reference) {
            if (string.IsNullOrWhiteSpace(reference)) {
                return false;
            }

            foreach (ReferenceListPart part in SplitReferenceList(reference!.Trim())) {
                if (!TryParseReference(part, out _)) {
                    return false;
                }
            }

            return true;
        }

        private static bool IsValidHyperlinkTarget(Hyperlink hyperlink, HashSet<string> relationshipIds) {
            string? relationshipId = hyperlink.Id?.Value;
            if (!string.IsNullOrWhiteSpace(relationshipId)) {
                return relationshipIds.Contains(relationshipId);
            }

            return !string.IsNullOrWhiteSpace(hyperlink.Location?.Value);
        }

        private void CleanupUnreferencedHyperlinkRelationships(HashSet<string> referencedRelationshipIds) {
            foreach (var relationship in _worksheetPart.HyperlinkRelationships.ToList()) {
                if (!referencedRelationshipIds.Contains(relationship.Id)) {
                    _worksheetPart.DeleteReferenceRelationship(relationship);
                }
            }
        }
    }
}
