namespace OfficeIMO.Word {
    public sealed partial class WordFeatureReport {
        /// <summary>Returns actionable repair or routing hints for an unavailable capability.</summary>
        public IReadOnlyList<WordPreflightRepairHint> GetRepairHints(WordPreflightCapability capability) {
            if (Can(capability)) return Array.Empty<WordPreflightRepairHint>();
            var hints = new List<WordPreflightRepairHint>();
            foreach (WordFeatureFinding finding in GetCapabilityFindings(capability)) {
                AddRepairHint(hints, capability, finding);
            }
            return hints
                .GroupBy(hint => hint.FeatureName + "\u001f" + hint.Action + "\u001f" + (hint.Command ?? string.Empty),
                    StringComparer.Ordinal)
                .Select(group => group.First())
                .ToArray();
        }

        private static void AddRepairHint(List<WordPreflightRepairHint> hints,
            WordPreflightCapability capability, WordFeatureFinding finding) {
            switch (finding.Name) {
                case "Digital signatures":
                    hints.Add(new WordPreflightRepairHint(capability, finding.Name,
                        "Use a read-only workflow, work on an unsigned copy, or explicitly accept signature invalidation.",
                        "WordSaveOptions.SignedDocumentPolicy",
                        "A package mutation cannot preserve the validity of an existing digital signature."));
                    break;
                case "Alternative format imports":
                    hints.Add(new WordPreflightRepairHint(capability, finding.Name,
                        "Materialize alternative-format content in Word-compatible software before fixed-layout export, or extract and replace it with native Word content.",
                        "ExtractEmbeddedDocument(...) / RemoveEmbeddedDocument(...)"));
                    break;
                case "External linked images":
                    hints.Add(new WordPreflightRepairHint(capability, finding.Name,
                        "Embed linked image bytes before offline fixed-layout export.",
                        "AddImage(...) or replace the external image relationship",
                        "Dependency-free export does not fetch external resources."));
                    break;
                case "SmartArt":
                    hints.Add(new WordPreflightRepairHint(capability, finding.Name,
                        "Accept the dependency-free SmartArt fallback, flatten the diagram to an image, or route exact layout through Word-compatible software.",
                        "ExportImages(...)"));
                    break;
                case "ActiveX controls":
                    hints.Add(new WordPreflightRepairHint(capability, finding.Name,
                        "Remove or flatten ActiveX UI before fixed-layout export, or use a preserve-only route.",
                        "GetEmbeddedPayloads() / RemoveEmbeddedPayload(...)"));
                    break;
                case "Attached templates":
                    hints.Add(new WordPreflightRepairHint(capability, finding.Name,
                        "Detach the external template or use a content-only edit that leaves the relationship untouched.",
                        null, finding.Note));
                    break;
                default:
                    if (finding.SupportLevel == WordFeatureSupportLevel.Preserved
                        || finding.SupportLevel == WordFeatureSupportLevel.Unsupported) {
                        hints.Add(new WordPreflightRepairHint(capability, finding.Name,
                            "Use a preserve-only workflow, remove the feature, or route the document through Word-compatible software.",
                            "SaveCopy(...) or package-preserving save", finding.Note));
                    }
                    break;
            }
        }
    }
}
