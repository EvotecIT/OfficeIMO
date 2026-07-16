namespace OfficeIMO.Word {
    /// <summary>Workflow-level Word operations covered by feature preflight checks.</summary>
    public enum WordPreflightCapability {
        /// <summary>Read paragraphs, tables, fields, reviews, and other supported document content.</summary>
        ReadDocumentContent,

        /// <summary>Edit ordinary document content while preserving package parts not fully authored by OfficeIMO.</summary>
        EditDocumentContent,

        /// <summary>Perform structure-changing edits across sections, relationships, and package-backed content.</summary>
        EditDocumentStructure,

        /// <summary>Bind values into a Word template and save the generated document.</summary>
        BindTemplate,

        /// <summary>Render the document through first-party fixed-layout image or PDF paths.</summary>
        RenderFixedLayout,

        /// <summary>Save a package round trip without a known unsupported mutation contract.</summary>
        SavePackageRoundTrip
    }

    public sealed partial class WordFeatureReport {
        private static readonly string[] FixedLayoutBlockerNames = {
            "Alternative format imports",
            "External linked images",
            "SmartArt",
            "ActiveX controls"
        };

        /// <summary>True when core document content can be read despite non-content package metadata.</summary>
        public bool CanReadDocumentContent => UnsupportedFeatures.All(IsReadSafeUnsupportedFeature);

        /// <summary>True when ordinary content edits have no known unsupported package blocker.</summary>
        public bool CanEditDocumentContent => UnsupportedFeatures.Count == 0;

        /// <summary>True when structure-changing edits have no preserve-only or unsupported blocker.</summary>
        public bool CanEditDocumentStructure => !HasAdvancedFeatures;

        /// <summary>True when template binding has no preserve-only or unsupported blocker.</summary>
        public bool CanBindTemplate => !HasAdvancedFeatures;

        /// <summary>True when no discovered feature is known to require external materialization or exact unsupported layout.</summary>
        public bool CanRenderFixedLayout => GetFixedLayoutBlockers().Count == 0;

        /// <summary>True when saving is not known to invalidate an unsupported package contract.</summary>
        public bool CanSavePackageRoundTrip => UnsupportedFeatures.Count == 0;

        /// <summary>Returns whether the requested workflow capability can be attempted.</summary>
        public bool Can(WordPreflightCapability capability) {
            switch (capability) {
                case WordPreflightCapability.ReadDocumentContent:
                    return CanReadDocumentContent;
                case WordPreflightCapability.EditDocumentContent:
                    return CanEditDocumentContent;
                case WordPreflightCapability.EditDocumentStructure:
                    return CanEditDocumentStructure;
                case WordPreflightCapability.BindTemplate:
                    return CanBindTemplate;
                case WordPreflightCapability.RenderFixedLayout:
                    return CanRenderFixedLayout;
                case WordPreflightCapability.SavePackageRoundTrip:
                    return CanSavePackageRoundTrip;
                default:
                    throw new ArgumentOutOfRangeException(nameof(capability), capability,
                        "Unsupported Word preflight capability.");
            }
        }

        /// <summary>Throws with operation-specific diagnostics when a capability is unavailable.</summary>
        public WordFeatureReport EnsureCan(WordPreflightCapability capability) {
            if (Can(capability)) return this;
            IReadOnlyList<string> diagnostics = GetCapabilityDiagnostics(capability);
            string detail = diagnostics.Count == 0
                ? "No additional diagnostics were reported."
                : string.Join("; ", diagnostics);
            throw new InvalidOperationException(
                $"Word preflight capability '{capability}' is not available: {detail}");
        }

        /// <summary>Explains why a workflow capability is unavailable.</summary>
        public IReadOnlyList<string> GetCapabilityDiagnostics(WordPreflightCapability capability) {
            if (Can(capability)) return Array.Empty<string>();
            var messages = new List<string>();
            foreach (WordFeatureFinding finding in GetCapabilityFindings(capability)) {
                AddDistinct(messages, FormatCapabilityFinding(finding));
            }
            if (messages.Count == 0) {
                AddDistinct(messages, "The requested Word workflow is not available for this document.");
            }
            return messages.AsReadOnly();
        }

        internal IEnumerable<WordFeatureFinding> GetCapabilityFindings(WordPreflightCapability capability) {
            switch (capability) {
                case WordPreflightCapability.ReadDocumentContent:
                    return UnsupportedFeatures.Where(finding => !IsReadSafeUnsupportedFeature(finding));
                case WordPreflightCapability.EditDocumentContent:
                case WordPreflightCapability.SavePackageRoundTrip:
                    return UnsupportedFeatures;
                case WordPreflightCapability.EditDocumentStructure:
                case WordPreflightCapability.BindTemplate:
                    return UnsupportedFeatures.Concat(PreservedFeatures);
                case WordPreflightCapability.RenderFixedLayout:
                    return GetFixedLayoutBlockers();
                default:
                    throw new ArgumentOutOfRangeException(nameof(capability), capability,
                        "Unsupported Word preflight capability.");
            }
        }

        private IReadOnlyList<WordFeatureFinding> GetFixedLayoutBlockers() =>
            FindFeatures(FixedLayoutBlockerNames);

        private static bool IsReadSafeUnsupportedFeature(WordFeatureFinding finding) =>
            string.Equals(finding.Name, "Digital signatures", StringComparison.OrdinalIgnoreCase);

        private static string FormatCapabilityFinding(WordFeatureFinding finding) {
            string message = $"{finding.Name} ({finding.Count}, {finding.SupportLevel}): {finding.Note}";
            if (finding.Details.Count == 0) return message;
            const int maxDetails = 3;
            string details = string.Join("; ", finding.Details.Take(maxDetails));
            if (finding.Details.Count > maxDetails) {
                details += $"; +{finding.Details.Count - maxDetails} more";
            }
            return message + " [" + details + "]";
        }

        private static void AddDistinct(List<string> messages, string message) {
            if (string.IsNullOrWhiteSpace(message)) return;
            if (!messages.Contains(message, StringComparer.Ordinal)) messages.Add(message);
        }
    }
}
