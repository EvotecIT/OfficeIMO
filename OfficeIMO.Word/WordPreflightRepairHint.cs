namespace OfficeIMO.Word {
    /// <summary>Actionable guidance for resolving or routing around a failed Word capability preflight.</summary>
    public sealed class WordPreflightRepairHint {
        internal WordPreflightRepairHint(WordPreflightCapability capability, string featureName,
            string action, string? command = null, string? details = null) {
            Capability = capability;
            FeatureName = featureName;
            Action = action;
            Command = command;
            Details = details;
        }

        /// <summary>Capability the hint applies to.</summary>
        public WordPreflightCapability Capability { get; }

        /// <summary>Feature or blocker that triggered the hint.</summary>
        public string FeatureName { get; }

        /// <summary>Recommended action.</summary>
        public string Action { get; }

        /// <summary>Optional OfficeIMO API or command-shaped remediation.</summary>
        public string? Command { get; }

        /// <summary>Optional context for the recommendation.</summary>
        public string? Details { get; }
    }
}
