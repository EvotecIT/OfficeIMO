namespace OfficeIMO.Excel {
    /// <summary>
    /// Actionable guidance for resolving or routing around a failed Excel preflight capability.
    /// </summary>
    public sealed class ExcelPreflightRepairHint {
        internal ExcelPreflightRepairHint(
            ExcelPreflightCapability capability,
            string featureName,
            string action,
            string? command = null,
            string? details = null) {
            Capability = capability;
            FeatureName = featureName;
            Action = action;
            Command = command;
            Details = details;
        }

        /// <summary>Capability the hint applies to.</summary>
        public ExcelPreflightCapability Capability { get; }

        /// <summary>Feature or blocker that triggered the hint.</summary>
        public string FeatureName { get; }

        /// <summary>Recommended action.</summary>
        public string Action { get; }

        /// <summary>Optional OfficeIMO API or command-shaped remediation.</summary>
        public string? Command { get; }

        /// <summary>Optional additional context.</summary>
        public string? Details { get; }
    }
}
