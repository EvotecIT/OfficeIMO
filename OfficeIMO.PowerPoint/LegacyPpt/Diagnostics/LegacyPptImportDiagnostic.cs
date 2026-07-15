namespace OfficeIMO.PowerPoint.LegacyPpt.Diagnostics {
    /// <summary>Describes a finding produced while reading a PowerPoint 97-2003 binary presentation.</summary>
    public sealed class LegacyPptImportDiagnostic {
        /// <summary>Creates an import diagnostic.</summary>
        public LegacyPptImportDiagnostic(string code, string message, LegacyPptDiagnosticSeverity severity,
            long? streamOffset = null) {
            Code = string.IsNullOrWhiteSpace(code) ? "PPT-UNKNOWN" : code;
            Message = message ?? string.Empty;
            Severity = severity;
            StreamOffset = streamOffset;
        }

        /// <summary>Gets the stable diagnostic code.</summary>
        public string Code { get; }

        /// <summary>Gets the human-readable diagnostic message.</summary>
        public string Message { get; }

        /// <summary>Gets the diagnostic severity.</summary>
        public LegacyPptDiagnosticSeverity Severity { get; }

        /// <summary>Gets the offset in the PowerPoint Document stream, when known.</summary>
        public long? StreamOffset { get; }

        /// <inheritdoc />
        public override string ToString() => StreamOffset.HasValue
            ? $"{Code} at 0x{StreamOffset.Value:X}: {Message}"
            : $"{Code}: {Message}";
    }
}
