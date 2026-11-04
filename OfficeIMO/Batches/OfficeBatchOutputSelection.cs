namespace OfficeIMO;

/// <summary>Shared regular-file output selection for native batches and file-oriented hosts.</summary>
internal static class OfficeBatchOutputSelection {
    internal const string OutputExists = "OutputExists";
    internal const string UpToDate = "UpToDate";

    /// <summary>Observes an existing destination without decoding it or treating it as a verified converted artifact.</summary>
    internal static OfficeBatchItemReport? ObserveSkip(string input, string output, bool skipExisting, bool skipUpToDate) {
        if (!skipExisting && !skipUpToDate) return null;
        OfficeBatchExecutor.EnsureNoLinks(input);
        OfficeBatchExecutor.EnsureNoLinks(output);
        var destination = new FileInfo(output);
        if (!destination.Exists) return null;
        string reason;
        if (skipExisting) reason = OutputExists;
        else {
            var source = new FileInfo(input);
            // A package directory timestamp does not establish the freshness of its members.
            if (!source.Exists || destination.LastWriteTimeUtc <= source.LastWriteTimeUtc) return null;
            reason = UpToDate;
        }
        string message = reason == OutputExists ? "Existing output left unchanged."
            : "Output is newer than its source; existing bytes left unchanged.";
        return new(input, null, null, false, message,
            [new OfficeDiagnostic(OfficeDiagnosticCodes.ConversionInputSkipped, OfficeDiagnosticSeverity.Information,
                message, location: output, format: "selection")], true) { SkippedReason = reason };
    }
}
