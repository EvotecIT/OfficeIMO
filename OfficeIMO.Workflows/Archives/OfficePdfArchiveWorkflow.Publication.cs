namespace OfficeIMO.Workflows;

public static partial class OfficePdfArchiveWorkflow {
    private static string GetPendingStagePath(string output, string stageId) {
        if (!Guid.TryParseExact(stageId, "N", out _)) throw new InvalidDataException("Archive publication identity is invalid.");
        return Path.Combine(Path.GetDirectoryName(output)!, "." + stageId + ".archive.pdf");
    }

    private static async Task CompletePublicationAsync(OfficePdfArchiveRequest settings, string input, string output,
        string receiptPath, OfficePdfArchiveReceipt receipt, CancellationToken token, IOfficeWorkflowPublicationGuard? hostGuard) {
        if (receipt.OutputSha256 == null) throw new InvalidDataException("Completed archive output has no recorded hash.");
        EnsureNoLinks(output);
        if (File.Exists(output)) {
            if (await HashFileAsync(output, settings.OutputDirectory, settings.MaximumOutputBytes, token).ConfigureAwait(false) != receipt.OutputSha256)
                throw new InvalidDataException("Completed archive output changed; no automatic replacement is allowed.");
        } else {
            if (receipt.PendingStageId == null)
                throw new InvalidDataException("Completed archive output is missing; no automatic retry is allowed.");
            string staged = GetPendingStagePath(output, receipt.PendingStageId);
            if (hostGuard != null && !await hostGuard.CanPublishAsync(output, false, token).ConfigureAwait(false))
                throw new UnauthorizedAccessException("Archive output is protected by the host publication policy.");
            if (await HashFileAsync(input, settings.InputDirectory, settings.MaximumInputBytes, token).ConfigureAwait(false) != receipt.InputSha256)
                throw new InvalidDataException("Archive source changed before publication; no output was replaced.");
            EnsureNoLinks(staged);
            if (!File.Exists(staged) ||
                await HashFileAsync(staged, settings.OutputDirectory, settings.MaximumOutputBytes, token).ConfigureAwait(false) != receipt.OutputSha256)
                throw new InvalidDataException("Pending archive PDF is missing or changed; no output was replaced.");
            EnsureNoLinks(staged);
            EnsureNoLinks(output);
            token.ThrowIfCancellationRequested();
            File.Move(staged, output, overwrite: false);
            // Confirm the published path before clearing the intent, including replacement
            // races after the last stage check. A mismatch retains the recovery record.
            EnsureNoLinks(output);
            if (await HashFileAsync(output, settings.OutputDirectory, settings.MaximumOutputBytes, CancellationToken.None).ConfigureAwait(false) != receipt.OutputSha256)
                throw new InvalidDataException("Published archive PDF differs from its recorded artifact; completion was not recorded.");
        }
        // A crash after the move is reconciled by verifying the final artifact on restart.
        if (receipt.PendingStageId != null)
            StoreReceipt(receiptPath, receipt with { PendingStageId = null });
    }
}
