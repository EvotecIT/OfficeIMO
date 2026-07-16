using OfficeIMO.Email;

namespace OfficeIMO.Email.AddressBook;

public sealed partial class OfflineAddressBookSession {
    /// <summary>Runs an explicit checksum, record-framing, or full-decode integrity pass.</summary>
    public OfflineAddressBookValidationReport Validate(
        OfflineAddressBookValidationOptions? options = null,
        IProgress<OfflineAddressBookValidationProgress>? progress = null,
        CancellationToken cancellationToken = default) {
        ThrowIfDisposed();
        OfflineAddressBookValidationOptions effective = options ?? new OfflineAddressBookValidationOptions();
        OabAddressListSource[] selectedSources = SelectSources(effective.AddressListId).ToArray();
        var results = new List<OfflineAddressBookValidationResult>(selectedSources.Length);
        long totalBytesHashed = 0;
        long totalEntriesScanned = 0;
        long totalEntriesSkipped = 0;

        for (int sourceIndex = 0; sourceIndex < selectedSources.Length; sourceIndex++) {
            cancellationToken.ThrowIfCancellationRequested();
            OabAddressListSource selected = selectedSources[sourceIndex];
            var diagnostics = new List<EmailDiagnostic>();
            uint? checksum = null;
            if (effective.ValidateChecksum) {
                long before = totalBytesHashed;
                checksum = OabCrc32.Compute(
                    selected.Source,
                    effective.MaxChecksumBytesPerFile,
                    effective.ProgressByteInterval,
                    bytes => {
                        totalBytesHashed = before + bytes;
                        ReportValidationProgress(progress, sourceIndex,
                            totalBytesHashed, totalEntriesScanned, totalEntriesSkipped);
                    },
                    cancellationToken);
                if (checksum.Value != selected.Info.Serial) {
                    diagnostics.Add(new EmailDiagnostic(
                        "OAB_CHECKSUM_MISMATCH",
                        "The recalculated OAB payload checksum does not match the header value.",
                        EmailDiagnosticSeverity.Error,
                        selected.Info.Id));
                }
            }

            long recordsScanned = 0;
            long recordsDecoded = 0;
            long recordsSkipped = 0;
            bool framingComplete = effective.Mode == OfflineAddressBookValidationMode.ChecksumOnly;
            bool consumedDeclaredPayload = effective.Mode == OfflineAddressBookValidationMode.ChecksumOnly;
            if (effective.Mode != OfflineAddressBookValidationMode.ChecksumOnly) {
                ValidateRecords(selected, effective, diagnostics,
                    ref recordsScanned, ref recordsDecoded, ref recordsSkipped,
                    out framingComplete, out consumedDeclaredPayload,
                    sourceIndex, progress, totalBytesHashed,
                    ref totalEntriesScanned, ref totalEntriesSkipped,
                    cancellationToken);
            }

            results.Add(new OfflineAddressBookValidationResult(
                selected.Info, checksum, recordsScanned, recordsDecoded, recordsSkipped,
                framingComplete, consumedDeclaredPayload, diagnostics));
            ReportValidationProgress(progress, sourceIndex + 1,
                totalBytesHashed, totalEntriesScanned, totalEntriesSkipped);
        }
        return new OfflineAddressBookValidationReport(results);
    }

    private void ValidateRecords(OabAddressListSource selected,
        OfflineAddressBookValidationOptions options,
        List<EmailDiagnostic> diagnostics,
        ref long recordsScanned,
        ref long recordsDecoded,
        ref long recordsSkipped,
        out bool framingComplete,
        out bool consumedDeclaredPayload,
        int completedSources,
        IProgress<OfflineAddressBookValidationProgress>? progress,
        long totalBytesHashed,
        ref long totalEntriesScanned,
        ref long totalEntriesSkipped,
        CancellationToken cancellationToken) {
        framingComplete = false;
        consumedDeclaredPayload = false;
        using (OabStreamLease lease = selected.Source.OpenRead()) {
            Stream stream = lease.Stream;
            OabBinary.Seek(selected.Source, stream, selected.Info.EntriesOffset, selected.Info.Id);
            long maximum = Math.Min(selected.Info.DeclaredEntryCount, options.MaxEntriesPerAddressList);
            for (long entryIndex = 0; entryIndex < maximum; entryIndex++) {
                cancellationToken.ThrowIfCancellationRequested();
                long offset = stream.Position - selected.Source.BaseOffset;
                string location = BuildEntryLocation(selected.Info, entryIndex);
                OabRecordEnvelope envelope;
                try {
                    envelope = OabV4RecordReader.ReadEnvelope(selected.Source, stream, _options, location);
                } catch (Exception exception) when (IsRecoverableReadException(exception)) {
                    diagnostics.Add(new EmailDiagnostic(
                        "OAB_VALIDATION_FRAMING_STOPPED",
                        exception.Message,
                        EmailDiagnosticSeverity.Error,
                        location));
                    return;
                }
                recordsScanned++;
                totalEntriesScanned++;
                if (options.Mode == OfflineAddressBookValidationMode.FullDecode) {
                    try {
                        OabParsedRecord parsed = OabV4RecordReader.Parse(envelope,
                            selected.Info.EntryPropertyDefinitions, _options, location);
                        foreach (EmailDiagnostic diagnostic in parsed.Diagnostics) diagnostics.Add(diagnostic);
                        recordsDecoded++;
                    } catch (Exception exception) when (
                        options.ContinueOnEntryError && IsRecoverableReadException(exception)) {
                        recordsSkipped++;
                        totalEntriesSkipped++;
                        diagnostics.Add(new EmailDiagnostic(
                            "OAB_VALIDATION_ENTRY_SKIPPED",
                            exception.Message,
                            exception is OfflineAddressBookLimitExceededException
                                ? EmailDiagnosticSeverity.Warning
                                : EmailDiagnosticSeverity.Error,
                            string.Concat(selected.Info.Id, ":", offset.ToString(CultureInfo.InvariantCulture))));
                    }
                }
                if (recordsScanned % options.ProgressEntryInterval == 0) {
                    ReportValidationProgress(progress, completedSources,
                        totalBytesHashed, totalEntriesScanned, totalEntriesSkipped);
                }
            }

            framingComplete = recordsScanned == selected.Info.DeclaredEntryCount;
            if (!framingComplete) {
                diagnostics.Add(new EmailDiagnostic(
                    "OAB_VALIDATION_ENTRY_LIMIT",
                    "Validation stopped at the configured per-address-list entry limit.",
                    EmailDiagnosticSeverity.Warning,
                    selected.Info.Id));
                return;
            }
            long relativePosition = stream.Position - selected.Source.BaseOffset;
            consumedDeclaredPayload = relativePosition == selected.Source.Length;
            if (!consumedDeclaredPayload) {
                diagnostics.Add(new EmailDiagnostic(
                    "OAB_VALIDATION_TRAILING_DATA",
                    string.Concat((selected.Source.Length - relativePosition).ToString(CultureInfo.InvariantCulture),
                        " byte(s) remain after the declared record sequence."),
                    EmailDiagnosticSeverity.Error,
                    selected.Info.Id));
            }
        }
    }

    private static void ReportValidationProgress(
        IProgress<OfflineAddressBookValidationProgress>? progress,
        int addressListsCompleted,
        long bytesHashed,
        long entriesScanned,
        long entriesSkipped) =>
        progress?.Report(new OfflineAddressBookValidationProgress(
            addressListsCompleted, bytesHashed, entriesScanned, entriesSkipped));
}
