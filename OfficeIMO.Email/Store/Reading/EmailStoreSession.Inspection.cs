namespace OfficeIMO.Email.Store;

public sealed partial class EmailStoreSession {
    /// <summary>Returns the constant-memory catalog snapshot built while opening this session.</summary>
    public EmailStoreInspectionReport Inspect() {
        ThrowIfDisposed();
        return new EmailStoreInspectionReport(this);
    }

    /// <summary>Validates a bounded store scope at structural, summary, or fully projected depth.</summary>
    public EmailStoreValidationReport Validate(EmailStoreValidationOptions? options = null,
        CancellationToken cancellationToken = default) {
        ThrowIfDisposed();
        EmailStoreValidationOptions effective = options ?? new EmailStoreValidationOptions();
        EmailStoreStructuralValidationResult? structural = null;
        if (effective.VerifyStructuralIntegrity) {
            structural = _backend is PstStoreSessionBackend pst
                ? pst.ValidateStructure(effective, cancellationToken)
                : EmailStoreStructuralValidationResult.NotSupported();
        }
        if (effective.Mode == EmailStoreValidationMode.Shallow) {
            return new EmailStoreValidationReport(
                effective.Mode, 0, 0, 0, false,
                Diagnostics.Concat(structural?.Diagnostics ?? Array.Empty<EmailStoreDiagnostic>()).ToArray(),
                effective.VerifyStructuralIntegrity, structural);
        }

        int enumerationLimit = effective.MaxItems == int.MaxValue
            ? int.MaxValue
            : effective.MaxItems + 1;
        var enumeration = new EmailStoreEnumerationOptions(
            effective.FolderId,
            effective.IncludeDescendants,
            effective.IncludeAssociatedItems,
            effective.IncludeOrphanedItems,
            enumerationLimit);
        int examined = 0;
        int failed = 0;
        int orphaned = 0;
        bool truncated = false;
        var validationDiagnostics = new List<EmailStoreDiagnostic>();
        foreach (EmailStoreItemReference reference in EnumerateItems(enumeration, cancellationToken)) {
            cancellationToken.ThrowIfCancellationRequested();
            if (examined >= effective.MaxItems) {
                truncated = true;
                break;
            }
            examined++;
            if (reference.IsOrphaned) orphaned++;
            try {
                if (effective.Mode == EmailStoreValidationMode.FullItems) {
                    ReadItem(reference, cancellationToken);
                } else {
                    ReadSummary(reference, cancellationToken);
                }
            } catch (EmailStoreLimitExceededException exception) {
                failed++;
                validationDiagnostics.Add(new EmailStoreDiagnostic(
                    "EMAIL_STORE_VALIDATION_ITEM_LIMIT",
                    exception.Message,
                    EmailStoreDiagnosticSeverity.Warning,
                    string.Concat("item/", reference.Id)));
            } catch (Exception exception) when (
                exception is InvalidDataException ||
                exception is NotSupportedException ||
                exception is KeyNotFoundException) {
                failed++;
                validationDiagnostics.Add(new EmailStoreDiagnostic(
                    "EMAIL_STORE_VALIDATION_ITEM_FAILED",
                    exception.Message,
                    EmailStoreDiagnosticSeverity.Error,
                    string.Concat("item/", reference.Id)));
            }
        }

        EmailStoreDiagnostic[] diagnostics = Diagnostics
            .Concat(structural?.Diagnostics ?? Array.Empty<EmailStoreDiagnostic>())
            .Concat(validationDiagnostics)
            .ToArray();
        return new EmailStoreValidationReport(
            effective.Mode, examined, failed, orphaned, truncated, diagnostics,
            effective.VerifyStructuralIntegrity, structural);
    }

    /// <summary>
    /// Discovers indexed items absent from normal folder contents tables without modifying the source.
    /// </summary>
    public EmailStoreRecoveryReport DiscoverRecoverableItems(EmailStoreRecoveryOptions? options = null,
        CancellationToken cancellationToken = default) {
        ThrowIfDisposed();
        EmailStoreRecoveryOptions effective = options ?? new EmailStoreRecoveryOptions();
        int enumerationLimit = effective.MaxItemsScanned == int.MaxValue
            ? int.MaxValue
            : effective.MaxItemsScanned + 1;
        var enumeration = new EmailStoreEnumerationOptions(
            effective.FolderId,
            effective.IncludeDescendants,
            effective.IncludeAssociatedItems,
            includeOrphanedItems: true,
            maxItems: enumerationLimit);
        int scanned = 0;
        bool stoppedAtLimit = false;
        var recovered = new List<EmailStoreItemReference>();
        foreach (EmailStoreItemReference reference in EnumerateItems(enumeration, cancellationToken)) {
            cancellationToken.ThrowIfCancellationRequested();
            if (scanned >= effective.MaxItemsScanned) {
                stoppedAtLimit = true;
                break;
            }
            scanned++;
            if (!reference.IsOrphaned) continue;
            if (recovered.Count >= effective.MaxRecoveredItems) {
                stoppedAtLimit = true;
                break;
            }
            recovered.Add(reference);
        }
        return new EmailStoreRecoveryReport(
            scanned, stoppedAtLimit, recovered, Diagnostics.ToArray());
    }
}
