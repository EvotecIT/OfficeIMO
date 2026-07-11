namespace OfficeIMO.Pdf;

/// <summary>Chooses a proven full-rewrite, append-only, or blocked path for existing-document mutations.</summary>
public static class PdfMutationPlanner {
    /// <summary>Plans a mutation for a PDF byte array.</summary>
    public static PdfMutationPlan Plan(
        byte[] pdf,
        PdfMutationOperation operation,
        PdfReadOptions? options = null,
        IEnumerable<string>? fieldNames = null,
        PdfMutationExecutionPreference executionPreference = PdfMutationExecutionPreference.Automatic) {
        Guard.NotNull(pdf, nameof(pdf));
        return Plan(PdfInspector.Preflight(pdf, options), operation, fieldNames, executionPreference);
    }

    /// <summary>Plans a mutation for a readable PDF stream.</summary>
    public static PdfMutationPlan Plan(
        Stream input,
        PdfMutationOperation operation,
        PdfReadOptions? options = null,
        IEnumerable<string>? fieldNames = null,
        PdfMutationExecutionPreference executionPreference = PdfMutationExecutionPreference.Automatic) {
        Guard.NotNull(input, nameof(input));
        if (!input.CanRead) {
            throw new ArgumentException("Stream must be readable.", nameof(input));
        }

        using var buffer = new MemoryStream();
        input.CopyTo(buffer);
        return Plan(buffer.ToArray(), operation, options, fieldNames, executionPreference);
    }

    /// <summary>Plans a mutation for a PDF file.</summary>
    public static PdfMutationPlan Plan(
        string inputPath,
        PdfMutationOperation operation,
        PdfReadOptions? options = null,
        IEnumerable<string>? fieldNames = null,
        PdfMutationExecutionPreference executionPreference = PdfMutationExecutionPreference.Automatic) {
        Guard.NotNullOrWhiteSpace(inputPath, nameof(inputPath));
        return Plan(File.ReadAllBytes(inputPath), operation, options, fieldNames, executionPreference);
    }

    /// <summary>Plans a mutation from an existing general preflight report.</summary>
    public static PdfMutationPlan Plan(
        PdfDocumentPreflight preflight,
        PdfMutationOperation operation,
        IEnumerable<string>? fieldNames = null,
        PdfMutationExecutionPreference executionPreference = PdfMutationExecutionPreference.Automatic) {
        Guard.NotNull(preflight, nameof(preflight));
        ValidateOperation(operation);
        ValidateExecutionPreference(executionPreference);

        string[] requestedFields = NormalizeFieldNames(fieldNames);
        PdfDocumentSecurityInfo security = preflight.Probe.Security;
        PdfAppendOnlyMutationReport appendOnly = PdfIncrementalUpdater.AnalyzeAppendOnlyMutation(security, requestedFields);
        bool fullRewriteImplemented = IsFullRewriteImplemented(operation);
        bool appendOnlyImplemented = IsAppendOnlyImplemented(operation);
        bool fullRewriteCapability = CanFullRewrite(preflight, operation);
        bool securityRewrite = operation == PdfMutationOperation.ChangeEncryption && fullRewriteCapability;
        bool fullRewriteAvailable =
            fullRewriteImplemented &&
            fullRewriteCapability &&
            (securityRewrite ||
                (!security.RequiresAppendOnlyMutation &&
                (!security.BlocksOfficeIMOFullRewriteMutation || CanExtractEncryptedPages(preflight, operation))));
        bool appendOnlyAvailable = appendOnlyImplemented && CanAppend(appendOnly, operation);

        PdfMutationExecutionMode mode;
        if (executionPreference == PdfMutationExecutionPreference.RequireFullRewrite) {
            mode = fullRewriteAvailable ? PdfMutationExecutionMode.FullRewrite : PdfMutationExecutionMode.Blocked;
        } else if (executionPreference == PdfMutationExecutionPreference.RequireAppendOnly || RequiresAppendOnlyByDefinition(operation)) {
            mode = appendOnlyAvailable ? PdfMutationExecutionMode.AppendOnly : PdfMutationExecutionMode.Blocked;
        } else if (fullRewriteAvailable) {
            mode = PdfMutationExecutionMode.FullRewrite;
        } else if (appendOnlyAvailable) {
            mode = PdfMutationExecutionMode.AppendOnly;
        } else {
            mode = PdfMutationExecutionMode.Blocked;
        }

        IReadOnlyList<PdfMutationStructure> structures = GetAffectedStructures(operation);
        IReadOnlyList<PdfMutationPermissionCheck> permissions = GetPermissionChecks(operation, mode);
        IReadOnlyList<PdfMutationProof> proofs = GetRequiredProofs(operation, mode, security);
        IReadOnlyList<string> blockers = mode == PdfMutationExecutionMode.Blocked
            ? GetBlockerCodes(preflight, appendOnly, operation, fullRewriteImplemented, appendOnlyImplemented, security)
            : Array.Empty<string>();
        IReadOnlyList<PdfMutationCapabilityRecord> capabilityRecords = BuildCapabilityRecords(
            operation,
            structures,
            fullRewriteImplemented,
            appendOnlyImplemented,
            fullRewriteAvailable,
            appendOnlyAvailable,
            permissions,
            proofs,
            blockers);
        IReadOnlyList<string> warnings = GetWarnings(preflight, appendOnly, mode, security);
        IReadOnlyList<string> diagnostics = GetDiagnostics(preflight, appendOnly, operation, mode, blockers, fullRewriteCapability, appendOnlyAvailable, security);

        return new PdfMutationPlan(
            operation,
            executionPreference,
            mode,
            preflight,
            appendOnly,
            fullRewriteAvailable,
            appendOnlyAvailable,
            structures,
            capabilityRecords,
            permissions,
            proofs,
            blockers,
            warnings,
            diagnostics);
    }

    private static bool CanFullRewrite(PdfDocumentPreflight preflight, PdfMutationOperation operation) {
        switch (operation) {
            case PdfMutationOperation.FillFormFields:
                return preflight.CanFillSimpleFormFields;
            case PdfMutationOperation.FlattenFormFields:
                return preflight.CanFlattenSimpleFormFields;
            case PdfMutationOperation.FillAndFlattenFormFields:
                return preflight.CanFillAndFlattenSimpleFormFields;
            case PdfMutationOperation.PrepareExternalSignature:
            case PdfMutationOperation.ModifyAttachments:
                return false;
            case PdfMutationOperation.ChangeEncryption:
                return CanChangeEncryption(preflight);
            case PdfMutationOperation.ExtractPages:
                return preflight.CanRewrite || CanExtractEncryptedPages(preflight, operation);
            default:
                return preflight.CanRewrite;
        }
    }

    private static bool CanAppend(PdfAppendOnlyMutationReport report, PdfMutationOperation operation) {
        switch (operation) {
            case PdfMutationOperation.UpdateMetadata:
                return report.CanAppendMetadata;
            case PdfMutationOperation.FillFormFields:
                return report.CanAppendFormFields;
            case PdfMutationOperation.PrepareExternalSignature:
                return report.CanPrepareExternalSignature;
            default:
                return false;
        }
    }

    private static bool IsFullRewriteImplemented(PdfMutationOperation operation) {
        return operation != PdfMutationOperation.PrepareExternalSignature &&
            operation != PdfMutationOperation.ModifyAttachments;
    }

    private static bool IsAppendOnlyImplemented(PdfMutationOperation operation) {
        return operation == PdfMutationOperation.UpdateMetadata ||
            operation == PdfMutationOperation.FillFormFields ||
            operation == PdfMutationOperation.PrepareExternalSignature;
    }

    private static bool RequiresAppendOnlyByDefinition(PdfMutationOperation operation) =>
        operation == PdfMutationOperation.PrepareExternalSignature;

    private static System.Collections.ObjectModel.ReadOnlyCollection<PdfMutationStructure> GetAffectedStructures(PdfMutationOperation operation) {
        switch (operation) {
            case PdfMutationOperation.UpdateMetadata:
                return ReadOnly(PdfMutationStructure.InfoDictionary);
            case PdfMutationOperation.FillFormFields:
                return ReadOnly(PdfMutationStructure.AcroForm, PdfMutationStructure.AppearanceStreams, PdfMutationStructure.Annotations);
            case PdfMutationOperation.FlattenFormFields:
            case PdfMutationOperation.FillAndFlattenFormFields:
                return ReadOnly(PdfMutationStructure.AcroForm, PdfMutationStructure.AppearanceStreams, PdfMutationStructure.Annotations, PdfMutationStructure.PageContent, PdfMutationStructure.PageResources);
            case PdfMutationOperation.PrepareExternalSignature:
                return ReadOnly(PdfMutationStructure.Signatures, PdfMutationStructure.AcroForm, PdfMutationStructure.Catalog, PdfMutationStructure.ObjectGraph);
            case PdfMutationOperation.ModifyPageTree:
            case PdfMutationOperation.ExtractPages:
                return ReadOnly(PdfMutationStructure.PageTree, PdfMutationStructure.PageResources, PdfMutationStructure.Navigation, PdfMutationStructure.Catalog);
            case PdfMutationOperation.ModifyPageContent:
                return ReadOnly(PdfMutationStructure.PageContent, PdfMutationStructure.PageResources);
            case PdfMutationOperation.ModifyCatalog:
                return ReadOnly(PdfMutationStructure.Catalog, PdfMutationStructure.Navigation);
            case PdfMutationOperation.ModifyAnnotations:
                return ReadOnly(PdfMutationStructure.Annotations, PdfMutationStructure.AppearanceStreams, PdfMutationStructure.PageResources);
            case PdfMutationOperation.ModifyAttachments:
                return ReadOnly(PdfMutationStructure.Attachments, PdfMutationStructure.Catalog, PdfMutationStructure.ObjectGraph);
            case PdfMutationOperation.ChangeEncryption:
                return ReadOnly(PdfMutationStructure.Encryption, PdfMutationStructure.ObjectGraph);
            case PdfMutationOperation.Optimize:
                return ReadOnly(PdfMutationStructure.ObjectGraph, PdfMutationStructure.Catalog, PdfMutationStructure.PageResources);
            case PdfMutationOperation.Redact:
                return ReadOnly(PdfMutationStructure.PageContent, PdfMutationStructure.PageResources, PdfMutationStructure.Annotations, PdfMutationStructure.TaggedContent, PdfMutationStructure.InfoDictionary, PdfMutationStructure.XmpMetadata, PdfMutationStructure.Attachments);
            default:
                throw new ArgumentOutOfRangeException(nameof(operation), operation, "Unsupported PDF mutation operation.");
        }
    }

    private static System.Collections.ObjectModel.ReadOnlyCollection<PdfMutationPermissionCheck> GetPermissionChecks(PdfMutationOperation operation, PdfMutationExecutionMode mode) {
        var permissions = new List<PdfMutationPermissionCheck> { PdfMutationPermissionCheck.ReadDocument };
        switch (operation) {
            case PdfMutationOperation.FillFormFields:
            case PdfMutationOperation.FlattenFormFields:
            case PdfMutationOperation.FillAndFlattenFormFields:
                Add(permissions, PdfMutationPermissionCheck.FillForms);
                Add(permissions, PdfMutationPermissionCheck.DocMdp);
                Add(permissions, PdfMutationPermissionCheck.FieldMdp);
                break;
            case PdfMutationOperation.ModifyPageTree:
            case PdfMutationOperation.ExtractPages:
                Add(permissions, PdfMutationPermissionCheck.ModifyDocument);
                Add(permissions, PdfMutationPermissionCheck.AssembleDocument);
                Add(permissions, PdfMutationPermissionCheck.DocMdp);
                break;
            case PdfMutationOperation.ModifyAnnotations:
                Add(permissions, PdfMutationPermissionCheck.ModifyAnnotations);
                Add(permissions, PdfMutationPermissionCheck.DocMdp);
                break;
            case PdfMutationOperation.ChangeEncryption:
                Add(permissions, PdfMutationPermissionCheck.OwnerAuthorization);
                break;
            default:
                Add(permissions, PdfMutationPermissionCheck.ModifyDocument);
                if (operation != PdfMutationOperation.Optimize) {
                    Add(permissions, PdfMutationPermissionCheck.DocMdp);
                }
                break;
        }

        if (mode == PdfMutationExecutionMode.AppendOnly || operation == PdfMutationOperation.PrepareExternalSignature) {
            Add(permissions, PdfMutationPermissionCheck.AppendRevision);
        }

        return permissions.AsReadOnly();
    }

    private static IReadOnlyList<PdfMutationProof> GetRequiredProofs(
        PdfMutationOperation operation,
        PdfMutationExecutionMode mode,
        PdfDocumentSecurityInfo security) {
        if (mode == PdfMutationExecutionMode.Blocked) {
            return Array.Empty<PdfMutationProof>();
        }

        var proofs = new List<PdfMutationProof> { PdfMutationProof.ReadableOutput };
        if (mode == PdfMutationExecutionMode.FullRewrite) {
            Add(proofs, PdfMutationProof.RewritePreservation);
        } else {
            Add(proofs, PdfMutationProof.BytePrefixPreservation);
            Add(proofs, PdfMutationProof.RevisionChain);
            if (security.HasSignatures || security.HasDocMDPPermissions) {
                Add(proofs, PdfMutationProof.SignatureByteRanges);
                Add(proofs, PdfMutationProof.SignaturePermissions);
            }
        }

        switch (operation) {
            case PdfMutationOperation.UpdateMetadata:
                Add(proofs, PdfMutationProof.MetadataReadback);
                break;
            case PdfMutationOperation.FillFormFields:
                Add(proofs, PdfMutationProof.FormFieldReadback);
                Add(proofs, PdfMutationProof.VisualRendering);
                break;
            case PdfMutationOperation.FlattenFormFields:
            case PdfMutationOperation.FillAndFlattenFormFields:
                Add(proofs, PdfMutationProof.FormFieldReadback);
                Add(proofs, PdfMutationProof.VisualRendering);
                break;
            case PdfMutationOperation.PrepareExternalSignature:
                Add(proofs, PdfMutationProof.SignatureByteRanges);
                Add(proofs, PdfMutationProof.SignaturePermissions);
                break;
            case PdfMutationOperation.ModifyPageTree:
            case PdfMutationOperation.ExtractPages:
                Add(proofs, PdfMutationProof.PageStructureReadback);
                break;
            case PdfMutationOperation.ModifyPageContent:
                Add(proofs, PdfMutationProof.VisualRendering);
                break;
            case PdfMutationOperation.ModifyAnnotations:
                Add(proofs, PdfMutationProof.AnnotationReadback);
                Add(proofs, PdfMutationProof.VisualRendering);
                break;
            case PdfMutationOperation.ModifyAttachments:
                Add(proofs, PdfMutationProof.AttachmentReadback);
                break;
            case PdfMutationOperation.ChangeEncryption:
                Add(proofs, PdfMutationProof.EncryptionRoundTrip);
                break;
            case PdfMutationOperation.Redact:
                Add(proofs, PdfMutationProof.RedactionResidue);
                Add(proofs, PdfMutationProof.VisualRendering);
                break;
        }

        return proofs.AsReadOnly();
    }

    private static System.Collections.ObjectModel.ReadOnlyCollection<string> GetBlockerCodes(
        PdfDocumentPreflight preflight,
        PdfAppendOnlyMutationReport appendOnly,
        PdfMutationOperation operation,
        bool fullRewriteImplemented,
        bool appendOnlyImplemented,
        PdfDocumentSecurityInfo security) {
        var blockers = new List<string>();
        if (!preflight.CanRead) {
            for (int i = 0; i < preflight.ReadBlockers.Count; i++) {
                Add(blockers, "Read." + preflight.ReadBlockers[i].Kind);
            }
        }

        if (!fullRewriteImplemented) {
            Add(blockers, "FullRewrite.NotImplemented." + operation);
        } else if (operation == PdfMutationOperation.ChangeEncryption) {
            if (security.HasEncryption && !security.HasOwnerAuthorization) {
                Add(blockers, "FullRewrite.Encryption.OwnerAuthorizationRequired");
            }

            if (security.HasSignatures) {
                Add(blockers, "FullRewrite.SignaturePreservation");
            }

            if (security.HasDocMDPPermissions) {
                Add(blockers, "FullRewrite.DocMdpPreservation");
            }

            if (security.HasUsageRights) {
                Add(blockers, "FullRewrite.UsageRightsPreservation");
            }
        } else {
            for (int i = 0; i < preflight.RewriteBlockers.Count; i++) {
                Add(blockers, "FullRewrite." + preflight.RewriteBlockers[i].Kind);
            }

            if (security.HasXrefStreams) {
                Add(blockers, "FullRewrite.XrefStreamPreservation");
            }

            if (security.HasObjectStreams) {
                Add(blockers, "FullRewrite.ObjectStreamPreservation");
            }

            if (security.RequiresAppendOnlyMutation) {
                Add(blockers, "FullRewrite.AppendOnlyRequired");
            }
        }

        if (!appendOnlyImplemented) {
            Add(blockers, "AppendOnly.NotImplemented." + operation);
        } else {
            for (int i = 0; i < appendOnly.Blockers.Count; i++) {
                Add(blockers, "AppendOnly." + appendOnly.Blockers[i]);
            }

            string appendAction = GetAppendAction(operation);
            if (!appendOnly.SupportedActions.Contains(appendAction, StringComparer.Ordinal)) {
                Add(blockers, "AppendOnly.ActionBlocked." + appendAction);
            }
        }

        if (blockers.Count == 0) {
            Add(blockers, "Mutation.NoProvenExecutionPath");
        }

        return blockers.AsReadOnly();
    }

    private static IReadOnlyList<string> GetWarnings(
        PdfDocumentPreflight preflight,
        PdfAppendOnlyMutationReport appendOnly,
        PdfMutationExecutionMode mode,
        PdfDocumentSecurityInfo security) {
        var warnings = new List<string>();
        if (mode == PdfMutationExecutionMode.AppendOnly) {
            for (int i = 0; i < appendOnly.Warnings.Count; i++) {
                Add(warnings, "AppendOnly." + appendOnly.Warnings[i]);
            }
        }

        if (preflight.Probe.HasActiveContent || preflight.DocumentInfo?.HasActiveContent == true) {
            Add(warnings, "Input.ActiveContentPreserved");
        }

        if (security.HasDocumentSecurityStore) {
            Add(warnings, "Input.DocumentSecurityStorePresent");
        }

        if (mode == PdfMutationExecutionMode.FullRewrite && security.HasIncrementalUpdates) {
            Add(warnings, "Input.RevisionHistoryWillBeNormalized");
        }

        return warnings.Count == 0 ? Array.Empty<string>() : warnings.AsReadOnly();
    }

    private static System.Collections.ObjectModel.ReadOnlyCollection<string> GetDiagnostics(
        PdfDocumentPreflight preflight,
        PdfAppendOnlyMutationReport appendOnly,
        PdfMutationOperation operation,
        PdfMutationExecutionMode mode,
        IReadOnlyList<string> blockers,
        bool fullRewriteCapability,
        bool appendOnlyAvailable,
        PdfDocumentSecurityInfo security) {
        var diagnostics = new List<string>();
        if (mode == PdfMutationExecutionMode.FullRewrite) {
            Add(diagnostics, operation + " can use a full rewrite for this PDF; rewrite-preservation proof is required.");
        } else if (mode == PdfMutationExecutionMode.AppendOnly) {
            string reason = security.RequiresAppendOnlyMutation
                ? "the input requires prior bytes and revisions to be preserved"
                : fullRewriteCapability
                    ? "append-only mutation is the safer available path"
                    : "full rewrite is blocked for this document structure";
            Add(diagnostics, operation + " will use an append-only revision because " + reason + ".");
        } else {
            Add(diagnostics, operation + " is blocked because neither a proven full rewrite nor append-only path is available.");
            for (int i = 0; i < blockers.Count; i++) {
                Add(diagnostics, "Mutation blocker: " + blockers[i] + ".");
            }

            if (!fullRewriteCapability) {
                for (int i = 0; i < preflight.Diagnostics.Count; i++) {
                    Add(diagnostics, preflight.Diagnostics[i]);
                }
            }

            if (!appendOnlyAvailable) {
                for (int i = 0; i < appendOnly.Blockers.Count; i++) {
                    Add(diagnostics, "Append-only blocker: " + appendOnly.Blockers[i] + ".");
                }
            }
        }

        return diagnostics.AsReadOnly();
    }

    private static string GetAppendAction(PdfMutationOperation operation) {
        switch (operation) {
            case PdfMutationOperation.UpdateMetadata:
                return "Metadata";
            case PdfMutationOperation.FillFormFields:
                return "FormFill";
            case PdfMutationOperation.PrepareExternalSignature:
                return "SignaturePrepare";
            case PdfMutationOperation.ModifyAnnotations:
                return "Annotations";
            case PdfMutationOperation.ModifyPageTree:
                return "PageTree";
            case PdfMutationOperation.ModifyAttachments:
                return "Attachments";
            default:
                return operation.ToString();
        }
    }

    private static string[] NormalizeFieldNames(IEnumerable<string>? fieldNames) {
        if (fieldNames is null) {
            return Array.Empty<string>();
        }

        return fieldNames
            .Where(static fieldName => !string.IsNullOrWhiteSpace(fieldName))
            .Distinct(StringComparer.Ordinal)
            .ToArray();
    }

    private static bool CanExtractEncryptedPages(PdfDocumentPreflight preflight, PdfMutationOperation operation) {
        if (operation != PdfMutationOperation.ExtractPages ||
            !preflight.CanRead ||
            !preflight.Probe.HasEncryption) {
            return false;
        }

        for (int i = 0; i < preflight.RewriteBlockers.Count; i++) {
            if (preflight.RewriteBlockers[i].Kind != PdfRewriteBlockerKind.Encryption) {
                return false;
            }
        }

        return true;
    }

    private static bool CanChangeEncryption(PdfDocumentPreflight preflight) {
        if (!preflight.CanRead) {
            return false;
        }

        PdfDocumentSecurityInfo security = preflight.Probe.Security;
        if (security.HasSignatures || security.HasDocMDPPermissions || security.HasUsageRights) {
            return false;
        }

        return !security.HasEncryption || security.HasOwnerAuthorization;
    }

    private static System.Collections.ObjectModel.ReadOnlyCollection<PdfMutationCapabilityRecord> BuildCapabilityRecords(
        PdfMutationOperation operation,
        IReadOnlyList<PdfMutationStructure> structures,
        bool fullRewriteImplemented,
        bool appendOnlyImplemented,
        bool fullRewriteAvailable,
        bool appendOnlyAvailable,
        IReadOnlyList<PdfMutationPermissionCheck> permissions,
        IReadOnlyList<PdfMutationProof> proofs,
        IReadOnlyList<string> blockers) {
        var grouped = new Dictionary<PdfMutationCapabilityKind, List<PdfMutationStructure>>();
        for (int i = 0; i < structures.Count; i++) {
            PdfMutationCapabilityKind kind = GetCapabilityKind(operation, structures[i]);
            if (!grouped.TryGetValue(kind, out List<PdfMutationStructure>? values)) {
                values = new List<PdfMutationStructure>();
                grouped.Add(kind, values);
            }

            Add(values, structures[i]);
        }

        var records = new List<PdfMutationCapabilityRecord>();
        foreach (KeyValuePair<PdfMutationCapabilityKind, List<PdfMutationStructure>> group in grouped.OrderBy(static item => item.Key)) {
            records.Add(new PdfMutationCapabilityRecord(
                group.Key,
                group.Value.AsReadOnly(),
                fullRewriteImplemented,
                appendOnlyImplemented,
                fullRewriteAvailable,
                appendOnlyAvailable,
                permissions,
                proofs,
                blockers));
        }

        return records.AsReadOnly();
    }

    private static PdfMutationCapabilityKind GetCapabilityKind(PdfMutationOperation operation, PdfMutationStructure structure) {
        switch (structure) {
            case PdfMutationStructure.PageTree:
                return PdfMutationCapabilityKind.PageTreeChanges;
            case PdfMutationStructure.PageContent:
            case PdfMutationStructure.PageResources:
            case PdfMutationStructure.TaggedContent:
                return PdfMutationCapabilityKind.ContentChanges;
            case PdfMutationStructure.Catalog:
            case PdfMutationStructure.Navigation:
                return PdfMutationCapabilityKind.CatalogChanges;
            case PdfMutationStructure.AcroForm:
            case PdfMutationStructure.AppearanceStreams when operation != PdfMutationOperation.ModifyAnnotations:
                return PdfMutationCapabilityKind.FormChanges;
            case PdfMutationStructure.AppearanceStreams:
            case PdfMutationStructure.Annotations:
                return PdfMutationCapabilityKind.AnnotationChanges;
            case PdfMutationStructure.InfoDictionary:
            case PdfMutationStructure.XmpMetadata:
                return PdfMutationCapabilityKind.MetadataChanges;
            case PdfMutationStructure.Attachments:
                return PdfMutationCapabilityKind.AttachmentChanges;
            case PdfMutationStructure.Encryption:
                return PdfMutationCapabilityKind.EncryptionChanges;
            case PdfMutationStructure.Signatures:
                return PdfMutationCapabilityKind.SignatureChanges;
            case PdfMutationStructure.ObjectGraph when operation == PdfMutationOperation.PrepareExternalSignature:
                return PdfMutationCapabilityKind.SignatureChanges;
            case PdfMutationStructure.ObjectGraph when operation == PdfMutationOperation.ModifyAttachments:
                return PdfMutationCapabilityKind.AttachmentChanges;
            case PdfMutationStructure.ObjectGraph when operation == PdfMutationOperation.ChangeEncryption:
                return PdfMutationCapabilityKind.EncryptionChanges;
            case PdfMutationStructure.ObjectGraph:
                return PdfMutationCapabilityKind.ContentChanges;
            default:
                throw new ArgumentOutOfRangeException(nameof(structure), structure, "Unsupported PDF mutation structure.");
        }
    }

    private static System.Collections.ObjectModel.ReadOnlyCollection<T> ReadOnly<T>(params T[] values) => Array.AsReadOnly(values);

    private static void Add<T>(List<T> values, T value) where T : notnull {
        if (!values.Contains(value)) {
            values.Add(value);
        }
    }

    private static void ValidateOperation(PdfMutationOperation operation) {
        int value = (int)operation;
        if (value < (int)PdfMutationOperation.UpdateMetadata || value > (int)PdfMutationOperation.Redact) {
            throw new ArgumentOutOfRangeException(nameof(operation), operation, "Unsupported PDF mutation operation.");
        }
    }

    private static void ValidateExecutionPreference(PdfMutationExecutionPreference executionPreference) {
        int value = (int)executionPreference;
        if (value < (int)PdfMutationExecutionPreference.Automatic || value > (int)PdfMutationExecutionPreference.RequireAppendOnly) {
            throw new ArgumentOutOfRangeException(nameof(executionPreference), executionPreference, "Unsupported PDF mutation execution preference.");
        }
    }
}
