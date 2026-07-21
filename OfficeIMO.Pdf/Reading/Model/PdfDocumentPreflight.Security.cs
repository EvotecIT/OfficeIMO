namespace OfficeIMO.Pdf;

public sealed partial class PdfDocumentPreflight {
    private IReadOnlyList<string>? _securityDiagnostics;
    private IReadOnlyList<string>? _signatureMutationDiagnostics;
    private IReadOnlyList<string>? _appendOnlyMutationDiagnostics;
    private PdfAppendOnlyMutationReport? _appendOnlyMutationReport;

    /// <summary>Human-readable security, signature, and revision diagnostics derived from the lightweight probe.</summary>
    public IReadOnlyList<string> SecurityDiagnostics {
        get {
            if (_securityDiagnostics is not null) {
                return _securityDiagnostics;
            }

            var messages = new List<string>();
            PdfDocumentSecurityInfo security = Probe.Security;
            if (security.HasEncryption) {
                string encryption = "PDF encryption was detected";
                if (!string.IsNullOrEmpty(security.EncryptionFilter)) {
                    encryption += " using /Filter /" + security.EncryptionFilter;
                }

                if (!string.IsNullOrEmpty(security.EncryptionSubFilter)) {
                    encryption += " and /SubFilter /" + security.EncryptionSubFilter;
                }

                if (security.EncryptionRevision.HasValue || security.EncryptionLengthBits.HasValue) {
                    encryption += " (";
                    var parts = new List<string>();
                    if (security.EncryptionRevision.HasValue) {
                        parts.Add("R=" + security.EncryptionRevision.Value.ToString(System.Globalization.CultureInfo.InvariantCulture));
                    }

                    if (security.EncryptionLengthBits.HasValue) {
                        parts.Add(security.EncryptionLengthBits.Value.ToString(System.Globalization.CultureInfo.InvariantCulture) + "-bit");
                    }

                    encryption += string.Join(", ", parts) + ")";
                }

                encryption += security.PasswordAuthenticationRole == PdfPasswordAuthenticationRole.None
                    ? ". No password authorization was established."
                    : ". The supplied password authenticated as " + security.PasswordAuthenticationRole + ".";
                AddDistinct(messages, encryption);

                if (security.EncryptionPermissions.HasValue) {
                    AddDistinct(messages, "Raw encryption permissions /P=" + security.EncryptionPermissions.Value.ToString(System.Globalization.CultureInfo.InvariantCulture) + " were detected and are enforced for user-password operations unless the caller explicitly ignores restrictions.");
                }

                if (PermissionRestrictionsIgnored) {
                    AddDistinct(messages, "Authenticated user-password permission restrictions are being explicitly ignored for this operation.");
                }
            }

            AddRange(messages, SignatureMutationDiagnostics);

            if (security.HasDocMDPPermissions) {
                AddDistinct(messages, "Catalog /Perms contains DocMDP permissions; rewrite requires certification-signature preservation semantics.");
                if (security.DocMDPPermissionLevel.HasValue) {
                    AddDistinct(messages, "DocMDP certification permission level /P=" + security.DocMDPPermissionLevel.Value.ToString(System.Globalization.CultureInfo.InvariantCulture) + " was detected.");
                }
            }

            if (security.HasUsageRights) {
                AddDistinct(messages, "Catalog /Perms contains usage-rights entries; rewrite may invalidate viewer-extended rights.");
            }

            if (security.HasDocumentSecurityStore) {
                string dss = "Document Security Store (/DSS) was detected";
                if (security.DocumentSecurityStore.VriEntryCount > 0) {
                    dss += " with " + security.DocumentSecurityStore.VriEntryCount.ToString(System.Globalization.CultureInfo.InvariantCulture) + " VRI entr";
                    dss += security.DocumentSecurityStore.VriEntryCount == 1 ? "y" : "ies";
                }

                if (security.DocumentSecurityStore.HasValidationEvidence) {
                    dss += "; signature validation evidence must be preserved during mutation.";
                } else {
                    dss += ".";
                }

                AddDistinct(messages, dss);
            }

            if (security.HasIncrementalUpdates) {
                string revision = "Incremental update markers were detected";
                if (security.StartXrefCount > 0) {
                    revision += " (" + security.StartXrefCount.ToString(System.Globalization.CultureInfo.InvariantCulture) + " startxref section";
                    revision += security.StartXrefCount == 1 ? ")" : "s)";
                }

                revision += "; safe mutation requires append-only revision preservation.";
                AddDistinct(messages, revision);
            }

            if (security.HasXrefStreams) {
                AddDistinct(messages, "XRef stream markers were detected; rewrite must preserve or safely normalize cross-reference stream state.");
            }

            AddRange(messages, AppendOnlyMutationDiagnostics);
            _securityDiagnostics = messages.Count == 0 ? Array.Empty<string>() : messages.AsReadOnly();
            return _securityDiagnostics;
        }
    }

    /// <summary>True when security-specific diagnostics were produced.</summary>
    public bool HasSecurityDiagnostics => SecurityDiagnostics.Count > 0;

    /// <summary>True when signatures, append-only flags, rights, or existing revisions require mutation by adding a new PDF revision.</summary>
    public bool RequiresAppendOnlyMutation => Probe.Security.RequiresAppendOnlyMutation;

    /// <summary>True when the current OfficeIMO.Pdf writer can safely attempt append-only mutation for this input.</summary>
    public bool CanAppendOnlyMutate => RequiresAppendOnlyMutation && AppendOnlyMutationDiagnostics.Count == 0;

    /// <summary>Append-only mutation policy derived from the same security markers used by the incremental updater.</summary>
    public PdfAppendOnlyMutationReport AppendOnlyMutationReport {
        get {
            if (_appendOnlyMutationReport is not null) {
                return _appendOnlyMutationReport;
            }

            _appendOnlyMutationReport = PdfIncrementalUpdater.AnalyzeAppendOnlyMutation(Probe.Security);
            return _appendOnlyMutationReport;
        }
    }

    /// <summary>True when OfficeIMO.Pdf can append a metadata-only revision to this input.</summary>
    public bool CanAppendMetadataRevision => AppendOnlyMutationReport.CanAppendMetadata;

    /// <summary>True when OfficeIMO.Pdf can append simple AcroForm field-value revisions to this input.</summary>
    public bool CanAppendFormFieldRevision => AppendOnlyMutationReport.CanAppendFormFields;

    /// <summary>True when OfficeIMO.Pdf can append an external-signature placeholder revision to this input.</summary>
    public bool CanPrepareExternalSignatureRevision => AppendOnlyMutationReport.CanPrepareExternalSignature;

    /// <summary>Human-readable diagnostics explaining why append-only mutation cannot be attempted yet.</summary>
    public IReadOnlyList<string> AppendOnlyMutationDiagnostics {
        get {
            if (_appendOnlyMutationDiagnostics is not null) {
                return _appendOnlyMutationDiagnostics;
            }

            var messages = new List<string>();
            PdfDocumentSecurityInfo security = Probe.Security;
            if (!security.RequiresAppendOnlyMutation) {
                _appendOnlyMutationDiagnostics = Array.Empty<string>();
                return _appendOnlyMutationDiagnostics;
            }

            if (!CanRead) {
                if (ReadBlockers.Count == 0) {
                    AddDistinct(messages, "PDF append-only mutation is not available because OfficeIMO.Pdf cannot read this PDF.");
                } else {
                    for (int i = 0; i < ReadBlockers.Count; i++) {
                        AddDistinct(messages, ReadBlockers[i].Message);
                    }
                }
            }

            if (security.HasEncryption) {
                AddDistinct(messages, "Encrypted PDFs cannot be append-only mutated by OfficeIMO.Pdf yet.");
            }

            if (security.HasSignatures) {
                if (security.Signatures.Any(static signature => signature.HasFieldLock)) {
                    AddDistinct(messages, "Signature field locks restrict append-only form filling; requested fields must be checked before update.");
                }

                if (security.HasDocMDPPermissions &&
                    security.DocMDPPermissionLevel.HasValue &&
                    security.DocMDPPermissionLevel.Value >= 2 &&
                    security.DocMDPPermissionLevel.Value <= 3) {
                } else {
                    AddDistinct(messages, "Append-only signature preservation is not implemented by OfficeIMO.Pdf yet.");
                }
            }

            if (security.HasDocMDPPermissions) {
                if (security.DocMDPPermissionLevel.HasValue &&
                    security.DocMDPPermissionLevel.Value >= 2 &&
                    security.DocMDPPermissionLevel.Value <= 3) {
                } else {
                    AddDistinct(messages, "DocMDP certification permissions do not allow append-only form filling.");
                }
            }

            if (security.HasUsageRights) {
                AddDistinct(messages, "Usage-rights entries must be preserved before append-only mutation.");
            }

            if (security.HasXrefStreams) {
                AddDistinct(messages, "Append-only mutation for xref-stream PDFs is not implemented by OfficeIMO.Pdf yet.");
            }

            _appendOnlyMutationDiagnostics = messages.AsReadOnly();
            return _appendOnlyMutationDiagnostics;
        }
    }

    private IReadOnlyList<string> SignatureMutationDiagnostics {
        get {
            if (_signatureMutationDiagnostics is not null) {
                return _signatureMutationDiagnostics;
            }

            var messages = new List<string>();
            PdfDocumentSecurityInfo security = Probe.Security;
            if (security.HasSignatures) {
                string signature = "PDF signature markers were detected";
                if (security.SignatureFieldCount > 0) {
                    signature += " in " + security.SignatureFieldCount.ToString(System.Globalization.CultureInfo.InvariantCulture) + " signature field";
                    signature += security.SignatureFieldCount == 1 ? string.Empty : "s";
                }

                if (security.SignatureFieldNames.Count > 0) {
                    signature += " (" + string.Join(", ", security.SignatureFieldNames) + ")";
                }

                signature += "; rewrite would invalidate signatures unless append-only signature preservation is implemented.";
                AddDistinct(messages, signature);
            }

            if (security.HasByteRange) {
                string byteRange = "Signature /ByteRange markers were detected";
                if (security.ByteRangeSegmentCount > 0) {
                    byteRange += " with " + security.ByteRangeSegmentCount.ToString(System.Globalization.CultureInfo.InvariantCulture) + " segment";
                    byteRange += security.ByteRangeSegmentCount == 1 ? string.Empty : "s";
                }

                byteRange += ".";
                AddDistinct(messages, byteRange);
            }

            if (security.AcroFormAppendOnly) {
                AddDistinct(messages, "AcroForm /SigFlags indicates append-only updates are expected.");
            }

            _signatureMutationDiagnostics = messages.Count == 0 ? Array.Empty<string>() : messages.AsReadOnly();
            return _signatureMutationDiagnostics;
        }
    }
}
