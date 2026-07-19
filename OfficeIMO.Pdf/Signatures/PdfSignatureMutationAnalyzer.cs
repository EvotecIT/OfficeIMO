namespace OfficeIMO.Pdf;

/// <summary>Compares signature coverage, revision state, and structural permissions before and after a PDF mutation.</summary>
internal static class PdfSignatureMutationAnalyzer {
    /// <summary>Builds a before/after signature preservation report for a requested mutation.</summary>
    public static PdfSignatureMutationReport Analyze(
        byte[] before,
        byte[] after,
        PdfMutationOperation operation,
        IEnumerable<string>? fieldNames = null,
        PdfReadOptions? readOptions = null,
        PdfMutationExecutionPreference executionPreference = PdfMutationExecutionPreference.Automatic) {
        Guard.NotNull(before, nameof(before));
        Guard.NotNull(after, nameof(after));

        PdfMutationPlan plan = PdfMutationPlanner.Plan(before, operation, readOptions, fieldNames, executionPreference);
        PdfSignatureValidationReport beforeValidation = PdfSignatureValidator.Validate(before, readOptions);
        PdfReadOptions afterReadOptions = PdfReadOptions.WithMinimumInputBytes(readOptions, after.LongLength);
        PdfSignatureValidationReport afterValidation = PdfSignatureValidator.Validate(after, afterReadOptions);
        bool prefixPreserved = HasExactPrefix(after, before);
        bool revisionChainExtended = HasExtendedRevisionChain(beforeValidation.Security, afterValidation.Security);
        int[] beforeRevisionEnds = FindRevisionEnds(before);
        int[] afterRevisionEnds = FindRevisionEnds(after);
        PdfSignatureMutationPermissionStatus permission = GetPermissionStatus(plan, beforeValidation.HasSignatures);
        var results = new List<PdfSignatureMutationResult>(beforeValidation.Signatures.Count);

        for (int i = 0; i < beforeValidation.Signatures.Count; i++) {
            PdfSignatureValidationResult beforeSignature = beforeValidation.Signatures[i];
            PdfSignatureValidationResult? afterSignature = FindMatchingSignature(afterValidation.Signatures, beforeSignature.Signature);
            long? beforeSignedLength = GetSignedLength(beforeSignature.Signature);
            long? afterSignedLength = afterSignature is null ? null : GetSignedLength(afterSignature.Signature);
            int[] coveredBefore = GetCoveredRevisions(beforeRevisionEnds, beforeSignedLength);
            int[] coveredAfter = GetCoveredRevisions(afterRevisionEnds, afterSignedLength);
            bool byteRangePreserved = afterSignature is not null &&
                beforeSignature.Signature.ByteRangeValues.SequenceEqual(afterSignature.Signature.ByteRangeValues);

            results.Add(new PdfSignatureMutationResult(
                beforeSignature,
                afterSignature,
                coveredBefore.Length == 0 ? null : coveredBefore[coveredBefore.Length - 1],
                coveredAfter.Length == 0 ? null : coveredAfter[coveredAfter.Length - 1],
                Array.AsReadOnly(coveredBefore),
                Array.AsReadOnly(coveredAfter),
                prefixPreserved,
                byteRangePreserved,
                HasLaterRevisions(beforeSignature, coveredBefore, beforeValidation.Security),
                afterSignature is not null && HasLaterRevisions(afterSignature, coveredAfter, afterValidation.Security),
                permission));
        }

        IReadOnlyList<string> diagnostics = BuildDiagnostics(plan, beforeValidation, prefixPreserved, revisionChainExtended, results);
        return new PdfSignatureMutationReport(
            plan,
            beforeValidation,
            afterValidation,
            prefixPreserved,
            revisionChainExtended,
            results.AsReadOnly(),
            diagnostics);
    }

    private static PdfSignatureMutationPermissionStatus GetPermissionStatus(PdfMutationPlan plan, bool hasSignatures) {
        if (!hasSignatures) {
            return PdfSignatureMutationPermissionStatus.NotApplicableUnsigned;
        }

        if (!plan.CanExecute) {
            return PdfSignatureMutationPermissionStatus.Forbidden;
        }

        return plan.ExecutionMode == PdfMutationExecutionMode.AppendOnly
            ? PdfSignatureMutationPermissionStatus.Permitted
            : PdfSignatureMutationPermissionStatus.Indeterminate;
    }

    private static bool HasExtendedRevisionChain(PdfDocumentSecurityInfo before, PdfDocumentSecurityInfo after) =>
        before.LastStartXrefOffset.HasValue &&
        after.RevisionCount > before.RevisionCount &&
        after.Revisions.Any(revision => revision.PreviousXrefOffset == before.LastStartXrefOffset);

    private static PdfSignatureValidationResult? FindMatchingSignature(
        IReadOnlyList<PdfSignatureValidationResult> signatures,
        PdfSignatureInfo signature) {
        for (int i = 0; i < signatures.Count; i++) {
            PdfSignatureInfo candidate = signatures[i].Signature;
            if (candidate.ObjectNumber == signature.ObjectNumber &&
                string.Equals(candidate.FieldName, signature.FieldName, StringComparison.Ordinal)) {
                return signatures[i];
            }
        }

        return null;
    }

    private static long? GetSignedLength(PdfSignatureInfo signature) {
        if (signature.ByteRangeValues.Count < 4) {
            return null;
        }

        long secondOffset = signature.ByteRangeValues[2];
        long secondLength = signature.ByteRangeValues[3];
        return secondOffset >= 0 && secondLength >= 0 && secondOffset <= long.MaxValue - secondLength
            ? secondOffset + secondLength
            : null;
    }

    private static int[] FindRevisionEnds(byte[] pdf) {
        byte[] marker = Encoding.ASCII.GetBytes("%%EOF");
        var ends = new List<int>();
        for (int i = 0; i <= pdf.Length - marker.Length; i++) {
            if (!MatchesAt(pdf, marker, i)) {
                continue;
            }

            int end = i + marker.Length;
            while (end < pdf.Length && (pdf[end] == (byte)'\r' || pdf[end] == (byte)'\n')) {
                end++;
            }

            ends.Add(end);
            i += marker.Length - 1;
        }

        return ends.ToArray();
    }

    private static int[] GetCoveredRevisions(int[] revisionEnds, long? signedLength) {
        if (!signedLength.HasValue) {
            return Array.Empty<int>();
        }

        var covered = new List<int>();
        for (int i = 0; i < revisionEnds.Length; i++) {
            if (revisionEnds[i] <= signedLength.Value) {
                covered.Add(i + 1);
            }
        }

        return covered.ToArray();
    }

    private static bool HasLaterRevisions(
        PdfSignatureValidationResult signature,
        int[] coveredRevisions,
        PdfDocumentSecurityInfo security) =>
        coveredRevisions.Length > 0
            ? security.RevisionCount > coveredRevisions[coveredRevisions.Length - 1]
            : signature.UnsignedByteCount.GetValueOrDefault() > 0 || security.HasIncrementalUpdates;

    private static bool HasExactPrefix(byte[] value, byte[] prefix) {
        if (value.Length < prefix.Length) {
            return false;
        }

        for (int i = 0; i < prefix.Length; i++) {
            if (value[i] != prefix[i]) {
                return false;
            }
        }

        return true;
    }

    private static bool MatchesAt(byte[] value, byte[] expected, int offset) {
        for (int i = 0; i < expected.Length; i++) {
            if (value[offset + i] != expected[i]) {
                return false;
            }
        }

        return true;
    }

    private static System.Collections.ObjectModel.ReadOnlyCollection<string> BuildDiagnostics(
        PdfMutationPlan plan,
        PdfSignatureValidationReport before,
        bool prefixPreserved,
        bool revisionChainExtended,
        IReadOnlyList<PdfSignatureMutationResult> results) {
        var diagnostics = new List<string>();
        diagnostics.Add(plan.CanExecute ? "Mutation.Permitted" : "Mutation.Forbidden");
        diagnostics.Add(prefixPreserved ? "Bytes.InputPrefixPreserved" : "Bytes.InputPrefixChanged");
        diagnostics.Add(revisionChainExtended ? "Revisions.ChainExtended" : "Revisions.ChainNotExtended");
        if (!before.HasSignatures) {
            diagnostics.Add("Signatures.None");
        } else if (results.All(static result => result.IsStructurallyPreserved)) {
            diagnostics.Add("Signatures.StructurallyPreserved");
        } else {
            diagnostics.Add("Signatures.StructuralPreservationFailed");
        }

        return diagnostics.AsReadOnly();
    }
}
